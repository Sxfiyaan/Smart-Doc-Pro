"""Upload and download API for SmartDoc milestone 1."""

from __future__ import annotations

import asyncio
import os
import re
import shutil
import tempfile
import time
import uuid
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from starlette.concurrency import run_in_threadpool

from backend.session_file import SessionFile
from backend import agent

MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_FILES_PER_UPLOAD = 20
MAX_BATCH_BYTES = 100 * 1024 * 1024
SESSION_TTL_SECONDS = 60 * 60
SUPPORTED_EXTENSIONS = {".xlsx", ".csv", ".docx", ".txt"}
STORAGE_ROOT = Path(tempfile.gettempdir()) / "smartdoc-sessions"
FRONTEND_DIR = Path(__file__).resolve().parents[1] / "frontend"


@dataclass
class Session:
    session_id: str
    directory: Path
    files: dict[str, SessionFile] = field(default_factory=dict)
    touched_at: float = field(default_factory=time.monotonic)
    current_file_ids: list[str] = field(default_factory=list)
    undo_history: list[dict] = field(default_factory=list)
    chat_history: list[dict[str, str]] = field(default_factory=list)
    pending_command: agent.CommandContext | None = None
    command_lock: asyncio.Lock = field(default_factory=asyncio.Lock)
    active_command_id: str | None = None
    command_contexts: dict[str, agent.CommandContext] = field(default_factory=dict)
    command_results: dict[str, dict] = field(default_factory=dict)
    command_tasks: set[asyncio.Task] = field(default_factory=set)


class CommandRequest(BaseModel):
    command: str
    output_format: Literal["docx", "xlsx"] | None = None


class ConfirmationRequest(BaseModel):
    confirmation_id: str
    confirmed: bool


sessions: dict[str, Session] = {}
sessions_lock = asyncio.Lock()
owned_sessions: set[str] = set()


def sanitize_filename(filename: str | None) -> str:
    """Keep only a safe basename and a conservative set of characters."""
    original = (filename or "upload").replace("\\", "/").split("/")[-1]
    cleaned = re.sub(r"[^A-Za-z0-9._ -]", "_", original).strip(" .")
    return cleaned[:180] or "upload"


async def cleanup_expired_sessions() -> None:
    now = time.monotonic()
    async with sessions_lock:
        expired = [key for key, session in sessions.items()
                   if now - session.touched_at > SESSION_TTL_SECONDS]
        for key in expired:
            session = sessions.pop(key)
            shutil.rmtree(session.directory, ignore_errors=True)


async def create_session() -> Session:
    await cleanup_expired_sessions()
    session_id = uuid.uuid4().hex
    directory = STORAGE_ROOT / session_id
    directory.mkdir(parents=True, exist_ok=False)
    session = Session(session_id=session_id, directory=directory)
    async with sessions_lock:
        sessions[session_id] = session
        owned_sessions.add(session_id)
    return session


async def get_session(session_id: str) -> Session:
    await cleanup_expired_sessions()
    async with sessions_lock:
        session = sessions.get(session_id)
        if session is None:
            raise HTTPException(status_code=404, detail="Session not found or expired")
        session.touched_at = time.monotonic()
        return session


async def store_upload(session: Session, upload: UploadFile) -> SessionFile:
    display_name = sanitize_filename(upload.filename)
    suffix = Path(display_name).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        await upload.close()
        raise HTTPException(status_code=415, detail=f"Unsupported file type: {suffix or '(none)'}")

    file_id = uuid.uuid4().hex
    target = session.directory / f"{file_id}{suffix}"
    size = 0
    try:
        with target.open("wb") as destination:
            while chunk := await upload.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_FILE_BYTES:
                    raise HTTPException(status_code=413, detail="Each file must be 20 MB or smaller")
                destination.write(chunk)
    except Exception:
        target.unlink(missing_ok=True)
        raise
    finally:
        await upload.close()

    item = SessionFile(file_id=file_id, display_name=display_name, path=target, size=size)
    session.files[file_id] = item
    return item


@asynccontextmanager
async def lifespan(_: FastAPI):
    STORAGE_ROOT.mkdir(parents=True, exist_ok=True)
    # Session metadata lives in memory, so directories left by a prior process
    # are unreachable and safe to remove when this process starts.
    for stale_directory in STORAGE_ROOT.iterdir():
        if stale_directory.is_dir() and re.fullmatch(r"[0-9a-f]{32}", stale_directory.name):
            shutil.rmtree(stale_directory, ignore_errors=True)
    yield
    async with sessions_lock:
        for session_id in owned_sessions:
            session = sessions.pop(session_id, None)
            if session is not None:
                shutil.rmtree(session.directory, ignore_errors=True)
        owned_sessions.clear()


app = FastAPI(title="SmartDoc", version="0.1.0", lifespan=lifespan)

_default_frontend_origins = {
    "null",  # Standalone local index.html files.
    "http://localhost:3000", "http://127.0.0.1:3000",
    "http://localhost:5173", "http://127.0.0.1:5173",
    "http://localhost:5500", "http://127.0.0.1:5500",
    "http://localhost:8000", "http://127.0.0.1:8000",
}
_extra_frontend_origins = {
    origin.strip() for origin in os.getenv("SMARTDOC_ALLOWED_ORIGINS", "").split(",")
    if origin.strip()
}
app.add_middleware(
    CORSMiddleware,
    allow_origins=sorted(_default_frontend_origins | _extra_frontend_origins),
    allow_credentials=False,
    allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
    allow_headers=["Content-Type"],
)


@app.get("/api/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/sessions", status_code=201)
async def new_session() -> dict[str, str]:
    session = await create_session()
    return {"session_id": session.session_id}


@app.post("/api/sessions/{session_id}/files", status_code=201)
async def upload_files(session_id: str, files: list[UploadFile] = File(...)) -> dict:
    session = await get_session(session_id)
    if session.active_command_id or session.pending_command:
        raise HTTPException(status_code=409, detail="Wait for the active command to finish first")
    if len(files) > MAX_FILES_PER_UPLOAD:
        for upload in files:
            await upload.close()
        raise HTTPException(status_code=413, detail="Upload at most 20 files at a time")
    uploaded = []
    total_size = 0
    try:
        for upload in files:
            item = await store_upload(session, upload)
            uploaded.append({"file_id": item.file_id, "file_name": item.display_name,
                             "size": item.size})
            session.current_file_ids.append(item.file_id)
            total_size += item.size
            if total_size > MAX_BATCH_BYTES:
                raise HTTPException(status_code=413, detail="Combined upload must be 100 MB or smaller")
    except Exception:
        for item in uploaded:
            session.files.pop(item["file_id"], None)
            if item["file_id"] in session.current_file_ids:
                session.current_file_ids.remove(item["file_id"])
            matching_path = session.directory / f'{item["file_id"]}{Path(item["file_name"]).suffix.lower()}'
            matching_path.unlink(missing_ok=True)
        for upload in files:
            await upload.close()
        raise
    return {"files": uploaded}


@app.get("/api/sessions/{session_id}/files")
async def list_files(session_id: str) -> dict:
    session = await get_session(session_id)
    return {"files": [{"file_id": item.file_id, "file_name": item.display_name,
                        "size": item.size, "is_current": key in session.current_file_ids,
                        "is_output": item.is_output}
                       for key, item in session.files.items()]}


@app.post("/api/sessions/{session_id}/inspect")
async def inspect_files(session_id: str) -> dict:
    session = await get_session(session_id)
    async with session.command_lock:
        if session.active_command_id or session.pending_command:
            raise HTTPException(status_code=409, detail="Wait for the active command to finish first")
        try:
            return await run_in_threadpool(agent.inspect_session, session)
        except Exception as exc:
            raise HTTPException(status_code=502, detail=f"File inspection failed: {exc}") from exc


@app.delete("/api/sessions/{session_id}/files/{file_id}")
async def delete_file(session_id: str, file_id: str) -> dict:
    session = await get_session(session_id)
    async with session.command_lock:
        if session.active_command_id or session.pending_command:
            raise HTTPException(status_code=409, detail="Wait for the active command to finish first")
        item = session.files.get(file_id)
        if item is None:
            raise HTTPException(status_code=404, detail="File not found")
        if item.is_output:
            raise HTTPException(status_code=409, detail="Generated files are removed with Undo")

        session.files.pop(file_id)
        session.current_file_ids = [current_id for current_id in session.current_file_ids
                                    if current_id != file_id]
        for snapshot in session.undo_history:
            snapshot["previous_file_ids"] = [saved_id for saved_id in snapshot["previous_file_ids"]
                                              if saved_id != file_id]
            snapshot["created_file_ids"] = [saved_id for saved_id in snapshot["created_file_ids"]
                                            if saved_id != file_id]
        item.path.unlink(missing_ok=True)
        session.chat_history.append({
            "role": "assistant",
            "content": f"Removed {item.display_name} from this session. I will use the remaining files for the next command.",
        })
        session.chat_history = session.chat_history[-agent.MAX_CHAT_HISTORY_MESSAGES:]
        return {"status": "deleted", "file_id": file_id,
                "current_file_ids": session.current_file_ids}


@app.post("/api/sessions/{session_id}/commands")
async def run_command(session_id: str, request: CommandRequest) -> dict:
    session = await get_session(session_id)
    async with session.command_lock:
        if session.pending_command is not None or session.active_command_id:
            raise HTTPException(status_code=409, detail="Confirm or cancel the pending action first")
        try:
            context = agent.start_command(session, request.command, request.output_format)
        except (ValueError, RuntimeError) as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return _schedule_command(session, context)


def _schedule_command(session: Session, context: agent.CommandContext,
                      *, confirmed: bool = False) -> dict:
    session.active_command_id = context.command_id
    session.command_contexts[context.command_id] = context
    session.command_results[context.command_id] = {"status": "running", "activity": []}
    task = asyncio.create_task(_run_command_task(session, context, confirmed=confirmed))
    session.command_tasks.add(task)
    task.add_done_callback(session.command_tasks.discard)
    return {"status": "running", "command_id": context.command_id}


async def _run_command_task(session: Session, context: agent.CommandContext,
                            *, confirmed: bool = False) -> None:
    try:
        result = await run_in_threadpool(agent.run_loop, session, context, confirmed=confirmed)
        async with session.command_lock:
            session.command_results[context.command_id] = result
            session.active_command_id = None
            session.pending_command = context if result["status"] == "confirmation_required" else None
    except Exception as exc:
        agent.abort_context(context)
        async with session.command_lock:
            session.command_results[context.command_id] = {
                "status": "error", "detail": f"Agent command failed: {exc}",
                "activity": context.activity,
                "usage": agent.groq_usage.summarize(context.usage_records),
            }
            session.active_command_id = None
            session.pending_command = None


@app.get("/api/sessions/{session_id}/commands/{command_id}")
async def command_status(session_id: str, command_id: str) -> dict:
    session = await get_session(session_id)
    result = session.command_results.get(command_id)
    if result is None:
        raise HTTPException(status_code=404, detail="Command not found")
    if result["status"] == "running":
        context = session.command_contexts.get(command_id)
        result = {"status": "running", "command_id": command_id,
                  "activity": list(context.activity) if context else [],
                  "usage": agent.groq_usage.summarize(context.usage_records) if context else None}
    return result


@app.post("/api/sessions/{session_id}/commands/confirm")
async def confirm_command(session_id: str, request: ConfirmationRequest) -> dict:
    session = await get_session(session_id)
    async with session.command_lock:
        context = session.pending_command
        if context is None or context.command_id != request.confirmation_id:
            raise HTTPException(status_code=404, detail="Pending confirmation not found")
        if not request.confirmed:
            session.pending_command = None
            agent.abort_context(context)
            result = {"status": "cancelled", "summary": "Action cancelled. No uploaded files were changed."}
            session.command_results[context.command_id] = result
            return result
        return _schedule_command(session, context, confirmed=True)


@app.post("/api/sessions/{session_id}/undo")
async def undo_command(session_id: str) -> dict:
    session = await get_session(session_id)
    async with session.command_lock:
        if session.pending_command is not None or session.active_command_id:
            raise HTTPException(status_code=409, detail="Resolve the pending confirmation before undoing")
        if not session.undo_history:
            raise HTTPException(status_code=409, detail="There is no command to undo")
        snapshot = session.undo_history.pop()
        for file_id in snapshot["created_file_ids"]:
            item = session.files.pop(file_id, None)
            if item is not None:
                item.path.unlink(missing_ok=True)
        session.current_file_ids = snapshot["previous_file_ids"]
        return {"status": "undone", "current_files": agent._summary_for_ids(session, session.current_file_ids),
                "can_undo": bool(session.undo_history)}


@app.get("/api/sessions/{session_id}/files/{file_id}/download")
async def download_file(session_id: str, file_id: str) -> FileResponse:
    session = await get_session(session_id)
    item = session.files.get(file_id)
    if item is None or not item.path.is_file():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(item.path, filename=item.display_name)


@app.delete("/api/sessions/{session_id}", status_code=204)
async def delete_session(session_id: str) -> None:
    async with sessions_lock:
        session = sessions.get(session_id)
        if session is None:
            raise HTTPException(status_code=404, detail="Session not found")
        if session.active_command_id or session.pending_command:
            raise HTTPException(status_code=409, detail="Finish or cancel the active command first")
        sessions.pop(session_id, None)
        owned_sessions.discard(session_id)
        shutil.rmtree(session.directory, ignore_errors=True)


@app.get("/", response_class=HTMLResponse)
async def frontend() -> FileResponse:
    page = FRONTEND_DIR / "index.html"
    if not page.is_file():
        raise HTTPException(status_code=404, detail="Frontend is not installed")
    return FileResponse(page)


if FRONTEND_DIR.is_dir():
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")
