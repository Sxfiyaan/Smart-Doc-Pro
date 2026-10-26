"""Upload and download API for SmartDoc milestone 1."""

from __future__ import annotations

import asyncio
import re
import shutil
import tempfile
import time
import uuid
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_FILES_PER_UPLOAD = 20
MAX_BATCH_BYTES = 100 * 1024 * 1024
SESSION_TTL_SECONDS = 60 * 60
SUPPORTED_EXTENSIONS = {".xlsx", ".csv", ".docx", ".txt"}
STORAGE_ROOT = Path(tempfile.gettempdir()) / "smartdoc-sessions"
FRONTEND_DIR = Path(__file__).resolve().parents[1] / "frontend"


@dataclass
class SessionFile:
    file_id: str
    display_name: str
    path: Path
    size: int


@dataclass
class Session:
    session_id: str
    directory: Path
    files: dict[str, SessionFile] = field(default_factory=dict)
    touched_at: float = field(default_factory=time.monotonic)


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
            total_size += item.size
            if total_size > MAX_BATCH_BYTES:
                raise HTTPException(status_code=413, detail="Combined upload must be 100 MB or smaller")
    except Exception:
        for item in uploaded:
            session.files.pop(item["file_id"], None)
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
                        "size": item.size}
                       for item in session.files.values()]}


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
        session = sessions.pop(session_id, None)
        if session is None:
            raise HTTPException(status_code=404, detail="Session not found")
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
