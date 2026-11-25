"""Groq tool-calling loop for safe Excel and CSV edits."""

from __future__ import annotations

import json
import os
import re
import shutil
import time
import uuid
from copy import deepcopy
from dataclasses import dataclass, field
from importlib import import_module
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from groq import Groq
from groq.types.chat import ChatCompletion
import httpx
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries

from backend.session_file import SessionFile
from backend.tools import generation
from backend.tools import excel_tools as excel
from backend.tools import word_tools as word

groq_usage = import_module("groq usage.tracker")

load_dotenv()

MAX_TOOL_TURNS = 12
MAX_CHAT_HISTORY_MESSAGES = 6
DESTRUCTIVE_TOOLS = {"clear_cells", "clear_columns", "clear_rows", "delete_column", "delete_columns", "delete_row", "delete_rows", "delete_table", "delete_sheet", "remove_duplicates", "word_delete_paragraph", "word_delete_table", "word_delete_table_row", "word_delete_table_column"}


def _tool(name: str, description: str, properties: dict[str, Any], required: list[str]) -> dict[str, Any]:
    return {"type": "function", "function": {
        "name": name, "description": description,
        "parameters": {"type": "object", "properties": properties, "required": required,
                        "additionalProperties": False},
    }}


FILE = {"type": "string", "description": "File ID listed in the provided file summaries"}
SHEET = {"type": ["string", "null"], "description": "Workbook sheet name; use null for CSV"}
RANGE = {"type": "string", "description": "Excel-style cell range such as A1:C10"}
TOOLS = [
    _tool("draft_from_template", "Create a separate Word document from an existing .docx. Preserve its layout. Replace exact source phrases (including placeholders in paragraphs or tables) and optionally append styled paragraphs. Supply all changes in one call.", {
        "file_id": FILE, "title": {"type": "string"},
        "replacements": {"type": "array", "items": {"type": "object", "properties": {
            "source": {"type": "string"}, "text": {"type": "string"}}, "required": ["source", "text"]}},
        "append": {"type": "array", "items": {"type": "object", "properties": {
            "style": {"type": "string"}, "text": {"type": "string"}}, "required": ["style", "text"]}},
    }, ["file_id", "title", "replacements", "append"]),
    _tool("create_document", "Write a complete new Word document in one call. Supply the actual content, with one short section per page (up to 1200 characters each). Respect the requested page limit.", {
        "title": {"type": "string"}, "pages": {"type": "array", "minItems": 1,
        "items": {"type": "object", "properties": {"heading": {"type": "string"}, "text": {"type": "string"}},
                  "required": ["heading", "text"]}}}, ["title", "pages"]),
    _tool("create_workbook", "Write a complete new Excel workbook in one call. Include the header row and actual data in rows.", {
        "title": {"type": "string"}, "sheets": {"type": "array", "minItems": 1,
        "items": {"type": "object", "properties": {"name": {"type": "string"},
                  "rows": {"type": "array", "items": {"type": "array", "items": {"type": ["string", "number", "boolean", "null"]}}}},
                  "required": ["name", "rows"]}}}, ["title", "sheets"]),
    _tool("list_sheets", "List sheet names in a workbook. CSV has one Sheet1 data sheet.", {"file_id": FILE}, ["file_id"]),
    _tool("get_summary", "Get sheet names, headers, row counts and sample rows.", {"file_id": FILE}, ["file_id"]),
    _tool("read_range", "Read cell values from an Excel or CSV range.", {"file_id": FILE, "sheet": SHEET, "cell_range": RANGE}, ["file_id", "sheet", "cell_range"]),
    _tool("create_table", "Create a formatted Excel table. Requires a header row and data rows.", {"file_id": FILE, "sheet": SHEET, "cell_range": RANGE, "headers": {"type": ["array", "null"], "items": {"type": "string"}}, "style": {"type": "string"}}, ["file_id", "sheet", "cell_range"]),
    _tool("modify_table", "Change an Excel table's range or name.", {"file_id": FILE, "sheet": SHEET, "table_name": {"type": "string"}, "cell_range": {"type": ["string", "null"]}, "new_name": {"type": ["string", "null"]}}, ["file_id", "sheet", "table_name"]),
    _tool("update_cells", "Set one or more existing cell values by address.", {"file_id": FILE, "sheet": SHEET, "updates": {"type": "object", "additionalProperties": True}}, ["file_id", "sheet", "updates"]),
    _tool("clear_cells", "Clear the contents of explicitly listed Excel or CSV data cells, preserving formatting. Requires confirmation. Use this for removing selected cell values; do not pass blank cells or whole columns.", {"file_id": FILE, "sheet": SHEET, "cells": {"type": "array", "items": {"type": "string"}, "minItems": 1, "maxItems": 1000}}, ["file_id", "sheet", "cells"]),
    _tool("clear_columns", "Clear all data values in a contiguous column range such as G:J, while retaining the column headers, columns, and formatting. Requires confirmation.", {"file_id": FILE, "sheet": SHEET, "columns": {"type": "string", "description": "Excel column letters/range such as G:J"}}, ["file_id", "sheet", "columns"]),
    _tool("clear_rows", "Clear all values in a contiguous data-row range such as 5:8, while retaining the rows and their formatting. Row 1 is the header and cannot be cleared. Requires confirmation.", {"file_id": FILE, "sheet": SHEET, "rows": {"type": "string", "description": "Worksheet row number or range, such as 5:8"}}, ["file_id", "sheet", "rows"]),
    _tool("add_column", "Add a column with a header and optional values.", {"file_id": FILE, "sheet": SHEET, "header": {"type": "string"}, "values": {"type": ["array", "null"], "items": {}}, "index": {"type": ["integer", "null"]}}, ["file_id", "sheet", "header"]),
    _tool("delete_column", "Delete a column by letter, index, or CSV header. This requires confirmation.", {"file_id": FILE, "sheet": SHEET, "column": {"type": ["string", "integer"]}}, ["file_id", "sheet", "column"]),
    _tool("delete_columns", "Delete a contiguous range of columns, such as G:J. This requires confirmation. For CSV, columns can also be addressed by header or index range.", {"file_id": FILE, "sheet": SHEET, "columns": {"type": "string", "description": "Excel column letters/range such as G:J"}}, ["file_id", "sheet", "columns"]),
    _tool("add_row", "Append or insert a data row. Row values must match the column count.", {"file_id": FILE, "sheet": SHEET, "values": {"type": "array", "items": {}}, "index": {"type": ["integer", "null"]}}, ["file_id", "sheet", "values"]),
    _tool("delete_row", "Delete a data row by its sheet row number. This requires confirmation.", {"file_id": FILE, "sheet": SHEET, "index": {"type": "integer"}}, ["file_id", "sheet", "index"]),
    _tool("delete_rows", "Delete a contiguous range of data rows by worksheet row numbers, such as 5:8. Row 1 is the header and cannot be deleted. This requires confirmation.", {"file_id": FILE, "sheet": SHEET, "rows": {"type": "string", "description": "Worksheet row number or range, such as 5:8"}}, ["file_id", "sheet", "rows"]),
    _tool("delete_table", "Delete an Excel table definition. This requires confirmation.", {"file_id": FILE, "sheet": {"type": "string"}, "table_name": {"type": "string"}}, ["file_id", "sheet", "table_name"]),
    _tool("delete_sheet", "Delete a worksheet. This requires confirmation.", {"file_id": FILE, "sheet": {"type": "string"}}, ["file_id", "sheet"]),
    _tool("rename_sheet", "Rename a worksheet.", {"file_id": FILE, "sheet": {"type": "string"}, "new_name": {"type": "string"}}, ["file_id", "sheet", "new_name"]),
    _tool("sort", "Sort data rows in the selected range by a column (ascending or descending).", {"file_id": FILE, "sheet": SHEET, "cell_range": RANGE, "column": {"type": ["integer", "string"]}, "order": {"type": "string", "enum": ["asc", "desc"]}}, ["file_id", "sheet", "cell_range", "column"]),
    _tool("filter", "Return rows matching a condition without changing the file.", {"file_id": FILE, "sheet": SHEET, "cell_range": RANGE, "condition": {"type": "object", "properties": {"column": {"type": ["integer", "string"]}, "operator": {"type": "string", "enum": ["equals", "not_equals", "contains", "greater_than", "greater_or_equal", "less_than", "less_or_equal"]}, "value": {}}, "required": ["column", "value"], "additionalProperties": False}}, ["file_id", "sheet", "cell_range", "condition"]),
    _tool("remove_duplicates", "Remove duplicate data rows using all columns or selected key columns. This requires confirmation.", {"file_id": FILE, "sheet": SHEET, "cell_range": {"type": ["string", "null"]}, "columns": {"type": ["array", "null"], "items": {"type": ["string", "integer"]}}}, ["file_id", "sheet"]),
    _tool("apply_formatting", "Apply bold, fill, borders, or number format to an Excel range.", {"file_id": FILE, "sheet": {"type": "string"}, "cell_range": RANGE, "bold": {"type": ["boolean", "null"]}, "fill": {"type": ["string", "null"]}, "borders": {"type": ["boolean", "null"]}, "number_format": {"type": ["string", "null"]}}, ["file_id", "sheet", "cell_range"]),
    _tool("add_formula", "Write an Excel formula to one cell.", {"file_id": FILE, "sheet": {"type": "string"}, "cell": {"type": "string"}, "formula": {"type": "string"}}, ["file_id", "sheet", "cell", "formula"]),
    _tool("merge_files", "Merge two or more uploaded/current files into a new Excel workbook by appending rows or joining on a key.", {"file_ids": {"type": "array", "items": {"type": "string"}}, "mode": {"type": "string", "enum": ["append", "join_on_key"]}, "key": {"type": ["string", "null"]}}, ["file_ids", "mode"]),
    _tool("word_merge_documents", "Create a new .docx by appending the body content of two or more uploaded Word documents in the supplied order. The source files are preserved.", {"file_ids": {"type": "array", "items": {"type": "string"}, "minItems": 2}}, ["file_ids"]),
    _tool("word_read_document", "Read Word paragraphs, headings, and tables. Paragraph and table indexes are zero-based.", {"file_id": FILE}, ["file_id"]),
    _tool("word_insert_paragraph", "Insert a paragraph before an index, or append it when index is null.", {"file_id": FILE, "text": {"type": "string"}, "index": {"type": ["integer", "null"]}, "style": {"type": ["string", "null"]}}, ["file_id", "text"]),
    _tool("word_edit_paragraph", "Replace one paragraph's text and optionally its style.", {"file_id": FILE, "index": {"type": "integer"}, "text": {"type": "string"}, "style": {"type": ["string", "null"]}}, ["file_id", "index", "text"]),
    _tool("word_apply_style", "Apply a built-in Word paragraph style such as Heading 1 or Normal.", {"file_id": FILE, "index": {"type": "integer"}, "style": {"type": "string"}}, ["file_id", "index", "style"]),
    _tool("word_delete_paragraph", "Delete one paragraph. Requires confirmation.", {"file_id": FILE, "index": {"type": "integer"}}, ["file_id", "index"]),
    _tool("word_create_table", "Create a Word table with optional headers, style, and paragraph location.", {"file_id": FILE, "rows": {"type": "integer", "minimum": 1}, "columns": {"type": "integer", "minimum": 1}, "headers": {"type": ["array", "null"], "items": {"type": "string"}}, "style": {"type": ["string", "null"]}, "location": {"type": ["integer", "null"]}}, ["file_id", "rows", "columns"]),
    _tool("word_modify_table_cell", "Replace a Word table cell's text. Indexes are zero-based.", {"file_id": FILE, "table_index": {"type": "integer"}, "row": {"type": "integer"}, "column": {"type": "integer"}, "text": {"type": "string"}}, ["file_id", "table_index", "row", "column", "text"]),
    _tool("word_add_table_row", "Append a row to a Word table.", {"file_id": FILE, "table_index": {"type": "integer"}, "values": {"type": ["array", "null"], "items": {}}}, ["file_id", "table_index"]),
    _tool("word_add_table_column", "Append a column to a Word table. The first row is treated as the header row.", {"file_id": FILE, "table_index": {"type": "integer"}, "header": {"type": "string"}, "values": {"type": ["array", "null"], "items": {}}}, ["file_id", "table_index"]),
    _tool("word_delete_table", "Delete a Word table. Requires confirmation.", {"file_id": FILE, "table_index": {"type": "integer"}}, ["file_id", "table_index"]),
    _tool("word_delete_table_row", "Delete one Word table row. Requires confirmation.", {"file_id": FILE, "table_index": {"type": "integer"}, "row": {"type": "integer"}}, ["file_id", "table_index", "row"]),
    _tool("word_delete_table_column", "Delete one Word table column. Requires confirmation.", {"file_id": FILE, "table_index": {"type": "integer"}, "column": {"type": "integer"}}, ["file_id", "table_index", "column"]),
]

ALL_TOOL_NAMES = [tool["function"]["name"] for tool in TOOLS]
TOOLS_BY_NAME = {tool["function"]["name"]: tool for tool in TOOLS}

# Explicit groups keep routing independent of naming conventions. Several Excel
# tools intentionally have no "excel_" prefix in the current public registry.
TOOL_GROUPS = {
    "create": ["create_document", "create_workbook"],
    "docx": ["create_document"],
    "xlsx": ["create_workbook"],
    "template": ["draft_from_template"],
    "merge_excel": ["merge_files"],
    "merge_word": ["word_merge_documents"],
    "word_all": [
        "create_document", "draft_from_template",
        "word_merge_documents", "word_read_document", "word_insert_paragraph",
        "word_edit_paragraph", "word_apply_style", "word_delete_paragraph",
        "word_create_table", "word_modify_table_cell", "word_add_table_row",
        "word_add_table_column", "word_delete_table", "word_delete_table_row",
        "word_delete_table_column",
    ],
    "excel_all": [
        "create_workbook",
        "list_sheets", "get_summary", "read_range", "create_table", "modify_table",
        "update_cells", "clear_cells", "clear_columns", "clear_rows", "add_column",
        "delete_column", "delete_columns", "add_row", "delete_row", "delete_rows",
        "delete_table", "delete_sheet", "rename_sheet", "sort", "filter",
        "remove_duplicates", "apply_formatting", "add_formula", "merge_files",
    ],
}
FULL_TOOL_SET = ALL_TOOL_NAMES
SUPPORTED_WORD_EXTENSIONS = {"docx"}
SUPPORTED_EXCEL_EXTENSIONS = {"xlsx", "csv"}
MERGE_RE = re.compile(
    r"\b(?:merge|combine|join)\b|\bappend\s+(?:(?:the|all|these|those)\s+)?(?:files?|documents?|workbooks?)\b",
    re.IGNORECASE,
)
IN_FILE_MERGE_RE = re.compile(r"\b(?:cells?|tables?|rows?|columns?|paragraphs?|sheets?)\b", re.IGNORECASE)


def file_kinds(files: list[Any]) -> tuple[bool, bool, bool]:
    """Return (has_word, has_excel, has_unknown) from names or SessionFile objects."""
    has_word = has_excel = has_unknown = False
    for file in files:
        name = getattr(file, "display_name", file)
        extension = Path(str(name)).suffix.lower().lstrip(".")
        if extension in SUPPORTED_WORD_EXTENSIONS:
            has_word = True
        elif extension in SUPPORTED_EXCEL_EXTENSIONS:
            has_excel = True
        else:
            has_unknown = True
    return has_word, has_excel, has_unknown


def route(command: str, files: list[Any]) -> tuple[str, bool]:
    """Select only tools compatible with uploaded files; mixed/unknown inputs fall back safely."""
    if not files:
        return "create", False
    has_word, has_excel, has_unknown = file_kinds(files)
    if has_word and re.search(r"\b(?:use|using|follow|based on|refer(?:ence)?|from)\b.*\btemplate\b", command, re.I | re.S):
        return "template", False
    if has_unknown or (has_word and has_excel) or not (has_word or has_excel):
        return "full", False

    if len(files) > 1 and MERGE_RE.search(command) and not IN_FILE_MERGE_RE.search(command):
        return ("merge_word" if has_word else "merge_excel"), False
    if has_word:
        return "word_all", False
    return "excel_all", False


def tools_for(group: str) -> list[dict[str, Any]]:
    """Return the registered tool schemas for a routed group."""
    names = FULL_TOOL_SET if group == "full" else TOOL_GROUPS[group]
    return [TOOLS_BY_NAME[name] for name in names]


@dataclass
class CommandContext:
    command_id: str
    command: str
    work_dir: Path
    original_ids: list[str]
    working_paths: dict[str, Path]
    messages: list[dict[str, Any]]
    prompt_component_chars: dict[str, int] = field(default_factory=dict)
    initial_message_count: int = 0
    tool_group: str = "full"
    dirty_ids: set[str] = field(default_factory=set)
    generated_names: dict[str, str] = field(default_factory=dict)
    generated_ids: list[str] = field(default_factory=list)
    consumed_ids: set[str] = field(default_factory=set)
    activity: list[dict[str, Any]] = field(default_factory=list)
    usage_records: list[dict[str, Any]] = field(default_factory=list)
    pending_call: dict[str, Any] | None = None


def _model_client():
    provider = os.getenv("SMARTDOC_PROVIDER", "groq").strip().lower()
    if provider == "ollama":
        return _ollama_completion, os.getenv("OLLAMA_MODEL", "qwen3.5:4b")
    if provider != "groq":
        raise RuntimeError("SMARTDOC_PROVIDER must be groq or ollama")
    key = os.getenv("GROQ_API_KEY")
    if not key:
        raise RuntimeError("GROQ_API_KEY is missing. Set it in the environment or .env file.")
    return Groq(api_key=key).chat.completions.create, os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")


def _ollama_completion(*, model, messages, max_completion_tokens, temperature,
                       tools=None, tool_choice="auto") -> ChatCompletion:
    native_messages = []
    for message in messages:
        entry = {"role": message["role"], "content": message.get("content") or ""}
        if message.get("tool_calls"):
            entry["tool_calls"] = [{"function": {
                "name": call["function"]["name"],
                "arguments": json.loads(call["function"]["arguments"]),
            }} for call in message["tool_calls"]]
        if message["role"] == "tool":
            entry["tool_name"] = message["name"]
        native_messages.append(entry)
    payload = {"model": model, "messages": native_messages, "stream": False, "think": False,
               "options": {"num_ctx": 16384, "num_predict": max_completion_tokens,
                           "temperature": temperature}}
    forced = isinstance(tool_choice, dict)
    if forced:
        payload["format"] = tools[0]["function"]["parameters"]
        native_messages[0]["content"] += " Return the completed file content as JSON matching the supplied schema."
    elif tools:
        payload["tools"] = tools
    try:
        response = httpx.post(
            "http://127.0.0.1:11434/api/chat", json=payload,
            timeout=httpx.Timeout(600, connect=5), trust_env=False,
        )
        response.raise_for_status()
    except httpx.ConnectError as exc:
        raise RuntimeError("Cannot reach local Ollama. Start the Ollama app or run ollama serve.") from exc
    except httpx.HTTPStatusError as exc:
        raise RuntimeError(f"Ollama returned {response.status_code}: {response.text}") from exc
    result = response.json()
    reply = result["message"]
    if forced:
        reply = {"content": "", "tool_calls": [{"function": {
            "name": tool_choice["function"]["name"], "arguments": json.loads(reply["content"]),
        }}]}
    calls = [{"id": uuid.uuid4().hex, "type": "function", "function": {
        "name": call["function"]["name"], "arguments": json.dumps(call["function"]["arguments"]),
    }} for call in reply.get("tool_calls") or []]
    prompt, completion = result.get("prompt_eval_count", 0), result.get("eval_count", 0)
    return ChatCompletion.model_validate({
        "id": uuid.uuid4().hex, "model": model, "object": "chat.completion", "created": int(time.time()),
        "choices": [{"index": 0, "finish_reason": "tool_calls" if calls else "stop",
                     "message": {"role": "assistant", "content": reply.get("content"), "tool_calls": calls or None}}],
        "usage": {"prompt_tokens": prompt, "completion_tokens": completion, "total_tokens": prompt + completion},
    })


def _summary_for_ids(session, file_ids: list[str]) -> list[dict[str, Any]]:
    summaries = []
    for file_id in file_ids:
        item = session.files[file_id]
        try:
            summary = word.read_document(item.path) if item.path.suffix.lower() == ".docx" else excel.get_summary(item.path)
        except Exception as exc:
            summary = {"error": str(exc)}
        summaries.append({"file_id": file_id, "file_name": item.display_name,
                          "is_output": item.is_output,
                          "download_url": f"/api/sessions/{session.session_id}/files/{file_id}/download",
                          "summary": summary})
    return summaries


def inspect_session(session) -> dict[str, Any]:
    """Summarize active workbooks and ask Groq for a user-facing orientation."""
    files = _summary_for_ids(session, list(session.current_file_ids))
    file_description = json.dumps(files, ensure_ascii=False, default=str)
    fallback_lines = []
    for entry in files:
        if "paragraph_count" in entry["summary"]:
            summary = entry["summary"]
            fallback_lines.append(
                f"{entry['file_name']}: {summary['paragraph_count']} paragraphs and {len(summary['tables'])} tables."
            )
            continue
        sheets = entry["summary"].get("sheets", [])
        if not sheets:
            fallback_lines.append(f"{entry['file_name']}: {entry['summary'].get('error', 'No readable sheets found.')}")
            continue
        for sheet in sheets:
            fallback_lines.append(
                f"{entry['file_name']} · {sheet['sheet_name']}: {sheet['row_count']} data rows, "
                f"{sheet['column_count']} columns ({', '.join(sheet['headers'])})."
            )
    fallback = "I inspected the uploaded files. " + (" ".join(fallback_lines) or "No files are currently available.")
    overview, ai_enabled, usage_record = fallback, False, None
    inspection_system = (
        "You are SmartDoc. Explain what the uploaded spreadsheets contain using only the supplied "
        "sheet names, used ranges, headers, and sample rows. Be concise and practical. Mention each "
        "file and sheet, describe the columns in plain language, and suggest a few useful tasks. "
        "Do not claim to have changed anything or infer information absent from the samples."
    )
    inspection_user = f"Inspect these uploaded files and explain what you understand:\n{file_description}"
    inspection_messages = [
        {"role": "system", "content": inspection_system},
        {"role": "user", "content": inspection_user},
    ]
    try:
        complete, model = _model_client()
        started = time.perf_counter()
        response = complete(
            model=model,
            messages=inspection_messages,
            temperature=0,
            max_completion_tokens=500,
        )
        usage_record = groq_usage.record_success(
            response=response, model=model, phase="file_inspection", elapsed_seconds=time.perf_counter() - started,
            component_chars={
                "system_instructions": len(inspection_system),
                "file_summaries": len(inspection_user),
            },
        )
        overview = response.choices[0].message.content or fallback
        ai_enabled = bool(response.choices[0].message.content)
    except Exception:
        if "model" in locals():
            usage_record = groq_usage.record_failure(
                model=model, phase="file_inspection", elapsed_seconds=time.perf_counter() - started,
                component_chars={
                    "system_instructions": len(inspection_system),
                    "file_summaries": len(inspection_user),
                },
            )
    if overview:
        session.chat_history.append({"role": "assistant", "content": overview})
        session.chat_history = session.chat_history[-MAX_CHAT_HISTORY_MESSAGES:]
    return {"files": files, "overview": overview, "ai_enabled": ai_enabled,
            "usage": groq_usage.summarize([usage_record]) if usage_record else None}


def start_command(session, command: str, output_format: str | None = None) -> CommandContext:
    if not command.strip():
        raise ValueError("Enter a command first")
    input_ids = [] if output_format else list(session.current_file_ids)
    tool_group = output_format or route(command, [session.files[key] for key in input_ids])[0]
    work_dir = session.directory / f"command-{uuid.uuid4().hex}"
    work_dir.mkdir(parents=True, exist_ok=False)
    working_paths: dict[str, Path] = {}
    for file_id in input_ids:
        item = session.files[file_id]
        path = work_dir / f"{file_id}{item.path.suffix.lower()}"
        shutil.copy2(item.path, path)
        working_paths[file_id] = path
    summaries = _summary_for_ids(session, input_ids)
    history = [] if output_format else list(session.chat_history[-MAX_CHAT_HISTORY_MESSAGES:])
    system_content = (
            "You are SmartDoc, an Excel, CSV, and Word editing agent. Use only the supplied tools; never invent file IDs. "
            "Use the file summaries to resolve references. Ask a short clarifying question if the target or meaning is ambiguous. "
            "When revising a draft you just generated, target the generated file identified in the conversation. "
            "Call tools to make requested changes, then report what changed in plain language. Do not claim a change unless a tool succeeded. "
            "Interpret the user's target precisely before selecting a tool: numeric row ranges such as 'delete rows 5-7' mean delete_rows with rows '5:7'. "
            "Excel-letter ranges such as 'delete columns I-L' mean delete_columns with columns 'I:L'. Letters always refer to columns, even if the user accidentally says 'rows I-L'. "
            "Never pass a letter range as a cell range. "
            "To remove cell contents without deleting rows or columns, use clear_cells with the exact non-empty cell addresses. "
            "For all values across selected rows or columns, use clear_rows or clear_columns; these retain the row/column structure and column headers. "
            "When the user refers to values identified in the conversation, use only those identified cells and request confirmation; never clear an entire column unless explicitly requested. "
            "For removing duplicates from every/all columns or the whole sheet, pass cell_range=null and columns=null; the tool uses the actual used range. "
            "Never estimate the last row from the summary row count. If the user gives an explicit range, keep that range. "
            "For other range operations, cell ranges include the header row. For CSV, use sheet null. "
            "For Word documents, inspect with word_read_document before editing. Word paragraph, table, row, and column indexes are zero-based."
            " When asked to combine Word documents into a new file, use word_merge_documents with their file IDs in the requested order; create a separate output and do not ask whether the user wants a new file. Do not ask the user to identify embedded objects during a merge; the merge tool copies embedded package parts when possible, and any remaining failure must be explained using the concrete tool error."
    )
    if output_format:
        system_content = (f"Create a new .{output_format} file from the user's brief. "
                          "Call the supplied tool once with the complete content. Respect requested length and structure.")
        page_limit = generation.requested_page_limit(command) if output_format == "docx" else None
        if page_limit is not None:
            system_content += (f" The pages array must contain at most {page_limit} entries. "
                               "Combine topics in the same entry. Keep each entry below 1200 characters.")
    elif tool_group == "create":
        system_content = ("You are SmartDoc. To create a Word document or Excel workbook, call the corresponding "
                          "creation tool once with the finished content. Respect the requested topic, format and length. "
                          "No upload is needed. If the user asks to edit a missing source, ask them to upload it. "
                          "For a question or clarification, reply directly.")
    elif tool_group == "template":
        system_content = ("Draft a new Word file using the named template. Use the supplied summaries to select "
                          "its file ID, match exact source text, and preserve its structure. Add styled paragraphs "
                          "if more space is needed. Call draft_from_template once with the finished content. "
                          "Ask which file to use if the template is ambiguous.")
    current_request = f"Current files and summaries:\n{json.dumps(summaries, ensure_ascii=False, default=str)}\n\nUser: {command}"
    messages = [{"role": "system", "content": system_content}, *history,
                {"role": "user", "content": current_request}]
    prompt_component_chars = {
        "system_instructions": len(system_content),
        "prior_chat_history": sum(len(json.dumps(message, ensure_ascii=False, default=str)) for message in history),
        "current_files_and_command": len(current_request),
    }
    context = CommandContext(uuid.uuid4().hex, command, work_dir, input_ids,
                             working_paths, messages, prompt_component_chars=prompt_component_chars,
                             initial_message_count=len(messages), tool_group=tool_group)
    if output_format:
        context.activity.append({"tool": "agent", "detail": f"Generating a new .{output_format} file"})
    else:
        context.activity.append({"tool": "agent", "detail": "Reviewing file summaries and planning the command"})
        context.activity.append({"tool": "router", "detail": f"Selected {tool_group} tools based on uploaded file types"})
    return context


def _json_value(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, dict):
        return {str(k): _json_value(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_value(v) for v in value]
    return str(value)


def _path_for(context: CommandContext, file_id: str) -> Path:
    if file_id not in context.working_paths:
        raise excel.ExcelToolError(f"File ID '{file_id}' is not available in this command")
    return context.working_paths[file_id]


def _execute(name: str, args: dict[str, Any], context: CommandContext) -> Any:
    if name in (*TOOL_GROUPS["create"], "draft_from_template"):
        out_id = uuid.uuid4().hex
        suffix = ".xlsx" if name == "create_workbook" else ".docx"
        path = context.work_dir / f"{out_id}{suffix}"
        if name == "draft_from_template":
            generation.write_from_template(path, _path_for(context, args["file_id"]), args)
        elif suffix == ".docx":
            generation.write_document(path, args, context.command)
        else:
            generation.write_workbook(path, args)
        title = re.sub(r'[^\w -]+', '', args["title"]).strip()[:100] or "Untitled"
        context.generated_names[out_id] = title + suffix
        context.generated_ids.append(out_id)
        context.working_paths[out_id] = path
        context.activity.append({"tool": name, "detail": f"Created {title}{suffix}"})
        return title + suffix
    if name == "word_merge_documents":
        ids = args.get("file_ids") or []
        if len(ids) < 2:
            raise word.WordToolError("Select at least two Word document file IDs to merge")
        paths = [_path_for(context, file_id) for file_id in ids]
        if any(path.suffix.lower() != ".docx" for path in paths):
            raise word.WordToolError("Word document merging accepts .docx files only")
        out_id = uuid.uuid4().hex
        output = context.work_dir / f"{out_id}.docx"
        result = word.merge_documents(paths, output)
        context.working_paths[out_id] = Path(result)
        context.generated_ids.append(out_id)
        context.activity.append({"tool": name, "detail": f"Created a new Word document from {len(ids)} source files"})
        return {"file_id": out_id, "file_name": "merged.docx", "path": result}
    if name == "merge_files":
        ids = args.get("file_ids") or []
        if len(ids) < 2:
            raise excel.ExcelToolError("Select at least two file IDs to merge")
        paths = [_path_for(context, file_id) for file_id in ids]
        out_id = uuid.uuid4().hex
        output = context.work_dir / f"{out_id}.xlsx"
        result = excel.merge_files(paths, args.get("mode", "append"), output, args.get("key"))
        context.consumed_ids.update(ids)
        context.working_paths[out_id] = Path(result)
        context.generated_ids.append(out_id)
        context.activity.append({"tool": name, "detail": f"Created merged workbook from {len(ids)} files"})
        return {"file_id": out_id, "file_name": "merged.xlsx", "path": result}

    file_id = args.pop("file_id", None)
    if not file_id:
        raise excel.ExcelToolError("Tool call requires file_id")
    path = _path_for(context, file_id)
    sheet, cell_range = args.get("sheet"), args.get("cell_range")
    calls = {
        "list_sheets": lambda: excel.list_sheets(path),
        "get_summary": lambda: excel.get_summary(path),
        "read_range": lambda: excel.read_range(path, sheet, cell_range),
        "create_table": lambda: excel.create_table(path, sheet, cell_range, args.get("headers"), args.get("style", "TableStyleMedium2")),
        "modify_table": lambda: excel.modify_table(path, sheet, args["table_name"], args.get("cell_range"), args.get("new_name")),
        "update_cells": lambda: excel.update_cells(path, sheet, args["updates"]),
        "clear_cells": lambda: excel.clear_cells(path, sheet, args["cells"]),
        "clear_columns": lambda: excel.clear_columns(path, sheet, args["columns"]),
        "clear_rows": lambda: excel.clear_rows(path, sheet, args["rows"]),
        "add_column": lambda: excel.add_column(path, sheet, args["header"], args.get("values"), args.get("index")),
        "delete_column": lambda: excel.delete_column(path, sheet, args["column"]),
        "delete_columns": lambda: excel.delete_columns(path, sheet, args["columns"]),
        "add_row": lambda: excel.add_row(path, sheet, args["values"], args.get("index")),
        "delete_row": lambda: excel.delete_row(path, sheet, args["index"]),
        "delete_rows": lambda: excel.delete_rows(path, sheet, args["rows"]),
        "delete_table": lambda: excel.delete_table(path, sheet, args["table_name"]),
        "delete_sheet": lambda: excel.delete_sheet(path, sheet),
        "rename_sheet": lambda: excel.rename_sheet(path, sheet, args["new_name"]),
        "sort": lambda: excel.sort(path, sheet, cell_range, args["column"], args.get("order", "asc")),
        "filter": lambda: excel.filter(path, sheet, cell_range, args["condition"]),
        "remove_duplicates": lambda: excel.remove_duplicates(path, sheet, cell_range, args.get("columns")),
        "apply_formatting": lambda: excel.apply_formatting(path, sheet, cell_range, args.get("bold"), args.get("fill"), args.get("borders"), args.get("number_format")),
        "add_formula": lambda: excel.add_formula(path, sheet, args["cell"], args["formula"]),
        "word_read_document": lambda: word.read_document(path),
        "word_insert_paragraph": lambda: word.insert_paragraph(path, args["text"], args.get("index"), args.get("style")),
        "word_edit_paragraph": lambda: word.edit_paragraph(path, args["index"], args["text"], args.get("style")),
        "word_apply_style": lambda: word.apply_style(path, args["index"], args["style"]),
        "word_delete_paragraph": lambda: word.delete_paragraph(path, args["index"]),
        "word_create_table": lambda: word.create_table(path, args["rows"], args["columns"], args.get("headers"), args.get("style", "Table Grid"), args.get("location")),
        "word_modify_table_cell": lambda: word.modify_table_cell(path, args["table_index"], args["row"], args["column"], args["text"]),
        "word_add_table_row": lambda: word.add_table_row(path, args["table_index"], args.get("values")),
        "word_add_table_column": lambda: word.add_table_column(path, args["table_index"], args.get("header", ""), args.get("values")),
        "word_delete_table": lambda: word.delete_table(path, args["table_index"]),
        "word_delete_table_row": lambda: word.delete_table_row(path, args["table_index"], args["row"]),
        "word_delete_table_column": lambda: word.delete_table_column(path, args["table_index"], args["column"]),
    }
    if name not in calls:
        raise excel.ExcelToolError(f"Tool '{name}' is not available")
    if name == "sort":
        context.activity.append({"tool": name, "detail": f"Sorting rows in {session_filename(context, file_id)}"})
    started = time.perf_counter()
    result = calls[name]()
    if name == "read_range" and isinstance(result, dict):
        values = result.get("values", [])
        if len(values) > 51:
            result["values"] = values[:51]
            result["truncated"] = True
            result["returned_rows"] = 50
    elif name == "filter" and isinstance(result, list) and len(result) > 51:
        result = {"headers": result[0], "rows": result[1:51], "truncated": True,
                  "returned_rows": 50, "total_matching_rows": len(result) - 1}
    if name not in {"list_sheets", "get_summary", "read_range", "filter", "word_read_document"}:
        context.dirty_ids.add(file_id)
    detail = f"{name} on {session_filename(context, file_id)}"
    elapsed = time.perf_counter() - started
    context.activity.append({"tool": name, "detail": detail, "elapsed_seconds": round(elapsed, 3),
                             "result": _json_value(result)})
    return result


def _normalize_tool_request(name: str, args: dict[str, Any], command: str) -> tuple[str, dict[str, Any]]:
    """Enforce unambiguous row/column intent and safe whole-sheet deduplication.

    The LLM proposes a tool call, but this boundary owns spreadsheet-coordinate
    semantics: numbers identify rows and Excel letters identify columns.
    """
    asks_delete = re.search(r"\b(?:delete|remove|clear|erase)\b", command, re.IGNORECASE)
    clears_values = re.search(
        r"\b(?:clear|erase)\b|\b(?:delete|remove)\b(?:\s+(?:all|the|unwanted))*\s+(?:data|values?|contents?|cells?)\b",
        command,
        re.IGNORECASE,
    )
    row_range = re.search(r"\brows?\s+(\d+)\s*(?:-|:|\bto\b)\s*(\d+)\b", command, re.IGNORECASE)
    range_tools = {"clear_cells", "clear_columns", "clear_rows", "delete_column", "delete_columns", "delete_row", "delete_rows"}
    if row_range and clears_values and name in range_tools:
        rows = f"{row_range.group(1)}:{row_range.group(2)}"
        normalized = {key: value for key, value in args.items()
                      if key not in {"column", "columns", "index", "cell_range", "axis", "cells"}}
        return "clear_rows", {**normalized, "rows": rows}
    if row_range and asks_delete and name in range_tools:
        rows = f"{row_range.group(1)}:{row_range.group(2)}"
        normalized = {key: value for key, value in args.items()
                      if key not in {"column", "columns", "index", "cell_range", "axis"}}
        return "delete_rows", {**normalized, "rows": rows}
    axis_range = re.search(r"\b(?:columns?|cols?|rows?)\s+(?:from\s+)?([A-Z]{1,3})\s*(?:-|\bto\b|:)\s*([A-Z]{1,3})\b", command, re.IGNORECASE)
    if axis_range and clears_values and name in range_tools:
        columns = f"{axis_range.group(1).upper()}:{axis_range.group(2).upper()}"
        normalized = {key: value for key, value in args.items()
                      if key not in {"column", "columns", "index", "cell_range", "axis", "cells"}}
        return "clear_columns", {**normalized, "columns": columns}
    if axis_range and asks_delete and name in range_tools:
        columns = f"{axis_range.group(1).upper()}:{axis_range.group(2).upper()}"
        normalized = {key: value for key, value in args.items()
                      if key not in {"column", "index", "cell_range", "axis"}}
        return "delete_columns", {**normalized, "columns": columns}
    # Tool calls occasionally put a letter range in ``index`` while using a
    # row-oriented tool name. Letters always identify spreadsheet columns, so
    # keep the confirmation and resulting operation on the safe column path.
    supplied_range = args.get("columns") or args.get("column") or args.get("index")
    if (clears_values and name in range_tools and isinstance(supplied_range, str)
            and re.fullmatch(r"\s*[A-Z]{1,3}\s*(?:-|:|\bto\b)\s*[A-Z]{1,3}\s*", supplied_range, re.IGNORECASE)):
        columns = re.sub(r"\s*(?:-|\bto\b)\s*", ":", supplied_range.strip(), flags=re.IGNORECASE).upper()
        normalized = {key: value for key, value in args.items()
                      if key not in {"column", "columns", "index", "cell_range", "axis", "cells", "rows"}}
        return "clear_columns", {**normalized, "columns": columns}
    if name in {"clear_columns", "clear_rows", "delete_columns"}:
        return name, args
    if name == "update_cells" and asks_delete:
        updates = args.get("updates")
        if isinstance(updates, dict) and updates and all(value is None for value in updates.values()):
            return "clear_cells", {key: value for key, value in args.items() if key != "updates"} | {
                "cells": list(updates)
            }
    if name == "clear_cells":
        return name, args
    if name != "remove_duplicates":
        return name, args
    has_explicit_range = re.search(
        r"\b\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+\b",
        command,
        re.IGNORECASE,
    )
    full_scope = re.search(
        r"\b(?:every|all)\s+columns?\b|\bentire\s+(?:sheet|worksheet|workbook|file)\b|\bwhole\s+(?:sheet|worksheet|workbook|file)\b",
        command,
        re.IGNORECASE,
    )
    if full_scope and not has_explicit_range:
        return name, {**args, "cell_range": None, "columns": None}
    return name, args


def session_filename(context: CommandContext, file_id: str) -> str:
    return context.working_paths[file_id].name


def _confirm_info(name: str, args: dict[str, Any], session, context: CommandContext) -> dict[str, Any]:
    file_id = args.get("file_id")
    item = session.files.get(file_id)
    label = item.display_name if item else "generated workbook"
    target = {"clear_cells": "Clear cell values", "clear_columns": "Clear column values", "clear_rows": "Clear row values", "delete_column": "Delete column", "delete_columns": "Delete columns", "delete_row": "Delete row", "delete_rows": "Delete rows",
              "delete_table": "Delete table", "delete_sheet": "Delete sheet",
              "remove_duplicates": "Remove duplicate rows", "word_delete_paragraph": "Delete paragraph",
              "word_delete_table": "Delete Word table", "word_delete_table_row": "Delete Word table row",
              "word_delete_table_column": "Delete Word table column"}[name]
    detail = f"SmartDoc is ready to {target.lower()} in {label}."
    payload = {key: value for key, value in args.items() if key != "file_id"}
    if name == "clear_cells":
        detail += f" It will clear {len(args.get('cells', []))} selected cell value(s), while keeping their formatting."
    elif name == "clear_columns":
        detail += f" It will clear the values in column(s) {args.get('columns')}; the headers and formatting will stay."
    elif name == "clear_rows":
        detail += f" It will clear the values in row(s) {args.get('rows')}; the rows and formatting will stay."
    elif name == "delete_column":
        detail += f" It will permanently remove column {args.get('column')}."
    elif name == "delete_columns":
        detail += f" It will permanently remove columns {args.get('columns')}."
    elif name == "delete_row":
        detail += f" It will permanently remove row {args.get('index')}."
    elif name == "delete_rows":
        detail += f" It will permanently remove rows {args.get('rows')}."
    elif name == "delete_table":
        detail += f" It will remove the table named '{args.get('table_name')}' from sheet '{args.get('sheet')}'."
    elif name == "delete_sheet":
        detail += f" It will permanently remove the '{args.get('sheet')}' worksheet."
    elif name == "remove_duplicates":
        detail += f" It will permanently remove duplicate rows from {args.get('cell_range') or 'all data'}, comparing {args.get('columns') or 'all columns'}."
    elif name == "word_delete_paragraph":
        detail += f" It will permanently remove paragraph {args.get('index') + 1}."
    elif name == "word_delete_table":
        detail += f" It will permanently remove table {args.get('table_index') + 1}."
    elif name == "word_delete_table_row":
        detail += f" It will permanently remove row {args.get('row') + 1} from table {args.get('table_index') + 1}."
    elif name == "word_delete_table_column":
        detail += f" It will permanently remove column {args.get('column') + 1} from table {args.get('table_index') + 1}."
    exact = _destructive_preview(name, args, _path_for(context, file_id))
    return {"action": name, "title": target, "detail": detail, "target": payload,
            "file_name": label, "exact": exact}


def _sheet_rows(path: Path, sheet: str | None) -> list[list[Any]]:
    if path.suffix.lower() == ".csv":
        frame = pd.read_csv(path, dtype=object, keep_default_na=False)
        return [list(frame.columns), *frame.values.tolist()]
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in workbook.sheetnames:
            raise excel.ExcelToolError(f"Sheet '{sheet}' not found")
        return [list(row) for row in workbook[sheet].iter_rows(values_only=True)]
    finally:
        workbook.close()


def _destructive_preview(name: str, args: dict[str, Any], path: Path) -> dict[str, Any]:
    if name.startswith("word_"):
        return word.destructive_preview(path, name.removeprefix("word_"), args)
    if name == "clear_cells":
        return _clear_cells_preview(args, path)
    if name == "clear_columns":
        return _column_range_preview(args, path, action="clear")
    if name == "clear_rows":
        return _row_range_preview(args, path, action="clear")
    if name == "delete_columns":
        return _column_range_preview(args, path, action="delete")
    if name == "delete_rows":
        return _row_range_preview(args, path, action="delete")
    sheet = args.get("sheet")
    if name == "remove_duplicates":
        return _duplicate_preview(args, path)
    rows = _sheet_rows(path, sheet)
    if name == "delete_row":
        index = int(args["index"])
        if index < 2 or index > len(rows):
            raise excel.ExcelToolError(f"Row {index} is outside the sheet")
        return {"row_number": index, "values": rows[index - 1]}
    if name == "delete_sheet":
        if path.suffix.lower() == ".csv":
            raise excel.ExcelToolError("CSV files do not contain worksheets")
        workbook = load_workbook(path, read_only=True)
        try:
            if len(workbook.sheetnames) <= 1:
                raise excel.ExcelToolError("Cannot delete the only worksheet in a workbook")
        finally:
            workbook.close()
        return {"sheet": sheet, "row_count": len(rows), "values": rows}
    if name == "delete_column":
        column = args["column"]
        if path.suffix.lower() == ".csv" and isinstance(column, str) and not column.isdigit():
            if column not in rows[0]:
                raise excel.ExcelToolError(f"Column '{column}' not found")
            index = rows[0].index(column)
        else:
            index = int(column) - 1 if str(column).isdigit() else column_index_from_string(str(column).upper()) - 1
        if index < 0 or not rows or index >= len(rows[0]):
            raise excel.ExcelToolError(f"Column '{column}' is outside the sheet")
        return {"column": column, "header": rows[0][index],
                "values": [{"row_number": i + 1, "value": row[index] if index < len(row) else None}
                           for i, row in enumerate(rows)]}
    if name == "delete_table":
        if path.suffix.lower() == ".csv":
            raise excel.ExcelToolError("CSV files do not contain tables")
        workbook = load_workbook(path, read_only=False, data_only=True)
        try:
            ws = workbook[sheet]
            table = ws.tables.get(args["table_name"])
            if table is None:
                raise excel.ExcelToolError(f"Table '{args['table_name']}' not found")
            cell_range = table.ref
        finally:
            workbook.close()
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        return {"table": args["table_name"], "range": cell_range,
                "values": [row[min_col - 1:max_col] for row in rows[min_row - 1:max_row]]}
    return {}


def _column_range_preview(args: dict[str, Any], path: Path, *, action: str) -> dict[str, Any]:
    """Read only selected columns for a fast, exact destructive-action preview."""
    columns = args.get("columns")
    sheet = args.get("sheet")
    if path.suffix.lower() == ".csv":
        frame = pd.read_csv(path, dtype=object, keep_default_na=False)
        start, end = excel._column_bounds(columns, len(frame.columns))
        headers = list(frame.columns[start - 1:end])
        selected = frame.iloc[:, start - 1:end]
        non_empty = [int(((selected.iloc[:, index].notna()) & (selected.iloc[:, index].astype(str) != "")).sum())
                     for index in range(selected.shape[1])]
        sample = selected.head(8).values.tolist()
        row_count = len(frame)
    else:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            if sheet not in workbook.sheetnames:
                raise excel.ExcelToolError(f"Sheet '{sheet}' not found")
            worksheet = workbook[sheet]
            start, end = excel._column_bounds(columns, worksheet.max_column)
            rows = worksheet.iter_rows(min_col=start, max_col=end, values_only=True)
            headers = list(next(rows, ()))
            non_empty = [0] * (end - start + 1)
            sample = []
            row_count = max(worksheet.max_row - 1, 0)
            for row in rows:
                values = list(row)
                for index, value in enumerate(values):
                    if value is not None and str(value).strip():
                        non_empty[index] += 1
                if len(sample) < 8:
                    sample.append(values)
        finally:
            workbook.close()
    labels = [f"{column_index_from_string_to_letter(index)}: {headers[index - start] if index - start < len(headers) else ''}"
              for index in range(start, end + 1)]
    scope = "column values (headers retained)" if action == "clear" else "entire columns"
    return {"scope": scope, "columns": labels, "data_rows": row_count,
            "non_empty_cells": non_empty, "sample_values_first_8_rows": sample,
            "sample_truncated": row_count > len(sample)}


def column_index_from_string_to_letter(index: int) -> str:
    return get_column_letter(index)


def _row_range_preview(args: dict[str, Any], path: Path, *, action: str) -> dict[str, Any]:
    """Preview only selected rows so confirmation stays fast."""
    sheet, requested = args.get("sheet"), args.get("rows")
    if path.suffix.lower() == ".csv":
        frame = pd.read_csv(path, dtype=object, keep_default_na=False)
        start, end = excel._row_bounds(requested, len(frame) + 1)
        headers = list(frame.columns)
        sample = frame.iloc[start - 2:min(end - 1, start - 2 + 8)].values.tolist()
        total = end - start + 1
    else:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            if sheet not in workbook.sheetnames:
                raise excel.ExcelToolError(f"Sheet '{sheet}' not found")
            worksheet = workbook[sheet]
            start, end = excel._row_bounds(requested, worksheet.max_row)
            headers = list(next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ()))
            sample = [list(row) for row in worksheet.iter_rows(
                min_row=start, max_row=min(end, start + 7), values_only=True
            )]
            total = end - start + 1
        finally:
            workbook.close()
    scope = "row values (rows retained)" if action == "clear" else "entire rows"
    return {"scope": scope, "rows": f"{start}:{end}", "row_count": total,
            "headers": headers, "sample_values_first_8_rows": sample,
            "sample_truncated": total > len(sample)}


def _clear_cells_preview(args: dict[str, Any], path: Path) -> dict[str, Any]:
    cells = args.get("cells")
    if not isinstance(cells, list) or not cells or len(cells) > 1000:
        raise excel.ExcelToolError("Clear operation needs between 1 and 1000 cell addresses")
    values: list[dict[str, Any]] = []
    if path.suffix.lower() == ".csv":
        frame = pd.read_csv(path, dtype=object, keep_default_na=False)
        rows = [list(frame.columns), *frame.values.tolist()]
        for address in cells:
            if not isinstance(address, str) or not re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+", address.upper()):
                raise excel.ExcelToolError(f"Invalid cell address '{address}'")
            col, row, end_col, end_row = range_boundaries(address.replace("$", "").upper())
            if row == 1 or row > len(rows) or col != end_col or row != end_row or col > len(frame.columns):
                raise excel.ExcelToolError(f"Cell '{address}' is outside CSV data bounds")
            values.append({"cell": address.replace("$", "").upper(), "value": rows[row - 1][col - 1]})
    else:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = args.get("sheet")
            if sheet not in workbook.sheetnames:
                raise excel.ExcelToolError(f"Sheet '{sheet}' not found")
            worksheet = workbook[sheet]
            for address in cells:
                if not isinstance(address, str) or not re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+", address.upper()):
                    raise excel.ExcelToolError(f"Invalid cell address '{address}'")
                col, row, end_col, end_row = range_boundaries(address.replace("$", "").upper())
                if row == 1 or row > worksheet.max_row or col != end_col or row != end_row or col > worksheet.max_column:
                    raise excel.ExcelToolError(f"Cell '{address}' is outside data bounds")
                values.append({"cell": address.replace("$", "").upper(),
                               "value": worksheet.cell(row, col).value})
        finally:
            workbook.close()
    if len({item["cell"] for item in values}) != len(values):
        raise excel.ExcelToolError("Cell addresses must be unique")
    nonempty = [item for item in values if item["value"] is not None and str(item["value"]).strip()]
    if not nonempty:
        raise excel.ExcelToolError("The selected cells are already empty")
    return {"scope": "selected cell values", "cell_count": len(nonempty), "cells": nonempty}


def _duplicate_preview(args: dict[str, Any], path: Path) -> dict[str, Any]:
    """Scan only the requested cells and retain duplicate rows for confirmation."""
    cell_range = args.get("cell_range")
    requested = args.get("columns")
    if path.suffix.lower() == ".csv":
        frame = pd.read_csv(path, dtype=object, keep_default_na=False)
        min_col, min_row, max_col, max_row = range_boundaries(cell_range) if cell_range else (1, 1, len(frame.columns), len(frame) + 1)
        if min_row != 1 or max_row > len(frame) + 1 or max_col > len(frame.columns):
            raise excel.ExcelToolError("Duplicate range is outside CSV bounds")
        selected_headers = list(frame.columns[min_col - 1:max_col])
        indices = _duplicate_column_indices(requested, selected_headers)
        selected = frame.iloc[min_row:max_row - 1, min_col - 1:max_col]
        duplicate_positions = selected.index[selected.duplicated(subset=[selected_headers[i] for i in indices], keep="first")]
        duplicates = [
            {"row_number": int(index) + 2, "values": frame.iloc[int(index), min_col - 1:max_col].tolist()}
            for index in duplicate_positions
        ]
        return {"duplicate_rows": duplicates, "count": len(duplicates)}

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = args.get("sheet")
        if sheet not in workbook.sheetnames:
            raise excel.ExcelToolError(f"Sheet '{sheet}' not found")
        ws = workbook[sheet]
        if cell_range:
            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        else:
            min_col, min_row = 1, 1
            max_col, max_row = ws.max_column, ws.max_row
        if max_row > ws.max_row or max_col > ws.max_column:
            raise excel.ExcelToolError("Duplicate range is outside sheet bounds")
        header = next(ws.iter_rows(min_row=min_row, max_row=min_row, min_col=min_col,
                                   max_col=max_col, values_only=True), ())
        headers = list(header)
        indices = _duplicate_column_indices(requested, headers)
        seen, duplicates = set(), []
        for row_number, row_values in enumerate(ws.iter_rows(
            min_row=min_row + 1, max_row=max_row, min_col=min_col,
            max_col=max_col, values_only=True,
        ), start=min_row + 1):
            row = list(row_values)
            key = tuple(row[index] if index < len(row) else None for index in indices)
            if key in seen:
                duplicates.append({"row_number": row_number, "values": row})
            else:
                seen.add(key)
        return {"duplicate_rows": duplicates, "count": len(duplicates)}
    finally:
        workbook.close()


def _duplicate_column_indices(requested: list[Any] | None, headers: list[Any]) -> list[int]:
    if not requested:
        return list(range(len(headers)))
    indices = []
    for column in requested:
        if isinstance(column, int) or str(column).isdigit():
            index = int(column) - 1
            if index < 0 or index >= len(headers):
                raise excel.ExcelToolError(f"Duplicate column '{column}' is outside range")
        else:
            try:
                index = [str(value) for value in headers].index(str(column))
            except ValueError as exc:
                raise excel.ExcelToolError(f"Duplicate column '{column}' not found") from exc
        indices.append(index)
    return indices


def _message_dict(message) -> dict[str, Any]:
    if hasattr(message, "model_dump"):
        return message.model_dump(exclude_none=True)
    return message.to_dict()


def _request_component_chars(context: CommandContext) -> dict[str, int]:
    component_chars = dict(context.prompt_component_chars)
    loop_messages = context.messages[context.initial_message_count:]
    component_chars["previous_tool_loop_messages"] = sum(
        len(json.dumps(message, ensure_ascii=False, default=str)) for message in loop_messages
    )
    component_chars["tool_definitions"] = len(json.dumps(
        tools_for(context.tool_group), ensure_ascii=False, separators=(",", ":")
    ))
    return component_chars


def _recover_malformed_clear_call(error: Exception, command: str) -> dict[str, Any] | None:
    """Recover only Groq's narrowly malformed update_cells calls that clear explicit cells."""
    text = str(error)
    if not re.search(r"tool_use_failed|Failed to parse tool call arguments as JSON", text, re.IGNORECASE):
        return None
    if not re.search(r'"name"\s*:\s*"update_cells"', text):
        return None
    if not re.search(r"\b(?:delete|remove|clear|erase)\b", command, re.IGNORECASE):
        return None
    file_match = re.search(r'"file_id"\s*:\s*"([A-Za-z0-9_-]+)"', text)
    sheet_match = re.search(r'"sheet"\s*:\s*(null|"(?:[^"\\]|\\.)*")', text)
    updates_match = re.search(r'"updates"\s*:\s*\{([^{}]*)\}', text)
    if not file_match or not sheet_match or not updates_match:
        return None
    raw_updates = updates_match.group(1)
    pairs = re.findall(r'"(\$?[A-Z]{1,3}\$?\d+)"\s*:\s*null', raw_updates, re.IGNORECASE)
    residue = re.sub(r'"\$?[A-Z]{1,3}\$?\d+"\s*:\s*null', "", raw_updates, flags=re.IGNORECASE)
    if not pairs or re.sub(r"[\s,]", "", residue):
        return None
    cells = list(dict.fromkeys(address.replace("$", "").upper() for address in pairs))
    if len(cells) > 1000:
        return None
    sheet_value = sheet_match.group(1)
    sheet = None if sheet_value == "null" else json.loads(sheet_value)
    return {"file_id": file_match.group(1), "sheet": sheet, "cells": cells}


def run_loop(session, context: CommandContext, *, confirmed: bool = False) -> dict[str, Any]:
    complete, model = _model_client()
    available_tools = tools_for(context.tool_group)
    tool_choice = "auto"
    if context.tool_group in {"docx", "xlsx"}:
        available_tools = deepcopy(available_tools)
        function = available_tools[0]["function"]
        tool_choice = {"type": "function", "function": {"name": function["name"]}}
        if context.tool_group == "docx":
            pages = function["parameters"]["properties"]["pages"]
            limit = generation.requested_page_limit(context.command)
            if limit is not None and limit > 0:
                pages["maxItems"] = limit
    if confirmed and context.pending_call:
        call = context.pending_call
        args = json.loads(call["arguments"])
        result = _execute(call["name"], args, context)
        context.messages.append({"role": "tool", "tool_call_id": call["id"],
                                 "name": call["name"], "content": json.dumps(_json_value(result), default=str)})
        context.pending_call = None
    for turn in range(1, MAX_TOOL_TURNS + 1):
        component_chars = _request_component_chars(context)
        started = time.perf_counter()
        try:
            response = complete(model=model, messages=context.messages, tools=available_tools,
                                tool_choice=tool_choice, temperature=0, max_completion_tokens=6000 if context.tool_group in {"create", "template", "docx", "xlsx"} else 1200)
        except Exception as exc:
            context.usage_records.append(groq_usage.record_failure(
                model=model, phase="command", component_chars=component_chars, turn=turn,
                elapsed_seconds=time.perf_counter() - started,
                operation_id=context.command_id,
            ))
            args = _recover_malformed_clear_call(exc, context.command)
            if args is None:
                raise
            call_id = "recovered-clear-" + uuid.uuid4().hex
            context.messages.append({"role": "assistant", "tool_calls": [{
                "id": call_id, "type": "function", "function": {
                    "name": "clear_cells", "arguments": json.dumps(args)
                }
            }]})
            confirmation = _confirm_info("clear_cells", args, session, context)
            context.pending_call = {"id": call_id, "name": "clear_cells",
                                    "arguments": json.dumps(args)}
            context.activity.append({"tool": "clear_cells",
                                     "detail": "Recovered the exact selected cells from a malformed tool request; awaiting confirmation"})
            return {"status": "confirmation_required", "confirmation_id": context.command_id,
                    "confirmation": confirmation, "activity": context.activity,
                    "usage": groq_usage.summarize(context.usage_records)}
        context.usage_records.append(groq_usage.record_success(
            response=response, model=model, phase="command",
            elapsed_seconds=time.perf_counter() - started,
            component_chars=component_chars, turn=turn, operation_id=context.command_id,
        ))
        message = response.choices[0].message
        serialized = _message_dict(message)
        context.messages.append(serialized)
        calls = serialized.get("tool_calls") or []
        if not calls:
            if context.tool_group in {"docx", "xlsx"}:
                raise RuntimeError("The model returned text without generating the requested file. Please try again.")
            return _commit_result(session, context, message.content or "Done.")
        if len(calls) > 1:
            for call in calls:
                context.messages.append({"role": "tool", "tool_call_id": call["id"],
                                         "name": call["function"]["name"],
                                         "content": json.dumps({"error": "Call one tool at a time"})})
            continue
        call = calls[0]
        function = call["function"]
        name = function["name"]
        if name not in {entry["function"]["name"] for entry in tools_for(context.tool_group)}:
            context.messages.append({"role": "tool", "tool_call_id": call["id"], "name": name,
                                     "content": json.dumps({"error": f"Unknown tool: {name}"})})
            continue
        try:
            args = json.loads(function.get("arguments") or "{}")
            name, args = _normalize_tool_request(name, args, context.command)
            if name in DESTRUCTIVE_TOOLS:
                if name == "remove_duplicates":
                    context.activity.append({"tool": name, "detail": "Checking the selected rows for duplicate values before confirmation"})
                confirmation = _confirm_info(name, args, session, context)
                context.pending_call = {"id": call["id"], "name": name, "arguments": json.dumps(args)}
                return {"status": "confirmation_required", "confirmation_id": context.command_id,
                        "confirmation": confirmation,
                        "activity": context.activity,
                        "usage": groq_usage.summarize(context.usage_records)}
            result = _execute(name, args, context)
            if name in (*TOOL_GROUPS["create"], "draft_from_template"):
                break
            content = _json_value(result)
            if isinstance(content, dict) and content.get("truncated"):
                content["note"] = "Output limited to the first 50 matching/data rows plus header. Request a smaller range for more."
            context.messages.append({"role": "tool", "tool_call_id": call["id"], "name": name,
                                     "content": json.dumps(content, default=str)})
        except Exception as exc:
            context.activity.append({"tool": name, "detail": f"Tool error: {exc}"})
            context.messages.append({"role": "tool", "tool_call_id": call["id"], "name": name,
                                     "content": json.dumps({"error": str(exc)})})
    else:
        raise RuntimeError("The agent reached its tool-call limit before completing the command")
    return _commit_result(session, context, f"Created {result}.")


def _commit_result(session, context: CommandContext, summary: str) -> dict[str, Any]:
    previous = list(session.current_file_ids)
    created: list[str] = []
    replacements: dict[str, str] = {}
    generated_outputs: list[str] = []
    for old_id in context.dirty_ids:
        if old_id in context.generated_ids:
            continue
        old = session.files[old_id]
        path = context.working_paths[old_id]
        new_id = uuid.uuid4().hex
        output_path = session.directory / f"{new_id}{old.path.suffix.lower()}"
        shutil.copy2(path, output_path)
        output = type(old)(new_id, f"{Path(old.display_name).stem}_edited{old.path.suffix.lower()}",
                           output_path, output_path.stat().st_size, True)
        session.files[new_id] = output
        replacements[old_id] = new_id
        created.append(new_id)
    for generated_id in context.generated_ids:
        temp_path = context.working_paths[generated_id]
        new_id = uuid.uuid4().hex
        suffix = temp_path.suffix.lower()
        output_path = session.directory / f"{new_id}{suffix}"
        shutil.copy2(temp_path, output_path)
        output = SessionFile(new_id, context.generated_names.get(generated_id, f"merged{suffix}"),
                             output_path, output_path.stat().st_size, True)
        session.files[new_id] = output
        created.append(new_id)
        generated_outputs.append(new_id)
    current = [replacements.get(file_id, file_id) for file_id in previous
               if file_id not in context.consumed_ids]
    current.extend(generated_outputs)
    session.current_file_ids = current
    session.chat_history.extend([
        {"role": "user", "content": context.command},
        {"role": "assistant", "content": summary},
    ])
    session.chat_history = session.chat_history[-MAX_CHAT_HISTORY_MESSAGES:]
    if created:
        session.undo_history.append({"previous_file_ids": previous, "created_file_ids": created})
    result_files = [{"file_id": file_id, "file_name": session.files[file_id].display_name,
                     "download_url": f"/api/sessions/{session.session_id}/files/{file_id}/download"}
                    for file_id in created]
    preview = _summary_for_ids(session, session.current_file_ids)
    shutil.rmtree(context.work_dir, ignore_errors=True)
    return {"status": "complete", "summary": summary, "files": result_files,
            "current_files": preview, "activity": context.activity,
            "can_undo": bool(session.undo_history),
            "usage": groq_usage.summarize(context.usage_records)}


def abort_context(context: CommandContext) -> None:
    shutil.rmtree(context.work_dir, ignore_errors=True)
