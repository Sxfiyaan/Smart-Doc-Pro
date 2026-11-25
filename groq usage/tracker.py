"""Privacy-conscious Groq request accounting for SmartDoc.

Provider-reported token totals are exact. Per-component prompt estimates use
character counts because the Groq tokenizer for the configured model is not
installed locally; the tracker never stores prompts, file values, or commands.
"""

from __future__ import annotations

import json
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

LOG_PATH = Path(__file__).resolve().parent / "requests.jsonl"
_LOG_LOCK = threading.Lock()
_CHARS_PER_TOKEN = 4


def _usage_field(usage: Any, name: str) -> int | None:
    value = getattr(usage, name, None) if usage is not None else None
    return int(value) if isinstance(value, (int, float)) else None


def _nested_usage_field(usage: Any, group: str, name: str) -> int | None:
    details = getattr(usage, group, None) if usage is not None else None
    value = getattr(details, name, None) if details is not None else None
    return int(value) if isinstance(value, (int, float)) else None


def _write(record: dict[str, Any]) -> None:
    """Append metadata only; logging must never make a file command fail."""
    try:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with _LOG_LOCK, LOG_PATH.open("a", encoding="utf-8") as log:
            log.write(json.dumps(record, ensure_ascii=False, separators=(",", ":")) + "\n")
    except OSError:
        pass


def record_success(*, response: Any, model: str, phase: str, elapsed_seconds: float,
                   component_chars: dict[str, int], turn: int | None = None,
                   operation_id: str | None = None) -> dict[str, Any]:
    usage = getattr(response, "usage", None)
    prompt_tokens = _usage_field(usage, "prompt_tokens")
    completion_tokens = _usage_field(usage, "completion_tokens")
    total_tokens = _usage_field(usage, "total_tokens")
    if total_tokens is None and prompt_tokens is not None and completion_tokens is not None:
        total_tokens = prompt_tokens + completion_tokens

    estimates = {name: max(0, round(max(0, count) / _CHARS_PER_TOKEN))
                 for name, count in component_chars.items()}
    request_id = uuid.uuid4().hex
    record = {
        "request_id": request_id,
        "operation_id": operation_id or request_id,
        "time_utc": datetime.now(timezone.utc).isoformat(),
        "phase": phase,
        "turn": turn,
        "model": model,
        "status": "success",
        "elapsed_seconds": elapsed_seconds,
        "prompt_tokens": prompt_tokens,
        "completion_tokens": completion_tokens,
        "total_tokens": total_tokens,
        "cached_prompt_tokens": _nested_usage_field(usage, "prompt_tokens_details", "cached_tokens"),
        "reasoning_tokens": _nested_usage_field(usage, "completion_tokens_details", "reasoning_tokens"),
        "prompt_component_chars": component_chars,
        "approx_prompt_tokens_by_component": estimates,
        "estimate_method": "characters divided by 4; approximate, not provider-tokenized",
    }
    _write(record)
    return record


def record_failure(*, model: str, phase: str, component_chars: dict[str, int], elapsed_seconds: float,
                   turn: int | None = None, operation_id: str | None = None) -> dict[str, Any]:
    request_id = uuid.uuid4().hex
    record = {
        "request_id": request_id,
        "operation_id": operation_id or request_id,
        "time_utc": datetime.now(timezone.utc).isoformat(),
        "phase": phase,
        "turn": turn,
        "model": model,
        "status": "failed",
        "elapsed_seconds": elapsed_seconds,
        "prompt_tokens": None,
        "completion_tokens": None,
        "total_tokens": None,
        "prompt_component_chars": component_chars,
        "approx_prompt_tokens_by_component": {
            name: max(0, round(max(0, count) / _CHARS_PER_TOKEN))
            for name, count in component_chars.items()
        },
        "estimate_method": "characters divided by 4; approximate, not provider-tokenized",
    }
    _write(record)
    return record


def summarize(records: list[dict[str, Any]]) -> dict[str, Any]:
    """Aggregate exact provider totals and approximate prompt component sizes."""
    categories: dict[str, int] = {}
    for record in records:
        for name, count in record.get("approx_prompt_tokens_by_component", {}).items():
            categories[name] = categories.get(name, 0) + count
    return {
        "request_count": len(records),
        "successful_request_count": sum(record.get("status") == "success" for record in records),
        "prompt_tokens": sum(record.get("prompt_tokens") or 0 for record in records),
        "completion_tokens": sum(record.get("completion_tokens") or 0 for record in records),
        "total_tokens": sum(record.get("total_tokens") or 0 for record in records),
        "approx_prompt_tokens_by_component": categories,
        "estimate_method": "Provider totals are exact; component token counts are rough character-based estimates.",
        "requests": records,
    }
