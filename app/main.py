"""Backward-compatible ASGI entry point for SmartDoc.

Use either ``uvicorn backend.main:app`` or the earlier project command
``uvicorn app.main:app``. Both names expose the same API and routes.
"""

from backend.main import app

__all__ = ["app"]
