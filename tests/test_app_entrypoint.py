from fastapi.testclient import TestClient

from app.main import app as legacy_app
from backend.main import app as backend_app


def test_legacy_entrypoint_exposes_the_full_smartdoc_api():
    assert legacy_app is backend_app
    with TestClient(legacy_app) as client:
        assert client.get("/").status_code == 200
        assert client.post("/api/sessions").status_code == 201
