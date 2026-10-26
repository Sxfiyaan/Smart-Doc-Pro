from __future__ import annotations

import io
import time

import pytest
from fastapi.testclient import TestClient

from backend import main


@pytest.fixture
def client():
    with TestClient(main.app) as test_client:
        yield test_client


def create_session(client: TestClient) -> str:
    response = client.post("/api/sessions")
    assert response.status_code == 201
    return response.json()["session_id"]


def test_upload_list_download_and_delete(client: TestClient):
    session_id = create_session(client)
    response = client.post(
        f"/api/sessions/{session_id}/files",
        files=[("files", ("report.txt", b"hello", "text/plain"))],
    )
    assert response.status_code == 201
    uploaded = response.json()["files"][0]
    assert uploaded["file_name"] == "report.txt"
    assert uploaded["size"] == 5
    assert client.get(f"/api/sessions/{session_id}/files").json()["files"] == [uploaded]

    download = client.get(
        f"/api/sessions/{session_id}/files/{uploaded['file_id']}/download"
    )
    assert download.status_code == 200
    assert download.content == b"hello"
    assert "report.txt" in download.headers["content-disposition"]

    directory = main.sessions[session_id].directory
    assert client.delete(f"/api/sessions/{session_id}").status_code == 204
    assert not directory.exists()
    assert client.get(f"/api/sessions/{session_id}/files").status_code == 404


def test_filename_is_sanitized_and_storage_is_isolated(client: TestClient):
    session_a = create_session(client)
    session_b = create_session(client)
    response = client.post(
        f"/api/sessions/{session_a}/files",
        files=[("files", ("../../payroll 😀.txt", b"data", "text/plain"))],
    )
    assert response.status_code == 201
    item = response.json()["files"][0]
    assert "/" not in item["file_name"] and "\\" not in item["file_name"]
    assert client.get(f"/api/sessions/{session_b}/files").json()["files"] == []


def test_rejects_unsupported_extension_without_storing_files(client: TestClient):
    session_id = create_session(client)
    response = client.post(
        f"/api/sessions/{session_id}/files",
        files=[
            ("files", ("ok.txt", b"valid", "text/plain")),
            ("files", ("program.exe", b"no", "application/octet-stream")),
        ],
    )
    assert response.status_code == 415
    assert client.get(f"/api/sessions/{session_id}/files").json()["files"] == []
    assert list(main.sessions[session_id].directory.iterdir()) == []


def test_rejects_oversize_file_and_removes_partial_upload(client: TestClient):
    session_id = create_session(client)
    too_large = io.BytesIO(b"x" * (main.MAX_FILE_BYTES + 1))
    response = client.post(
        f"/api/sessions/{session_id}/files",
        files=[("files", ("large.txt", too_large, "text/plain"))],
    )
    assert response.status_code == 413
    assert client.get(f"/api/sessions/{session_id}/files").json()["files"] == []
    assert list(main.sessions[session_id].directory.iterdir()) == []


def test_rejects_batch_over_aggregate_limit_atomically(client: TestClient):
    session_id = create_session(client)
    size = 17 * 1024 * 1024
    response = client.post(
        f"/api/sessions/{session_id}/files",
        files=[
            ("files", (f"{index}.txt", io.BytesIO(bytes([index]) * size), "text/plain"))
            for index in range(6)
        ],
    )
    assert response.status_code == 413
    assert client.get(f"/api/sessions/{session_id}/files").json()["files"] == []
    assert list(main.sessions[session_id].directory.iterdir()) == []


def test_expired_session_is_removed(client: TestClient):
    session_id = create_session(client)
    session = main.sessions[session_id]
    session.touched_at = time.monotonic() - main.SESSION_TTL_SECONDS - 1
    directory = session.directory
    assert client.get(f"/api/sessions/{session_id}/files").status_code == 404
    assert not directory.exists()


def test_home_page_and_health_endpoint(client: TestClient):
    assert client.get("/").status_code == 200
    assert "Drop files here" in client.get("/").text
    assert client.get("/api/health").json() == {"status": "ok"}
