from __future__ import annotations

import json
import time
from io import BytesIO
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook
from docx import Document as WordDocument

from backend import agent, main


class FakeMessage:
    def __init__(self, content=None, tool_calls=None):
        self.content = content
        self.tool_calls = tool_calls

    def model_dump(self, exclude_none=True):
        result = {"role": "assistant"}
        if self.content is not None:
            result["content"] = self.content
        if self.tool_calls is not None:
            result["tool_calls"] = self.tool_calls
        return result


class FakeGroq:
    def __init__(self, messages):
        self.messages = list(messages)
        self.chat = SimpleNamespace(completions=self)

    def create(self, **kwargs):
        message = self.messages.pop(0)
        if isinstance(message, Exception):
            raise message
        return SimpleNamespace(choices=[SimpleNamespace(message=message)])


def xlsx_bytes() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Employees"
    sheet.append(["Name", "Age"])
    sheet.append(["Asha", 30])
    sheet.append(["Ben", 25])
    workbook.create_sheet("Notes").append(["Private note"])
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def docx_bytes() -> bytes:
    document = WordDocument()
    document.add_heading("Project update", level=1)
    document.add_paragraph("Draft status")
    table = document.add_table(rows=2, cols=2, style="Table Grid")
    table.cell(0, 0).text, table.cell(0, 1).text = "Task", "Owner"
    table.cell(1, 0).text, table.cell(1, 1).text = "Review", "Asha"
    output = BytesIO()
    document.save(output)
    return output.getvalue()


@pytest.fixture
def client():
    with TestClient(main.app) as test_client:
        yield test_client


def upload_workbook(client: TestClient) -> tuple[str, str, bytes]:
    session_id = client.post("/api/sessions").json()["session_id"]
    original = xlsx_bytes()
    response = client.post(f"/api/sessions/{session_id}/files",
                           files=[("files", ("staff.xlsx", original,
                                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))])
    return session_id, response.json()["files"][0]["file_id"], original


def wait_for_command(client: TestClient, session_id: str, response):
    assert response.status_code in {200, 202}, response.text
    result = response.json()
    if result.get("status") != "running":
        return result
    for _ in range(200):
        result = client.get(
            f"/api/sessions/{session_id}/commands/{result['command_id']}"
        ).json()
        if result.get("status") != "running":
            return result
        time.sleep(0.01)
    raise AssertionError("Agent command did not finish")


def tool_call(call_id: str, name: str, args: dict) -> FakeMessage:
    return FakeMessage(tool_calls=[{"id": call_id, "type": "function", "function": {
        "name": name, "arguments": json.dumps(args)}}])


def use_fake_groq(monkeypatch, *messages):
    fake = FakeGroq(messages)
    monkeypatch.setattr(agent, "_groq_client", lambda: (fake, "test-model"))


def test_clear_range_language_uses_the_correct_axis():
    name, args = agent._normalize_tool_request(
        "delete_rows", {"file_id": "file", "sheet": "Data", "rows": "5:8"},
        "delete all values in rows 5-8",
    )
    assert (name, args["rows"]) == ("clear_rows", "5:8")
    name, args = agent._normalize_tool_request(
        "delete_row", {"file_id": "file", "sheet": "Data", "index": "I-J"},
        "clear rows from I-J",
    )
    assert (name, args["columns"]) == ("clear_columns", "I:J")
    name, args = agent._normalize_tool_request(
        "delete_columns", {"file_id": "file", "sheet": "Data", "columns": "I:J"},
        "clear rows I-J",
    )
    assert (name, args["columns"]) == ("clear_columns", "I:J")
    name, args = agent._normalize_tool_request(
        "delete_rows", {"file_id": "file", "sheet": "Data", "rows": "I-K"},
        "delete rows I-K because they contain unwanted noisy data",
    )
    assert (name, args["columns"]) == ("delete_columns", "I:K")


def test_explicit_delete_row_and_column_commands_select_the_right_tools():
    """Groq may suggest a nearby tool; the server must preserve coordinate intent."""
    name, args = agent._normalize_tool_request(
        "delete_columns", {"file_id": "file", "sheet": "Data", "columns": "5-7"},
        "delete rows 5-7 (it is unwanted data)",
    )
    assert (name, args["rows"]) == ("delete_rows", "5:7")

    name, args = agent._normalize_tool_request(
        "delete_rows", {"file_id": "file", "sheet": "Data", "rows": "I-L"},
        "delete column I-L",
    )
    assert (name, args["columns"]) == ("delete_columns", "I:L")


def test_tool_command_creates_downloadable_copy_and_undoes(client, monkeypatch):
    session_id, file_id, original = upload_workbook(client)
    use_fake_groq(monkeypatch,
                  tool_call("call-1", "update_cells", {"file_id": file_id, "sheet": "Employees",
                                                          "updates": {"A2": "Anya"}}),
                  FakeMessage(content="Updated the first employee name."))
    response = client.post(f"/api/sessions/{session_id}/commands",
                           json={"command": "Change the first employee name to Anya"})
    result = wait_for_command(client, session_id, response)
    assert result["status"] == "complete" and result["can_undo"] is True
    assert result["files"][0]["file_name"] == "staff_edited.xlsx"
    current_files = client.get(f"/api/sessions/{session_id}/files").json()["files"]
    edited_listing = next(item for item in current_files if item["is_current"])
    assert edited_listing["is_output"] is True
    assert client.get(
        f"/api/sessions/{session_id}/files/{edited_listing['file_id']}/download"
    ).status_code == 200
    output = client.get(f"/api/sessions/{session_id}/files/{result['files'][0]['file_id']}/download")
    changed = load_workbook(BytesIO(output.content), data_only=True)
    assert changed["Employees"]["A2"].value == "Anya"
    changed.close()
    source = client.get(f"/api/sessions/{session_id}/files/{file_id}/download")
    assert source.content == original

    undone = client.post(f"/api/sessions/{session_id}/undo")
    assert undone.status_code == 200 and undone.json()["can_undo"] is False
    assert undone.json()["current_files"][0]["file_name"] == "staff.xlsx"
    assert client.get(f"/api/sessions/{session_id}/files/{result['files'][0]['file_id']}/download").status_code == 404


def test_destructive_call_requires_confirmation_and_can_be_cancelled(client, monkeypatch):
    session_id, file_id, original = upload_workbook(client)
    use_fake_groq(monkeypatch, tool_call("call-del", "delete_sheet",
                                         {"file_id": file_id, "sheet": "Notes"}))
    response = client.post(f"/api/sessions/{session_id}/commands", json={"command": "Delete Notes sheet"})
    pending = wait_for_command(client, session_id, response)
    assert pending["status"] == "confirmation_required"
    assert "Notes" in pending["confirmation"]["detail"]
    assert pending["confirmation"]["exact"]["values"] == [["Private note"]]
    assert client.get(f"/api/sessions/{session_id}/files/{file_id}/download").content == original
    cancelled = client.post(f"/api/sessions/{session_id}/commands/confirm",
                            json={"confirmation_id": pending["confirmation_id"], "confirmed": False})
    assert cancelled.status_code == 200 and cancelled.json()["status"] == "cancelled"
    assert client.get(f"/api/sessions/{session_id}/files/{file_id}/download").content == original


def test_confirmed_destructive_call_creates_copy_and_undo_restores(client, monkeypatch):
    session_id, file_id, original = upload_workbook(client)
    use_fake_groq(monkeypatch,
                  tool_call("call-del", "delete_sheet", {"file_id": file_id, "sheet": "Notes"}),
                  FakeMessage(content="Deleted the Notes sheet."))
    pending_response = client.post(f"/api/sessions/{session_id}/commands", json={"command": "Delete Notes sheet"})
    pending = wait_for_command(client, session_id, pending_response)
    confirmed_response = client.post(f"/api/sessions/{session_id}/commands/confirm",
                                     json={"confirmation_id": pending["confirmation_id"], "confirmed": True})
    result = wait_for_command(client, session_id, confirmed_response)
    assert result["status"] == "complete"
    edited = client.get(f"/api/sessions/{session_id}/files/{result['files'][0]['file_id']}/download")
    workbook = load_workbook(BytesIO(edited.content))
    assert workbook.sheetnames == ["Employees"]
    workbook.close()
    assert client.get(f"/api/sessions/{session_id}/files/{file_id}/download").content == original
    assert client.post(f"/api/sessions/{session_id}/undo").status_code == 200


def test_merge_command_uses_output_as_current_and_undo_restores_sources(client, monkeypatch):
    session_id = client.post("/api/sessions").json()["session_id"]
    response = client.post(f"/api/sessions/{session_id}/files", files=[
        ("files", ("one.csv", "Name,Age\nAsha,30\n".encode(), "text/csv")),
        ("files", ("two.csv", "Name,Age\nBen,25\n".encode(), "text/csv")),
    ])
    ids = [item["file_id"] for item in response.json()["files"]]
    use_fake_groq(monkeypatch,
                  tool_call("merge-1", "merge_files", {"file_ids": ids, "mode": "append"}),
                  FakeMessage(content="Merged the two files."))
    merged = client.post(f"/api/sessions/{session_id}/commands", json={"command": "Merge all files"})
    result = wait_for_command(client, session_id, merged)
    assert len(result["current_files"]) == 1
    assert result["current_files"][0]["file_name"] == "merged.xlsx"
    output_id = result["files"][0]["file_id"]
    workbook = load_workbook(BytesIO(client.get(
        f"/api/sessions/{session_id}/files/{output_id}/download").content), data_only=True)
    assert workbook["Merged"]["A2"].value == "Asha"
    assert workbook["Merged"]["A3"].value == "Ben"
    workbook.close()
    undone = client.post(f"/api/sessions/{session_id}/undo")
    assert [item["file_name"] for item in undone.json()["current_files"]] == ["one.csv", "two.csv"]


def test_all_column_dedupe_uses_actual_sheet_end_even_if_model_range_is_short(client, monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "List - 2"
    sheet.append(["SI. No.", "Superset ID", "Name", "Degree", "Department"])
    source_rows = []
    for number in range(1, 24):
        row = [number, 8000000 + number, f"Student {number}", "B.Tech.", "CSE"]
        sheet.append(row)
        source_rows.append(row)
    duplicate_rows = [source_rows[index - 1] for index in (14, 17, 18, 19)]
    for row in duplicate_rows:
        sheet.append(row)
    buffer = BytesIO()
    workbook.save(buffer)
    original = buffer.getvalue()

    session_id = client.post("/api/sessions").json()["session_id"]
    upload = client.post(f"/api/sessions/{session_id}/files", files=[
        ("files", ("students.xlsx", original,
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    ])
    file_id = upload.json()["files"][0]["file_id"]
    use_fake_groq(monkeypatch,
                  tool_call("dedupe", "remove_duplicates", {
                      "file_id": file_id, "sheet": "List - 2",
                      "cell_range": "A1:E27", "columns": None,
                  }),
                  FakeMessage(content="Removed four duplicate rows across the full sheet."))
    initial = client.post(f"/api/sessions/{session_id}/commands",
                          json={"command": "delete duplicate from every column"})
    pending = wait_for_command(client, session_id, initial)
    assert pending["status"] == "confirmation_required"
    preview = pending["confirmation"]["exact"]["duplicate_rows"]
    assert len(preview) == 4
    assert preview[-1]["row_number"] == 28
    assert preview[-1]["values"][0] == 19

    confirmed = client.post(f"/api/sessions/{session_id}/commands/confirm", json={
        "confirmation_id": pending["confirmation_id"], "confirmed": True,
    })
    result = wait_for_command(client, session_id, confirmed)
    output_id = result["files"][0]["file_id"]
    output = client.get(f"/api/sessions/{session_id}/files/{output_id}/download")
    edited = load_workbook(BytesIO(output.content), data_only=True)
    edited_sheet = edited["List - 2"]
    assert edited_sheet.max_row == 24
    assert sum(1 for row in edited_sheet.iter_rows(min_row=2, values_only=True)
               if row[0] == 19) == 1
    edited.close()
    assert client.get(f"/api/sessions/{session_id}/files/{file_id}/download").content == original


def test_sort_then_duplicate_confirmation_reports_progress_and_exact_rows(client, monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Students"
    sheet.append(["Superset ID", "Name"])
    sheet.append([300, "C"])
    sheet.append([100, "A"])
    sheet.append([100, "A"])
    buffer = BytesIO()
    workbook.save(buffer)
    session_id = client.post("/api/sessions").json()["session_id"]
    uploaded = client.post(f"/api/sessions/{session_id}/files", files=[
        ("files", ("students.xlsx", buffer.getvalue(),
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    ]).json()["files"][0]
    use_fake_groq(monkeypatch,
                  tool_call("sort", "sort", {"file_id": uploaded["file_id"], "sheet": "Students",
                                                "cell_range": "A1:B4", "column": "A", "order": "asc"}),
                  tool_call("dedupe", "remove_duplicates", {"file_id": uploaded["file_id"],
                                                               "sheet": "Students", "cell_range": None,
                                                               "columns": None}))
    started = client.post(f"/api/sessions/{session_id}/commands",
                          json={"command": "Sort by Superset ID then remove duplicate rows across all columns"})
    pending = wait_for_command(client, session_id, started)
    assert pending["status"] == "confirmation_required"
    tools = [event["tool"] for event in pending["activity"]]
    assert "sort" in tools
    assert "remove_duplicates" in tools
    assert pending["confirmation"]["exact"]["duplicate_rows"] == [
        {"row_number": 3, "values": [100, "A"]}
    ]


def test_delete_rows_g_to_j_maps_to_confirmed_column_range(client, monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["A", "B", "C", "D", "E", "F", "G value", "H value", "I value", "J value", "K value"])
    sheet.append(list(range(11)))
    buffer = BytesIO()
    workbook.save(buffer)
    session_id = client.post("/api/sessions").json()["session_id"]
    uploaded = client.post(f"/api/sessions/{session_id}/files", files=[
        ("files", ("wide.xlsx", buffer.getvalue(),
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    ]).json()["files"][0]

    # Simulate a model mistake: it routes the G:J request to the single-row tool.
    use_fake_groq(monkeypatch,
                  tool_call("wrong-axis", "delete_row", {"file_id": uploaded["file_id"],
                                                            "sheet": "Data", "index": "G-J"}),
                  FakeMessage(content="Removed columns G through J."))
    started = client.post(f"/api/sessions/{session_id}/commands",
                          json={"command": "delete rows G-J because they contain unwanted noisy data"})
    pending = wait_for_command(client, session_id, started)
    assert pending["status"] == "confirmation_required"
    assert pending["confirmation"]["target"]["columns"] == "G:J"
    assert pending["confirmation"]["exact"]["columns"] == [
        "G: G value", "H: H value", "I: I value", "J: J value"
    ]
    assert pending["confirmation"]["exact"]["scope"] == "entire columns"
    assert pending["confirmation"]["exact"]["data_rows"] == 1
    assert pending["confirmation"]["exact"]["sample_values_first_8_rows"] == [[6, 7, 8, 9]]

    confirmed = client.post(f"/api/sessions/{session_id}/commands/confirm", json={
        "confirmation_id": pending["confirmation_id"], "confirmed": True,
    })
    result = wait_for_command(client, session_id, confirmed)
    output_id = result["files"][0]["file_id"]
    output = client.get(f"/api/sessions/{session_id}/files/{output_id}/download")
    edited = load_workbook(BytesIO(output.content), data_only=True)
    assert [cell.value for cell in edited["Data"][1]] == ["A", "B", "C", "D", "E", "F", "K value"]
    assert [cell.value for cell in edited["Data"][2]] == [0, 1, 2, 3, 4, 5, 10]
    edited.close()


def test_clear_column_and_row_ranges_keep_structure(client, monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row in [["ID", "Name", "Score"], [1, "A", 10], [2, "B", 20], [3, "C", 30], [4, "D", 40]]:
        sheet.append(row)
    buffer = BytesIO()
    workbook.save(buffer)
    session_id = client.post("/api/sessions").json()["session_id"]
    uploaded = client.post(f"/api/sessions/{session_id}/files", files=[
        ("files", ("clear.xlsx", buffer.getvalue(),
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    ]).json()["files"][0]

    # Recover a model that chooses row deletion for an instruction to remove row values.
    use_fake_groq(monkeypatch,
                  tool_call("wrong-operation", "delete_rows", {"file_id": uploaded["file_id"],
                                                                   "sheet": "Data", "rows": "3:4"}),
                  FakeMessage(content="Cleared the requested values."))
    started = client.post(f"/api/sessions/{session_id}/commands",
                          json={"command": "delete all values in rows 3-4"})
    pending = wait_for_command(client, session_id, started)
    assert pending["status"] == "confirmation_required"
    assert pending["confirmation"]["title"] == "Clear row values"
    assert pending["confirmation"]["exact"]["scope"] == "row values (rows retained)"

    confirmed = client.post(f"/api/sessions/{session_id}/commands/confirm", json={
        "confirmation_id": pending["confirmation_id"], "confirmed": True,
    })
    result = wait_for_command(client, session_id, confirmed)
    output = client.get(f"/api/sessions/{session_id}/files/{result['files'][0]['file_id']}/download")
    edited = load_workbook(BytesIO(output.content), data_only=True)
    assert list(edited["Data"].values) == [
        ("ID", "Name", "Score"), (1, "A", 10), (None, None, None), (None, None, None), (4, "D", 40)
    ]
    edited.close()


def test_numeric_row_range_requires_exact_confirmation_and_preserves_header(client, monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row in [["ID", "Name"], [1, "A"], [2, "B"], [3, "C"], [4, "D"]]:
        sheet.append(row)
    buffer = BytesIO()
    workbook.save(buffer)
    session_id = client.post("/api/sessions").json()["session_id"]
    uploaded = client.post(f"/api/sessions/{session_id}/files", files=[
        ("files", ("rows.xlsx", buffer.getvalue(),
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    ]).json()["files"][0]

    # Recover a model choosing the single-row tool for a numeric span.
    use_fake_groq(monkeypatch,
                  tool_call("wrong-axis", "delete_row", {"file_id": uploaded["file_id"],
                                                            "sheet": "Data", "index": "3-4"}),
                  FakeMessage(content="Deleted rows 3 through 4."))
    started = client.post(f"/api/sessions/{session_id}/commands",
                          json={"command": "delete rows 3-4"})
    pending = wait_for_command(client, session_id, started)
    assert pending["status"] == "confirmation_required"
    assert pending["confirmation"]["target"]["rows"] == "3:4"
    assert pending["confirmation"]["exact"]["rows"] == "3:4"
    assert pending["confirmation"]["exact"]["sample_values_first_8_rows"] == [[2, "B"], [3, "C"]]

    confirmed = client.post(f"/api/sessions/{session_id}/commands/confirm", json={
        "confirmation_id": pending["confirmation_id"], "confirmed": True,
    })
    result = wait_for_command(client, session_id, confirmed)
    output_id = result["files"][0]["file_id"]
    output = client.get(f"/api/sessions/{session_id}/files/{output_id}/download")
    edited = load_workbook(BytesIO(output.content), data_only=True)
    assert list(edited["Data"].values) == [("ID", "Name"), (1, "A"), (4, "D")]
    edited.close()


def test_word_command_creates_an_edited_copy_and_requires_confirmation_for_deletion(client, monkeypatch):
    session_id = client.post("/api/sessions").json()["session_id"]
    uploaded = client.post(f"/api/sessions/{session_id}/files", files=[
        ("files", ("update.docx", docx_bytes(),
                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
    ]).json()["files"][0]
    inspected = client.post(f"/api/sessions/{session_id}/inspect")
    assert inspected.status_code == 200
    assert inspected.json()["files"][0]["summary"]["paragraph_count"] == 2

    use_fake_groq(monkeypatch,
                  tool_call("word-edit", "word_edit_paragraph", {"file_id": uploaded["file_id"],
                                                                    "index": 1, "text": "Approved status"}),
                  FakeMessage(content="Updated the status paragraph."))
    started = client.post(f"/api/sessions/{session_id}/commands", json={"command": "Change the status to approved"})
    result = wait_for_command(client, session_id, started)
    assert result["status"] == "complete"
    output = client.get(f"/api/sessions/{session_id}/files/{result['files'][0]['file_id']}/download")
    edited = WordDocument(BytesIO(output.content))
    assert edited.paragraphs[1].text == "Approved status"

    use_fake_groq(monkeypatch,
                  tool_call("word-delete", "word_delete_table", {"file_id": result["files"][0]["file_id"],
                                                                      "table_index": 0}),
                  FakeMessage(content="Removed the table."))
    started = client.post(f"/api/sessions/{session_id}/commands", json={"command": "Delete the table"})
    pending = wait_for_command(client, session_id, started)
    assert pending["status"] == "confirmation_required"
    assert pending["confirmation"]["title"] == "Delete Word table"
    assert pending["confirmation"]["exact"]["sample_rows"] == [["Task", "Owner"], ["Review", "Asha"]]


def test_malformed_groq_null_updates_recover_as_confirmed_cell_clear(client, monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "List - 2"
    sheet["G11"] = "remove me"
    sheet["G12"] = "also remove"
    sheet["H6"] = 42
    buffer = BytesIO()
    workbook.save(buffer)
    session_id = client.post("/api/sessions").json()["session_id"]
    uploaded = client.post(f"/api/sessions/{session_id}/files", files=[
        ("files", ("messy.xlsx", buffer.getvalue(),
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    ]).json()["files"][0]
    malformed = RuntimeError(
        "Error code: 400 - {'error': {'code': 'tool_use_failed', "
        "'message': 'Failed to parse tool call arguments as JSON', "
        "'failed_generation': '{\"name\": \"update_cells\", \"arguments\": "
        "{\"file_id\":\"" + uploaded["file_id"] + "\",\"sheet\":\"List - 2\", "
        "\"updates\":{\"G11\":null,\"G12\":null,\"H6\":null}}\"}'}}"
    )
    use_fake_groq(monkeypatch, malformed, FakeMessage(content="Cleared the selected values."))
    started = client.post(f"/api/sessions/{session_id}/commands",
                          json={"command": "delete all the unwanted data values you mentioned"})
    pending = wait_for_command(client, session_id, started)
    assert pending["status"] == "confirmation_required"
    confirmation = pending["confirmation"]
    assert confirmation["title"] == "Clear cell values"
    assert confirmation["exact"]["cells"] == [
        {"cell": "G11", "value": "remove me"},
        {"cell": "G12", "value": "also remove"},
        {"cell": "H6", "value": 42},
    ]

    confirmed = client.post(f"/api/sessions/{session_id}/commands/confirm", json={
        "confirmation_id": pending["confirmation_id"], "confirmed": True,
    })
    result = wait_for_command(client, session_id, confirmed)
    assert result["status"] == "complete"
    output = client.get(f"/api/sessions/{session_id}/files/{result['files'][0]['file_id']}/download")
    edited = load_workbook(BytesIO(output.content), data_only=True)
    assert edited["List - 2"]["G11"].value is None
    assert edited["List - 2"]["G12"].value is None
    assert edited["List - 2"]["H6"].value is None
    edited.close()
