from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

from backend.tools import word_tools as wt


@pytest.fixture
def document_path(tmp_path: Path) -> Path:
    path = tmp_path / "brief.docx"
    document = Document()
    document.add_heading("Project brief", level=1)
    document.add_paragraph("First draft")
    table = document.add_table(rows=2, cols=2, style="Table Grid")
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Status"
    table.cell(1, 0).text = "Asha"
    table.cell(1, 1).text = "Draft"
    document.save(path)
    return path


def test_read_and_edit_paragraphs(document_path: Path):
    summary = wt.read_document(document_path)
    assert summary["paragraph_count"] == 2
    assert summary["headings"][0]["text"] == "Project brief"
    assert summary["tables"][0]["sample_rows"] == [["Name", "Status"], ["Asha", "Draft"]]
    assert wt.insert_paragraph(document_path, "Purpose", 1, "Heading 2") == 1
    wt.edit_paragraph(document_path, 2, "Updated draft")
    wt.apply_style(document_path, 2, "Normal")
    wt.delete_paragraph(document_path, 2)
    summary = wt.read_document(document_path)
    assert [paragraph["text"] for paragraph in summary["paragraphs"]] == ["Project brief", "Purpose"]


def test_create_modify_and_delete_word_tables(document_path: Path):
    index = wt.create_table(document_path, 2, 2, ["Task", "Owner"])
    assert index == 1
    wt.modify_table_cell(document_path, 1, 1, 0, "Review")
    assert wt.add_table_row(document_path, 1, ["Ship", "Ben"]) == 2
    assert wt.add_table_column(document_path, 1, "State", ["Open", "Ready"]) == 2
    summary = wt.read_document(document_path)
    assert summary["tables"][1]["sample_rows"] == [
        ["Task", "Owner", "State"], ["Review", "", "Open"], ["Ship", "Ben", "Ready"]
    ]
    wt.delete_table_row(document_path, 1, 2)
    wt.delete_table_column(document_path, 1, 1)
    wt.delete_table(document_path, 1)
    assert len(wt.read_document(document_path)["tables"]) == 1


def test_word_validation_and_destructive_previews(document_path: Path):
    paragraph_preview = wt.destructive_preview(document_path, "delete_paragraph", {"index": 1})
    assert paragraph_preview["paragraph_text"] == "First draft"
    table_preview = wt.destructive_preview(document_path, "delete_table_column", {"table_index": 0, "column": 1})
    assert table_preview["values"] == ["Status", "Draft"]
    with pytest.raises(wt.WordToolError, match="outside table bounds"):
        wt.modify_table_cell(document_path, 0, 9, 0, "No")
    with pytest.raises(wt.WordToolError, match="only column"):
        single = Document(); single.add_table(rows=1, cols=1); single.save(document_path)
        wt.delete_table_column(document_path, 0, 0)
