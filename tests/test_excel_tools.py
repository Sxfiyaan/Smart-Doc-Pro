from __future__ import annotations

from copy import copy
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from backend.tools import excel_tools as et


@pytest.fixture
def workbook_path(tmp_path: Path) -> Path:
    path = tmp_path / "book.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Employees"
    ws.append(["Name", "Department", "Age"])
    ws.append(["Asha", "Sales", 30])
    ws.append(["Ben", "IT", 25])
    ws.append(["Asha", "Sales", 30])
    wb.create_sheet("Notes").append(["Keep"])
    wb.save(path)
    return path


def test_inspection_and_range_validation(workbook_path: Path):
    assert et.list_sheets(workbook_path) == ["Employees", "Notes"]
    summary = et.get_summary(workbook_path)
    assert summary["sheets"][0]["headers"] == ["Name", "Department", "Age"]
    assert summary["sheets"][0]["row_count"] == 3
    assert et.read_range(workbook_path, "Employees", "A1:B2")["values"] == [
        ["Name", "Department"], ["Asha", "Sales"]
    ]
    with pytest.raises(et.ExcelToolError, match="not found"):
        et.read_range(workbook_path, "Missing", "A1:A2")
    with pytest.raises(et.ExcelToolError, match="bounds"):
        et.read_range(workbook_path, "Employees", "A1:Z20")


def test_create_modify_and_delete_table(workbook_path: Path):
    table_name = et.create_table(workbook_path, "Employees", "A1:C4", style="TableStyleMedium4")
    wb = load_workbook(workbook_path)
    assert table_name in wb["Employees"].tables
    wb.close()
    result = et.modify_table(workbook_path, "Employees", table_name, "A1:C3", "StaffTable")
    assert result == {"table_name": "StaffTable", "range": "A1:C3"}
    et.delete_table(workbook_path, "Employees", "StaffTable")
    wb = load_workbook(workbook_path)
    assert not wb["Employees"].tables
    wb.close()


def test_table_name_validation_and_column_guard(workbook_path: Path):
    table_name = et.create_table(workbook_path, "Employees", "A1:C4")
    with pytest.raises(et.ExcelToolError, match="overlaps"):
        et.create_table(workbook_path, "Employees", "A1:B3")
    with pytest.raises(et.ExcelToolError, match="inside an Excel table"):
        et.add_column(workbook_path, "Employees", "Region", ["N", "S", "E"], index=2)
    et.modify_table(workbook_path, "Employees", table_name, new_name="Renamed")
    with pytest.raises(et.ExcelToolError, match="not found"):
        et.modify_table(workbook_path, "Employees", table_name, new_name="Another")


def test_update_cells_rows_columns_sheet_and_formatting(workbook_path: Path):
    assert et.update_cells(workbook_path, "Employees", {"A2": "Anya"}) == 1
    assert et.add_column(workbook_path, "Employees", "Active", [True, False, True]) == 4
    assert et.add_row(workbook_path, "Employees", ["Cia", "HR", 28, True]) == 5
    assert et.apply_formatting(workbook_path, "Employees", "A1:D1", bold=True,
                               fill="FFE080", borders=True) == 4
    assert et.add_formula(workbook_path, "Employees", "E2", "=C2*2") == "E2"
    wb = load_workbook(workbook_path, data_only=False)
    ws = wb["Employees"]
    assert ws["A2"].value == "Anya" and ws["D2"].value is True
    assert ws["A1"].font.bold is True and ws["A1"].fill.fgColor.rgb.endswith("FFE080")
    assert ws["E2"].value == "=C2*2"
    wb.close()
    et.delete_column(workbook_path, "Employees", "D")
    et.delete_row(workbook_path, "Employees", 5)
    et.rename_sheet(workbook_path, "Notes", "Archive")
    et.delete_sheet(workbook_path, "Archive")
    wb = load_workbook(workbook_path)
    assert wb.sheetnames == ["Employees"]
    assert wb["Employees"].max_column == 4
    wb.close()


def test_delete_row_range_preserves_header_for_xlsx_and_csv(tmp_path: Path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["ID", "Name"])
    for row in [[1, "A"], [2, "B"], [3, "C"], [4, "D"]]:
        sheet.append(row)
    xlsx = tmp_path / "rows.xlsx"
    workbook.save(xlsx)
    assert et.delete_rows(xlsx, "Sheet", "3-4") == (3, 4)
    wb = load_workbook(xlsx, data_only=True)
    assert list(wb.active.values) == [("ID", "Name"), (1, "A"), (4, "D")]
    wb.close()

    csv = tmp_path / "rows.csv"
    pd.DataFrame([[1, "A"], [2, "B"], [3, "C"], [4, "D"]], columns=["ID", "Name"]).to_csv(csv, index=False)
    assert et.delete_rows(csv, None, "3:4") == (3, 4)
    assert pd.read_csv(csv).values.tolist() == [[1, "A"], [4, "D"]]
    with pytest.raises(et.ExcelToolError, match="outside data rows"):
        et.delete_rows(csv, None, "1:2")


def test_clear_cells_clears_contents_and_preserves_formatting(workbook_path: Path):
    wb = load_workbook(workbook_path)
    font = copy(wb["Employees"]["A2"].font)
    font.bold = True
    wb["Employees"]["A2"].font = font
    original_style = wb["Employees"]["A2"]._style
    wb.save(workbook_path)
    wb.close()

    assert et.clear_cells(workbook_path, "Employees", ["A2", "B2"]) == 2
    wb = load_workbook(workbook_path)
    assert wb["Employees"]["A2"].value is None
    assert wb["Employees"]["B2"].value is None
    assert wb["Employees"]["A2"]._style == original_style
    wb.close()
    with pytest.raises(et.ExcelToolError, match="data bounds"):
        et.clear_cells(workbook_path, "Employees", ["A1"])


def test_clear_column_and_row_ranges_preserve_structure(tmp_path: Path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["ID", "Name", "Score"])
    sheet.append([1, "A", 10])
    sheet.append([2, "B", 20])
    sheet.append([3, "C", 30])
    sheet.append([4, "D", 40])
    xlsx = tmp_path / "clear-ranges.xlsx"
    workbook.save(xlsx)

    assert et.clear_columns(xlsx, "Data", "B:C") == {"columns": "2:3", "cleared_cells": 8}
    wb = load_workbook(xlsx, data_only=True)
    assert list(wb["Data"].values) == [("ID", "Name", "Score"), (1, None, None), (2, None, None), (3, None, None), (4, None, None)]
    wb.close()
    assert et.clear_rows(xlsx, "Data", "3:4") == {"rows": "3:4", "cleared_cells": 2}
    wb = load_workbook(xlsx, data_only=True)
    assert list(wb["Data"].values) == [("ID", "Name", "Score"), (1, None, None), (None, None, None), (None, None, None), (4, None, None)]
    wb.close()

    csv = tmp_path / "clear-ranges.csv"
    pd.DataFrame([[1, "A"], [2, "B"], [3, "C"]], columns=["ID", "Name"]).to_csv(csv, index=False)
    assert et.clear_columns(csv, None, "B") == {"columns": "2:2", "cleared_cells": 3}
    assert et.clear_rows(csv, None, "3:4") == {"rows": "3:4", "cleared_cells": 2}
    assert pd.read_csv(csv, keep_default_na=False).values.tolist() == [["1", ""], ["", ""], ["", ""]]


def test_sort_filter_and_remove_duplicates(workbook_path: Path):
    assert et.sort(workbook_path, "Employees", "A1:C4", "C", "asc") == 3
    matches = et.filter(workbook_path, "Employees", "A1:C4",
                        {"column": "Department", "operator": "equals", "value": "Sales"})
    assert len(matches) == 3
    assert et.remove_duplicates(workbook_path, "Employees", "A1:C4", ["Name", "Department"]) == 1
    summary = et.get_summary(workbook_path)
    assert summary["sheets"][0]["row_count"] == 2


def test_csv_data_operations_and_summary(tmp_path: Path):
    path = tmp_path / "people.csv"
    pd.DataFrame([["Asha", "Sales", 30], ["Ben", "IT", 25], ["Asha", "Sales", 30]],
                 columns=["Name", "Department", "Age"]).to_csv(path, index=False)
    assert et.list_sheets(path) == ["Sheet1"]
    assert et.get_summary(path)["sheets"][0]["row_count"] == 3
    assert et.read_range(path, None, "A1:B2")["values"] == [["Name", "Department"], ["Asha", "Sales"]]
    assert len(et.filter(path, None, "A1:C4", {"column": "Age", "operator": "greater_than", "value": 26})) == 3
    assert et.remove_duplicates(path, None, columns=["Name"]) == 1
    et.update_cells(path, None, {"A2": "Bea"})
    assert et.sort(path, None, "A1:C3", "Age", "desc") == 2
    frame = pd.read_csv(path)
    assert frame.iloc[0]["Name"] == "Bea"
    assert "Bea" in frame["Name"].values
    assert et.add_row(path, None, ["Cal", "HR", 28]) == 4
    assert et.add_column(path, None, "Active", [True, False, True]) == 4
    et.delete_row(path, None, 4)
    et.delete_column(path, None, "Active")
    assert len(pd.read_csv(path)) == 2
    with pytest.raises(et.ExcelToolError, match="do not support"):
        et.create_table(path, "Sheet1", "A1:C2")


def test_merge_append_and_join(tmp_path: Path):
    first = tmp_path / "first.csv"
    second = tmp_path / "second.csv"
    pd.DataFrame({"ID": [1, 2], "Name": ["A", "B"]}).to_csv(first, index=False)
    pd.DataFrame({"ID": [2, 3], "Score": [90, 80]}).to_csv(second, index=False)
    appended = Path(et.merge_files([first, second], "append", tmp_path / "append.xlsx"))
    joined = Path(et.merge_files([first, second], "join_on_key", tmp_path / "join.xlsx", key="ID"))
    assert pd.read_excel(appended).shape == (4, 3)
    joined_frame = pd.read_excel(joined).sort_values("ID")
    assert joined_frame["ID"].tolist() == [1, 2, 3]
    assert joined_frame.loc[joined_frame.ID == 2, "Score"].iloc[0] == 90


def test_missing_file_and_invalid_arguments(workbook_path: Path):
    with pytest.raises(et.ExcelToolError, match="File not found"):
        et.get_summary(workbook_path.parent / "missing.xlsx")
    with pytest.raises(et.ExcelToolError, match="start with '='"):
        et.add_formula(workbook_path, "Employees", "A2", "SUM(A1:A4)")
    with pytest.raises(et.ExcelToolError, match="only worksheet"):
        path = workbook_path.parent / "single.xlsx"
        wb = Workbook(); wb.save(path)
        et.delete_sheet(path, "Sheet")
