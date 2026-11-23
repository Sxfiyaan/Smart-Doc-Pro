"""Excel and CSV operations with explicit file, sheet, and range validation."""

from __future__ import annotations

import re
import uuid
from copy import copy
from pathlib import Path
from typing import Any, Mapping, Sequence

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import get_column_letter, range_boundaries

SUPPORTED_EXTENSIONS = {".xlsx", ".csv"}
DEFAULT_SHEET = "Sheet1"


class ExcelToolError(ValueError):
    """An actionable validation or operation error returned to the agent."""


def _path(file_path: str | Path) -> Path:
    path = Path(file_path)
    if not path.is_file():
        raise ExcelToolError(f"File not found: {path}")
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ExcelToolError("Supported data files are .xlsx and .csv")
    return path


def _is_csv(path: Path) -> bool:
    return path.suffix.lower() == ".csv"


def _workbook(path: Path):
    if _is_csv(path):
        raise ExcelToolError("This operation requires .xlsx; CSV files do not support sheets, tables, or formatting")
    try:
        return load_workbook(path)
    except Exception as exc:
        raise ExcelToolError(f"Could not open workbook: {exc}") from exc


def _sheet_name(path: Path, sheet: str | None) -> str:
    if _is_csv(path):
        if sheet not in (None, DEFAULT_SHEET, path.stem):
            raise ExcelToolError(f"CSV files have one data sheet named '{DEFAULT_SHEET}'")
        return DEFAULT_SHEET
    if not sheet:
        raise ExcelToolError("A sheet name is required for .xlsx files")
    return sheet


def _load_csv(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path, dtype=object, keep_default_na=False)
    except pd.errors.EmptyDataError:
        return pd.DataFrame()
    except Exception as exc:
        raise ExcelToolError(f"Could not read CSV: {exc}") from exc


def _save_csv(frame: pd.DataFrame, path: Path) -> None:
    frame.to_csv(path, index=False)


def _get_sheet(workbook, sheet_name: str):
    if sheet_name not in workbook.sheetnames:
        raise ExcelToolError(f"Sheet '{sheet_name}' not found. Available sheets: {', '.join(workbook.sheetnames)}")
    return workbook[sheet_name]


def _range(sheet, cell_range: str, *, allow_empty: bool = False) -> tuple[int, int, int, int]:
    if not isinstance(cell_range, str) or not re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?", cell_range.upper()):
        raise ExcelToolError(f"Invalid cell range '{cell_range}'")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range.replace("$", "").upper())
    except Exception as exc:
        raise ExcelToolError(f"Invalid cell range '{cell_range}'") from exc
    if not all((min_col, min_row, max_col, max_row)):
        raise ExcelToolError(f"Invalid cell range '{cell_range}'")
    if max_row > sheet.max_row or max_col > sheet.max_column:
        if not allow_empty:
            raise ExcelToolError(
                f"Range '{cell_range}' exceeds sheet bounds ({sheet.max_row} rows, {sheet.max_column} columns)"
            )
    return min_col, min_row, max_col, max_row


def _parse_range(cell_range: str) -> tuple[int, int, int, int]:
    if not isinstance(cell_range, str) or not re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?", cell_range.upper()):
        raise ExcelToolError(f"Invalid cell range '{cell_range}'")
    min_col, min_row, max_col, max_row = range_boundaries(cell_range.replace("$", "").upper())
    if not all((min_col, min_row, max_col, max_row)):
        raise ExcelToolError(f"Invalid cell range '{cell_range}'")
    return min_col, min_row, max_col, max_row


def _values(ws, min_col: int, min_row: int, max_col: int, max_row: int) -> list[list[Any]]:
    return [list(row) for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                               min_col=min_col, max_col=max_col,
                                               values_only=True)]


def _header_frame(frame: pd.DataFrame, file_name: str) -> dict[str, Any]:
    return {
        "file_name": file_name,
        "sheets": [{"sheet_name": DEFAULT_SHEET, "row_count": len(frame),
                    "column_count": len(frame.columns), "headers": [str(c) for c in frame.columns],
                    "sample_rows": frame.head(5).values.tolist(),
                    "used_range": f"A1:{get_column_letter(len(frame.columns))}{len(frame) + 1}" if len(frame.columns) else "A1",
                    "tables": []}],
    }


def list_sheets(file_path: str | Path) -> list[str]:
    path = _path(file_path)
    if _is_csv(path):
        return [DEFAULT_SHEET]
    wb = _workbook(path)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def read_range(file_path: str | Path, sheet: str | None, cell_range: str) -> dict[str, Any]:
    path = _path(file_path)
    sheet_name = _sheet_name(path, sheet)
    if _is_csv(path):
        frame = _load_csv(path)
        # CSV coordinates include its header as row 1, matching Excel's A1 convention.
        min_col, min_row, max_col, max_row = _parse_range(cell_range)
        if not min_col or not min_row or max_row > len(frame) + 1 or max_col > len(frame.columns):
            raise ExcelToolError(f"Range '{cell_range}' is outside CSV bounds")
        rows = [list(frame.columns)] + frame.values.tolist()
        return {"sheet": sheet_name, "range": cell_range,
                "values": [row[min_col - 1:max_col] for row in rows[min_row - 1:max_row]]}
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, sheet_name)
        bounds = _range(ws, cell_range)
        return {"sheet": sheet_name, "range": cell_range, "values": _values(ws, *bounds)}
    finally:
        wb.close()


def get_summary(file_path: str | Path) -> dict[str, Any]:
    path = _path(file_path)
    if _is_csv(path):
        return _header_frame(_load_csv(path), path.name)
    wb = _workbook(path)
    try:
        sheets = []
        for ws in wb.worksheets:
            headers = [str(ws.cell(1, col).value) if ws.cell(1, col).value is not None
                       else f"column_{col}" for col in range(1, ws.max_column + 1)]
            samples = _values(ws, 1, 2, ws.max_column, min(ws.max_row, 6)) if ws.max_row > 1 else []
            sheets.append({"sheet_name": ws.title, "row_count": max(ws.max_row - 1, 0),
                           "column_count": ws.max_column, "headers": headers,
                           "sample_rows": samples, "used_range": ws.calculate_dimension(),
                           "tables": list(ws.tables.keys())})
        return {"file_name": path.name, "sheets": sheets}
    finally:
        wb.close()


def create_table(file_path: str | Path, sheet: str, cell_range: str,
                 headers: Sequence[str] | None = None, style: str = "TableStyleMedium2") -> str:
    path = _path(file_path)
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, sheet)
        min_col, min_row, max_col, max_row = _range(ws, cell_range)
        if max_row <= min_row:
            raise ExcelToolError("A table range must include a header row and at least one data row")
        if headers is not None:
            if len(headers) != max_col - min_col + 1:
                raise ExcelToolError("Header count must match the number of columns in the range")
            for col, header in enumerate(headers, start=min_col):
                ws.cell(min_row, col).value = header
        names = [ws.cell(min_row, col).value for col in range(min_col, max_col + 1)]
        if any(name is None or str(name).strip() == "" for name in names):
            raise ExcelToolError("Every table column must have a non-empty header")
        if len({str(name).casefold() for name in names}) != len(names):
            raise ExcelToolError("Table headers must be unique")
        for existing_name in ws.tables:
            left, top, right, bottom = range_boundaries(ws.tables[existing_name].ref)
            if not (max_col < left or min_col > right or max_row < top or min_row > bottom):
                raise ExcelToolError(f"Table range overlaps existing table '{existing_name}'")
        table_name = "Table_" + re.sub(r"\W+", "_", ws.title).strip("_")[:20] + "_" + uuid.uuid4().hex[:6]
        table = Table(displayName=table_name, ref=cell_range.replace("$", "").upper())
        table.tableStyleInfo = TableStyleInfo(name=style, showFirstColumn=False,
                                              showLastColumn=False, showRowStripes=True,
                                              showColumnStripes=False)
        ws.add_table(table)
        wb.save(path)
        return table_name
    finally:
        wb.close()


def _find_table(ws, table_name: str):
    if table_name not in ws.tables:
        raise ExcelToolError(f"Table '{table_name}' not found on sheet '{ws.title}'")
    return ws.tables[table_name]


def modify_table(file_path: str | Path, sheet: str, table_name: str,
                 cell_range: str | None = None, new_name: str | None = None) -> dict[str, str]:
    path = _path(file_path)
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, sheet)
        table = _find_table(ws, table_name)
        if cell_range:
            bounds = _range(ws, cell_range)
            if bounds[2] - bounds[0] + 1 != len(table.tableColumns):
                raise ExcelToolError("Modified table range must keep the same number of columns")
            table.ref = cell_range.replace("$", "").upper()
        if new_name:
            if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", new_name):
                raise ExcelToolError("Table names must start with a letter or underscore and contain only letters, numbers, dots, or underscores")
            if any(other_ws.tables[other_key].displayName.casefold() == new_name.casefold()
                   for other_ws in wb.worksheets for other_key in other_ws.tables
                   if other_key != table_name):
                raise ExcelToolError(f"Table name '{new_name}' is already in use")
            del ws.tables[table_name]
            table.name = new_name
            table.displayName = new_name
            ws.add_table(table)
        wb.save(path)
        return {"table_name": table.displayName, "range": table.ref}
    finally:
        wb.close()


def update_cells(file_path: str | Path, sheet: str | None,
                 updates: Mapping[str, Any]) -> int:
    path = _path(file_path)
    if not updates:
        return 0
    if _is_csv(path):
        frame = _load_csv(path)
        for address, value in updates.items():
            min_col, min_row, max_col, max_row = range_boundaries(address.upper())
            if not min_col or not min_row or min_col != max_col or min_row != max_row:
                raise ExcelToolError(f"Expected one cell address, got '{address}'")
            if min_row == 1 or min_row > len(frame) + 1 or min_col > len(frame.columns):
                raise ExcelToolError(f"Cell '{address}' is outside CSV data bounds")
            frame.iat[min_row - 2, min_col - 1] = value
        _save_csv(frame, path)
        return len(updates)
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        for address, value in updates.items():
            if not re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+", address.upper()):
                raise ExcelToolError(f"Expected one cell address, got '{address}'")
            cell = ws[address.replace("$", "").upper()]
            if cell.row > ws.max_row or cell.column > ws.max_column:
                raise ExcelToolError(f"Cell '{address}' is outside sheet bounds")
        for address, value in updates.items():
            ws[address.replace("$", "").upper()] = value
        wb.save(path)
        return len(updates)
    finally:
        wb.close()


def clear_cells(file_path: str | Path, sheet: str | None,
                cells: Sequence[str]) -> int:
    """Clear explicit data cells while preserving their formatting."""
    if not cells:
        raise ExcelToolError("Select at least one cell to clear")
    if len(cells) > 1000:
        raise ExcelToolError("A single clear operation is limited to 1000 cells")
    normalized = [str(cell).replace("$", "").upper() for cell in cells]
    if len(set(normalized)) != len(normalized):
        raise ExcelToolError("Cell addresses must be unique")
    for address in normalized:
        if not re.fullmatch(r"[A-Z]{1,3}\d+", address):
            raise ExcelToolError(f"Invalid cell address '{address}'")
        min_col, min_row, max_col, max_row = range_boundaries(address)
        if min_col != max_col or min_row != max_row or min_row == 1:
            raise ExcelToolError(f"Cell '{address}' is outside data bounds")
    path = _path(file_path)
    if not _is_csv(path):
        wb = _workbook(path)
        try:
            ws = _get_sheet(wb, _sheet_name(path, sheet))
            for address in normalized:
                col, row, _, _ = range_boundaries(address)
                if row > ws.max_row or col > ws.max_column:
                    raise ExcelToolError(f"Cell '{address}' is outside data bounds")
        finally:
            wb.close()
    updates = {address: None for address in normalized}
    changed = update_cells(file_path, sheet, updates)
    if _is_csv(path):
        frame = _load_csv(path)
        remaining = [address for address in normalized
                     if frame.iat[range_boundaries(address)[1] - 2,
                                  range_boundaries(address)[0] - 1] not in (None, "")]
    else:
        wb = _workbook(path)
        try:
            ws = _get_sheet(wb, _sheet_name(path, sheet))
            remaining = [address for address in normalized if ws[address].value is not None]
        finally:
            wb.close()
    if remaining:
        raise ExcelToolError(f"Could not verify that these cells were cleared: {', '.join(remaining)}")
    return changed


def clear_columns(file_path: str | Path, sheet: str | None,
                  columns: str | int) -> dict[str, int | str]:
    """Clear data values in contiguous columns while retaining headers and formatting."""
    path = _path(file_path)
    if _is_csv(path):
        frame = _load_csv(path)
        start, end = _column_bounds(columns, len(frame.columns))
        values = frame.iloc[:, start - 1:end]
        cleared = int(((values.notna()) & (values.astype(str) != "")).sum().sum())
        frame.iloc[:, start - 1:end] = ""
        _save_csv(frame, path)
        checked = _load_csv(path).iloc[:, start - 1:end]
        if (checked.astype(str) != "").any().any():
            raise ExcelToolError("Could not verify that the selected column values were cleared")
        return {"columns": f"{start}:{end}", "cleared_cells": cleared}
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        start, end = _column_bounds(columns, ws.max_column)
        cleared = 0
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=start, max_col=end):
            for cell in row:
                if cell.value is not None:
                    cleared += 1
                cell.value = None
        wb.save(path)
    finally:
        wb.close()
    check = _workbook(path)
    try:
        ws = _get_sheet(check, _sheet_name(path, sheet))
        if any(cell.value is not None for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                                               min_col=start, max_col=end) for cell in row):
            raise ExcelToolError("Could not verify that the selected column values were cleared")
    finally:
        check.close()
    return {"columns": f"{start}:{end}", "cleared_cells": cleared}


def clear_rows(file_path: str | Path, sheet: str | None,
               rows: str | int) -> dict[str, int | str]:
    """Clear all cell values in contiguous data rows while retaining rows and formatting."""
    path = _path(file_path)
    if _is_csv(path):
        frame = _load_csv(path)
        start, end = _row_bounds(rows, len(frame) + 1)
        values = frame.iloc[start - 2:end - 1, :]
        cleared = int(((values.notna()) & (values.astype(str) != "")).sum().sum())
        frame.iloc[start - 2:end - 1, :] = ""
        _save_csv(frame, path)
        checked = _load_csv(path).iloc[start - 2:end - 1, :]
        if (checked.astype(str) != "").any().any():
            raise ExcelToolError("Could not verify that the selected row values were cleared")
        return {"rows": f"{start}:{end}", "cleared_cells": cleared}
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        start, end = _row_bounds(rows, ws.max_row)
        cleared = 0
        for row_number, row in enumerate(ws.iter_rows(min_row=start, max_row=end), start=start):
            for cell in row:
                if cell.value is not None:
                    cleared += 1
                cell.value = None
            # Empty trailing rows otherwise disappear when openpyxl saves the workbook.
            # Matching the sheet default keeps the visual row height unchanged.
            dimension = ws.row_dimensions[row_number]
            if dimension.height is None:
                dimension.height = ws.sheet_format.defaultRowHeight or 15
        wb.save(path)
    finally:
        wb.close()
    check = _workbook(path)
    try:
        ws = _get_sheet(check, _sheet_name(path, sheet))
        if any(cell.value is not None for row in ws.iter_rows(min_row=start, max_row=end) for cell in row):
            raise ExcelToolError("Could not verify that the selected row values were cleared")
    finally:
        check.close()
    return {"rows": f"{start}:{end}", "cleared_cells": cleared}


def add_column(file_path: str | Path, sheet: str | None, header: str,
               values: Sequence[Any] | None = None, index: int | None = None) -> int:
    path = _path(file_path)
    if _is_csv(path):
        frame = _load_csv(path)
        position = len(frame.columns) if index is None else index - 1
        if position < 0 or position > len(frame.columns):
            raise ExcelToolError("Column index is outside the CSV bounds")
        if header in frame.columns:
            raise ExcelToolError(f"Column '{header}' already exists")
        if values is not None and len(values) != len(frame):
            raise ExcelToolError("Column value count must match the number of data rows")
        frame.insert(position, header, list(values) if values is not None else [None] * len(frame))
        _save_csv(frame, path)
        return position + 1
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        position = ws.max_column + 1 if index is None else index
        if position < 1 or position > ws.max_column + 1:
            raise ExcelToolError("Column index is outside the sheet bounds")
        if values is not None and len(values) != max(ws.max_row - 1, 0):
            raise ExcelToolError("Column value count must match the data row count")
        for table_name in ws.tables:
            table = ws.tables[table_name]
            min_col, min_row, max_col, max_row = range_boundaries(table.ref)
            if min_col <= position <= max_col:
                raise ExcelToolError("Adding a column inside an Excel table is not supported; modify the table separately")
        ws.insert_cols(position)
        source_col = position - 1 if position > 1 else position + 1
        for row in range(1, ws.max_row + 1):
            source, target = ws.cell(row, source_col), ws.cell(row, position)
            if source.has_style:
                target._style = copy(source._style)
            if source.number_format:
                target.number_format = source.number_format
        ws.cell(1, position).value = header
        for offset, value in enumerate(values or [], start=2):
            ws.cell(offset, position).value = value
        wb.save(path)
        return position
    finally:
        wb.close()


def delete_column(file_path: str | Path, sheet: str | None, column: int | str) -> None:
    delete_columns(file_path, sheet, column)


def _column_bounds(specification: str | int, max_columns: int) -> tuple[int, int]:
    """Return inclusive 1-based column bounds from a letter/index/range."""
    from openpyxl.utils.cell import column_index_from_string

    if isinstance(specification, int) or str(specification).isdigit():
        start = end = int(specification)
    else:
        value = str(specification).strip().replace("-", ":").replace(" ", "")
        parts = value.split(":")
        if len(parts) > 2 or any(not part for part in parts):
            raise ExcelToolError(f"Invalid column range '{specification}'. Use a letter or range such as G:J")
        try:
            start = column_index_from_string(parts[0].upper())
            end = column_index_from_string(parts[-1].upper())
        except ValueError as exc:
            raise ExcelToolError(f"Invalid column range '{specification}'. Use a letter or range such as G:J") from exc
    if start < 1 or end < start or end > max_columns:
        raise ExcelToolError(f"Column range '{specification}' is outside sheet bounds (1–{max_columns})")
    return start, end


def _row_bounds(specification: str | int, max_row: int) -> tuple[int, int]:
    """Return inclusive 1-based row bounds, preserving row 1 as the header."""
    if isinstance(specification, int) or str(specification).isdigit():
        start = end = int(specification)
    else:
        parts = str(specification).strip().replace("-", ":").replace(" ", "").split(":")
        if len(parts) > 2 or any(not part.isdigit() for part in parts):
            raise ExcelToolError(f"Invalid row range '{specification}'. Use row numbers such as 5:8")
        start, end = int(parts[0]), int(parts[-1])
    if start < 2 or end < start or end > max_row:
        raise ExcelToolError(f"Row range '{specification}' is outside data rows (2–{max_row})")
    return start, end


def delete_columns(file_path: str | Path, sheet: str | None,
                   columns: str | int) -> tuple[int, int]:
    """Delete one or a contiguous range of columns, shifting remaining columns left."""
    path = _path(file_path)
    if _is_csv(path):
        frame = _load_csv(path)
        if isinstance(columns, str) and columns in frame.columns:
            start = end = list(frame.columns).index(columns) + 1
        else:
            start, end = _column_bounds(columns, len(frame.columns))
        frame = frame.drop(columns=list(frame.columns[start - 1:end]))
        _save_csv(frame, path)
        return start, end
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        start, end = _column_bounds(columns, ws.max_column)
        ws.delete_cols(start, end - start + 1)
        wb.save(path)
        return start, end
    finally:
        wb.close()


def delete_rows(file_path: str | Path, sheet: str | None,
                rows: str | int) -> tuple[int, int]:
    """Delete one or a contiguous range of data rows, preserving the header row."""
    path = _path(file_path)
    if _is_csv(path):
        frame = _load_csv(path)
        start, end = _row_bounds(rows, len(frame) + 1)
        frame = frame.drop(frame.index[start - 2:end - 1]).reset_index(drop=True)
        _save_csv(frame, path)
        return start, end
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        start, end = _row_bounds(rows, ws.max_row)
        ws.delete_rows(start, end - start + 1)
        wb.save(path)
        return start, end
    finally:
        wb.close()


def add_row(file_path: str | Path, sheet: str | None, values: Sequence[Any],
            index: int | None = None) -> int:
    path = _path(file_path)
    if _is_csv(path):
        frame = _load_csv(path)
        if len(values) != len(frame.columns):
            raise ExcelToolError("Row value count must match the number of columns")
        position = len(frame) if index is None else index - 2
        if position < 0 or position > len(frame):
            raise ExcelToolError("Row index is outside the data bounds")
        top, bottom = frame.iloc[:position], frame.iloc[position:]
        frame = pd.concat([top, pd.DataFrame([values], columns=frame.columns), bottom], ignore_index=True)
        _save_csv(frame, path)
        return position + 2
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        if len(values) != ws.max_column:
            raise ExcelToolError("Row value count must match the number of columns")
        position = ws.max_row + 1 if index is None else index
        if position < 2 or position > ws.max_row + 1:
            raise ExcelToolError("Data rows start at row 2; index is outside sheet bounds")
        ws.insert_rows(position)
        for col, value in enumerate(values, start=1):
            ws.cell(position, col).value = value
        wb.save(path)
        return position
    finally:
        wb.close()


def delete_row(file_path: str | Path, sheet: str | None, index: int) -> None:
    path = _path(file_path)
    if _is_csv(path):
        frame = _load_csv(path)
        position = index - 2
        if position < 0 or position >= len(frame):
            raise ExcelToolError("Data rows start at row 2; index is outside CSV bounds")
        _save_csv(frame.drop(frame.index[position]).reset_index(drop=True), path)
        return
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        if index < 2 or index > ws.max_row:
            raise ExcelToolError("Data rows start at row 2; index is outside sheet bounds")
        ws.delete_rows(index)
        wb.save(path)
    finally:
        wb.close()


def delete_table(file_path: str | Path, sheet: str, table_name: str) -> None:
    path = _path(file_path)
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, sheet)
        _find_table(ws, table_name)
        del ws.tables[table_name]
        wb.save(path)
    finally:
        wb.close()


def delete_sheet(file_path: str | Path, sheet: str) -> None:
    path = _path(file_path)
    wb = _workbook(path)
    try:
        _get_sheet(wb, sheet)
        if len(wb.worksheets) == 1:
            raise ExcelToolError("Cannot delete the only worksheet in a workbook")
        del wb[sheet]
        wb.save(path)
    finally:
        wb.close()


def rename_sheet(file_path: str | Path, sheet: str, new_name: str) -> str:
    path = _path(file_path)
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, sheet)
        if not new_name or len(new_name) > 31 or re.search(r"[\\/*?:\[\]]", new_name):
            raise ExcelToolError("Sheet names must be 1–31 characters and cannot contain \\ / * ? : [ ]")
        if new_name in wb.sheetnames and new_name != sheet:
            raise ExcelToolError(f"Sheet '{new_name}' already exists")
        ws.title = new_name
        wb.save(path)
        return new_name
    finally:
        wb.close()


def sort(file_path: str | Path, sheet: str | None, cell_range: str,
         column: int | str, order: str = "asc") -> int:
    path = _path(file_path)
    order = order.lower()
    if order not in {"asc", "desc"}:
        raise ExcelToolError("Sort order must be 'asc' or 'desc'")
    reverse = order == "desc"
    if _is_csv(path):
        frame = _load_csv(path)
        min_col, min_row, max_col, max_row = _parse_range(cell_range)
        if min_row != 1 or min_col != 1 or max_col != len(frame.columns) or max_row > len(frame) + 1:
            raise ExcelToolError("CSV sorting requires a range covering the header and all columns within CSV bounds")
        if str(column).isdigit():
            column_index = int(column)
            if column_index < 1 or column_index > len(frame.columns):
                raise ExcelToolError("Sort column index is outside CSV bounds")
            sort_column = frame.columns[column_index - 1]
        else:
            sort_column = column
        if sort_column not in frame.columns:
            raise ExcelToolError(f"Sort column '{sort_column}' not found")
        first_data, last_data = min_row - 1, max_row - 1
        sorted_rows = frame.iloc[first_data:last_data].sort_values(
            sort_column, ascending=not reverse, kind="stable"
        )
        frame.iloc[first_data:last_data] = sorted_rows.values
        _save_csv(frame, path)
        return len(sorted_rows)
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        min_col, min_row, max_col, max_row = _range(ws, cell_range)
        col_index = int(column) if isinstance(column, int) or str(column).isdigit() else None
        if col_index is None:
            from openpyxl.utils.cell import column_index_from_string
            try:
                col_index = column_index_from_string(str(column).upper())
            except ValueError as exc:
                raise ExcelToolError(f"Invalid sort column '{column}'") from exc
        if col_index < min_col or col_index > max_col:
            raise ExcelToolError("Sort column must be inside the selected range")
        offset = col_index - min_col
        try:
            body = [list(row) for row in ws.iter_rows(
                min_row=min_row + 1, max_row=max_row,
                min_col=min_col, max_col=max_col, values_only=True,
            )]
            body.sort(key=lambda row: (row[offset] is None, row[offset]), reverse=reverse)
        except TypeError as exc:
            raise ExcelToolError("Sort column contains values that cannot be compared") from exc
        for row_offset, record in enumerate(body, start=min_row + 1):
            for col_offset, value in enumerate(record, start=min_col):
                ws.cell(row_offset, col_offset).value = value
        wb.save(path)
        return len(body)
    finally:
        wb.close()


def filter(file_path: str | Path, sheet: str | None, cell_range: str,
           condition: Mapping[str, Any]) -> list[list[Any]]:
    """Return matching rows; filtering is a read operation and leaves source intact."""
    path = _path(file_path)
    if "column" not in condition or "value" not in condition:
        raise ExcelToolError("Filter condition requires 'column' and 'value'")
    if _is_csv(path):
        frame = _load_csv(path)
        min_col, min_row, max_col, max_row = _parse_range(cell_range)
        if min_row != 1 or max_row > len(frame) + 1 or max_col > len(frame.columns):
            raise ExcelToolError(f"Range '{cell_range}' is outside CSV bounds")
        selected = frame.iloc[:max_row - 1, min_col - 1:max_col]
        selected_headers = list(selected.columns)
        col = condition["column"]
        if str(col).isdigit():
            position = int(col) - 1
            if position < 0 or position >= len(selected_headers):
                raise ExcelToolError("Filter column must be inside the selected range")
            col = selected_headers[position]
        if col not in selected_headers:
            raise ExcelToolError(f"Filter column '{col}' not found in selected range")
        op, value = condition.get("operator", "equals"), condition["value"]
        mask = _filter_mask(selected[col].astype(str), op, value)
        return [selected_headers, *selected.loc[mask].values.tolist()]
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        min_col, min_row, max_col, max_row = _range(ws, cell_range)
        rows = _values(ws, min_col, min_row, max_col, max_row)
        column = condition["column"]
        if isinstance(column, str) and not column.isdigit():
            try:
                offset = [str(value) for value in rows[0]].index(column)
            except ValueError as exc:
                raise ExcelToolError(f"Filter column '{column}' not found in range headers") from exc
        else:
            offset = int(column) - 1
        if offset < 0 or offset >= len(rows[0]):
            raise ExcelToolError("Filter column must be inside the selected range")
        return [rows[0], *[row for row in rows[1:]
                           if _matches(row[offset], condition.get("operator", "equals"), condition["value"])]]
    finally:
        wb.close()


def _matches(actual: Any, operator: str, expected: Any) -> bool:
    op = operator.lower()
    if op == "equals": return str(actual) == str(expected)
    if op == "not_equals": return str(actual) != str(expected)
    if op == "contains": return str(expected).casefold() in str(actual).casefold()
    try:
        a, b = float(actual), float(expected)
    except (ValueError, TypeError):
        raise ExcelToolError(f"Operator '{op}' requires numeric values")
    if op == "greater_than": return a > b
    if op == "greater_or_equal": return a >= b
    if op == "less_than": return a < b
    if op == "less_or_equal": return a <= b
    raise ExcelToolError(f"Unsupported filter operator '{operator}'")


def _filter_mask(series: pd.Series, operator: str, value: Any) -> pd.Series:
    if operator.lower() == "contains":
        return series.str.contains(str(value), case=False, regex=False)
    return series.map(lambda actual: _matches(actual, operator, value))


def _delete_row_runs(worksheet, row_numbers: Sequence[int]) -> None:
    """Delete contiguous row groups from bottom to top to avoid repeated shifts."""
    if not row_numbers:
        return
    ordered = sorted(set(row_numbers))
    runs: list[tuple[int, int]] = []
    run_start = previous = ordered[0]
    for row_number in ordered[1:]:
        if row_number == previous + 1:
            previous = row_number
            continue
        runs.append((run_start, previous - run_start + 1))
        run_start = previous = row_number
    runs.append((run_start, previous - run_start + 1))
    for start, amount in reversed(runs):
        worksheet.delete_rows(start, amount)


def remove_duplicates(file_path: str | Path, sheet: str | None,
                      cell_range: str | None = None,
                      columns: Sequence[str | int] | None = None) -> int:
    path = _path(file_path)
    if _is_csv(path):
        frame = _load_csv(path)
        subset = None
        if columns:
            subset = [frame.columns[int(c) - 1] if str(c).isdigit() else c for c in columns]
            missing = [c for c in subset if c not in frame.columns]
            if missing:
                raise ExcelToolError(f"Duplicate check columns not found: {missing}")
        if cell_range:
            min_col, min_row, max_col, max_row = _parse_range(cell_range)
            if min_row != 1 or max_row > len(frame) + 1 or max_col > len(frame.columns):
                raise ExcelToolError(f"Range '{cell_range}' is outside CSV bounds")
            range_columns = list(frame.columns[min_col - 1:max_col])
            if subset is not None and any(name not in range_columns for name in subset):
                raise ExcelToolError("Duplicate check columns must be inside the selected range")
            subset = subset or range_columns
            selected = frame.iloc[min_row - 1:max_row - 1]
            duplicates = selected.index[selected.duplicated(subset=subset, keep="first")]
            _save_csv(frame.drop(index=duplicates).reset_index(drop=True), path)
            return len(duplicates)
        before = len(frame)
        _save_csv(frame.drop_duplicates(subset=subset, keep="first"), path)
        return before - len(_load_csv(path))
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, _sheet_name(path, sheet))
        cell_range = cell_range or f"A1:{ws.cell(ws.max_row, ws.max_column).coordinate}"
        min_col, min_row, max_col, max_row = _range(ws, cell_range)
        if max_row <= min_row:
            return 0
        selected = list(range(min_col, max_col + 1))
        if columns:
            selected = []
            headers = [ws.cell(min_row, col).value for col in range(min_col, max_col + 1)]
            for column in columns:
                if isinstance(column, int) or str(column).isdigit():
                    col = int(column)
                    if min_col <= col <= max_col:
                        selected.append(col)
                    elif 1 <= col <= len(headers):
                        selected.append(min_col + col - 1)
                    else:
                        raise ExcelToolError(f"Duplicate column '{column}' is outside selected range")
                else:
                    if column not in headers:
                        raise ExcelToolError(f"Duplicate column '{column}' not found in selected range")
                    selected.append(min_col + headers.index(column))
        seen, duplicate_rows = set(), []
        for row in range(min_row + 1, max_row + 1):
            key = tuple(ws.cell(row, col).value for col in selected)
            if key in seen:
                duplicate_rows.append(row)
            else:
                seen.add(key)
        _delete_row_runs(ws, duplicate_rows)
        wb.save(path)
        return len(duplicate_rows)
    finally:
        wb.close()


def apply_formatting(file_path: str | Path, sheet: str, cell_range: str,
                     bold: bool | None = None, fill: str | None = None,
                     borders: bool | None = None, number_format: str | None = None) -> int:
    path = _path(file_path)
    wb = _workbook(path)
    try:
        ws = _get_sheet(wb, sheet)
        min_col, min_row, max_col, max_row = _range(ws, cell_range)
        count = 0
        for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                min_col=min_col, max_col=max_col):
            for cell in row:
                if bold is not None:
                    font = copy(cell.font); font.bold = bold; cell.font = font
                if fill is not None:
                    color = fill.lstrip("#")
                    if not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
                        raise ExcelToolError("Fill must be a six-digit hex color, such as 'FFFF00'")
                    cell.fill = PatternFill(fill_type="solid", fgColor=color.upper())
                if borders:
                    thin = Side(style="thin", color="000000")
                    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                if number_format is not None:
                    cell.number_format = number_format
                count += 1
        wb.save(path)
        return count
    finally:
        wb.close()


def add_formula(file_path: str | Path, sheet: str, cell: str, formula: str) -> str:
    if not formula.startswith("="):
        raise ExcelToolError("Formula must start with '='")
    return update_cells(file_path, sheet, {cell: formula}) and cell


def merge_files(files: Sequence[str | Path], mode: str = "append",
                output_path: str | Path | None = None, key: str | None = None) -> str:
    """Merge CSV/XLSX first sheets into a new workbook without modifying sources."""
    if len(files) < 2:
        raise ExcelToolError("Merge requires at least two input files")
    if mode not in {"append", "join_on_key"}:
        raise ExcelToolError("Merge mode must be 'append' or 'join_on_key'")
    frames = []
    for file in files:
        path = _path(file)
        if _is_csv(path):
            frames.append(_load_csv(path))
        else:
            wb = _workbook(path)
            try:
                ws = wb.active
                records = list(ws.values)
            finally:
                wb.close()
            frames.append(pd.DataFrame(records[1:], columns=records[0]) if records else pd.DataFrame())
    if mode == "append":
        frame = pd.concat(frames, ignore_index=True, sort=False)
    else:
        if not key:
            raise ExcelToolError("join_on_key requires a key column name")
        for frame in frames:
            if key not in frame.columns:
                raise ExcelToolError(f"Join key '{key}' is missing from an input file")
        frame = frames[0]
        for other in frames[1:]:
            frame = frame.merge(other, on=key, how="outer", suffixes=("", "_right"))
    if output_path is None:
        parent = Path(files[0]).parent
        output_path = parent / f"merged_{uuid.uuid4().hex[:8]}.xlsx"
    output = Path(output_path)
    if output.suffix.lower() != ".xlsx":
        raise ExcelToolError("Merged output path must end in .xlsx")
    output_resolved = output.resolve()
    if any(Path(source).resolve() == output_resolved for source in files):
        raise ExcelToolError("Merged output must be a new file and cannot replace an input")
    if output.exists():
        raise ExcelToolError(f"Output file already exists: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            frame.to_excel(writer, index=False, sheet_name="Merged")
    except Exception as exc:
        output.unlink(missing_ok=True)
        raise ExcelToolError(f"Could not write merged workbook: {exc}") from exc
    return str(output)
