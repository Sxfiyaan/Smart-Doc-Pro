
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def inspect_excel(file_path: str) -> dict[str, Any]:
    """
    Inspect an Excel workbook and return its sheet names,
    dimensions, headers, and a small sample of rows.
    """
    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Only .xlsx and .xlsm files are supported.")

    workbook = load_workbook(
        filename=path,
        read_only=True,
        data_only=True,
    )

    try:
        result = {"file_name": path.name, "sheets": []}

        for worksheet in workbook.worksheets:
            rows = worksheet.iter_rows(values_only=True)

            header_row = next(rows, ())
            headers = [
                str(value) if value is not None else f"column_{index + 1}"
                for index, value in enumerate(header_row)
            ]

            sample_rows = []
            for row in rows:
                sample_rows.append(list(row))
                if len(sample_rows) >= 5:
                    break

            result["sheets"].append({
                "sheet_name": worksheet.title,
                "row_count": max(worksheet.max_row - 1, 0),
                "column_count": worksheet.max_column,
                "headers": headers,
                "sample_rows": sample_rows,
            })

        return result

    finally:
        workbook.close()
