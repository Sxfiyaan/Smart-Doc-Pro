
from langchain_core.tools import tool

from app.documents.excel import inspect_excel


@tool
def inspect_excel_tool(file_path: str) -> dict:
    """Inspect an Excel workbook and return its sheet names,
    dimensions, headers, and sample rows."""
    return inspect_excel(file_path)
