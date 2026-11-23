"""Safe, index-based Word document operations for SmartDoc."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any, Sequence

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.shared import Inches
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from lxml import etree

RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class WordToolError(ValueError):
    """An actionable validation or operation error returned to the agent."""


def _path(file_path: str | Path) -> Path:
    path = Path(file_path)
    if not path.is_file():
        raise WordToolError(f"File not found: {path}")
    if path.suffix.lower() != ".docx":
        raise WordToolError("Word operations require a .docx file")
    return path


def _document(file_path: str | Path) -> tuple[Path, Document]:
    path = _path(file_path)
    try:
        return path, Document(path)
    except Exception as exc:
        raise WordToolError(f"Could not open Word document: {exc}") from exc


def _paragraph(document: Document, index: int) -> Paragraph:
    if not isinstance(index, int) or index < 0 or index >= len(document.paragraphs):
        raise WordToolError(f"Paragraph {index} was not found; valid indexes are 0 to {max(len(document.paragraphs) - 1, 0)}")
    return document.paragraphs[index]


def _table(document: Document, index: int) -> Table:
    if not isinstance(index, int) or index < 0 or index >= len(document.tables):
        raise WordToolError(f"Table {index} was not found; valid indexes are 0 to {max(len(document.tables) - 1, 0)}")
    return document.tables[index]


def _cell(table: Table, row: int, column: int) -> _Cell:
    if not isinstance(row, int) or row < 0 or row >= len(table.rows):
        raise WordToolError(f"Row {row} is outside table bounds")
    if not isinstance(column, int) or column < 0 or column >= len(table.columns):
        raise WordToolError(f"Column {column} is outside table bounds")
    return table.cell(row, column)


def _paragraph_style(document: Document, style: str) -> str:
    try:
        selected = document.styles[style]
    except KeyError as exc:
        raise WordToolError(f"Paragraph style '{style}' was not found") from exc
    if selected.type != WD_STYLE_TYPE.PARAGRAPH:
        raise WordToolError(f"Style '{style}' is not a paragraph style")
    return style


def _table_values(table: Table, limit: int = 10) -> list[list[str]]:
    return [[cell.text for cell in row.cells] for row in table.rows[:limit]]


def read_document(file_path: str | Path) -> dict[str, Any]:
    """Return a compact, stable summary that lets the agent address document content."""
    _, document = _document(file_path)
    paragraphs = []
    headings = []
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        entry = {"index": index, "text": text, "style": paragraph.style.name}
        paragraphs.append(entry)
        if paragraph.style.name.lower().startswith("heading"):
            headings.append(entry)
    tables = [{
        "index": index,
        "rows": len(table.rows),
        "columns": len(table.columns),
        "sample_rows": _table_values(table),
        "sample_truncated": len(table.rows) > 10,
    } for index, table in enumerate(document.tables)]
    return {
        "file_name": Path(file_path).name,
        "paragraph_count": len(document.paragraphs),
        "paragraphs": paragraphs[:60],
        "paragraphs_truncated": len(paragraphs) > 60,
        "headings": headings,
        "tables": tables,
    }


def merge_documents(file_paths: Sequence[str | Path], output_path: str | Path) -> str:
    """Create a new DOCX containing each source document's body in the given order."""
    if len(file_paths) < 2:
        raise WordToolError("Merging Word documents requires at least two files")
    paths = [_path(file_path) for file_path in file_paths]
    output = Path(output_path)
    if output.suffix.lower() != ".docx":
        raise WordToolError("Merged Word output must end in .docx")
    output_resolved = output.resolve()
    if any(path.resolve() == output_resolved for path in paths):
        raise WordToolError("Merged output must be a new file and cannot replace an input")
    if output.exists():
        raise WordToolError(f"Output file already exists: {output}")

    try:
        combined = Document(paths[0])
        drawing_id_next = _next_package_xml_id(
            combined.part.package,
            "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr",
        )
        bookmark_id_next = _next_package_xml_id(
            combined.part.package,
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart",
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkEnd",
        )
        for source_path in paths[1:]:
            source = Document(source_path)
            copied_parts: dict[tuple[int, str], Part] = {}
            reserved_names = {str(part.partname) for part in combined.part.package.iter_parts()}
            bookmark_ids: dict[str, str] = {}
            combined.add_page_break()
            body = combined._element.body
            section_properties = body.sectPr
            insertion_index = list(body).index(section_properties) if section_properties is not None else len(body)
            for child in source._element.body:
                if child.tag.endswith("}sectPr"):
                    continue
                copied = deepcopy(child)
                _remap_document_local_ids(copied, drawing_id_next, bookmark_id_next, bookmark_ids)
                _copy_referenced_styles(copied, source, combined)
                _remap_document_relationships(
                    copied, source, combined, copied_parts, reserved_names,
                    drawing_id_next, bookmark_id_next, bookmark_ids,
                )
                body.insert(insertion_index, copied)
                insertion_index += 1
        output.parent.mkdir(parents=True, exist_ok=True)
        combined.save(output)
    except WordToolError:
        output.unlink(missing_ok=True)
        raise
    except Exception as exc:
        output.unlink(missing_ok=True)
        raise WordToolError(f"Could not merge Word documents: {exc}") from exc
    return str(output)


def _next_package_xml_id(package, *tags: str) -> list[int]:
    """Return a mutable counter above IDs used anywhere in the document package."""
    values = []
    for part in package.iter_parts():
        if not (part.content_type.endswith("+xml") or part.content_type in {"application/xml", "text/xml"}):
            continue
        try:
            root = etree.fromstring(part.blob)
        except Exception:
            continue
        for node in root.iter():
            if node.tag not in tags:
                continue
            value = node.get("id") or node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
            try:
                values.append(int(value))
            except (TypeError, ValueError):
                continue
    return [max(values, default=0) + 1]


def _remap_document_local_ids(element, drawing_id_next: list[int], bookmark_id_next: list[int],
                              bookmark_ids: dict[str, str]) -> None:
    """Avoid document-wide drawing and bookmark ID collisions from appended sources."""
    word_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    drawing_docpr = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"
    for node in element.iter():
        if node.tag == drawing_docpr:
            node.set("id", str(drawing_id_next[0]))
            drawing_id_next[0] += 1
        elif node.tag in {f"{word_ns}bookmarkStart", f"{word_ns}bookmarkEnd"}:
            old_id = node.get(f"{word_ns}id")
            if old_id is None:
                continue
            if old_id not in bookmark_ids:
                bookmark_ids[old_id] = str(bookmark_id_next[0])
                bookmark_id_next[0] += 1
            node.set(f"{word_ns}id", bookmark_ids[old_id])


def _copy_referenced_styles(element, source: Document, target: Document) -> None:
    """Copy missing paragraph/table styles used by imported body content."""
    style_id_tag = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId"
    style_reference_tags = {
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle",
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblStyle",
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rStyle",
    }
    source_styles = source.styles.element
    target_styles = target.styles.element
    source_by_id = {style.get(style_id_tag): style for style in source_styles if style.get(style_id_tag)}
    target_ids = {style.get(style_id_tag) for style in target_styles if style.get(style_id_tag)}
    pending = [node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
               for node in element.iter() if node.tag in style_reference_tags]
    while pending:
        style_id = pending.pop()
        if not style_id or style_id in target_ids or style_id not in source_by_id:
            continue
        style = deepcopy(source_by_id[style_id])
        target_styles.append(style)
        target_ids.add(style_id)
        for dependency in ("basedOn", "next", "link"):
            node = style.find(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{dependency}")
            if node is not None:
                pending.append(node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"))


def _copy_related_part(source_part: Part, target: Document, copied_parts: dict,
                       reserved_names: set[str], drawing_id_next: list[int],
                       bookmark_id_next: list[int], bookmark_ids: dict[str, str]) -> Part:
    """Copy a related OPC part and its relationship graph into the target package."""
    key = (id(source_part.package), str(source_part.partname))
    if key in copied_parts:
        return copied_parts[key]

    # Keep the source directory and extension so Word recognizes charts, OLE objects,
    # diagrams, and other package parts. The package chooses a collision-free name.
    source_name = str(source_part.partname)
    directory, _, filename = source_name.rpartition("/")
    stem, dot, extension = filename.rpartition(".")
    if not dot:
        template = f"{directory}/{filename}%d" if directory else f"/{filename}%d"
    else:
        template = f"{directory}/{stem}%d.{extension}" if directory else f"/{stem}%d.{extension}"
    number = 1
    new_name_text = template % number
    while new_name_text in reserved_names:
        number += 1
        new_name_text = template % number
    new_name = PackURI(new_name_text)
    reserved_names.add(new_name_text)
    copied = Part.load(new_name, source_part.content_type, source_part.blob, target.part.package)
    # Memoize before walking relationships to safely handle cycles in package graphs.
    copied_parts[key] = copied

    relationship_ids: dict[str, str] = {}
    try:
        for relationship in source_part.rels.values():
            if relationship.is_external:
                new_id = copied.relate_to(
                    relationship.target_ref, relationship.reltype, is_external=True
                )
            else:
                related_part = _copy_related_part(
                    relationship.target_part, target, copied_parts, reserved_names,
                    drawing_id_next, bookmark_id_next, bookmark_ids,
                )
                new_id = copied.relate_to(related_part, relationship.reltype)
            relationship_ids[relationship.rId] = new_id
    except Exception as exc:
        raise WordToolError(
            f"Could not copy embedded Word content '{source_name}' "
            f"({source_part.content_type}): {exc}"
        ) from exc

    # Relationship IDs are local to each package part. Rewrite them in copied XML
    # parts because the destination package may assign different IDs.
    if source_part.content_type.endswith("+xml") or source_part.content_type in {"application/xml", "text/xml"}:
        try:
            root = etree.fromstring(source_part.blob)
            _remap_document_local_ids(root, drawing_id_next, bookmark_id_next, bookmark_ids)
            for node in root.iter():
                for attribute, relationship_id in list(node.attrib.items()):
                    if attribute.startswith(f"{{{RELATIONSHIP_NS}}}"):
                        replacement = relationship_ids.get(relationship_id)
                        if replacement:
                            node.set(attribute, replacement)
            copied._blob = etree.tostring(root, encoding="UTF-8", xml_declaration=True)
        except Exception as exc:
            raise WordToolError(
                f"Could not update relationships in embedded Word content '{source_name}': {exc}"
            ) from exc
    return copied


def _remap_document_relationships(element, source: Document, target: Document,
                                  copied_parts: dict, reserved_names: set[str],
                                  drawing_id_next: list[int], bookmark_id_next: list[int],
                                  bookmark_ids: dict[str, str]) -> None:
    """Retarget body relationships, including recursively related embedded parts."""
    for node in element.iter():
        for attribute, relationship_id in list(node.attrib.items()):
            if not attribute.startswith(f"{{{RELATIONSHIP_NS}}}"):
                continue
            relationship = source.part.rels.get(relationship_id)
            if relationship is None:
                raise WordToolError(f"A source document relationship '{relationship_id}' could not be resolved")
            if relationship.is_external:
                new_id = target.part.relate_to(
                    relationship.target_ref, relationship.reltype, is_external=True
                )
            elif relationship.reltype.endswith("/image"):
                new_id, _ = target.part.get_or_add_image(BytesIO(relationship.target_part.blob))
            else:
                related_part = _copy_related_part(
                    relationship.target_part, target, copied_parts, reserved_names,
                    drawing_id_next, bookmark_id_next, bookmark_ids,
                )
                new_id = target.part.relate_to(related_part, relationship.reltype)
            node.set(attribute, new_id)


def insert_paragraph(file_path: str | Path, text: str, index: int | None = None,
                     style: str | None = None) -> int:
    path, document = _document(file_path)
    if not isinstance(text, str):
        raise WordToolError("Paragraph text must be a string")
    if style is not None:
        _paragraph_style(document, style)
    if index is None or index == len(document.paragraphs):
        paragraph = document.add_paragraph(text, style=style)
        result = len(document.paragraphs) - 1
    else:
        target = _paragraph(document, index)
        element = OxmlElement("w:p")
        target._p.addprevious(element)
        paragraph = Paragraph(element, target._parent)
        paragraph.text = text
        if style is not None:
            paragraph.style = style
        result = index
    document.save(path)
    return result


def edit_paragraph(file_path: str | Path, index: int, text: str,
                   style: str | None = None) -> None:
    path, document = _document(file_path)
    paragraph = _paragraph(document, index)
    if not isinstance(text, str):
        raise WordToolError("Paragraph text must be a string")
    paragraph.text = text
    if style is not None:
        paragraph.style = _paragraph_style(document, style)
    document.save(path)


def delete_paragraph(file_path: str | Path, index: int) -> None:
    path, document = _document(file_path)
    paragraph = _paragraph(document, index)
    paragraph._element.getparent().remove(paragraph._element)
    document.save(path)


def apply_style(file_path: str | Path, index: int, style: str) -> None:
    path, document = _document(file_path)
    paragraph = _paragraph(document, index)
    paragraph.style = _paragraph_style(document, style)
    document.save(path)


def create_table(file_path: str | Path, rows: int, columns: int,
                 headers: Sequence[str] | None = None, style: str | None = "Table Grid",
                 location: int | None = None) -> int:
    path, document = _document(file_path)
    if not isinstance(rows, int) or rows < 1 or not isinstance(columns, int) or columns < 1:
        raise WordToolError("A table needs at least one row and one column")
    if headers is not None and len(headers) != columns:
        raise WordToolError("Header count must match the number of table columns")
    result_index = len(document.tables)
    table = document.add_table(rows=rows, cols=columns)
    if style is not None:
        try:
            table.style = style
        except (KeyError, ValueError) as exc:
            raise WordToolError(f"Table style '{style}' was not found") from exc
    if headers is not None:
        for column, header in enumerate(headers):
            table.cell(0, column).text = str(header)
    if location is not None:
        target = _paragraph(document, location)
        target._p.addnext(table._tbl)
    document.save(path)
    return result_index


def modify_table_cell(file_path: str | Path, table_index: int, row: int, column: int,
                      text: str) -> None:
    path, document = _document(file_path)
    if not isinstance(text, str):
        raise WordToolError("Cell text must be a string")
    _cell(_table(document, table_index), row, column).text = text
    document.save(path)


def add_table_row(file_path: str | Path, table_index: int,
                  values: Sequence[Any] | None = None) -> int:
    path, document = _document(file_path)
    table = _table(document, table_index)
    if values is not None and len(values) != len(table.columns):
        raise WordToolError("Row values must match the table column count")
    row = table.add_row()
    if values is not None:
        for cell, value in zip(row.cells, values):
            cell.text = str(value)
    document.save(path)
    return len(table.rows) - 1


def add_table_column(file_path: str | Path, table_index: int, header: str = "",
                     values: Sequence[Any] | None = None) -> int:
    path, document = _document(file_path)
    table = _table(document, table_index)
    if values is not None and len(values) != max(len(table.rows) - 1, 0):
        raise WordToolError("Column values must match the number of data rows (excluding the header row)")
    column = table.add_column(Inches(1))
    column.cells[0].text = str(header)
    if values is not None:
        for cell, value in zip(column.cells[1:], values):
            cell.text = str(value)
    document.save(path)
    return len(table.columns) - 1


def delete_table_row(file_path: str | Path, table_index: int, row: int) -> None:
    path, document = _document(file_path)
    table = _table(document, table_index)
    if len(table.rows) <= 1:
        raise WordToolError("Cannot delete the only row in a table")
    if not isinstance(row, int) or row < 0 or row >= len(table.rows):
        raise WordToolError(f"Row {row} is outside table bounds")
    table._tbl.remove(table.rows[row]._tr)
    document.save(path)


def delete_table_column(file_path: str | Path, table_index: int, column: int) -> None:
    path, document = _document(file_path)
    table = _table(document, table_index)
    if len(table.columns) <= 1:
        raise WordToolError("Cannot delete the only column in a table")
    if not isinstance(column, int) or column < 0 or column >= len(table.columns):
        raise WordToolError(f"Column {column} is outside table bounds")
    for row in table.rows:
        row._tr.remove(row.cells[column]._tc)
    grid = table._tbl.tblGrid
    grid.remove(grid.gridCol_lst[column])
    document.save(path)


def delete_table(file_path: str | Path, table_index: int) -> None:
    path, document = _document(file_path)
    table = _table(document, table_index)
    table._tbl.getparent().remove(table._tbl)
    document.save(path)


def destructive_preview(file_path: str | Path, operation: str, args: dict[str, Any]) -> dict[str, Any]:
    """Return only the exact content a user should review before destructive edits."""
    _, document = _document(file_path)
    if operation == "delete_paragraph":
        index = args["index"]
        paragraph = _paragraph(document, index)
        return {"paragraph_index": index, "paragraph_text": paragraph.text, "style": paragraph.style.name}
    if operation == "delete_table":
        index = args["table_index"]
        table = _table(document, index)
        return {"table_index": index, "rows": len(table.rows), "columns": len(table.columns),
                "sample_rows": _table_values(table), "sample_truncated": len(table.rows) > 10}
    table = _table(document, args["table_index"])
    if operation == "delete_table_row":
        row = args["row"]
        return {"table_index": args["table_index"], "row": row,
                "values": [cell.text for cell in table.rows[row].cells]}
    if operation == "delete_table_column":
        column = args["column"]
        _cell(table, 0, column)
        return {"table_index": args["table_index"], "column": column,
                "values": [row.cells[column].text for row in table.rows]}
    raise WordToolError(f"No preview is available for '{operation}'")
