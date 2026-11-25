"""Write complete Office files in a command's temporary directory."""

import re
import textwrap
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.shared import Inches, Pt
from openpyxl import Workbook


def write_from_template(path: Path, source: Path, args: dict) -> None:
    if source.suffix.lower() != ".docx":
        raise ValueError("Choose a Word .docx file as the template")
    replacements = {}
    for entry in args.get("replacements", []):
        phrase, text = entry["source"], entry["text"]
        if not phrase or phrase in replacements or not isinstance(text, str):
            raise ValueError("Use a unique nonempty source phrase for each replacement")
        replacements[phrase] = text
    additions = args.get("append", [])
    if not replacements and not additions:
        raise ValueError("Supply replacement text or paragraphs for the new draft")
    document = Document(source)
    document.core_properties.title = args["title"]
    if replacements:
        pattern = re.compile("|".join(re.escape(s) for s in sorted(replacements, key=len, reverse=True)))
        found = set()
        parts = [document.part] + [part for part in document.part.package.parts
                                  if isinstance(part, (FooterPart, HeaderPart))]
        for part in parts:
            for paragraph in part.element.xpath(".//w:p"):
                nodes = paragraph.xpath("./w:r/w:t | ./w:hyperlink/w:r/w:t")
                original = "".join(node.text or "" for node in nodes)
                # Work backwards to keep original character offsets valid across formatted runs.
                for match in reversed(list(pattern.finditer(original))):
                    found.add(match.group())
                    offset = 0
                    for node in nodes:
                        text = node.text or ""
                        end = offset + len(text)
                        if offset < match.end() and end > match.start():
                            left = max(0, match.start() - offset)
                            right = min(len(text), match.end() - offset)
                            insert = replacements[match.group()] if offset <= match.start() else ""
                            node.text = text[:left] + insert + text[right:]
                            node.set(qn("xml:space"), "preserve")
                        offset = end
        missing = replacements.keys() - found
        if missing:
            raise ValueError(f"Template text not found: {', '.join(sorted(missing))}")
    for paragraph in additions:
        document.add_paragraph(paragraph["text"], style=paragraph.get("style") or "Normal")
    document.save(path)


def requested_page_limit(command: str) -> int | None:
    numbers = "one two three four five six seven eight nine ten".split()
    limits = re.findall(r"\b(\d+|" + "|".join(numbers) + r")[ -]+pages?\b", command.lower())
    if limits:
        value = limits[-1]
        return int(value) if value.isdigit() else numbers.index(value) + 1
    return None


def write_document(path: Path, args: dict, command: str) -> None:
    title, pages = args["title"], args["pages"]
    if not isinstance(title, str) or not title.strip() or not isinstance(pages, list) or not pages:
        raise ValueError("Supply a document title and at least one page")
    limit = requested_page_limit(command)
    if limit is not None and len(pages) > limit:
        raise ValueError(f"Condense the document to at most {limit} pages. You supplied {len(pages)} "
                         f"entries in pages; combine sections into at most {limit} entries.")

    document = Document()
    document.core_properties.title = title
    section = document.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(1)
    normal = document.styles["Normal"]
    normal.font.name, normal.font.size = "Calibri", Pt(11)
    for index, page in enumerate(pages):
        heading, body = page["heading"], page["text"]
        if not isinstance(body, str) or not body.strip() or not isinstance(heading, str) or not heading.strip():
            raise ValueError("Each page needs a heading and body text")
        if len(heading) > 100 or len(body) > 1200:
            raise ValueError("Shorten each heading to 100 characters and each page to 1200 characters")
        lines = [line for paragraph in body.splitlines()
                 for line in (textwrap.wrap(paragraph, width=46) or [""])]
        if len(lines) > 32:
            raise ValueError("Condense the text to fit on the page")
        heading_paragraph = document.add_heading(" ".join(heading.split()), level=1)
        heading_paragraph.paragraph_format.page_break_before = bool(index)
        paragraph = document.add_paragraph("\n".join(lines))
        paragraph.paragraph_format.line_spacing = Pt(14)
    document.save(path)


def write_workbook(path: Path, args: dict) -> None:
    if not isinstance(args["title"], str) or not args["title"].strip() or not args["sheets"]:
        raise ValueError("Supply a workbook title and at least one sheet")
    workbook = Workbook()
    workbook.remove(workbook.active)
    try:
        for data in args["sheets"]:
            name, rows = data["name"], data["rows"]
            if not isinstance(name, str) or not name.strip() or len(name) > 31:
                raise ValueError("Supply a sheet name of 1 to 31 characters")
            if not isinstance(rows, list) or not rows:
                raise ValueError("Include at least a header row")
            sheet = workbook.create_sheet(name)
            for row in rows:
                if not isinstance(row, list) or not row:
                    raise ValueError("Each row must be a nonempty list of cell values")
                sheet.append(row)
        workbook.save(path)
    finally:
        workbook.close()
