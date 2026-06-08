"""Rewrite uploaded documents while preserving native file formats when possible."""

from __future__ import annotations

import io

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

from .file_parser import ESTIMATED_LINES_PER_PAGE, parse_file
from .inclusive_scanner import apply_all_accepted_replacements

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def build_corrected_text(parsed_lines: list, accepted_by_key: dict) -> str:
    """Render corrected parsed lines into the current plain-text output format."""
    corrected_chunks = []
    current_page = None

    for entry in parsed_lines:
        key = (entry["page"], entry["line"])
        text = entry["text"]
        if key in accepted_by_key:
            text = apply_all_accepted_replacements(text, accepted_by_key[key])

        if entry["page"] != current_page:
            if corrected_chunks:
                corrected_chunks.append("")
            corrected_chunks.append(f"--- Page/Section {entry['page']} ---")
            current_page = entry["page"]

        corrected_chunks.append(f"L{entry['line']:>4}: {text}")

    return "\n".join(corrected_chunks).strip() + "\n"


def _rewrite_runs_in_paragraph(paragraph, accepted: list) -> None:
    runs = list(paragraph.runs)
    if not runs:
        updated_text = apply_all_accepted_replacements(paragraph.text, accepted)
        if updated_text != paragraph.text:
            paragraph.text = updated_text
        return

    original_text = "".join(run.text for run in runs)
    updated_text = apply_all_accepted_replacements(original_text, accepted)
    if updated_text == original_text:
        return

    # A target word can be split across runs. Replacing against the combined
    # run text avoids misses, then we write the full result into run[0]. This
    # keeps most formatting but can collapse intra-word run boundaries.
    runs[0].text = updated_text
    for run in runs[1:]:
        run.text = ""


def rewrite_docx(file_bytes: bytes, accepted_by_key: dict) -> bytes:
    doc = Document(io.BytesIO(file_bytes))
    line_num = 0

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            line_num += 1
            page_num = ((line_num - 1) // ESTIMATED_LINES_PER_PAGE) + 1
            key = (page_num, line_num)
            accepted = accepted_by_key.get(key)
            if accepted:
                _rewrite_runs_in_paragraph(para, accepted)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if text:
                    line_num += 1
                    page_num = ((line_num - 1) // ESTIMATED_LINES_PER_PAGE) + 1
                    key = (page_num, line_num)
                    accepted = accepted_by_key.get(key)
                    if accepted:
                        for para in cell.paragraphs:
                            _rewrite_runs_in_paragraph(para, accepted)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def rewrite_pptx(file_bytes: bytes, accepted_by_key: dict) -> bytes:
    prs = Presentation(io.BytesIO(file_bytes))

    for slide_num, slide in enumerate(prs.slides, start=1):
        line_num = 0

        for shape in slide.shapes:
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        line_num += 1
                        key = (slide_num, line_num)
                        accepted = accepted_by_key.get(key)
                        if accepted:
                            _rewrite_runs_in_paragraph(para, accepted)

            if hasattr(shape, "has_table") and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text:
                            line_num += 1
                            key = (slide_num, line_num)
                            accepted = accepted_by_key.get(key)
                            if accepted:
                                for para in cell.text_frame.paragraphs:
                                    _rewrite_runs_in_paragraph(para, accepted)

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


def rewrite_xlsx(file_bytes: bytes, accepted_by_key: dict) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=False, data_only=False)

    for sheet_num, ws in enumerate(wb.worksheets, start=1):
        for row_num, row in enumerate(ws.iter_rows(), start=1):
            cell_texts = [
                str(cell.value).strip()
                for cell in row
                if cell.value is not None and str(cell.value).strip()
            ]
            if not cell_texts:
                continue

            key = (sheet_num, row_num)
            accepted = accepted_by_key.get(key)
            if not accepted:
                continue

            for cell in row:
                if cell.data_type == "f":
                    continue
                if isinstance(cell.value, str) and cell.value.strip():
                    updated = apply_all_accepted_replacements(cell.value, accepted)
                    if updated != cell.value:
                        cell.value = updated

    output = io.BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def rewrite_file(file_bytes: bytes, filename: str, accepted_by_key: dict) -> tuple[bytes, str, str]:
    """Return (output_bytes, output_filename, mime_type) for corrected copy."""
    base_name = filename.rsplit(".", 1)[0]
    ext = filename.lower().rsplit(".", 1)[-1]

    if ext == "docx":
        return rewrite_docx(file_bytes, accepted_by_key), f"{base_name}_corrected.docx", DOCX_MIME
    if ext == "pptx":
        return rewrite_pptx(file_bytes, accepted_by_key), f"{base_name}_corrected.pptx", PPTX_MIME
    if ext == "xlsx":
        return rewrite_xlsx(file_bytes, accepted_by_key), f"{base_name}_corrected.xlsx", XLSX_MIME
    if ext == "pdf":
        parsed_lines = parse_file(file_bytes, filename)
        corrected_text = build_corrected_text(parsed_lines, accepted_by_key)
        return corrected_text.encode("utf-8"), f"{base_name}_corrected.txt", "text/plain"

    raise ValueError(
        f"Unsupported file type: .{ext}. Supported: .docx, .pdf, .pptx, .xlsx"
    )
