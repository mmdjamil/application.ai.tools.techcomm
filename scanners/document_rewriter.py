"""
document_rewriter.py

Re-open the original uploaded document and apply accepted (term, replacement)
edits per (page, line) key, returning the modified document in its original
binary format.

Supported formats:
  .docx  → python-docx  (preserves fonts, headings, tables, images)
  .pptx  → python-pptx  (preserves layout, theme, images)
  .xlsx  → openpyxl     (preserves sheets, styling, formulas in other cells)
  .pdf   → plain-text fallback (in-place PDF rewriting is not supported)

The walk order in each rewriter mirrors the corresponding parser in
file_parser.py exactly so that (page, line) keys align correctly.
"""

import io

from docx import Document
from pptx import Presentation
from openpyxl import load_workbook

from .inclusive_scanner import apply_all_accepted_replacements
from .file_parser import ESTIMATED_LINES_PER_PAGE, parse_pdf


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _apply_runs(runs, new_text: str) -> None:
    """
    Write *new_text* into the first run and clear the remaining runs.

    A word can be split across multiple runs (formatting boundaries inside a
    paragraph).  The strategy here is: concatenate all runs into one string,
    do the replacement on that string, then write the result back into run[0]
    and zero out run[1:].  This preserves run[0]'s character formatting for
    the whole paragraph but discards intra-word formatting splits inside
    replaced words — an acceptable trade-off.
    """
    if not runs:
        return
    runs[0].text = new_text
    for run in runs[1:]:
        run.text = ""


# ---------------------------------------------------------------------------
# Format-specific rewriters
# ---------------------------------------------------------------------------

def rewrite_docx(file_bytes: bytes, accepted_by_key: dict) -> bytes:
    """
    Re-open a .docx, apply accepted replacements per (page, line), return bytes.

    Walk order mirrors parse_docx: body paragraphs first, then table cells.
    Line numbering uses the same "non-empty after .strip()" rule as the parser
    to prevent (page, line) key drift.
    """
    doc = Document(io.BytesIO(file_bytes))
    line_num = 0

    # --- Body paragraphs (same order as parse_docx) ---
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        line_num += 1
        estimated_page = ((line_num - 1) // ESTIMATED_LINES_PER_PAGE) + 1
        key = (estimated_page, line_num)
        if key in accepted_by_key:
            run_text = "".join(run.text for run in para.runs)
            new_text = apply_all_accepted_replacements(run_text, accepted_by_key[key])
            if new_text != run_text:
                _apply_runs(para.runs, new_text)

    # --- Table cells (same order as parse_docx) ---
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if not cell.text.strip():
                    continue
                line_num += 1
                estimated_page = ((line_num - 1) // ESTIMATED_LINES_PER_PAGE) + 1
                key = (estimated_page, line_num)
                if key in accepted_by_key:
                    # Apply the replacement to each paragraph in the cell
                    # individually; non-inclusive terms are single words and
                    # won't span cell-paragraph boundaries.
                    for cell_para in cell.paragraphs:
                        if not cell_para.text.strip():
                            continue
                        run_text = "".join(run.text for run in cell_para.runs)
                        new_text = apply_all_accepted_replacements(
                            run_text, accepted_by_key[key]
                        )
                        if new_text != run_text:
                            _apply_runs(cell_para.runs, new_text)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def rewrite_pptx(file_bytes: bytes, accepted_by_key: dict) -> bytes:
    """
    Re-open a .pptx, apply accepted replacements per (page, line), return bytes.

    Walk order mirrors parse_pptx: for each slide, text frames then tables.
    Line counter resets to 0 at the start of each slide (page = slide number).
    """
    prs = Presentation(io.BytesIO(file_bytes))

    for slide_num, slide in enumerate(prs.slides, start=1):
        line_num = 0

        for shape in slide.shapes:
            # --- Text frames (text boxes, placeholders) ---
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    if not para.text.strip():
                        continue
                    line_num += 1
                    key = (slide_num, line_num)
                    if key in accepted_by_key:
                        run_text = "".join(run.text for run in para.runs)
                        new_text = apply_all_accepted_replacements(
                            run_text, accepted_by_key[key]
                        )
                        if new_text != run_text:
                            _apply_runs(para.runs, new_text)

            # --- Tables ---
            if hasattr(shape, "has_table") and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if not cell.text.strip():
                            continue
                        line_num += 1
                        key = (slide_num, line_num)
                        if key in accepted_by_key:
                            for cell_para in cell.text_frame.paragraphs:
                                if not cell_para.text.strip():
                                    continue
                                run_text = "".join(
                                    run.text for run in cell_para.runs
                                )
                                new_text = apply_all_accepted_replacements(
                                    run_text, accepted_by_key[key]
                                )
                                if new_text != run_text:
                                    _apply_runs(cell_para.runs, new_text)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def rewrite_xlsx(file_bytes: bytes, accepted_by_key: dict) -> bytes:
    """
    Re-open a .xlsx, apply accepted replacements per (page, line), return bytes.

    Opens in read-write / keep-formulas mode so untouched cells are preserved
    as-is (including formulas).  Walk order mirrors parse_xlsx: sheets in order,
    rows by enumerate index (row_num = actual 1-based position in the sheet,
    matching the parser's line numbering).
    """
    wb = load_workbook(io.BytesIO(file_bytes), read_only=False, data_only=False)

    for sheet_num, sheet_name in enumerate(wb.sheetnames, start=1):
        ws = wb[sheet_name]

        for row_num, row in enumerate(ws.iter_rows(), start=1):
            # Mirror the parser's non-empty row check
            has_content = any(
                cell.value is not None and str(cell.value).strip()
                for cell in row
            )
            if not has_content:
                continue

            key = (sheet_num, row_num)
            if key in accepted_by_key:
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str):
                        new_val = apply_all_accepted_replacements(
                            cell.value, accepted_by_key[key]
                        )
                        cell.value = new_val

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _rewrite_pdf_as_text(file_bytes: bytes, accepted_by_key: dict) -> bytes:
    """
    Generate a plain-text fallback for PDFs.

    True in-place PDF rewriting (font embedding, glyph positioning) is out of
    scope.  Instead, the extracted text is re-assembled with replacements applied
    and returned as UTF-8 plain text, mirroring the original flat-text export.
    """
    parsed_lines = parse_pdf(file_bytes)
    chunks = []
    current_page = None

    for entry in parsed_lines:
        key = (entry["page"], entry["line"])
        text = entry["text"]
        if key in accepted_by_key:
            text = apply_all_accepted_replacements(text, accepted_by_key[key])

        if entry["page"] != current_page:
            if chunks:
                chunks.append("")
            chunks.append(f"--- Page/Section {entry['page']} ---")
            current_page = entry["page"]

        chunks.append(f"L{entry['line']:>4}: {text}")

    result = "\n".join(chunks).strip() + "\n" if chunks else ""
    return result.encode("utf-8")


# ---------------------------------------------------------------------------
# Public router
# ---------------------------------------------------------------------------

def rewrite_file(
    file_bytes: bytes, filename: str, accepted_by_key: dict
) -> tuple:
    """
    Apply accepted (term, replacement) edits to *file_bytes* and return
    ``(output_bytes, output_filename, mime_type)``.

    For .pdf the function falls back to a plain-text export because in-place
    PDF rewriting is not supported.  The caller should surface a caption in
    the UI to explain this.
    """
    base = filename.rsplit(".", 1)[0]
    ext = filename.lower().rsplit(".", 1)[-1]

    if ext == "docx":
        return (
            rewrite_docx(file_bytes, accepted_by_key),
            f"{base}_corrected.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    if ext == "pptx":
        return (
            rewrite_pptx(file_bytes, accepted_by_key),
            f"{base}_corrected.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    if ext == "xlsx":
        return (
            rewrite_xlsx(file_bytes, accepted_by_key),
            f"{base}_corrected.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if ext == "pdf":
        return (
            _rewrite_pdf_as_text(file_bytes, accepted_by_key),
            f"{base}_corrected.txt",
            "text/plain",
        )

    raise ValueError(
        f"Unsupported file type: .{ext}. Supported: .docx, .pptx, .xlsx, .pdf"
    )
