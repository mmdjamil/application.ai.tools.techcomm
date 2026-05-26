"""
file_parser.py
Extracts text content line-by-line from Word, PDF, PowerPoint, and Excel files.
Returns a list of dicts: [{"page": int, "line": int, "text": str}, ...]
"""

import pdfplumber
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
import io


def parse_docx(file_bytes: bytes) -> list:
    """Parse a Word document. Treats each paragraph as a line."""
    doc = Document(io.BytesIO(file_bytes))
    results = []
    line_num = 0
    for i, para in enumerate(doc.paragraphs):
        line_num += 1
        if para.text.strip():
            results.append({
                "page": 1,
                "line": line_num,
                "text": para.text
            })
    # Also scan tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                line_num += 1
                if cell.text.strip():
                    results.append({
                        "page": 1,
                        "line": line_num,
                        "text": cell.text
                    })
    return results


def parse_pdf(file_bytes: bytes) -> list:
    """Parse a PDF document page-by-page, line-by-line."""
    results = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                for line_num, line in enumerate(text.split("\n"), start=1):
                    if line.strip():
                        results.append({
                            "page": page_num,
                            "line": line_num,
                            "text": line
                        })
    return results


def parse_pptx(file_bytes: bytes) -> list:
    """Parse a PowerPoint file. Each slide = 1 page."""
    prs = Presentation(io.BytesIO(file_bytes))
    results = []
    for slide_num, slide in enumerate(prs.slides, start=1):
        line_num = 0
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line_num += 1
                    if para.text.strip():
                        results.append({
                            "page": slide_num,
                            "line": line_num,
                            "text": para.text
                        })
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        line_num += 1
                        if cell.text.strip():
                            results.append({
                                "page": slide_num,
                                "line": line_num,
                                "text": cell.text
                            })
    return results


def parse_xlsx(file_bytes: bytes) -> list:
    """Parse an Excel file. Each sheet = 1 page."""
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    results = []
    for sheet_num, sheet_name in enumerate(wb.sheetnames, start=1):
        ws = wb[sheet_name]
        for row_num, row in enumerate(ws.iter_rows(values_only=True), start=1):
            cell_texts = [str(cell) for cell in row if cell is not None]
            if cell_texts:
                combined = " | ".join(cell_texts)
                results.append({
                    "page": sheet_num,
                    "line": row_num,
                    "text": combined
                })
    wb.close()
    return results


def parse_file(file_bytes: bytes, filename: str) -> list:
    """Route to the correct parser based on file extension."""
    ext = filename.lower().rsplit(".", 1)[-1]
    parsers = {
        "docx": parse_docx,
        "pdf": parse_pdf,
        "pptx": parse_pptx,
        "xlsx": parse_xlsx,
    }
    parser = parsers.get(ext)
    if parser is None:
        raise ValueError(f"Unsupported file type: .{ext}. Supported: .docx, .pdf, .pptx, .xlsx")
    return parser(file_bytes)