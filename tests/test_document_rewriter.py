import io
import unittest
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation
from pptx.util import Inches

from scanners.document_rewriter import rewrite_file


class DocumentRewriterTests(unittest.TestCase):
    def test_rewrite_file_docx_round_trip(self):
        doc = Document()
        doc.add_paragraph("Use blacklist and Whitelist.")

        source = io.BytesIO()
        doc.save(source)

        accepted_by_key = {
            (1, 1): [("blacklist", "denylist"), ("whitelist", "allowlist")]
        }
        output_bytes, output_filename, mime = rewrite_file(
            source.getvalue(), "sample.docx", accepted_by_key
        )

        self.assertEqual(output_filename, "sample_corrected.docx")
        self.assertEqual(
            mime,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        rewritten = Document(io.BytesIO(output_bytes))
        self.assertEqual(rewritten.paragraphs[0].text, "Use denylist and Allowlist.")

    def test_rewrite_file_pptx_round_trip(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        textbox.text_frame.paragraphs[0].text = "The master controls access."

        source = io.BytesIO()
        prs.save(source)

        accepted_by_key = {(1, 1): [("master", "primary / initiator")]}
        output_bytes, output_filename, mime = rewrite_file(
            source.getvalue(), "slides.pptx", accepted_by_key
        )

        self.assertEqual(output_filename, "slides_corrected.pptx")
        self.assertEqual(
            mime,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

        rewritten = Presentation(io.BytesIO(output_bytes))
        texts = []
        for shape in rewritten.slides[0].shapes:
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    if para.text.strip():
                        texts.append(para.text)
        self.assertIn("The primary controls access.", texts)

    def test_rewrite_file_xlsx_round_trip(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "slave"

        source = io.BytesIO()
        wb.save(source)

        accepted_by_key = {(1, 1): [("slave", "secondary / target")]}
        output_bytes, output_filename, mime = rewrite_file(
            source.getvalue(), "sheet.xlsx", accepted_by_key
        )

        self.assertEqual(output_filename, "sheet_corrected.xlsx")
        self.assertEqual(
            mime,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        rewritten = load_workbook(io.BytesIO(output_bytes), data_only=False)
        self.assertEqual(rewritten.active["A1"].value, "secondary")

    def test_rewrite_file_pdf_falls_back_to_text(self):
        accepted_by_key = {(1, 1): [("blacklist", "denylist")]}

        with patch(
            "scanners.document_rewriter.parse_file",
            return_value=[{"page": 1, "line": 1, "text": "blacklist item"}],
        ):
            output_bytes, output_filename, mime = rewrite_file(
                b"%PDF-1.4", "source.pdf", accepted_by_key
            )

        self.assertEqual(output_filename, "source_corrected.txt")
        self.assertEqual(mime, "text/plain")
        self.assertIn("denylist", output_bytes.decode("utf-8"))

    def test_docx_line_numbering_skips_empty_paragraphs(self):
        doc = Document()
        doc.add_paragraph("blacklist stays")
        doc.add_paragraph("   ")
        doc.add_paragraph("whitelist changes")

        source = io.BytesIO()
        doc.save(source)

        accepted_by_key = {(1, 2): [("whitelist", "allowlist")]}
        output_bytes, _, _ = rewrite_file(source.getvalue(), "line_guard.docx", accepted_by_key)

        rewritten = Document(io.BytesIO(output_bytes))
        self.assertEqual(rewritten.paragraphs[0].text, "blacklist stays")
        self.assertEqual(rewritten.paragraphs[2].text, "allowlist changes")


if __name__ == "__main__":
    unittest.main()
