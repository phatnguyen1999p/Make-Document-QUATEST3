"""
Unit tests for PDF document ID verification (English vs. Vietnamese).
"""

import os
import tempfile
import unittest
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from core.pdf_service import PdfService, detect_language


def create_pdf(file_path: str, pages_content: list):
    """Create a PDF with specified text content on each page."""
    c = canvas.Canvas(file_path, pagesize=letter)
    for idx, page_lines in enumerate(pages_content):
        if idx > 0:
            c.showPage()
        y = 750
        for line in page_lines:
            c.drawString(72, y, line)
            y -= 30
    c.save()


class TestPdfIdVerification(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_language_detection(self):
        self.assertEqual(detect_language("Doc - E.pdf"), "EN")
        self.assertEqual(detect_language("Doc - EN.docx"), "EN")
        self.assertEqual(detect_language("Doc - ENG.docx"), "EN")
        self.assertEqual(detect_language("Doc - e.pdf"), "EN")
        self.assertEqual(detect_language("Doc.docx"), "VN")
        self.assertEqual(detect_language("Doc.xlsx", sheet_names=["Sheet1", "Report - E"]), "EN")
        self.assertEqual(detect_language("Doc.xlsx", sheet_names=["Sheet1", "BaoCao"]), "VN")

    def test_english_valid_single_page(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - EN.pdf")
        pages = [
            [
                "Header No.: Q326D000001/E",
                "Date: 20/08/2026",
                "Body: Pursuant to Request Form Q326D000001",
                "Results: Pass"
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertEqual(result, "ID: OK [Q326D000001/E]")

    def test_english_valid_multipage(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - E.pdf")
        pages = [
            [
                "No.: Q326D000001/E",
                "Date: 20/08/2026",
                "Body: Reference Q326D000001"
            ],
            [
                "No.: Q326D000001/E",
                "Page 2 content details..."
            ],
            [
                "No.: Q326D000001/E",
                "Page 3 final conclusions..."
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertEqual(result, "ID: OK [Q326D000001/E]")

    def test_english_missing_e_on_page2(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - EN.pdf")
        pages = [
            [
                "No.: Q326D000001/E",
                "Date: 20/08/2026",
                "Body: Reference Q326D000001"
            ],
            [
                "No.: Q326D000001",  # Forgot /E on page 2
                "Page 2 content details..."
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertIn("P2(thiếu /E header)", result)

    def test_english_missing_body_base_id_on_page1(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - EN.pdf")
        pages = [
            [
                "No.: Q326D000001/E",
                "Date: 20/08/2026",
                "Body: General test without referencing base ID"
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertIn("P1(thiếu số gốc body)", result)

    def test_english_missing_e_on_page1_header(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - EN.pdf")
        pages = [
            [
                "No.: Q326D000001",  # Missing /E
                "Date: 20/08/2026",
                "Body: Reference Q326D000001"
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertIn("P1(thiếu /E header)", result)

    def test_vietnamese_valid_multipage(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report.pdf")
        pages = [
            [
                "Số: Q326D000001",
                "Ngày: 20/08/2026",
                "Căn cứ phiếu yêu cầu số Q326D000001"
            ],
            [
                "Số: Q326D000001",
                "Nội dung trang 2"
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="VN")
        self.assertEqual(result, "ID: OK [Q326D000001]")

    def test_vietnamese_with_extraneous_e(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report.pdf")
        pages = [
            [
                "Số: Q326D000001/E",  # Extraneous /E in VN version
                "Ngày: 20/08/2026",
                "Căn cứ phiếu Q326D000001"
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="VN")
        self.assertIn("P1(thừa /E)", result)

    def test_mismatched_ids_across_pages(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - EN.pdf")
        pages = [
            [
                "No.: Q326D000001/E",
                "Date: 20/08/2026",
                "Body: Reference Q326D000001"
            ],
            [
                "No.: Q326D000002/E",  # Mismatch 000002 vs 000001
                "Page 2 content..."
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertIn("Không khớp mã", result)

    def test_missing_ids(self):
        pdf_path = os.path.join(self.temp_dir.name, "Empty.pdf")
        pages = [
            ["General document without any case registration numbers"]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path)
        self.assertEqual(result, "ID: MISSING")

    def test_subitem_en_id(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - EN.pdf")
        pages = [
            [
                "No.: Q326D000001/1/E",
                "Date: 20/08/2026",
                "Body: Ref Q326D000001/1"
            ],
            [
                "No.: Q326D000001/1/E",
                "Page 2..."
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertEqual(result, "ID: OK [Q326D000001/1/E]")

    def test_subitem_vn_id(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report.pdf")
        pages = [
            [
                "Số: Q326D000001/2",
                "Ngày: 20/08/2026",
                "Căn cứ phiếu Q326D000001"
            ],
            [
                "Số: Q326D000001/2",
                "Nội dung trang 2"
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="VN")
        self.assertEqual(result, "ID: OK [Q326D000001/2]")

    def test_revision_dot_r_vn(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report.pdf")
        pages = [
            [
                "Số: Q326D000001/1.R",
                "Ngày: 20/08/2026",
                "Căn cứ phiếu Q326D000001"
            ],
            [
                "Số: Q326D000001/1.R",
                "Nội dung trang 2"
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="VN")
        self.assertEqual(result, "ID: OK [Q326D000001/1.R]")

    def test_revision_dot_r_en(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - EN.pdf")
        pages = [
            [
                "No.: Q326D000001/1/E.R",
                "Date: 20/08/2026",
                "Body: Reference Q326D000001"
            ],
            [
                "No.: Q326D000001/1/E.R",
                "Page 2..."
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertEqual(result, "ID: OK [Q326D000001/1/E.R]")

    def test_header_mismatch_across_pages(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report - EN.pdf")
        pages = [
            [
                "No.: Q326D000001/1/E.R",
                "Date: 20/08/2026",
                "Body: Reference Q326D000001"
            ],
            [
                "No.: Q326D000001/2/E.R",  # Header mismatch between page 1 (/1) and page 2 (/2)
                "Page 2..."
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="EN")
        self.assertIn("P2(header khác: Q326D000001/2/E.R)", result)

    def test_vietnamese_spaced_id_multipage(self):
        pdf_path = os.path.join(self.temp_dir.name, "Report.pdf")
        pages = [
            [
                "Số: Q326D0 84627 /3 - Ngày 25/08/2026",
                "2. Văn bản yêu cầu giám định: Q326D084627 ngày 13/07/2026"
            ],
            [
                "Phụ lục",
                "Số: Q326D0 84627 /3 - Ngày 25/08/2026",
                "Nội dung trang 2..."
            ]
        ]
        create_pdf(pdf_path, pages)
        result = PdfService.verify_document_number(pdf_path, lang="VN")
        self.assertEqual(result, "ID: OK [Q326D084627/3]")

    def test_testdata_bug_pdf(self):
        pdf_path = os.path.abspath("TestData_Bug/84627 - TBKQ - 3 - vn - N2.pdf")
        if os.path.exists(pdf_path):
            result = PdfService.verify_document_number(pdf_path, lang="VN")
            self.assertEqual(result, "ID: OK [Q326D084627/3]")


if __name__ == "__main__":
    unittest.main()
