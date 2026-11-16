"""
Word automation service using MS Word COM interface (win32com).
Handles document date updating in headers/footers/shapes/appendices and exporting to PDF.
"""

import os
import re
from datetime import datetime
from typing import Dict, Any, Optional, Callable
import pythoncom
import win32com.client


class WordEngine:
    """
    Service for MS Word COM automation tasks.
    """

    PATTERNS = {
        "%d/%m/%Y": r'\b\d{1,2}/\d{1,2}/\d{4}\b',
        "%B %d, %Y": r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)[\s\xa0]+\d{1,2},[\s\xa0]+\d{4}\b',
        "VN_Format": r'(?i)\bng[àa]y[\s\xa0]+\d{1,2}[\s\xa0]+th[áa]ng[\s\xa0]+\d{1,2}[\s\xa0]+n[ăa]m[\s\xa0]+\d{4}\b'
    }

    APPENDIX_PATTERN = re.compile(
        r'(?i)(?:Issued together with|Ban hành kèm theo)[\s\S]{1,150}?(?:dated|ngày)\s+(\d{1,2}/\d{1,2}/\d{4})'
    )

    @classmethod
    def update_word_date(cls, doc, target_time: datetime) -> str:
        """
        Scan and replace dates strictly in Word document headers.
        """
        # 1. Extract text strictly from Headers
        header_text = ""
        for section in doc.Sections:
            for header in section.Headers:
                if header.Exists:
                    header_text += header.Range.Text + " "
                    try:
                        for shape in header.Shapes:
                            if shape.TextFrame.HasText:
                                header_text += shape.TextFrame.TextRange.Text + " "
                    except Exception:
                        pass

        def extract_dates(text_to_search: str):
            results = []
            for fmt, reg in cls.PATTERNS.items():
                for m in re.findall(reg, text_to_search):
                    try:
                        clean_m = re.sub(r'\s+', ' ', m.replace('\xa0', ' ')).strip()
                        if fmt == "VN_Format":
                            nums = re.findall(r'\d+', clean_m)
                            if len(nums) == 3:
                                d_obj = datetime(int(nums[2]), int(nums[1]), int(nums[0]))
                                results.append({"text": m, "date": d_obj, "fmt": fmt})
                        else:
                            d_obj = datetime.strptime(clean_m, fmt)
                            results.append({"text": m, "date": d_obj, "fmt": fmt})
                    except Exception:
                        continue
            return results

        # 2. Only extract dates found in header
        header_dates = extract_dates(header_text)
        if not header_dates:
            return "➖ Bỏ qua (Không có ngày ở Header)"

        unique_targets = {}
        for t in header_dates:
            unique_targets[t['text']] = t['fmt']

        replacement_count = 0
        current_time = target_time
        needs_update = False

        # 3. Perform Find & Replace strictly in Headers
        for target_text, fmt in unique_targets.items():
            if fmt == "%d/%m/%Y":
                new_val = current_time.strftime("%d/%m/%Y")
            elif fmt == "VN_Format":
                new_val = f"Ngày {current_time.strftime('%d')} tháng {current_time.strftime('%m')} năm {current_time.strftime('%Y')}"
            else:
                new_val = current_time.strftime("%B %d, %Y")

            if target_text.lower() == new_val.lower():
                continue

            needs_update = True
            target_for_word = target_text.replace('\xa0', '^w').replace(' ', '^w')

            def replace_in_range(rng):
                nonlocal replacement_count
                try:
                    rng.Find.ClearFormatting()
                    rng.Find.Replacement.ClearFormatting()
                    # Forward=True, Wrap=0 (wdFindStop), Replace=2 (wdReplaceAll)
                    # Wrap=0 ensures the search stays strictly within the specified Range
                    if rng.Find.Execute(target_for_word, False, False, False, False, False, True, 0, False, new_val, 2):
                        replacement_count += 1
                except Exception:
                    pass

            # ONLY replace in Section Headers (and text boxes inside headers)
            for section in doc.Sections:
                for header in section.Headers:
                    if header.Exists:
                        replace_in_range(header.Range)
                        try:
                            for shape in header.Shapes:
                                if shape.TextFrame.HasText:
                                    replace_in_range(shape.TextFrame.TextRange)
                        except Exception:
                            pass

        if not needs_update:
            return "✅ Ngày đã đúng"
        elif replacement_count > 0:
            doc.Save()
            return f"✅ Đã cập nhật ({replacement_count} chỗ)"
        else:
            return "❌ Lỗi Find"

    @staticmethod
    def export_to_pdf(doc, pdf_path: str):
        """
        Export opened Word document to PDF using FileFormat=17 (wdFormatPDF).
        """
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
