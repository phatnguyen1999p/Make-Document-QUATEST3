"""
PDF processing and verification services:
- Background PDF overlay / watermark insertion
- Language detection (VN / EN)
- Signer verification (GDV & Boss) with accent and non-accent matching
- Document date verification
- Document ID format checking (Q3 pattern)
- Page numbering sequence checking
"""

import os
import re
from datetime import datetime
from typing import Tuple, Optional, List, Dict
from PyPDF2 import PdfReader, PdfWriter
from .utils import remove_accents


def detect_language(file_name_or_path: Optional[str], sheet_names: Optional[List[str]] = None) -> str:
    """
    Detect document language (VN or EN) based on file naming convention or sheet names.
    Pattern: '- E', '- EN', '- ENG'.
    """
    if file_name_or_path:
        base_name = os.path.basename(file_name_or_path)
        stem, _ = os.path.splitext(base_name)
        if re.search(r'-\s*(?:E|EN|ENG)(?:\b|_|\s|\d|$)', stem, re.IGNORECASE):
            return "EN"

    if sheet_names:
        if isinstance(sheet_names, str):
            sheet_names = [sheet_names]
        for s in sheet_names:
            if s and re.search(r'-\s*(?:E|EN|ENG)(?:\b|_|\s|\d|$)', s, re.IGNORECASE):
                return "EN"

    return "VN"


class PdfService:
    """
    Core service for merging PDF backgrounds and running comprehensive post-export checks.
    """

    @staticmethod
    def verify_signers(pdf_path: str, inspector_library: List[str], boss_library: List[str],
                       expected_insp: Optional[str] = None, expected_boss: Optional[str] = None) -> Tuple[str, Optional[str], Optional[str]]:
        """
        Verify presence and consistency of inspector (GDV) and boss signers in PDF text.
        """
        try:
            if not inspector_library and not boss_library:
                return "Signer: Not match", None, None

            reader = PdfReader(pdf_path)
            full_text = ""
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    full_text += t + " "

            # Normalize text by collapsing consecutive whitespaces
            full_text_clean = re.sub(r'\s+', ' ', full_text)

            def find_name(library: List[str]) -> Optional[str]:
                if not library:
                    return None
                for name in library:
                    p1 = re.sub(r'\s+', r'\\s+', name)
                    name_no_accent = remove_accents(name)
                    p2 = re.sub(r'\s+', r'\\s+', name_no_accent)

                    if (re.search(p1, full_text_clean, re.IGNORECASE | re.UNICODE) is not None or
                        re.search(p2, full_text_clean, re.IGNORECASE | re.UNICODE) is not None):
                        return name
                return None

            detected_insp = find_name(inspector_library)
            detected_boss = find_name(boss_library)

            # Consistency Check if expected signers are provided
            if expected_insp or expected_boss:
                is_match = True
                if expected_insp and detected_insp != expected_insp:
                    is_match = False
                if expected_boss and detected_boss != expected_boss:
                    is_match = False

                if is_match:
                    return "Signer: OK (Đồng nhất)", detected_insp, detected_boss
                else:
                    return f"Signer: KHÁC MẪU ({detected_insp or '?'}/{detected_boss or '?'})", detected_insp, detected_boss

            # Existence check (first file in batch)
            if detected_insp and detected_boss:
                return "Signer: OK", detected_insp, detected_boss
            else:
                miss = []
                if not detected_insp:
                    miss.append("GDV")
                if not detected_boss:
                    miss.append("Boss")
                return f"Signer: THIẾU ({', '.join(miss)})", detected_insp, detected_boss

        except Exception as e:
            return f"Signer: Err ({str(e)})", None, None

    @staticmethod
    def verify_pdf_dates(pdf_path: str, target_date: Optional[datetime] = None) -> str:
        """
        Verify document date across all pages in the PDF file.
        Strict rule: Each page must contain exactly 1 occurrence of the issue date.
        """
        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            if total_pages == 0:
                return "Date: N/A (0 trang)"

            now = target_date if target_date else datetime.now()
            d = now.day
            m = now.month
            yyyy = now.strftime('%Y')
            month_full = now.strftime('%B')
            month_abbr = now.strftime('%b')

            # Build comprehensive date pattern matching all standard formats
            date_pattern = re.compile(
                rf'(?i)(?:'
                rf'\bng[àa]y\s+0?{d}\s+th[áa]ng\s+0?{m}\s+n[ăa]m\s+{yyyy}\b|'
                rf'\b(?:{month_full}|{month_abbr})\s+0?{d}(?:st|nd|rd|th)?(?:[\s,]+){yyyy}\b|'
                rf'\b0?{d}(?:st|nd|rd|th)?\s+(?:{month_full}|{month_abbr})(?:[\s,]+){yyyy}\b|'
                rf'\b0?{d}\s*[/.-]\s*0?{m}\s*[/.-]\s*{yyyy}\b'
                rf')'
            )

            pages_with_issue = []
            total_occurrences = 0

            for idx, page in enumerate(reader.pages, 1):
                raw_txt = page.extract_text() or ""
                # Normalize spaces and line breaks
                clean_txt = re.sub(r'[\xa0\t\r\n]+', ' ', raw_txt)
                clean_txt = re.sub(r'\s+', ' ', clean_txt).strip()

                matches = date_pattern.findall(clean_txt)
                p_count = len(matches)
                total_occurrences += p_count

                if p_count != 1:
                    if p_count == 0:
                        pages_with_issue.append(f"P{idx}(0 lần - thiếu)")
                    else:
                        pages_with_issue.append(f"P{idx}({p_count} lần - thừa)")

            if not pages_with_issue:
                return "Date: OK"
            elif total_occurrences == 0:
                return f"Date: THIẾU ({total_pages} trang)"
            else:
                return f"Date: CHECK ({', '.join(pages_with_issue)} / {total_pages} trang)"
        except Exception as e:
            return f"Date: Err ({str(e)})"

    @staticmethod
    def verify_document_number(pdf_path: str, lang: Optional[str] = None) -> str:
        """
        Verify document registration number format (e.g., Q326A123456, Q326A123456/1, Q326A123456/1/E, Q326A123456/1/E.R).
        Rules:
        - Base number (e.g. Q326D000001) must always be present in the body of Page 1.
        - Multiple results / sub-reports add '/1', '/2', '/3' (e.g. Q326D000001/1 for VN, Q326D000001/1/E for EN).
        - Revision suffix '.R' (e.g. Q326D000001/1.R, Q326D000001/1/E.R) is valid and marked as OK.
        - All headers across every page in the document must have the exact same case number.
        - English version (lang == "EN"): Headers must contain '/E'.
        - Vietnamese version (lang == "VN"): Headers must NOT contain '/E'.
        """
        try:
            if not lang:
                lang = detect_language(pdf_path)

            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            if total_pages == 0:
                return "ID: N/A (0 trang)"

            pattern = re.compile(
                r'(?i)(?:^|[^A-Z0-9])(Q\s*3\s*\d\s*\d\s*[A-D](?:\s*\d){6}(?:\s*/\s*\d+)?(?:\s*/\s*E)?(?:\s*\.\s*R)?)(?![A-Z0-9])'
            )
            fallback_pattern = re.compile(r'Q3\d{2}[A-D]\d{6}(?:/\d+)?(?:/E)?(?:\.R)?')

            page_ids: Dict[int, List[str]] = {}
            all_found_ids: List[str] = []

            for idx, page in enumerate(reader.pages, 1):
                txt = page.extract_text() or ""
                t = txt.replace('\xa0', ' ')
                raw_matches = pattern.findall(t)
                matches = list(dict.fromkeys([re.sub(r'\s+', '', m).upper() for m in raw_matches]))

                # Also search cleaned text to ensure nothing is missed (e.g., unusual spacing/punctuation)
                clean = re.sub(r'[^a-zA-Z0-9/.]', '', txt).upper()
                clean = re.sub(r'\.R\b', '.R', clean)
                for fm in fallback_pattern.findall(clean):
                    if fm not in matches:
                        matches.append(fm)

                page_ids[idx] = matches
                all_found_ids.extend(matches)

            if not all_found_ids:
                return "ID: MISSING"

            def extract_root(raw_id: str) -> str:
                m = re.match(r'(Q3\d{2}[A-D]\d{6})', raw_id, re.I)
                return m.group(1).upper() if m else raw_id

            # Check root case number consistency across all extracted IDs
            all_roots = [extract_root(i) for i in all_found_ids]
            unique_roots = list(dict.fromkeys(all_roots))
            if len(unique_roots) > 1:
                return f"ID: CHECK (Không khớp mã gốc: {', '.join(unique_roots)})"

            root_id = unique_roots[0]
            issues = []

            if lang == "EN":
                # ENGLISH VERSION:
                # 1. Candidate header IDs must have /E
                en_header_candidates = [i for i in all_found_ids if '/E' in i]
                p1_en_candidates = [i for i in page_ids.get(1, []) if '/E' in i]

                if p1_en_candidates:
                    master_header = p1_en_candidates[0]
                elif en_header_candidates:
                    master_header = max(set(en_header_candidates), key=lambda x: (en_header_candidates.count(x), len(x)))
                else:
                    master_header = f"{root_id}/E"

                # Check that all headers across all pages match the master header
                for p in range(1, total_pages + 1):
                    p_ids = page_ids.get(p, [])
                    if master_header not in p_ids:
                        other_en = [i for i in p_ids if '/E' in i]
                        if other_en:
                            issues.append(f"P{p}(header khác: {other_en[0]})")
                        else:
                            non_en = [i for i in p_ids if extract_root(i) == root_id]
                            if non_en:
                                issues.append(f"P{p}(thiếu /E header)")
                            else:
                                issues.append(f"P{p}(thiếu ID)")

                # 2. Body check on Page 1: Must contain base number (without /E)
                p1_ids = page_ids.get(1, [])
                has_base_p1 = any('/E' not in i and extract_root(i) == root_id for i in p1_ids)
                if not has_base_p1:
                    issues.append("P1(thiếu số gốc body)")

                if not issues:
                    return f"ID: OK [{master_header}]"
                else:
                    return f"ID: CHECK ({', '.join(issues)})"

            else:
                # VIETNAMESE VERSION:
                # 1. No /E allowed on any page
                vn_candidates = [i for i in all_found_ids if '/E' not in i]
                if not vn_candidates:
                    return "ID: MISSING"

                p1_vn_candidates = [i for i in page_ids.get(1, []) if '/E' not in i]
                if p1_vn_candidates:
                    # Pick most specific candidate from Page 1 (e.g. /1.R over root)
                    master_header = max(p1_vn_candidates, key=len)
                else:
                    master_header = max(set(vn_candidates), key=lambda x: (vn_candidates.count(x), len(x)))

                # Check all pages for header consistency and no /E
                for p in range(1, total_pages + 1):
                    p_ids = page_ids.get(p, [])
                    en_ids = [i for i in p_ids if '/E' in i]
                    if en_ids:
                        issues.append(f"P{p}(thừa /E)")
                    elif master_header not in p_ids and root_id not in p_ids:
                        issues.append(f"P{p}(thiếu ID)")
                    elif master_header not in p_ids and p > 1:
                        other_vn = [i for i in p_ids if extract_root(i) == root_id]
                        if other_vn and other_vn[0] != master_header:
                            issues.append(f"P{p}(header khác: {other_vn[0]})")

                # 2. Body check on Page 1: Must contain base number
                p1_ids = page_ids.get(1, [])
                has_base_p1 = any(extract_root(i) == root_id for i in p1_ids)
                if not has_base_p1:
                    issues.append("P1(thiếu số gốc body)")

                if not issues:
                    return f"ID: OK [{master_header}]"
                else:
                    return f"ID: CHECK ({', '.join(issues)})"
        except Exception as e:
            return f"ID: Err ({str(e)})"

    @staticmethod
    def verify_page_numbers(pdf_path: str) -> str:
        """
        Verify consecutive page numbering format (e.g. 1/3, 2/3) across all pages after page 1.
        """
        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)

            if total_pages <= 1:
                return "Page: N/A"

            expected_total = total_pages - 1
            correct_count = 0
            missed_pages = []

            for i in range(1, total_pages):
                page = reader.pages[i]
                text = page.extract_text()

                if not text:
                    missed_pages.append(f"P{i+1}(Trống)")
                    continue

                text_clean = re.sub(r'\s+', '', text)
                current_num = i
                pattern = rf"{current_num}/{expected_total}"

                if re.search(pattern, text_clean):
                    correct_count += 1
                else:
                    missed_pages.append(f"P{i+1}")

            if correct_count == expected_total:
                return "Page: OK"
            else:
                return f"Page: {correct_count}/{expected_total} (Miss: {','.join(missed_pages)})"

        except Exception as e:
            return f"Page: Error ({str(e)})"

    @classmethod
    def process_pdf_post_export(cls, pdf_path: str, bg_path_vn: str, bg_path_en: str,
                                add_bg: bool = True, check_date: bool = True, check_id: bool = True,
                                check_page: bool = True, check_signer: bool = True,
                                target_date: Optional[datetime] = None,
                                inspector_library: Optional[List[str]] = None,
                                boss_library: Optional[List[str]] = None,
                                expected_insp: Optional[str] = None,
                                expected_boss: Optional[str] = None,
                                lang: Optional[str] = None) -> Tuple[str, str, Tuple[Optional[str], Optional[str]]]:
        """
        Full post-export PDF pipeline:
        1. Insert background template onto first page.
        2. Run date, ID, page numbering, and signer checks.
        3. Returns (overall_status_string, tag, (detected_insp, detected_boss)).
        """
        try:
            if not lang:
                lang = detect_language(pdf_path)

            bg_file = bg_path_en if lang == "EN" else bg_path_vn
            issues = []

            # 1. Overlay background on page 1
            if add_bg:
                if not os.path.exists(bg_file):
                    issues.append(f"Lỗi BG ({lang})")
                else:
                    reader = PdfReader(pdf_path)
                    writer = PdfWriter()
                    bg_reader = PdfReader(bg_file)
                    background = bg_reader.pages[0]

                    for i in range(len(reader.pages)):
                        page = reader.pages[i]
                        if i == 0:
                            background.merge_page(page)
                            writer.add_page(background)
                        else:
                            writer.add_page(page)

                    with open(pdf_path, "wb") as f:
                        writer.write(f)

            # 2. Date verification
            if check_date:
                raw_date = cls.verify_pdf_dates(pdf_path, target_date)
                if "OK" not in raw_date and "Bỏ qua" not in raw_date:
                    clean = raw_date.replace("Date:", "").strip()
                    issues.append(f"Ngày ({clean})")

            # 3. ID verification
            if check_id:
                raw_id = cls.verify_document_number(pdf_path, lang=lang)
                if "OK" not in raw_id and "Bỏ qua" not in raw_id:
                    clean = raw_id.replace("ID:", "").strip()
                    issues.append(f"ID ({clean})")

            # 4. Page numbering verification
            if check_page:
                raw_page = cls.verify_page_numbers(pdf_path)
                if "OK" not in raw_page and "N/A" not in raw_page and "Bỏ qua" not in raw_page:
                    clean = raw_page.replace("Page:", "").strip()
                    issues.append(f"Trang ({clean})")

            # 5. Signers verification
            d_insp, d_boss = None, None
            if check_signer and inspector_library is not None and boss_library is not None:
                raw_signer, d_insp, d_boss = cls.verify_signers(
                    pdf_path, inspector_library, boss_library, expected_insp, expected_boss
                )
                if "OK" not in raw_signer and "Bỏ qua" not in raw_signer:
                    clean = raw_signer.replace("Signer:", "").strip()
                    issues.append(f"Chữ ký ({clean})")

            # 6. Synthesize results
            if issues:
                overall_check = f"❌ {', '.join(issues)}"
                tag = "bad"
            else:
                any_check = add_bg or check_date or check_id or check_page or check_signer
                overall_check = "✅ OK" if any_check else "➖ N/A"
                tag = "good" if any_check else "ready"

            return overall_check, tag, (d_insp, d_boss)
        except Exception as e:
            return f"❌ Lỗi: {str(e)}", "bad", (None, None)
