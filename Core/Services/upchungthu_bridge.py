# -*- coding: utf-8 -*-
"""
UpChungThu Bridge Service.
Integrates the UpChungThu software (located at C:\\Users\\phat.nt\\Desktop\\Source\\Source Code\\UpChungThu\\UpChungThu)
with Làm KQ pipeline to:
1. Automatically overlay official QUATEST 3 letterhead background ('GIAY TIEU DE.pdf' / 'GIAY TIEU DE - E.pdf')
   onto Page 1 of TBKQ (Thông Báo Kết Quả) deliverables.
2. Run UpChungThu's comprehensive error-checking mechanisms:
   - Signer verification (GDV & Boss matching against LAN/local signer library)
   - Document registration number formatting (Q3 pattern, header consistency, English /E, Vietnamese no /E)
   - Page numbering sequence (i/total on multi-page appendices)
   - Date verification across all pages.
"""

import os
import sys
import re
from datetime import datetime
from typing import Dict, Any, Optional, Tuple, List
from PyPDF2 import PdfReader, PdfWriter

_HOME = os.path.expanduser("~")
# Machine-independent default; override with env var UPCHUNGTHU_PATH.
DEFAULT_UPCHUNGTHU_DIR = os.path.join(_HOME, "Desktop", "Source", "Source Code", "UpChungThu", "UpChungThu")
# Bản chép kèm workspace (chỉ mã nguồn + giấy tiêu đề, 11/11/2026) — dùng khi máy không có bản gốc.
_WORKSPACE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BUNDLED_UPCHUNGTHU_DIR = os.path.join(_WORKSPACE, "CongCu", "UpChungThu")

_UPCHUNGTHU_INITIALIZED = False
_PDF_SERVICE = None
_DETECT_LANGUAGE = None
_CONFIG_MANAGER = None
_CONFIG_INSTANCE = None


def get_upchungthu_dir() -> str:
    """Finds and returns the resolved root path of UpChungThu."""
    env_dir = os.environ.get("UPCHUNGTHU_PATH")
    if env_dir and os.path.isdir(env_dir):
        return os.path.abspath(env_dir)

    candidates = [
        DEFAULT_UPCHUNGTHU_DIR,
        os.path.abspath(r"..\Source Code\UpChungThu\UpChungThu"),
        os.path.abspath(r"..\UpChungThu\UpChungThu"),
        os.path.join(_HOME, "Desktop", "Source", "UpChungThu", "UpChungThu"),
        BUNDLED_UPCHUNGTHU_DIR,
    ]
    for c in candidates:
        if os.path.isdir(c) and os.path.isfile(os.path.join(c, "GIAY TIEU DE.pdf")):
            return os.path.abspath(c)

    return DEFAULT_UPCHUNGTHU_DIR


def _init_upchungthu():
    """Dynamically imports UpChungThu modules and config."""
    global _UPCHUNGTHU_INITIALIZED, _PDF_SERVICE, _DETECT_LANGUAGE, _CONFIG_MANAGER, _CONFIG_INSTANCE
    if _UPCHUNGTHU_INITIALIZED and _PDF_SERVICE is not None:
        return

    uct_dir = get_upchungthu_dir()
    if not os.path.isdir(uct_dir):
        raise FileNotFoundError(f"UpChungThu directory not found at: {uct_dir}")

    if uct_dir not in sys.path:
        sys.path.insert(0, uct_dir)

    try:
        from core.pdf_service import PdfService, detect_language
        from core.config_manager import ConfigManager

        _PDF_SERVICE = PdfService
        _DETECT_LANGUAGE = detect_language
        _CONFIG_MANAGER = ConfigManager
        _CONFIG_INSTANCE = ConfigManager(uct_dir)
        _UPCHUNGTHU_INITIALIZED = True
    except Exception as e:
        raise ImportError(f"Failed to import UpChungThu core modules from {uct_dir}: {e}")


def get_config():
    """Returns the loaded UpChungThu ConfigManager instance (with LAN signers)."""
    _init_upchungthu()
    return _CONFIG_INSTANCE


def get_background_paths() -> Tuple[str, str]:
    """Returns (bg_path_vn, bg_path_en) absolute paths."""
    uct_dir = get_upchungthu_dir()
    bg_vn = os.path.join(uct_dir, "GIAY TIEU DE.pdf")
    bg_en = os.path.join(uct_dir, "GIAY TIEU DE - E.pdf")

    if not os.path.isfile(bg_vn):
        candidate = os.path.join(uct_dir, "GIAY TIEU DE .pdf")
        if os.path.isfile(candidate):
            bg_vn = candidate

    return bg_vn, bg_en


def is_tbkq_file(file_name_or_path: str) -> bool:
    """Checks if a file represents a TBKQ deliverable (and not BCKQ or internal calculation)."""
    fn = os.path.basename(file_name_or_path).lower()
    if "bckq" in fn:
        return False
    if "tinhtoan" in fn or "calc" in fn:
        return False
    if "tem tu" in fn:
        return False
    return "tbkq" in fn


def detect_document_language(file_name_or_path: str) -> str:
    """Detects document language: 'VN' or 'EN'."""
    _init_upchungthu()
    return _DETECT_LANGUAGE(file_name_or_path)


def has_letterhead_background(pdf_path: str) -> bool:
    """Checks if Page 1 of the PDF already has the QUATEST 3 letterhead background."""
    try:
        reader = PdfReader(pdf_path)
        if len(reader.pages) == 0:
            return False
        p0 = reader.pages[0].extract_text() or ""
        has_vn = any(kw in p0 for kw in ["ỦY BAN TIÊU CHUẨN", "TRUNG TÂM KỸ THUẬT TIÊU CHUẨN", "COMMISSION FOR THE STANDARDS", "QUALITY ASSURANCE AND TESTING"])
        has_addr = "49 Pasteur" in p0
        return has_vn or has_addr
    except Exception:
        return False


def extract_cert_date_from_pdf(pdf_path: str) -> Optional[datetime]:
    """Extracts issue date from PDF text (e.g. 'Ngày 22/10/2026' or 'Date: 22/10/2026')."""
    try:
        reader = PdfReader(pdf_path)
        for p in reader.pages:
            txt = p.extract_text() or ""
            m = re.search(r"(?:Ngày|Date)[:\s\-]+(\d{1,2}/\d{1,2}/\d{4})", txt, re.IGNORECASE)
            if m:
                d_str = m.group(1)
                return datetime.strptime(d_str, "%d/%m/%Y")
    except Exception as _err:
        print(f"[Warning] upchungthu_bridge.py:140 {type(_err).__name__}: {_err}")
    return None


def apply_tbkq_background(pdf_path: str, lang: Optional[str] = None, force: bool = False) -> Tuple[bool, str]:
    """
    Overlays the official QUATEST 3 letterhead background onto Page 1 of a TBKQ PDF.
    Returns (success_bool, message).
    """
    abs_pdf = os.path.abspath(pdf_path)
    if not os.path.isfile(abs_pdf):
        return False, f"File not found: {abs_pdf}"

    if not is_tbkq_file(abs_pdf):
        return False, "Not a TBKQ document (background skipped)"

    if not force and has_letterhead_background(abs_pdf):
        return False, "Background already present (skipped double overlay)"

    if not lang:
        lang = detect_document_language(abs_pdf)

    bg_vn, bg_en = get_background_paths()
    bg_file = bg_en if lang == "EN" else bg_vn

    if not os.path.isfile(bg_file):
        return False, f"Background template missing: {bg_file}"

    temp_out = abs_pdf + ".tmp_bg"
    try:
        reader = PdfReader(abs_pdf)
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

        with open(temp_out, "wb") as f:
            writer.write(f)

        # Replace original file atomically
        os.replace(temp_out, abs_pdf)
        return True, f"Background applied successfully ({lang}: {os.path.basename(bg_file)})"
    except Exception as e:
        if os.path.isfile(temp_out):
            try:
                os.remove(temp_out)
            except Exception as _err:
                print(f"[Warning] upchungthu_bridge.py:194 {type(_err).__name__}: {_err}")
        return False, f"Error applying background: {e}"


def smart_verify_dates(pdf_path: str, target_date: Optional[datetime] = None) -> str:
    """
    Wrapper around UpChungThu verify_pdf_dates that handles edge cases gracefully:
    If Page 1 has 2 occurrences because the inspection date ('Thời gian giám định')
    happened to be on the same date as the certificate issue date, and Pages 2+ strictly
    have 1 occurrence, it validates as OK.
    """
    _init_upchungthu()
    raw = _PDF_SERVICE.verify_pdf_dates(pdf_path, target_date)
    if "OK" in raw or "Bỏ qua" in raw:
        return "Date: OK"

    clean = raw.replace("Date:", "").strip()
    # Check if only Page 1 is reported as having 2 occurrences
    if "P1(2 lần - thừa)" in clean and not re.search(r"P[2-9]\d*", clean):
        try:
            reader = PdfReader(pdf_path)
            p1_txt = reader.pages[0].extract_text() or ""
            target_dt = target_date if target_date else extract_cert_date_from_pdf(pdf_path)
            if target_dt:
                d_str = target_dt.strftime("%d/%m/%Y")
                # Look for inspection date / application date in body matching the same date
                if re.search(rf"(?:Từ\s+ngày|From|nhận\s+giám\s+định\s+số.*?ngày)\s+{d_str}", p1_txt, re.IGNORECASE):
                    return "Date: OK (P1 có ngày giám định/tiếp nhận trùng ngày ban hành)"
        except Exception as _err:
            print(f"[Warning] upchungthu_bridge.py:223 {type(_err).__name__}: {_err}")

    return raw


def verify_tbkq_pdf(pdf_path: str, cert_date: Optional[Any] = None, lang: Optional[str] = None,
                    expected_insp: Optional[str] = None, expected_boss: Optional[str] = None) -> Dict[str, Any]:
    """
    Runs the 4 UpChungThu error-checking mechanisms on a TBKQ PDF:
    1. Signers (GDV & Boss against library)
    2. Document Registration Number (Q3 pattern, /E for English, base number in Page 1 body)
    3. Page Numbering (i/total on multi-page appendices)
    4. Date consistency across all pages
    """
    _init_upchungthu()
    abs_pdf = os.path.abspath(pdf_path)
    if not os.path.isfile(abs_pdf):
        return {"overall": "FAIL", "message": f"File not found: {abs_pdf}", "issues": ["File not found"]}

    if not lang:
        lang = detect_document_language(abs_pdf)

    # Parse target date
    target_dt = None
    if isinstance(cert_date, datetime):
        target_dt = cert_date
    elif isinstance(cert_date, str) and cert_date.strip():
        try:
            target_dt = datetime.strptime(cert_date.strip(), "%d/%m/%Y")
        except Exception as _err:
            print(f"[Warning] upchungthu_bridge.py:253 {type(_err).__name__}: {_err}")
    if target_dt is None:
        target_dt = extract_cert_date_from_pdf(abs_pdf)

    cfg = get_config()
    issues: List[str] = []

    # 1. Signers verification
    raw_signer, d_insp, d_boss = _PDF_SERVICE.verify_signers(
        abs_pdf, cfg.inspector_library, cfg.boss_library, expected_insp, expected_boss
    )
    if "OK" not in raw_signer and "Bỏ qua" not in raw_signer:
        clean = raw_signer.replace("Signer:", "").strip()
        issues.append(f"Chữ ký ({clean})")

    # 2. Document Number verification
    raw_id = _PDF_SERVICE.verify_document_number(abs_pdf, lang=lang)
    if "OK" not in raw_id and "Bỏ qua" not in raw_id:
        clean = raw_id.replace("ID:", "").strip()
        issues.append(f"Mã hồ sơ ({clean})")

    # 3. Page Numbering verification
    raw_page = _PDF_SERVICE.verify_page_numbers(abs_pdf)
    if "OK" not in raw_page and "N/A" not in raw_page and "Bỏ qua" not in raw_page:
        clean = raw_page.replace("Page:", "").strip()
        issues.append(f"Số trang ({clean})")

    # 4. Date verification
    raw_date = smart_verify_dates(abs_pdf, target_dt)
    if "OK" not in raw_date and "Bỏ qua" not in raw_date:
        clean = raw_date.replace("Date:", "").strip()
        issues.append(f"Ngày ({clean})")

    status = "OK" if not issues else "FAIL"
    message = "✅ Tất cả kiểm tra hợp lệ" if status == "OK" else f"❌ {', '.join(issues)}"

    return {
        "status": status,
        "message": message,
        "lang": lang,
        "signers": (d_insp, d_boss),
        "signer_status": raw_signer,
        "id_status": raw_id,
        "page_status": raw_page,
        "date_status": raw_date,
        "target_date": target_dt.strftime("%d/%m/%Y") if target_dt else "N/A",
        "issues": issues,
    }


def process_tbkq_pdf(pdf_path: str, cert_date: Optional[Any] = None, add_bg: bool = True,
                     force_bg: bool = False, print_log: bool = True) -> Dict[str, Any]:
    """
    Combined pipeline:
    1. Detects document language (VN or EN).
    2. Overlays the official QUATEST 3 letterhead background on Page 1 (if add_bg=True).
    3. Runs the 4 comprehensive QA checks from UpChungThu.
    4. Outputs structured log to console.
    """
    abs_pdf = os.path.abspath(pdf_path)
    fname = os.path.basename(abs_pdf)
    lang = detect_document_language(abs_pdf)

    bg_applied = False
    bg_msg = ""
    if add_bg and is_tbkq_file(abs_pdf):
        bg_applied, bg_msg = apply_tbkq_background(abs_pdf, lang=lang, force=force_bg)

    qa_result = verify_tbkq_pdf(abs_pdf, cert_date=cert_date, lang=lang)
    qa_result["bg_applied"] = bg_applied
    qa_result["bg_message"] = bg_msg
    qa_result["file"] = abs_pdf

    if print_log:
        insp, boss = qa_result["signers"]
        status_icon = "✅ [UpChungThu QA: PASSED]" if qa_result["status"] == "OK" else "❌ [UpChungThu QA: ISSUES]"
        bg_icon = f"🖼️ BG ({lang}): {bg_msg}" if bg_applied else f"🖼️ BG: {bg_msg or 'Skipped'}"

        print(f"  {status_icon} {fname} ({lang})")
        print(f"     - {bg_icon}")
        print(f"     - Chữ ký: GDV='{insp or '?'}' | Boss='{boss or '?'}' ({qa_result['signer_status']})")
        print(f"     - Mã hồ sơ: {qa_result['id_status']}")
        print(f"     - Phụ lục trang: {qa_result['page_status']}")
        print(f"     - Ngày ban hành: {qa_result['date_status']}")
        if qa_result["issues"]:
            print(f"     - Cảnh báo: {', '.join(qa_result['issues'])}")

    return qa_result
