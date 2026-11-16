# -*- coding: utf-8 -*-
"""
Core/Tools/tax_helper.py
Tra cứu thông tin doanh nghiệp từ Mã số thuế (MST) và Đối soát sai lệch với Biên bản hiện trường.
Nguồn dữ liệu: Tổng cục Thuế (gdt.gov.vn) qua Open API Gateway.
Tuân thủ nghiêm ngặt nguyên tắc chung:
- Đồng bộ song ngữ (tiếng Việt & tiếng Anh).
- Đối soát sai lệch tên và địa chỉ giữa Biên bản và Dữ liệu thuế.
- Cảnh báo rõ ràng cho người dùng (User) khi phát hiện sai lệch để xử lý.
- Tự động fallback an toàn khi mất mạng hoặc không tìm thấy MST.
"""

import os
import sys
import re
import json
import time
import difflib
import unicodedata
import urllib.request
import urllib.error
from typing import Dict, Any, List, Optional

# Đảm bảo in tiếng Việt UTF-8 mượt mà trên Windows console
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception as _err:
        print(f"[Warning] tax_helper.py:28 {type(_err).__name__}: {_err}")

WORKSPACE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)

TAX_CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tax_cache.json")


def load_tax_cache() -> Dict[str, Any]:
    """Tải bộ nhớ đệm MST cục bộ (offline cache)."""
    if os.path.isfile(TAX_CACHE_PATH):
        try:
            with open(TAX_CACHE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_tax_cache(cache: Dict[str, Any]) -> None:
    """Lưu bộ nhớ đệm MST cục bộ."""
    try:
        with open(TAX_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[Warning] Không thể ghi tax_cache.json: {e}")


def remove_accents(input_str: str) -> str:
    """Xóa dấu tiếng Việt phục vụ so sánh văn bản OCR."""
    if not input_str:
        return ""
    s = unicodedata.normalize("NFD", input_str)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    return unicodedata.normalize("NFC", s)


def normalize_company_name(name: str, unaccent: bool = True) -> str:
    """
    Chuẩn hóa tên công ty:
    - Loại bỏ ký tự đặc biệt, dấu câu, khoảng trắng thừa.
    - Chuẩn hóa các từ viết tắt phổ biến trong doanh nghiệp Việt Nam.
    """
    if not name:
        return ""
    s = name.upper()
    if unaccent:
        s = remove_accents(s)
    # Loại bỏ ký hiệu dấu câu
    s = re.sub(r"[\.,\-\(\)\[\]\"'/:]", " ", s)

    # Chuẩn hóa từ viết tắt
    replacements = [
        (r"\bCTY\b", "CONG TY" if unaccent else "CÔNG TY"),
        (r"\bTNHH\b", "TRACH NHIEM HUU HAN" if unaccent else "TRÁCH NHIỆM HỮU HẠN"),
        (r"\bCP\b", "CO PHAN" if unaccent else "CỔ PHẦN"),
        (r"\bJSC\b", "CO PHAN" if unaccent else "CỔ PHẦN"),
        (r"\bLTD\b", "TRACH NHIEM HUU HAN" if unaccent else "TRÁCH NHIỆM HỮU HẠN"),
        (r"\bMTV\b", "MOT THANH VIEN" if unaccent else "MỘT THÀNH VIÊN"),
        (r"\bTMDV\b", "THUONG MAI DICH VU" if unaccent else "THƯƠNG MẠI DỊCH VỤ"),
        (r"\bTM\s*&\s*DV\b", "THUONG MAI DICH VU" if unaccent else "THƯƠNG MẠI DỊCH VỤ"),
        (r"\bSX\b", "SAN XUAT" if unaccent else "SẢN XUẤT"),
        (r"\bXNK\b", "XUAT NHAP KHAU" if unaccent else "XUẤT NHẬP KHẨU"),
    ]
    for pattern, repl in replacements:
        s = re.sub(pattern, repl, s)

    s = re.sub(r"\s+", " ", s).strip()
    return s


def extract_core_keywords(norm_name: str) -> set:
    """Trích xuất các từ khóa cốt lõi định danh doanh nghiệp (loại bỏ loại hình công ty)."""
    stop_words = {
        "CONG", "TY", "TRACH", "NHIEM", "HUU", "HAN", "CO", "PHAN", "MOT", "THANH", "VIEN",
        "THUONG", "MAI", "DICH", "VU", "SAN", "XUAT", "DAU", "TU", "VIET", "NAM"
    }
    words = set(norm_name.split())
    core = words - stop_words
    return core if core else words


def lookup_tax_code(tax_code: str, timeout: int = 8, max_retries: int = 2) -> Optional[Dict[str, Any]]:
    """
    Tra cứu thông tin Mã số thuế qua cổng API Open Gateway (Tổng cục Thuế).
    Hỗ trợ kiểm tra bộ nhớ đệm cục bộ (Cache) trước để chạy tức thì (0ms) và tránh bị rate limit.
    """
    clean_mst = re.sub(r"[^0-9\-]", "", tax_code.strip())
    if not clean_mst:
        return None

    # 1. Kiểm tra cache cục bộ trước
    cache = load_tax_cache()
    if clean_mst in cache:
        res = dict(cache[clean_mst])
        res["from_cache"] = True
        return res

    url = f"https://api.vietqr.io/v2/business/{clean_mst}"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept": "application/json"
        }
    )

    for attempt in range(max_retries + 1):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                if data.get("code") == "00" and data.get("data"):
                    d = data["data"]
                    info = {
                        "tax_code": clean_mst,
                        "name_vn": d.get("name", "").strip(),
                        "name_en": (d.get("internationalName") or "").strip(),
                        "short_name": (d.get("shortName") or "").strip(),
                        "address": d.get("address", "").strip(),
                        "status": d.get("status", "").strip(),
                        "source": "Tổng cục Thuế (gdt.gov.vn)"
                    }
                    # Lưu vào cache cục bộ
                    cache[clean_mst] = info
                    save_tax_cache(cache)
                    info["from_cache"] = False
                    return info
                elif data.get("code") in ("51", "52"):
                    return {
                        "tax_code": clean_mst,
                        "error": "Mã số thuế không tồn tại hoặc không hợp lệ",
                        "code": data.get("code")
                    }
        except urllib.error.HTTPError as he:
            if he.code == 429 and attempt < max_retries:
                # Quá số lượng yêu cầu -> nghỉ 1.5s rồi thử lại
                time.sleep(1.5)
                continue
            return {
                "tax_code": clean_mst,
                "error": f"Máy chủ tra cứu thuế phản hồi HTTP {he.code}",
                "is_network_error": True
            }
        except Exception as e:
            if attempt < max_retries:
                time.sleep(1.0)
                continue
            return {
                "tax_code": clean_mst,
                "error": f"Không thể kết nối đến máy chủ tra cứu thuế ({e})",
                "is_network_error": True
            }

    return None


def compare_client_names(name_doc: str, name_tax: str) -> Dict[str, Any]:
    """
    Đối soát so sánh tên ghi trên Biên bản/Yêu cầu với Tên chính thức theo MST.
    Trả về tỷ lệ trùng khớp, phân loại mức độ và khuyến nghị xử lý.
    """
    if not name_doc or not name_tax:
        return {
            "status": "UNKNOWN",
            "score": 0.0,
            "message": "Thiếu dữ liệu tên để đối soát."
        }

    # 1. So sánh nguyên bản sau khi bỏ dấu cách thừa
    clean_doc = re.sub(r"\s+", " ", name_doc.strip())
    clean_tax = re.sub(r"\s+", " ", name_tax.strip())

    if clean_doc.upper() == clean_tax.upper():
        return {
            "status": "EXACT_MATCH",
            "score": 1.0,
            "message": "Trùng khớp chính xác 100% (cả dấu tiếng Việt và loại hình)."
        }

    # 2. Chuẩn hóa không dấu và từ viết tắt
    norm_doc = normalize_company_name(clean_doc, unaccent=True)
    norm_tax = normalize_company_name(clean_tax, unaccent=True)

    if norm_doc == norm_tax:
        return {
            "status": "NORMALIZED_MATCH",
            "score": 0.99,
            "message": "Trùng khớp hoàn toàn sau khi chuẩn hóa viết tắt/dấu tiếng Việt (OCR scan)."
        }

    # 3. Tính độ tương đồng chuỗi (SequenceMatcher)
    seq_ratio = difflib.SequenceMatcher(None, norm_doc, norm_tax).ratio()

    # 4. Tính độ phủ từ khóa cốt lõi (Core Keywords Overlap)
    core_doc = extract_core_keywords(norm_doc)
    core_tax = extract_core_keywords(norm_tax)
    overlap = len(core_doc & core_tax) / max(len(core_doc | core_tax), 1)

    # Đánh giá tổng hợp
    if seq_ratio >= 0.88 or (overlap >= 0.8 and seq_ratio >= 0.75):
        status = "MATCH_WITH_TYPO"
        msg = f"Tương đồng cao ({seq_ratio:.1%}). Có thể do lỗi gõ máy hoặc sai sót nhỏ về dấu/chữ."
    elif overlap >= 0.5 or (core_tax and core_doc.issubset(core_tax)):
        status = "PARTIAL_MATCH"
        msg = f"Trùng một phần tên cốt lõi ({seq_ratio:.1%}, từ khóa: {', '.join(core_doc & core_tax)}). Biên bản có thể ghi vắn tắt hoặc ghi tên chi nhánh/nhà máy."
    else:
        status = "MISMATCH"
        msg = f"SAI LỆCH NGHIÊM TRỌNG ({seq_ratio:.1%}). Tên trên biên bản khác biệt hoàn toàn với tên pháp nhân của Mã số thuế."

    return {
        "status": status,
        "score": round(seq_ratio, 3),
        "keyword_overlap": round(overlap, 3),
        "norm_doc": norm_doc,
        "norm_tax": norm_tax,
        "message": msg
    }


def detect_tax_codes(text: str) -> List[str]:
    """Tự động tìm kiếm mã số thuế trong văn bản OCR hoặc tài liệu."""
    if not text:
        return []
    # Dạng chuẩn: 10 chữ số hoặc 10 chữ số gạch nối 3 chữ số
    patterns = [
        r"(?:mã số thuế|mst|tax\s*code|m\.s\.t)[\s:\.]*([0-9]{10}(?:-[0-9]{3})?)",
        r"\b([0-9]{10}-[0-9]{3})\b",
        r"\b([0-9]{10})\b"
    ]
    found = []
    for pat in patterns:
        matches = re.findall(pat, text, re.IGNORECASE)
        for m in matches:
            # Lọc bỏ các số năm hoặc số điện thoại thông thường
            if len(m) == 10 and (m.startswith("202") or m.startswith("199") or m.startswith("09") or m.startswith("08")):
                # 09, 08 có thể là MST (Hà Nội/TPHCM), kiểm tra ngữ cảnh nếu có tiền tố MST
                if "mst" in text.lower() or "mã số thuế" in text.lower():
                    if m not in found:
                        found.append(m)
            else:
                if m not in found:
                    found.append(m)
    return found


def verify_customer_with_tax(
    name_on_doc: str,
    tax_code: str,
    address_on_doc: str = ""
) -> Dict[str, Any]:
    """
    Quy trình kiểm tra trọn gói:
    1. Tra cứu MST từ nguồn dữ liệu thuế chính thức.
    2. Đối soát tên khách hàng trên biên bản với dữ liệu thuế.
    3. Đối soát địa chỉ (nếu có).
    4. Trả về báo cáo kết quả và khuyến nghị hành động cho User.
    """
    res: Dict[str, Any] = {
        "tax_code": tax_code,
        "name_on_doc": name_on_doc,
        "address_on_doc": address_on_doc,
        "tax_data": None,
        "name_check": None,
        "has_discrepancy": False,
        "alert_level": "INFO",  # INFO, WARNING, CRITICAL
        "recommendation": "",
        "suggested_client_vn": name_on_doc,
        "suggested_client_en": "",
        "suggested_address_vn": address_on_doc
    }

    # 1. Tra cứu thuế
    tax_info = lookup_tax_code(tax_code)
    if not tax_info or tax_info.get("error"):
        res["alert_level"] = "WARNING"
        res["has_discrepancy"] = True
        err_msg = tax_info.get("error") if tax_info else "Không tìm thấy thông tin"
        res["recommendation"] = f"Lỗi tra cứu MST {tax_code}: {err_msg}. Vui lòng kiểm tra lại MST trên hồ sơ gốc."
        return res

    res["tax_data"] = tax_info
    name_tax = tax_info.get("name_vn", "")
    name_tax_en = tax_info.get("name_en", "")
    addr_tax = tax_info.get("address", "")

    # 2. Đối soát tên
    check = compare_client_names(name_on_doc, name_tax)
    res["name_check"] = check

    if check["status"] in ("EXACT_MATCH", "NORMALIZED_MATCH"):
        res["has_discrepancy"] = False
        res["alert_level"] = "INFO"
        res["recommendation"] = "Tên khách hàng trùng khớp hoàn toàn. Tự động áp dụng tên chính thức theo MST và tên tiếng Anh quốc tế."
        res["suggested_client_vn"] = name_tax
        res["suggested_client_en"] = name_tax_en
    elif check["status"] == "MATCH_WITH_TYPO":
        res["has_discrepancy"] = True
        res["alert_level"] = "WARNING"
        res["recommendation"] = (
            f"Phát hiện sai lệch nhỏ giữa Biên bản ('{name_on_doc}') và MST ('{name_tax}'). "
            "Khuyến nghị: Sử dụng tên chính thức theo MST để đảm bảo giá trị pháp lý của chứng thư."
        )
        res["suggested_client_vn"] = name_tax
        res["suggested_client_en"] = name_tax_en
    elif check["status"] == "PARTIAL_MATCH":
        res["has_discrepancy"] = True
        res["alert_level"] = "WARNING"
        res["recommendation"] = (
            f"Tên trên Biên bản ghi vắn tắt ('{name_on_doc}') so với MST ('{name_tax}'). "
            "Báo lại User lựa chọn: [1] Cập nhật tên đầy đủ theo MST, hoặc [2] Giữ nguyên tên theo yêu cầu cụ thể của khách hàng."
        )
        res["suggested_client_vn"] = name_tax
        res["suggested_client_en"] = name_tax_en
    else:  # MISMATCH
        res["has_discrepancy"] = True
        res["alert_level"] = "CRITICAL"
        res["recommendation"] = (
            f"CẢNH BÁO ĐỎ: Tên trên Biên bản ('{name_on_doc}') KHÁC HOÀN TOÀN với chủ sở hữu MST ('{name_tax}'). "
            "Cần thông báo ngay cho User để xác minh: Có thể khách hàng điền nhầm MST hoặc đây là đơn vị ủy quyền/thuê địa điểm."
        )
        # Giữ nguyên tên biên bản để tránh ghi sai đối tượng khi chưa có xác nhận của User
        res["suggested_client_vn"] = name_on_doc

    # Đề xuất địa chỉ: Nếu trên biên bản chưa có hoặc ghi sơ sài, gợi ý địa chỉ MST
    if addr_tax and not address_on_doc:
        res["suggested_address_vn"] = addr_tax

    return res


def format_verification_report(result: Dict[str, Any]) -> str:
    """Định dạng báo cáo trực quan cho User xem trong terminal hoặc chat."""
    lines = []
    lines.append("=" * 72)
    lines.append("   KẾT QUẢ ĐỐI SOÁT THÔNG TIN KHÁCH HÀNG VỚI MÃ SỐ THUẾ (MST)")
    lines.append("=" * 72)
    lines.append(f"Mã số thuế tra cứu : {result.get('tax_code')}")
    lines.append(f"Mức độ cảnh báo    : [{result.get('alert_level')}]")
    lines.append("-" * 72)

    tax_data = result.get("tax_data")
    if tax_data and not tax_data.get("error"):
        lines.append("[DỮ LIỆU TỪ TỔNG CỤC THUẾ / MST]")
        lines.append(f"  • Tên pháp nhân (VN): {tax_data.get('name_vn')}")
        lines.append(f"  • Tên quốc tế (EN)  : {tax_data.get('name_en') or '(Chưa đăng ký)'}")
        lines.append(f"  • Tên viết tắt       : {tax_data.get('short_name') or '(Không)'}")
        lines.append(f"  • Địa chỉ ĐKKD       : {tax_data.get('address')}")
        lines.append(f"  • Tình trạng MST     : {tax_data.get('status')}")
    else:
        err = tax_data.get("error") if tax_data else "Không tìm thấy thông tin"
        lines.append(f"[LỖI TRA CỨU]: {err}")

    lines.append("-" * 72)
    lines.append("[THÔNG TIN TRÊN BIÊN BẢN HIỆN TRƯỜNG]")
    lines.append(f"  • Tên ghi nhận (Doc) : {result.get('name_on_doc')}")
    lines.append(f"  • Địa chỉ ghi nhận   : {result.get('address_on_doc') or '(Không ghi)'}")

    lines.append("-" * 72)
    check = result.get("name_check")
    if check:
        lines.append(f"[ĐỐI SOÁT]: {check.get('status')} (Độ tương đồng: {check.get('score', 0):.1%})")
        lines.append(f"  Chi tiết: {check.get('message')}")

    lines.append("-" * 72)
    lines.append(f"[KHUYẾN NGHỊ XỬ LÝ]:\n  👉 {result.get('recommendation')}")
    lines.append("=" * 72)
    return "\n".join(lines)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Tra cứu MST & Đối soát tên khách hàng")
    parser.add_argument("--mst", required=True, help="Mã số thuế cần tra cứu (ví dụ: 0315562387)")
    parser.add_argument("--name", default="", help="Tên khách hàng ghi trên biên bản hiện trường")
    parser.add_argument("--addr", default="", help="Địa chỉ ghi trên biên bản")
    args = parser.parse_args()

    # Chạy kiểm tra
    res = verify_customer_with_tax(args.name, args.mst, args.addr)
    print(format_verification_report(res))
