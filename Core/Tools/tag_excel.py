# -*- coding: utf-8 -*-
"""
Dữ liệu điền TAG cho các bảng tính Excel mẫu (Core/Templates/TinhToan*_template.xlsx).

Các ô thông tin hồ sơ/khách hàng trong template là TAG {{ ... }} (danh mục: Core/Templates/TAGS.md).
Hàm build_* trong bsc_calc / fumehood_calc / khinen_calc gọi fill_workbook_tags(wb, data) ngay sau khi mở template,
sau đó vẫn ghi số liệu đo vào đúng ô như trước (công thức Excel giữ nguyên — AGENTS.md mục 11).
"""

import datetime
import re
from typing import Any, Dict

from ..Models.dossier import inspection_scope
from .tag_engine import render_xlsx


def _dt(v: Any):
    """'22/10/2026' -> datetime; datetime giữ nguyên; không đọc được -> chuỗi gốc."""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(v or ""))
    if m:
        try:
            return datetime.datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    return v or ""


def _tm(v: Any):
    if isinstance(v, datetime.time):
        return v
    m = re.search(r"(\d{1,2})[:h](\d{2})", str(v or ""))
    if m:
        try:
            return datetime.time(int(m.group(1)), int(m.group(2)))
        except ValueError:
            pass
    return v or ""


def _quoted_in(txt: Any) -> str:
    """'Measuring/ validating "Fume hood X"' -> '"Fume hood X"' (tên tủ tiếng Anh lấy từ nội dung yêu cầu)."""
    m = re.search(r'["“](.+?)["”]', str(txt or ""))
    return f'"{m.group(1)}"' if m else ""


def _scope(txt: Any, is_en: bool, is_cabinet: Any) -> str:
    """Nội dung yêu cầu: phiếu tủ (BSC, tủ hút, tủ cấy) ghi 'Giám định tủ ...' / 'Inspection of ...'."""
    txt = str(txt or "")
    return inspection_scope(txt, is_en) if is_cabinet else txt


def _last_line(txt: str) -> str:
    lines = [l.strip() for l in str(txt or "").split("\n") if l.strip()]
    return lines[-1] if lines else ""


def excel_context(data: Dict[str, Any]) -> Dict[str, Any]:
    """Dựng context từ dict dữ liệu mà DossierRunner truyền cho bộ dựng Excel (c_data / point_data)."""
    meta = data.get("metadata", {}) or {}
    cab = data.get("cabinet", {}) or {}
    code = data.get("code") or meta.get("code", "")
    sub = data.get("sub_no") or meta.get("sub_no", "")
    cert_date = data.get("cert_date") or data.get("sample_date") or ""
    cab_name = cab.get("cabinet_name") or data.get("cabinet_name", "")
    req = data.get("requirements", {}) or {}
    return {
        "ma_hs": code,
        "so_phu": sub if sub not in (None, "") else "",
        "so_ct": f"{code}/{sub}" if sub else code,
        "ngay_ct": _dt(cert_date),
        "ngay_dk": _dt(data.get("reg_date") or meta.get("reg_date") or data.get("sample_date")),
        "ngay_gd": _dt(data.get("sample_date") or meta.get("sample_date") or cert_date),
        "gio_bd": _tm(data.get("start_time") or meta.get("start_time")),
        "gio_kt": _tm(data.get("end_time") or meta.get("end_time")),
        "ten_tu": (f'"{cab_name}"' if cab_name and not cab_name.startswith('"') else cab_name),
        "ten_tu_en": cab.get("cabinet_name_en", "") or data.get("cabinet_name_en", "") or _quoted_in(data.get("scope_en")),
        "nhan_hieu": cab.get("brand", ""),
        "cap_tu": cab.get("class", "II"),
        "loai_tu": cab.get("cabinet_type") or cab.get("type", ""),
        "model": cab.get("model", ""),
        "serial": cab.get("serial", ""),
        "ma_tb": cab.get("device_id", "N/A"),
        "kh_ten": data.get("client_name_vn", ""),
        "kh_ten_en": data.get("client_name_en", ""),
        "kh_dia_chi": data.get("client_address_vn", ""),
        "kh_dia_chi_en": data.get("client_address_en", ""),
        "noi_dung_yc": _scope(data.get("scope_vn") or data.get("request_content_vn", ""), False, cab_name or cab),
        "noi_dung_yc_en": _scope(data.get("scope_en") or data.get("request_content_en", ""), True, cab_name or cab),
        "dia_diem": data.get("location_vn") or data.get("location_desc_vn", ""),
        "dia_diem_en": data.get("location_en") or data.get("location_desc_en", ""),
        "gdv_ten": data.get("gdv_name_vn", "") or _last_line(data.get("sign_gdv_vn", "")),
        "gdv_ten_en": data.get("gdv_name_en", "") or _last_line(data.get("sign_gdv_en", "")),
        # khí nén
        "vi_tri_ten": data.get("location_name", ""),
        "ky_hieu": data.get("location_symbol") or data.get("point_label", ""),
        "yc_bui": req.get("dust") or data.get("dust_requirement", ""),
        "yc_nuoc": req.get("water") or data.get("water_requirement", ""),
        "yc_dau": req.get("oil") or data.get("oil_requirement", ""),
    }


def fill_workbook_tags(wb, data: Dict[str, Any]) -> set:
    """Điền TAG trong workbook mở bằng openpyxl. Ô thiếu dữ liệu -> để trống (không giữ dữ liệu hồ sơ cũ)."""
    return render_xlsx(wb, excel_context(data), warn=False)
