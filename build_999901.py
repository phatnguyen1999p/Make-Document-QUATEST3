# -*- coding: utf-8 -*-
"""
build_999901.py
===============
Script dựng Hồ sơ ảo thử nghiệm Quan trắc Môi trường Lao động (QTMTLĐ):
- Mã hồ sơ ảo: Q326D999901 (lưu tại `Input/999901/`, KHÔNG ghi lên ổ mạng Z: theo AGENTS.md mục 0).
- Khách hàng ảo: CÔNG TY TNHH SẢN XUẤT LINH KIỆN CƠ ĐIỆN TỬ ALPHA VIỆT NAM
- Địa chỉ: Lô B12, đường số 5, khu công nghiệp Long Thành, xã Tam An, huyện Long Thành, tỉnh Đồng Nai
- Ngày quan trắc: 10/11/2026 | Ngày ban hành: 15/11/2026 (date.today())
- Kiểm thử toàn diện:
  1. Phiếu tính toán Vi khí hậu - Vật lý - Bụi (`Q326D999901 - TinhToanViKhiHauVatLyBui.xlsx`):
     - 03 vị trí chính (T1..T3) với cả mẫu Đạt và Không đạt (Nhiệt độ T2 = 34,6 °C Không đạt;
       Ánh sáng T3 = 255 Lux Không đạt; Tiếng ồn T1 = 87,2 dBA Không đạt).
  2. Phiếu tính toán Hơi khí độc (`Q326D999901 - TinhToanHoiKhiDoc.xlsx`):
     - Kiểm thử tra cứu VLOOKUP tự động từ `Sheet1` (678 dòng) đối với:
       + `Butyl acrylat` (chất mới bổ sung từ QCVN 01:2025/BYT, ngưỡng TWA = 55 mg/m³): 02 mẫu Đạt (`18,0` và `24,0` mg/m³).
       + `Toluen` (QCVN 01:2025/BYT, ngưỡng TWA = 100 mg/m³): 01 mẫu Đạt (`46,0` mg/m³) và 01 mẫu Không đạt (`110,0` mg/m³ tại T5).
       + `SO2`, `NOx`, `CO` (QCVN 03:2019/BYT): mỗi chỉ tiêu 03 mẫu Đạt.
       + Dòng kiểm thử bổ sung trong bảng tính cho `Aceton cyanohydrin` (TWA = 4), `Caprolactam (dạng hơi)` (TWA = 20),
         `Khói hàn` (QCVN 31:2026/BYT = 5), `n-Heptan` (QCVN 31:2026/BYT = 1200).
  3. Phiếu tính toán Ecgonomy & Tâm sinh lý (`Q326D999901 - TinhToanEcgonomy.xlsx`).
  4. Báo cáo kết quả QTMTLĐ chính thức (`TBKQ - viet - mtld - Q326D999901.docx`) + mục
     "KHUYẾN NGHỊ GIẢI PHÁP KHẮC PHỤC" do AI / Contextual Seeded Composer viết với giọng văn
     nhẹ nhàng, mang tính ghi nhận và sẽ khắc phục.
  5. Báo cáo kết quả nội bộ (`BCKQ - 1 - mtld - Q326D999901.docx`, biểu mẫu BM3030.03 HL:01/8/2025)
     đồng bộ 100% bảng Tổng hợp kết quả quan trắc (Tổng cộng: 31 mẫu | 27 mẫu đạt | 04 mẫu không đạt).
  6. Xuất toàn bộ PDF sang `Input/999901/pdf/` (an toàn máy in tem qua `avoid_label_printer`).
"""

from __future__ import annotations

import os
from datetime import date
from pathlib import Path
from typing import Any, Dict, List

import docx
import openpyxl
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from Core.Builders.handlers.mtld import (
    _set_cell_text_preserve_font,
    apply_mtld_recommendations_to_doc,
    ensure_mtld_table_headers_and_cantsplit,
    fix_mtld_header_and_page_fields,
    format_count_2digits,
    generate_mtld_recommendations,
    sync_mtld_summary_table,
)
from Core.Tools.mtld_calc import (
    ROOT_DIR,
    TEMPLATE_MTLD_DIR,
    create_ecgonomy_template,
    create_hoi_khi_doc_template,
    create_vikhihau_vatly_bui_template,
    create_word_templates,
    export_mtld_pdfs,
    populate_ecgonomy_xlsx,
    populate_hoi_khi_doc_xlsx,
    populate_vikhihau_vatly_bui_xlsx,
)

# Khóa cứng biến môi trường để tuyệt đối không đồng bộ lên Z:
os.environ["QUATEST3_NETWORK_ROOT"] = r"C:\__NON_EXISTENT_Z_SYNC__"

DOSSIER_ID = "999901"
DOSSIER_CODE = f"Q326D{DOSSIER_ID}"
COMPANY_NAME_UPPER = "CÔNG TY TNHH SẢN XUẤT LINH KIỆN CƠ ĐIỆN TỬ ALPHA VIỆT NAM"
COMPANY_NAME_TITLE = "Công ty TNHH Sản xuất Linh kiện Cơ điện tử Alpha Việt Nam"
COMPANY_ADDRESS = "Lô B12, đường số 5, khu công nghiệp Long Thành, xã Tam An, huyện Long Thành, tỉnh Đồng Nai"
INDUSTRY_DESC = "Gia công dập cơ khí, hàn sóng bản mạch điện tử, phủ keo bảo vệ (conformal coating) và lắp ráp linh kiện"
INSPECTION_DATE = date(2026, 11, 10)
ISSUE_DATE = date.today()

BASE_DIR = ROOT_DIR / "Input" / DOSSIER_ID
THUCHIEN_DIR = BASE_DIR / "ThucHien"
TBKQ_DIR = BASE_DIR / "TBKQ"
PDF_DIR = BASE_DIR / "pdf"


# Dữ liệu 3 vị trí đo Vi khí hậu - Ánh sáng - Tiếng ồn - Bụi hô hấp - Khí SO2/NOx/CO
LOCATIONS_MAIN = [
    {
        "code": "T1",
        "name": "T1 - Khu vực xưởng gia công cơ khí & máy dập liên hoàn",
        "temp": 28.4,       # Đạt (20 - 34)
        "hum": 54.2,        # Đạt (40 - 80)
        "wind": 0.24,       # Đạt (0,1 - 1,5)
        "lux_raw": 400,     # Sau hệ số * 1.05 -> 420 Lux (Đạt >= 300)
        "lux": 420,
        "noise": 87.2,      # Không đạt (> 85 dBA)
        "dust_resp": 0.042, # Đạt (<= 8,0 mg/m3)
        "co2_ppm": 520,
    },
    {
        "code": "T2",
        "name": "T2 - Khu vực lò hàn sóng & lắp ráp bản mạch điện tử (SMT)",
        "temp": 34.6,       # Không đạt (> 34 °C)
        "hum": 48.5,        # Đạt (40 - 80)
        "wind": 0.19,       # Đạt (0,1 - 1,5)
        "lux_raw": 490,     # Sau hệ số * 1.05 -> 515 Lux (Đạt >= 300)
        "lux": 515,
        "noise": 74.6,      # Đạt (<= 85 dBA)
        "dust_resp": 0.028, # Đạt (<= 8,0 mg/m3)
        "co2_ppm": 580,
    },
    {
        "code": "T3",
        "name": "T3 - Khu vực kiểm tra ngoại quan & đóng gói thành phẩm (KCS)",
        "temp": 26.2,       # Đạt (20 - 34)
        "hum": 56.8,        # Đạt (40 - 80)
        "wind": 0.15,       # Đạt (0,1 - 1,5)
        "lux_raw": 243,     # Sau hệ số * 1.05 -> 255 Lux (Không đạt < 300)
        "lux": 255,
        "noise": 68.3,      # Đạt (<= 85 dBA)
        "dust_resp": 0.019, # Đạt (<= 8,0 mg/m3)
        "co2_ppm": 490,
    },
]

# Dữ liệu 2 vị trí đo hơi dung môi hữu cơ (Butyl acrylat & Toluen theo QCVN 01:2025/BYT)
LOCATIONS_SOLVENT = [
    {
        "code": "T4",
        "name": "T4 - Khu vực phủ keo bảo vệ bản mạch (Conformal Coating)",
        "butyl_acrylat_ug": 540.0,  # V = 30 L -> KQTT = 18,0 mg/m3 (Đạt <= 55)
        "butyl_acrylat_str": "18,0",
        "butyl_acrylat_pass": True,
        "toluen_ug": 1380.0,        # V = 30 L -> KQTT = 46,0 mg/m3 (Đạt <= 100)
        "toluen_str": "46,0",
        "toluen_pass": True,
    },
    {
        "code": "T5",
        "name": "T5 - Khu vực tẩy rửa khuôn & vệ sinh bản mạch bằng dung môi",
        "butyl_acrylat_ug": 720.0,  # V = 30 L -> KQTT = 24,0 mg/m3 (Đạt <= 55)
        "butyl_acrylat_str": "24,0",
        "butyl_acrylat_pass": True,
        "toluen_ug": 3300.0,        # V = 30 L -> KQTT = 110,0 mg/m3 (Không đạt > 100)
        "toluen_str": "110",
        "toluen_pass": False,
    },
]

# Bảng tổng hợp kết quả quan trắc đồng bộ giữa TBKQ và BCKQ nội bộ (Tổng 31 mẫu: 27 Đạt, 04 Không đạt)
FACTOR_SUMMARY: List[Dict[str, Any]] = [
    {"name": "Nhiệt độ", "total": 3, "pass": 2, "fail": 1},
    {"name": "Độ ẩm", "total": 3, "pass": 3, "fail": 0},
    {"name": "Tốc độ gió", "total": 3, "pass": 3, "fail": 0},
    {"name": "Ánh sáng", "total": 3, "pass": 2, "fail": 1},
    {"name": "Tiếng ồn chung", "total": 3, "pass": 2, "fail": 1},
    {"name": "Bụi hô hấp", "total": 3, "pass": 3, "fail": 0},
    {"name": "Hexan", "display_name": "- Butyl acrylat", "total": 2, "pass": 2, "fail": 0},
    {"name": "Heptane", "display_name": "- Toluen", "total": 2, "pass": 1, "fail": 1},
    {"name": "SO2", "display_name": "- SO2", "total": 3, "pass": 3, "fail": 0},
    {"name": "NOx", "display_name": "- NOx", "total": 3, "pass": 3, "fail": 0},
    {"name": "CO", "display_name": "- CO", "total": 3, "pass": 3, "fail": 0},
]

FAILING_DETAILS: Dict[str, List[str]] = {
    "Nhiệt độ": ["T2 - Khu vực lò hàn sóng & lắp ráp bản mạch điện tử (34,6 °C / giới hạn 20 – 34 °C)"],
    "Ánh sáng": ["T3 - Khu vực kiểm tra ngoại quan & đóng gói thành phẩm (255 Lux / giới hạn ≥ 300 Lux)"],
    "Tiếng ồn chung": ["T1 - Khu vực xưởng gia công cơ khí & máy dập liên hoàn (87,2 dBA / giới hạn ≤ 85 dBA)"],
    "Hơi khí độc (Toluen)": ["T5 - Khu vực tẩy rửa khuôn & vệ sinh bản mạch bằng dung môi (110 mg/m³ / giới hạn ≤ 100 mg/m³)"],
}


def build_calc_vikhihau_vatly_bui() -> Path:
    """Sinh phiếu tính toán Q326D999901 - TinhToanViKhiHauVatLyBui.xlsx giữ nguyên công thức động."""
    out_path = THUCHIEN_DIR / f"{DOSSIER_CODE} - TinhToanViKhiHauVatLyBui.xlsx"
    return populate_vikhihau_vatly_bui_xlsx(
        out_path=out_path,
        locations_main=LOCATIONS_MAIN,
        company_name=COMPANY_NAME_TITLE,
        dossier_code=DOSSIER_CODE,
        inspection_date_str=INSPECTION_DATE.strftime("%d/%m/%Y"),
        file_suffix="Tính Toán Nhiệt Ẩm",
    )


def build_calc_hoi_khi_doc() -> Path:
    """Sinh phiếu tính toán Q326D999901 - TinhToanHoiKhiDoc.xlsx kiểm thử tra cứu Sheet1 (678 chất)."""
    out_path = THUCHIEN_DIR / f"{DOSSIER_CODE} - TinhToanHoiKhiDoc.xlsx"
    samples = [
        # 1. Butyl acrylat (chất mới bổ sung từ QCVN 01:2025/BYT, ngưỡng TWA = 55 mg/m3)
        {
            "location": LOCATIONS_SOLVENT[0]["name"],
            "flow_lpm": 1.0,
            "time_min": 30,
            "kqtn_no": "03120/2026/KQTN",
            "parameter": "Butyl acrylat",
            "kqtn_mg_l": "-",
            "vol_sol_ml": "-",
            "kqtn_ug": LOCATIONS_SOLVENT[0]["butyl_acrylat_ug"],
        },
        {
            "location": LOCATIONS_SOLVENT[1]["name"],
            "flow_lpm": 1.0,
            "time_min": 30,
            "kqtn_no": "03120/2026/KQTN",
            "parameter": "Butyl acrylat",
            "kqtn_mg_l": "-",
            "vol_sol_ml": "-",
            "kqtn_ug": LOCATIONS_SOLVENT[1]["butyl_acrylat_ug"],
        },
        # 2. Toluen (QCVN 01:2025/BYT, ngưỡng TWA = 100 mg/m3 — T4 Đạt, T5 Không đạt)
        {
            "location": LOCATIONS_SOLVENT[0]["name"],
            "flow_lpm": 1.0,
            "time_min": 30,
            "kqtn_no": "03120/2026/KQTN",
            "parameter": "Toluen",
            "kqtn_mg_l": "-",
            "vol_sol_ml": "-",
            "kqtn_ug": LOCATIONS_SOLVENT[0]["toluen_ug"],
        },
        {
            "location": LOCATIONS_SOLVENT[1]["name"],
            "flow_lpm": 1.0,
            "time_min": 30,
            "kqtn_no": "03120/2026/KQTN",
            "parameter": "Toluen",
            "kqtn_mg_l": "-",
            "vol_sol_ml": "-",
            "kqtn_ug": LOCATIONS_SOLVENT[1]["toluen_ug"],
        },
        # 3. SO2, NOx, CO tại 3 vị trí T1, T2, T3 (đo bằng ống phát hiện nhanh Kitagawa)
        *[
            {
                "location": loc["name"],
                "flow_lpm": "-",
                "time_min": "-",
                "kqtn_no": "Ống chỉ thị Kitagawa",
                "parameter": gas_name,
                "kqtn_mg_l": "-",
                "vol_sol_ml": "-",
                "direct_mg_m3": gas_val,
            }
            for loc in LOCATIONS_MAIN
            for gas_name, gas_val in (("SO2", "< 1,3"), ("NOx", "< 0,94"), ("CO", "< 1,1"))
        ],
        # 4. Các dòng kiểm thử bổ sung cho những chất đặc biệt trong QCVN 01:2025/BYT & QCVN 31:2026/BYT
        {
            "location": "T2 - Khu vực lò hàn sóng (Kiểm thử QCVN 31:2026/BYT)",
            "flow_lpm": 2.0,
            "time_min": 30,
            "kqtn_no": "03121/2026/KQTN",
            "parameter": "Khói hàn",
            "kqtn_mg_l": 0.012,
            "vol_sol_ml": 10,
        },
        {
            "location": "T4 - Khu vực phủ keo (Kiểm thử Aceton cyanohydrin QCVN 01:2025/BYT)",
            "flow_lpm": 1.0,
            "time_min": 30,
            "kqtn_no": "03121/2026/KQTN",
            "parameter": "Aceton cyanohydrin",
            "kqtn_mg_l": "< 0,03",
            "vol_sol_ml": 10,
        },
        {
            "location": "T5 - Khu vực tẩy rửa (Kiểm thử Caprolactam dạng hơi QCVN 01:2025/BYT)",
            "flow_lpm": 1.0,
            "time_min": 30,
            "kqtn_no": "03121/2026/KQTN",
            "parameter": "Caprolactam (dạng hơi)",
            "kqtn_mg_l": "-",
            "vol_sol_ml": "-",
            "kqtn_ug": 180.0,
        },
    ]
    return populate_hoi_khi_doc_xlsx(
        out_path,
        samples,
        dossier_code=DOSSIER_CODE,
        date_str=ISSUE_DATE.strftime("%d/%m/%Y"),
        file_suffix="Tính Toán Hơi Khí Độc",
    )


def build_calc_ecgonomy() -> Path:
    """Sinh phiếu tính toán Q326D999901 - TinhToanEcgonomy.xlsx."""
    out_path = THUCHIEN_DIR / f"{DOSSIER_CODE} - TinhToanEcgonomy.xlsx"
    return populate_ecgonomy_xlsx(
        out_path=out_path,
        locations=LOCATIONS_MAIN + LOCATIONS_SOLVENT,
        dossier_code=DOSSIER_CODE,
        date_str=INSPECTION_DATE.strftime("%d/%m/%Y"),
        file_suffix="Tính Toán Ecgonomy",
    )


def build_tbkq_and_bckq() -> Tuple[Path, Path, Dict[str, str]]:
    """Dựng báo cáo TBKQ và BCKQ nội bộ hoàn chỉnh cho hồ sơ ảo Q326D999901."""
    tpl_tbkq = TEMPLATE_MTLD_DIR / "FORM - TBKQ - viet - mtld - 2026.docx"
    tpl_bckq = TEMPLATE_MTLD_DIR / "FORM - BCKQ - mtld - 2026.docx"
    if not tpl_tbkq.exists() or not tpl_bckq.exists():
        create_word_templates()

    out_tbkq = TBKQ_DIR / f"TBKQ - viet - mtld - {DOSSIER_CODE}.docx"
    out_bckq = THUCHIEN_DIR / f"BCKQ - 1 - mtld - {DOSSIER_CODE}.docx"

    issue_str = ISSUE_DATE.strftime("%d/%m/%Y")
    insp_str = INSPECTION_DATE.strftime("%d/%m/%Y")

    # 1. Gọi AI / Contextual Seeded Composer viết mục Khuyến nghị giải pháp khắc phục
    recs = generate_mtld_recommendations(
        company_name=COMPANY_NAME_TITLE,
        industry_desc=INDUSTRY_DESC,
        factor_summary=FACTOR_SUMMARY,
        failing_details=FAILING_DETAILS,
        dossier_code=DOSSIER_CODE,
        use_ai=True,
    )

    # 2. Dựng TBKQ
    doc = docx.Document(tpl_tbkq)
    fix_mtld_header_and_page_fields(doc, DOSSIER_CODE, issue_str, is_bckq=False)

    # Table 0: Trang bìa
    t0 = doc.tables[0]
    _set_cell_text_preserve_font(t0.rows[6].cells[2], f"Ngày {ISSUE_DATE.day:02d}", font_size_pt=13.0, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t0.rows[6].cells[3], f"Tháng {ISSUE_DATE.month:02d}", font_size_pt=13.0, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t0.rows[6].cells[4], f"Năm {ISSUE_DATE.year}", font_size_pt=13.0, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t0.rows[10].cells[0], COMPANY_NAME_UPPER, font_size_pt=16.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t0.rows[11].cells[1], COMPANY_ADDRESS, font_size_pt=13.0, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)

    # Table 1: Số & Ngày trang trong
    t1 = doc.tables[1]
    _set_cell_text_preserve_font(t1.rows[1].cells[0], f"Số: {DOSSIER_CODE}/MTLĐ", font_size_pt=13.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(
        t1.rows[1].cells[1],
        f"TPHCM, ngày {ISSUE_DATE.day:02d} tháng {ISSUE_DATE.month:02d} năm {ISSUE_DATE.year}",
        font_size_pt=13.0,
        italic=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    # Table 2: Thông tin chung & phạm vi quan trắc
    t2 = doc.tables[2]
    _set_cell_text_preserve_font(
        t2.rows[4].cells[0],
        f"Do Ông: Nguyễn Xuân Nam làm đại diện đã tiến hành quan trắc môi trường lao động tại:\n"
        f"{COMPANY_NAME_TITLE}\n"
        f"Địa chỉ: {COMPANY_ADDRESS}",
        font_size_pt=13.0,
    )
    _set_cell_text_preserve_font(
        t2.rows[5].cells[0],
        f"Ngày {INSPECTION_DATE.day:02d} tháng {INSPECTION_DATE.month:02d} năm {INSPECTION_DATE.year}",
        font_size_pt=13.0,
    )
    _set_cell_text_preserve_font(
        t2.rows[24].cells[-1],
        "Butyl acrylat, Toluen, CO, SO2, NOx phát sinh từ hoạt động gia công, hàn mạch và sử dụng dung môi",
        font_size_pt=13.0,
    )

    # Cập nhật đoạn văn tiêu đề mục 1 của Hơi khí độc
    for p in doc.paragraphs:
        if "Hexan, Heptane" in p.text:
            for r in p.runs:
                r.text = ""
            run = p.runs[0] if p.runs else p.add_run()
            run.text = "1.\tNồng độ hơi Butyl acrylat, Toluen"
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
            run.font.bold = True

    # Table 3: Vi khí hậu (Nhiệt độ, Độ ẩm, Tốc độ gió)
    t3 = doc.tables[3]
    for idx, loc in enumerate(LOCATIONS_MAIN):
        row = t3.rows[3 + idx]
        _set_cell_text_preserve_font(row.cells[0], str(idx + 1), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[1], loc["name"], align=WD_ALIGN_PARAGRAPH.LEFT)
        t_str = f"{loc['temp']:.1f}".replace(".", ",")
        h_str = f"{loc['hum']:.1f}".replace(".", ",")
        w_str = f"{loc['wind']:.2f}".replace(".", ",")
        # Nhiệt độ (20 - 34)
        if 20.0 <= loc["temp"] <= 34.0:
            _set_cell_text_preserve_font(row.cells[2], t_str, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(row.cells[3], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            _set_cell_text_preserve_font(row.cells[2], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(row.cells[3], t_str, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        # Độ ẩm (40 - 80)
        _set_cell_text_preserve_font(row.cells[4], h_str, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[5], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        # Tốc độ gió (0,1 - 1,5)
        _set_cell_text_preserve_font(row.cells[6], w_str, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[7], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

    _set_cell_text_preserve_font(t3.rows[6].cells[2], "02", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t3.rows[6].cells[3], "01", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t3.rows[6].cells[4], "03", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t3.rows[6].cells[5], "00", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t3.rows[6].cells[6], "03", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t3.rows[6].cells[7], "00", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t3.rows[10].cells[0], "Tổng số mẫu nhiệt độ vượt giới hạn cho phép: 01 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)

    # Table 4: Ánh sáng (Lux, >= 300)
    t4 = doc.tables[4]
    for idx, loc in enumerate(LOCATIONS_MAIN):
        row = t4.rows[3 + idx]
        _set_cell_text_preserve_font(row.cells[0], str(idx + 1), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[1], loc["name"], align=WD_ALIGN_PARAGRAPH.LEFT)
        if loc["lux"] >= 300:
            _set_cell_text_preserve_font(row.cells[2], str(loc["lux"]), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(row.cells[3], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            _set_cell_text_preserve_font(row.cells[2], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(row.cells[3], str(loc["lux"]), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t4.rows[6].cells[2], "02", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t4.rows[6].cells[3], "01", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t4.rows[9].cells[0], "Tổng số mẫu ánh sáng vượt giới hạn cho phép: 01 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)

    # Table 5: Tiếng ồn (dBA, <= 85)
    t5 = doc.tables[5]
    for idx, loc in enumerate(LOCATIONS_MAIN):
        row = t5.rows[3 + idx]
        _set_cell_text_preserve_font(row.cells[0], str(idx + 1), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[1], loc["name"], align=WD_ALIGN_PARAGRAPH.LEFT)
        n_str = f"{loc['noise']:.1f}".replace(".", ",")
        if loc["noise"] <= 85.0:
            _set_cell_text_preserve_font(row.cells[2], n_str, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(row.cells[3], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            _set_cell_text_preserve_font(row.cells[2], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(row.cells[3], n_str, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t5.rows[6].cells[2], "02", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t5.rows[6].cells[3], "01", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t5.rows[10].cells[0], "Tổng số mẫu độ ồn vượt giới hạn cho phép: 01 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)

    # Table 6: Bụi hô hấp (<= 8,0 mg/m3)
    t6 = doc.tables[6]
    for idx, loc in enumerate(LOCATIONS_MAIN):
        row = t6.rows[3 + idx]
        _set_cell_text_preserve_font(row.cells[0], str(idx + 1), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[1], loc["name"], align=WD_ALIGN_PARAGRAPH.LEFT)
        d_str = f"{loc['dust_resp']:.3f}".replace(".", ",")
        _set_cell_text_preserve_font(row.cells[2], d_str, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[3], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t6.rows[9].cells[0], "Tổng số mẫu bụi hô hấp: 03 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell_text_preserve_font(t6.rows[10].cells[0], "Tổng số mẫu bụi hô hấp vượt giới hạn cho phép: 00 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)

    # Table 7: Hơi dung môi hữu cơ (Butyl acrylat <= 55 mg/m3 và Toluen <= 100 mg/m3 theo QCVN 01:2025/BYT)
    t7 = doc.tables[7]
    _set_cell_text_preserve_font(t7.rows[0].cells[0], "Giới hạn cho phép theo QCVN 01:2025/BYT (a)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t7.rows[0].cells[3], "Butyl acrylat (mg/m³)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t7.rows[0].cells[5], "Toluen (mg/m³)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t7.rows[1].cells[3], "≤ 55", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t7.rows[1].cells[5], "≤ 100", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    for idx, sloc in enumerate(LOCATIONS_SOLVENT):
        row = t7.rows[3 + idx]
        _set_cell_text_preserve_font(row.cells[0], str(idx + 1), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[2], sloc["name"], align=WD_ALIGN_PARAGRAPH.LEFT)
        # Butyl acrylat
        _set_cell_text_preserve_font(row.cells[3], sloc["butyl_acrylat_str"], bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[4], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        # Toluen
        if sloc["toluen_pass"]:
            _set_cell_text_preserve_font(row.cells[5], sloc["toluen_str"], bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(row.cells[6], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            _set_cell_text_preserve_font(row.cells[5], "-", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text_preserve_font(row.cells[6], sloc["toluen_str"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    _set_cell_text_preserve_font(t7.rows[5].cells[3], "02", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t7.rows[5].cells[4], "00", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t7.rows[5].cells[5], "01", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t7.rows[5].cells[6], "01", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text_preserve_font(t7.rows[7].cells[0], "Tổng số mẫu Butyl acrylat: 02 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell_text_preserve_font(t7.rows[8].cells[0], "Tổng số mẫu Butyl acrylat vượt giới hạn cho phép: 00 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell_text_preserve_font(t7.rows[9].cells[0], "Tổng số mẫu Toluen: 02 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell_text_preserve_font(t7.rows[10].cells[0], "Tổng số mẫu Toluen vượt giới hạn cho phép: 01 mẫu", align=WD_ALIGN_PARAGRAPH.LEFT)

    # Table 8: SO2, NOx, CO
    t8 = doc.tables[8]
    for idx, loc in enumerate(LOCATIONS_MAIN):
        row = t8.rows[3 + idx]
        _set_cell_text_preserve_font(row.cells[0], str(idx + 1), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text_preserve_font(row.cells[2], loc["name"], align=WD_ALIGN_PARAGRAPH.LEFT)

    # Table 9: Bảng Tổng hợp kết quả quan trắc
    sync_mtld_summary_table(doc, FACTOR_SUMMARY)

    # Table 10: Bảng Khuyến nghị giải pháp khắc phục (do AI / Contextual Seeded Composer sinh)
    apply_mtld_recommendations_to_doc(doc, recs)

    # Đặt lặp lại dòng tiêu đề bảng (<w:tblHeader/>), chống tách dòng (<w:cantSplit/>) và chống gãy dòng (<w:noWrap/>)
    ensure_mtld_table_headers_and_cantsplit(doc)

    doc.save(out_tbkq)

    # 3. Dựng BCKQ nội bộ (BM3030.03 HL:01/8/2025)
    import re
    doc_bckq = docx.Document(tpl_bckq)
    fix_mtld_header_and_page_fields(doc_bckq, DOSSIER_CODE, issue_str, is_bckq=True)
    tb0 = doc_bckq.tables[0]
    _set_cell_text_preserve_font(tb0.rows[0].cells[8], DOSSIER_CODE, font_size_pt=13.0, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell_text_preserve_font(tb0.rows[0].cells[-1], "", font_size_pt=13.0, bold=False)
    _set_cell_text_preserve_font(tb0.rows[2].cells[-1], COMPANY_NAME_UPPER, font_size_pt=13.0, bold=False)
    _set_cell_text_preserve_font(tb0.rows[3].cells[-1], COMPANY_ADDRESS, font_size_pt=13.0, bold=False)
    _set_cell_text_preserve_font(tb0.rows[6].cells[-1], f"Nhà máy của {COMPANY_NAME_TITLE}, {COMPANY_ADDRESS}", font_size_pt=13.0, bold=False)
    _set_cell_text_preserve_font(tb0.rows[7].cells[-1], f"Bắt đầu từ 08:30 đến 15:30 ngày {insp_str}", font_size_pt=13.0, bold=False)
    _set_cell_text_preserve_font(
        tb0.rows[28].cells[-1],
        "Butyl acrylat, Toluen, CO, SO2, NOx phát sinh từ hoạt động gia công, hàn mạch và sử dụng dung môi",
        font_size_pt=13.0,
        bold=False,
    )

    # Cập nhật toàn bộ mã hồ sơ và ngày trong các đoạn văn của BCKQ (mục 4.5, Gửi mẫu thử nghiệm, Ngày báo cáo)
    for p in doc_bckq.paragraphs:
        if any(k in p.text for k in ("Q326", "014464", "04/2026")):
            new_txt = p.text
            new_txt = re.sub(r"Q3\d{2}[A-D]\d{6}", DOSSIER_CODE, new_txt)
            new_txt = new_txt.replace("014464/VYTCC ngày 17/04/2026 và 014465/VYTCC ngày 17/04/2026", "03120/2026/KQTN ngày 14/11/2026 và 03121/2026/KQTN ngày 14/11/2026")
            new_txt = new_txt.replace("10/04/2026", insp_str)
            new_txt = new_txt.replace("08/04/2026", "10/11/2026")
            new_txt = new_txt.replace("20/04/2026", "14/11/2026")
            new_txt = new_txt.replace("21/04/2026", issue_str)
            new_txt = new_txt.replace("22/04/2026", issue_str)
            for r in p.runs:
                r.text = ""
            run = p.runs[0] if p.runs else p.add_run()
            run.text = new_txt
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)

    # Bảng 1 của BCKQ: ghi nhận kết quả đo ống phát hiện nhanh tại T1, T2, T3
    tb1 = doc_bckq.tables[1]
    for idx, loc_code in enumerate(("T1", "T2", "T3")):
        r_idx = 2 + idx * 2
        if r_idx < len(tb1.rows):
            _set_cell_text_preserve_font(tb1.rows[r_idx].cells[0], loc_code, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Bảng 2 của BCKQ: Bảng Tổng hợp kết quả quan trắc đồng bộ 100% với TBKQ
    sync_mtld_summary_table(doc_bckq, FACTOR_SUMMARY)
    ensure_mtld_table_headers_and_cantsplit(doc_bckq)

    doc_bckq.save(out_bckq)
    return out_tbkq, out_bckq, recs


def export_all_pdfs(files: List[Path]) -> List[Path]:
    """Xuất PDF cho các file DOCX/XLSX qua hàm chuẩn export_mtld_pdfs trong Core."""
    return export_mtld_pdfs(files, PDF_DIR)


if __name__ == "__main__":
    THUCHIEN_DIR.mkdir(parents=True, exist_ok=True)
    TBKQ_DIR.mkdir(parents=True, exist_ok=True)
    PDF_DIR.mkdir(parents=True, exist_ok=True)

    p_vkh = build_calc_vikhihau_vatly_bui()
    p_hkd = build_calc_hoi_khi_doc()
    p_eco = build_calc_ecgonomy()
    p_tbkq, p_bckq, recs = build_tbkq_and_bckq()

    print(f"[OK] Đã tạo phiếu tính toán VKH-VL-Bụi: {p_vkh}")
    print(f"[OK] Đã tạo phiếu tính toán Hơi khí độc: {p_hkd}")
    print(f"[OK] Đã tạo phiếu tính toán Ecgonomy:   {p_eco}")
    print(f"[OK] Đã tạo báo cáo TBKQ:               {p_tbkq}")
    print(f"[OK] Đã tạo báo cáo nội bộ BCKQ:        {p_bckq}")

    pdfs = export_all_pdfs([p_tbkq, p_bckq, p_vkh, p_hkd, p_eco])
    print(f"[OK] Đã xuất {len(pdfs)} file PDF vào {PDF_DIR}: {[p.name for p in pdfs]}")
    print("\n=== NỘI DUNG KHUYẾN NGHỊ GIẢI PHÁP KHẮC PHỤC ĐÃ SINH CHO Q326D999901 ===")
    for k, v in recs.items():
        print(f"\n[{k}]\n{v}")
