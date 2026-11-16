# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D138143 - Công ty TNHH Cường Thịnh Medical (phòng sạch, 29/10/2026) — 1 chứng thư (mã trần).

Địa điểm: Bệnh viện Hoàn Mỹ Sài Gòn (BBLM). Khu vực: "Phòng hậu phẫu" - tầng 4 (26 m², ISO 8, trạng thái nghỉ).
Chỉ tiêu (BBLM): nồng độ hạt bụi 05 vị trí; nhiệt độ - độ ẩm 01; chênh áp 01; lưu lượng + ACH 03 (HEPA box); độ ồn 01.
Mức quy định (đề cương khách hàng "Đề cương kiểm định hệ khí sạch - Bệnh viện Hoàn Mỹ Sài Gòn – Giai đoạn 3",
scan tr. 13–18, bảng phòng hậu phẫu): 20 - 24 °C; 40 - 60 %RH; chênh áp 3 - 7 Pa; ồn ≤ 50 dBA; ACH ≥ 20; ISO 8.

Nguồn: Input/138143/20261104124713226.pdf (26 trang: PKSCV 1, PYC 3–6, báo giá 7–9, BBLM 11–12,
       đề cương 13–18, mặt bằng có đánh dấu vị trí 19, ghi chép 21–25, giấy tính ACH 26; trang 2, 10, 20 trắng),
       Input/138143/6ChannelFixed.xlsx (AeroTrak 9350, hạt/m³ cộng dồn Σ; file 1148 zero, 1149–1153 đo),
       Input/138143/Data Do Luu Luong/*.csv (testo 420, ghi chú 1-1 … 3-3 = HEPA 1..3, từng phần miệng cấp),
       Input/138143/Auto_0178 (Rion NL-53, LAeq 1 phút = 50,3 dBA),
       Input/138143/Photos-1-00/Use (5 ảnh chèn chứng thư), Old test result/88904 (TBKQ cũ cùng khách — tham khảo).

Người dùng chốt (04/11/2026):
  - Thể tích phòng = 26 m² x 2,5 m = 65 m³ (giấy tính tay tr. 26); ACH = 1 381 / 65 = 21,2 (giấy ghi "ACH: 20").
  - Vị trí bụi 1..5 = file 1149..1153 theo thứ tự giờ đo (phiếu ghi chép ghi "1153, 1152, 1151, 1150, 1149").
  - Nhiệt độ/độ ẩm hiệu chính theo giấy hiệu chuẩn Testo 625 – máy 3: giá trị đúng =
    chỉ thị − sai số, sai số nội suy tuyến tính giữa 2 điểm hiệu chuẩn kẹp hai bên; giá trị = TB 3 lần sau hiệu chính.
  - Ngày ban hành = 04/11/2026 (ngày làm hồ sơ, không có KQTN).
  - Chênh áp "Phòng hậu phẫu → Hành lang" (người dùng chốt), = TB 3 lần đọc.
  - Lệch nguồn giữ theo BBLM: ngày 29/10/2026 (phiếu ghi chép in sẵn 21/07/2026); thiết bị chênh áp TSI 8375.
  - Độ ồn KHÔNG đánh giá: khách hàng xác nhận lúc đo xung quanh có hoạt động tạo nguồn ồn khác -> TBKQ chỉ trả số,
    mức "-", ghi chú (4) NOISE_NOTE; BCKQ ghi "Chỉ tiêu không đánh giá"; Excel mức/đánh giá "-".

Chạy:  python build_138143.py          (DOCX + XLSX + Edoc + sơ đồ)
       python build_138143.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: trong code (network_dir để trống) — chép lên Z bằng robocopy (AGENTS mục 0).
"""

import os
import re
import sys
import glob
import copy
import argparse

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)

import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, within
from Core.Styling.paragraphs import set_assessment_paragraph
from Core.Tools.excel_borders import thin_outside_borders
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)
from Core.Tools.cleanroom_calc import (CleanroomCalcBook, CalTable, read_aerotrak, apply_cal,
                                       replace_reference_lines)

CODE = "Q326D138143"
CODE6 = "138143"
REG_DATE = "29/10/2026"           # PYC + BBLM mục 1
SAMPLE_DATE = "29/10/2026"        # BBLM (phiếu ghi chép in sẵn "21/07/2026" — lệch, báo người dùng)
import datetime as _dt
CERT_DATE = _dt.date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày hiện tại, tự cập nhật mỗi lần dựng (người dùng chốt 06/11/2026, AGENTS mục 2)
START, END = "14:00", "15:30"     # BBLM mục 6
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN = os.path.join(INPUT_DIR, "20261104124713226.pdf")
DUST_XLSX = os.path.join(INPUT_DIR, "6ChannelFixed.xlsx")
FLOW_DIR = os.path.join(INPUT_DIR, "Data Do Luu Luong")
NOISE_RND = os.path.join(INPUT_DIR, "Auto_0178", "Auto_Leq", "NL_0001_OCT_Leq_0178_0001.rnd")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photos-1-00", "Use")
DIAGRAM_DIR = os.path.join(INPUT_DIR, "TBKQ", "_diagrams")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Cường Thịnh Medical",
    name_en="Cuong Thinh Medical Company Limited",
    # địa chỉ theo BBLM + PYC bổ sung (không "ấp 3" như PYC cũ trong scan — người dùng chốt 05/11/2026)
    address_vn="Số 22/2G, đường Xuân Thới Sơn 20C, xã Xuân Thới Sơn, Thành phố Hồ Chí Minh",
    address_en="No. 22/2G, Xuan Thoi Son 20C street, Xuan Thoi Son commune, Ho Chi Minh City",
)
# PYC bổ sung ký ngày 28/10/2026 (Z: "PHIẾU YÊU CẦU & NHẬN THẨM ĐỊNH-GIÁM ĐỊNH KHÍ SẠCH GIAI ĐOẠN 3.pdf") ghi tên
# liên danh -> người dùng chốt 05/11/2026: CHỈ TBKQ ghi tên liên danh; BCKQ, Excel giữ tên theo BBLM + ghi chú.
PYC_NEW_DATE = "28/10/2026"
CLIENT_TBKQ = ClientInfo(
    name_vn="Công ty TNHH Cường Thịnh Medical (liên danh Quang Tùng – Cường Thịnh Medical)",
    name_en="Cuong Thinh Medical Company Limited (Quang Tung – Cuong Thinh Medical Joint Venture)",
    address_vn=CLIENT.address_vn, address_en=CLIENT.address_en,
)
LOCATION_VN = "Bệnh viện Hoàn Mỹ Sài Gòn, 60-60A Phan Xích Long, phường Cầu Kiệu, Thành phố Hồ Chí Minh"
LOCATION_EN = "Hoan My Sai Gon Hospital, 60-60A Phan Xich Long street, Cau Kieu ward, Ho Chi Minh City"
OUTLINE_VN = "“Đề cương kiểm định hệ khí sạch - Đối với: Bệnh viện Hoàn Mỹ Sài Gòn – Giai đoạn 3”"
OUTLINE_EN = "“Clean air system inspection outline – For: Hoan My Sai Gon Hospital – Phase 3”"

ROOM_VN, ROOM_EN = "Phòng hậu phẫu", "Post-Anesthesia Care Unit (PACU)"
AREA_M2, HEIGHT_M = 26.0, 2.5      # diện tích: báo giá + đề cương; chiều cao: giấy tính tay tr. 26
ISO8_LIMITS = [3520000, 832000, 29300]   # ISO 14644-1:2015 Bảng 1 (≥ 0,5; ≥ 1; ≥ 5 µm)
ZERO_FILE = 1148
POINT_FILES = [1149, 1150, 1151, 1152, 1153]     # vị trí 1..5 theo giờ đo (người dùng chốt 04/11/2026)
TEMP_READ = [21.9, 21.8, 21.6]                   # phiếu ghi chép tr. 1
RH_READ = [61.9, 62.3, 62.0]
# Giấy hiệu chuẩn Testo 625 – máy 3, mục 14 (người dùng gửi 04/11/2026): (giá trị chỉ thị, giá trị chuẩn, sai số)
CAL_T = [(15.2, 15.00, 0.2), (20.0, 20.00, 0.0), (24.9, 25.00, -0.1), (29.8, 30.00, -0.2), (39.7, 40.00, -0.3)]
CAL_RH = [(36.1, 30.0, 6.1), (45.8, 40.0, 5.8), (54.5, 50.0, 4.5), (64.8, 60.0, 4.8), (74.3, 70.0, 4.3),
          (84.0, 80.0, 4.0)]
CAL_DEVICE = "Testo 625 – máy 3"
# Giấy hiệu chuẩn lấy từ thư viện HieuChuan/ (AGENTS 8C): Testo 625 máy 3 = MT33-3; chênh áp TSI 8375 = MT31.
# Lưu lượng Testo 420 và độ ồn Rion NL-53 chưa có bảng hiệu chuẩn trong thư viện -> không hiệu chính.
CAL_T_TBL = CalTable.from_library("MT33-3", "nhiệt độ", SAMPLE_DATE, device=CAL_DEVICE, unit="°C", fallback=CAL_T)
CAL_RH_TBL = CalTable.from_library("MT33-3", "độ ẩm", SAMPLE_DATE, device=CAL_DEVICE, unit="%RH", fallback=CAL_RH)
CAL_DP_TBL = CalTable.from_library("MT31", "Chênh áp / áp suất", SAMPLE_DATE, device="TSI 8375", unit="Pa")
DP_READ = [3.416, 3.352, 3.330]                  # Pa
LIM = {"t": "20 - 24", "rh": "40 - 60", "dp": "3 - 7", "noise": "≤ 50", "ach": 20}
TRANG_SO_DO = "2"                                # trang phụ lục chứa sơ đồ — kiểm lại trên PDF


# ---------------------------------------------------------------------------
# Đọc số liệu
# ---------------------------------------------------------------------------
def read_particles():
    """6ChannelFixed.xlsx -> {file: {"time", "vol", "c": [0,3; 0,5; 1; 2; 3; 5 µm]}} (hạt/m³, cộng dồn)."""
    ws = openpyxl.load_workbook(DUST_XLSX, data_only=True).active
    hdr, out = None, {}
    for row in ws.iter_rows(values_only=True):
        vals = list(row)
        if hdr is None:
            if "#" in vals and "Location" in vals:
                hdr = {str(v).split("\n")[0].strip(): i for i, v in enumerate(vals) if v not in (None, "")}
            continue
        n = vals[hdr["#"]]
        if isinstance(n, (int, float)):
            out[int(n)] = {"time": vals[hdr["Date/Time (dd/mm/yyyy)"]], "vol": vals[hdr["Vol."]],
                           "c": [vals[hdr[f"Ch{i}"]] for i in range(1, 7)]}
    missing = [f for f in [ZERO_FILE] + POINT_FILES if f not in out]
    if missing:
        raise SystemExit(f"[FATAL] 6ChannelFixed.xlsx thiếu file {missing}")
    return out


def read_flows():
    """CSV testo 420 -> [(hepa, phần, thời điểm, [5 số đọc m³/h], trung bình máy)] theo ghi chú '1-1' … '3-3'."""
    out = []
    for f in sorted(glob.glob(os.path.join(FLOW_DIR, "*.csv"))):
        t = open(f, "rb").read().decode("utf-8", "ignore")
        t = t[t.find("Date/Time,380"):]
        lines = t.splitlines()
        reads = [int(l.split(",")[4]) for l in lines if re.match(r"\d+/\d+/\d+ ", l)]
        avg = int([l for l in lines if l.startswith("Overall Average")][0].split(",")[4])
        m = re.search(r"Basic View,([^,]+),(\d)-(\d),", t)
        out.append((int(m.group(2)), int(m.group(3)), m.group(1), reads, avg))
    if len(out) != 7:
        raise SystemExit(f"[FATAL] cần 7 file lưu lượng, có {len(out)}")
    return out


def read_noise():
    for line in open(NOISE_RND, encoding="utf-8", errors="ignore"):
        if line.startswith("Leq"):
            return float(line.split(",")[1])
    raise SystemExit("[FATAL] không đọc được LAeq")


PARTICLES = read_particles()
FLOWS = read_flows()
NOISE = read_noise()                                       # 50,3 dBA (khớp phiếu ghi chép)


def r1(x):
    """Làm tròn 1 số lẻ kiểu nửa lên (như ROUND của Excel)."""
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(str(x)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))


def vn(x, nd=1):
    return f"{x:.{nd}f}".replace(".", ",")


def cal_bracket(x, table):
    """Chỉ số i sao cho table[i], table[i+1] kẹp giá trị chỉ thị x (ngoài khoảng -> đoạn đầu/cuối)."""
    for i in range(len(table) - 1):
        if table[i][0] <= x <= table[i + 1][0]:
            return i
    return 0 if x < table[0][0] else len(table) - 2


def calibrated(x, table):
    """Giá trị sau hiệu chính = chỉ thị − sai số (nội suy tuyến tính, như FORECAST của Excel)."""
    i = cal_bracket(x, table)
    (x1, _, e1), (x2, _, e2) = table[i], table[i + 1]
    return x - (e1 + (e2 - e1) * (x - x1) / (x2 - x1))


T_AVG = r1(sum(apply_cal(v, CAL_T_TBL) for v in TEMP_READ) / 3)
RH_AVG = r1(sum(apply_cal(v, CAL_RH_TBL) for v in RH_READ) / 3)
DP_AVG = r1(sum(apply_cal(v, CAL_DP_TBL) for v in DP_READ) / 3)
HEPA = {}
for h, part, _, reads, _avg in FLOWS:
    HEPA[h] = HEPA.get(h, 0) + round(sum(reads) / len(reads))     # = ROUND(AVERAGE(...),0) từng phần
TOTAL = sum(HEPA.values())
VOLUME = AREA_M2 * HEIGHT_M
ACH = r1(TOTAL / VOLUME)

# Độ ồn: khách hàng xác nhận lúc đo xung quanh có hoạt động tạo nguồn ồn khác -> chỉ trả số liệu,
# mức "-", không đánh giá, thêm ghi chú (người dùng chốt 04/11/2026). Mức đề cương ≤ 50 dBA chỉ để tham khảo.
NOISE_NOTE_VN = "Tại thời điểm thực hiện, khu vực xung quanh có các hoạt động tạo nguồn ồn khác."
NOISE_NOTE_EN = "At the time of measurement, there were other noise-generating activities around the area."  # 1 dòng

OK = {
    "t": within(vn(T_AVG), LIM["t"]),
    "rh": within(vn(RH_AVG), LIM["rh"]),
    "noise": None,     # không đánh giá (khách hàng xác nhận có nguồn ồn khác xung quanh — NOISE_NOTE)
    "dp": within(vn(DP_AVG), LIM["dp"]),
    "ach": ACH >= LIM["ach"],
    "dust": all(PARTICLES[f]["c"][i] <= lim for f in POINT_FILES for i, lim in zip((1, 2, 5), ISO8_LIMITS)),
}


# ---------------------------------------------------------------------------
# Ảnh, sơ đồ
# ---------------------------------------------------------------------------
PHOTOS = [
    ("20261029_143626.jpg", "Đo nồng độ hạt bụi", "Airborne particle counting"),
    ("20261029_144843.jpg", "Đo nhiệt độ, độ ẩm", "Temperature and humidity measurement"),
    ("20261029_145229.jpg", "Đo độ ồn", "Noise level measurement"),
    ("20261029_150836.jpg", "Đo lưu lượng khí cấp qua lọc HEPA", "Airflow rate measurement through HEPA filter"),
    ("20261029_153617.jpg", "Đo chênh lệch áp suất", "Pressure differential measurement"),
]


def _optimized(path: str, max_w: int = 1200) -> str:
    """Ảnh nén trong <thư mục ảnh>/_optimized (AGENTS mục 1, 7): xoay EXIF, thu nhỏ, không crop."""
    from PIL import Image, ImageOps
    out_dir = os.path.join(os.path.dirname(path), "_optimized")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".jpg")
    im = ImageOps.exif_transpose(Image.open(path))
    if im.height > im.width:
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=88)
    return out


# Số viết tay (bút xanh) của GĐV trên mặt bằng — hộp (x0, y0, x1, y1) trên ảnh đã xoay (scan tr. 19, 600 dpi,
# clip 72,455,180,615) -> xóa nét bút, gõ lại bằng chữ in (người dùng yêu cầu 04/11/2026).
# Ô chữ nhật xanh đậm (vị trí đo ồn, ~303,428–410,490) là ký hiệu, không phải chữ viết -> giữ.
HANDWRITTEN = [
    ((1020, 377, 1085, 469), "1", "dust"), ((720, 389, 788, 478), "2", "dust"), ((489, 356, 560, 451), "3", "dust"),
    ((180, 284, 251, 376), "4", "dust"), ((504, 497, 581, 568), "5", "dust"),
    ((984, 260, 1040, 328), "HEPA 1", "hepa"), ((684, 260, 752, 310), "HEPA 2", "hepa"),
    ((462, 260, 494, 301), "HEPA 3", "hepa"),
]


def _erase_handwriting(plan):
    """Xóa nét bút xanh trong các hộp HANDWRITTEN (số in thay bằng ký hiệu draw.io kéo thả được)."""
    import numpy as np
    from PIL import Image, ImageFilter
    a = np.asarray(plan).astype(int)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    ink = (b > r + 20) & (b > g + 5)                   # nét bút xanh (kể cả viền nhạt)
    out = np.asarray(plan).copy()
    for (x0, y0, x1, y1), _, _ in HANDWRITTEN:
        x0, y0, x1, y1 = x0 - 4, y0 - 4, x1 + 4, y1 + 4
        m = np.zeros(ink.shape, bool)
        m[y0:y1, x0:x1] = ink[y0:y1, x0:x1]
        mm = Image.fromarray((m * 255).astype("uint8")).filter(ImageFilter.MaxFilter(5))
        out[np.asarray(mm) > 0] = 255
    return Image.fromarray(out)


def _make_plan():
    import fitz
    from PIL import Image
    os.makedirs(DIAGRAM_DIR, exist_ok=True)
    raw = os.path.join(DIAGRAM_DIR, "_zone1_raw.png")
    fitz.open(SCAN)[18].get_pixmap(dpi=600, clip=fitz.Rect(72, 455, 180, 615)).save(raw)
    plan = Image.open(raw).convert("RGB").rotate(-90, expand=True)
    os.remove(raw)
    return _erase_handwriting(plan)


def build_layouts(overwrite: bool = False) -> dict:
    """Sơ đồ = vùng "Zone 1" trên mặt bằng khách hàng (scan tr. 19, GĐV đánh dấu vị trí), xoay ngang, số viết tay
    thay bằng ký hiệu draw.io kéo thả được, chú thích VN/EN Times New Roman (AGENTS 8C) -> TBKQ/_diagrams.
    File TBKQ/_diagrams/so_do_vi_tri.drawio: người dùng sửa tab _vn rồi chạy lại (Core/Tools/drawio_layout.py)."""
    from Core.Tools.drawio_layout import ensure_layout
    markers = []
    for (x0, y0, x1, y1), text, kind in HANDWRITTEN:
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        if kind == "hepa":
            markers.append({"preset": "hepa", "text": text, "x": cx - 12, "y": cy + 4})
        else:
            markers.append({"preset": "dust", "text": text, "x": cx, "y": cy, "font_size": 46})
    sym = [{"shape": "ellipse", "stroke": "#E60000", "stroke_width": 5, "w": 28, "h": 28},
           {"shape": "ellipse", "stroke": "#F0D200", "stroke_width": 5, "w": 46, "h": 26},
           {"shape": "rect", "stroke": "#1428AA", "stroke_width": 5, "w": 46, "h": 22},
           {"shape": "arrow", "color": "#006E28", "stroke_width": 5, "w": 42, "dir": "down"},
           {"shape": "rect", "fill": "#828282", "stroke": "#000000", "stroke_width": 2, "w": 34, "h": 34}]
    legend = {
        "vn": list(zip(sym, ["Vị trí đo nồng độ hạt bụi (1 - 5)", "Vị trí đo nhiệt độ, độ ẩm", "Vị trí đo độ ồn",
                             "Vị trí đo chênh lệch áp suất", "Miệng cấp khí qua lọc HEPA (HEPA 1 - 3)"])),
        "en": list(zip(sym, ["Particle concentration measuring point (1 - 5)",
                             "Temperature and humidity measuring point", "Noise level measuring point",
                             "Pressure differential measuring point", "Supply air through HEPA filter (HEPA 1 - 3)"])),
    }
    # chú thích EN dài -> 1 cột (tránh chồng chữ)
    pngs = ensure_layout(DIAGRAM_DIR, _make_plan, markers, legend=legend, legend_cols={"vn": 2, "en": 1},
                         overwrite=overwrite)
    return {lang: [{"path": p, "caption_vn": "", "caption_en": ""}] for lang, p in pngs.items()}


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def build_dossier(layouts) -> Dossier:
    photos = [{"path": _optimized(os.path.join(PHOTO_DIR, f)), "caption_vn": c_vn, "caption_en": c_en}
              for f, c_vn, c_en in PHOTOS]
    std_vn = ("ISO 14644-1:2015;\nMức quy định đối với nhiệt độ, độ ẩm, chênh lệch áp suất và số lần trao đổi khí "
              f"được cung cấp trong {OUTLINE_VN}")          # độ ồn không đánh giá (NOISE_NOTE)
    std_en = ("ISO 14644-1:2015;\nSpecified limits for temperature, humidity, pressure differential and air "
              f"changes per hour provided in {OUTLINE_EN}")
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False, dossier_type="cleanroom",
        scope_vn="Giám định chất lượng phòng sạch", scope_en="Cleanroom performance inspection",
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01", procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01",
        standards_vn=std_vn, standards_en=std_en,
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),          # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam
        equipments=[   # theo BBLM mục 7 (AGENTS mục 14)
            eq("aerotrak_9350"),
            eq("testo_625"),
            eq("tsi_8375"),
            eq("rion_nl53"),
            eq("testo_420"),
        ],
        locations=[MonitoringLocation(
            name_vn=ROOM_VN, name_en=ROOM_EN, area_m2="26", clean_class="ISO 8",
            state_vn="Trạng thái nghỉ", state_en="At rest",
            points_summary_vn=["Nồng độ hạt bụi: 05 vị trí;", "Nhiệt độ - độ ẩm: 01 vị trí;",
                               "Chênh lệch áp suất: 01 vị trí;", "Độ ồn: 01 vị trí;",
                               "Lưu lượng gió qua lọc HEPA: 03 vị trí."],
            points_summary_en=["Particle concentration: 05 points;", "Temperature - Humidity: 01 point;",
                               "Pressure differential: 01 point;", "Noise level: 01 point;",
                               "Airflow rate through HEPA filter: 03 points."],
        )],
        photos=photos, layouts=layouts["vn"],
        results_data={
            "bui_3kenh": [{"cap": "ISO 8", "gioi_han": ISO8_LIMITS, "trang_so_do": TRANG_SO_DO,
                           "diem": [{"khu_vuc": ROOM_VN, "vi_tri": i + 1,
                                     "gia_tri": [PARTICLES[f]["c"][1], PARTICLES[f]["c"][2], PARTICLES[f]["c"][5]]}
                                    for i, f in enumerate(POINT_FILES)]}],
            "nhiet_am": {"dong": [{"khu_vuc": ROOM_VN, "nhiet_do": vn(T_AVG), "do_am": vn(RH_AVG), "do_on": vn(NOISE)}],
                         "muc_nhiet_do": LIM["t"], "muc_do_am": LIM["rh"], "muc_do_on": "-"},
            "chenh_ap": {"dong": [{"khu_vuc_1": ROOM_VN, "khu_vuc_2": "Hành lang", "khu_vuc_2_en": "Corridor",
                                   "gia_tri": vn(DP_AVG), "muc": LIM["dp"]}]},
            "luu_luong": [{"muc_ach": LIM["ach"], "trang_so_do": TRANG_SO_DO,
                           "diem": [{"khu_vuc": ROOM_VN, "vi_tri": f"HEPA {h}", "luu_luong": HEPA[h]} for h in sorted(HEPA)],
                           "phong": {ROOM_VN: {"tong": TOTAL, "the_tich": vn(VOLUME), "ach": vn(ACH)}}}],
        },
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX (cùng cách build_133266 / build_131671)
# ---------------------------------------------------------------------------
def _ptext(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _highlight_placeholders(doc):
    for r in list(doc.element.body.iter(qn("w:r"))):
        t = r.find(qn("w:t"))
        if t is None or PLACEHOLDER not in (t.text or ""):
            continue
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        if rpr.find(qn("w:highlight")) is None:
            h = OxmlElement("w:highlight")
            h.set(qn("w:val"), "yellow")
            rpr.append(h)


def _fix_cover_lists(doc):
    for row in doc.tables[0].rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if not p.text.lstrip().startswith("- "):
                    continue
                pPr = p._p.get_or_add_pPr()
                for el in pPr.findall(qn("w:numPr")) + pPr.findall(qn("w:ind")):
                    pPr.remove(el)
                ind = OxmlElement("w:ind")
                ind.set(qn("w:left"), "0")
                ind.set(qn("w:hanging"), "0")
                pPr.append(ind)


def _keep(p_el):
    pPr = p_el.get_or_add_pPr()
    if pPr.find(qn("w:keepNext")) is None:
        pPr.insert(0, OxmlElement("w:keepNext"))


def _keep_headings_with_tables(doc):
    for el in doc.element.body.iterchildren(qn("w:tbl")):
        prev = el.getprevious()
        if prev is not None and prev.tag == qn("w:p") and _ptext(prev).strip():
            _keep(prev)
            rows = el.findall(qn("w:tr"))
            for tr in rows[:min(5, len(rows) - 1)]:
                for p_el in tr.iter(qn("w:p")):
                    _keep(p_el)


def _is_note_table(el) -> bool:
    return el is not None and el.tag == qn("w:tbl") and _ptext(el).strip().startswith(("Ghi chú", "Note"))


def _keep_tails_together(doc):
    def keep(tr_list):
        for tr in tr_list:
            for p_el in tr.iter(qn("w:p")):
                _keep(p_el)
    for tbl in list(doc.element.body.iterchildren(qn("w:tbl"))):
        note, gaps = tbl.getnext(), []
        while note is not None and note.tag == qn("w:p") and not _ptext(note).strip():
            gaps.append(note)
            note = note.getnext()
        if _is_note_table(tbl) or not _is_note_table(note):
            continue
        keep(gaps)
        after = note.getnext()
        assess = after is not None and after.tag == qn("w:p") and _ptext(after).strip().startswith(("Đánh giá", "Assessment"))
        keep(tbl.findall(qn("w:tr"))[-2:])
        note_rows = note.findall(qn("w:tr"))
        keep(note_rows if assess else note_rows[:-1])


def _superscript_note_marks(doc):
    from docx.text.run import Run
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        if not _is_note_table(tbl):
            continue
        for r_el in list(tbl.iter(qn("w:r"))):
            ts = r_el.findall(qn("w:t"))
            m = re.match(r"^\s*(\(\d\))(.*)$", ts[0].text or "") if len(ts) == 1 else None
            if not m:
                continue
            run = Run(r_el, None)
            if run.font.superscript:
                continue
            if m.group(2):
                rest = copy.deepcopy(r_el)
                r_el.addnext(rest)
                Run(rest, None).text = m.group(2)
            run.text = m.group(1)
            run.font.superscript = True


def _drop_section_gaps(doc):
    """Đoạn trống (không run) còn sót giữa câu Đánh giá và tiêu đề mục kế (chỗ template xóa mục vi sinh/HEPA/scan…)
    -> bỏ hết, tiêu đề mục sau đặt space_before 6pt (người dùng yêu cầu 04/11/2026: bỏ khoảng trắng giữa các mục).
    Đoạn trống giữa bảng kết quả và bảng Ghi chú (đứng sau bảng) giữ nguyên."""
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
    body = doc.element.body
    for el in list(body.iterchildren(qn("w:p"))):
        if el.getparent() is None or _ptext(el).strip() or el.find(qn("w:r")) is not None \
                or el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            continue
        prev = el.getprevious()
        while prev is not None and prev.tag == qn("w:p") and not _ptext(prev).strip() and prev.find(qn("w:r")) is None:
            prev = prev.getprevious()
        nxt = el.getnext()
        while nxt is not None and nxt.tag == qn("w:p") and not _ptext(nxt).strip() and nxt.find(qn("w:r")) is None:
            nxt = nxt.getnext()
        if prev is None or nxt is None or prev.tag != qn("w:p") or nxt.tag != qn("w:p") or not _ptext(nxt).strip():
            continue
        el.getparent().remove(el)
        Paragraph(nxt, None).paragraph_format.space_before = Pt(6)
    # Câu kết luận sau "Đánh giá:" -> keepLines, KHÔNG keepNext (AGENTS mục 3): bỏ đoạn trống xong, keepNext ở câu
    # kết luận nối chuỗi sang mục sau (9.4 + 9.5 dài hơn 1 trang) -> Word bỏ tiêu đề 9.4 lơ lửng cuối trang.
    for el in body.iterchildren(qn("w:p")):
        prev = el.getprevious()
        if prev is None or prev.tag != qn("w:p") or not _ptext(prev).strip().startswith(("Đánh giá", "Assessment")):
            continue
        pf = Paragraph(el, None).paragraph_format
        pf.keep_with_next = False
        pf.keep_together = True


def _drop_trailing_empty(doc):
    """Bỏ đoạn trống cuối văn bản, kể cả đoạn trống nằm ngay trước đoạn chứa đường kết thúc
    (đoạn trống space_before 12pt đẩy đường kết thúc sang trang mới — 138143 lần dựng đầu ra trang 6 trắng)."""
    MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
    paras = list(doc.element.body.iterchildren(qn("w:p")))
    for el in reversed(paras):
        if el.find(qn("w:r")) is not None or el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            break
        el.getparent().remove(el)
    paras = list(doc.element.body.iterchildren(qn("w:p")))
    ending = next((p for p in reversed(paras) if next(p.iter(MC), None) is not None), None)
    if ending is None:
        return
    prev = ending.getprevious()
    while prev is not None and prev.tag == qn("w:p") and not _ptext(prev).strip() \
            and next(prev.iter(MC), None) is None and prev.find(qn("w:r")) is None:
        nxt = prev.getprevious()
        prev.getparent().remove(prev)
        prev = nxt


def _appendix_total_from_numpages(doc):
    """Footer phụ lục: SECTIONPAGES -> {= {NUMPAGES} - 1}, cùng cỡ chữ (Core: tag_context.appendix_total_field)."""
    from Core.Builders.tag_context import appendix_total_field
    return appendix_total_field(doc)


def _set_cell_runs(cell, text):
    """Ghi lại chữ trong ô (giữ định dạng run đầu)."""
    p = cell.paragraphs[0]
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def _fix_dust_point_note(doc, is_en, client=CLIENT):
    """Ghi chú bảng khu vực: số vị trí đo bụi do khách hàng chỉ định (BBLM), không theo Bảng A.1 (26 m² -> 7 vị trí)."""
    new = (f"Number of particle concentration measuring points was indicated by {client.name_en}." if is_en else
           f"Số lượng vị trí đo nồng độ hạt bụi được chỉ định bởi {client.name_vn}.")
    hit = 0
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    if re.search(r"(bảng A1|Table A\.?1|table A\.?1)", p.text) and p.text.strip() != new:
                        runs = p.runs
                        runs[0].text = new
                        for r in runs[1:]:
                            r._r.getparent().remove(r._r)
                        hit += 1
    if not hit:
        print("[WARNING] không tìm thấy ghi chú 'bảng A1' để sửa")


def _bold_fail_values(doc, is_en):
    """Giá trị không đạt in đậm (AGENTS mục 3): độ ẩm, độ ồn bảng nhiệt ẩm."""
    targets = []
    if OK["t"] is False:
        targets.append(vn(T_AVG))
    if OK["rh"] is False:
        targets.append(vn(RH_AVG))
    if OK["noise"] is False:
        targets.append(vn(NOISE))
    if OK["dp"] is False:
        targets.append(vn(DP_AVG))
    if OK["ach"] is False:
        targets.append(vn(ACH))
    targets = [x.replace(",", ".") if is_en else x for x in targets]
    for t in doc.tables:
        head = t.rows[0].cells[0].text if t.rows else ""
        if not any(k in t._tbl.xpath("string(.)") for k in ("%RH", "Pa", "ACH")):
            continue
        for row in t.rows[1:]:
            if "Mức quy định" in row.cells[0].text or "limit" in row.cells[0].text.lower():
                continue
            for c in row.cells:
                if c.text.strip() in targets:
                    for p in c.paragraphs:
                        for r in p.runs:
                            r.bold = True


def _fix_temp_assessment(doc, is_en):
    """Câu đánh giá nhiệt độ - độ ẩm - độ ồn: kết quả lẫn đạt / không đạt -> tách theo chỉ tiêu."""
    names_vn = {"t": "nhiệt độ", "rh": "độ ẩm", "noise": "độ ồn"}
    names_en = {"t": "temperature", "rh": "humidity", "noise": "noise level"}
    names = names_en if is_en else names_vn
    ok = [names[k] for k in ("t", "rh", "noise") if OK[k] is True]
    ng = [names[k] for k in ("t", "rh", "noise") if OK[k] is False]
    def join(xs):
        return xs[0] if len(xs) == 1 else ", ".join(xs[:-1]) + (" and " if is_en else " và ") + xs[-1]
    room = f"“{ROOM_EN if is_en else ROOM_VN}”"
    src = OUTLINE_EN if is_en else OUTLINE_VN
    if not ng or not ok:                           # một kết luận chung (độ ồn không đánh giá -> không nêu)
        xs, good = (ok, True) if ok else (ng, False)
        if is_en:
            txt = (f"The measurement result of {join(xs)} in {room} {'complied' if good else 'did not comply'} with "
                   f"the corresponding permissible limits provided in {src}.")
        else:
            txt = (f"Kết quả đo {join(xs)} tại {room} {'phù hợp' if good else 'không phù hợp'} với mức quy định "
                   f"tương ứng được cung cấp trong {src}.")
    elif is_en:
        txt = (f"The measurement result of {join(ok)} in {room} complied with, and the measurement result of "
               f"{join(ng)} did not comply with the corresponding permissible limits provided in {OUTLINE_EN}.")
    else:
        txt = (f"Kết quả đo {join(ok)} tại {room} phù hợp, kết quả đo {join(ng)} không phù hợp với mức quy định "
               f"tương ứng được cung cấp trong {OUTLINE_VN}.")
    key = ("temperature", "humidity") if is_en else ("nhiệt độ", "độ ẩm")
    for p in doc.paragraphs:
        s = p.text.strip()
        if all(k in s.lower() for k in key) and any(w in s for w in ("phù hợp", "compl", "did not", "………")) and not s.startswith(("Đánh giá", "Assessment")):
            set_assessment_paragraph(p, txt, font_size=13)
            return
    print("[WARNING] không tìm thấy câu đánh giá nhiệt độ - độ ẩm")


def _noise_not_assessed(doc, is_en):
    """Độ ồn không đánh giá: tiêu đề cột 'Độ ồn (4), dBA' + dòng ghi chú (4) dưới bảng nhiệt ẩm (NOISE_NOTE)."""
    import copy
    from docx.table import Table, _Row
    head = "Noise" if is_en else "Độ ồn"
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        s = tbl.xpath("string(.)")
        if "%RH" not in s or head not in s:
            continue
        t = Table(tbl, doc)
        cell = next(c for c in t.rows[0].cells if c.text.strip().startswith(head))
        r0 = cell.paragraphs[0].runs[0]
        assert r0.text.startswith(head), r0.text
        rest = r0.text[len(head):]
        r0.text = head + " "
        sup = copy.deepcopy(r0._r)
        r0._r.addnext(sup)
        tail = copy.deepcopy(r0._r)
        sup.addnext(tail)
        from docx.text.run import Run
        Run(sup, r0._parent).text = "(4)"
        Run(sup, r0._parent).font.superscript = True
        Run(tail, r0._parent).text = rest
        nt = tbl.getnext()
        while nt is not None and nt.tag != qn("w:tbl"):
            nt = nt.getnext()
        assert nt is not None and _is_note_table(nt)
        notes = Table(nt, doc)
        src = next(r for r in notes.rows if r.cells[0].text.strip().startswith("(3)"))
        new = copy.deepcopy(src._tr)
        src._tr.addnext(new)
        row = _Row(new, notes)
        _set_cell_runs(row.cells[0], "(4) …")
        _set_cell_runs(row.cells[1], NOISE_NOTE_EN if is_en else NOISE_NOTE_VN)
        return
    print("[WARNING] không tìm thấy bảng nhiệt ẩm - độ ồn")


def _outline_in_assessments(doc, is_en, client=CLIENT):
    """'… được cung cấp bởi <KH>.' -> '… được cung cấp trong “Đề cương …”.' ở câu Đánh giá (như TBKQ 88904)."""
    old = f"provided by {client.name_en}." if is_en else f"được cung cấp bởi {client.name_vn}."
    new = f"provided in {OUTLINE_EN}." if is_en else f"được cung cấp trong {OUTLINE_VN}."
    for p in doc.paragraphs:
        s = p.text.strip()
        if s.endswith(old) and any(w in s for w in ("phù hợp", "compl")):
            set_assessment_paragraph(p, s[: -len(old)] + new, font_size=13)


def _fit_layout(doc):
    """Lấp trang phụ lục (người dùng yêu cầu 05/11/2026): dòng dữ liệu bảng nhiệt ẩm / chênh áp cao bằng các dòng
    khác (454 twips, template để 1361 / 1247); cột 'Khu vực đo 1' rộng ra để tên phòng nằm 1 dòng;
    sơ đồ 13 -> 15,8 cm (khổ chữ 16,5 cm); ảnh grid 3 cột 4,8 -> 5,0 cm (AGENTS mục 3: 4 - 5 cm)."""
    W = qn("w:w")
    for t in doc.tables:
        s = t._tbl.xpath("string(.)")
        if "%RH" in s and ("Độ ồn" in s or "Noise" in s):
            rows = t.rows[1:-1]                              # dòng dữ liệu (bỏ tiêu đề, dòng mức)
        elif "Khu vực đo 2" in s or "Measurement area 2" in s or "area 2" in s.lower():
            rows = t.rows[1:]
            # pct (tổng 5000): EN tên phòng dài -> cột 1 rộng ~7,3 cm để nằm 1 dòng (giữ 9.4 cùng trang với VN)
            widths = [2212, 303, 790, 848, 847] if "Area 1" in s else [1300, 303, 1150, 1100, 1147]
            grid = t._tbl.find(qn("w:tblGrid"))
            tot = sum(int(c.get(W)) for c in grid)
            for c, w in zip(grid, widths):
                c.set(W, str(round(tot * w / 5000)))
            for r in t.rows:
                for c, w in zip(r._tr.findall(qn("w:tc")), widths):
                    tcw = c.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
                    if tcw is not None and tcw.get(qn("w:type")) == "pct":
                        tcw.set(W, str(w))
        else:
            continue
        for r in rows:
            h = r._tr.find(f"{qn('w:trPr')}/{qn('w:trHeight')}")
            if h is not None:
                h.set(qn("w:val"), "454")
                h.set(qn("w:hRule"), "atLeast")
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for inl in doc.element.body.iter(f"{ns}inline"):
        ext = inl.find(f"{ns}extent")
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
        new = {4680000: 5688000, 1728000: 1800000}.get(cx)      # 13 -> 15,8 cm; 4,8 -> 5,0 cm
        if not new:
            continue
        ncy = round(cy * new / cx)
        ext.set("cx", str(new)); ext.set("cy", str(ncy))
        for e in inl.iter(f"{a}ext"):
            if e.get("cx") == str(cx):
                e.set("cx", str(new)); e.set("cy", str(ncy))


def finish_tbkq(src: str, is_en: bool) -> str:
    doc = Document(src)
    _fix_cover_lists(doc)
    _fix_dust_point_note(doc, is_en, CLIENT_TBKQ)
    _outline_in_assessments(doc, is_en, CLIENT_TBKQ)
    _fix_temp_assessment(doc, is_en)
    _noise_not_assessed(doc, is_en)
    _bold_fail_values(doc, is_en)
    _fit_layout(doc)
    _keep_headings_with_tables(doc)
    _keep_tails_together(doc)
    _drop_section_gaps(doc)
    _drop_trailing_empty(doc)
    _appendix_total_from_numpages(doc)
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    lang = "en" if is_en else "viet"
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {lang} - phong sach - {CODE}.docx")
    from Core.Builders.tag_context import style_note_tables
    style_note_tables(doc)      # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3, 09/11/2026)
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)             # tên handler đặt "TBKQ - vn - 1 - …" lệch AGENTS mục 1 (E41)
    return dst


# ---------------------------------------------------------------------------
# BCKQ
# ---------------------------------------------------------------------------
def _bckq_noise_note(doc):
    """Đoạn '- Chỉ tiêu không đánh giá: Độ ồn …' ngay dưới bảng Đánh giá của BCKQ (NOISE_NOTE)."""
    import copy
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
    tbl = next(t for t in doc.tables if "Chỉ tiêu không đạt" in t._tbl.xpath("string(.)"))
    ref = next(p for p in doc.paragraphs if p.text.strip().startswith("- Kết quả đo nồng độ hạt bụi"))
    el = copy.deepcopy(ref._p)
    tbl._tbl.addnext(el)
    p = Paragraph(el, ref._parent)
    runs = p.runs
    runs[0].text = (f"- Chỉ tiêu không đánh giá: Độ ồn ({vn(NOISE)} dBA) — khách hàng xác nhận tại thời điểm thực hiện, "
                    f"khu vực xung quanh có các hoạt động tạo nguồn ồn khác.")
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    p.paragraph_format.space_before = Pt(6)


def _bckq_client_note(doc):
    """Báo cáo GĐV: 'Các bất thường …: [X] Có' + dòng nghiêng KH bổ sung PYC tên liên danh.
    Theo bản người dùng sửa tay trên Z 05/11/2026 16:17 (chuyển ghi chú khỏi ô 'Đơn vị yêu cầu', bỏ ngày PYC)."""
    import copy
    p = next(p for p in doc.paragraphs if p.text.strip().startswith("Các bất thường trong quá trình giám định"))
    runs = p.runs
    for r in runs:
        if "] Có" in r.text:
            r.text = r.text.replace("[  ] Có", "[X] Có").replace("[X] không", "[  ] không")
    assert p.text.startswith("Các bất thường trong quá trình giám định: [X] Có"), p.text
    br = copy.deepcopy(runs[-1]._r)
    for t in [c for c in br if c.tag != qn("w:rPr")]:     # bỏ cả w:t lẫn w:tab, chỉ giữ định dạng
        br.remove(t)
    br.append(OxmlElement("w:br"))
    runs[-1]._r.addnext(br)
    note = copy.deepcopy(runs[0]._r)
    br.addnext(note)
    from docx.text.run import Run
    r = Run(note, p)
    r.text = (f"Khách hàng bổ sung Phiếu yêu cầu & nhận thẩm định/ giám định ngày với tên mới "
              f"“{CLIENT_TBKQ.name_vn}”; TBKQ ghi theo tên mới.")
    r.italic = True


def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["tham_chieu"] = [
        f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {SO_BUI} ngày {SAMPLE_DATE}.",
        *REF_NHIET_AM_ON,
        f"- Kết quả đo chênh lệch áp suất: xem tại phiếu Kết quả đo chênh lệch áp suất "
        f"số {SO_CHENH_AP} ngày {SAMPLE_DATE}.",
        f"- Kết quả đo lưu lượng và số lần trao đổi khí: xem tại phiếu Tính toán vận tốc gió và lưu lượng qua lọc HEPA "
        f"số {SO_LUU_LUONG} ngày {SAMPLE_DATE}.",
    ]
    items = [("dust", "Nồng độ hạt bụi (ISO 8)"),
             ("t", f"Nhiệt độ ({vn(T_AVG)} °C; mức {LIM['t']} °C)"),
             ("rh", f"Độ ẩm ({vn(RH_AVG)} %RH; mức {LIM['rh']} %RH)"),
             ("noise", f"Độ ồn ({vn(NOISE)} dBA; mức {LIM['noise']} dBA)"),
             ("dp", f"Chênh lệch áp suất ({vn(DP_AVG)} Pa; mức {LIM['dp']} Pa)"),
             ("ach", f"Số lần trao đổi khí ({vn(ACH)} lần/giờ; mức ≥ {LIM['ach']})")]
    dat = [t for k, t in items if OK[k] is True]
    khong = [t for k, t in items if OK[k] is False]
    ctx["danh_gia"] = [{
        "khu_vuc": f"“{ROOM_VN}”",
        "dat": ";\n".join(dat) or "Không",
        "khong_dat": ";\n".join(khong) or "Không",
        "can_cu": ("ISO 14644-1:2015 (cấp sạch ISO 8);\nMức quy định nhiệt độ, độ ẩm, chênh áp, số lần trao đổi "
                   f"khí trong {OUTLINE_VN} do {CLIENT.name_vn} cung cấp."),
    }]
    out = b.build(ctx=ctx, out_name=f"BCKQ - 1 - phong sach - {CODE}.docx")
    doc = Document(out)
    _fix_dust_point_note(doc, is_en=False)
    _bckq_noise_note(doc)
    _bckq_client_note(doc)
    _keep_headings_with_tables(doc)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Hình ảnh", "Sơ đồ", "Báo cáo của GĐV")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    return out


# ---------------------------------------------------------------------------
# Excel phiếu tính toán (ThucHien) — công thức động (AGENTS mục 11)
# ---------------------------------------------------------------------------
# Số phiếu tính toán = mã + đuôi chỉ tiêu (AGENTS mục 7); nhiệt ẩm và độ ồn tách 2 phiếu (người dùng chốt 07/11/2026)
SO_BUI = f"{CODE}/Bụi"
SO_NHIET_AM = f"{CODE}/Nhiệt độ, độ ẩm"
SO_DO_ON = f"{CODE}/Độ ồn"
SO_LUU_LUONG = f"{CODE}/Lưu lượng"
SO_CHENH_AP = f"{CODE}/Chênh áp"
REF_NHIET_AM_ON = [
    f"- Kết quả đo nhiệt độ, độ ẩm: xem tại phiếu Tính toán kết quả nhiệt độ, độ ẩm số {SO_NHIET_AM} ngày {SAMPLE_DATE}.",
    f"- Kết quả đo độ ồn: xem tại phiếu Tính toán kết quả độ ồn số {SO_DO_ON} ngày {SAMPLE_DATE}.",
]


def build_excel() -> str:
    """Phiếu tính toán theo form phòng sạch chung (Core/Tools/cleanroom_calc.py, người dùng chốt 07/11/2026)."""
    book = CleanroomCalcBook()
    labels, recs = read_aerotrak(DUST_XLSX)
    book.dust_aerotrak([{"room": ROOM_VN, "cap": "ISO 8",
                         "limits": dict(zip(("Ch2", "Ch3", "Ch6"), ISO8_LIMITS)),
                         "zero": recs[ZERO_FILE], "points": [recs[f] for f in POINT_FILES]}],
                       so=SO_BUI, date=SAMPLE_DATE, labels=labels)
    book.temp_rh([{"room": ROOM_VN, "loc": 1, "t": TEMP_READ, "rh": RH_READ,
                   "lim_t": LIM["t"], "lim_rh": LIM["rh"]}], CAL_T_TBL, CAL_RH_TBL, so=SO_NHIET_AM, date=SAMPLE_DATE)
    # Độ ồn không đánh giá (khách hàng xác nhận có nguồn ồn khác xung quanh — người dùng chốt 04/11/2026) -> mức "-"
    book.noise([{"room": ROOM_VN, "loc": 1, "reads": [NOISE], "lim": None}],
               so=SO_DO_ON, date=SAMPLE_DATE)
    book.pressure([{"a": ROOM_VN, "b": "Hành lang", "sign": "PA1", "reads": DP_READ, "lim": LIM["dp"]}],
                  so=SO_CHENH_AP, date=SAMPLE_DATE, cal=CAL_DP_TBL)
    hepas = []
    for hp in sorted(HEPA):
        hepas.append({"name": f"HEPA {hp}", "parts": [
            {"part": f"{hp}-{part}", "date": SAMPLE_DATE, "time": when.split(" ", 1)[1], "reads": reads}
            for h2, part, when, reads, _avg in FLOWS if h2 == hp]})
    book.flow([{"room": ROOM_VN, "area": AREA_M2, "height": HEIGHT_M, "lim_ach": f"≥ {LIM['ach']}", "hepas": hepas}],
              so=SO_LUU_LUONG, date=SAMPLE_DATE)
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - TinhToanPhongSach.xlsx")
    book.save(out)
    print(f"[build_138143] Excel -> {out}")
    return out


def patch_bckq_refs() -> str:
    """Chỉ dựng lại phiếu: sửa dòng viện dẫn nhiệt ẩm - ồn trong BCKQ đã có (giữ ngày, định dạng)."""
    path = os.path.join(INPUT_DIR, "ThucHien", f"BCKQ - 1 - phong sach - {CODE}.docx")
    ok = replace_reference_lines(path, "- Kết quả đo nhiệt độ, độ ẩm, độ ồn", REF_NHIET_AM_ON)
    print(f"[build_138143] BCKQ viện dẫn nhiệt ẩm / ồn: {'đã sửa' if ok else 'không thấy dòng cũ (giữ nguyên)'}")
    return path


def split_edoc() -> list:
    """BBLM = scan tr. 11–12; Ghi chép = tr. 21–25 (BM3030.01) + tr. 26 (giấy tính ACH) -> TBKQ/Edoc (AGENTS 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(SCAN)
    outs = []
    for name, (a, b) in ((f"{CODE6} - BBLM.pdf", (10, 11)), (f"{CODE6} - Ghi Chep.pdf", (20, 25))):
        dst = fitz.open()
        dst.insert_pdf(src, from_page=a, to_page=b)
        p = os.path.join(edoc, name)
        finish_edoc(dst, INPUT_DIR, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_138143] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--tao-lai-so-do", action="store_true",
                    help="Dựng lại so_do_vi_tri.drawio từ tọa độ trong code (bỏ chỉnh tay trên draw.io)")
    ap.add_argument("--chi-so-do", action="store_true", help="Chỉ dựng/xuất sơ đồ, không dựng DOCX/XLSX")
    ap.add_argument("--chi-excel", action="store_true",
                    help="Chỉ dựng lại phiếu tính toán (+ sửa dòng viện dẫn BCKQ); giữ TBKQ, ngày ban hành đã giao")
    args = ap.parse_args()
    if args.chi_so_do:
        print(build_layouts(args.tao_lai_so_do))
        return
    if args.chi_excel:
        xlsx = build_excel()
        bckq = patch_bckq_refs()
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            ex = PDFExporter(build_dossier(build_layouts(False)))
            ex.export_files([bckq])
            ex.export_excel_files([xlsx])
        return

    print(f"[build_138143] T {vn(T_AVG)} °C, RH {vn(RH_AVG)} %RH, ồn {vn(NOISE)} dBA, ΔP {vn(DP_AVG)} Pa, "
          f"HEPA {HEPA}, tổng {TOTAL} m³/h, V {vn(VOLUME)} m³, ACH {vn(ACH)} | đạt: {OK}")
    layouts = build_layouts(args.tao_lai_so_do)
    d = build_dossier(layouts)
    d.ensure_output_directories()
    split_edoc()
    d.client = CLIENT_TBKQ                 # chỉ TBKQ ghi tên liên danh (PYC bổ sung 28/10/2026)
    h = CleanroomHandler(d)
    vn_doc = finish_tbkq(h.build_tbkq(is_en=False), is_en=False)
    d.layouts = layouts["en"]
    en_doc = finish_tbkq(h.build_tbkq(is_en=True), is_en=True)
    d.layouts = layouts["vn"]
    d.client = CLIENT                      # BCKQ + Excel: tên theo BBLM
    bckq = build_bckq(d)
    xlsx = build_excel()
    docs = [bckq, vn_doc, en_doc]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(d)
        ex.export_files(docs)
        ex.export_excel_files([xlsx])
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=d.code).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
