# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D132809 - Công ty TNHH Gene Solutions Lab (phòng sạch ISO 5 + ISO 7) — 1 chứng thư (mã trần).

Địa điểm: 186-188 Nguyễn Duy Dương, phường Vườn Lài, Thành phố Hồ Chí Minh (BBLM).
Khu vực (PYC, báo giá, phiếu ghi chép): "Phòng sạch ISO 5" 8,67 m²; "Phòng sạch ISO 7" 5,64 m².
Chỉ tiêu: nồng độ hạt bụi ISO 5 (0,1–5,0 µm) 05 vị trí x 02 trạng thái; ISO 7 (0,5–5,0 µm) 03 vị trí x 02 trạng thái;
          mỗi phòng nhiệt độ - độ ẩm, độ ồn, chênh áp 01 vị trí.
Hai lần đi hiện trường:
  - 29/10/2026 09:00–10:30 (BBLM scan tr. 15–16): nhiệt độ - độ ẩm, độ ồn, chênh áp (bụi "hoạt động" bị gạch trên BBLM;
    các file bụi sáng 29/10 trong database không được dùng — phiếu ghi chép chỉ ghi file ngày 04/11).
  - 04/11/2026 14:30–16:30 (BBLM lần 2 "Q326D132809/2", người dùng tự đặt ở TBKQ/Edoc/132809 - BBLM - 2.pdf — file của
    người dùng, không ghép/không ghi đè; database Lasair III 110 15:25–15:57, file 1666 zero, 1667–1689 đo):
    nồng độ hạt bụi nghỉ + hoạt động, ISO 5: 5 + 5, ISO 7: 3 + 3 — khớp BBLM.
Mức quy định: chỉ ISO 14644-1:2015 cho bụi (ISO 5, ISO 7). Khách hàng không cung cấp mức nhiệt độ, độ ẩm, độ ồn, chênh áp
          -> không ghi dòng mức, không đánh giá (AGENTS mục 3).
Người dùng yêu cầu (04/11/2026): mục bụi tách 9.2 "Kết quả đo nồng độ hạt bụi" -> 9.2.1 trạng thái nghỉ, 9.2.2 hoạt động.

Nguồn: Input/132809/20261104165659841.pdf (22 tr.: PKSCV 1–2, PYC 3–6, báo giá 7–9, PO 11–13, BBLM 15–16,
       mặt bằng có đánh dấu vị trí 17, ghi chép 19–21; tr. 10, 14, 18, 22 trắng),
       Input/132809/database_all 0411.csv (Lasair III 110 S/N 137662; Data ID = số thứ tự dòng),
       Input/132809/132809 - TinhPhongSach.xlsx (người dùng — chỉ đối chiếu sheet "Bụi 0,1"; các sheet khác là số liệu cũ),
       Input/132809/Auto_0176, Auto_0177 (Rion NL-53, LAeq 1 phút), Photos-1-001 (4)/Use (4 ảnh).

Chạy:  python build_132809.py          (DOCX + XLSX + Edoc + sơ đồ)
       python build_132809.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: trong code (network_dir để trống) — chép lên Z bằng robocopy (AGENTS mục 0).
"""

import os
import re
import sys
import csv
import copy
import argparse
from decimal import Decimal, ROUND_HALF_UP

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
from Core.Builders.tag_context import PLACEHOLDER
from Core.Styling.paragraphs import set_assessment_paragraph
from Core.Tools.excel_borders import thin_outside_borders
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)
from Core.Tools.cleanroom_calc import (CleanroomCalcBook, CalTable, read_lasair_csv, apply_cal,
                                       replace_reference_lines)

CODE = "Q326D132809"
CODE6 = "132809"
REG_DATE = "20/10/2026"            # PYC
DATE1, START1, END1 = "29/10/2026", "09:00", "10:30"      # BBLM lần 1
DATE2, START2, END2 = "04/11/2026", "14:30", "16:30"      # BBLM lần 2 (máy Lasair 15:25–15:57 nằm trong khoảng)
import datetime as _dt
CERT_DATE = _dt.date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày hiện tại, tự cập nhật mỗi lần dựng (người dùng chốt 06/11/2026, AGENTS mục 2)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN = os.path.join(INPUT_DIR, "20261104165659841.pdf")
DUST_CSV = os.path.join(INPUT_DIR, "database_all 0411.csv")
USER_XLSX = os.path.join(INPUT_DIR, "132809 - TinhPhongSach.xlsx")
NOISE_RND = {"iso5": os.path.join(INPUT_DIR, "Auto_0176", "Auto_Leq", "NL_0001_OCT_Leq_0176_0001.rnd"),
             "iso7": os.path.join(INPUT_DIR, "Auto_0177", "Auto_Leq", "NL_0001_OCT_Leq_0177_0001.rnd")}
PHOTO_DIR = os.path.join(INPUT_DIR, "Photos-1-001 (4)", "Use")
DIAGRAM_DIR = os.path.join(INPUT_DIR, "TBKQ", "_diagrams")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Gene Solutions Lab",
    name_en="Gene Solutions Lab Limited Liability Company",
    # PYC mục 3 (địa chỉ pháp lý); BBLM mục 3 ghi địa chỉ nơi thực hiện -> báo người dùng
    address_vn="Tầng 3, Phòng 301, Tòa nhà WMC, Số 102 A-B-C Cống Quỳnh, phường Bến Thành, Thành phố Hồ Chí Minh",
    address_en="3rd Floor, Room 301, WMC Building, No. 102 A-B-C Cong Quynh street, Ben Thanh ward, Ho Chi Minh City",
)
LOCATION_VN = "Phòng sạch của Công ty TNHH Gene Solutions Lab - 186-188 Nguyễn Duy Dương, phường Vườn Lài, Thành phố Hồ Chí Minh"
LOCATION_EN = "Cleanroom of Gene Solutions Lab Limited Liability Company - 186-188 Nguyen Duy Duong street, Vuon Lai ward, Ho Chi Minh City"

ROOMS = {
    "iso5": {"vn": "Phòng sạch ISO 5", "en": "ISO 5 cleanroom", "area": "8,67", "cap": "ISO 5"},
    "iso7": {"vn": "Phòng sạch ISO 7", "en": "ISO 7 cleanroom", "area": "5,64", "cap": "ISO 7"},
}
ANTE_VN, ANTE_EN = "Phòng đệm", "Anteroom"         # phiếu ghi chép "P. ISO 7 → Đệm" (chữ viết tay)
STATES = {"nghi": ("trạng thái nghỉ", "at-rest state", "Nghỉ"),
          "hd": ("trạng thái hoạt động", "operational state", "Hoạt động")}

# Phiếu ghi chép tr. 19–20: file Lasair theo vị trí (ISO 5 nghỉ VT3 ghi "1669" sửa thành "1672")
FILES = {
    ("iso5", "nghi"): [1667, 1668, 1672, 1674, 1675],
    ("iso5", "hd"): [1682, 1683, 1684, 1688, 1689],
    ("iso7", "nghi"): [1676, 1677, 1678],
    ("iso7", "hd"): [1679, 1680, 1681],
}
ZERO_FILE = 1666
SIZES = ["0,1", "0,15", "0,2", "0,25", "0,3", "0,5", "1,0", "5,0"]     # Counts 1..8 (vi sai)
REPORT = {"iso5": [0, 2, 4, 5, 6, 7],      # ≥ 0,1; 0,2; 0,3; 0,5; 1; 5 µm
          "iso7": [5, 6, 7]}               # ≥ 0,5; 1; 5 µm
LIMITS = {"iso5": [100000, 23700, 10200, 3520, 832, None],   # ISO 14644-1:2015 Bảng 1 (cấp 5: ≥ 5 µm không quy định)
          "iso7": [352000, 83200, 2930]}

# Phiếu ghi chép tr. 20 (ngày 29/10/2026), Testo 625 – máy 3. ISO 5 RH lần 1 đọc "80,1" (chữ sửa đè, chưa chắc).
TEMP_READ = {"iso5": [20.5, 20.5, 20.4], "iso7": [21.1, 21.0, 21.1]}
RH_READ = {"iso5": [80.1, 79.8, 79.9], "iso7": [94.3, 94.4, 94.1]}
# Giấy hiệu chuẩn Testo 625 – máy 3 (người dùng gửi 04/11/2026, dùng ở 138143): (chỉ thị, chuẩn, sai số)
CAL_T = [(15.2, 15.00, 0.2), (20.0, 20.00, 0.0), (24.9, 25.00, -0.1), (29.8, 30.00, -0.2), (39.7, 40.00, -0.3)]
CAL_RH = [(36.1, 30.0, 6.1), (45.8, 40.0, 5.8), (54.5, 50.0, 4.5), (64.8, 60.0, 4.8), (74.3, 70.0, 4.3),
          (84.0, 80.0, 4.0)]
CAL_DEVICE = "Testo 625 – máy 3"
# Giấy hiệu chuẩn lấy từ thư viện HieuChuan/ (AGENTS 8C): Testo 625 máy 3 = MT33-3 (đo 29/10); chênh áp TSI 8375 = MT31.
# Độ ồn Rion NL-53 chưa có bảng hiệu chuẩn trong thư viện -> không hiệu chính.
CAL_T_TBL = CalTable.from_library("MT33-3", "nhiệt độ", DATE1, device=CAL_DEVICE, unit="°C", fallback=CAL_T)
CAL_RH_TBL = CalTable.from_library("MT33-3", "độ ẩm", DATE1, device=CAL_DEVICE, unit="%RH", fallback=CAL_RH)
CAL_DP_TBL = CalTable.from_library("MT31", "Chênh áp / áp suất", DATE1, device="TSI 8375", unit="Pa")
# Chênh áp TSI 8375 (Pa) — phiếu ghi chép tr. 20 (ký hiệu PA1/PA2 theo phiếu)
DP = [("iso5", "iso7", "PA1", [6.22, 6.14, 6.15]),
      ("iso7", "ante", "PA2", [12.8, 12.8, 12.7])]
TRANG_SO_DO = "2"
# Đuôi số phiếu tính toán phòng sạch (người dùng chốt 05/11/2026, AGENTS mục 7 / 8C): header sheet Excel "Số: <mã>/<đuôi>"
# và câu viện dẫn trong BCKQ. Lưu lượng (nếu có): "Lưu lượng".
# Nhiệt ẩm và độ ồn tách 2 phiếu (người dùng chốt 07/11/2026 — form phòng sạch chung Core/Tools/cleanroom_calc.py)
SUFFIX = {"bui": "Bụi", "nhiet_am": "Nhiệt độ, độ ẩm", "do_on": "Độ ồn", "chenh_ap": "Chênh áp",
          "luu_luong": "Lưu lượng"}                  # trang phụ lục chứa sơ đồ — kiểm lại trên PDF


# ---------------------------------------------------------------------------
# Đọc số liệu
# ---------------------------------------------------------------------------
def r0(x):
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def r1(x):
    return float(Decimal(str(x)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))


def vn(x, nd=1):
    return f"{x:.{nd}f}".replace(".", ",")


def read_particles():
    """database CSV -> {Data ID: {time, vol (m³), counts[8], loc}} (Data ID = số thứ tự dòng dữ liệu)."""
    rows = list(csv.DictReader(open(DUST_CSV, encoding="utf-8-sig")))
    out = {}
    for i, r in enumerate(rows, 1):
        out[i] = {"date": r["Sample Date"], "time": r["Sample Time"], "vol": float(r["Sample Volume"]),
                  "interval": r["Sample Interval"], "loc": r["Location"],
                  "counts": [int(r[f"Counts {k}"]) for k in range(1, 9)]}
    for f in [ZERO_FILE] + [x for v in FILES.values() for x in v]:
        p = out[f]
        if p["date"] != "2026/11/04" or p["interval"] != "00:01:00":
            raise SystemExit(f"[FATAL] file {f}: {p['date']} {p['interval']} — không phải mẫu 1 phút ngày 04/11")
    return out


def conc(f, k):
    """Nồng độ cộng dồn ≥ SIZES[k], hạt/m³ = Σ counts[k:] / thể tích (m³)."""
    p = PARTICLES[f]
    return sum(p["counts"][k:]) / p["vol"]


def check_user_excel():
    """Đối chiếu sheet 'Bụi 0,1' của file người dùng (chỉ đọc) với số tính lại từ CSV."""
    try:
        ws = openpyxl.load_workbook(USER_XLSX, data_only=True).worksheets[0]
    except Exception as e:      # noqa: BLE001
        print(f"[WARNING] không đọc được {USER_XLSX}: {e}")
        return
    cols = {"0,1": 53, "0,2": 57, "0,3": 61, "0,5": 63, "1,0": 65, "5,0": 67}
    bad = 0
    for r in range(2, ws.max_row + 1):
        fid = ws.cell(r, 1).value
        if not isinstance(fid, int) or fid not in PARTICLES:
            continue
        for s, c in cols.items():
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and abs(v - conc(fid, SIZES.index(s))) > 0.5:
                bad += 1
                print(f"[WARNING] [Data Mismatch] file {fid} ≥{s}: Excel {v:.1f} ≠ CSV {conc(fid, SIZES.index(s)):.1f}")
    print(f"[build_132809] Đối chiếu Excel người dùng (sheet Bụi): {'khớp' if not bad else f'{bad} ô lệch'}")


def read_noise(path):
    for line in open(path, encoding="utf-8", errors="ignore"):
        if line.startswith("Leq"):
            return float(line.split(",")[1])
    raise SystemExit(f"[FATAL] không đọc được LAeq {path}")


def cal_bracket(x, table):
    for i in range(len(table) - 1):
        if table[i][0] <= x <= table[i + 1][0]:
            return i
    return 0 if x < table[0][0] else len(table) - 2


def out_of_range(x, table):
    """Chỉ số điểm hiệu chuẩn biên khi x nằm ngoài khoảng hiệu chuẩn, None nếu trong khoảng."""
    if x > table[-1][0]:
        return len(table) - 1
    if x < table[0][0]:
        return 0
    return None


def calibrated(x, table):
    """Giá trị sau hiệu chính = chỉ thị − sai số (nội suy tuyến tính như FORECAST). Ngoài khoảng hiệu chuẩn -> trừ
    đúng sai số của điểm biên, không ngoại suy (người dùng chốt 04/11/2026: ISO 7 RH 94 %RH lấy sai số điểm 84,0)."""
    j = out_of_range(x, table)
    if j is not None:
        return x - table[j][2]
    i = cal_bracket(x, table)
    (x1, _, e1), (x2, _, e2) = table[i], table[i + 1]
    return x - (e1 + (e2 - e1) * (x - x1) / (x2 - x1))


PARTICLES = read_particles()
NOISE = {k: read_noise(p) for k, p in NOISE_RND.items()}                  # 63,6 / 62,1 dBA (khớp phiếu)
T_AVG = {k: r1(sum(apply_cal(v, CAL_T_TBL) for v in TEMP_READ[k]) / 3) for k in ROOMS}
RH_AVG = {k: r1(sum(apply_cal(v, CAL_RH_TBL) for v in RH_READ[k]) / 3) for k in ROOMS}
RH_EXTRAPOLATED = [k for k in ROOMS if max(RH_READ[k]) > CAL_RH[-1][0]]
DP_AVG = [r1(sum(apply_cal(x, CAL_DP_TBL) for x in v) / 3) for *_, v in DP]
VALUES = {key: [[r0(conc(f, k)) for k in REPORT[key[0]]] for f in files] for key, files in FILES.items()}
DUST_OK = {key: all(v <= lim for row in rows for v, lim in zip(row, LIMITS[key[0]]) if lim is not None)
           for key, rows in VALUES.items()}


# ---------------------------------------------------------------------------
# Ảnh, sơ đồ
# ---------------------------------------------------------------------------
PHOTOS = [
    ("20261104_152623.jpg", "Đo nồng độ hạt bụi ở trạng thái nghỉ", "Particle concentration measurement at rest"),
    ("20261104_155517.jpg", "Đo nồng độ hạt bụi ở trạng thái hoạt động",
     "Particle concentration measurement in operation"),
    ("20261029_100149.jpg", "Đo độ ồn", "Noise level measurement"),
    ("20261029_100233.jpg", "Đo độ ồn", "Noise level measurement"),
]


def _optimized(path: str, max_w: int = 1200) -> str:
    """Ảnh nén trong <thư mục ảnh>/_optimized (AGENTS mục 1, 7): xoay EXIF, ảnh dọc xoay ngang, không crop."""
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


# Chữ viết tay (bút xanh) của GĐV trên mặt bằng — hộp (x0, y0, x1, y1) trên ảnh đã xoay
# (scan tr. 17, 300 dpi, clip 290,222,552,660, xoay -90°) -> xóa nét bút, gõ lại bằng chữ in (như 138143).
# Mũi tên chênh áp (nét xanh dài) giữ nguyên, chỉ gõ lại nhãn PA1/PA2.
HANDWRITTEN = [
    # phòng ISO 5 (trái): vị trí bụi 1..5
    ((212, 648, 292, 748), "+1", "dust"), ((460, 655, 525, 752), "+2", "dust"), ((460, 452, 525, 550), "+3", "dust"),
    ((670, 448, 748, 548), "+4", "dust"), ((855, 468, 948, 578), "+5", "dust"),
    # phòng ISO 7 (phải): vị trí bụi 1..3
    ((1502, 782, 1572, 872), "+1", "dust"), ((1332, 652, 1418, 728), "+2", "dust"), ((1176, 432, 1250, 518), "+3", "dust"),
    # vị trí nhiệt độ - độ ẩm - độ ồn (khoanh tròn)
    ((602, 612, 725, 732), "2", "point"), ((1212, 556, 1356, 668), "1", "point"),
    # nhãn chênh áp — GĐV vẽ ngược phiếu ghi chép; đổi cho khớp phiếu (người dùng chốt 04/11/2026):
    # PA1 = ISO 5 -> ISO 7 (vách giữa 2 phòng), PA2 = ISO 7 -> phòng đệm
    ((1288, 262, 1385, 324), "PA2", "dp"), ((1052, 382, 1120, 430), "PA1", "dp"),
]


def _erase_handwriting(plan):
    """Xóa nét bút xanh trong các hộp HANDWRITTEN (chữ in thay bằng ký hiệu draw.io kéo thả được)."""
    import numpy as np
    from PIL import Image, ImageFilter
    a = np.asarray(plan).astype(int)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    ink = (b > r + 40) & (b > g + 25)                  # nét bút xanh dương (tường xanh lơ có g ≈ b -> loại)
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
    raw = os.path.join(DIAGRAM_DIR, "_plan_raw.png")
    fitz.open(SCAN)[16].get_pixmap(dpi=300, clip=fitz.Rect(290, 222, 552, 660)).save(raw)
    plan = Image.open(raw).convert("RGB").rotate(-90, expand=True)
    os.remove(raw)
    return _erase_handwriting(plan)


def build_layouts(overwrite: bool = False) -> dict:
    """Sơ đồ = vùng 2 phòng sạch + phòng đệm trên mặt bằng khách hàng (scan tr. 17, GĐV đánh dấu vị trí), xoay ngang,
    chữ viết tay xóa và thay bằng ký hiệu draw.io kéo thả được, chú thích VN/EN Times New Roman (AGENTS 8C).
    File TBKQ/_diagrams/so_do_vi_tri.drawio: người dùng sửa tab _vn rồi chạy lại (Core/Tools/drawio_layout.py)."""
    from Core.Tools.drawio_layout import ensure_layout
    markers = [{"preset": kind, "text": text, "x": (x0 + x1) / 2, "y": (y0 + y1) / 2}
               for (x0, y0, x1, y1), text, kind in HANDWRITTEN]
    rooms = {"vn": [(530, 1010, "PHÒNG SẠCH ISO 5"), (1390, 1010, "PHÒNG SẠCH ISO 7"), (1290, 95, "PHÒNG ĐỆM")],
             "en": [(530, 1010, "ISO 5 CLEANROOM"), (1390, 1010, "ISO 7 CLEANROOM"), (1290, 95, "ANTEROOM")]}
    sym = [{"preset": "dust", "font_size": 40, "text": "+n"},
           {"shape": "ellipse", "stroke": "#1428AA", "stroke_width": 5, "w": 36, "h": 36},
           {"shape": "arrow", "color": "#1E3CC8", "stroke_width": 5, "w": 66}]
    legend = {
        "vn": list(zip(sym, ["Vị trí đo nồng độ hạt bụi (ISO 5: 1 - 5; ISO 7: 1 - 3)",
                             "Vị trí đo nhiệt độ, độ ẩm, độ ồn", "Vị trí đo chênh lệch áp suất (PA1, PA2)"])),
        "en": list(zip(sym, ["Particle concentration measuring point (ISO 5: 1 - 5; ISO 7: 1 - 3)",
                             "Temperature, humidity and noise level measuring point",
                             "Pressure differential measuring point (PA1, PA2)"])),
    }
    pngs = ensure_layout(DIAGRAM_DIR, _make_plan, markers, rooms, legend, overwrite=overwrite)
    return {lang: [{"path": p, "caption_vn": "", "caption_en": ""}] for lang, p in pngs.items()}


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def points_summary(key, en):
    n = len(FILES[(key, "nghi")])
    rng = "0,1 – 5,0 µm" if key == "iso5" else "0,5 – 5,0 µm"
    if en:
        rng = rng.replace(",", ".")
        return [f"Particle concentration ({rng}): {n:02d} points × 02 states (at rest, operational);",
                "Temperature - Humidity: 01 point;", "Noise level: 01 point;", "Pressure differential: 01 point."]
    return [f"Nồng độ hạt bụi ({rng}): {n:02d} vị trí × 02 trạng thái (nghỉ, hoạt động);",
            "Nhiệt độ - độ ẩm: 01 vị trí;", "Độ ồn: 01 vị trí;", "Chênh lệch áp suất: 01 vị trí."]


def dust_group(key, state, en_flag=False):
    room = ROOMS[key]
    lim = [x if x is not None else "- (3)" for x in LIMITS[key]]
    ok = DUST_OK[(key, state)]
    st_vn, st_en, _ = STATES[state]
    return {"cap": room["cap"], "gioi_han": lim, "trang_so_do": TRANG_SO_DO,
            "ket_luan_vn": f"ở {st_vn} {'phù hợp' if ok else 'không phù hợp'}",
            "ket_luan_en": f"in {st_en} {'complied' if ok else 'did not comply'}",
            "diem": [{"khu_vuc": room["vn"], "khu_vuc_en": room["en"], "vi_tri": i + 1, "gia_tri": v}
                     for i, v in enumerate(VALUES[(key, state)])]}


def build_dossier(layouts) -> Dossier:
    photos = [{"path": _optimized(os.path.join(PHOTO_DIR, f)), "caption_vn": c_vn, "caption_en": c_en}
              for f, c_vn, c_en in PHOTOS]
    room_name = lambda k: ANTE_VN if k == "ante" else ROOMS[k]["vn"]          # noqa: E731
    room_en = lambda k: ANTE_EN if k == "ante" else ROOMS[k]["en"]            # noqa: E731
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=DATE1, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {DATE1} từ {START1} đến {END1}; ngày {DATE2} từ {START2} đến {END2}",
        sample_time_en=f"From {DATE1}",
        inspection_dates=[DATE1, DATE2],      # Table 0 mục 4: "Từ ngày 29/10/2026 và 04/11/2026" (AGENTS mục 2)
        sub_no=None, is_multi_cert=False, dossier_type="cleanroom",
        scope_vn="Giám định chất lượng phòng sạch", scope_en="Cleanroom performance inspection",
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01", procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01",
        standards_vn="ISO 14644-1:2015", standards_en="ISO 14644-1:2015",
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),          # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam
        equipments=[   # chỉ thiết bị thực sự dùng (phiếu ghi chép tick Lasair III 110; AeroTrak 9350 không dùng)
            eq("lasair_iii_110"),
            eq("testo_625"),
            eq("rion_nl53"),
            eq("tsi_8375"),
        ],
        locations=[MonitoringLocation(
            name_vn=ROOMS[k]["vn"], name_en=ROOMS[k]["en"], area_m2=ROOMS[k]["area"], clean_class=ROOMS[k]["cap"],
            state_vn="Nghỉ và hoạt động", state_en="At rest and operational",
            points_summary_vn=points_summary(k, False), points_summary_en=points_summary(k, True),
        ) for k in ROOMS],
        photos=photos, layouts=layouts["vn"],
        results_data={
            "photo_width_cm": 7.0,   # ảnh 2 cột 7,0 cm (khung 7–8 cm) để TBKQ EN vừa trang như VN (07/11/2026)
            "layout_width_cm": 11.0,  # sơ đồ 11 cm (13 -> 12 -> 11: ghi chú 11pt 09/11) để 9.2.1 lên cùng trang sơ đồ, không trống nửa trang
            "bui_6kenh": [dust_group("iso5", "nghi"), dust_group("iso5", "hd")],
            "bui_3kenh": [dust_group("iso7", "nghi"), dust_group("iso7", "hd")],
            "nhiet_am": {"dong": [{"khu_vuc": ROOMS[k]["vn"], "khu_vuc_en": ROOMS[k]["en"], "nhiet_do": vn(T_AVG[k]),
                                   "do_am": vn(RH_AVG[k]), "do_on": vn(NOISE[k])} for k in ROOMS],
                         "muc_nhiet_do": "-", "muc_do_am": "-", "muc_do_on": "-",
                         "ket_luan_vn": "-", "ket_luan_en": "-"},
            "chenh_ap": {"dong": [{"khu_vuc_1": room_name(a), "khu_vuc_1_en": room_en(a),
                                   "khu_vuc_2": room_name(b), "khu_vuc_2_en": room_en(b),
                                   "gia_tri": vn(v), "muc": "-"} for (a, b, _, _), v in zip(DP, DP_AVG)],
                         "ket_luan_vn": "-", "ket_luan_en": "-"},
        },
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX (cùng cách build_138143)
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
        if next(el.iter(qn("w:drawing")), None) is not None:
            continue          # bảng ảnh/sơ đồ: cho ngắt giữa các hàng ảnh (giữ cả khối -> trang trắng nửa trang)
        prev = el.getprevious()
        if prev is not None and prev.tag == qn("w:p") and _ptext(prev).strip():
            _keep(prev)
            rows = el.findall(qn("w:tr"))
            # tiêu đề + 2 dòng tiêu đề bảng + 4 dòng đầu liền nhau; bảng ngắn (≤ 10 hàng) giữ trọn — tránh chẻ bảng
            # bụi ISO 5 (8 hàng) mà VN/EN ngắt khác chỗ (05/11/2026, sau khi bảng lên 13pt)
            n = len(rows) - 1 if len(rows) <= 10 else min(5, len(rows) - 1)
            for tr in rows[:n]:
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
    """(n) đầu dòng bảng Ghi chú và '(3)' trong ô giới hạn bụi -> chỉ số trên."""
    from docx.text.run import Run
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        note = _is_note_table(tbl)
        for r_el in list(tbl.iter(qn("w:r"))):
            ts = r_el.findall(qn("w:t"))
            if len(ts) != 1:
                continue
            txt = ts[0].text or ""
            m = re.match(r"^\s*(\(\d\))(.*)$", txt) if note else re.match(r"^(.*?)\s*(\(3\))$", txt)
            if not m:
                continue
            run = Run(r_el, None)
            if run.font.superscript:
                continue
            if note:
                mark, rest = m.group(1), m.group(2)
                if rest:
                    el = copy.deepcopy(r_el)
                    r_el.addnext(el)
                    Run(el, None).text = rest
                run.text = mark
                run.font.superscript = True
            else:
                before, mark = m.group(1), m.group(2)
                el = copy.deepcopy(r_el)
                r_el.addnext(el)
                run.text = before + " "
                Run(el, None).text = mark
                Run(el, None).font.superscript = True


def _drop_section_gaps(doc):
    """Đoạn trống giữa câu Đánh giá / bảng và tiêu đề mục kế -> bỏ, tiêu đề mục sau space_before 6pt (như 138143)."""
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
        if prev is None or nxt is None or nxt.tag != qn("w:p") or not _ptext(nxt).strip():
            continue
        if prev.tag == qn("w:tbl") and not _is_note_table(prev):
            continue                                    # khoảng giữa bảng kết quả và bảng Ghi chú giữ nguyên
        el.getparent().remove(el)
        Paragraph(nxt, None).paragraph_format.space_before = Pt(6)
    for el in body.iterchildren(qn("w:p")):
        prev = el.getprevious()
        if prev is None or prev.tag != qn("w:p") or not _ptext(prev).strip().startswith(("Đánh giá", "Assessment")):
            continue
        pf = Paragraph(el, None).paragraph_format
        pf.keep_with_next = False
        pf.keep_together = True


def _drop_trailing_empty(doc):
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


def _remove_no_limit_parts(doc, is_en):
    """Nhiệt độ - độ ẩm - độ ồn và chênh áp: khách hàng không cung cấp mức -> bỏ dòng/cột mức, ghi chú (3), Đánh giá
    (AGENTS mục 3)."""
    body = doc.element.body
    heads = ("Measurement result of temperature", "Measurement result of pressure") if is_en else \
        ("Kết quả đo nhiệt độ", "Kết quả đo chênh lệch áp suất")
    for head in heads:
        hp = next((p for p in body.iterchildren(qn("w:p")) if _ptext(p).strip().startswith(head)), None)
        if hp is None:
            print(f"[WARNING] không thấy mục '{head}'")
            continue
        tbl = hp.getnext()
        while tbl is not None and tbl.tag != qn("w:tbl"):
            tbl = tbl.getnext()
        from docx.table import Table
        t = Table(tbl, doc)
        if "nhiệt" in head or "temperature" in head:
            last = t.rows[-1]._tr
            if re.search(r"Mức quy định|limit|Specif", _ptext(last)):
                last.getparent().remove(last)
        else:
            _drop_last_column(t)
        note = tbl.getnext()
        while note is not None and not _is_note_table(note):
            note = note.getnext()
        for tr in note.findall(qn("w:tr")):
            if "(3)" in _ptext(tr):
                tr.getparent().remove(tr)
        a = note.getnext()
        if a is not None and a.tag == qn("w:p") and _ptext(a).strip().startswith(("Đánh giá", "Assessment")):
            s = a.getnext()
            a.getparent().remove(a)
            if s is not None and s.tag == qn("w:p") and _ptext(s).strip():
                s.getparent().remove(s)
        else:
            print(f"[WARNING] '{head}': không thấy đoạn Đánh giá để bỏ")


def _drop_last_column(t):
    """Bỏ cột cuối (Mức quy định) của bảng chênh áp; độ rộng cột bỏ chia đều cho 2 cột tên khu vực."""
    tbl = t._tbl
    grid = tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    w_last = int(cols[-1].get(qn("w:w")))
    grid.remove(cols[-1])
    for i in (0, 2):
        cols[i].set(qn("w:w"), str(int(cols[i].get(qn("w:w"))) + w_last // 2))
    for tr in tbl.findall(qn("w:tr")):
        tcs = tr.findall(qn("w:tc"))
        last = tcs[-1]
        span = last.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
        tr.remove(last)
        for i in (0, 2):
            if i < len(tcs) - 1:
                w = tcs[i].find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
                if w is not None and w.get(qn("w:type")) == "dxa":
                    w.set(qn("w:w"), str(int(w.get(qn("w:w"))) + w_last // 2))
        if span is not None:
            print("[WARNING] cột mức chênh áp có gridSpan — kiểm tra bảng")


def _restructure_dust(doc, is_en):
    """Mục bụi: 9.2 'Kết quả đo nồng độ hạt bụi' -> 9.2.1 trạng thái nghỉ (ISO 5, ISO 7), 9.2.2 hoạt động (ISO 5, ISO 7)
    (người dùng yêu cầu 04/11/2026). Template render ra: [tiêu đề 3 kênh][ISO 7 nghỉ][ISO 7 HĐ][tiêu đề 6 kênh]
    [ISO 5 nghỉ][ISO 5 HĐ] -> sắp lại."""
    body = doc.element.body
    paras = list(body.iterchildren(qn("w:p")))
    k3 = "(3 particle sizes)" if is_en else "(3 kênh hạt)"
    k6 = "(6 particle sizes)" if is_en else "(6 kênh hạt)"
    h3 = next(p for p in paras if _ptext(p).strip().endswith(k3))
    h6 = next(p for p in paras if _ptext(p).strip().endswith(k6))
    nxt_head = next(p for p in paras if _ptext(p).strip().startswith(
        "Measurement result of temperature" if is_en else "Kết quả đo nhiệt độ"))

    def groups(start, stop):
        els, el = [], start.getnext()
        while el is not None and el is not stop:
            els.append(el)
            el = el.getnext()
        out, cur = [], None
        for e in els:
            if e.tag == qn("w:tbl") and not _is_note_table(e):
                cur = [e]
                out.append(cur)
            elif cur is not None:
                cur.append(e)
        if len(out) != 2:
            raise SystemExit(f"[FATAL] mục bụi: cần 2 nhóm (nghỉ, hoạt động), thấy {len(out)}")
        return out

    iso7 = groups(h3, h6)
    iso5 = groups(h6, nxt_head)

    def set_text(p_el, text):
        rs = p_el.findall(qn("w:r"))
        ts = rs[0].findall(qn("w:t"))
        ts[0].text = text
        for t in ts[1:]:
            t.getparent().remove(t)
        for r in rs[1:]:
            r.getparent().remove(r)

    def sub(text):
        p = copy.deepcopy(h3)
        lvl = p.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:ilvl')}")
        lvl.set(qn("w:val"), "2")
        set_text(p, text)
        return p

    if is_en:
        set_text(h3, "Measurement result of particle concentration")
        s1 = sub("Measurement result of particle concentration in at-rest state")
        s2 = sub("Measurement result of particle concentration in operational state")
    else:
        set_text(h3, "Kết quả đo nồng độ hạt bụi")
        s1 = sub("Kết quả đo nồng độ hạt bụi ở trạng thái nghỉ")
        s2 = sub("Kết quả đo nồng độ hạt bụi ở trạng thái hoạt động")
    for p_el in (h3, s1, s2):
        _keep(p_el)
    order = [s1] + iso5[0] + iso7[0] + [s2] + iso5[1] + iso7[1]
    anchor = h3
    for e in order:
        anchor.addnext(e)
        anchor = e
    h6.getparent().remove(h6)
    # Bảng ISO 7 đứng ngay sau câu kết luận ISO 5: bỏ đoạn trống ở giữa (thay bằng space_after 12pt của câu kết luận)
    # và giữ tiêu đề bảng + 4 dòng đầu liền nhau như bảng ISO 5 (đoạn trống làm _keep_headings_with_tables bỏ sót).
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
    for g in (iso7[0], iso5[1], iso7[1]):
        tbl = g[0]
        prev = tbl.getprevious()
        while prev is not None and prev.tag == qn("w:p") and not _ptext(prev).strip() and prev.find(qn("w:r")) is None:
            p0 = prev.getprevious()
            prev.getparent().remove(prev)
            prev = p0
        if prev is not None and prev.tag == qn("w:p") and not _ptext(prev).strip().startswith(("Kết quả đo nồng độ hạt bụi ở",
                                                                                           "Measurement result of particle concentration in")):
            Paragraph(prev, None).paragraph_format.space_after = Pt(12)
        rows = tbl.findall(qn("w:tr"))
        for tr in rows[:min(5, len(rows) - 1)]:
            for p_el in tr.iter(qn("w:p")):
                _keep(p_el)


def _iso5_note(doc, is_en):
    """Bảng ISO 5: ghi chú ⁽³⁾ cho ô giới hạn '-' ở d ≥ 5 µm."""
    txt = ("No limit is specified for particles d ≥ 5 µm for ISO Class 5 (ISO 14644-1:2015, Table 1)." if is_en else
           "ISO 14644-1:2015 (Bảng 1) không quy định giới hạn nồng độ hạt d ≥ 5 µm đối với cấp sạch ISO 5.")
    body = doc.element.body
    n = 0
    for tbl in body.iterchildren(qn("w:tbl")):
        if _is_note_table(tbl) or "(3)" not in _ptext(tbl):
            continue
        note = tbl.getnext()
        while note is not None and not _is_note_table(note):
            note = note.getnext()
        rows = note.findall(qn("w:tr"))
        src = next(tr for tr in rows if "(2)" in _ptext(tr))
        new = copy.deepcopy(src)
        tcs = new.findall(qn("w:tc"))
        for tc, val in ((tcs[0], "(3) …"), (tcs[-1], txt)):
            ts = list(tc.iter(qn("w:t")))
            ts[0].text = val
            for t in ts[1:]:
                t.text = ""
        src.addnext(new)
        n += 1
    if n != 2:
        print(f"[WARNING] ghi chú ⁽³⁾ ISO 5: thêm {n}/2")


def _shrink_one_line_rows(doc):
    """Hàng 1 dòng chữ mà template ép cao tối thiểu > 760 twips (bảng nhiệt ẩm 1361, chênh áp 1247) -> 440 twips
    (~0,78 cm, AGENTS mục 9; người dùng yêu cầu 04/11/2026: hàng dữ liệu không để cao thừa). Hàng nhiều dòng chữ
    (vd dòng giới hạn bụi 4 dòng) giữ nguyên."""
    n = 0
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        if _is_note_table(tbl) or next(tbl.iter(qn("w:drawing")), None) is not None:
            continue
        for tr in tbl.findall(qn("w:tr")):
            h = tr.find(f"{qn('w:trPr')}/{qn('w:trHeight')}")
            if h is None or int(h.get(qn("w:val"), "0")) <= 760:
                continue
            cells = [tc for tc in tr.findall(qn("w:tc"))]
            one_line = all(len([p for p in tc.findall(qn("w:p")) if _ptext(p).strip()]) <= 1
                           and len(_ptext(tc).strip()) <= 22 for tc in cells)
            if one_line:
                h.set(qn("w:val"), "440")
                h.set(qn("w:hRule"), "atLeast")
                n += 1
    print(f"[build_132809] thu {n} hàng 1 dòng về 440 twips")


def finish_tbkq(src: str, is_en: bool) -> str:
    doc = Document(src)
    _shrink_one_line_rows(doc)
    _fix_cover_lists(doc)
    _restructure_dust(doc, is_en)
    _remove_no_limit_parts(doc, is_en)
    _iso5_note(doc, is_en)
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
def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["tham_chieu"] = [
        f"- Kết quả đo nồng độ hạt bụi (trạng thái nghỉ và hoạt động): xem tại phiếu Tính toán kết quả nồng độ hạt bụi "
        f"số {CODE}/{SUFFIX['bui']} ngày {DATE2}.",
        *REF_NHIET_AM_ON,
        f"- Kết quả đo chênh lệch áp suất: xem tại phiếu Kết quả đo chênh lệch áp suất "
        f"số {CODE}/{SUFFIX['chenh_ap']} ngày {DATE1}.",
    ]
    rows = []
    for k in ROOMS:
        cap = ROOMS[k]["cap"]
        dat = [f"Nồng độ hạt bụi ở {STATES[s][0]} ({cap})" for s in STATES if DUST_OK[(k, s)]]
        khong = [f"Nồng độ hạt bụi ở {STATES[s][0]} ({cap})" for s in STATES if not DUST_OK[(k, s)]]
        rows.append({"khu_vuc": f"“{ROOMS[k]['vn']}”", "dat": ";\n".join(dat) or "Không",
                     "khong_dat": ";\n".join(khong) or "Không",
                     "can_cu": f"ISO 14644-1:2015 (cấp sạch {cap})."})
    ctx["danh_gia"] = rows
    out = b.build(ctx=ctx, out_name=f"BCKQ - 1 - phong sach - {CODE}.docx")
    doc = Document(out)
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
REF_NHIET_AM_ON = [
    f"- Kết quả đo nhiệt độ, độ ẩm: xem tại phiếu Tính toán kết quả nhiệt độ, độ ẩm "
    f"số {CODE}/{SUFFIX['nhiet_am']} ngày {DATE1}.",
    f"- Kết quả đo độ ồn: xem tại phiếu Tính toán kết quả độ ồn số {CODE}/{SUFFIX['do_on']} ngày {DATE1}.",
]


def build_excel() -> str:
    """Phiếu tính toán theo form phòng sạch chung (Core/Tools/cleanroom_calc.py, người dùng chốt 07/11/2026)."""
    book = CleanroomCalcBook()
    recs = read_lasair_csv(DUST_CSV)          # Data ID = số thứ tự dòng (CSV không có cột Data ID)
    groups = []
    for key in ROOMS:
        lim = {k: v for k, v in zip(REPORT[key], LIMITS[key])}
        for s_ in STATES:
            groups.append({"room": ROOMS[key]["vn"], "cap": ROOMS[key]["cap"], "state": STATES[s_][2],
                           "limits": lim, "zero": recs[ZERO_FILE] if not groups else None,
                           "points": [recs[f] for f in FILES[(key, s_)]]})
    book.dust_lasair(groups, so=f"{CODE}/{SUFFIX['bui']}", date=DATE2)
    book.temp_rh([{"room": ROOMS[k]["vn"], "loc": 1, "t": TEMP_READ[k], "rh": RH_READ[k], "lim_t": None,
                   "lim_rh": None} for k in ROOMS], CAL_T_TBL, CAL_RH_TBL,
                 so=f"{CODE}/{SUFFIX['nhiet_am']}", date=DATE1)
    book.noise([{"room": ROOMS[k]["vn"], "loc": 1, "reads": [NOISE[k]], "lim": None} for k in ROOMS],
               so=f"{CODE}/{SUFFIX['do_on']}", date=DATE1)
    names = {**{k: ROOMS[k]["vn"] for k in ROOMS}, "ante": ANTE_VN}
    book.pressure([{"a": names[a], "b": names[b_], "sign": sign, "reads": reads, "lim": None}
                   for a, b_, sign, reads in DP], so=f"{CODE}/{SUFFIX['chenh_ap']}", date=DATE1, cal=CAL_DP_TBL)
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - TinhToanPhongSach.xlsx")
    book.save(out)
    print(f"[build_132809] Excel -> {out}")
    return out


def patch_bckq_refs() -> str:
    """Chỉ dựng lại phiếu: sửa dòng viện dẫn nhiệt ẩm - ồn trong BCKQ đã có (giữ ngày, định dạng)."""
    path = os.path.join(INPUT_DIR, "ThucHien", f"BCKQ - 1 - phong sach - {CODE}.docx")
    ok = replace_reference_lines(path, "- Kết quả đo nhiệt độ, độ ẩm, độ ồn", REF_NHIET_AM_ON)
    print(f"[build_132809] BCKQ viện dẫn nhiệt ẩm / ồn: {'đã sửa' if ok else 'không thấy dòng cũ (giữ nguyên)'}")
    return path


def split_edoc() -> list:
    """BBLM = scan tr. 15–16; Ghi chép = tr. 19–21 (BM3030.01) + tr. 17 (mặt bằng GĐV đánh dấu vị trí) -> TBKQ/Edoc."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(SCAN)
    outs = []
    for name, pages in ((f"{CODE6} - BBLM.pdf", [14, 15]), (f"{CODE6} - Ghi Chep.pdf", [18, 19, 20, 16])):
        dst = fitz.open()
        for pg in pages:
            dst.insert_pdf(src, from_page=pg, to_page=pg)
        p = os.path.join(edoc, name)
        finish_edoc(dst, INPUT_DIR, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_132809] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--tao-lai-so-do", action="store_true",
                    help="Dựng lại so_do_vi_tri.drawio từ tọa độ trong code (bỏ chỉnh tay trên draw.io)")
    ap.add_argument("--chi-so-do", action="store_true", help="Chỉ dựng/xuất sơ đồ, không dựng DOCX/XLSX")
    ap.add_argument("--chi-excel", action="store_true",
                    help="Chỉ dựng lại phiếu tính toán (+ sửa dòng viện dẫn BCKQ, + PDF nếu --pdf); giữ TBKQ, ngày ban hành đã giao")
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

    check_user_excel()
    for key, rows in VALUES.items():
        print(f"[build_132809] {key}: {rows} -> {'đạt' if DUST_OK[key] else 'KHÔNG ĐẠT'}")
    print(f"[build_132809] T {T_AVG}, RH {RH_AVG} (ngoại suy: {RH_EXTRAPOLATED}), ồn {NOISE}, ΔP {DP_AVG}")
    layouts = build_layouts(args.tao_lai_so_do)
    d = build_dossier(layouts)
    d.ensure_output_directories()
    split_edoc()
    h = CleanroomHandler(d)
    vn_doc = finish_tbkq(h.build_tbkq(is_en=False), is_en=False)
    d.layouts = layouts["en"]
    en_doc = finish_tbkq(h.build_tbkq(is_en=True), is_en=True)
    d.layouts = layouts["vn"]
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
