# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D120423 - Công ty TNHH Kỹ thuật Nhật Khang (BV Đa khoa Tâm Anh TP.HCM) — ĐỢT 8 (16/11/2026),
chứng thư Q326D120423/8 (đợt 7 = /7, script build_120423.py). Phòng sạch Khu nội trú C2: Khu ICU ngoại + VIP 3 +
P. Đệm + sảnh / hành lang ICU (tầng 6), Phòng mổ 10 (tầng trệt), Phòng mổ tầng hầm: bụi + vi sinh không khí.

Nguồn (Input/120423/Dot8, chép từ Z "…\\Q326D120423-Nhật Khang - Tâm Anh - HCM\\Q326D120423 - Đợt 8 - 161126"
17/11/2026 — chỉ đọc):
  - 20261117170142289.pdf (scan 18 trang):
      tr.1–2 BBLM (BM3030.02) ngày 16/11/2026: Nhật Khang, 49/24/7A Nguyễn Hữu Cảnh, phường Thạnh Mỹ Tây;
        địa điểm "Khu ICU Ngoại Tầng 6 C2, Phòng mổ 10 Trệt C2, Phòng mổ Hầm C2 - Bệnh viện Đa khoa Tâm Anh Thành
        phố Hồ Chí Minh, số 2B Phổ Quang, phường Tân Sơn Hòa"; 9h00 – 14h00 ngày 16/11/2026; AeroTrak 9350 + Buck
        B30120; QUATEST 3 3030:2025; N2.TC.PS.01; N2.TC.09; vi sinh 19 mẫu (1 → 19), bụi 51 vị trí; "Khu ICU2, P.Vip3,
        Hành lang, P.Đệm, Phòng mổ 10 Trệt C2, Phòng mổ Hầm C2"; người thực hiện Nguyễn Minh Quân (chữ ký);
        ký hiệu phiếu ghi chép Q326D120423/8.
      tr.3 phiếu ghi chép chung (BM3030.01); tr.4–7 bụi AeroTrak: ICU ngoại zero 579, 1 → 13 = 580 → 592 (11–13 sửa
        từ 560–562); P.Vip3 14 → 19 = 593 → 598; P.Đệm zero 599, 20 → 23 = 600 → 603; hành lang 24 → 33 = 604 → 613;
        tầng trệt P.Mổ 10 "zero 615", 34 → 42 = 616 → 624; P.Mổ hầm C2 43 → 52 = 627 → 636.
      tr.8–10 vi sinh (không ghi lưu lượng / thời gian / blank): P.Vip3 1–2; Khu ICU ngoại (Khu ICU2) 3–7;
        P.Đệm + hành lang ICU 8–12 (sơ đồ: Đệm 8, hành lang 9–12); P.Mổ 10 tầng trệt 13–15; P.Mổ hầm C2 16–19.
      tr.11 mặt bằng PM10 đánh dấu vi sinh 13–15; tr.13 mặt bằng tầng 6 + hầm đánh dấu vi sinh; tr.15 mặt bằng PM10
        đánh dấu bụi 34–42; tr.17 mặt bằng tầng 6 + hầm đánh dấu bụi 1–51 (hầm 43–51). tr.12, 14, 16, 18 mặt sau
        (chỉ chữ thấm) — bỏ.
  - KQ Do/6ChannelFixed.xlsx: AeroTrak, Σ, hạt/m³, file 185 – 644 (dùng 579 – 635; 614 = 0 hạt).
  - Không có thư mục ảnh trong Đợt 8.
  - Mặt bằng: Input/120423/Layout/"Mặt bằng Màng Hepa - Nội trú C2, C3.pdf" tr.2 (tầng 6 + khung tầng hầm C2),
    "Mặt bằng Màng Lọc Hepa - Phòng mổ số 10.pdf" (chép từ thư mục hồ sơ trên Z).
  - Diện tích, cấp sạch: "Bảng khối lượng quan trắc chất lượng phòng sạch – Khu nội trú C2" (scan đợt 7 tr.3):
    ICU Ngoại 64,9 m²; P. VIP 3 (cạnh ICU Lây) 12,8; P. Đệm (ICU Lây) 7,3; Sảnh / hành lang ICU 38,0 (ước tính);
    Phòng mổ 10 33,0 (ước tính); Phòng mổ (tầng hầm) 35,0 — ICU ISO 8, 2 phòng mổ ISO 7.

Quyết định của người dùng (17/11/2026):
  - 2 phòng mổ đánh giá theo "ISO 8" (bảng khối lượng ghi ISO 7 — theo ISO 7 thì bụi không đạt); vi sinh 2 phòng mổ
    "GMP Grade D ≤ 200" -> toàn bộ hồ sơ ISO 8, GMP Grade D như đợt 7.
  - Phòng mổ tầng hầm: "Bỏ vị trí 52 (file 636)" -> 9 vị trí 43–51 (= BBLM 51, bảng khối lượng 9).
  - Phòng mổ 10: "Dùng 615–623" (file 614 = zero theo dữ liệu máy; phiếu ghi "zero 615", 616–624).
  - Sơ đồ: vị trí căn đều theo vách (lưới), vi sinh hình thoi xanh lá, ký hiệu đè chữ -> nhãn dẫn mũi tên (17/11/2026).

Suy luận của AI (cần người dùng xác nhận):
  (1) Chỉ dựng TBKQ tiếng Việt như đợt 5–7.
  (2) Vi sinh 100 L/phút × 01 phút như đợt 7 (phiếu đợt 8 để trống) -> LOD < 10 CFU/m³; không có mẫu blank.
  (3) Phòng mổ tầng hầm không ghi zero -> dùng zero 614 (lần zero gần nhất trước file 627); file 625, 626 không có.
  (4) Tên khu vực theo bảng khối lượng; P. Đệm / hành lang chia theo sơ đồ đánh dấu tay.
  (5) Số phiếu tính toán Q326D120423/8/Bụi, /8/Vi sinh; ngày ban hành = ngày KQTN (chờ).
  (6) Không có ảnh -> TBKQ / BCKQ không có mục ảnh.

Vi sinh CHƯA có KQTN: điền VI_SINH_CFU_DIA (ký hiệu mẫu 1..19), KQTN_SO, KQTN_NGAY rồi chạy lại.

Chạy:  python build_120423_dot8.py [--pdf] [--tao-lai-so-do] [--chi-so-do]
Chép Z: python Core/Tools/ai_files.py chep 120423/Dot8 --z "<thư mục Đợt 8 trên Z>" --ghi
"""

import os
import re
import sys
import copy
import argparse

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_dong_bo_Z"))

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, within, verdict, set_assessment_paragraph, style_note_tables
from Core.Tools.qa_check import check_single_docx
from Core.Tools.cleanroom_calc import CleanroomCalcBook, read_aerotrak, LEFT
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D120423"
CODE6 = "120423"
SUB_NO = 8                       # đợt 8 -> chứng thư /8
REG_DATE = "25/09/2026"          # BBLM mục 1
SAMPLE_DATE = "16/11/2026"
START, END = "09:00", "14:00"    # BBLM
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6, "Dot8")
LAYOUT_DIR = os.path.join(WORKSPACE, "Input", CODE6, "Layout")
SCAN = os.path.join(INPUT_DIR, "20261117170142289.pdf")
RAW_XLSX = os.path.join(INPUT_DIR, "KQ Do", "6ChannelFixed.xlsx")
PHOTO_DIR = None                 # đợt 8 không có ảnh — suy luận (6)
PLAN_C2 = os.path.join(LAYOUT_DIR, "Mặt bằng Màng Hepa - Nội trú C2, C3.pdf")
PLAN_PM10 = os.path.join(LAYOUT_DIR, "Mặt bằng Màng Lọc Hepa - Phòng mổ số 10.pdf")
DIAGRAM_DIR = os.path.join(INPUT_DIR, "TBKQ", "_diagrams")

# --- Vi sinh: chưa có KQTN. Điền khi có kết quả rồi chạy lại ---
KQTN_SO = ""
KQTN_NGAY = ""
VI_SINH_CFU_DIA = {i: None for i in range(1, 20)}   # CFU/đĩa theo ký hiệu mẫu 1..19; 0 = "< 1"
VI_SINH_BLANKS = {}                                  # phiếu không ghi mẫu blank
VI_SINH_KQ = {}
VS_FLOW_LPM, VS_MINUTES = 100, 1      # suy luận (2)
BLANK_LIST = ", ".join(str(s) for s in VI_SINH_BLANKS)
N_VS = len(VI_SINH_CFU_DIA)

CERT_DATE = KQTN_NGAY or PLACEHOLDER

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Kỹ thuật Nhật Khang",
    name_en="Nhat Khang Technical Company Limited",     # không dùng (chỉ TBKQ VN) — AI dịch
    address_vn="49/24/7A Nguyễn Hữu Cảnh, phường Thạnh Mỹ Tây, Thành phố Hồ Chí Minh",
    address_en="49/24/7A Nguyen Huu Canh Street, Thanh My Tay Ward, Ho Chi Minh City",
)
LOCATION_VN = ("Khu ICU Ngoại tầng 6 C2, Phòng mổ 10 tầng trệt C2, Phòng mổ tầng hầm C2 - Bệnh viện Đa khoa Tâm Anh "
               "Thành phố Hồ Chí Minh - Số 2B Phổ Quang, phường Tân Sơn Hòa, Thành phố Hồ Chí Minh")
LOCATION_EN = ("External ICU area 6th floor C2, Operating room 10 ground floor C2, Basement operating room C2 - Tam Anh "
               "General Hospital Ho Chi Minh City - No. 2B Pho Quang Street, Tan Son Hoa Ward, Ho Chi Minh City")

ISO_CLASS = 8
ISO8_LIMITS = (3520000, 832000, 29300)          # ISO 14644-1:2015 Bảng 1, d ≥ 0,5 / 1 / 5 µm, hạt/m³
GRADE = "D"
VS_LIMIT = 200                                  # GMP Grade D, CFU/m³
BASIS_VN = "WHO TRS 1044 (2022) – Phụ lục 2"
BASIS_EN = "WHO TRS 1044 (2022) – Annex 2"


def GMP(g: str) -> str:
    return f"GMP Grade {g}"


# ---------------------------------------------------------------------------
# Mặt bằng ghép: tầng 6 (trên) + tầng hầm và Phòng mổ 10 (dưới, thu 0,9)
# Ảnh nguồn @300 dpi: C2 tr.2 xoay 270° (3508 × 4961), PM10 tr.1 (3508 × 4961).
# Scan đánh dấu tay @200 dpi ghép SIFT + RANSAC (similarity) — tr.17/13 xoay 180° rồi M; tr.15/11 dùng M trực tiếp.
# ---------------------------------------------------------------------------
PLAN_DPI = 300
T6_BOX = (0, 2950, 2700, 4560)          # tầng 6 trên ảnh C2 xoay
HAM_BOX = (300, 430, 1900, 1450)        # khung tầng hầm trên ảnh C2 xoay
PM10_BOX = (950, 1950, 2300, 2950)      # Phòng mổ 10 trên ảnh PM10
BOT_S = 0.9
BOT_Y = (T6_BOX[3] - T6_BOX[1]) + 20
PM10_X = round((HAM_BOX[2] - HAM_BOX[0]) * BOT_S) + 30
PLAN_W = T6_BOX[2] - T6_BOX[0]
PLAN_H = BOT_Y + round(max(HAM_BOX[3] - HAM_BOX[1], PM10_BOX[3] - PM10_BOX[1]) * BOT_S)

M_SCAN = {      # trang scan -> (xoay 180°?, ma trận 2×3 scan@200 -> ảnh nguồn @300)
    17: (True, ((2.20794, 0.0162058, -87.9181), (-0.0162058, 2.20794, -63.754))),     # bụi tầng 6 + hầm
    13: (True, ((2.20617, 0.0128656, -75.8306), (-0.0128656, 2.20617, -53.9424))),    # vi sinh tầng 6 + hầm
    15: (False, ((-2.20870, -0.00851691, 3592.26), (0.00851691, -2.20870, 5085.13))),  # bụi PM10
    11: (False, ((-2.20966, -0.0161429, 3604.42), (0.0161429, -2.20966, 5082.53))),    # vi sinh PM10
}


def src_to_plan(part: str, x: float, y: float):
    """Tọa độ ảnh nguồn (@300 dpi) của phần 't6' | 'ham' | 'pm10' -> tọa độ mặt bằng ghép."""
    if part == "t6":
        return x - T6_BOX[0], y - T6_BOX[1]
    if part == "ham":
        return (x - HAM_BOX[0]) * BOT_S, (y - HAM_BOX[1]) * BOT_S + BOT_Y
    return (x - PM10_BOX[0]) * BOT_S + PM10_X, (y - PM10_BOX[1]) * BOT_S + BOT_Y


def scan_to_plan(page: int, part: str, x: float, y: float):
    rot, M = M_SCAN[page]
    if rot:
        x, y = 1654 - x, 2339 - y
    sx = M[0][0] * x + M[0][1] * y + M[0][2]
    sy = M[1][0] * x + M[1][1] * y + M[1][2]
    return src_to_plan(part, sx, sy)


def _page_img(pdf, page, rot_k):
    import fitz
    import numpy as np
    from PIL import Image
    pm = fitz.open(pdf)[page].get_pixmap(dpi=PLAN_DPI)
    a = np.frombuffer(pm.samples, np.uint8).reshape(pm.height, pm.width, pm.n)[..., :3]
    return Image.fromarray(np.rot90(a, rot_k).copy() if rot_k else a.copy())


def _make_plan():
    from PIL import Image, ImageDraw
    c2 = _page_img(PLAN_C2, 1, 3)
    pm10 = _page_img(PLAN_PM10, 0, 0)
    out = Image.new("RGB", (PLAN_W, PLAN_H), "white")
    out.paste(c2.crop(T6_BOX), (0, 0))
    ham = c2.crop(HAM_BOX)
    out.paste(ham.resize((round(ham.width * BOT_S), round(ham.height * BOT_S))), (0, BOT_Y))
    pm = pm10.crop(PM10_BOX)
    out.paste(pm.resize((round(pm.width * BOT_S), round(pm.height * BOT_S))), (PM10_X, BOT_Y))
    d = ImageDraw.Draw(out)                     # đường phân cách 3 mặt bằng
    d.line([(0, BOT_Y - 10), (PLAN_W, BOT_Y - 10)], fill="black", width=4)
    d.line([(PM10_X - 15, BOT_Y - 10), (PM10_X - 15, PLAN_H)], fill="black", width=4)
    return out


# Phòng theo bảng khối lượng. dust: [(nhãn, file AeroTrak, (trang scan, phần, x, y) số viết tay @200 dpi)]
# vs: [(nhãn, ký hiệu mẫu, (trang, phần, x, y))]. Nhãn đánh số liên tục như phiếu ghi chép / sơ đồ viết tay.
def _pts(page, part, items):
    return [(n, f, (page, part, x, y)) for n, f, (x, y) in items]


ROOMS = [
    {"vn": "Khu ICU Ngoại", "en": "External ICU area", "area": "64,9", "zero": 579,
     "dust": _pts(17, "t6", [(1, 580, (1491, 558)), (2, 581, (1416, 489)), (3, 582, (1359, 589)),
                             (4, 583, (1246, 564)), (5, 584, (1265, 464)), (6, 585, (1164, 451)),
                             (7, 586, (1146, 571)), (8, 587, (1030, 520)), (9, 588, (1040, 430)),
                             (10, 589, (915, 405)), (11, 590, (915, 505)), (12, 591, (785, 490)),
                             (13, 592, (790, 385))]),
     "vs": _pts(13, "t6", [(3, 3, (1412, 537)), (4, 4, (1290, 530)), (5, 5, (1150, 520)), (6, 6, (1015, 480)),
                           (7, 7, (870, 470))])},
    {"vn": "P. VIP 3 (cạnh ICU Lây)", "en": "VIP room 3 (next to infectious ICU)", "area": "12,8", "zero": 579,
     "dust": _pts(17, "t6", [(14, 593, (615, 335)), (15, 594, (715, 345)), (16, 595, (700, 430)),
                             (17, 596, (590, 405)), (18, 597, (580, 480)), (19, 598, (680, 495))]),
     "vs": _pts(13, "t6", [(1, 1, (670, 350)), (2, 2, (640, 480))])},
    {"vn": "P. Đệm (ICU Lây)", "en": "Anteroom (infectious ICU)", "area": "7,3", "zero": 599,
     "dust": _pts(17, "t6", [(20, 600, (560, 610)), (21, 601, (605, 570)), (22, 602, (585, 635)),
                             (23, 603, (650, 600))]),
     "vs": _pts(13, "t6", [(8, 8, (565, 605))])},
    {"vn": "Sảnh / hành lang ICU", "en": "ICU lobby / corridor", "area": "38,0", "zero": 599,
     "dust": _pts(17, "t6", [(24, 604, (740, 640)), (25, 605, (785, 595)), (26, 606, (845, 645)),
                             (27, 607, (895, 625)), (28, 608, (935, 650)), (29, 609, (975, 630)),
                             (30, 610, (980, 745)), (31, 611, (945, 780)), (32, 612, (960, 820)),
                             (33, 613, (905, 840))]),
     "vs": _pts(13, "t6", [(9, 9, (925, 840)), (10, 10, (965, 710)), (11, 11, (895, 625)), (12, 12, (765, 600))])},
    {"vn": "Phòng mổ 10 (tầng trệt)", "en": "Operating room 10 (ground floor)", "area": "33,0", "zero": 614,
     "dust": _pts(15, "pm10", [(34, 615, (1105, 1110)), (35, 616, (900, 1115)), (36, 617, (730, 1110)),
                               (37, 618, (725, 1210)), (38, 619, (890, 1210)), (39, 620, (1090, 1195)),
                               (40, 621, (1085, 1320)), (41, 622, (910, 1320)), (42, 623, (730, 1320))]),
     "vs": _pts(11, "pm10", [(13, 13, (1080, 1195)), (14, 14, (895, 1190)), (15, 15, (720, 1190))])},
    {"vn": "Phòng mổ (tầng hầm)", "en": "Operating room (basement)", "area": "35,0", "zero": 614,
     "dust": _pts(17, "ham", [(43, 627, (940, 1920)), (44, 628, (940, 1750)), (45, 629, (1005, 1880)),
                              (46, 630, (1050, 1945)), (47, 631, (1080, 1785)), (48, 632, (1215, 1800)),
                              (49, 633, (1350, 1815)), (50, 634, (1290, 1990)), (51, 635, (1180, 1960))]),
     "vs": _pts(13, "ham", [(16, 16, (985, 1930)), (17, 17, (1020, 1780)), (18, 18, (1250, 1815)),
                            (19, 19, (1235, 1980))])},
]
TRANG_SO_DO_BUI, TRANG_SO_DO_VS = "2", "3"   # trang phụ lục chứa sơ đồ 1 / sơ đồ 2 — kiểm tra lại trên PDF

LABELS, RECS = read_aerotrak(RAW_XLSX)
_need = [f for r in ROOMS for f in [r["zero"]] + [x[1] for x in r["dust"]]]
if [f for f in _need if f not in RECS]:
    raise SystemExit(f"[FATAL] 6ChannelFixed.xlsx thiếu file {[f for f in _need if f not in RECS]}")
for r in ROOMS:
    for _, f, _ in r["dust"]:
        if RECS[f].get("Date/Time (dd/mm/yyyy)", "")[:10] != SAMPLE_DATE:
            raise SystemExit(f"[FATAL] file {f} không đo ngày {SAMPLE_DATE}")


def dust_vals(f):
    return [RECS[f]["Ch2"], RECS[f]["Ch3"], RECS[f]["Ch6"]]


def dust_ok():
    return all(v <= lim for r in ROOMS for _, f, _ in r["dust"] for v, lim in zip(dust_vals(f), ISO8_LIMITS))


def vs_value(sample):
    """Kết quả CFU/m³ của một ký hiệu mẫu; chưa có -> PLACEHOLDER."""
    vol_m3 = VS_FLOW_LPM * VS_MINUTES / 1000.0
    kq = VI_SINH_KQ.get(sample)
    if kq not in (None, ""):
        return kq
    cfu = VI_SINH_CFU_DIA.get(sample)
    if cfu is None:
        return PLACEHOLDER
    if cfu == 0:
        return f"< {round(1 / vol_m3)}"
    return round(cfu / vol_m3)


def vs_ok():
    vals = [vs_value(s) for r in ROOMS for _, s, _ in r["vs"]]
    if any(v == PLACEHOLDER for v in vals):
        return None
    return all(within(str(v).replace(".", ","), f"≤ {VS_LIMIT}") for v in vals)


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


# ---------------------------------------------------------------------------
# Vách trong từng phòng (tọa độ ảnh nguồn @300 dpi, đổi sang mặt bằng ghép) — vị trí căn đều theo lưới
# (Core/Tools/even_grid.py, người dùng yêu cầu 17/11/2026). Cánh ICU ngoại nghiêng ~10° (vách trên
# (1250,3750) -> (2021,3886)); tầng hầm nghiêng ~9°; Phòng mổ 10 thẳng.
# Sảnh / hành lang ICU hình chữ L -> 2 tứ giác: A sảnh thang máy (vị trí 30–33, vi sinh 9–10), B hành lang ngang
# (24–29, vi sinh 11–12); nhãn chia theo tứ giác chứa chỗ đánh dấu tay.
# ---------------------------------------------------------------------------
def _q(part, *pts):
    return tuple(src_to_plan(part, x, y) for x, y in pts)


def _room_quads():
    return {
        "Khu ICU Ngoại": [_q("t6", (145, 3555), (2010, 3884), (1945, 4405), (88, 4082))],
        "P. VIP 3 (cạnh ICU Lây)": [_q("t6", (2025, 3893), (2379, 3950), (2300, 4464), (1960, 4436))],
        "P. Đệm (ICU Lây)": [_q("t6", (2043, 3589), (2436, 3639), (2405, 3925), (2035, 3868))],
        "Sảnh / hành lang ICU": [_q("t6", (1360, 3100), (1680, 3140), (1618, 3515), (1345, 3480)),
                                 _q("t6", (1345, 3480), (2030, 3588), (2022, 3880), (1330, 3764))],
        "Phòng mổ 10 (tầng trệt)": [_q("pm10", (1015, 2050), (2223, 2050), (2223, 2835), (1015, 2835))],
        "Phòng mổ (tầng hầm)": [_q("ham", (463, 525), (1790, 745), (1680, 1385), (355, 1160))],
    }


def _orig(item):
    page, part, x, y = item[-1]
    return scan_to_plan(page, part, x, y)


def snapped_positions() -> dict:
    """{(phòng, 'dust'|'vs', nhãn): ((x, y), tứ giác)}: mỗi loại một lưới căn đều theo vách; phòng nhiều tứ giác thì
    mỗi nhãn vào tứ giác chứa chỗ đánh dấu tay (gần nhất)."""
    from Core.Tools.even_grid import snap_to_grid, _inside
    quads, pos = _room_quads(), {}
    for r in ROOMS:
        qs = quads[r["vn"]]
        for key in ("dust", "vs"):
            groups = {i: [] for i in range(len(qs))}
            for item in r[key]:
                o = _orig(item)
                k = next((i for i, q in enumerate(qs) if _inside(q, o)), None)
                if k is None:
                    k = min(range(len(qs)), key=lambda i: sum((o[0] - c[0]) ** 2 + (o[1] - c[1]) ** 2 for c in qs[i]))
                groups[k].append(item)
            for k, items in groups.items():
                if not items:
                    continue
                # vi sinh 5 vị trí = 4 góc + 1 trung tâm (2-1-2) — người dùng yêu cầu 17/11/2026 (ICU)
                counts = [2, 1, 2] if key == "vs" and len(items) == 5 else None
                for item, p in zip(items, snap_to_grid([_orig(it) for it in items], qs[k], counts)):
                    pos[(r["vn"], key, item[0])] = (p, qs[k])
    return pos


# Chữ in sẵn trên mặt bằng (tọa độ ảnh nguồn) mà ký hiệu không được đè -> nhãn dẫn mũi tên (add_callouts)
PLAN_TEXT_BOXES = [
    ("t6", (980, 3815, 1265, 3955)),        # "KHU ICU NGOẠI 64.9m2"
    ("t6", (2130, 4060, 2255, 4170)),       # "P. VIP 12.8m2"
    ("t6", (2160, 3715, 2300, 3810)),       # "P.ĐỆM 7.3m2"
    ("ham", (950, 955, 1255, 1085)),        # "PHÒNG MỔ 35m2"
    ("pm10", (2025, 2260, 2090, 2530)),     # "PHÒNG MỔ 10" (chữ dọc)
]
# Nhãn code thêm (lớp Chú thích): tiêu đề 3 mặt bằng + tên phòng chưa có trên mặt bằng — tọa độ mặt bằng ghép
ROOM_LABELS = [
    (380, 60, "TẦNG 6 – NỘI TRÚ C2"),
    (300, BOT_Y + 50, "TẦNG HẦM – NỘI TRÚ C2"),
    (PM10_X + 330, BOT_Y + 50, "TẦNG TRỆT – PHÒNG MỔ 10"),
    (*src_to_plan("t6", 2193, 4040), "VIP 3"),
    (*src_to_plan("t6", 1500, 3040), "SẢNH / HÀNH LANG"),
]
ROOM_FONT = 40


def _text_boxes() -> list:
    boxes = []
    for part, (x0, y0, x1, y1) in PLAN_TEXT_BOXES:
        (a, b), (c, d) = src_to_plan(part, x0, y0), src_to_plan(part, x1, y1)
        boxes.append((min(a, c), min(b, d), max(a, c), max(b, d)))
    for cx, cy, t in ROOM_LABELS:
        w, h = len(t) * ROOM_FONT * 0.62 + 16, ROOM_FONT * 1.25 + 6
        boxes.append((cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2))
    return boxes


GREEN = {"preset": "point", "fill": "#66D95B", "stroke": "#2E7D32", "stroke_width": 3, "color": "#000000",
         "bold": False, "font_size": 40, "w": 66, "h": 66}
# vị trí lấy mẫu vi sinh: hình thoi đều màu xanh lá cây (người dùng yêu cầu 17/11/2026)
DIAMOND = {"shape": "rect", "extra": {"shape": "rhombus"}, "fill": "#66D95B", "stroke": "#2E7D32", "stroke_width": 3,
           "color": "#000000", "bold": False, "font_size": 36, "w": 88, "h": 88}


def build_layouts(overwrite: bool = False) -> list:
    from Core.Tools.drawio_layout import ensure_layout
    from Core.Tools.even_grid import add_callouts
    rooms = {lang: list(ROOM_LABELS) for lang in ("vn", "en")}
    out = []
    pos = snapped_positions()
    for sub, sym, key, cap_vn, leg_vn, leg_en in (
            ("bui", GREEN, "dust", "Sơ đồ 1: Vị trí đo nồng độ hạt bụi",
             "Vị trí đo nồng độ hạt bụi", "Particle concentration measuring point"),
            ("vi_sinh", DIAMOND, "vs", "Sơ đồ 2: Vị trí lấy mẫu vi sinh vật trong không khí",
             "Vị trí lấy mẫu vi sinh vật trong không khí", "Airborne microbial sampling point")):
        markers = []
        for r in ROOMS:
            for item in r[key]:
                (x, y), quad = pos[(r["vn"], key, item[0])]
                markers.append({**sym, "text": str(item[0]), "x": x, "y": y, "region": quad})
        led = add_callouts(markers, _text_boxes(), bounds=(10, 10, PLAN_W - 10, PLAN_H - 10),
                           radii=(110, 150, 200, 260, 330, 420))
        if led:
            print(f"[build_120423_dot8] Sơ đồ {sub}: nhãn {', '.join(led)} đè chữ -> dẫn mũi tên")
        legend = {"vn": [({**sym, "text": "1"}, leg_vn)], "en": [({**sym, "text": "1"}, leg_en)]}
        pngs = ensure_layout(os.path.join(DIAGRAM_DIR, sub), _make_plan, markers, rooms=rooms, legend=legend,
                             legend_cols={"vn": 1, "en": 1}, overwrite=overwrite)
        out.append({"path": pngs["vn"], "caption_vn": cap_vn, "caption_en": cap_vn})
    return out


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def build_dossier(layouts) -> Dossier:
    photos = ([{"path": _optimized(os.path.join(PHOTO_DIR, f)), "caption_vn": "", "caption_en": ""}
               for f in sorted(os.listdir(PHOTO_DIR)) if f.lower().endswith(".jpg")] if PHOTO_DIR else [])
    locations = []
    for r in ROOMS:
        nd, nv = len(r["dust"]), len(r["vs"])
        locations.append(MonitoringLocation(
            name_vn=r["vn"], name_en=r["en"], area_m2=r["area"], clean_class=f"ISO {ISO_CLASS}",
            state_vn="Trạng thái nghỉ", state_en="At rest",
            points_summary_vn=[f"Nồng độ hạt bụi: {nd:02d} vị trí;", f"Vi sinh vật trong không khí: {nv:02d} vị trí."],
            points_summary_en=[f"Particle concentration: {nd:02d} points;", f"Microorganisms: {nv:02d} points."],
        ))
    bui = [{"cap": f"ISO {ISO_CLASS}", "gioi_han": list(ISO8_LIMITS), "trang_so_do": TRANG_SO_DO_BUI,
            "diem": [{"khu_vuc": r["vn"], "vi_tri": n, "gia_tri": dust_vals(f)} for r in ROOMS for n, f, _ in r["dust"]]}]
    vi_sinh = [{"cap": GRADE, "gioi_han": VS_LIMIT, "trang_so_do": TRANG_SO_DO_VS,
                "diem": [{"khu_vuc": r["vn"], "vi_tri": n, "gia_tri": vs_value(s)} for r in ROOMS for n, s, _ in r["vs"]]}]
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=SUB_NO, is_multi_cert=True,
        dossier_type="cleanroom",
        scope_vn="Quan trắc chất lượng phòng sạch",            # như TBKQ đợt 6
        scope_en="Cleanroom quality monitoring",
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01, N2.TC.09",
        procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01, N2.TC.09",
        standards_vn=("- ISO 14644-1:2015 - Cleanrooms and associated controlled environments — Part 1: Classification "
                      "of air cleanliness by particle concentration;\n- WHO TRS 1044 (2022) – Phụ lục 2 – WHO good "
                      "manufacturing practices for sterile pharmaceutical products."),
        standards_en=("- ISO 14644-1:2015 - Cleanrooms and associated controlled environments — Part 1: Classification "
                      "of air cleanliness by particle concentration;\n- WHO TRS 1044 (2022) – Annex 2 – WHO good "
                      "manufacturing practices for sterile pharmaceutical products."),
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),     # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam
        equipments=[eq("aerotrak_9350"), eq("buck_b30120")],
        locations=locations,
        photos=photos, layouts=layouts,
        results_data={"bui_3kenh": bui, "vi_sinh": vi_sinh, "layout_cols": 1, "layout_width_cm": 13.5,
                      "photo_width_cm": 7.0},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX (helper chép từ build_127215.py)
# ---------------------------------------------------------------------------
def _replace_in_paragraph(p_el, old: str, new: str) -> bool:
    ts = [t for t in p_el.iter(qn("w:t"))]
    full = "".join(t.text or "" for t in ts)
    i = full.find(old)
    if i < 0:
        return False
    j = i + len(old)
    pos = 0
    done = False
    for t in ts:
        s, e = pos, pos + len(t.text or "")
        pos = e
        if e <= i or s >= j:
            continue
        txt = t.text or ""
        a, b = max(i, s) - s, min(j, e) - s
        t.text = txt[:a] + ("" if done else new) + txt[b:]
        t.set(qn("xml:space"), "preserve")
        done = True
    return True


def _replace_in_doc(doc, old, new) -> int:
    n = 0
    if old == new or old in new:          # tránh lặp vô hạn (vd LOD < 10 -> < 10)
        return 0
    for p in doc.element.body.iter(qn("w:p")):
        while _replace_in_paragraph(p, old, new):
            n += 1
    return n


def _grid_cells(tr):
    out, col = [], 0
    trPr = tr.find(qn("w:trPr"))
    gb = trPr.find(qn("w:gridBefore")) if trPr is not None else None
    if gb is not None:
        col = int(gb.get(qn("w:val")))
    for tc in tr.findall(qn("w:tc")):
        tcPr = tc.find(qn("w:tcPr"))
        gs = tcPr.find(qn("w:gridSpan")) if tcPr is not None else None
        span = int(gs.get(qn("w:val"))) if gs is not None else 1
        out.append((tc, col, span))
        col += span
    return out


def _shift_width(tbl, src: int, dst: int, frac: float):
    """Chuyển `frac` độ rộng cột lưới `src` sang cột `dst`, cập nhật tcW các ô."""
    cols = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    d = int(int(cols[src].get(qn("w:w"))) * frac)
    cols[src].set(qn("w:w"), str(int(cols[src].get(qn("w:w"))) - d))
    cols[dst].set(qn("w:w"), str(int(cols[dst].get(qn("w:w"))) + d))
    widths = [int(c.get(qn("w:w"))) for c in cols]
    for tr in tbl.findall(qn("w:tr")):
        for tc, c0, span in _grid_cells(tr):
            tcPr = tc.find(qn("w:tcPr"))
            tcW = tcPr.find(qn("w:tcW")) if tcPr is not None else None
            if tcW is None:
                continue
            w = sum(widths[c0:c0 + span])
            if tcW.get(qn("w:type")) == "pct":
                tcW.set(qn("w:w"), str(round(5000 * w / sum(widths))))
            elif tcW.get(qn("w:type")) == "dxa":
                tcW.set(qn("w:w"), str(w))


def _fix_cover_lists(doc):
    """Thiết bị / Căn cứ (Table 0): bỏ numPr (tránh '- - '), thụt 0 (AGENTS mục 14)."""
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


def _keep_table_together(tbl):
    rows = tbl.rows
    for i, row in enumerate(rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        for tc in row._tr.tc_lst:
            for p_el in tc.iter(qn("w:p")):
                pPr = p_el.get_or_add_pPr()
                for el in pPr.findall(qn("w:keepNext")):
                    pPr.remove(el)
                if i < len(rows) - 1:
                    pPr.insert(0, OxmlElement("w:keepNext"))


def _superscript_header_marks(tbl):
    from docx.text.run import Run
    for row in tbl.rows[:2]:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in list(p.runs):
                    m = re.search(r"\(\d\)", r.text)
                    if not m:
                        continue
                    if r.text.strip() == m.group(0):
                        r.font.superscript = True
                        continue
                    if r.font.superscript:
                        continue
                    before, mark, after = r.text[:m.start()], m.group(0), r.text[m.end():]
                    r.text = before
                    nr = copy.deepcopy(r._r)
                    r._r.addnext(nr)
                    sr = Run(nr, p)
                    sr.text = mark
                    sr.font.superscript = True
                    if after.strip():
                        ar = copy.deepcopy(r._r)
                        nr.addnext(ar)
                        Run(ar, p).text = after


def _keep_headings_with_tables(doc):
    for el in doc.element.body.iterchildren(qn("w:tbl")):
        if el.find(".//" + qn("w:drawing")) is not None:
            continue          # bảng ảnh / sơ đồ: không keepNext cả khối (gây nửa trang trắng — AGENTS mục 15)
        prev = el.getprevious()
        if prev is not None and prev.tag == qn("w:p") and "".join(t.text or "" for t in prev.iter(qn("w:t"))).strip():
            pPr = prev.get_or_add_pPr()
            if pPr.find(qn("w:keepNext")) is None:
                pPr.insert(0, OxmlElement("w:keepNext"))
            rows = el.findall(qn("w:tr"))
            for tr in rows[:min(5, len(rows) - 1)]:
                for p_el in tr.iter(qn("w:p")):
                    ppr = p_el.get_or_add_pPr()
                    if ppr.find(qn("w:keepNext")) is None:
                        ppr.insert(0, OxmlElement("w:keepNext"))


def _highlight_placeholders(doc):
    roots = [doc.element.body]
    for s in doc.sections:
        roots += [s.header._element, s.footer._element, s.first_page_header._element]
    for root in roots:
        for r in list(root.iter(qn("w:r"))):
            t = r.find(qn("w:t"))
            if t is None or PLACEHOLDER not in (t.text or "") or len(r.findall(qn("w:t"))) != 1:
                continue
            parts = t.text.split(PLACEHOLDER)
            pieces = []
            for i, part in enumerate(parts):
                if part:
                    pieces.append((part, False))
                if i < len(parts) - 1:
                    pieces.append((PLACEHOLDER, True))
            parent, idx = r.getparent(), r.getparent().index(r)
            extras = [c for c in r if c.tag not in (qn("w:rPr"), qn("w:t"))]
            if extras:
                lead = copy.deepcopy(r)
                for c in list(lead):
                    if c.tag == qn("w:t"):
                        lead.remove(c)
                parent.insert(idx + 1, lead)
                idx += 1
            for k, (txt, hl) in enumerate(pieces):
                nr = copy.deepcopy(r)
                for c in list(nr):
                    if c.tag not in (qn("w:rPr"), qn("w:t")):
                        nr.remove(c)
                nt = nr.find(qn("w:t"))
                nt.text = txt
                nt.set(qn("xml:space"), "preserve")
                if hl:
                    rpr = nr.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr")
                        nr.insert(0, rpr)
                    h = OxmlElement("w:highlight")
                    h.set(qn("w:val"), "yellow")
                    rpr.append(h)
                parent.insert(idx + 1 + k, nr)
            parent.remove(r)


def _keep_result_blocks(doc):
    """Bảng kết quả ngắn + Ghi chú + 'Đánh giá:' + câu kết luận cùng trang; câu kết luận chỉ keepLines (AGENTS 3, 15)."""
    def kn(el, on=True):
        for p_el in ([el] if el.tag == qn("w:p") else list(el.iter(qn("w:p")))):
            ppr = p_el.get_or_add_pPr()
            for x in ppr.findall(qn("w:keepNext")):
                ppr.remove(x)
            if on:
                ppr.insert(0, OxmlElement("w:keepNext"))

    tbls = list(doc.element.body.iterchildren(qn("w:tbl")))[1:]
    for tbl in tbls:
        nxt = tbl.getnext()
        gaps = []
        while nxt is not None and nxt.tag == qn("w:p") and not "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip():
            gaps.append(nxt)
            nxt = nxt.getnext()
        if nxt is None or nxt.tag != qn("w:tbl"):
            continue
        first = "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip()
        if not first.startswith(("Ghi chú", "Note")):
            continue
        rows = tbl.findall(qn("w:tr"))
        for tr in (rows if len(rows) <= 12 else rows[-2:]):
            kn(tr)
        for g in gaps:
            kn(g)
        head = nxt.getnext()
        if head is None or head.tag != qn("w:p"):
            continue
        if "".join(t.text or "" for t in head.iter(qn("w:t"))).strip() not in ("Đánh giá:", "Assessment:"):
            kn(nxt, False)
            continue
        kn(nxt)
        kn(head)
        sent = head.getnext()
        if sent is not None and sent.tag == qn("w:p"):
            kn(sent, False)
            ppr = sent.get_or_add_pPr()
            if ppr.find(qn("w:keepLines")) is None:
                ppr.insert(0, OxmlElement("w:keepLines"))


def _keep_room_groups(doc):
    """Bảng kết quả có cột khu vực gộp dọc: các dòng của cùng một khu vực đi liền nhau (không chẻ 'Airlock' qua trang)."""
    for tbl in list(doc.element.body.iterchildren(qn("w:tbl")))[1:]:
        rows = tbl.findall(qn("w:tr"))

        def cont(tr):
            tcs = tr.findall(qn("w:tc"))
            vm = tcs[0].find(f"{qn('w:tcPr')}/{qn('w:vMerge')}") if tcs else None
            return vm is not None and vm.get(qn("w:val")) in (None, "continue")
        size, group, k = {}, {}, None                   # số dòng của nhóm chứa mỗi dòng
        for i, tr in enumerate(rows):
            k = k if cont(tr) and k is not None else i
            size[k] = size.get(k, 0) + 1
            group[i] = k
        for i, (tr, nxt) in enumerate(zip(rows[2:], rows[3:]), 2):          # bỏ 2 dòng tiêu đề
            if cont(nxt) and size[group[i]] <= 8:          # cùng khu vực, nhóm ngắn -> đi liền (nhóm dài cho ngắt)
                for p_el in tr.iter(qn("w:p")):
                    ppr = p_el.get_or_add_pPr()
                    if ppr.find(qn("w:keepNext")) is None:
                        ppr.insert(0, OxmlElement("w:keepNext"))


def _fix_vi_sinh(doc):
    """Bảng vi sinh: dòng giới hạn theo WHO TRS 1044 – GMP Grade D; thêm đoạn Đánh giá sau bảng Ghi chú (E44c)."""
    vs_tbls = [t for t in doc.tables if "vi sinh" in t.rows[-1].cells[0].text and "WHO TRS" in t.rows[-1].cells[0].text]
    if len(vs_tbls) != 1:
        print(f"[Warning] build_120423: thấy {len(vs_tbls)} bảng vi sinh (cần 1)")
        return
    t = vs_tbls[0]
    last = t._tbl.findall(qn("w:tr"))[-1]
    old = "".join(x.text or "" for x in last.iter(qn("w:t")))
    new = (f"Giới hạn mức độ ô nhiễm vi sinh vật đối với khu vực sạch {GMP(GRADE)} nêu tại {BASIS_VN}")
    cell_p = next(p for p in last.iter(qn("w:p")) if "".join(x.text or "" for x in p.iter(qn("w:t"))).strip())
    _replace_in_paragraph(cell_p, "".join(x.text or "" for x in cell_p.iter(qn("w:t"))), new)
    print(f"[build_120423_dot8] dòng giới hạn vi sinh: '{old}' -> '{new}'")
    _superscript_header_marks(t)       # 18 dòng: không ép cả bảng đi liền (để bảng ngắt trang, bỏ trang trống cuối)
    note_el = t._tbl.getnext()
    while note_el is not None and note_el.tag != qn("w:tbl"):
        note_el = note_el.getnext()
    paras = list(doc.paragraphs)
    head = next(p for p in paras if p.text.strip() == "Đánh giá:")
    sent = paras[paras.index(head) + 1]
    new_head, new_sent = copy.deepcopy(head._p), copy.deepcopy(sent._p)
    note_el.addnext(new_head)
    new_head.addnext(new_sent)
    rooms = "; ".join(f"“{r['vn']}”" for r in ROOMS)
    txt = (f"Kết quả phân tích tổng số vi sinh vật hiếu khí trong không khí tại {rooms} {verdict(vs_ok(), False)} "
           f"với mức quy định đối với khu vực sạch {GMP(GRADE)} theo {BASIS_VN}.")
    set_assessment_paragraph(Paragraph(new_sent, head._parent), txt, font_size=13, keep_next=False)


def _fix_misc(doc):
    """Cấp sạch bảng khu vực: 'ISO 8' + 'GMP Grade D' (vi sinh); câu WHO TRS 961 còn sót -> TRS 1044."""
    for t in doc.tables:
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if "Cấp sạch" in hdr and len(t.columns) >= 3:
            if len(t._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))) == 4:
                _shift_width(t._tbl, src=3, dst=2, frac=0.06)     # 'ISO 8; GMP / Grade D' 2 dòng, cột chỉ tiêu 2 dòng
            for row in t.rows[1:]:
                for c in row.cells:
                    for p in c.paragraphs:
                        if p.text.strip() == f"ISO {ISO_CLASS}":
                            _replace_in_paragraph(p._p, f"ISO {ISO_CLASS}", f"ISO {ISO_CLASS}; {GMP(GRADE)}")
            for tr in t._tbl.findall(qn("w:tr"))[1:]:
                for h in tr.iter(qn("w:trHeight")):
                    if int(h.get(qn("w:val"), "0")) > 760:
                        h.set(qn("w:val"), "560")
    _replace_in_doc(doc, f"đối với cấp sạch {GRADE} nêu tại", f"đối với khu vực sạch {GMP(GRADE)} nêu tại")
    _replace_in_doc(doc, "WHO TRS 961 (2011) – Phụ lục 6", BASIS_VN)
    lod = round(1000 / (VS_FLOW_LPM * VS_MINUTES))
    _replace_in_doc(doc, "“< 10 CFU/m³”", f"“< {lod} CFU/m³”")


def _drop_trailing_empty(doc):
    """Xóa đoạn trống ngay trước đường kết thúc (mục bị bỏ khỏi template) — tránh trang cuối chỉ có đường kết thúc."""
    body = doc.element.body
    paras = [el for el in body if el.tag == qn("w:p")]
    end = next((p for p in reversed(paras) if "AlternateContent" in p.xml), None)
    if end is None:
        return
    from Core.Styling.elements import is_blank_paragraph   # không xét chuỗi xml ('drawingml' trong namespace)
    prev = end.getprevious()
    while is_blank_paragraph(prev):
        nxt = prev.getprevious()
        body.remove(prev)
        prev = nxt


def finish_tbkq(src: str) -> str:
    doc = Document(src)
    _drop_trailing_empty(doc)
    _fix_cover_lists(doc)
    _fix_misc(doc)
    _fix_vi_sinh(doc)
    _keep_headings_with_tables(doc)
    _keep_result_blocks(doc)
    _keep_room_groups(doc)
    style_note_tables(doc)      # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3)
    _highlight_placeholders(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - viet - {SUB_NO} - phong sach - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    return dst


# ---------------------------------------------------------------------------
# BCKQ
# ---------------------------------------------------------------------------
SO_BUI = f"{CODE}/{SUB_NO}/Bụi"
SO_VI_SINH = f"{CODE}/{SUB_NO}/Vi sinh"


def bckq_ref_lines(kqtn, vs_date):
    mapping = "; ".join(f"{r['vn']}: mẫu {', '.join(str(s) for _, s, _ in r['vs'])}" for r in ROOMS)
    n_blank = f" và {len(VI_SINH_BLANKS):02d} mẫu blank" if VI_SINH_BLANKS else ""
    return [
        f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {SO_BUI} ngày {SAMPLE_DATE}.",
        f"- Kết quả lấy mẫu vi sinh vật trong không khí ({N_VS} mẫu{n_blank}): xem tại phiếu Kết quả thử nghiệm "
        f"số {kqtn} ngày {vs_date} và phiếu Tính toán kết quả vi sinh vật trong không khí số {SO_VI_SINH} "
        f"ngày {vs_date}.",
        f"- Ký hiệu mẫu vi sinh (phiếu ghi chép, trùng ký hiệu vị trí trên sơ đồ): {mapping}"
        + (f"; mẫu blank: {BLANK_LIST}." if VI_SINH_BLANKS else "."),
    ]


def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["noi_dung_yc"] = "Lấy mẫu, phân tích vi sinh, và đo nồng độ bụi phòng sạch"     # BBLM mục 4
    vs_date = KQTN_NGAY or PLACEHOLDER
    ctx["tham_chieu"] = [
        *bckq_ref_lines(KQTN_SO or PLACEHOLDER, vs_date),
        *(["- KQTN của mẫu Blank: " + "; ".join(
            f"mẫu {s}: {v if v is not None else PLACEHOLDER} CFU/đĩa" for s, v in VI_SINH_BLANKS.items()) + "."]
          if VI_SINH_BLANKS else []),
    ]
    vs_name = "Vi sinh vật trong không khí (tổng số vi sinh vật hiếu khí)"
    dok, vok = dust_ok(), vs_ok()
    dat = (["Nồng độ hạt bụi"] if dok else []) + ([vs_name] if vok else [])
    khong = (["Nồng độ hạt bụi"] if dok is False else []) + ([vs_name] if vok is False else [])
    if vok is None:
        dat.append(f"{vs_name}: {PLACEHOLDER}")
    ctx["danh_gia"] = [{
        "khu_vuc": "; ".join(f"“{r['vn']}”" for r in ROOMS),
        "dat": ";\n".join(dat),
        "khong_dat": ";\n".join(khong) if khong else ("Không" if vok is not None else PLACEHOLDER),
        "can_cu": (f"ISO 14644-1:2015 (cấp sạch ISO {ISO_CLASS});\n{BASIS_VN} (khu vực sạch {GMP(GRADE)}: "
                   f"vi sinh vật ≤ {VS_LIMIT} CFU/m³)"),
    }]
    out = b.build(ctx=ctx, out_name=f"BCKQ - {SUB_NO} - phong sach - {CODE}.docx")
    doc = Document(out)
    _fix_misc(doc)
    _keep_headings_with_tables(doc)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Hình ảnh", "Sơ đồ", "Báo cáo của GĐV")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    return out


# ---------------------------------------------------------------------------
# Excel phiếu tính toán (form phòng sạch chung, AGENTS mục 8C) — "Bụi 0,3" + "Vi sinh"
# ---------------------------------------------------------------------------
def build_excel() -> str:
    from openpyxl.styles import PatternFill
    yellow = PatternFill("solid", fgColor="FFFF00")
    book = CleanroomCalcBook()
    groups = [{"room": r["vn"], "cap": f"ISO {ISO_CLASS}",
               "limits": {"Ch2": ISO8_LIMITS[0], "Ch3": ISO8_LIMITS[1], "Ch6": ISO8_LIMITS[2]},
               "zero": RECS[r["zero"]], "points": [RECS[f] for _, f, _ in r["dust"]],
               "labels": [n for n, _, _ in r["dust"]]} for r in ROOMS]
    book.dust_aerotrak(groups, so=SO_BUI, date=SAMPLE_DATE, labels=LABELS)

    vs_date = KQTN_NGAY or PLACEHOLDER
    wv = book.extra_sheet("Vi sinh", so=SO_VI_SINH, date=vs_date, widths=(6, 36, 9, 11, 10, 10, 15, 13, 12, 11),
                          title="TÍNH TOÁN KẾT QUẢ VI SINH VẬT TRONG KHÔNG KHÍ")
    hdr = ("STT", "Vị trí lấy mẫu", "Ký hiệu mẫu", "Lưu lượng, L/phút", "Thời gian, phút", "Thể tích, L",
           "Số khuẩn lạc, CFU/đĩa (theo KQTN)", "Kết quả, CFU/m³", "Mức quy định, CFU/m³", "Đánh giá")
    for c, lbl in enumerate(hdr, 1):
        book.put(wv, 1, c, lbl, bold=True)
    wv.row_dimensions[1].height = 45
    items = sorted(((s, r["vn"], n) for r in ROOMS for n, s, _ in r["vs"]))
    for i, it in enumerate(items + [("blank", s) for s in VI_SINH_BLANKS]):
        r = 2 + i
        book.put(wv, r, 1, i + 1)
        if it[0] == "blank":
            bv = VI_SINH_BLANKS[it[1]]
            book.put(wv, r, 2, "Mẫu blank (TS)", align=LEFT)
            book.put(wv, r, 3, it[1])
            for c in (4, 5, 6, 8, 9, 10):
                book.put(wv, r, c, "-")
            c_ = book.put(wv, r, 7, bv)
            if bv is None:
                c_.fill = yellow
            continue
        sample, room, n = it
        book.put(wv, r, 2, f"{room} - vị trí {n}", align=LEFT)
        book.put(wv, r, 3, sample)
        book.put(wv, r, 4, VS_FLOW_LPM)
        book.put(wv, r, 5, VS_MINUTES)
        book.put(wv, r, 6, f"=D{r}*E{r}")
        cfu = VI_SINH_CFU_DIA.get(sample)
        c_ = book.put(wv, r, 7, "< 1" if cfu == 0 else cfu)      # ghi như KQTN (AGENTS mục 13)
        if cfu is None:
            c_.fill = yellow
        kq = VI_SINH_KQ.get(sample)
        if kq not in (None, ""):
            book.put(wv, r, 8, kq)
        else:
            book.put(wv, r, 8, f'=IF(G{r}="","",IF(ISNUMBER(G{r}),ROUND(G{r}*1000/F{r},0),'
                               f'"< "&ROUND(1000/F{r},0)))', fmt="0")
        book.put(wv, r, 9, f"≤ {VS_LIMIT}")
        book.put(wv, r, 10, f'=IF(H{r}="","",IF(ISNUMBER(H{r}),IF(H{r}<={VS_LIMIT},"Đạt","Không đạt"),"Đạt"))')

    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - {SUB_NO} - TinhToanPhongSach.xlsx")
    book.save(out)
    for n in book.notes:
        print(f"[build_120423_dot8] {n}")
    print(f"[build_120423_dot8] Excel -> {out}")
    return out


def split_edoc() -> list:
    """BBLM = tr.1–2; Ghi chép = tr.7–10, 3 (bảng khối lượng có đánh dấu tay), 5 (ghi tay số file), 6 (mặt bằng
    đánh số vị trí); tr.4 trắng bỏ; không có trang 'Khác' (AGENTS mục 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(SCAN)
    outs = []
    # BBLM = tr.1–2; Ghi chép = tr.3–11 (phiếu + mặt bằng vi sinh PM10), 13, 15, 17 (mặt bằng đánh dấu);
    # tr.12, 14, 16, 18 mặt sau chỉ chữ thấm — bỏ
    for name, pages in ((f"{CODE6} - BBLM.pdf", [0, 1]),
                        (f"{CODE6} - Ghi Chep.pdf", [2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 14, 16])):
        dst = fitz.open()
        for pg in pages:
            dst.insert_pdf(src, from_page=pg, to_page=pg)
        p = os.path.join(edoc, name)
        finish_edoc(dst, INPUT_DIR, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_120423_dot8] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--tao-lai-so-do", action="store_true", help="Dựng lại 2 file .drawio từ tọa độ trong code")
    ap.add_argument("--chi-so-do", action="store_true", help="Chỉ dựng/xuất sơ đồ")
    args = ap.parse_args()
    layouts = build_layouts(overwrite=args.tao_lai_so_do)
    if args.chi_so_do:
        return
    d = build_dossier(layouts)
    d.ensure_output_directories()
    split_edoc()
    vn = finish_tbkq(CleanroomHandler(d).build_tbkq(is_en=False))
    bckq = build_bckq(d)
    xlsx = build_excel()
    docs = [bckq, vn]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(d)
        ex.export_files(docs)
        ex.export_excel_files([xlsx])
    print("\n===== QA =====")
    print(f"  Bụi: {'phù hợp' if dust_ok() else 'KHÔNG phù hợp'} ISO {ISO_CLASS}; vi sinh: "
          f"{ {None: 'chờ KQTN', True: 'phù hợp', False: 'KHÔNG phù hợp'}[vs_ok()] }")
    for p in docs:
        print(f"  {check_single_docx(p, code=d.code).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
