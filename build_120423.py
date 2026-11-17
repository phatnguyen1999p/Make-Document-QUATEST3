# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D120423 - Công ty TNHH Kỹ thuật Nhật Khang (BV Đa khoa Tâm Anh TP.HCM) — ĐỢT 7 (09/11/2026),
chứng thư Q326D120423/7 (đợt 6 = /6). Phòng sạch Khu nội trú C2 tầng 6 – ICU: bụi + vi sinh không khí.

Nguồn (Input/120423, chép từ Z "…\\Q326D120423-Nhật Khang - Tâm Anh - HCM\\Q326D120423 - Đợt 7 - 091126" 10/11/2026 — chỉ đọc):
  - 20261110183959477.pdf (scan 10 trang):
      tr.1–2 BBLM (BM3030.02) ngày 09/11/2026: Nhật Khang, 49/24/7A Nguyễn Hữu Cảnh, phường Thạnh Mỹ Tây;
        địa điểm "Khu vực ICU lây và nội - Bệnh viện Đa khoa Tâm Anh Thành phố Hồ Chí Minh - Số 2B Phổ Quang,
        Phường Tân Sơn Hòa"; 09h00 – 11h00 ngày 09/11/2026; AeroTrak 9350 + Buck BioCulture B30120;
        QUATEST3 3030:2025; N2.TC.PS.01, N2.TC.09; vi sinh 15 mẫu (1 → 15), bụi 43 vị trí; trạng thái nghỉ;
        "Thực hiện tại các phòng: Khu ICU nội, Khu ICU lây, phòng VIP1, phòng VIP2, Airlock"; người thực hiện
        Nguyễn Tiến Phát. Ngày PYC để trống ("…/…/2026") -> lấy 25/09/2026 như TBKQ đợt 6.
      tr.3 "Bảng khối lượng quan trắc chất lượng phòng sạch – Khu nội trú C2 (tầng hầm, tầng trệt – PM 10, tầng 6 – ICU)"
        có đánh dấu tay: diện tích, cấp sạch ISO 8 các khu tầng 6 (Airlock 9,0 m² "ước tính"); WHO TRS 961 (2011) Phụ lục 6.
      tr.4 trắng (bỏ). tr.5 ghi tay số file AeroTrak từng phòng. tr.6 mặt bằng tầng 6 đánh số vị trí bụi (bút xanh).
      tr.7–10 phiếu ghi chép (BM3030.01): bụi AeroTrak 1 phút — ICU nội zero 1280, 1281 → 1296; ICU lây zero 1297,
        1298 → 1307; VIP1 zero 1260, 1261 → 1266; Airlock zero 1267, 1268 → 1272; VIP2 zero 1273, 1274 → 1279
        (tr.8, tr.10 gạch bỏ). Vi sinh APBuck 100 L/phút × 01 phút: VIP2 6 → 7; VIP1 8 → 9; Airlock 10 → 11;
        ICU nội 1 → 5; ICU lây 12 → 15; TS blank 16.
  - 6ChannelFixed.xlsx: AeroTrak zone "120423-ISO8"; file 1259 – 1308 ngày 09/11/2026 (+ 682 – 742 ngày 05/10/2026
    của đợt trước — không dùng). File 1259 (trước zero 1260) và 1308 không có trên phiếu -> không dùng.
  - Photos-1-001 (5)/Use/*.jpg: 4 ảnh người dùng chọn.
  - Layout/Mặt bằng Màng Hepa - Nội trú C2, C3.pdf (chép từ thư mục hồ sơ trên Z), trang 2 = tầng 6 Nội trú C2.

Quyết định của người dùng (10/11/2026):
  - "Cấp sạch toàn khu vực là ISO Class 8, và GMP Grade D cho vi sinh"; tham khảo các đợt 4, 5, 6.
  - "Giữ như hiện tại": đồng ý cả 6 điểm suy luận dưới đây (10/11/2026).

Suy luận của AI (người dùng đã đồng ý 10/11/2026):
  (1) Chỉ dựng TBKQ tiếng Việt như đợt 5, 6 (không có PYC trong scan đợt 7).
  (2) Căn cứ vi sinh theo TBKQ đợt 6: WHO TRS 1044 (2022) – Phụ lục 2, GMP Grade D ≤ 200 CFU/m³ (bảng khối lượng tr.3
      ghi WHO TRS 961 (2011) Phụ lục 6 — cùng mức 200 CFU/m³).
  (3) Vị trí số viết tay trên tr.6 chuyển sang mặt bằng sạch (ghép SIFT, tỉ lệ 1,1046); vị trí lấy mẫu VI SINH không
      được đánh dấu trên mặt bằng -> code đặt tạm (gần các vị trí bụi), người dùng kéo lại trên draw.io.
  (4) Diện tích theo bảng khối lượng tr.3 (Airlock 9,0 m² là "ước tính").
  (5) Số phiếu tính toán theo số chứng thư: Q326D120423/7/Bụi, Q326D120423/7/Vi sinh.
  (6) Địa điểm giám định ghi đúng BBLM mục 5.

KQTN vi sinh (về 17/11/2026, Z …\\Đợt 7\\KQTN\\20261117161902377.pdf, 2 bản giống nhau): số 060424/VYTCC ngày 13/11/2026,
  ngày nhận mẫu 09/11/2026; "Vi sinh không khí 1 … 17" CFU/đĩa. Mẫu n = ký hiệu mẫu n trên phiếu ghi chép; 16 = blank;
  mẫu 17 ("< 1") không có trên BBLM / phiếu ghi chép -> người dùng chốt 17/11/2026: "Là blank thứ 2".
Ngày ban hành TBKQ = ngày KQTN 13/11/2026 (AGENTS mục 2).

Chạy:  python build_120423.py          (DOCX + XLSX + Edoc + sơ đồ)
       python build_120423.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
       python build_120423.py --tao-lai-so-do   (dựng lại 2 file .drawio từ tọa độ trong code, bỏ chỉnh tay)
Không đồng bộ Z trong code — chép lên thư mục "Đợt 7" trên Z bằng robocopy (AGENTS mục 0).
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
SUB_NO = 7                       # đợt 7 -> chứng thư /7
REG_DATE = "25/09/2026"          # PYC (BBLM để trống ngày; theo TBKQ đợt 6)
SAMPLE_DATE = "09/11/2026"
START, END = "09:00", "11:00"    # BBLM
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN = os.path.join(INPUT_DIR, "20261110183959477.pdf")
RAW_XLSX = os.path.join(INPUT_DIR, "6ChannelFixed.xlsx")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photos-1-001 (5)", "Use")
PLAN_PDF = os.path.join(INPUT_DIR, "Layout", "Mặt bằng Màng Hepa - Nội trú C2, C3.pdf")
DIAGRAM_DIR = os.path.join(INPUT_DIR, "TBKQ", "_diagrams")

# --- Vi sinh: chưa có KQTN. Điền khi có kết quả rồi chạy lại ---
KQTN_SO = "060424/VYTCC"   # phiếu KQTN Viện Y tế Công cộng TP.HCM (Z …\Đợt 7\KQTN\20261117161902377.pdf)
KQTN_NGAY = "13/11/2026"   # ngày phiếu KQTN -> ngày ban hành TBKQ (AGENTS mục 2); ngày nhận mẫu 09/11/2026
# CFU/đĩa theo KQTN "Vi sinh không khí n" = ký hiệu mẫu n; 0 = "< 1" (không phát hiện)
VI_SINH_CFU_DIA = {1: 0, 2: 5, 3: 5, 4: 9, 5: 7, 6: 0, 7: 2, 8: 4, 9: 6, 10: 0, 11: 0, 12: 9, 13: 7, 14: 3, 15: 8}
# Mẫu blank (CFU/đĩa): 16 theo phiếu ghi chép; 17 không có trên phiếu — người dùng chốt 17/11/2026 "Là blank thứ 2"
VI_SINH_BLANKS = {16: "< 1", 17: "< 1"}
VI_SINH_KQ = {}          # hoặc ghi thẳng kết quả CFU/m³ theo KQTN, vd {1: "< 10", 2: 20}
VS_FLOW_LPM, VS_MINUTES = 100, 1      # phiếu ghi chép: 100 L/phút × 01 phút (đợt 6: 2 phút)
BLANK_LIST = ", ".join(str(s) for s in VI_SINH_BLANKS)

CERT_DATE = KQTN_NGAY or PLACEHOLDER

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Kỹ thuật Nhật Khang",
    name_en="Nhat Khang Technical Company Limited",     # không dùng (chỉ TBKQ VN) — AI dịch
    address_vn="49/24/7A Nguyễn Hữu Cảnh, phường Thạnh Mỹ Tây, Thành phố Hồ Chí Minh",
    address_en="49/24/7A Nguyen Huu Canh Street, Thanh My Tay Ward, Ho Chi Minh City",
)
LOCATION_VN = ("Khu vực ICU lây và nội - Bệnh viện Đa khoa Tâm Anh Thành phố Hồ Chí Minh - Số 2B Phổ Quang, "
               "phường Tân Sơn Hòa, Thành phố Hồ Chí Minh")
LOCATION_EN = ("Infectious and internal ICU area - Tam Anh General Hospital Ho Chi Minh City - No. 2B Pho Quang Street, "
               "Tan Son Hoa Ward, Ho Chi Minh City")

ISO_CLASS = 8
ISO8_LIMITS = (3520000, 832000, 29300)          # ISO 14644-1:2015 Bảng 1, d ≥ 0,5 / 1 / 5 µm, hạt/m³
GRADE = "D"
VS_LIMIT = 200                                  # GMP Grade D, CFU/m³
BASIS_VN = "WHO TRS 1044 (2022) – Phụ lục 2"
BASIS_EN = "WHO TRS 1044 (2022) – Annex 2"


def GMP(g: str) -> str:
    return f"GMP Grade {g}"


# Phòng theo bảng khối lượng tr.3. dust: [(nhãn, file AeroTrak, (x, y) số viết tay trên scan tr.6 @200 dpi)]
# vs: [(nhãn vị trí, ký hiệu mẫu, (x, y) đặt tạm — suy luận (3))]
ROOMS = [
    {"vn": "Khu ICU Nội", "en": "Internal ICU area", "area": "89,6", "zero": 1280,
     "dust": list(zip(range(1, 17), range(1281, 1297),
                      [(1160, 1445), (1045, 1435), (920, 1420), (765, 1390), (765, 1270), (905, 1300), (1035, 1300),
                       (1155, 1315), (1145, 1210), (1020, 1215), (905, 1205), (780, 1220), (760, 1060), (885, 1065),
                       (1020, 1075), (1160, 1080)])),
     "vs": [(1, 1, (820, 1060)), (2, 2, (1100, 1080)), (3, 3, (975, 1265)), (4, 4, (850, 1345)), (5, 5, (1100, 1445))]},
    {"vn": "Khu ICU Lây", "en": "Infectious ICU area", "area": "44,0", "zero": 1297,
     "dust": list(zip(range(1, 11), range(1298, 1308),
                      [(1515, 1760), (1360, 1760), (1240, 1760), (1185, 1680), (1280, 1670), (1385, 1670),
                       (1540, 1680), (1480, 1555), (1370, 1545), (1255, 1565)])),
     "vs": [(1, 12, (1255, 1565)), (2, 13, (1480, 1555)), (3, 14, (1240, 1760)), (4, 15, (1515, 1760))]},
    {"vn": "Phòng VIP 1 (cạnh ICU Nội)", "en": "VIP room 1 (next to internal ICU)", "area": "11,6", "zero": 1260,
     "dust": list(zip(range(1, 7), range(1261, 1267),
                      [(580, 1010), (465, 1008), (455, 1108), (555, 1122), (568, 1162), (445, 1152)])),
     "vs": [(1, 8, (465, 1008)), (2, 9, (568, 1162))]},
    {"vn": "Phòng VIP 2 (cạnh Airlock)", "en": "VIP room 2 (next to airlock)", "area": "12,9", "zero": 1273,
     "dust": list(zip(range(1, 7), range(1274, 1280),
                      [(537, 1365), (437, 1360), (440, 1310), (537, 1320), (550, 1245), (437, 1228)])),
     "vs": [(1, 6, (437, 1228)), (2, 7, (537, 1365))]},
    {"vn": "Airlock", "en": "Airlock", "area": "9,0", "zero": 1267,
     "dust": list(zip(range(1, 6), range(1268, 1273),
                      [(660, 1407), (617, 1372), (675, 1330), (635, 1285), (677, 1248)])),
     "vs": [(1, 10, (677, 1248)), (2, 11, (617, 1372))]},
]
# nhãn phòng trên sơ đồ (đè chữ "P. VIP" in sẵn của mặt bằng) — tọa độ scan @200 dpi
ROOM_LABELS = [(515, 1120, "VIP 1"), (492, 1245, "VIP 2")]
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
# Sơ đồ (draw.io) — mặt bằng sạch trang 2 của PDF mặt bằng, xoay cùng chiều scan tr.6
# Ghép scan tr.6 (@200 dpi) -> mặt bằng (@150 dpi, xoay 270°) bằng SIFT + RANSAC (313 điểm khớp):
#   x' = 1,1046 (x − 350) + 351,91 ; y' = 1,1046 (y − 900) + 957,79
# ---------------------------------------------------------------------------
PLAN_DPI = 300
_S, _TX, _TY = 1.1046, 351.9087, 957.7921
_K = PLAN_DPI / 150
CROP = (770, 2048, 3508, 4014)        # vùng tầng 6 – ICU trên mặt bằng @300 dpi


def to_plan(x, y):
    """Tọa độ scan tr.6 (@200 dpi) -> tọa độ ảnh mặt bằng đã cắt (@300 dpi)."""
    return (_K * (_S * (x - 350) + _TX) - CROP[0], _K * (_S * (y - 900) + _TY) - CROP[1])


def _make_plan():
    import fitz
    import numpy as np
    from PIL import Image
    pm = fitz.open(PLAN_PDF)[1].get_pixmap(dpi=PLAN_DPI)
    a = np.frombuffer(pm.samples, np.uint8).reshape(pm.height, pm.width, pm.n)[..., :3]
    return Image.fromarray(np.rot90(a, 3).copy()).crop(CROP)


# Vách trong từng phòng trên ảnh mặt bằng đã cắt (@300 dpi) — vị trí bụi căn đều theo lưới (Core/Tools/even_grid.py,
# người dùng yêu cầu 17/11/2026). Mặt bằng nghiêng 6° (vách trên ICU Nội: (762,106) -> (1831,219)): hệ trục u/v xoay 6°.
# ICU Lây hình thang (vách trên gần ngang, vách dưới nghiêng 6°) -> khai 4 góc; mép trái lấy vách phải phòng đồ dịch.
# ICU Nội lấy mép phải phần trên (u = 1850), bỏ phần lồi ra ~130 px ở nửa dưới (ICU Nội không phải hình chữ nhật).
WALL_ANGLE = 6.0


def _room_quads():
    from Core.Tools.even_grid import rect_quad
    return {
        # vách dưới thật v ≈ 968; lấy 840 (mép trên hộp kỹ thuật PCCC + cửa vào ở góc dưới-trái) để hàng dưới không
        # rơi vào hộp kỹ thuật
        "Khu ICU Nội": rect_quad(754, 1850, 26, 840, WALL_ANGLE),
        "Khu ICU Lây": ((1883, 1199), (2672, 1189), (2653, 1914), (1883, 1839)),
        "Phòng VIP 1 (cạnh ICU Nội)": rect_quad(89, 542, 29, 471, WALL_ANGLE),
        "Phòng VIP 2 (cạnh Airlock)": rect_quad(85, 540, 490, 972, WALL_ANGLE),
        "Airlock": rect_quad(555, 762, 490, 970, WALL_ANGLE),
    }


def snapped_positions() -> dict:
    """{(phòng, 'dust'|'vs', nhãn): (x, y) trên mặt bằng}: bụi vào lưới căn đều, nhãn gán theo chỗ gần sơ đồ viết tay
    nhất; vi sinh đặt tại vị trí bụi (mới) gần chỗ đặt tạm nhất."""
    from Core.Tools.even_grid import snap_to_grid
    quads, pos = _room_quads(), {}
    for r in ROOMS:
        orig = [to_plan(*xy) for _, _, xy in r["dust"]]
        new = snap_to_grid(orig, quads[r["vn"]])
        for (n, _, _), p in zip(r["dust"], new):
            pos[(r["vn"], "dust", n)] = p
        for n, _, xy in r["vs"]:
            o = to_plan(*xy)
            k = min(range(len(orig)), key=lambda i: (orig[i][0] - o[0]) ** 2 + (orig[i][1] - o[1]) ** 2)
            pos[(r["vn"], "vs", n)] = new[k]
    return pos


GREEN = {"preset": "point", "fill": "#66D95B", "stroke": "#2E7D32", "stroke_width": 3, "color": "#000000",
         "bold": False, "font_size": 40, "w": 66, "h": 66}
ORANGE = {**GREEN, "fill": "#FFB74D", "stroke": "#E65100"}


def build_layouts(overwrite: bool = False) -> list:
    from Core.Tools.drawio_layout import ensure_layout
    rooms = {lang: [(*to_plan(x, y), t) for x, y, t in ROOM_LABELS] for lang in ("vn", "en")}
    out = []
    pos = snapped_positions()
    for sub, sym, key, cap_vn, leg_vn, leg_en in (
            ("bui", GREEN, "dust", "Sơ đồ 1: Vị trí đo nồng độ hạt bụi",
             "Vị trí đo nồng độ hạt bụi (đánh số theo từng khu vực)", "Particle concentration measuring point"),
            ("vi_sinh", ORANGE, "vs", "Sơ đồ 2: Vị trí lấy mẫu vi sinh vật trong không khí",
             "Vị trí lấy mẫu vi sinh vật trong không khí (đánh số theo từng khu vực)", "Airborne microbial sampling point")):
        markers = []
        for r in ROOMS:
            for item in r[key]:
                x, y = pos[(r["vn"], key, item[0])]
                markers.append({**sym, "text": str(item[0]), "x": x, "y": y})
        legend = {"vn": [({**sym, "text": "1"}, leg_vn)], "en": [({**sym, "text": "1"}, leg_en)]}
        rooms_spec = {lang: [(x, y, t) for x, y, t in v] for lang, v in rooms.items()}
        pngs = ensure_layout(os.path.join(DIAGRAM_DIR, sub), _make_plan, markers,
                             rooms={k: v for k, v in rooms_spec.items()}, legend=legend,
                             legend_cols={"vn": 1, "en": 1}, overwrite=overwrite)
        out.append({"path": pngs["vn"], "caption_vn": cap_vn, "caption_en": cap_vn})
    return out


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def build_dossier(layouts) -> Dossier:
    photos = [{"path": _optimized(os.path.join(PHOTO_DIR, f)), "caption_vn": "", "caption_en": ""}
              for f in sorted(os.listdir(PHOTO_DIR)) if f.lower().endswith(".jpg")]
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
    print(f"[build_120423] dòng giới hạn vi sinh: '{old}' -> '{new}'")
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
    mapping = "; ".join(f"{r['vn']}: vị trí {', '.join(str(n) for n, _, _ in r['vs'])} – mẫu "
                        f"{', '.join(str(s) for _, s, _ in r['vs'])}" for r in ROOMS)
    return [
        f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {SO_BUI} ngày {SAMPLE_DATE}.",
        f"- Kết quả lấy mẫu vi sinh vật trong không khí (15 mẫu và {len(VI_SINH_BLANKS):02d} mẫu blank): xem tại phiếu Kết quả thử nghiệm "
        f"số {kqtn} ngày {vs_date} và phiếu Tính toán kết quả vi sinh vật trong không khí số {SO_VI_SINH} "
        f"ngày {vs_date}.",
        f"- Ký hiệu mẫu vi sinh (phiếu ghi chép) – vị trí trên sơ đồ: {mapping}; mẫu blank: {BLANK_LIST}.",
    ]


def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["noi_dung_yc"] = "Lấy mẫu, phân tích vi sinh, và đo nồng độ bụi phòng sạch"     # BBLM mục 4
    vs_date = KQTN_NGAY or PLACEHOLDER
    ctx["tham_chieu"] = [
        *bckq_ref_lines(KQTN_SO or PLACEHOLDER, vs_date),
        "- KQTN của mẫu Blank: " + "; ".join(
            f"mẫu {s}: {v if v is not None else PLACEHOLDER} CFU/đĩa" for s, v in VI_SINH_BLANKS.items()) + ".",
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
        print(f"[build_120423] {n}")
    print(f"[build_120423] Excel -> {out}")
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
    for name, pages in ((f"{CODE6} - BBLM.pdf", [0, 1]), (f"{CODE6} - Ghi Chep.pdf", [6, 7, 8, 9, 2, 4, 5])):
        dst = fitz.open()
        for pg in pages:
            dst.insert_pdf(src, from_page=pg, to_page=pg)
        p = os.path.join(edoc, name)
        finish_edoc(dst, INPUT_DIR, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_120423] Edoc -> {p}")
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
