# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D133266 - Công ty TNHH Nextern Việt Nam (20/10/2026) — 2 chứng thư:
  /1 Phòng sạch: nồng độ hạt bụi 31 vị trí (Lasair III 110, 1 phút x 28,3 L/phút), trạng thái hoạt động, ISO 8
     ("Phòng sản xuất" 21 vị trí, "Air lock" 2, "Phòng thay đồ" 8 — diện tích, cấp sạch, sơ đồ theo hồ sơ 108033).
  /2 Khí nén: KN1 (đường ống khí nén tại khu vực phòng sạch), KN2 (đường ống khí nén sau lọc tại khu vực máy nén):
     tổng dầu, điểm sương, nồng độ hạt — đánh giá theo NGƯỠNG NỘI BỘ NHÀ MÁY ghi trên phiếu ghi chép:
     tổng dầu < 1 mg/m³; hàm lượng nước < 1 200 mg/m³; hạt d > 0,5 µm < 27 615 hạt/m³.

Nguồn: Input/133266/20261102150827100.pdf (BBLM tr. 1–2, ghi chép tr. 3–7, tr. 8 trắng),
       Input/133266/20261102150843557.pdf + KQDo/20261021150426651.pdf (phiếu in Lasair khí nén),
       Input/133266/133266 - TinhPhongSach.xlsx sheet "Bụi 0,1" (số đếm Lasair phòng sạch, file 1531–1562),
       Input/133266/KQDo/KN1|KN2/Use (ảnh màn hình CS DP510 đầu/cuối), KN1.png, KN2.png (ảnh vị trí),
       Input/133266/Q326D133266 - KQTN (HDTN26013537 dầu hạt 26/10; MTTN26019695.01–.03 dầu hơi 30/10).

Quyết định của người dùng (02/11/2026):
  - Hạt bụi khí nén theo Excel 24/10 của người dùng: KN1 = mẫu 5 phút 16:29–16:34 (lần 1, theo yêu cầu khách hàng
    "đo 5 phút, 28,3 L/phút"), KN2 = mẫu 1 m³ 12:05–12:40 (đúng nhãn ghi tay trên phiếu in).
  - Điểm sương KN1 theo 2 ảnh trong KQDo/KN1/Use (15:08:58 và 15:21:46), KN2 theo KQDo/KN2/Use (16:07:49, 16:20:56).
Suy luận của AI (cần người dùng xác nhận — xem PROJECT_CONTEXT mục 1):
  - Phòng sạch: diện tích/cấp sạch/sơ đồ lấy theo TBKQ 108033 (cùng 3 khu vực, cùng số vị trí).
  - Sơ đồ hệ thống khí nén theo TBKQ 108033 và 122266 (giống nhau).
  - Dầu hơi: MTTN26019695.01 -> KN1, .02 -> KN2, .03 -> blank (cả 3 KPH).
  - Ghi chép trang 4–5: Phòng sản xuất 1–21 = file 1531–1551; Air lock 1–2 = 1552–1553;
    Phòng thay đồ 4–8 = 1554–1558, 1–3 = 1559–1561; file 1562 không có trên phiếu -> không dùng.

Chạy:  python build_133266.py          (DOCX + XLSX + Edoc)
       python build_133266.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: trong code (network_dir để trống) — chép lên Z bằng robocopy (AGENTS mục 0).
"""

import os
import sys
import copy
import shutil
import zipfile
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

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.handlers.compressed_air import CompressedAirHandler
from Core.Builders.tag_context import PLACEHOLDER
from Core.Tools.khinen_calc import build_compressed_air_workbook, create_dossier_from_compressed_air_workbooks, khinen_cals
from Core.Tools.excel_borders import thin_outside_borders
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)
from Core.Tools.cleanroom_calc import CleanroomCalcBook, read_lasair_sheet

CODE = "Q326D133266"
CODE6 = "133266"
REG_DATE = "20/10/2026"           # BBLM mục 1
SAMPLE_DATE = "20/10/2026"
CERT_DATE = "30/10/2026"          # KQTN muộn nhất: MTTN26019695.01–.03 ngày 30/10/2026 (AGENTS mục 2)
KN_START, KN_END = "09:00", "17:00"           # BBLM mục 6
PS_START, PS_END = "16:42", "17:22"           # Lasair: file 1531 lúc 16:42 – file 1561 lúc 17:21 (1 phút/vị trí)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN = os.path.join(INPUT_DIR, "20261102150827100.pdf")
SCAN_LASAIR = [os.path.join(INPUT_DIR, "20261102150843557.pdf"),
               os.path.join(INPUT_DIR, "KQDo", "20261021150426651.pdf")]
RAW_XLSX = os.path.join(INPUT_DIR, f"{CODE6} - TinhPhongSach.xlsx")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photo")
DIAGRAM_DIR = os.path.join(INPUT_DIR, "TBKQ", "_diagrams")
REF_108033 = (r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\Q326D108033 - NEXTERN - Khí nén + Phòng sạch"
              r"\TBKQ - Q326D108033")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Nextern Việt Nam",
    name_en="Nextern Vietnam Limited Liability Company",
    address_vn="Nhà xưởng C1B-D5, lô C-1B-CN, đường DE5, Khu công nghiệp Mỹ Phước 3, phường Thới Hòa, Thành phố Hồ Chí Minh",
    address_en="Factory C1B-D5, lot C-1B-CN, DE5 street, My Phuoc 3 Industrial Park, Thoi Hoa ward, Ho Chi Minh City",
)

# ---------------------------------------------------------------------------
# Phòng sạch
# ---------------------------------------------------------------------------
ISO8_LIMITS = [3520000, 832000, 29300]        # ISO 14644-1:2015 Bảng 1, hạt/m³ (≥ 0,5; ≥ 1; ≥ 5 µm)
ROOMS = [  # (VN, EN, diện tích m², [(vị trí, file Lasair)]) — theo phiếu ghi chép trang 4–5
    ("Phòng sản xuất", "Production room", "139,6", [(i + 1, 1531 + i) for i in range(21)]),
    ("Air lock", "Air lock", "4,08", [(1, 1552), (2, 1553)]),
    ("Phòng thay đồ", "Gowning room", "27,6", [(1, 1559), (2, 1560), (3, 1561)] + [(i, 1550 + i) for i in range(4, 9)]),
]
TRANG_SO_DO = "2"                             # trang phụ lục chứa sơ đồ — đã kiểm trên PDF 02/11/2026


def read_particles():
    """Sheet 'Bụi 0,1' (xuất từ Lasair) -> {file: {"time", "vol" (m³), "c": [8 kênh đếm vi phân 0,1 … 5 µm]}}."""
    ws = openpyxl.load_workbook(RAW_XLSX, data_only=True)["Bụi 0,1"]
    hdr = {str(c.value).strip(): c.column - 1 for c in ws[1] if c.value}
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        n = row[hdr["Data ID"]]
        if not isinstance(n, (int, float)):
            continue
        out[int(n)] = {
            "time": row[hdr["Sample Time"]].strftime("%H:%M:%S"),
            "vol": float(str(row[hdr["Sample Volume"]]).replace(",", ".")),
            "c": [int(row[hdr[f"Counts {i}"]]) for i in range(1, 9)],
        }
    need = [f for _, _, _, pts in ROOMS for _, f in pts]
    missing = [f for f in need if f not in out]
    if missing:
        raise SystemExit(f"[FATAL] {RAW_XLSX} thiếu file đo {missing}")
    return out


PARTICLES = read_particles()


def cumulative(f: int):
    """[≥ 0,5; ≥ 1; ≥ 5 µm] hạt/m³ = tổng số đếm các kênh từ cỡ đó trở lên / thể tích mẫu (giống công thức sheet 'Bụi 0,1')."""
    p = PARTICLES[f]
    c, v = p["c"], p["vol"]
    return [round(sum(c[5:]) / v), round(sum(c[6:]) / v), round(c[7] / v)]


# ---------------------------------------------------------------------------
# Khí nén — số liệu đọc từ phiếu in Lasair, ảnh màn hình CS DP510, KQTN
# ---------------------------------------------------------------------------
REQ = {"oil": "< 1mg/m3", "water": "< 1200mg/m3", "dust": "< 27 615 hạt/m3"}   # ngưỡng nội bộ nhà máy (phiếu ghi chép tr. 2)
KQTN_NGAY_YC = "28/10/2026"       # TẠM: ngày lấy mẫu + 8 ngày (người dùng 02/11/2026) — thay bằng ngày dự kiến trả KQ trên PYC
KQTN_DAU_HAT = {"so": "HDTN26013537", "ngay": "26/10/2026", "ngay_nhan_mau": "21/10/2026", "ngay_yc": KQTN_NGAY_YC}
KQTN_DAU_HOI = {"so": "MTTN26019695", "ngay": "30/10/2026", "ngay_nhan_mau": "21/10/2026", "ngay_yc": KQTN_NGAY_YC}
KN_POINTS = [
    {
        "sym": "KN1",
        "name_vn": "Đường ống khí nén tại khu vực phòng sạch",
        "name_en": "Compressed air pipeline in the cleanroom area",
        "photo": "KN1.png",
        # Phiếu in Lasair 20/10/2026 16:29:09–16:34:09 (5 phút, Δ hạt/m³): 0,1 0,15 0,2 0,25 0,3 0,5 1,0 5,0 µm
        "dust": [406505, 148175, 71344, 98091, 206939, 134996, 32963, 71],
        "dew": {"dewp_start": -47.33, "dewp_end": -47.54, "temp_start": 23.56, "temp_end": 23.78,
                "pres_start": 3.41, "pres_end": 3.78, "time_start": "Lúc 15h08", "time_end": "Lúc 15h21",
                "photo_start": ("KN1", "BM00011.jpg"), "photo_end": ("KN1", "BM00021.jpg")},
        "kqtn_aerosol": "< 10", "kqtn_vapor": "< 5",     # HDTN M1 KPH (LOD 0,01 mg); MTTN .01 KPH (LOD 5 µg)
        "kqtn_vn": "- Dầu hơi: xem tại phiếu MTTN26019695.01\n- Dầu hạt: xem tại phiếu HDTN26013537",
    },
    {
        "sym": "KN2",
        "name_vn": "Đường ống khí nén sau lọc tại khu vực máy nén",
        "name_en": "Compressed air pipeline after filter at the compressor area",
        "photo": "KN2.png",
        # Phiếu in Lasair 20/10/2026 12:05:04–12:40:23 (1 m³)
        "dust": [114375, 54641, 36696, 69613, 154652, 92357, 25604, 11],
        "dew": {"dewp_start": -34.27, "dewp_end": -36.24, "temp_start": 25.20, "temp_end": 24.52,
                "pres_start": 3.76, "pres_end": 3.53, "time_start": "Lúc 16h07", "time_end": "Lúc 16h20",
                "photo_start": ("KN2", "BM00022.jpg"), "photo_end": ("KN2", "BM00028.jpg")},
        "kqtn_aerosol": "< 10", "kqtn_vapor": "< 5",     # HDTN M2 KPH; MTTN .02 KPH
        "kqtn_vn": "- Dầu hơi: xem tại phiếu MTTN26019695.02\n- Dầu hạt: xem tại phiếu HDTN26013537",
    },
]
SYSTEM_CHAIN_VN = ["Máy nén khí Tecbell TW-15A/8EPM", "Bình chứa", "Lọc Tecbell TB-007G-III (C)",
                   "Máy sấy khí Tecbell TB-026AH", "Lọc Tecbell TB-007G-III (H)", "Lọc Tecbell TB-007G-III (TR)",
                   "Máy sấy khí Tecbell J-2509041", "Lọc Tecbell TB-007G-III (A)", "Lọc Tecbell TB-007G-III (T)",
                   "Điểm sử dụng"]
SYSTEM_CHAIN_EN = ["Tecbell air compressor TW-15A/8EPM", "Air receiver", "Tecbell filter TB-007G-III (C)",
                   "Tecbell air dryer TB-026AH", "Tecbell filter TB-007G-III (H)", "Tecbell filter TB-007G-III (TR)",
                   "Tecbell air dryer J-2509041", "Tecbell filter TB-007G-III (A)", "Tecbell filter TB-007G-III (T)",
                   "Point of use"]


def _optimized(path: str, out_dir: str, max_w: int = 1200) -> str:
    """Ảnh nén trong <thư mục ảnh>/_optimized (AGENTS mục 1, 7): xoay EXIF, thu nhỏ, không crop."""
    from PIL import Image, ImageOps
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".jpg")
    im = ImageOps.exif_transpose(Image.open(path))
    if im.height > im.width:          # ảnh dọc -> xoay ngang (AGENTS mục 3)
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=88)
    return out


def _layouts():
    """Sơ đồ vị trí đo bụi (khách hàng vẽ, đã đánh số) lấy từ TBKQ 108033 VN/EN -> TBKQ/_diagrams (chép 1 lần)."""
    os.makedirs(DIAGRAM_DIR, exist_ok=True)
    out = {}
    for lang, src in (("vn", "TBKQ - viet - 1 - phong sach - Q326D108033.docx"),
                      ("en", "TBKQ - en - 1 - phong sach - Q326D108033.docx")):
        dst = os.path.join(DIAGRAM_DIR, f"layout_do_bui_{lang}.png")
        if not os.path.isfile(dst):
            with zipfile.ZipFile(os.path.join(REF_108033, src)) as z:
                with open(dst, "wb") as f:
                    f.write(z.read("word/media/image3.png"))
        out[lang] = [{"path": dst, "caption_vn": "", "caption_en": ""}]
    return out


LAYOUTS = _layouts()


def cleanroom_dossier() -> Dossier:
    opt = os.path.join(PHOTO_DIR, "_optimized")
    photos = [{"path": _optimized(os.path.join(PHOTO_DIR, "Use", f), opt),
               "caption_vn": "Đo nồng độ hạt bụi", "caption_en": "Airborne particle counting"}
              for f in sorted(os.listdir(os.path.join(PHOTO_DIR, "Use"))) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    locations = [MonitoringLocation(
        name_vn=vn, name_en=en, area_m2=area, clean_class="ISO 8",
        state_vn="Trạng thái hoạt động", state_en="Operational",
        points_summary_vn=[f"Nồng độ hạt bụi: {len(pts):02d} vị trí"],
        points_summary_en=[f"Airborne particle concentration: {len(pts):02d} points"],
    ) for vn, en, area, pts in ROOMS]

    def diem(rooms):
        return [{"khu_vuc": vn, "vi_tri": n, "gia_tri": cumulative(f)} for vn, _, _, pts in rooms for n, f in pts]

    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {PS_START} đến {PS_END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=1, dossier_type="cleanroom",
        scope_vn="Giám định chất lượng phòng sạch",
        scope_en="Cleanroom quality inspection",
        location_desc_vn=f"Phòng sạch của {CLIENT.name_vn} - {CLIENT.address_vn}",   # BBLM mục 5, ghi đủ (AGENTS 15)
        location_desc_en=f"Cleanroom of {CLIENT.name_en} - {CLIENT.address_en}",
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01", procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01",
        standards_vn="ISO 14644-1:2015", standards_en="ISO 14644-1:2015",
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),     # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam
        equipments=[eq("lasair_iii_110")],
        locations=locations, photos=photos, layouts=LAYOUTS["vn"],
        results_data={"bui_3kenh": [
            {"cap": "ISO 8", "gioi_han": ISO8_LIMITS, "trang_so_do": TRANG_SO_DO, "diem": diem(ROOMS[:1])},
            {"cap": "ISO 8", "gioi_han": ISO8_LIMITS, "trang_so_do": TRANG_SO_DO, "diem": diem(ROOMS[1:])},
        ]},
        assessment={"items_list_vn": ["nồng độ hạt bụi"]},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


def build_kn_workbooks() -> list:
    """Phiếu tính toán khí nén từng vị trí (template TinhToanKhiNen, giữ công thức) -> ThucHien."""
    out_dir = os.path.join(INPUT_DIR, "ThucHien")
    outs = []
    for p in KN_POINTS:
        dew = dict(p["dew"])
        for k in ("photo_start", "photo_end"):
            dew[k] = os.path.join(INPUT_DIR, "KQDo", dew[k][0], "Use", dew[k][1])
        point_data = {
            "code": CODE, "point_label": p["sym"], "cert_date": CERT_DATE,
            "location_name": p["name_vn"], "location_symbol": p["sym"],
            "location_photo": os.path.join(INPUT_DIR, p["photo"]),
            "requirements": REQ, "water_limit": REQ["water"],
            "dewpoint": dew,
            # Hiệu chỉnh DP510 theo giấy hiệu chuẩn — theo công tắc chung AP_HIEU_CHINH (09/11/2026: tắt -> {}).
            # "Máy 1 (DP 510)" = MT10-1 (suy luận, cần người dùng xác nhận khi bật).
            "cals": khinen_cals("MT10-1", SAMPLE_DATE),
            "dust": {"device": "Lasair III 110", "raw_channels": p["dust"], "mode": "diff",  # cột Δ phiếu in
                     "dilution_factor": 1.0,
                     "customer_limit_05": REQ["dust"]},
            "oil": {"flow_aerosol": 35, "flow_vapor": 1, "time_aerosol": 240, "time_vapor": 240,   # phiếu ghi chép: 35 / 1 L/phút, 04 giờ
                    "kqtn_aerosol": p["kqtn_aerosol"], "kqtn_vapor": p["kqtn_vapor"], "customer_limit": REQ["oil"],
                    "kqtn_no_vapor": p["kqtn_vn"].split("phiếu ")[1].split("\n")[0], "kqtn_no_aerosol": KQTN_DAU_HAT["so"]},
        }
        path = os.path.join(out_dir, f"{CODE} - {p['sym']} - TinhToanKhiNen.xlsx")
        build_compressed_air_workbook(point_data, path)
        outs.append(path)
    return outs


def compressed_air_dossier(xlsx_paths: list) -> Dossier:
    d = create_dossier_from_compressed_air_workbooks(
        xlsx_paths, code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {KN_START} đến {KN_END}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=2, client=CLIENT, personnel=PersonnelInfo(executor="Phát"),
        equipments=[
            eq("lasair_iii_110_hpd2"),
            eq("cs_dp510"),
            eq("oil_kit"),
        ],
        scope_vn="Lấy mẫu, phân tích chất lượng khí nén", scope_en="Sampling and analysis of compressed air quality",
        standards_vn=("ISO 8573-1:2010;\nMức yêu cầu đối với khí nén cung cấp bởi Công ty TNHH Nextern Việt Nam."),
        standards_en=("ISO 8573-1:2010;\nRequirements for compressed air provided by Nextern Vietnam Limited Liability Company."),
    )
    d.reg_date = REG_DATE
    d.procedure_vn = d.procedure_en = "QUATEST 3 3030:2025, N2.TC.KN.01"
    d.location_desc_vn = f"Nhà xưởng của {CLIENT.name_vn} - {CLIENT.address_vn}"   # BBLM mục 5 (AGENTS 15)
    d.location_desc_en = f"Factory of {CLIENT.name_en} - {CLIENT.address_en}"
    opt = os.path.join(PHOTO_DIR, "_optimized")
    d.photos = []
    for loc, p in zip(d.locations, KN_POINTS):
        loc.name_vn, loc.name_en = p["name_vn"], p["name_en"]
        d.photos.append({"path": _optimized(os.path.join(INPUT_DIR, p["photo"]), opt),
                         "caption_vn": f"Vị trí lấy mẫu {p['sym']}", "caption_en": f"Sampling location {p['sym']}"})
    rd = d.results_data
    for r in rd["dewpoint"]["rows"]:          # hàm lượng nước ghi số nguyên như TBKQ 108033, 122266
        r["values"][4] = str(round(float(r["values"][4].replace(",", "."))))
    rd["system_chain_vn"], rd["system_chain_en"] = SYSTEM_CHAIN_VN, SYSTEM_CHAIN_EN
    rd["kqtn_by_location"] = {p["sym"]: p["kqtn_vn"] for p in KN_POINTS}
    def late(k):  # KQTN trễ hạn = ngày nhận KQTN sau ngày yêu cầu cấp — người dùng "Giữ nguyên 133266"
        # (12/11/2026): hồ sơ đã giao giữ cách tính cũ, KHÔNG áp quy tắc 10 ngày làm việc (AGENTS 13).
        from datetime import datetime
        d_ = lambda s: datetime.strptime(s, "%d/%m/%Y")
        return d_(k["ngay"]) > d_(k["ngay_yc"])
    rd["kqtn"] = [{"so": k["so"], "ngay": k["ngay"], "ngay_gui": k["ngay_nhan_mau"], "ngay_nhan": k["ngay"],
                   "ngay_yc": k["ngay_yc"],
                   "tre_han": "[X] có\t\t[  ] không" if late(k) else "[  ] có\t\t[X] không"}
                  for k in (KQTN_DAU_HAT, KQTN_DAU_HOI)]
    rd["bckq_label"] = "khi nen"
    d.input_dir, d.output_dir, d.network_dir = INPUT_DIR, "", ""
    d.thuchien_dir = os.path.join(INPUT_DIR, "ThucHien")
    d.tbkq_dir = os.path.join(INPUT_DIR, "TBKQ")
    d.edoc_dir = os.path.join(d.tbkq_dir, "Edoc")
    d.pdf_dir = os.path.join(INPUT_DIR, "pdf")
    return d


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX
# ---------------------------------------------------------------------------
def _highlight_placeholders(doc):
    """Tô vàng mọi '………' còn sót (thân văn bản, bảng) để người dùng thấy ô chờ dữ liệu."""
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
    """Thiết bị / Căn cứ (Table 0): bỏ numPr (tránh '- - '), thụt 0 để mọi dòng '- ' thẳng hàng (AGENTS mục 14)."""
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


def _keep_headings_with_tables(doc):
    """Tiêu đề mục đứng ngay trước bảng kết quả -> keepNext; tiêu đề cột + 3 dòng đầu đi liền (AGENTS mục 3)."""
    for el in doc.element.body.iterchildren(qn("w:tbl")):
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


def _is_note_table(el) -> bool:
    return el is not None and el.tag == qn("w:tbl") and \
        "".join(t.text or "" for t in el.iter(qn("w:t"))).strip().startswith(("Ghi chú", "Note"))


def _keep_tails_together(doc):
    """2 dòng cuối bảng kết quả + bảng Ghi chú + câu Đánh giá cùng trang (AGENTS mục 3, cách của
    generate_all_reports.keep_tail_together; bảng Ghi chú không có Đánh giá theo sau thì dòng cuối không keepNext — E32)."""
    def keep(tr_list):
        for tr in tr_list:
            for p_el in tr.iter(qn("w:p")):
                pPr = p_el.get_or_add_pPr()
                if pPr.find(qn("w:keepNext")) is None:
                    pPr.insert(0, OxmlElement("w:keepNext"))
    for tbl in list(doc.element.body.iterchildren(qn("w:tbl"))):
        note, gaps = tbl.getnext(), []
        while note is not None and note.tag == qn("w:p") and not "".join(t.text or "" for t in note.iter(qn("w:t"))).strip():
            gaps.append(note)                     # đoạn trống giữa bảng kết quả và bảng Ghi chú
            note = note.getnext()
        if _is_note_table(tbl) or not _is_note_table(note):
            continue
        keep(gaps)
        after = note.getnext()
        assess = after is not None and after.tag == qn("w:p") and \
            "".join(t.text or "" for t in after.iter(qn("w:t"))).strip().startswith(("Đánh giá", "Assessment"))
        keep(tbl.findall(qn("w:tr"))[-2:])
        note_rows = note.findall(qn("w:tr"))
        keep(note_rows if assess else note_rows[:-1])


def _superscript_note_marks(doc):
    """'(n)' đầu dòng trong bảng Ghi chú -> chỉ số trên (AGENTS mục 3)."""
    import re
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


def _drop_trailing_empty(doc):
    """Đoạn trống cuối văn bản -> bỏ, kể cả đoạn trống NGAY TRƯỚC đoạn chứa đường kết thúc (AGENTS mục 3)."""
    def blank(el):
        return el.tag == qn("w:p") and not "".join(t.text or "" for t in el.iter(qn("w:t"))).strip() \
            and el.find(f".//{qn('w:drawing')}") is None and "AlternateContent" not in "".join(
                c.tag for c in el.iter()) and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is None
    for el in reversed(list(doc.element.body.iterchildren(qn("w:p")))):
        if blank(el):
            el.getparent().remove(el)
            continue
        if "AlternateContent" in "".join(c.tag for c in el.iter()):
            prev = el.getprevious()
            while prev is not None and blank(prev):
                nxt = prev.getprevious()
                prev.getparent().remove(prev)
                prev = nxt
        break


def _appendix_total_from_numpages(doc):
    """Footer phụ lục: SECTIONPAGES -> {= {NUMPAGES} - 1}, cùng cỡ chữ (Core: tag_context.appendix_total_field)."""
    from Core.Builders.tag_context import appendix_total_field
    return appendix_total_field(doc)


def finish_docx(src: str, dst_name: str, cleanroom: bool = False) -> str:
    """Hậu xử lý chung rồi đổi tên theo AGENTS mục 1 (handler đặt 'TBKQ - vn - n - …', E41).
    Phòng sạch: thêm giữ liền đuôi bảng + Ghi chú + Đánh giá (form khí nén đã tự ngắt trang đúng, thêm vào thì kéo cả bảng)."""
    doc = Document(src)
    _fix_cover_lists(doc)
    _keep_headings_with_tables(doc)
    if cleanroom:
        _keep_tails_together(doc)
        _drop_trailing_empty(doc)
        _appendix_total_from_numpages(doc)
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    dst = os.path.join(os.path.dirname(src), dst_name)
    from Core.Builders.tag_context import style_note_tables
    style_note_tables(doc)      # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3, 09/11/2026)
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    return dst


def build_bckq_cleanroom(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["tham_chieu"] = [f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số "
                         f"{d.cert_no_vn} ngày {SAMPLE_DATE}."]
    ctx["danh_gia"] = [{"khu_vuc": f"“{vn}”", "dat": "Nồng độ hạt bụi", "khong_dat": "-",
                        "can_cu": "Mức quy định nồng độ hạt bụi tối đa đối với phòng sạch cấp ISO 8 nêu tại ISO 14644-1:2015"}
                       for vn, *_ in ROOMS]
    out = b.build(ctx=ctx, out_name=f"BCKQ - 1 - phong sach - {CODE}.docx")
    doc = Document(out)
    _keep_headings_with_tables(doc)
    for p in doc.paragraphs:                      # "Hình ảnh…" / "Sơ đồ…" đi liền bảng ảnh ngay sau
        if p.text.strip().startswith(("Hình ảnh", "Sơ đồ", "Báo cáo của GĐV")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    return out


def build_bckq_compressed_air(d: Dossier) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 2 - khi nen - {CODE}.docx")
    doc = Document(out)
    _keep_headings_with_tables(doc)
    _highlight_placeholders(doc)
    doc.save(out)
    return out


# ---------------------------------------------------------------------------
# Excel phiếu tính toán nồng độ hạt bụi phòng sạch (ThucHien) — giữ công thức động (AGENTS mục 11)
# ---------------------------------------------------------------------------
FONT = "Times New Roman"
THIN = Side(style="thin")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _put(ws, ref, value, bold=False, align=CENTER, fmt=None, grid=True, italic=False):
    c = ws[ref]
    c.value = value
    c.font = Font(name=FONT, size=11, bold=bold, italic=italic)
    c.alignment = align
    if grid:
        c.border = GRID
    if fmt:
        c.number_format = fmt
    return c


def _merge(ws, rng, value, **kw):
    from openpyxl.utils import range_boundaries
    _put(ws, rng.split(":")[0], value, **kw)
    ws.merge_cells(rng)
    c0, r0, c1, r1 = range_boundaries(rng)
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            ws.cell(r, c).border = GRID


def build_excel_cleanroom(d: Dossier = None) -> str:
    """Phiếu tính toán /1 theo form phòng sạch chung (Core/Tools/cleanroom_calc.py, người dùng chốt 07/11/2026):
    sheet 'Bụi 0,1' = database Lasair + công thức hạt/m³ cộng dồn. Số phiếu giữ = số chứng thư /1 (khớp BCKQ đã ban hành)."""
    so = d.cert_no_vn if d is not None else f"{CODE}/1"
    recs = read_lasair_sheet(RAW_XLSX, "Bụi 0,1")
    lim = {5: ISO8_LIMITS[0], 6: ISO8_LIMITS[1], 7: ISO8_LIMITS[2]}      # kênh ≥ 0,5; ≥ 1; ≥ 5 µm
    book = CleanroomCalcBook()
    book.dust_lasair([{"room": vn_, "cap": "ISO 8", "state": "Hoạt động", "limits": lim,
                       "labels": [n for n, _ in pts], "points": [recs[f] for _, f in pts]}
                      for vn_, _, _, pts in ROOMS], so=so, date=SAMPLE_DATE)
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - 1 - TinhToanPhongSach.xlsx")
    book.save(out)
    print(f"[build_133266] Excel -> {out}")
    return out


def split_edoc() -> list:
    """BBLM = trang 1–2; Ghi chép = trang 3–7 + phiếu in Lasair khí nén (trang 8 trắng, bỏ) -> TBKQ/Edoc (AGENTS 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(SCAN)
    bblm = fitz.open()
    bblm.insert_pdf(src, from_page=0, to_page=1)
    gc = fitz.open()
    gc.insert_pdf(src, from_page=2, to_page=6)
    gc.insert_pdf(fitz.open(SCAN_LASAIR[0]))                  # phiếu in có ghi nhãn KN1, KN2 (scan 02/11)
    gc.insert_pdf(fitz.open(SCAN_LASAIR[1]), from_page=1, to_page=1)   # trang 2 scan 21/10 (các lần đo khác)
    outs = []
    for name, pdf in ((f"{CODE6} - BBLM.pdf", bblm), (f"{CODE6} - Ghi Chep.pdf", gc)):
        p = os.path.join(edoc, name)
        finish_edoc(pdf, INPUT_DIR, name)
        pdf.save(p)
        outs.append(p)
        print(f"[build_133266] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--chi-excel", action="store_true",
                    help="Chỉ dựng lại phiếu tính toán phòng sạch /1 (+ PDF nếu --pdf); giữ TBKQ/BCKQ đã ban hành")
    args = ap.parse_args()
    if args.chi_excel:
        xlsx = build_excel_cleanroom()
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            PDFExporter(cleanroom_dossier()).export_excel_files([xlsx])
        return

    kn_xlsx = build_kn_workbooks()
    ps = cleanroom_dossier()
    kn = compressed_air_dossier(kn_xlsx)
    assign_cert_numbers([ps, kn])
    for d in (ps, kn):
        d.ensure_output_directories()
    split_edoc()

    h = CleanroomHandler(ps)
    ps_vn = finish_docx(h.build_tbkq(is_en=False), f"TBKQ - viet - 1 - phong sach - {CODE}.docx", cleanroom=True)
    ps.layouts = LAYOUTS["en"]
    ps_en = finish_docx(h.build_tbkq(is_en=True), f"TBKQ - en - 1 - phong sach - {CODE}.docx", cleanroom=True)
    ps.layouts = LAYOUTS["vn"]
    ps_bckq = build_bckq_cleanroom(ps)
    ps_xlsx = build_excel_cleanroom(ps)

    k = CompressedAirHandler(kn)
    kn_vn = finish_docx(k.build_tbkq(is_en=False), f"TBKQ - viet - 2 - khi nen - {CODE}.docx")
    kn_en = finish_docx(k.build_tbkq(is_en=True), f"TBKQ - en - 2 - khi nen - {CODE}.docx")
    kn_bckq = build_bckq_compressed_air(kn)

    docs = [ps_bckq, ps_vn, ps_en, kn_bckq, kn_vn, kn_en]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(ps).export_files(docs[:3])
        PDFExporter(kn).export_files(docs[3:])
        PDFExporter(ps).export_excel_files([ps_xlsx] + kn_xlsx)
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
