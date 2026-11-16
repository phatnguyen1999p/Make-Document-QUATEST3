# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D137035 - Công ty TNHH Nextern Việt Nam (28/10/2026) — 2 chứng thư:
  /1 Phòng sạch: nồng độ hạt bụi 31 vị trí (TSI AeroTrak 9350, 1 phút x 50 L/phút, chế độ Σ), trạng thái hoạt động,
     ISO 8 — "Phòng sản xuất" 21 vị trí, "Air lock" 2, "Phòng thay đồ" 8.
  /2 Khí nén: CHỈ ĐO NỒNG ĐỘ HẠT BỤI (BBLM: tổng dầu, điểm sương gạch bỏ) tại KN1, KN2 — AeroTrak 9350 + HPD III,
     mẫu 1 m³ (20 phút); mức yêu cầu nội bộ nhà máy hạt d > 0,5 µm < 27 615 hạt/m³.

Nguồn (Input/137035, chép từ Z 07/11/2026):
  20261107134630841.pdf — tr. 1–2 BBLM (BM3030.02, PYC Q326D137035 ngày 27/10/2026, thực hiện 28/10/2026
      09:00–14:30, người thực hiện Nguyễn Tiến Phát); tr. 3 ghi chép trang 1/5; tr. 4 ghi chép khí nén Vị trí 01
      "Tại máy nén khí sau lọc" (zero 1108, đo 1111, 1 m³, 6,5–7 bar, khí 24 °C); tr. 5 Vị trí 02 "Bên trong khu
      vực phòng sạch" (zero 1112, đo 1113, khí 22 °C); tr. 6–7 ghi chép phòng sạch (DPC 9350, zero 1114; Phòng sản
      xuất 1–21 = 1115–1135; Air lock zero 1136, 1–2 = 1137–1138; Thay đồ zero 1139, 1–3 = 1145–1147, 4–8 =
      1140–1144); tr. 8 trắng.
  20261107134639556.pdf — phiếu in AeroTrak khí nén, ghi tay "KN1 - zero" 1112, "KN1 - đo" 1113, "KN2 - zero" 1108,
      "KN2 - đo" 1111.
  KQ DO/6ChannelFixed.xlsx (+ .pdf) — AeroTrak 9350-02 S/N 93502239001, 33 mẫu 1115–1147 (hạt/m³ Σ).
  Photos-1-001 (3)/Use/CleanRoom (2 ảnh), Use/Compressed Air/Kn1.jpg (phòng sạch), KN2.jpg (máy nén).

Quyết định của người dùng (07/11/2026):
  - "Đồng ý, chép Z": chấp nhận toàn bộ các điểm suy luận dưới đây (mã 137035, KN1/KN2, cột 0,3 µm, mức nội bộ,
    thông số phòng sạch theo 133266, dùng file 1115, giờ đo theo máy).

Suy luận của AI (người dùng đã đồng ý 07/11/2026):
  - Người dùng gõ "136035": trên Z không có, chỉ có Q326D137035 - NEXTERN - Bụi khí nén, phòng sạch.
  - KN1 = phòng sạch (file 1113), KN2 = sau lọc máy nén (file 1111) theo nhãn ghi tay trên phiếu in và tên ảnh
    (cùng cách đánh số với 133266); phiếu ghi chép lại đánh "Vị trí 01" = máy nén, "Vị trí 02" = phòng sạch.
  - AeroTrak đo từ 0,3 µm: cột ISO 8573-1 "0,1 – 0,5 µm" ghi số kênh 0,3 – 0,5 µm, tiêu đề cột đổi thành
    "0,3 µm < d ≤ 0,5 µm"; phân loại vẫn đúng vì cấp do dải 1,0 – 5,0 µm quyết định (Class 3).
  - N* = (273 + nhiệt độ khí lấy mẫu trên phiếu) / 293, áp suất 1 bar (như form khí nén).
  - Mức nội bộ hạt d > 0,5 µm < 27 615 hạt/m³ lấy theo hồ sơ 133266 (phiếu 137035 không ghi mức).
  - Phòng sạch: diện tích, cấp sạch, sơ đồ vị trí như 133266 / 108033 (cùng 3 khu vực, cùng số vị trí).
  - Sơ đồ hệ thống khí nén như 133266.
  - Thời gian đo theo máy: khí nén 10:12–11:19, phòng sạch 13:32–14:43 (BBLM ghi kết thúc 14:30).

Chạy:  python build_137035.py          (DOCX + XLSX + Edoc)
       python build_137035.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: trong code (network_dir để trống) — chép lên Z bằng robocopy (AGENTS mục 0).
"""

import os
import sys
import copy
import shutil
import argparse
from datetime import date

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_ton_tai_Z"))

import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.handlers.compressed_air import CompressedAirHandler, keep_result_heading_with_next
from Core.Builders.tag_context import PLACEHOLDER
from Core.Tools.khinen_calc import build_compressed_air_workbook, create_dossier_from_compressed_air_workbooks
from Core.Tools.excel_borders import thin_outside_borders, strike_box
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)
from Core.Tools.cleanroom_calc import CleanroomCalcBook, read_aerotrak

CODE = "Q326D137035"
CODE6 = "137035"
REG_DATE = "27/10/2026"           # BBLM mục 1 (PYC)
SAMPLE_DATE = "28/10/2026"
CERT_DATE = date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày dựng (AGENTS mục 2)
KN_START, KN_END = "10:12", "11:19"           # phiếu in AeroTrak: 1111 10:12:26–10:32:26, 1113 10:58:51–11:18:51
PS_START, PS_END = "13:32", "14:43"           # 6ChannelFixed: 1115 lúc 13:32:38 – 1147 lúc 14:42:20 (1 phút)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN = os.path.join(INPUT_DIR, "20261107134630841.pdf")
SCAN_KN = os.path.join(INPUT_DIR, "20261107134639556.pdf")
SCAN_PS = os.path.join(INPUT_DIR, "KQ DO", "6ChannelFixed.pdf")
DUST_XLSX = os.path.join(INPUT_DIR, "KQ DO", "6ChannelFixed.xlsx")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photos-1-001 (3)")
PHOTO_PS = os.path.join(PHOTO_DIR, "Use", "CleanRoom")
PHOTO_KN = os.path.join(PHOTO_DIR, "Use", "Compressed Air")
DIAGRAM_DIR = os.path.join(INPUT_DIR, "TBKQ", "_diagrams")
REF_LAYOUT_DIR = os.path.join(WORKSPACE, "Input", "133266", "TBKQ", "_diagrams")   # sơ đồ khách vẽ (từ 108033)

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Nextern Việt Nam",
    name_en="Nextern Vietnam Limited Liability Company",
    address_vn="Nhà xưởng C1B-D5, lô C-1B-CN, đường DE5, Khu công nghiệp Mỹ Phước 3, phường Thới Hòa, Thành phố Hồ Chí Minh",
    address_en="Factory C1B-D5, lot C-1B-CN, DE5 street, My Phuoc 3 Industrial Park, Thoi Hoa ward, Ho Chi Minh City",
)
# Địa điểm giám định (BBLM mục 5 "Nhà xưởng của Công ty NEXTERN - Nhà xưởng C1B-D5 …", ghi đủ — AGENTS mục 15)
LOC_PS_VN = f"Phòng sạch của {CLIENT.name_vn} - {CLIENT.address_vn}"
LOC_PS_EN = f"Cleanroom of {CLIENT.name_en} - {CLIENT.address_en}"
LOC_KN_VN = f"Nhà xưởng của {CLIENT.name_vn} - {CLIENT.address_vn}"
LOC_KN_EN = f"Factory of {CLIENT.name_en} - {CLIENT.address_en}"

# ---------------------------------------------------------------------------
# Phòng sạch
# ---------------------------------------------------------------------------
ISO8_LIMITS = [3520000, 832000, 29300]        # ISO 14644-1:2015 Bảng 1, hạt/m³ (≥ 0,5; ≥ 1; ≥ 5 µm)
ROOMS = [  # (VN, EN, diện tích m², zero, [(vị trí, file AeroTrak)]) — phiếu ghi chép tr. 6–7
    ("Phòng sản xuất", "Production room", "139,6", None, [(i + 1, 1115 + i) for i in range(21)]),
    ("Air lock", "Air lock", "4,08", 1136, [(1, 1137), (2, 1138)]),
    ("Phòng thay đồ", "Gowning room", "27,6", 1139,
     [(1, 1145), (2, 1146), (3, 1147)] + [(i, 1136 + i) for i in range(4, 9)]),
]
TRANG_SO_DO = "2"                             # trang phụ lục chứa sơ đồ — đã kiểm trên PDF 07/11/2026

LABELS, RECS = read_aerotrak(DUST_XLSX)
_need = [f for *_, z, pts in ROOMS for f in ([z] if z else []) + [f for _, f in pts]]
_missing = [f for f in _need if f not in RECS]
if _missing:
    raise SystemExit(f"[FATAL] {DUST_XLSX} thiếu file đo {_missing}")
_unused = sorted(set(RECS) - set(_need))
if _unused:
    print(f"[WARNING] File AeroTrak không có trên phiếu ghi chép, không dùng: {_unused}")


def cumulative(f: int):
    """[≥ 0,5; ≥ 1; ≥ 5 µm] hạt/m³ — AeroTrak chế độ Σ (cộng dồn) ghi thẳng giá trị máy."""
    r = RECS[f]
    return [int(r["Ch2"]), int(r["Ch3"]), int(r["Ch6"])]


# ---------------------------------------------------------------------------
# Khí nén — chỉ nồng độ hạt bụi (phiếu in AeroTrak 9350 + HPD III, Δ hạt/m³: 0,3 0,5 1 2 3 5 µm)
# ---------------------------------------------------------------------------
REQ_DUST = "< 27 615 hạt/m3"      # mức nội bộ nhà máy (theo 133266 — phiếu 137035 không ghi)
KN_POINTS = [
    {"sym": "KN1", "name_vn": "Đường ống khí nén tại khu vực phòng sạch",
     "name_en": "Compressed air pipeline in the cleanroom area",
     "photo": "Kn1.jpg", "gas_temp": 22.0,          # phiếu ghi chép tr. 5 (Vị trí 02)
     # Sample # 1113, 28/10/2026 10:58:51–11:18:51, 1000 L
     "dust": [611, 439, 93, 16, 1, 0]},
    {"sym": "KN2", "name_vn": "Đường ống khí nén sau lọc tại khu vực máy nén",
     "name_en": "Compressed air pipeline after filter at the compressor area",
     "photo": "KN2.jpg", "gas_temp": 24.0,          # phiếu ghi chép tr. 4 (Vị trí 01)
     # Sample # 1111, 28/10/2026 10:12:26–10:32:26, 1000 L
     "dust": [1164, 574, 164, 122, 1, 0]},
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
    """Sơ đồ vị trí đo bụi (khách hàng vẽ, đã đánh số; như 133266) -> TBKQ/_diagrams."""
    os.makedirs(DIAGRAM_DIR, exist_ok=True)
    out = {}
    for lang in ("vn", "en"):
        dst = os.path.join(DIAGRAM_DIR, f"layout_do_bui_{lang}.png")
        if not os.path.isfile(dst):
            shutil.copy2(os.path.join(REF_LAYOUT_DIR, f"layout_do_bui_{lang}.png"), dst)
        out[lang] = [{"path": dst, "caption_vn": "", "caption_en": ""}]
    return out


LAYOUTS = _layouts()


def cleanroom_dossier() -> Dossier:
    opt = os.path.join(PHOTO_DIR, "_optimized")
    photos = [{"path": _optimized(os.path.join(PHOTO_PS, f), opt), "caption_vn": "", "caption_en": ""}
              for f in sorted(os.listdir(PHOTO_PS)) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    locations = [MonitoringLocation(
        name_vn=vn, name_en=en, area_m2=area, clean_class="ISO 8",
        state_vn="Trạng thái hoạt động", state_en="Operational",
        points_summary_vn=[f"Nồng độ hạt bụi: {len(pts):02d} vị trí"],
        points_summary_en=[f"Airborne particle concentration: {len(pts):02d} points"],
    ) for vn, en, area, _z, pts in ROOMS]

    def diem(rooms):
        return [{"khu_vuc": vn, "vi_tri": n, "gia_tri": cumulative(f)} for vn, _, _, _z, pts in rooms for n, f in pts]

    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {PS_START} đến {PS_END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=1, dossier_type="cleanroom",
        scope_vn="Giám định chất lượng phòng sạch",
        scope_en="Cleanroom quality inspection",
        location_desc_vn=LOC_PS_VN, location_desc_en=LOC_PS_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01", procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01",
        standards_vn="ISO 14644-1:2015", standards_en="ISO 14644-1:2015",
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),     # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam
        equipments=[eq("aerotrak_9350")],
        locations=locations, photos=photos, layouts=LAYOUTS["vn"],
        results_data={
            "bui_3kenh": [
            {"cap": "ISO 8", "gioi_han": ISO8_LIMITS, "trang_so_do": TRANG_SO_DO, "diem": diem(ROOMS[:1])},
            {"cap": "ISO 8", "gioi_han": ISO8_LIMITS, "trang_so_do": TRANG_SO_DO, "diem": diem(ROOMS[1:])},
        ]},
        assessment={"items_list_vn": ["nồng độ hạt bụi"]},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


def _finalize_kn_dust_only(path: str, gas_temp: float):
    """Phiếu khí nén chỉ đo bụi bằng AeroTrak: tổng dầu, điểm sương, khung Lasair không đo -> gạch chéo (AGENTS 8B,
    như phiếu tủ); kết quả trỏ sang khối AeroTrak (dòng 68); N* theo nhiệt độ khí lấy mẫu trên phiếu ghi chép."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Tinh_Toan_Khi_Co_Ap_V6"]
    ws._images = [im for im in ws._images if im.anchor._from.row < 84]      # bỏ ảnh HumCalc / ảnh điểm sương
    ws["C36"].value = gas_temp
    ws["B36"].value = "Nhiệt độ khí lấy mẫu T (oC):"
    ws["C4"].value = "- Kích thước 0,3 - 0,5 µm"
    for cell, ref in (("D4", "C68"), ("D5", "D68"), ("D6", "E68"), ("D7", "F68")):
        ws[cell].value = f"={ref}"
    for cell in ("D8", "F8", "H8", "I8", "D9", "F9", "H9", "I9", "C85", "I43", "J43", "I44", "J44", "C52"):
        ws[cell].value = "-"
    strike_box(ws, "C43:D49")         # tổng dầu
    strike_box(ws, "C54:J59")         # khối Lasair III 110
    strike_box(ws, "C87:J89")         # điểm sương
    strike_box(ws, "B91:J111")        # ảnh màn hình DP510 + HumCalc
    if "DiemSuong" in wb.sheetnames:  # sheet phụ điểm sương không dùng -> ẩn (PDF không in trang trắng)
        wb["DiemSuong"].sheet_state = "hidden"
    thin_outside_borders(wb)
    wb.save(path)
    wb.close()
    try:                              # openpyxl làm mất giá trị đã tính -> tính lại (AGENTS mục 11)
        import win32com.client
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            x = xl.Workbooks.Open(os.path.abspath(path))
            xl.CalculateFull()
            x.Save()
            x.Close()
        finally:
            xl.Quit()
    except Exception as e:
        print(f"[build_137035] Excel COM warning: {e}")


def build_kn_workbooks() -> list:
    """Phiếu tính toán khí nén từng vị trí (template TinhToanKhiNen, giữ công thức) -> ThucHien."""
    out_dir = os.path.join(INPUT_DIR, "ThucHien")
    outs = []
    for p in KN_POINTS:
        t = p["gas_temp"]
        point_data = {
            "code": CODE, "point_label": p["sym"], "cert_date": CERT_DATE,
            "location_name": p["name_vn"], "location_symbol": p["sym"],
            "location_photo": os.path.join(PHOTO_KN, p["photo"]),
            "requirements": {"dust": REQ_DUST},
            # điểm sương không đo: chỉ nhiệt độ khí (để N* khớp phiếu ghi chép); khung bị gạch chéo sau đó
            "dewpoint": {"temp_start": t, "temp_end": t},
            "dust": {"device": "AeroTrak 9350", "raw_channels": p["dust"], "mode": "diff",  # cột Diff phiếu in
                     "dilution_factor": 1.0,
                     "customer_limit_05": REQ_DUST},
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
        equipments=[eq("aerotrak_9350_hpd3")],
        scope_vn="Lấy mẫu, phân tích chất lượng khí nén", scope_en="Sampling and analysis of compressed air quality",
        standards_vn=("ISO 8573-1:2010;\nMức yêu cầu đối với khí nén cung cấp bởi Công ty TNHH Nextern Việt Nam."),
        standards_en=("ISO 8573-1:2010;\nRequirements for compressed air provided by Nextern Vietnam Limited Liability Company."),
    )
    d.reg_date = REG_DATE
    d.procedure_vn = d.procedure_en = "QUATEST 3 3030:2025, N2.TC.KN.01"
    d.location_desc_vn, d.location_desc_en = LOC_KN_VN, LOC_KN_EN
    opt = os.path.join(PHOTO_DIR, "_optimized")
    d.photos = []
    for loc, p in zip(d.locations, KN_POINTS):
        loc.name_vn, loc.name_en = p["name_vn"], p["name_en"]
        d.photos.append({"path": _optimized(os.path.join(PHOTO_KN, p["photo"]), opt),
                         "caption_vn": f"Vị trí lấy mẫu {p['sym']}", "caption_en": f"Sampling location {p['sym']}"})
    rd = d.results_data
    rd["oil"] = {"rows": []}          # không đo tổng dầu, điểm sương (BBLM gạch bỏ)
    rd["dewpoint"] = {"rows": []}
    rd["system_chain_vn"], rd["system_chain_en"] = SYSTEM_CHAIN_VN, SYSTEM_CHAIN_EN
    rd["kqtn"] = []
    rd["bckq_label"] = "khi nen"
    d.input_dir, d.output_dir, d.network_dir = INPUT_DIR, "", ""
    d.thuchien_dir = os.path.join(INPUT_DIR, "ThucHien")
    d.tbkq_dir = os.path.join(INPUT_DIR, "TBKQ")
    d.edoc_dir = os.path.join(d.tbkq_dir, "Edoc")
    d.pdf_dir = os.path.join(INPUT_DIR, "pdf")
    return d


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX (như build_133266)
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
    """2 dòng cuối bảng kết quả + bảng Ghi chú + câu Đánh giá cùng trang (AGENTS mục 3; E32)."""
    def keep(tr_list):
        for tr in tr_list:
            for p_el in tr.iter(qn("w:p")):
                pPr = p_el.get_or_add_pPr()
                if pPr.find(qn("w:keepNext")) is None:
                    pPr.insert(0, OxmlElement("w:keepNext"))
    for tbl in list(doc.element.body.iterchildren(qn("w:tbl"))):
        note, gaps = tbl.getnext(), []
        while note is not None and note.tag == qn("w:p") and not "".join(t.text or "" for t in note.iter(qn("w:t"))).strip():
            gaps.append(note)
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
    """Đoạn trống cuối văn bản -> bỏ, kể cả các đoạn trống nằm NGAY TRƯỚC đoạn chứa đường kết thúc (AGENTS mục 3:
    không để đoạn trống trước đường kết thúc — 137035: 8 đoạn trống đẩy đường kết thúc sang trang 5 trắng)."""
    def blank(el):
        return el.tag == qn("w:p") and not "".join(t.text or "" for t in el.iter(qn("w:t"))).strip() \
            and el.find(f".//{qn('w:drawing')}") is None and "AlternateContent" not in "".join(
                c.tag for c in el.iter()) and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is None
    for el in reversed(list(doc.element.body.iterchildren(qn("w:p")))):
        if blank(el):
            el.getparent().remove(el)
            continue
        if "AlternateContent" in "".join(c.tag for c in el.iter()):   # đoạn đường kết thúc: bỏ đoạn trống phía trước
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


def _aerotrak_band_headers(doc):
    """Bảng hạt khí nén: AeroTrak đo từ 0,3 µm -> tiêu đề cột '0,1 µm < d ≤ 0,5 µm' đổi thành 0,3 µm."""
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        runs = [t for t in p.iter(qn("w:t"))]
        txt = "".join(t.text or "" for t in runs)
        if ("d ≤ 0,5" in txt or "d ≤ 0.5" in txt) and ("0,1" in txt or "0.1" in txt) and len(txt) < 60:
            for t in runs:
                if t.text and ("0,1" in t.text or "0.1" in t.text):
                    t.text = t.text.replace("0,1", "0,3").replace("0.1", "0.3")
                    n += 1
                    break
    n += _replace_in_paragraphs(doc, {"(0,1-5,0) µm": "(0,3-5,0) µm", "(0.1-5.0) µm": "(0.3-5.0) µm"})
    return n


def _replace_in_paragraphs(doc, mapping: dict) -> int:
    """Thay chuỗi trong từng đoạn (kể cả khi chuỗi nằm vắt qua nhiều run): dồn chữ của đoạn vào run chứa đầu chuỗi."""
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        txt = "".join(t.text or "" for t in ts)
        if not any(a in txt for a in mapping):
            continue
        for a, b in mapping.items():
            while a in txt:
                start = txt.index(a)
                pos, first = 0, None
                for t in ts:                      # run chứa ký tự đầu của chuỗi cần thay
                    L_ = len(t.text or "")
                    if first is None and pos + L_ > start:
                        first = t
                        break
                    pos += L_
                new = txt.replace(a, b, 1)
                # đặt lại: các run trước giữ nguyên, run 'first' nhận phần còn lại, các run sau xóa chữ
                before = sum(len(t.text or "") for t in ts[:ts.index(first)])
                first.text = new[before:]
                first.set(qn("xml:space"), "preserve")
                for t in ts[ts.index(first) + 1:]:
                    t.text = ""
                txt = new
                n += 1
    return n


def finish_docx(src: str, dst_name: str, cleanroom: bool = False) -> str:
    """Hậu xử lý chung rồi đổi tên theo AGENTS mục 1 (handler đặt 'TBKQ - vn - n - …', E41)."""
    doc = Document(src)
    _fix_cover_lists(doc)
    _keep_headings_with_tables(doc)
    if cleanroom:
        _keep_tails_together(doc)
        _drop_trailing_empty(doc)
        _appendix_total_from_numpages(doc)
    else:
        print(f"[build_137035] tiêu đề cột 0,3 µm: {_aerotrak_band_headers(doc)} chỗ")
        keep_result_heading_with_next(doc)    # '9.2. Kết quả giám định' đi liền 9.2.1 (đã soát VN = EN)
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    dst = os.path.join(os.path.dirname(src), dst_name)
    from Core.Builders.tag_context import style_note_tables
    style_note_tables(doc)      # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3, 09/11/2026)
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    return dst


def build_bckq_cleanroom(d: Dossier, so_phieu: str) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["tham_chieu"] = [f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số "
                         f"{so_phieu} ngày {SAMPLE_DATE}."]
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
    # bụi khí nén tính trong phiếu từng vị trí (không có phiếu bụi riêng) -> viện dẫn đúng số phiếu (AGENTS mục 7)
    kn_refs = ", ".join(f"{CODE} - {p['sym']}" for p in KN_POINTS)
    _replace_in_paragraphs(doc, {
        f"phiếu Tính toán kết quả nồng độ hạt bụi số {CODE} ngày": f"phiếu Kết quả tính toán số {kn_refs} ngày",
        "(0,1-5,0) µm": "(0,3-5,0) µm"})
    _keep_headings_with_tables(doc)
    _highlight_placeholders(doc)
    doc.save(out)
    return out


SO_PHIEU_BUI = f"{CODE}/Bụi"      # số phiếu tính toán = mã + đuôi chỉ tiêu (AGENTS mục 7)


def build_excel_cleanroom() -> str:
    """Phiếu tính toán /1 theo form phòng sạch chung (Core/Tools/cleanroom_calc.py): sheet 'Bụi 0,3' = bảng AeroTrak."""
    lim = dict(zip(("Ch2", "Ch3", "Ch6"), ISO8_LIMITS))
    book = CleanroomCalcBook()
    book.dust_aerotrak([{"room": vn, "cap": "ISO 8", "state": "Hoạt động", "limits": lim,
                         "zero": RECS[z] if z else None,
                         "labels": [n for n, _ in pts], "points": [RECS[f] for _, f in pts]}
                        for vn, _, _, z, pts in ROOMS], so=SO_PHIEU_BUI, date=SAMPLE_DATE, labels=LABELS)
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - 1 - TinhToanPhongSach.xlsx")
    book.save(out)
    print(f"[build_137035] Excel -> {out}")
    return out


def split_edoc() -> list:
    """BBLM = trang 1–2; Ghi chép = trang 3–7 + phiếu in AeroTrak khí nén + bảng AeroTrak phòng sạch (tr. 8 trắng, bỏ)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(SCAN)
    bblm = fitz.open()
    bblm.insert_pdf(src, from_page=0, to_page=1)
    gc = fitz.open()
    gc.insert_pdf(src, from_page=2, to_page=6)
    gc.insert_pdf(fitz.open(SCAN_KN))
    gc.insert_pdf(fitz.open(SCAN_PS))
    outs = []
    for name, pdf in ((f"{CODE6} - BBLM.pdf", bblm), (f"{CODE6} - Ghi Chep.pdf", gc)):
        p = os.path.join(edoc, name)
        finish_edoc(pdf, INPUT_DIR, name)
        pdf.save(p)
        outs.append(p)
        print(f"[build_137035] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()

    kn_xlsx = build_kn_workbooks()
    ps = cleanroom_dossier()
    kn = compressed_air_dossier(kn_xlsx)
    for p, pt in zip(kn_xlsx, KN_POINTS):        # gạch khung không đo SAU khi đã đọc số liệu vào dossier
        _finalize_kn_dust_only(p, pt["gas_temp"])
    assign_cert_numbers([ps, kn])
    for d in (ps, kn):
        d.ensure_output_directories()
    split_edoc()

    h = CleanroomHandler(ps)
    ps_vn = finish_docx(h.build_tbkq(is_en=False), f"TBKQ - viet - 1 - phong sach - {CODE}.docx", cleanroom=True)
    ps.layouts = LAYOUTS["en"]
    ps_en = finish_docx(h.build_tbkq(is_en=True), f"TBKQ - en - 1 - phong sach - {CODE}.docx", cleanroom=True)
    ps.layouts = LAYOUTS["vn"]
    ps_bckq = build_bckq_cleanroom(ps, SO_PHIEU_BUI)
    ps_xlsx = build_excel_cleanroom()

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
