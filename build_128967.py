# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D128967 - Chi nhánh Công ty Cổ phần Công nghiệp – Dịch vụ - Thương mại Ngọc Nghĩa – Nhà máy nhựa số 2
(06/11/2026) — khí nén, 2 chứng thư (như năm 2025, hồ sơ 00914/N2.25/DG):
  /1 KN1 (hệ khí nén thấp áp), /2 KN2 (hệ khí nén cao áp, khu vực máy Sidel) — mỗi vị trí: nồng độ hạt (AeroTrak 9350
  + HPD III), điểm sương (DP510 + CS16), tổng dầu (lấy mẫu), vi sinh khí nén (Pinocchio Super II).

Nguồn (Input/128967 — chép từ thư mục hồ sơ trên Z 16/11/2026, chỉ đọc):
  0.  BBLM  .pdf — BM3030.02 2 trang: PYC Q326D128967 ngày 12/10/2026; ngày 06/11/2026 từ 8h45 đến 14h15; nội dung
      "Lấy mẫu/ phân tích chất lượng khí nén"; địa điểm "tại địa chỉ nêu tại mục 3"; thiết bị Aerotrak 9350 + HPD III,
      giấy lọc Q3400 Sartorius + ORBO 32 + PS Cassettes 47 mm, DP 510 + CS16, Pinocchio Super II; quy trình N2.TC.KN.01;
      nhà máy hoạt động sản xuất; KN1 "Ống dẫn khí nén sau khi qua lọc tại KV máy nén (hệ khí nén thấp áp)" (viết tay),
      KN2 "Ống dẫn khí nén khu vực máy Sidel (khí nén hệ cao áp)" (viết tay); mỗi vị trí đo nồng độ hạt, điểm sương,
      lấy mẫu tổng dầu, vi sinh. Tr. 2 viết tay: KN1 cấp từ 02 máy nén Atlas Copco có dùng dầu bôi trơn; KN2 từ 03 máy
      nén AF Kushan / AF Francois / Gardner Denver có dùng dầu bôi trơn. Người thực hiện Phạm Quốc Nghĩa.
  THUC HIEN -/D128967 - DO BUI/128967 .pdf — phiếu in AeroTrak 9350-02 S/N 93502239001, Count Mode Δ, # / m³:
      zero 1254, 1256, 1257; đo 1255 (09:39, 20 phút, 1000 L) = KN1, 1258 (10:45) = KN2 (theo giờ đo điểm sương).
  THUC HIEN -/D261106 - DIEM SUONG — màn hình DP510: KN1 BM00004 - DO 1, BM00013 - DO 2; KN2 BM00017 - DO 1,
      BM00022 - DO 2. Số đọc lấy theo phiếu người dùng (sheet DiemSuong).
  1.  TinhToanKhiNen - KN1/KN2 .xlsx — phiếu người dùng: điểm sương KN1 2,30/1,93 °C, khí 35,6/37,4 °C, 7,4/7,4 bar,
      09h22/09h48, quy đổi 7 bar 1,42; KN2 −7,61/−6,89 °C, 31,1/30,9 °C, 5,37/5,35 bar, 10h26/10h52, quy đổi −4,71;
      AeroTrak KN1/KN2 như phiếu in Δ; khối Lasair, dầu "< 10"/"< 5" là số còn lại của form (không có KQTN dầu).
  THUC HIEN -/KQTN - VI SINH — VSTN26028699.01 (KN1), .02 (KN2), .03 (KN3 BLANK), ngày 13/11/2026, nhận mẫu
      06/11/2026: tổng số nấm men, nấm mốc < 1 CFU/đĩa; tổng số VSV hiếu khí < 1 CFU/đĩa.
  THUC HIEN -/914 - TBKQ - V - KN1/KN2 .docx — TBKQ 2025 cùng khách: tên vị trí, cách trình bày (không lấy số liệu).
  0.  KN1 / KN2 .jpg — ảnh vị trí (người dùng đặt sẵn ở gốc thư mục hồ sơ).

Quyết định của người dùng (16/11/2026): giấy lọc dầu "Giữ Pall" -> eq("oil_kit") (BBLM ghi Q3400 Sartorius).

Suy luận của AI (cần người dùng xác nhận):
  - 2 chứng thư /1 KN1, /2 KN2, VN + EN như 2025 (không có PYC trong thư mục).
  - Tên vị trí theo BBLM 2026 viết lại theo cách ghi 2025 ("khí nén từ hệ thống khí nén thấp/cao áp"); KN1 chữ viết tay
    "KV máy nén" đọc là "khu vực máy nén" (2025 ghi "sau khi qua lọc Atlas Copco QD300+").
  - Tổng dầu: chưa có KQTN (MTTN / HDTN) -> kết quả, phân loại, ngày ban hành để ……… tô vàng.
  - Vi sinh: KQTN gộp "Tổng số nấm men, nấm mốc" -> TBKQ 2 dòng (tổng số VSV hiếu khí; tổng số nấm men, nấm mốc),
    2025 ghi 3 dòng.
  - Tên tiếng Anh khách hàng: AI dịch (chưa có MST / GCN).
  - Ảnh: 1 ảnh/vị trí (0. KN1.jpg, 0. KN2.jpg).
  (giấy lọc dầu: xem Quyết định)
  - Ngày yêu cầu cấp KQTN = ngày lấy mẫu + 8 = 14/11/2026 (không có PYC).

Chạy:  python build_128967.py          (DOCX + XLSX + Edoc)
       python build_128967.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: trong code (network_dir để trống) — chép lên Z theo AGENTS mục 0.
"""

import os
import re
import sys
import copy
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

from Core.Models import ClientInfo, PersonnelInfo, Equipment, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.compressed_air import CompressedAirHandler
from Core.Builders.tag_context import PLACEHOLDER, style_note_tables
from Core.Styling.tables import add_tbl_header
from Core.Tools.khinen_calc import build_compressed_air_workbook, create_dossier_from_compressed_air_workbooks
from Core.Tools.excel_borders import thin_outside_borders, strike_box
from Core.Tools.kqtn_late import tre_han_text
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D128967"
CODE6 = "128967"
REG_DATE = "12/10/2026"            # BBLM mục 1 (PYC)
SAMPLE_DATE = "06/11/2026"         # BBLM
START, END = "08:45", "14:15"      # BBLM mục 6
CERT_DATE_1 = CERT_DATE_2 = date.today().strftime("%d/%m/%Y")   # người dùng 16/11/2026: ngày ban hành = ngày dựng (kể cả khi chờ KQTN)
NGAY_YC_KQTN = "14/11/2026"        # TẠM: ngày lấy mẫu + 8 ngày (không có PYC — AGENTS mục 13)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SRC_TH = os.path.join(INPUT_DIR, "THUC HIEN -")
SCAN_BBLM = os.path.join(INPUT_DIR, "0.  BBLM  .pdf")
DEW_DIR = os.path.join(SRC_TH, "D261106 - DIEM SUONG")
PHOTO_OPT = os.path.join(SRC_TH, "D128967 - PHOTO", "_optimized")

CLIENT = ClientInfo(
    name_vn="Chi nhánh Công ty Cổ phần Công nghiệp – Dịch vụ - Thương mại Ngọc Nghĩa – Nhà máy nhựa số 2",   # BBLM
    name_en="Branch of Ngoc Nghia Industry – Service – Trading Joint Stock Company – Plastic Factory No. 2",  # AI dịch
    address_vn="Lô B1-8, B1-9 Khu công nghiệp Tây Bắc Củ Chi, xã Tân An Hội, Thành phố Hồ Chí Minh",          # BBLM
    address_en="Lot B1-8, B1-9 Tay Bac Cu Chi Industrial Park, Tan An Hoi commune, Ho Chi Minh City",
)
LOCATION_VN = f"{CLIENT.name_vn} - {CLIENT.address_vn}"     # BBLM mục 5: "tại địa chỉ nêu tại mục 3"
LOCATION_EN = f"{CLIENT.name_en} - {CLIENT.address_en}"
EXECUTOR = "Nghĩa"                 # BBLM: Phạm Quốc Nghĩa
OIL_KIT_VN = "Giấy lọc Q3400 Sartorius, ống than hoạt tính ORBO® 32, PS Cassettes ϕ 47 mm"     # BBLM mục 7
OIL_KIT_EN = "Q3400 Sartorius filter paper, ORBO® 32 activated carbon tube, PS Cassettes ϕ 47 mm"

POINTS = [
    {"sym": "KN1",
     "name_vn": "Ống dẫn khí nén sau khi qua lọc tại khu vực máy nén (khí nén từ hệ thống khí nén thấp áp)",
     "name_en": "Compressed air pipeline after the filter at the compressor area (compressed air from the "
                "low-pressure compressed air system)",
     "dust": [1086112, 64406, 11010, 9257, 179, 7],          # AeroTrak #1255, Δ #/m³ (0,3 0,5 1 2 3 5 µm)
     "expect": [1147275, 68033, 21597, 7],                   # phiếu người dùng D63:F63
     "dew": {"dewp_start": 2.30, "dewp_end": 1.93, "temp_start": 35.6, "temp_end": 37.4,
             "pres_start": 7.4, "pres_end": 7.4, "time_start": "Lúc 09h22", "time_end": "Lúc 09h48"},
     "bmp": ("KN1 -/BM00004 - DO 1  .bmp", "KN1 -/BM00013 - DO 2.bmp"),
     "t2_user": 1.42, "vs": "VSTN26028699.01",
     "photo": os.path.join(INPUT_DIR, "0.  KN1  .jpg"),
     "chain_vn": ["02 máy nén khí Atlas Copco (có dùng dầu bôi trơn)", "Hệ thống khí nén thấp áp", "Lọc",
                  "Ống dẫn khí nén (KN1)"],
     "chain_en": ["02 Atlas Copco air compressors (oil-lubricated)", "Low-pressure compressed air system", "Filter",
                  "Compressed air pipeline (KN1)"]},
    {"sym": "KN2",
     "name_vn": "Ống dẫn khí nén tại khu vực máy Sidel (khí nén từ hệ thống khí nén cao áp)",
     "name_en": "Compressed air pipeline at Sidel machine area (compressed air from the high-pressure compressed air system)",
     "dust": [119695, 74646, 23235, 15611, 197, 4],           # AeroTrak #1258
     "expect": [124189, 77448, 40509, 4],
     "dew": {"dewp_start": -7.61, "dewp_end": -6.89, "temp_start": 31.1, "temp_end": 30.9,
             "pres_start": 5.37, "pres_end": 5.35, "time_start": "Lúc 10h26", "time_end": "Lúc 10h52"},
     "bmp": ("KN2 -/BM00017 - DO 1  .bmp", "KN2 -/BM00022 - DO 2 .bmp"),
     "t2_user": -4.71, "vs": "VSTN26028699.02",
     "photo": os.path.join(INPUT_DIR, "0.  KN2 .jpg"),
     "chain_vn": ["03 máy nén khí AF Kushan / AF Francois / Gardner Denver (có dùng dầu bôi trơn)",
                  "Hệ thống khí nén cao áp", "Ống dẫn khí nén khu vực máy Sidel (KN2)"],
     "chain_en": ["03 air compressors AF Kushan / AF Francois / Gardner Denver (oil-lubricated)",
                  "High-pressure compressed air system", "Compressed air pipeline at Sidel machine area (KN2)"]},
]
KQTN_VS = {"so": "VSTN26028699", "blank": "VSTN26028699.03", "ngay": "13/11/2026", "ngay_nhan_mau": "06/11/2026"}
VS_HOLES, VS_LIT = 218, 500                 # Pinocchio Super II — phiếu người dùng
# (chỉ tiêu VN, chỉ tiêu EN, KQTN, kết quả TBKQ VN, EN) — VSTN26028699
MICRO = [("Tổng số vi sinh vật hiếu khí", "Total aerobic microorganisms", "< 1 CFU/đĩa", "Không có chỉ thị", "No indication"),
         ("Tổng số nấm men, nấm mốc", "Total yeasts and molds", "< 1 CFU/đĩa", "Không có chỉ thị", "No indication")]


def _jpg(src: str, out_dir: str, max_w: int = 1200) -> str:
    """Ảnh -> JPEG trong thư mục _optimized (ngoại lệ được phép, AGENTS mục 1); ảnh dọc xoay ngang (mục 3)."""
    from PIL import Image, ImageOps
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, re.sub(r"\s+", "_", os.path.splitext(os.path.basename(src))[0].strip()) + ".jpg")
    im = ImageOps.exif_transpose(Image.open(src))
    if im.height > im.width:
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=90)
    return out

# ---------------------------------------------------------------------------
# Tiện ích DOCX (theo build_134179)
# ---------------------------------------------------------------------------
def _ptext(el):
    import unicodedata
    return unicodedata.normalize("NFC", "".join(t.text or "" for t in el.iter(qn("w:t"))))


def _keep(p_el):
    pPr = p_el.get_or_add_pPr()
    if pPr.find(qn("w:keepNext")) is None:
        pPr.insert(0, OxmlElement("w:keepNext"))


def _keep_rows(rows):
    for tr in rows:
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)


def _is_note_table(el) -> bool:
    return el is not None and el.tag == qn("w:tbl") and _ptext(el).strip().startswith(("Ghi chú", "Note"))


def _set_text(p_el, text: str, superscript_tail: str = ""):
    """Đặt lại chữ của đoạn, giữ định dạng run đầu; `superscript_tail` in chỉ số trên (vd '(1)')."""
    runs = p_el.findall(qn("w:r"))
    proto = copy.deepcopy(runs[0])
    for r in runs:
        p_el.remove(r)
    for ch in [c for c in p_el if c.tag != qn("w:pPr")]:
        p_el.remove(ch)

    def run(txt, sup=False):
        r = copy.deepcopy(proto)
        for c in [c for c in r if c.tag != qn("w:rPr")]:
            r.remove(c)
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        for va in rpr.findall(qn("w:vertAlign")):
            rpr.remove(va)
        if sup:
            va = OxmlElement("w:vertAlign")
            va.set(qn("w:val"), "superscript")
            rpr.append(va)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = txt
        r.append(t)
        p_el.append(r)
    run(text)
    if superscript_tail:
        run(superscript_tail, True)


def _set_cell(tc, text: str, superscript_tail: str = ""):
    ps = tc.findall(qn("w:p"))
    base = next((p for p in ps if p.findall(qn("w:r"))), ps[0])
    for p in ps:
        if p is not base:
            tc.remove(p)
    _set_text(base, text, superscript_tail)


def _vmerge(tc, restart: bool):
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:vMerge")):
        tcPr.remove(old)
    vm = OxmlElement("w:vMerge")
    if restart:
        vm.set(qn("w:val"), "restart")
    tcPr.append(vm)


def note_rows(proto, lines):
    """Bảng 'Ghi chú' chép từ bảng mẫu: dòng tiêu đề + mỗi ghi chú một dòng (marker | nội dung)."""
    el = copy.deepcopy(proto)
    rows = el.findall(qn("w:tr"))
    item = rows[-1]
    for r in rows[1:]:
        el.remove(r)
    for marker, text in lines:
        tr = copy.deepcopy(item)
        el.append(tr)
        tcs = tr.findall(qn("w:tc"))
        _set_cell(tcs[0], marker)
        _set_cell(tcs[1], text)
    return el


def _prev_text_p(el):
    prev = el.getprevious()
    while prev is not None and not (prev.tag == qn("w:p") and _ptext(prev).strip()):
        prev = prev.getprevious()
    return prev


def _replace_in_paragraphs(root, mapping: dict) -> int:
    """Thay chuỗi trong từng đoạn (kể cả khi chuỗi vắt qua nhiều run)."""
    n = 0
    for p in root.iter(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        txt = "".join(t.text or "" for t in ts)
        if not any(a in txt for a in mapping):
            continue
        for a, b in mapping.items():
            while a in txt:
                start = txt.index(a)
                pos, first = 0, None
                for t in ts:
                    L_ = len(t.text or "")
                    if pos + L_ > start:
                        first = t
                        break
                    pos += L_
                new = txt.replace(a, b, 1)
                before = sum(len(t.text or "") for t in ts[:ts.index(first)])
                first.text = new[before:]
                first.set(qn("xml:space"), "preserve")
                for t in ts[ts.index(first) + 1:]:
                    t.text = ""
                txt = new
                n += 1
    return n


def _highlight_placeholders(doc):
    for r in list(doc.element.body.iter(qn("w:r"))):
        if not any(PLACEHOLDER in (t.text or "") for t in r.findall(qn("w:t"))):
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


def _keep_headings_with_tables(doc):
    for el in doc.element.body.iterchildren(qn("w:tbl")):
        prev = el.getprevious()
        if prev is not None and prev.tag == qn("w:p") and _ptext(prev).strip():
            _keep(prev)
            rows = el.findall(qn("w:tr"))
            _keep_rows(rows[:min(5, len(rows) - 1)])


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


# ---------------------------------------------------------------------------
# Phiếu tính toán khí nén (ThucHien) — template TinhToanKhiNen, giữ công thức (AGENTS mục 11)
# ---------------------------------------------------------------------------
def build_workbooks() -> dict:
    out = {}
    for p in POINTS:
        dew = dict(p["dew"])
        dew["photo_start"] = _jpg(os.path.join(DEW_DIR, p["bmp"][0]), os.path.join(DEW_DIR, "_optimized"))
        dew["photo_end"] = _jpg(os.path.join(DEW_DIR, p["bmp"][1]), os.path.join(DEW_DIR, "_optimized"))
        point_data = {
            "code": CODE, "point_label": p["sym"], "cert_date": CERT_DATE_1,
            "location_name": p["name_vn"], "location_symbol": p["sym"],
            "location_photo": _jpg(p["photo"], PHOTO_OPT),
            "dewpoint": dew,
            # Hiệu chỉnh DP510: công tắc chung AP_HIEU_CHINH đang tắt (AGENTS 8B) -> không khai "cals".
            "dust": {"device": "AeroTrak 9350", "raw_channels": p["dust"], "mode": "diff", "dilution_factor": 1.0},
            # tổng dầu: chưa có KQTN -> chỉ lưu lượng / thời gian lấy mẫu (như phiếu người dùng)
            "oil": {"flow_aerosol": 35, "flow_vapor": 1, "time_aerosol": 240, "time_vapor": 240},
        }
        path = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - {p['sym']} - TinhToanKhiNen.xlsx")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        build_compressed_air_workbook(point_data, path)
        out[p["sym"]] = path
    return out


def prepare_workbook(path: str, p: dict):
    """Trước khi đọc vào dossier: kết quả bụi trỏ khối AeroTrak (dòng 68, như 137035), cột 0,3 µm; khung Lasair gạch
    chéo; vi sinh < 1 CFU/đĩa (KQTN); tổng dầu chưa có KQTN -> ………. Tính lại bằng Excel COM, đối chiếu số hạt."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Tinh_Toan_Khi_Co_Ap_V6"]
    ws["C4"].value = "- Kích thước 0,3 - 0,5 µm"
    for cell, ref in (("D4", "C68"), ("D5", "D68"), ("D6", "E68"), ("D7", "F68")):
        ws[cell].value = f"={ref}"
    strike_box(ws, "C54:J59")                  # khối Lasair III 110 (không dùng)
    ws["C46"].value = ws["D46"].value = PLACEHOLDER        # KQTN dầu hạt / dầu hơi chưa có
    # Chưa có KQTN dầu: các ô dầu ra ……… (người dùng 16/11/2026: "Có, sửa luôn") — giữ công thức, chỉ bọc điều kiện
    # chờ KQTN; số phiếu MTTN/HDTN cũ của form xóa thành ………. Có KQTN thì bỏ khối này (phiếu dựng lại từ template).
    wait = f'OR(C46="{PLACEHOLDER}",D46="{PLACEHOLDER}")'
    for ref in ("C47", "D47", "C48", "C49", "F9"):
        f = str(ws[ref].value or "")
        if f.startswith("="):
            ws[ref].value = f'=IF({wait},"{PLACEHOLDER}",{f[1:]})'
    for ref in ("I43", "J43", "I44", "J44", "H9"):
        ws[ref].value = PLACEHOLDER
    ws["C77"].value, ws["C78"].value = VS_HOLES, VS_LIT
    ws["C81"].value = ws["D81"].value = ws["E81"].value = "< 1"     # tổng số; nấm men, nấm mốc (KQTN gộp)
    ws["I77"].value = p["vs"]
    from openpyxl.styles import Alignment
    a = ws["H2"].alignment
    ws["H2"].alignment = Alignment(horizontal=a.horizontal, vertical=a.vertical, shrink_to_fit=True, wrap_text=True)
    # Ảnh vị trí (neo B17) cao tới dòng 33 nhưng Excel ngắt trang trước dòng 30 -> ảnh tràn 2 trang (người dùng báo
    # 16/11/2026). Thu ảnh (giữ tỉ lệ, không crop) cho vừa dòng 17–29: 13 dòng × 16pt ≈ 208pt -> cao 200pt.
    for im in ws._images:
        if im.anchor._from.row == 16:
            h_px = round(200 * 4 / 3)
            im.width, im.height = round(im.width * h_px / im.height), h_px
            from openpyxl.utils.units import pixels_to_EMU
            if getattr(im.anchor, "ext", None) is not None:
                im.anchor.ext.width, im.anchor.ext.height = pixels_to_EMU(im.width), pixels_to_EMU(im.height)
    if "DiemSuong" in wb.sheetnames:
        wb["DiemSuong"].sheet_state = "hidden"
    thin_outside_borders(wb)
    wb.save(path)
    wb.close()
    try:
        import win32com.client
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            x = xl.Workbooks.Open(os.path.abspath(path))
            xl.CalculateFull()
            s = x.Worksheets("Tinh_Toan_Khi_Co_Ap_V6")
            got = [s.Range(a).Value for a in ("D4", "D5", "D6", "D7")]
            t2 = s.Range("G87").Value
            x.Save()
            x.Close()
        finally:
            xl.Quit()
        if [round(v) for v in got] != p["expect"]:
            print(f"[WARNING] [Data Mismatch] bụi {p['sym']}: phiếu {got} khác phiếu người dùng {p['expect']}")
        else:
            print(f"[build_128967] phiếu {p['sym']}: bụi {got} = phiếu người dùng")
        if abs(round(t2, 2) - p["t2_user"]) > 0.01:
            print(f"[WARNING] [Data Mismatch] điểm sương 7 bar {p['sym']}: code {t2:.2f}, phiếu người dùng {p['t2_user']}")
    except ImportError as e:
        print(f"[build_128967] Excel COM warning: {e}")


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def make_dossier(xlsx: str, p: dict, sub_no: int, cert_date: str):
    equip = [eq("aerotrak_9350_hpd3"), eq("cs_dp510"), eq("oil_kit"),
             eq("sas_pinocchio")]
    d = create_dossier_from_compressed_air_workbooks(
        [xlsx], code=CODE, cert_date=cert_date, sample_date=SAMPLE_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=sub_no, client=CLIENT, personnel=PersonnelInfo(executor=EXECUTOR), equipments=equip,
        scope_vn="Lấy mẫu/ phân tích chất lượng khí nén", scope_en="Sampling/ analyzing compressed air quality",
        standards_vn="ISO 8573-1:2010", standards_en="ISO 8573-1:2010",
        basis_vn="ISO 8573-1:2010", basis_en="ISO 8573-1:2010",
    )
    d.reg_date = REG_DATE
    d.procedure_vn = d.procedure_en = "QUATEST 3 3030:2025, N2.TC.KN.01"      # BBLM mục 8
    d.location_desc_vn, d.location_desc_en = LOCATION_VN, LOCATION_EN
    loc = d.locations[0]
    loc.name_vn, loc.name_en = p["name_vn"], p["name_en"]
    d.photos = [{"path": _jpg(p["photo"], PHOTO_OPT), "caption_vn": "", "caption_en": ""}]
    rd = d.results_data
    rd["system_status_vn"], rd["factory_status_vn"] = "Đang hoạt động", "Đang hoạt động sản xuất"   # BBLM mục 9
    rd["system_status_en"], rd["factory_status_en"] = "In operation", "In operation"
    rd["system_chain_vn"], rd["system_chain_en"] = p["chain_vn"], p["chain_en"]
    for r in rd["dewpoint"]["rows"]:          # hàm lượng nước ghi số nguyên (như 133266)
        r["values"][4] = str(round(float(r["values"][4].replace(",", "."))))
    for r in rd["oil"]["rows"]:               # chưa có KQTN tổng dầu
        r["val"], r["class"] = PLACEHOLDER, PLACEHOLDER
    rd["kqtn_by_location"] = {p["sym"]: f"- Dầu hơi: xem tại phiếu MTTN{PLACEHOLDER}\n- Dầu hạt: xem tại phiếu HDTN{PLACEHOLDER}"}
    rd["kqtn"] = [{"so": f"{KQTN_VS['so']} (.01 KN1, .02 KN2, .03 blank)", "ngay": KQTN_VS["ngay"],
                   "ngay_gui": KQTN_VS["ngay_nhan_mau"], "ngay_nhan": KQTN_VS["ngay"], "ngay_yc": NGAY_YC_KQTN,
                   "tre_han": tre_han_text(KQTN_VS["ngay"], KQTN_VS["ngay_nhan_mau"], NGAY_YC_KQTN, PLACEHOLDER)},
                  {"so": f"MTTN{PLACEHOLDER}", "ngay": PLACEHOLDER, "ngay_gui": PLACEHOLDER, "ngay_nhan": PLACEHOLDER,
                   "ngay_yc": NGAY_YC_KQTN, "tre_han": PLACEHOLDER},
                  {"so": f"HDTN{PLACEHOLDER}", "ngay": PLACEHOLDER, "ngay_gui": PLACEHOLDER, "ngay_nhan": PLACEHOLDER,
                   "ngay_yc": NGAY_YC_KQTN, "tre_han": PLACEHOLDER}]
    print(f"[build_128967] {p['sym']}: hạt {rd['particles']['rows'][0]['values']} > 5 µm {rd['particles']['rows'][0]['d_gt_50']}"
          f"; điểm sương {rd['dewpoint']['rows'][0]['values']}; dầu {rd['oil']['rows'][0] if rd['oil']['rows'] else '-'}")
    rd["bckq_label"] = "khi nen"
    d.input_dir, d.output_dir, d.network_dir = INPUT_DIR, "", ""
    d.thuchien_dir = os.path.join(INPUT_DIR, "ThucHien")
    d.tbkq_dir = os.path.join(INPUT_DIR, "TBKQ")
    d.edoc_dir = os.path.join(d.tbkq_dir, "Edoc")
    d.pdf_dir = os.path.join(INPUT_DIR, "pdf")
    return d


# ---------------------------------------------------------------------------
# TBKQ
# ---------------------------------------------------------------------------
def _drop_system_and_photo(doc, is_en):
    """BBLM: không thể hiện sơ đồ hệ thống khí nén, không chụp hình -> bỏ đoạn mô tả hệ thống (+ chuỗi) và mục ảnh."""
    body = doc.element.body
    keys = ("Description of compressed air system", "Diagram of compressed air system") if is_en else \
        ("Mô tả hệ thống khí nén", "Sơ đồ hệ thống khí nén")
    for p in list(body.iterchildren(qn("w:p"))):
        if _ptext(p).strip().startswith(keys):
            nxt = p.getnext()
            body.remove(p)
            if nxt is not None and nxt.tag == qn("w:p") and PLACEHOLDER in _ptext(nxt):
                body.remove(nxt)


def _add_micro(doc, is_en, sym):
    """Mục vi sinh khí nén (template chưa có — như build_134179): chép tiêu đề + bảng + Ghi chú của mục tổng dầu."""
    body = doc.element.body
    # dòng chỉ tiêu trong bảng vị trí
    last_param = "Concentration of particle sizes (0.1-5.0) µm." if is_en else "Nồng độ hạt bụi có kích thước (0,1-5,0) µm."
    for p in list(body.iter(qn("w:p"))):
        if _ptext(p).strip() == last_param:
            _set_text(p, last_param[:-1] + ";")
            extra = copy.deepcopy(p)
            p.addnext(extra)
            _set_text(extra, "Microbial contamination." if is_en else "Độ nhiễm bẩn vi sinh vật.")
    tbls = list(body.iterchildren(qn("w:tbl")))
    oil = next(t for t in tbls if ("Total oil content, mg/m" if is_en else "Hàm lượng tổng dầu, mg/m") in _ptext(t))
    oil_head = _prev_text_p(oil)
    between, nxt = [], oil.getnext()
    while nxt is not None and not _is_note_table(nxt):
        between.append(nxt)
        nxt = nxt.getnext()
    oil_note = nxt
    last_note = [t for t in tbls if _is_note_table(t)][-1]
    after_note, nxt = [], oil_note.getnext()
    while nxt is not None and nxt.tag == qn("w:p") and not _ptext(nxt).strip():
        after_note.append(nxt)
        nxt = nxt.getnext()

    anchor = last_note

    def put(el):
        nonlocal anchor
        anchor.addnext(el)
        anchor = el
        return el
    for g in after_note:
        put(copy.deepcopy(g))
    head = put(copy.deepcopy(oil_head))
    head_txt = _ptext(oil_head).strip()
    m = re.match(r"^(9\.2\.)(\d)\s", head_txt)
    title = "Microbial contamination" if is_en else "Nhiễm bẩn vi sinh vật"
    if m:   # EN đánh số cứng: số kế tiếp sau mục cuối
        heads = [re.match(r"^9\.2\.(\d)\s", _ptext(p).strip()) for p in body.iterchildren(qn("w:p"))]
        n = max(int(h.group(1)) for h in heads if h) + 1
        title = f"9.2.{n} {title}"
    _set_text(head, title)
    _keep(head)
    tbl = put(copy.deepcopy(oil))
    rows = tbl.findall(qn("w:tr"))
    data_proto = rows[1]
    for r in rows[2:]:
        tbl.remove(r)
    tcs = rows[0].findall(qn("w:tc"))
    _set_cell(tcs[0], "Location" if is_en else "Vị trí thực hiện")
    _set_cell(tcs[1], "Parameter" if is_en else "Chỉ tiêu")
    _set_cell(tcs[2], "Colony-forming units in compressed air, CFU/m³ " if is_en else
              "Số khuẩn lạc trong khí nén, CFU/m³ ", "(1)")
    for i, (vn, en, _kq, val_vn, val_en) in enumerate(MICRO):
        tr = data_proto if i == 0 else copy.deepcopy(data_proto)
        if i:
            tbl.append(tr)
        tcs = tr.findall(qn("w:tc"))
        _set_cell(tcs[0], sym if i == 0 else "")
        _vmerge(tcs[0], restart=(i == 0))
        _set_cell(tcs[1], en if is_en else vn)
        _set_cell(tcs[2], val_en if is_en else val_vn)
    _keep_rows(tbl.findall(qn("w:tr")))
    grid = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    total = sum(int(g.get(qn("w:w"))) for g in grid)
    widths = [round(total * w) for w in (0.18, 0.34, 0.48)]
    for g, w in zip(grid, widths):
        g.set(qn("w:w"), str(w))
    for tr in tbl.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), widths):
            tcW = tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tcW is not None:
                tcW.set(qn("w:w"), str(w))
                tcW.set(qn("w:type"), "dxa")
    for g in between:
        put(copy.deepcopy(g))
    lines = [
        ("(1) …", "Result converted to reference conditions of absolute pressure of 1 bar and temperature of 20 °C "
                  "according to ISO 8573-7:2003." if is_en else
                  "Kết quả quy về điều kiện áp suất tuyệt đối 1 bar, nhiệt độ 20 °C theo ISO 8573-7:2003."),
        ("●", "The result is expressed as “No indication” when no target organism is detected in the sampled "
              "compressed air." if is_en else
              "Kết quả được biểu thị “Không có chỉ thị” khi không phát hiện vi sinh vật mục tiêu trong khí nén được lấy mẫu."),
    ]
    # câu công nhận: cùng vị trí như bảng Ghi chú mục tổng dầu của template
    accred = "The performance field is accredited to ISO 17020 - VIAS 004." if is_en else \
        "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004."
    oil_rows = [_ptext(tr) for tr in oil_note.findall(qn("w:tr"))[1:]]
    acc_first = bool(oil_rows) and "VIAS" in oil_rows[0]
    lines = ([("●", accred)] + lines) if acc_first else (lines + [("●", accred)])
    note = put(note_rows(oil_note, lines))
    _keep_rows(note.findall(qn("w:tr"))[:-1])


def _aerotrak_band_headers(doc):
    """AeroTrak đo từ 0,3 µm -> tiêu đề cột '0,1 µm < d ≤ 0,5 µm' đổi thành 0,3 µm (như build_137035)."""
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        runs = list(p.iter(qn("w:t")))
        txt = "".join(t.text or "" for t in runs)
        if ("d ≤ 0,5" in txt or "d ≤ 0.5" in txt) and ("0,1" in txt or "0.1" in txt) and len(txt) < 60:
            for t in runs:
                if t.text and ("0,1" in t.text or "0.1" in t.text):
                    t.text = t.text.replace("0,1", "0,3").replace("0.1", "0.3")
                    n += 1
                    break
    n += _replace_in_paragraphs(doc.element.body, {"(0,1-5,0) µm": "(0,3-5,0) µm", "(0.1-5.0) µm": "(0.3-5.0) µm"})
    return n


def _renumber_note3(doc):
    """Ghi chú 'Không phân loại' của template mang số (3); bảng không có (2) (chứng thư /2 chỉ đo hạt) -> đổi (3) thành
    (2) như TBKQ 2025 / 2026."""
    body = doc.element.body
    if any("(2)" in _ptext(p) for p in body.iter(qn("w:p"))):
        return
    _replace_in_paragraphs(body, {"(3)": "(2)"})


def build_tbkq(d, is_en: bool) -> str:
    src = CompressedAirHandler(d).build_tbkq(is_en=is_en)
    doc = Document(src)
    _add_micro(doc, is_en, d.locations[0].points_summary_vn[0])
    print(f"[build_128967] tiêu đề cột 0,3 µm: {_aerotrak_band_headers(doc)} chỗ")
    _renumber_note3(doc)            # 'Không phân loại' (3) -> (2) như TBKQ 2025
    _fix_cover_lists(doc)
    from Core.Builders.tag_context import shrink_cover_break
    shrink_cover_break(doc, compact=True)   # bìa EN tràn sang trang 2 (tên KH, địa điểm dài — AGENTS mục 14, 15)
    _keep_headings_with_tables(doc)
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    style_note_tables(doc)          # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - {d.sub_no} - khi nen - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)              # tên handler 'TBKQ - vn - n - khi nen - …' (file do code tạo, E41)
    print(f"[build_128967] TBKQ {'EN' if is_en else 'VN'} /{d.sub_no} -> {dst}")
    return dst


# ---------------------------------------------------------------------------
# BCKQ
# ---------------------------------------------------------------------------
def build_bckq(d) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - {d.sub_no} - khi nen - {CODE}.docx")
    doc = Document(out)
    body = doc.element.body
    sym = d.locations[0].points_summary_vn[0]
    _replace_in_paragraphs(body, {
        f"phiếu Tính toán kết quả nồng độ hạt bụi số {CODE} ngày": f"phiếu Kết quả tính toán số {CODE} - {sym} ngày",
        f"phiếu Tính toán kết quả nhiệt độ điểm sương số {CODE} ngày": f"phiếu Kết quả tính toán số {CODE} - {sym} ngày",
    })
    if True:
        _replace_in_paragraphs(body, {
            "(0,1-5,0) µm": "(0,3-5,0) µm",
            "Nồng độ hạt bụi có kích thước (0,3-5,0) µm.": "Nồng độ hạt bụi có kích thước (0,3-5,0) µm; Độ nhiễm bẩn vi sinh vật.",
        })
        kids = list(body.iterchildren(qn("w:p")))
        dp_head = next(p for p in kids if _ptext(p).strip() == "Nhiệt độ điểm sương")
        dp_ref = dp_head.getnext()
        head = copy.deepcopy(dp_head)
        dp_ref.addnext(head)
        _set_text(head, "Vi sinh trong khí nén")
        _keep(head)
        kq = "; ".join(f"{vn} {k}" for vn, _en, k, _v, _ve in MICRO)
        lines = [
            f"Xem tại phiếu KQTN số {KQTN_VS['so']}.0{d.sub_no} ngày {KQTN_VS['ngay']} (mẫu {sym}, Pinocchio Super II {VS_HOLES} lỗ, "
            f"{VS_LIT} lít: {kq}).",
            "- KQ được biểu thị < 1 CFU/đĩa → không phát hiện vi sinh vật mục tiêu trong thể tích khí nén được lấy mẫu "
            "~ Không có chỉ thị.",
            f"- KQTN của mẫu Blank (phiếu {KQTN_VS['blank']}, mẫu KN3 (BLANK)): tổng số vi sinh vật hiếu khí, tổng số "
            "nấm men, nấm mốc < 1 CFU/đĩa → môi trường lấy mẫu đảm bảo kết quả.",
        ]
        anchor = head
        for i, line in enumerate(lines):
            p = copy.deepcopy(dp_ref)
            anchor.addnext(p)
            _set_text(p, line)
            if i < len(lines) - 1:
                _keep(p)
            anchor = p
    _keep_headings_with_tables(doc)
    for t in doc.tables[1:]:
        if len(t.rows) > 1 and t.rows[0].cells[0].text.strip().startswith("Vị trí"):
            add_tbl_header(t.rows[0])
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "- Hình ảnh", "- Sơ đồ",
                                      "Kết quả đo đạc", "Hàm lượng tổng dầu", "Nồng độ hạt bụi")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    print(f"[build_128967] BCKQ /{d.sub_no} -> {out}")
    return out


def split_edoc() -> list:
    """BBLM 2 trang -> TBKQ/Edoc (AGENTS 8A). Thư mục hồ sơ không có scan phiếu ghi chép / PYC -> chỉ có BBLM."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    pdf = fitz.open()
    pdf.insert_pdf(fitz.open(SCAN_BBLM))
    name = f"{CODE6} - BBLM.pdf"
    finish_edoc(pdf, INPUT_DIR, name)
    p = os.path.join(edoc, name)
    pdf.save(p)
    print(f"[build_128967] Edoc -> {p}")
    return [p]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()

    xlsx = build_workbooks()
    for p in POINTS:
        prepare_workbook(xlsx[p["sym"]], p)
    d1 = make_dossier(xlsx["KN1"], POINTS[0], 1, CERT_DATE_1)
    d2 = make_dossier(xlsx["KN2"], POINTS[1], 2, CERT_DATE_2)
    assign_cert_numbers([d1, d2])
    for d in (d1, d2):
        d.ensure_output_directories()
    split_edoc()

    docs = {}
    for d in (d1, d2):
        docs[d.sub_no] = [build_bckq(d), build_tbkq(d, False), build_tbkq(d, True)]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(d1).export_files(docs[1])
        PDFExporter(d2).export_files(docs[2])
        PDFExporter(d1).export_excel_files([xlsx["KN1"], xlsx["KN2"]])
    print("\n===== QA =====")
    for n in (1, 2):
        for p in docs[n]:
            print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
