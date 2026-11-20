# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D137792 - Công ty TNHH Bao bì Nước giải khát Crown Đồng Nai (12/11/2026) — khí nén, CHỈ vi sinh, 3 vị trí
KN1–KN3, 1 chứng thư (như 2025, hồ sơ 01000/N2.25/DG).

Nguồn (Input/137792 — chép từ thư mục hồ sơ trên Z 16/11/2026, chỉ đọc):
  0. BBLM .pdf — BM3030.02 2 trang: PYC Q326D137792 ngày 29/10/2026; "CÔNG TY TNHH BAO BÌ NƯỚC GIẢI KHÁT CROWN ĐỒNG
      NAI"; "Đường số 4, KCN Tam Phước, Phường Tam Phước, Thành Phố Đồng Nai"; "Lấy mẫu/ phân tích chất lượng khí nén";
      địa điểm "tại địa chỉ nêu tại mục 3"; ngày 12/11/2026 từ 8h45 đến 11h45; thiết bị DP 510 + CS16, SAS Pinocchio
      Super II; quy trình N2.TC.KN.01; nhà máy hoạt động sản xuất; vị trí (viết tay) KN1 "Ống dẫn khí nén tại Can line 3
      - 4 và end line 3-4-5", KN2 "Ống dẫn khí nén tại Cap line 5" (đọc "Cap"/"Can" chưa chắc), KN3 "Ống dẫn khí nén
      tại End line 1-2-6-7"; mỗi vị trí lấy mẫu vi sinh 01 mẫu. Tr. 2: 16 máy nén khí — 11 máy Ingersoll Rand, 5 máy
      Atlas Copco, các máy có sử dụng dầu. Người thực hiện Phạm Quốc Nghĩa.
  THUC HIEN/D261112 - nhiet do — DP510: "KN1 - NHIET" 36,31 °C (09:53), "KN2 - NHIET" 41,62 °C (10:36),
      "KN3 - NHIET" 33,33 °C (11:28); áp suất không đo ("Range ?").
  THUC HIEN/137792 - BM.VS.CL.0007.03.xlsx — giao nhận môi trường 10/11/2026: đĩa Bacillus 6, VRBA 6, PCA 10, TBX 6,
      BPA 6, DG18 6.
  THUC HIEN/1000 - TBKQ - V / E.docx (= "- TBKQ - V/E - 2025") — TBKQ 2025: 8 chỉ tiêu (tổng số VSV hiếu khí, nấm men,
      nấm mốc, E. coli, Coliforms, S. aureus, B. cereus, Salmonella), tên vị trí, cách trình bày.
  THUC HIEN/137792 - PHOTO — KN1.jpg, KN2.jpg, KN3.jpg (người dùng chọn) + 4 ảnh khác.
  Chưa có KQTN vi sinh.

Quyết định của người dùng (16/11/2026):
  1. 8 chỉ tiêu như 2025: "oke, đúng rồi".
  2. KN2: "Can line".
  3. Tên EN theo 2025: "oke, highlight lại nhé" -> tô vàng tên EN trong TBKQ EN.
  5. Nhiệt độ khí: "36,42,33" (KN1 36, KN2 42, KN3 33 °C).
  6. Địa chỉ: "thành phố Đồng Nai".

Suy luận của AI (cần người dùng xác nhận):
  - 1 chứng thư 3 vị trí (như 2025); chỉ tiêu 8 như 2025 (môi trường giao nhận khớp) — chờ KQTN xác nhận.
  - Tên vị trí theo BBLM viết lại theo 2025 ("Ống dẫn khí nén tại …"); KN2 "Cap line 5" hay "Can line 5" (2025).
  - Nhiệt độ khí (N*) theo màn hình DP510 có nhãn KN1/KN2/KN3; áp suất lấy mẫu 1 bar.
  - Pinocchio Super II 218 lỗ, 500 L như phiếu khí nén các hồ sơ khác (chưa có phiếu ghi chép).
  - Thiết bị: SAS Pinocchio + DP510 (đo nhiệt độ khí).
  - Tên tiếng Anh khách hàng theo TBKQ EN 2025: "Crown Beverage Cans (Dong Nai) Limited" (chưa có MST / GCN).
  - Ngày yêu cầu cấp KQTN = ngày lấy mẫu + 8 = 20/11/2026 (không có PYC).

Chạy:  python build_137792.py [--pdf]
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

CODE = "Q326D137792"
CODE6 = "137792"
REG_DATE = "29/10/2026"            # BBLM mục 1 (PYC)
SAMPLE_DATE = "12/11/2026"         # BBLM
START, END = "08:45", "11:45"      # BBLM mục 6
CERT_DATE = date.today().strftime("%d/%m/%Y")   # người dùng 16/11/2026: ngày ban hành = ngày dựng (kể cả khi chờ KQTN)
NGAY_YC_KQTN = "20/11/2026"        # TẠM: ngày lấy mẫu + 8 ngày (không có PYC — AGENTS mục 13)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SRC_TH = os.path.join(INPUT_DIR, "THUC HIEN")
SCAN_BBLM = os.path.join(INPUT_DIR, "0. BBLM .pdf")
PHOTO_DIR = os.path.join(SRC_TH, "137792 - PHOTO")
PHOTO_OPT = os.path.join(PHOTO_DIR, "_optimized")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Bao bì Nước giải khát Crown Đồng Nai",                    # BBLM (viết hoa chữ đầu như 2025)
    name_en="Crown Beverage Cans (Dong Nai) Limited",                               # TBKQ EN 2025
    address_vn="Đường số 4, Khu công nghiệp Tam Phước, phường Tam Phước, Thành phố Đồng Nai",   # BBLM; người dùng 16/11
    address_en="Road No. 4, Tam Phuoc Industrial Park, Tam Phuoc ward, Dong Nai City",
)
LOCATION_VN = f"Nhà máy của {CLIENT.name_vn} - {CLIENT.address_vn}"   # BBLM mục 5: "tại địa chỉ nêu tại mục 3"
LOCATION_EN = f"Factory of {CLIENT.name_en} - {CLIENT.address_en}"
EXECUTOR = "Nghĩa"                 # BBLM: Phạm Quốc Nghĩa

POINTS = [
    {"sym": "KN1", "name_vn": "Ống dẫn khí nén tại Can line 3-4 và End line 3-4-5",
     "name_en": "Compressed air pipeline at Can line 3-4 and End line 3-4-5", "gas_temp": 36, "photo": "KN1  .jpg"},
    {"sym": "KN2", "name_vn": "Ống dẫn khí nén tại Can line 5",
     "name_en": "Compressed air pipeline at Can line 5", "gas_temp": 42, "photo": "KN2  .jpg"},
    {"sym": "KN3", "name_vn": "Ống dẫn khí nén tại End line 1-2-6-7",
     "name_en": "Compressed air pipeline at End line 1-2-6-7", "gas_temp": 33, "photo": "KN3 .jpg"},
]
VS_HOLES, VS_LIT = 218, 500
# (chỉ tiêu VN, chỉ tiêu EN) — theo TBKQ 2025; kết quả chờ KQTN
MICRO = [("Tổng số vi sinh vật hiếu khí", "Total aerobic microorganisms"),
         ("Tổng số nấm men", "Total yeasts"),
         ("Tổng số nấm mốc", "Total molds"),
         ("Escherichia coli", "Escherichia coli"),
         ("Coliforms", "Coliforms"),
         ("Staphylococcus aureus", "Staphylococcus aureus"),
         ("Bacillus cereus", "Bacillus cereus"),
         ("Salmonella", "Salmonella")]
SYSTEM_VN = ["16 máy nén khí: 11 máy Ingersoll Rand, 5 máy Atlas Copco (có sử dụng dầu)", "Ống dẫn khí nén (KN1, KN2, KN3)"]
SYSTEM_EN = ["16 air compressors: 11 Ingersoll Rand, 5 Atlas Copco (oil-lubricated)", "Compressed air pipeline (KN1, KN2, KN3)"]

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

def _highlight_text(doc, needles) -> int:
    """Tô vàng cả đoạn chứa một trong các chuỗi `needles` (thân văn bản + bảng) — điểm người dùng muốn xem lại."""
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        if not any(x in _ptext(p) for x in needles):
            continue
        for r in p.iter(qn("w:r")):
            if not r.findall(qn("w:t")):
                continue
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            if rpr.find(qn("w:highlight")) is None:
                h = OxmlElement("w:highlight")
                h.set(qn("w:val"), "yellow")
                rpr.append(h)
        n += 1
    return n



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


# ---------------------------------------------------------------------------
# Phiếu tính toán khí nén (ThucHien) — một phiếu / vị trí, chỉ vi sinh
# ---------------------------------------------------------------------------
def build_workbooks() -> dict:
    out = {}
    for p in POINTS:
        t = p["gas_temp"]
        point_data = {
            "code": CODE, "point_label": p["sym"], "cert_date": CERT_DATE,
            "location_name": p["name_vn"], "location_symbol": p["sym"],
            "location_photo": _jpg(os.path.join(PHOTO_DIR, p["photo"]), PHOTO_OPT),
            "dewpoint": {"temp_start": t, "temp_end": t},
            "oil": {"flow_aerosol": 35, "flow_vapor": 1, "time_aerosol": 240, "time_vapor": 240},
        }
        path = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - {p['sym']} - TinhToanKhiNen.xlsx")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        build_compressed_air_workbook(point_data, path)
        out[p["sym"]] = path
    return out


def prepare_workbook(path: str, p: dict):
    """Chỉ vi sinh: hạt bụi, điểm sương, tổng dầu gạch chéo, tổng hợp '-'; vi sinh chờ KQTN -> ………;
    ảnh vị trí vừa dòng 17–29 (không tràn trang, như 128967)."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Tinh_Toan_Khi_Co_Ap_V6"]
    ws._images = [im for im in ws._images if im.anchor._from.row < 84]      # bỏ ảnh HumCalc / DP510
    ws["C36"].value = p["gas_temp"]
    ws["B36"].value = "Nhiệt độ khí lấy mẫu T (oC):"
    for rng in ("C43:D49", "C54:J59", "C61:J68", "C87:J89", "B91:J111"):
        strike_box(ws, rng)
    for ref in ("D4", "D5", "D6", "D7", "D8", "F8", "H8", "D9", "F9", "H9", "F3", "C52", "C85",
                "I43", "J43", "I44", "J44"):
        ws[ref].value = "-"
    ws["C77"].value, ws["C78"].value = VS_HOLES, VS_LIT
    for c in "CDEFG":
        ws[f"{c}81"].value = PLACEHOLDER          # chờ KQTN vi sinh
    for r in range(11, 16):
        ws[f"D{r}"].value = PLACEHOLDER
    ws["I77"].value = PLACEHOLDER
    from openpyxl.styles import Alignment
    a = ws["H2"].alignment
    ws["H2"].alignment = Alignment(horizontal=a.horizontal, vertical=a.vertical, shrink_to_fit=True, wrap_text=True)
    from openpyxl.utils.units import pixels_to_EMU
    for im in ws._images:
        if im.anchor._from.row == 16:
            h_px = round(200 * 4 / 3)
            im.width, im.height = round(im.width * h_px / im.height), h_px
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
            x.Save()
            x.Close()
        finally:
            xl.Quit()
    except ImportError as e:
        print(f"[build_137792] Excel COM warning: {e}")


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def make_dossier(xlsx: dict):
    d = create_dossier_from_compressed_air_workbooks(
        [xlsx[p["sym"]] for p in POINTS], code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        client=CLIENT, personnel=PersonnelInfo(executor=EXECUTOR),
        equipments=[eq("sas_pinocchio"), eq("cs_dp510")],
        scope_vn="Lấy mẫu/ phân tích chất lượng khí nén", scope_en="Sampling/ analyzing compressed air quality",
        standards_vn="ISO 8573-1:2010", standards_en="ISO 8573-1:2010",
        basis_vn="ISO 8573-1:2010", basis_en="ISO 8573-1:2010",
    )
    d.reg_date = REG_DATE
    d.procedure_vn = d.procedure_en = "QUATEST 3 3030:2025, N2.TC.KN.01"      # BBLM mục 8
    d.location_desc_vn, d.location_desc_en = LOCATION_VN, LOCATION_EN
    for loc, p in zip(d.locations, POINTS):
        loc.name_vn, loc.name_en = p["name_vn"], p["name_en"]
    d.photos = [{"path": _jpg(os.path.join(PHOTO_DIR, p["photo"]), PHOTO_OPT),
                 "caption_vn": f"Vị trí lấy mẫu {p['sym']}", "caption_en": f"Sampling location {p['sym']}"} for p in POINTS]
    rd = d.results_data
    rd["system_status_vn"], rd["factory_status_vn"] = "Đang hoạt động", "Đang hoạt động sản xuất"   # BBLM mục 9
    rd["system_status_en"], rd["factory_status_en"] = "In operation", "In operation"
    rd["system_chain_vn"], rd["system_chain_en"] = SYSTEM_VN, SYSTEM_EN
    rd["particles"] = {"rows": []}
    rd["dewpoint"] = {"rows": []}
    for r in rd["oil"]["rows"]:      # giữ khung mục dầu để dựng mục vi sinh rồi xóa (template không có mục vi sinh)
        r["val"], r["class"] = PLACEHOLDER, PLACEHOLDER
    rd["kqtn_by_location"] = {p["sym"]: f"- Vi sinh: chờ KQTN {PLACEHOLDER}" for p in POINTS}
    rd["kqtn"] = [{"so": f"Vi sinh {PLACEHOLDER}", "ngay": PLACEHOLDER, "ngay_gui": PLACEHOLDER,
                   "ngay_nhan": PLACEHOLDER, "ngay_yc": NGAY_YC_KQTN, "tre_han": PLACEHOLDER}]
    rd["bckq_label"] = "khi nen"
    d.input_dir, d.output_dir, d.network_dir = INPUT_DIR, "", ""
    d.thuchien_dir = os.path.join(INPUT_DIR, "ThucHien")
    d.tbkq_dir = os.path.join(INPUT_DIR, "TBKQ")
    d.edoc_dir = os.path.join(d.tbkq_dir, "Edoc")
    d.pdf_dir = os.path.join(INPUT_DIR, "pdf")
    return d


# ---------------------------------------------------------------------------
# TBKQ — đổi mục tổng dầu thành mục vi sinh
# ---------------------------------------------------------------------------
def _micro_from_oil(doc, is_en):
    body = doc.element.body
    oil_param = "Total oil content." if is_en else "Hàm lượng tổng dầu."
    for p in list(body.iter(qn("w:p"))):
        if _ptext(p).strip() == oil_param:
            _set_text(p, "Microbial contamination." if is_en else "Độ nhiễm bẩn vi sinh vật.")
    tbls = list(body.iterchildren(qn("w:tbl")))
    oil = next(t for t in tbls if ("Total oil content, mg/m" if is_en else "Hàm lượng tổng dầu, mg/m") in _ptext(t))
    head = _prev_text_p(oil)
    nxt = oil.getnext()
    while nxt is not None and not _is_note_table(nxt):
        nxt = nxt.getnext()
    oil_note = nxt
    m = re.match(r"^(9\.2\.\d\s+)", _ptext(head).strip())
    _set_text(head, (m.group(1) if m else "") + ("Microbial contamination" if is_en else "Nhiễm bẩn vi sinh vật"))
    rows = oil.findall(qn("w:tr"))
    data_proto = rows[1]
    for r in rows[1:]:
        oil.remove(r)
    tcs = rows[0].findall(qn("w:tc"))
    _set_cell(tcs[0], "Location" if is_en else "Vị trí thực hiện")
    _set_cell(tcs[1], "Parameter" if is_en else "Chỉ tiêu")
    _set_cell(tcs[2], "Colony-forming units in compressed air, CFU/m³ " if is_en else
              "Số khuẩn lạc trong khí nén, CFU/m³ ", "(1)")
    for p in POINTS:
        for i, (vn, en) in enumerate(MICRO):
            tr = copy.deepcopy(data_proto)
            oil.append(tr)
            c = tr.findall(qn("w:tc"))
            _set_cell(c[0], p["sym"] if i == 0 else "")
            _vmerge(c[0], restart=(i == 0))
            _set_cell(c[1], en if is_en else vn)
            _set_cell(c[2], PLACEHOLDER)
    grid = oil.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    total = sum(int(g.get(qn("w:w"))) for g in grid)
    widths = [round(total * w) for w in (0.18, 0.40, 0.42)]
    for g, w in zip(grid, widths):
        g.set(qn("w:w"), str(w))
    for tr in oil.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), widths):
            tcW = tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tcW is not None:
                tcW.set(qn("w:w"), str(w))
                tcW.set(qn("w:type"), "dxa")
    lines = [
        ("(1) …", "Result converted to reference conditions of absolute pressure of 1 bar and temperature of 20 °C "
                  "according to ISO 8573-7:2003." if is_en else
                  "Kết quả quy về điều kiện áp suất tuyệt đối 1 bar, nhiệt độ 20 °C theo ISO 8573-7:2003."),
        ("●", "The result is expressed as “No indication” when no target organism is detected in the sampled "
              "compressed air." if is_en else
              "Kết quả được biểu thị “Không có chỉ thị” khi không phát hiện vi sinh vật mục tiêu trong khí nén được lấy mẫu."),
        ("●", "The performance field is accredited to ISO 17020 - VIAS 004." if is_en else
              "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004."),
    ]
    new_note = note_rows(oil_note, lines)
    oil_note.addnext(new_note)
    oil_note.getparent().remove(oil_note)
    _keep_rows(new_note.findall(qn("w:tr"))[:-1])


def build_tbkq(d, is_en: bool) -> str:
    src = CompressedAirHandler(d).build_tbkq(is_en=is_en)
    doc = Document(src)
    _micro_from_oil(doc, is_en)
    _fix_cover_lists(doc)
    _keep_headings_with_tables(doc)
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    if is_en:
        print(f"[build_137792] tô vàng tên EN: {_highlight_text(doc, [CLIENT.name_en])} đoạn")
    style_note_tables(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - khi nen - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    print(f"[build_137792] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


def build_bckq(d) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - khi nen - {CODE}.docx")
    doc = Document(out)
    body = doc.element.body
    _replace_in_paragraphs(body, {"Hàm lượng tổng dầu.": "Độ nhiễm bẩn vi sinh vật.",
                                  "Hàm lượng tổng dầu": "Vi sinh trong khí nén"})
    _keep_headings_with_tables(doc)
    for t in doc.tables[1:]:
        if len(t.rows) > 1 and t.rows[0].cells[0].text.strip().startswith("Vị trí"):
            add_tbl_header(t.rows[0])
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "- Hình ảnh", "- Sơ đồ",
                                      "Kết quả đo đạc", "Vi sinh trong khí nén")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    print(f"[build_137792] BCKQ -> {out}")
    return out


def split_edoc() -> list:
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
    print(f"[build_137792] Edoc -> {p}")
    return [p]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    args = ap.parse_args()
    xlsx = build_workbooks()
    d = make_dossier(xlsx)
    assign_cert_numbers([d])
    for p in POINTS:
        prepare_workbook(xlsx[p["sym"]], p)
    d.ensure_output_directories()
    split_edoc()
    docs = [build_bckq(d), build_tbkq(d, False), build_tbkq(d, True)]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(d).export_files(docs)
        PDFExporter(d).export_excel_files([xlsx[p["sym"]] for p in POINTS])
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
