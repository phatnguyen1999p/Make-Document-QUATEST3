# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D136248 - Công ty TNHH Đông Phương Vũng Tàu (09/11/2026) — khí nén, 1 vị trí, CHỈ tổng dầu (lấy mẫu).

Nguồn (Input/136248 — chép từ thư mục hồ sơ trên Z 16/11/2026, chỉ đọc):
  0.  BBLM  .pdf (= E DOC/Q326D136248 - BBGĐHT .pdf) — BM3030.02 2 trang: PYC Q326D136248 ngày 26/10/2026; "CÔNG TY TNHH
      ĐÔNG PHƯƠNG VŨNG TÀU"; "Đường số 11, KCN Đông Xuyên, Phường Rạch Dừa, Thành phố Hồ Chí Minh"; "Lấy mẫu/ phân tích
      chất lượng khí nén"; địa điểm "tại địa chỉ nêu tại mục 3"; ngày 09/11/2026 từ 08h45 đến 13h00; thiết bị giấy lọc
      Q3400 Sartorius + ORBO 32 + PS Cassettes 47 mm, DP 510 + CS16; quy trình N2.TC.KN.01; nhà máy hoạt động sản xuất;
      báo giá QONV26015248; vị trí (viết tay) "Ống dẫn khí nén sau lọc tại khu vực máy nén khí", lấy mẫu tổng dầu,
      01 mẫu, ký hiệu "KN". Tr. 2 viết tay: hệ khí nén gồm 07 máy nén khí SULLAIR có dùng dầu bôi trơn; "Theo quy định
      Công ty, không được dùng điện thoại trong nhà máy và không thể hiện sơ đồ hệ khí nén tại TBKQ". Người thực hiện
      Phạm Quốc Nghĩa.
  THUC HIEN/D261109 - nhiet do/BM00000–02.bmp — màn hình DP510 09:44–09:47: nhiệt độ khí 34,70 / 34,72 / 35,13 °C,
      áp suất 6,068 / 6,047 / 6,036 bar (dùng cho N*).
  THUC HIEN/D136248 - PHOTO — 5 ảnh IMG_73xx + "KN - vi tri .jpg" (người dùng chọn, cũng đặt ở gốc thư mục).
  3. TBKQ.docx — bản nháp người dùng chép từ hồ sơ khác (Medcen 120473): chỉ tham khảo cách trình bày (chỉ tổng dầu,
      ghi chú thêm "Phương pháp thực hiện theo ISO 8573 - 2 và ISO 8573 - 5"); không lấy thông tin khách.
  Không có KQTN tổng dầu, không có scan phiếu ghi chép (BM3030.01 trong THUC HIEN là form trống).

Quyết định của người dùng (16/11/2026):
  1. Ký hiệu vị trí "để KN" (theo BBLM, không đổi thành KN1).
  2. Tên vị trí "Ống dẫn khí nén sau lọc tại khu vực máy nén khí" — "Đúng rồi".
  3. Ảnh vị trí trong TBKQ: "giữ, có gì thay đổi sau".
  4. Nhiệt độ khí cho N*: "làm tròn 35 độ".
  5. Lưu lượng / thời gian lấy mẫu dầu theo form: "okay, khi nào thay đổi sẽ báo".
  6. Tên EN khách hàng AI dịch: "okay".
  7. Ghi chú "Phương pháp thực hiện theo ISO 8573-2 và ISO 8573-5" (bản nháp người dùng): "highlight vàng giúp tôi" ->
     thêm vào bảng Ghi chú tổng dầu TBKQ VN/EN, tô vàng (chờ người dùng chốt giữ hay bỏ).

Suy luận của AI (cần người dùng xác nhận):
  - Ký hiệu "KN1" (AGENTS mục 6: vị trí đầu ghi KN1), BBLM ghi "KN".
  - Nhiệt độ khí cho N* = trung bình 3 màn hình DP510 (34,85 °C); áp suất lấy mẫu 1 bar như form.
  - TBKQ có 1 ảnh vị trí (KN - vi tri.jpg); không có sơ đồ hệ thống (BBLM). BCKQ ghi mô tả hệ thống theo BBLM tr. 2.
  - Lưu lượng / thời gian lấy mẫu dầu như form (35 L/phút × 240 phút hạt, 1 L/phút × 240 phút hơi) — chưa có phiếu ghi chép.
  - Thiết bị: giấy lọc dầu (tên chuẩn Pall như 128967) + DP510 (đo nhiệt độ, áp suất).
  - Tên tiếng Anh khách hàng AI dịch (chưa có MST / GCN).
  - Ngày yêu cầu cấp KQTN = ngày lấy mẫu + 8 = 17/11/2026 (không có PYC).

Chạy:  python build_136248.py          (DOCX + XLSX + Edoc)
       python build_136248.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
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

CODE = "Q326D136248"
CODE6 = "136248"
REG_DATE = "26/10/2026"            # BBLM mục 1 (PYC)
SAMPLE_DATE = "09/11/2026"         # BBLM
START, END = "08:45", "13:00"      # BBLM mục 6
CERT_DATE = date.today().strftime("%d/%m/%Y")   # người dùng 16/11/2026: ngày ban hành = ngày dựng (kể cả khi chờ KQTN)
NGAY_YC_KQTN = "17/11/2026"        # TẠM: ngày lấy mẫu + 8 ngày (không có PYC — AGENTS mục 13)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SRC_TH = os.path.join(INPUT_DIR, "THUC HIEN")
SCAN_BBLM = os.path.join(INPUT_DIR, "0.  BBLM  .pdf")
PHOTO = os.path.join(INPUT_DIR, "0.  KN - vi tri .jpg")
PHOTO_OPT = os.path.join(SRC_TH, "D136248 - PHOTO", "_optimized")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Đông Phương Vũng Tàu",                                     # BBLM
    name_en="Dong Phuong Vung Tau Company Limited",                                 # AI dịch
    address_vn="Đường số 11, Khu công nghiệp Đông Xuyên, phường Rạch Dừa, Thành phố Hồ Chí Minh",   # BBLM
    address_en="Street No. 11, Dong Xuyen Industrial Park, Rach Dua ward, Ho Chi Minh City",
)
LOCATION_VN = f"Nhà máy của {CLIENT.name_vn} - {CLIENT.address_vn}"   # BBLM mục 5: "tại địa chỉ nêu tại mục 3"
LOCATION_EN = f"Factory of {CLIENT.name_en} - {CLIENT.address_en}"
EXECUTOR = "Nghĩa"                 # BBLM: Phạm Quốc Nghĩa
SYM = "KN"                         # người dùng 16/11/2026: "để KN" (BBLM)
POINT = {
    "name_vn": "Ống dẫn khí nén sau lọc tại khu vực máy nén khí",
    "name_en": "Compressed air pipeline after the filter at the air compressor area",
    "gas_temp": 35,                    # DP510 BM00000–02 (34,70–35,13); người dùng: "làm tròn 35 độ"
    "pressure": round((6.068 + 6.047 + 6.036) / 3, 2),
}
SYSTEM_VN = "07 máy nén khí SULLAIR (có dùng dầu bôi trơn)"     # BBLM tr. 2
NO_DIAGRAM_VN = f"(Theo quy định của {CLIENT.name_vn}, không thể hiện sơ đồ hệ thống khí nén trong Chứng thư giám định)"

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




# ---------------------------------------------------------------------------
# Phiếu tính toán khí nén (ThucHien) — template TinhToanKhiNen, giữ công thức (AGENTS mục 11)
# ---------------------------------------------------------------------------
def build_workbook() -> str:
    t, p = POINT["gas_temp"], POINT["pressure"]
    point_data = {
        "code": CODE, "point_label": SYM, "cert_date": CERT_DATE,
        "location_name": POINT["name_vn"], "location_symbol": SYM,
        "location_photo": _jpg(PHOTO, PHOTO_OPT),
        # chỉ nhiệt độ khí (N*); điểm sương không đo -> khung gạch chéo sau
        "dewpoint": {"temp_start": t, "temp_end": t, "pres_start": p, "pres_end": p},
        "oil": {"flow_aerosol": 35, "flow_vapor": 1, "time_aerosol": 240, "time_vapor": 240},
    }
    path = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - {SYM} - TinhToanKhiNen.xlsx")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    build_compressed_air_workbook(point_data, path)
    return path


def prepare_workbook(path: str):
    """Chỉ tổng dầu: khung hạt bụi (Lasair, AeroTrak), điểm sương, vi sinh gạch chéo, tổng hợp '-'; dầu chưa có KQTN
    -> ……… (giữ công thức, như 128967); ảnh vị trí vừa dòng 17–29 (không tràn trang)."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Tinh_Toan_Khi_Co_Ap_V6"]
    ws._images = [im for im in ws._images if im.anchor._from.row < 84]      # bỏ ảnh HumCalc / DP510
    ws["C36"].value = POINT["gas_temp"]
    ws["B36"].value = "Nhiệt độ khí lấy mẫu T (oC):"
    for rng in ("C54:J59", "C61:J68", "C77:J83", "C87:J89", "B91:J111"):
        strike_box(ws, rng)
    for r in range(4, 9):
        for c in ("D", "F", "H"):
            if ws[f"{c}{r}"].value not in (None, "") and c != "E":
                ws[f"{c}{r}"].value = "-"
    for r in range(11, 16):
        ws[f"D{r}"].value = "-"
    ws["F3"].value = ws["F10"].value = ws["C52"].value = ws["C85"].value = "-"
    ws["C46"].value = ws["D46"].value = PLACEHOLDER
    wait = f'OR(C46="{PLACEHOLDER}",D46="{PLACEHOLDER}")'
    for ref in ("C47", "D47", "C48", "C49", "F9"):
        f = str(ws[ref].value or "")
        if f.startswith("="):
            ws[ref].value = f'=IF({wait},"{PLACEHOLDER}",{f[1:]})'
    for ref in ("I43", "J43", "I44", "J44", "H9"):
        ws[ref].value = PLACEHOLDER
    from openpyxl.styles import Alignment
    a = ws["H2"].alignment
    ws["H2"].alignment = Alignment(horizontal=a.horizontal, vertical=a.vertical, shrink_to_fit=True, wrap_text=True)
    from openpyxl.utils.units import pixels_to_EMU
    for im in ws._images:             # ảnh vị trí: cao 200pt, nằm trong dòng 17–29 (như 128967)
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
        print(f"[build_136248] Excel COM warning: {e}")


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def make_dossier(xlsx: str):
    d = create_dossier_from_compressed_air_workbooks(
        [xlsx], code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        client=CLIENT, personnel=PersonnelInfo(executor=EXECUTOR),
        equipments=[eq("oil_kit"), eq("cs_dp510")],
        scope_vn="Lấy mẫu/ phân tích chất lượng khí nén", scope_en="Sampling/ analyzing compressed air quality",
        standards_vn="ISO 8573-1:2010", standards_en="ISO 8573-1:2010",
        basis_vn="ISO 8573-1:2010", basis_en="ISO 8573-1:2010",
    )
    d.reg_date = REG_DATE
    d.procedure_vn = d.procedure_en = "QUATEST 3 3030:2025, N2.TC.KN.01"      # BBLM mục 8
    d.location_desc_vn, d.location_desc_en = LOCATION_VN, LOCATION_EN
    loc = d.locations[0]
    loc.name_vn, loc.name_en = POINT["name_vn"], POINT["name_en"]
    d.photos = [{"path": _jpg(PHOTO, PHOTO_OPT), "caption_vn": "", "caption_en": ""}]
    rd = d.results_data
    rd["system_status_vn"], rd["factory_status_vn"] = "Đang hoạt động", "Đang hoạt động sản xuất"   # BBLM mục 9
    rd["system_status_en"], rd["factory_status_en"] = "In operation", "In operation"
    rd["particles"] = {"rows": []}
    rd["dewpoint"] = {"rows": []}
    for r in rd["oil"]["rows"]:
        r["val"], r["class"] = PLACEHOLDER, PLACEHOLDER
    rd["kqtn_by_location"] = {SYM: f"- Dầu hơi: xem tại phiếu MTTN{PLACEHOLDER}\n- Dầu hạt: xem tại phiếu HDTN{PLACEHOLDER}"}
    rd["kqtn"] = [{"so": f"MTTN{PLACEHOLDER}", "ngay": PLACEHOLDER, "ngay_gui": PLACEHOLDER, "ngay_nhan": PLACEHOLDER,
                   "ngay_yc": NGAY_YC_KQTN, "tre_han": PLACEHOLDER},
                  {"so": f"HDTN{PLACEHOLDER}", "ngay": PLACEHOLDER, "ngay_gui": PLACEHOLDER, "ngay_nhan": PLACEHOLDER,
                   "ngay_yc": NGAY_YC_KQTN, "tre_han": PLACEHOLDER}]
    rd["system_chain_vn"] = [SYSTEM_VN, f"Ống dẫn khí nén sau lọc ({SYM})"]
    rd["bckq_label"] = "khi nen"
    d.input_dir, d.output_dir, d.network_dir = INPUT_DIR, "", ""
    d.thuchien_dir = os.path.join(INPUT_DIR, "ThucHien")
    d.tbkq_dir = os.path.join(INPUT_DIR, "TBKQ")
    d.edoc_dir = os.path.join(d.tbkq_dir, "Edoc")
    d.pdf_dir = os.path.join(INPUT_DIR, "pdf")
    return d


def _add_method_note(doc, is_en):
    """Ghi chú phương pháp (bản nháp người dùng) thêm vào bảng Ghi chú tổng dầu, ngay trước câu công nhận; tô vàng
    theo yêu cầu người dùng 16/11/2026 ("highlight vàng giúp tôi") — chờ chốt giữ / bỏ."""
    text = ("The method is performed according to ISO 8573-2 and ISO 8573-5." if is_en else
            "Phương pháp thực hiện theo ISO 8573-2 và ISO 8573-5.")
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        if not _is_note_table(tbl):
            continue
        rows = tbl.findall(qn("w:tr"))
        acc = next((tr for tr in rows if "VIAS" in _ptext(tr)), None)
        if acc is None:
            continue
        tr = copy.deepcopy(acc)
        acc.addprevious(tr)
        _set_cell(tr.findall(qn("w:tc"))[1], text)
        for r in tr.findall(qn("w:tc"))[1].iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            h = OxmlElement("w:highlight")
            h.set(qn("w:val"), "yellow")
            rpr.append(h)
        return


def build_tbkq(d, is_en: bool) -> str:
    rd = d.results_data
    chain = rd.pop("system_chain_vn", None)      # BBLM: không thể hiện sơ đồ hệ thống tại TBKQ (chỉ BCKQ)
    try:
        src = CompressedAirHandler(d).build_tbkq(is_en=is_en)
    finally:
        if chain:
            rd["system_chain_vn"] = chain
    doc = Document(src)
    _drop_system_and_photo(doc, is_en)
    _add_method_note(doc, is_en)
    _fix_cover_lists(doc)
    _keep_headings_with_tables(doc)
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    style_note_tables(doc)          # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - khi nen - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    print(f"[build_136248] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


def build_bckq(d) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - khi nen - {CODE}.docx")
    doc = Document(out)
    body = doc.element.body
    intro = next((p for p in body.iterchildren(qn("w:p")) if _ptext(p).strip().startswith("- Sơ đồ hệ thống")), None)
    if intro is not None:
        chain_p = intro.getnext()
        note_p = copy.deepcopy(intro)
        chain_p.addnext(note_p)
        _set_text(note_p, NO_DIAGRAM_VN)
        for r in note_p.findall(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            rpr.append(OxmlElement("w:i"))
    _keep_headings_with_tables(doc)
    for t in doc.tables[1:]:
        if len(t.rows) > 1 and t.rows[0].cells[0].text.strip().startswith("Vị trí"):
            add_tbl_header(t.rows[0])
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "- Hình ảnh", "- Sơ đồ",
                                      "Kết quả đo đạc", "Hàm lượng tổng dầu")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    print(f"[build_136248] BCKQ -> {out}")
    return out


def split_edoc() -> list:
    """BBLM 2 trang -> TBKQ/Edoc (AGENTS 8A). Không có scan phiếu ghi chép / PYC."""
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
    print(f"[build_136248] Edoc -> {p}")
    return [p]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()
    xlsx = build_workbook()
    d = make_dossier(xlsx)
    assign_cert_numbers([d])        # 1 chứng thư -> mã trần (AGENTS mục 2)
    # đọc phiếu trước khi gạch chéo khung không đo (parser cần ô điểm sương)
    prepare_workbook(xlsx)
    d.ensure_output_directories()
    split_edoc()
    docs = [build_bckq(d), build_tbkq(d, False), build_tbkq(d, True)]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(d).export_files(docs)
        PDFExporter(d).export_excel_files([xlsx])
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
