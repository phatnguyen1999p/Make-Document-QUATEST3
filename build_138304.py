# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D138304 - Công ty TNHH United International Pharma (16/11/2026) — khí nitơ: đo hàm lượng CO trong 1 bình
khí nitơ bằng ống chỉ thị Kitagawa 600SP tại hiện trường (không gửi mẫu, không KQTN), 1 chứng thư.

Nguồn (Input/138304 — chép từ thư mục hồ sơ trên Z 18/11/2026, chỉ đọc):
  0.  BBLM .pdf — BM3030.02 1 trang: PYC Q326D138304 ngày 29/10/2026; "CÔNG TY TNHH UNITED INTERNATIONAL PHARMA"; địa
      chỉ = địa điểm "Số 16 đường số 7, Khu công nghiệp Việt Nam - Singapore II, Phường Bình Dương, Thành phố Hồ Chí
      Minh"; "Giám định chất lượng khí Nitơ"; ngày 16/11/2026 từ 9h35 đến 10h10; thiết bị "Ống chỉ thị Kitagawa 600SP";
      quy trình QUATEST3 3030:2025, N2.TC.KN.01; nhà máy hoạt động sản xuất; báo giá QONV26016947; vị trí (viết tay)
      "Bình khí Nitơ Nitrogen có số QC: VN-261004", khí nitơ, đo CO, 01 vị trí, ký hiệu N2; "Đo hàm lượng CO trong khí
      Nito: < 5 ppm" (viết tay). Người thực hiện Phạm Quốc Nghĩa.
  THUC HIEN/D138304 - PHOTO/USE — N2.jpg, MA BINH CHUA NITO.jpg (người dùng chọn); thư mục ảnh còn "KQ DO CO.jpg".
  2. BCKQ - 2025, 3. TBKQ - 2025 - VN/EN — thực ra là hồ sơ khác (Q326D068548, T-Thai, khí nitơ): chỉ tham khảo cách
      trình bày (1 vị trí, bảng Chỉ tiêu / Phương pháp / Kết quả, Căn cứ "-", Ghi chú VIAS 004).

Quyết định của người dùng (18/11/2026):
  1. Tên vị trí: "Bình khí Nitrogen có số QC: VN-261004".
  2. Tên EN khách hàng AI dịch: "Oke".
  3. Cột phương pháp: "ghi theo BBLM đi, nhớ tô vàng" -> "Ống chỉ thị Kitagawa 600SP" (BBLM mục 7), tô vàng ở TBKQ VN/EN.
  4. Bỏ trạng thái / sơ đồ hệ thống khí nén: "oke". 5. Ảnh 2 tấm trong USE: "oke".

Suy luận của AI (cần người dùng xác nhận):
  - Tên vị trí "Bình khí nitơ (Nitrogen) có số QC: VN-261004" (đọc chữ viết tay BBLM).
  - Dựng trên template TBKQ khí nén: mục tổng dầu đổi thành "Hàm lượng cacbon monoxit (CO)", bỏ mô tả / trạng thái hệ
    thống khí nén (đối tượng là bình khí); không có phiếu tính toán Excel (đọc trực tiếp ống chỉ thị).
  - Phương pháp ghi "Đo tại hiện trường (ống chỉ thị Kitagawa 600SP)"; kết quả "< 5" ppm như BBLM.
  - Tên tiếng Anh khách hàng "United International Pharma Company Limited" (AI dịch, chưa có MST / GCN).
  - Ngày ban hành = ngày dựng (không có KQTN — AGENTS mục 2).

Chạy:  python build_138304.py [--pdf]
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

CODE = "Q326D138304"
CODE6 = "138304"
REG_DATE = "29/10/2026"            # BBLM mục 1 (PYC)
SAMPLE_DATE = "16/11/2026"         # BBLM
START, END = "09:35", "10:10"      # BBLM mục 6
CERT_DATE = date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày dựng (AGENTS mục 2)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SRC_TH = os.path.join(INPUT_DIR, "THUC HIEN")
SCAN_BBLM = os.path.join(INPUT_DIR, "0.  BBLM .pdf")
PHOTO_USE = os.path.join(SRC_TH, "D138304 - PHOTO", "USE")
PHOTO_OPT = os.path.join(SRC_TH, "D138304 - PHOTO", "_optimized")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH United International Pharma",                              # BBLM
    name_en="United International Pharma Company Limited",                          # AI dịch
    address_vn="Số 16 đường số 7, Khu công nghiệp Việt Nam - Singapore II, phường Bình Dương, Thành phố Hồ Chí Minh",
    address_en="No. 16, Street No. 7, Vietnam - Singapore Industrial Park II, Binh Duong ward, Ho Chi Minh City",
)
LOCATION_VN = f"Nhà máy của {CLIENT.name_vn} - {CLIENT.address_vn}"   # BBLM mục 5 (cùng địa chỉ mục 3)
LOCATION_EN = f"Factory of {CLIENT.name_en} - {CLIENT.address_en}"
EXECUTOR = "Nghĩa"                 # BBLM: Phạm Quốc Nghĩa
SYM = "N2"                         # BBLM
POINT = {"name_vn": "Bình khí Nitrogen có số QC: VN-261004",
         "name_en": "Nitrogen gas cylinder, QC No.: VN-261004"}
CO_RESULT = "< 5"                  # ppm — BBLM viết tay
METHOD_VN = "Ống chỉ thị Kitagawa 600SP"          # BBLM mục 7 (người dùng 18/11: "ghi theo BBLM", tô vàng)
METHOD_EN = "Kitagawa 600SP detector tube"
EQUIP = Equipment(name_vn="Ống chỉ thị Kitagawa 600SP", name_en="Kitagawa 600SP detector tube")   # BBLM mục 7


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


# ---------------------------------------------------------------------------
# Dossier — dựng qua phiếu khí nén tạm (chỉ để đọc cấu trúc, không xuất phiếu: đo trực tiếp, không tính toán)
# ---------------------------------------------------------------------------
def make_dossier():
    import tempfile
    tmp = os.path.join(tempfile.mkdtemp(prefix="kn138304_"), f"{CODE} - {SYM} - tmp.xlsx")
    build_compressed_air_workbook({
        "code": CODE, "point_label": SYM, "cert_date": CERT_DATE,
        "location_name": POINT["name_vn"], "location_symbol": SYM,
        "dewpoint": {"temp_start": 25, "temp_end": 25},
        "oil": {"flow_aerosol": 35, "flow_vapor": 1, "time_aerosol": 240, "time_vapor": 240},
    }, tmp)
    d = create_dossier_from_compressed_air_workbooks(
        [tmp], code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        client=CLIENT, personnel=PersonnelInfo(executor=EXECUTOR), equipments=[EQUIP],
        scope_vn="Giám định chất lượng khí nitơ", scope_en="Nitrogen gas quality inspection",
        standards_vn="-", standards_en="-", basis_vn="-", basis_en="-",
    )
    d.reg_date = REG_DATE
    d.procedure_vn = d.procedure_en = "QUATEST 3 3030:2025, N2.TC.KN.01"      # BBLM mục 8
    d.location_desc_vn, d.location_desc_en = LOCATION_VN, LOCATION_EN
    loc = d.locations[0]
    loc.name_vn, loc.name_en = POINT["name_vn"], POINT["name_en"]
    d.photos = [{"path": _jpg(os.path.join(PHOTO_USE, f), PHOTO_OPT), "caption_vn": "", "caption_en": ""}
                for f in sorted(os.listdir(PHOTO_USE)) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    rd = d.results_data
    rd["factory_status_vn"], rd["factory_status_en"] = "Đang hoạt động sản xuất", "In operation"   # BBLM mục 9
    rd["particles"] = {"rows": []}
    rd["dewpoint"] = {"rows": []}
    for r in rd["oil"]["rows"]:
        r["val"], r["class"] = CO_RESULT, METHOD_VN
    rd["kqtn_by_location"] = {SYM: f"- Hàm lượng CO: {CO_RESULT} ppm, đo tại hiện trường bằng ống chỉ thị Kitagawa 600SP"
                                   "."}
    rd["kqtn"] = []
    rd["bckq_label"] = "khi nen"
    d.input_dir, d.output_dir, d.network_dir = INPUT_DIR, "", ""
    d.thuchien_dir = os.path.join(INPUT_DIR, "ThucHien")
    d.tbkq_dir = os.path.join(INPUT_DIR, "TBKQ")
    d.edoc_dir = os.path.join(d.tbkq_dir, "Edoc")
    d.pdf_dir = os.path.join(INPUT_DIR, "pdf")
    return d


# ---------------------------------------------------------------------------
# TBKQ — mục tổng dầu -> mục CO; bỏ mô tả / trạng thái hệ thống khí nén
# ---------------------------------------------------------------------------
def _drop_system(doc, is_en):
    body = doc.element.body
    keys = ("Description of compressed air system", "Diagram of compressed air system", "Status of compressed air system") \
        if is_en else ("Mô tả hệ thống khí nén", "Sơ đồ hệ thống khí nén", "Trạng thái hệ thống khí nén")
    for p in list(body.iterchildren(qn("w:p"))):
        if _ptext(p).strip().startswith(keys):
            nxt = p.getnext()
            body.remove(p)
            if nxt is not None and nxt.tag == qn("w:p") and PLACEHOLDER in _ptext(nxt):
                body.remove(nxt)


def _co_from_oil(doc, is_en):
    body = doc.element.body
    oil_param = "Total oil content." if is_en else "Hàm lượng tổng dầu."
    co_param = "Carbon monoxide (CO) content." if is_en else "Hàm lượng cacbon monoxit (CO)."
    for p in list(body.iter(qn("w:p"))):
        if _ptext(p).strip() == oil_param:
            _set_text(p, co_param)
    tbls = list(body.iterchildren(qn("w:tbl")))
    oil = next(t for t in tbls if ("Total oil content, mg/m" if is_en else "Hàm lượng tổng dầu, mg/m") in _ptext(t))
    head = _prev_text_p(oil)
    m = re.match(r"^(9\.2\.\d\s+)", _ptext(head).strip())
    _set_text(head, (m.group(1) if m else "") + co_param[:-1])
    rows = oil.findall(qn("w:tr"))
    h = rows[0].findall(qn("w:tc"))
    _set_cell(h[0], "Location" if is_en else "Vị trí thực hiện")
    _set_cell(h[1], "Carbon monoxide (CO) content, ppm" if is_en else "Hàm lượng cacbon monoxit (CO), ppm")
    _set_cell(h[2], "Method" if is_en else "Phương pháp")
    c = rows[1].findall(qn("w:tc"))
    _set_cell(c[0], SYM)
    _set_cell(c[1], CO_RESULT)
    _set_cell(c[2], METHOD_EN if is_en else METHOD_VN)
    for r in c[2].iter(qn("w:r")):                    # người dùng 18/11/2026: "nhớ tô vàng"
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        hl = OxmlElement("w:highlight")
        hl.set(qn("w:val"), "yellow")
        rpr.append(hl)
    for r in rows[2:]:
        oil.remove(r)
    nxt = oil.getnext()
    while nxt is not None and not _is_note_table(nxt):
        nxt = nxt.getnext()
    note = note_rows(nxt, [("●", "The performance field is accredited to ISO 17020 - VIAS 004." if is_en else
                                 "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.")])
    nxt.addnext(note)
    nxt.getparent().remove(nxt)


def build_tbkq(d, is_en: bool) -> str:
    src = CompressedAirHandler(d).build_tbkq(is_en=is_en)
    doc = Document(src)
    _drop_system(doc, is_en)
    _co_from_oil(doc, is_en)
    _fix_cover_lists(doc)
    _keep_headings_with_tables(doc)
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    style_note_tables(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - khi nito - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    print(f"[build_138304] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


def build_bckq(d) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - khi nito - {CODE}.docx")
    doc = Document(out)
    body = doc.element.body
    # không có phiếu tính toán (đọc trực tiếp ống chỉ thị) -> viện dẫn phiếu ghi chép (AGENTS mục 7)
    _replace_in_paragraphs(body, {f"phiếu Kết quả tính toán số {CODE} - {SYM} ngày {CERT_DATE}":
                                  f"phiếu Ghi chép giám định tại hiện trường số {CODE} ngày {SAMPLE_DATE}"})
    _replace_in_paragraphs(body, {"đo đạc khí nén": "đo đạc khí nitơ", "Hàm lượng tổng dầu.": "Hàm lượng cacbon monoxit (CO).",
                                  "Hàm lượng tổng dầu": "Hàm lượng cacbon monoxit (CO)"})
    for p in list(body.iterchildren(qn("w:p"))):      # sơ đồ hệ thống khí nén: không áp dụng (bình khí)
        if _ptext(p).strip().startswith("- Sơ đồ hệ thống"):
            nxt = p.getnext()
            body.remove(p)
            if nxt is not None and nxt.tag == qn("w:p") and (PLACEHOLDER in _ptext(nxt) or not _ptext(nxt).strip()):
                body.remove(nxt)
    # Không gửi mẫu PTN (đo tại hiện trường): khung KQTN thay bằng một dòng
    kids = list(body.iterchildren(qn("w:p")))
    i0 = next(i for i, p in enumerate(kids) if _ptext(p).strip().startswith("Gửi mẫu thử nghiệm"))
    i1 = next(i for i, p in enumerate(kids) if _ptext(p).strip().startswith("Báo cáo của GĐV"))
    block = kids[i0 + 1:i1]
    if block:
        _set_text(block[0], "Không gửi mẫu thử nghiệm (đo tại hiện trường bằng ống chỉ thị).")
        for p in block[1:]:
            body.remove(p)
    _keep_headings_with_tables(doc)
    for t in doc.tables[1:]:
        if len(t.rows) > 1 and t.rows[0].cells[0].text.strip().startswith("Vị trí"):
            add_tbl_header(t.rows[0])
    _highlight_placeholders(doc)
    doc.save(out)
    print(f"[build_138304] BCKQ -> {out}")
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
    return [p]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    args = ap.parse_args()
    d = make_dossier()
    assign_cert_numbers([d])
    d.ensure_output_directories()
    split_edoc()
    docs = [build_bckq(d), build_tbkq(d, False), build_tbkq(d, True)]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(d).export_files(docs)
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
