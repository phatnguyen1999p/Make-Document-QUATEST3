# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D141835 - Công ty TNHH Quốc tế Sampure Việt Nam — lấy mẫu nước siêu sạch (sau RO + EDI), 1 chứng thư
(mã trần), kết quả theo KQTN.

Nguồn (Z: Q326D141835 - QUOC TE SAMPURE, chép về Input/141835 ngày 16/11/2026, chỉ đọc):
- `0. BBLM .pdf` (1 trang, BM3030.02): PYC Q326D141835 ngày 06/11/2026; đơn vị yêu cầu CÔNG TY TNHH QUỐC TẾ
  SAMPURE VIỆT NAM, Tầng 6, số 107 Nguyễn Phong Sắc, phường Cầu Giấy, Thành phố Hà Nội; nội dung "Giám định chất lượng
  nước RO"; địa điểm CÔNG TY TNHH ELECTRONIC TRIPOD VIỆT NAM (CHÂU ĐỨC) tại Lô 40, Đường Đ.10, Khu công nghiệp Châu Đức,
  Xã Nghĩa Thành, Thành phố Hồ Chí Minh; thời gian ngày 14/11/2026 từ 9h25 đến (để trống); thiết bị Túi vô trùng
  Whirl Pak; quy trình QUATEST3 3030:2025, N2.TC.N.03; chỉ tiêu theo báo giá QONV26012970 (không có trong hồ sơ);
  bảng: vị trí "ống dẫn nước đầu ra EDI (9h25)", loại mẫu "Nước siêu tinh khiết", lấy mẫu, 01 vị trí, ký hiệu mẫu
  "Nước siêu sạch"; ghi tay "Theo quy định Công ty, hình ảnh do người nhà máy chụp và gửi lại."; người thực hiện
  Phạm Quốc Nghĩa -> GĐV ký Phạm Quốc Nghĩa, người báo cáo Nghĩa + Minh.
- `THUC HIEN/BM3030.01- ghi chep -  .docx`, `BM3030.02- BBLM -  .docx`: bản in trắng (chưa ghi số liệu) — không có scan
  phiếu ghi chép.
- `0.  NUOC SIEU SACH .jpg` (ảnh vị trí, có chữ "Nước siêu sạch"), `THUC HIEN/D141835 - PHOTO/*.jpg` (5 ảnh Timemark
  14/11/2026 09:24 – 09:26).
- `2. BCKQ  .docx` (Q326D131622 Metkraft) và `3. TBKQ.docx` (Q326D092610 BSE): mẫu người dùng để sẵn trên Z — chỉ tham
  khảo bố cục (bảng Stt | Chỉ tiêu | Đơn vị | Phương pháp | Kết quả), không lấy số liệu.
- Chưa có KQTN, chưa có PYC (mục 15 ngôn ngữ chứng thư), chưa có báo giá (danh mục chỉ tiêu).

Quyết định của người dùng (17/11/2026):
- Chỉ tiêu: giữ nguyên (chờ báo giá / KQTN). Dựng cả VN + EN.
- Giờ kết thúc 09:30, tên EN "Sampure Vietnam International Company Limited", ký hiệu mẫu "Nước siêu sạch" /
  "Ultrapure water", nội dung "Giám định chất lượng nước RO": giữ, tô vàng.
- Ảnh giữ nguyên. Ngày yêu cầu cấp KQTN: để trống, tô vàng.
- Câu công nhận: VIMCERTS 078, tô vàng.

Suy luận của AI (cần người dùng xác nhận):
- Giờ kết thúc lấy mẫu 09:30 (BBLM để trống; ảnh cuối 09:26).
- Tên EN khách hàng "Sampure Vietnam International Company Limited" (AI tự dịch — chưa có MST để tra tên quốc tế).
- Tên EN địa điểm, nội dung yêu cầu "Giám định chất lượng nước RO" (theo BBLM mục 4), Căn cứ "-" (chưa biết mức).
- Dựng cả TBKQ VN + EN (chưa có PYC mục 15).
- Ảnh: 1 ảnh vị trí + 5 ảnh lấy mẫu (lưới 3 cột, không chú thích).
- Không có phiếu tính toán Excel: kết quả ghi theo KQTN (AGENTS mục 13), như hồ sơ 137058.

Khi có KQTN: điền CHI_TIEU (tên VN, tên EN, đơn vị, phương pháp, kết quả), KQTN_SO, KQTN_NGAY, NGAY_NHAN_MAU,
NGAY_NHAN_KQTN rồi chạy lại.

Chạy:  python build_141835.py          (DOCX + Edoc)
       python build_141835.py --pdf    (thêm PDF qua Word COM + UpChungThu, chỉ Windows)
       python build_141835.py --chi-vn | --chi-en   (chỉ dựng một bản TBKQ)
Không đồng bộ ổ mạng Z trong code — chép bằng `python Core/Tools/ai_files.py chep 141835 --ghi` (AGENTS mục 0).
"""

import os
import re
import sys
import copy
import glob
import argparse
from datetime import date, datetime, timedelta

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", r"C:\__khong_dong_bo_Z__")

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER
from Core.Styling.elements import add_document_ending_line
from Core.Styling.styles import set_cell_shading
from Core.Styling.tables import (format_table_structure, set_table_autofit_window, insert_image_to_cell,
                                 add_tbl_header, add_cant_split, PHOTO_WIDTH_3COL)
from Core.Tools.qa_check import check_single_docx

CODE = "Q326D141835"
CODE6 = "141835"
REG_DATE = "06/11/2026"            # BBLM mục 1
SAMPLE_DATE = "14/11/2026"         # BBLM mục 6
START, END = "09:25", "09:30"      # BBLM mục 6 (giờ kết thúc để trống -> suy luận theo ảnh, cần xác nhận)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN_BBLM = os.path.join(INPUT_DIR, "0. BBLM .pdf")
PHOTO_LOC = os.path.join(INPUT_DIR, "0.  NUOC SIEU SACH .jpg")
PHOTO_DIR = os.path.join(INPUT_DIR, "THUC HIEN", "D141835 - PHOTO")

# --- KQTN (chưa có) ---------------------------------------------------------------------------------------------
KQTN_SO = ""              # vd "HDTN26xxxxxx"
KQTN_NGAY = ""            # dd/mm/yyyy — cũng là ngày ban hành TBKQ (AGENTS 2)
NGAY_NHAN_MAU = ""        # "Ngày nhận mẫu" trên phiếu KQTN (= ngày gửi mẫu, AGENTS 13)
NGAY_NHAN_KQTN = ""
# PYC chưa có -> "Ngày yêu cầu cấp KQTN" tạm = ngày lấy mẫu + 8 ngày (AGENTS 13)
NGAY_YEU_CAU_KQTN = PLACEHOLDER        # người dùng chốt 17/11/2026: để trống, tô vàng
# Chỉ tiêu theo báo giá QONV26012970 / KQTN: (tên VN, tên EN, đơn vị, phương pháp thử, kết quả VN). Trống -> 3 dòng "………".
CHI_TIEU = []
N_PLACEHOLDER_ROWS = 3
CERT_DATE = KQTN_NGAY or date.today().strftime("%d/%m/%Y")   # chưa KQTN -> ngày dựng (AGENTS mục 2, 16/11/2026)

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Quốc tế Sampure Việt Nam",
    name_en="Sampure Vietnam International Company Limited",      # AI tự dịch (chưa có MST)
    address_vn="Tầng 6, số 107 Nguyễn Phong Sắc, phường Cầu Giấy, Thành phố Hà Nội",
    address_en="6th Floor, No. 107 Nguyen Phong Sac, Cau Giay ward, Ha Noi City",
)
LOCATION_VN = ("Công ty TNHH Electronic Tripod Việt Nam (Châu Đức) tại Lô 40, đường Đ.10, Khu công nghiệp Châu Đức, "
               "xã Nghĩa Thành, Thành phố Hồ Chí Minh")
LOCATION_EN = ("Electronic Tripod Vietnam Company Limited (Chau Duc) at Lot 40, D.10 street, Chau Duc Industrial Park, "
               "Nghia Thanh commune, Ho Chi Minh City")
SCOPE_VN = "Giám định chất lượng nước RO"          # BBLM mục 4
SCOPE_EN = "Inspection of RO water quality"
PROCEDURE = "QUATEST 3 3030:2025, N2.TC.N.03"      # BBLM mục 8
EQUIP_VN = "Túi vô trùng Whirl-Pak"                # vật tư lấy mẫu (BBLM mục 7), ghi thẳng (AGENTS 14)
EQUIP_EN = "Whirl-Pak sterile sampling bag"

SAMPLE_POS_VN = "Ống dẫn nước đầu ra EDI"          # BBLM bảng
SAMPLE_POS_EN = "EDI outlet water pipe"
SAMPLE_TYPE_VN = "Nước siêu tinh khiết"
SAMPLE_TYPE_EN = "Ultrapure water"
SAMPLE_CODE = "Nước siêu sạch"                     # ký hiệu mẫu theo BBLM (AGENTS 13)
SAMPLE_CODE_EN = "Ultrapure water"
RES_ROW_TWIPS = 440


# ---------------------------------------------------------------------------
# Tiện ích
# ---------------------------------------------------------------------------
def en_num(s: str) -> str:
    """Dấu thập phân EN (AGENTS mục 6): '2,1 × 10²' -> '2.1 × 10²'."""
    return re.sub(r"(\d),(\d)", r"\1.\2", s)


def _ptext(el):
    import unicodedata          # template BCKQ dùng dấu tổ hợp (NFD) -> so chữ theo NFC
    return unicodedata.normalize("NFC", "".join(t.text or "" for t in el.iter(qn("w:t"))))


def _hl(run):
    rpr = run._r.get_or_add_rPr()
    if rpr.find(qn("w:highlight")) is None:
        h = OxmlElement("w:highlight")
        h.set(qn("w:val"), "yellow")
        rpr.append(h)


def add_runs(p, text, size, bold=False, italic=False):
    """Ghi chữ vào đoạn; chỉ phần '………' được tô vàng (AGENTS mục 13)."""
    for i, piece in enumerate(text.split(PLACEHOLDER)):
        if i:
            r = p.add_run(PLACEHOLDER)
            r.font.size, r.bold, r.italic = Pt(size), bold, italic
            r.font.name = "Times New Roman"
            _hl(r)
        if piece:
            r = p.add_run(piece)
            r.font.size, r.bold, r.italic = Pt(size), bold, italic
            r.font.name = "Times New Roman"


def _hl_only_placeholders(p):
    """Run mẫu tô vàng cả câu ('Ngày báo cáo: ………') -> tách run, chỉ '………' tô vàng."""
    for r in list(p.runs):
        if PLACEHOLDER not in r.text:
            continue
        rpr = r._r.find(qn("w:rPr"))
        if rpr is not None:
            for h in rpr.findall(qn("w:highlight")):
                rpr.remove(h)
        pieces = r.text.split(PLACEHOLDER)
        r.text = pieces[0]
        anchor = r._r
        for piece in pieces[1:]:
            for txt, hl in ((PLACEHOLDER, True), (piece, False)):
                if not txt:
                    continue
                el = copy.deepcopy(r._r)
                anchor.addnext(el)
                anchor = el
                from docx.text.run import Run
                nr = Run(el, p)
                nr.text = txt
                if hl:
                    _hl(nr)
        if not pieces[0]:
            r._r.getparent().remove(r._r)


def _highlight_header_placeholders(doc):
    for sec in doc.sections:
        for part in (sec.header, sec.first_page_header, sec.footer):
            if part.is_linked_to_previous:
                continue
            for r in part._element.iter(qn("w:r")):
                t = r.find(qn("w:t"))
                if t is not None and PLACEHOLDER in (t.text or ""):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr")
                        r.insert(0, rpr)
                    if rpr.find(qn("w:highlight")) is None:
                        h = OxmlElement("w:highlight")
                        h.set(qn("w:val"), "yellow")
                        rpr.append(h)


class Body:
    """Ghi tuần tự các khối vào cuối thân văn bản (trước sectPr cuối)."""

    def __init__(self, doc, text_proto, size):
        self.doc, self.proto, self.size = doc, text_proto, size
        self.sect = doc.element.body.find(qn("w:sectPr"))

    def emit(self, el):
        self.sect.addprevious(el)
        return el

    def para(self, text="", bold=False, italic=False, indent_cm=None, keep_next=False, space_before=None):
        el = copy.deepcopy(self.proto)
        for ch in [c for c in el if c.tag != qn("w:pPr")]:
            el.remove(ch)
        self.emit(el)
        p = Paragraph(el, self.doc._body)
        add_runs(p, text, self.size, bold, italic)
        pf = p.paragraph_format
        if indent_cm is not None:
            from docx.shared import Cm
            pf.left_indent = Cm(indent_cm)
            pf.first_line_indent = Cm(0)
        if keep_next:
            pf.keep_with_next = True
        if space_before is not None:
            pf.space_before = Pt(space_before)
        return p

    def table(self, rows, cols, widths_pct, header_rows=1, size=None, row_height=440):
        t = self.doc.add_table(rows=rows, cols=cols)
        self.emit(t._tbl)
        format_table_structure(t, row_height=row_height)
        # tcMar 40/40 của format_table_structure làm Word cộng thêm ~4,5 pt/dòng (đo PDF 05/11/2026: 24,5 pt thay
        # vì 20 pt) -> bảng 21 mẫu tràn trang. Bỏ, dùng lề ô mặc định của bảng.
        for mar in list(t._tbl.iter(qn("w:tcMar"))):
            mar.getparent().remove(mar)
        grid = t._tbl.tblGrid
        total = 9355
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths_pct):
            gc.set(qn("w:w"), str(round(total * w / 100)))
        for row in t.rows:
            for c, w in zip(row._tr.findall(qn("w:tc")), widths_pct):
                tcPr = c.get_or_add_tcPr()
                old = tcPr.find(qn("w:tcW"))
                if old is not None:
                    tcPr.remove(old)
                tcW = OxmlElement("w:tcW")
                tcW.set(qn("w:w"), str(round(50 * w)))
                tcW.set(qn("w:type"), "pct")
                tcPr.insert(0, tcW)
        for r in t.rows[:header_rows]:
            add_tbl_header(r)
            for c in r.cells:
                set_cell_shading(c, "D9D9D9")
        t._size = size or 12.5
        t._gap = 1 if row_height >= 440 else 0
        return t


def cell_text(table, cell, text, bold=False, italic=False, align="center"):
    for p in list(cell.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = cell.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    for i, line in enumerate(text.split("\n")):
        p = p0 if i == 0 else cell.add_paragraph()
        p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
        pf = p.paragraph_format
        pf.space_before, pf.space_after, pf.line_spacing = Pt(table._gap), Pt(table._gap), 1.0
        add_runs(p, line, table._size, bold, italic)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def note_table(body: Body, proto, lines):
    """Bảng 'Ghi chú/Note' theo mẫu template: dòng tiêu đề + mỗi ghi chú một dòng (marker | nội dung)."""
    el = copy.deepcopy(proto)
    rows = el.findall(qn("w:tr"))
    head, item = rows[0], rows[-1]
    for r in rows[1:]:
        el.remove(r)
    for marker, text in lines:
        tr = copy.deepcopy(item)
        el.append(tr)
        tcs = tr.findall(qn("w:tc"))
        for tc, val in ((tcs[0], marker), (tcs[1], text)):
            ps = tc.findall(qn("w:p"))
            for extra in ps[1:]:
                tc.remove(extra)
            p_el = ps[0]
            runs = p_el.findall(qn("w:r"))
            proto_r = copy.deepcopy(runs[0]) if runs else None
            for r in runs:
                p_el.remove(r)
            for ch in [c for c in p_el if c.tag not in (qn("w:pPr"),)]:
                p_el.remove(ch)
            for i, piece in enumerate(val.split(PLACEHOLDER)):
                for txt, hl in (([(PLACEHOLDER, True)] if i else []) + ([(piece, False)] if piece else [])):
                    r = copy.deepcopy(proto_r)
                    for c in [c for c in r if c.tag != qn("w:rPr")]:
                        r.remove(c)
                    rpr = r.find(qn("w:rPr"))
                    if rpr is not None:
                        for va in rpr.findall(qn("w:vertAlign")):
                            rpr.remove(va)
                    t = OxmlElement("w:t")
                    t.set(qn("xml:space"), "preserve")
                    t.text = txt
                    r.append(t)
                    p_el.append(r)
                    if hl:
                        from docx.text.run import Run
                        _hl(Run(r, None))
    body.emit(el)
    return el


def _appendix_total_from_numpages(doc):
    """Footer phụ lục: SECTIONPAGES -> {= {NUMPAGES} - 1}, cùng cỡ chữ (Core: tag_context.appendix_total_field)."""
    from Core.Builders.tag_context import appendix_total_field
    return appendix_total_field(doc)


def _keep(p_el):
    pPr = p_el.get_or_add_pPr()
    if pPr.find(qn("w:keepNext")) is None:
        pPr.insert(0, OxmlElement("w:keepNext"))


def _keep_table_head(tbl_el, n=4):
    rows = tbl_el.findall(qn("w:tr"))
    for tr in rows[:min(n, len(rows) - 1)]:
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)


def _keep_table_tail(tbl_el, n=2):
    for tr in tbl_el.findall(qn("w:tr"))[-n:]:
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)



def _optimized(path: str, max_w: int = 1000) -> str:
    """Ảnh nén trong <thư mục ảnh>/_optimized (AGENTS mục 1, 3): xoay EXIF, ảnh dọc xoay 90°, không crop."""
    from PIL import Image, ImageOps
    out_dir = os.path.join(os.path.dirname(path), "_optimized")
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0].strip()[:40]
    out = os.path.join(out_dir, stem + ".jpg")
    im = ImageOps.exif_transpose(Image.open(path))
    if im.height > im.width:
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=85)
    return out


def photos() -> list:
    """Ảnh vị trí (có chữ 'Nước siêu sạch') + 5 ảnh lấy mẫu theo giờ chụp (tên file Zalo tăng dần)."""
    return [PHOTO_LOC] + sorted(glob.glob(os.path.join(PHOTO_DIR, "*.jpg")))


def photo_grid(body: Body):
    """Lưới 3 cột 4,8 cm, không chú thích (AGENTS mục 3)."""
    ps = [_optimized(p) for p in photos()]
    ncol = 3
    nrow = (len(ps) + ncol - 1) // ncol
    t = body.doc.add_table(rows=nrow, cols=ncol)
    body.emit(t._tbl)
    set_table_autofit_window(t)
    for i, path in enumerate(ps):
        cell = t.rows[i // ncol].cells[i % ncol]
        insert_image_to_cell(cell, path, width=PHOTO_WIDTH_3COL)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for row in t.rows:
        add_cant_split(row)
    return t


def rows_chi_tieu(is_en: bool) -> list:
    if not CHI_TIEU:
        return [(PLACEHOLDER, PLACEHOLDER, PLACEHOLDER, PLACEHOLDER)] * N_PLACEHOLDER_ROWS
    out = []
    for vn, en, unit, method, kq in CHI_TIEU:
        kq = (kq or PLACEHOLDER)
        out.append((en if is_en else vn, unit, method, en_num(kq) if is_en else kq))
    return out


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def build_dossier() -> Dossier:
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False, dossier_type="cleanroom",   # dùng khung bìa phòng sạch, phụ lục dựng lại
        scope_vn=SCOPE_VN, scope_en=SCOPE_EN,
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn=PROCEDURE, procedure_en=PROCEDURE,
        standards_vn="-", standards_en="-",
        client=CLIENT,
        personnel=PersonnelInfo(executor="Nghĩa"),                   # BBLM: Phạm Quốc Nghĩa
        equipments=[Equipment(name_vn=EQUIP_VN, name_en=EQUIP_EN)],
        locations=[], photos=[], results_data={},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
# TBKQ
# ---------------------------------------------------------------------------
def build_tbkq(d: Dossier, is_en: bool) -> str:
    src = CleanroomHandler(d).build_tbkq(is_en=is_en)
    doc = Document(src)
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    sect_i = next(i for i, el in enumerate(kids) if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None)
    after = kids[sect_i + 1:]
    heading = next(el for el in after if el.tag == qn("w:p") and _ptext(el).strip() in ("Performance content", "Nội dung thực hiện"))
    text_proto = copy.deepcopy(heading.getnext())
    note_proto = copy.deepcopy(next(el for el in after if el.tag == qn("w:tbl") and _ptext(el).strip().startswith(("Note", "Ghi chú"))))
    head_proto = copy.deepcopy(heading)
    for el in after:
        if el.tag != qn("w:sectPr"):
            body_el.remove(el)

    b = Body(doc, text_proto, 13)

    def head(text):
        el = copy.deepcopy(head_proto)
        for r in el.findall(qn("w:r"))[1:]:
            el.remove(r)
        el.find(qn("w:r")).find(qn("w:t")).text = text
        b.emit(el)
        _keep(el)
        return el

    head("Performance content" if is_en else "Nội dung thực hiện")
    if is_en:
        b.para(f"- Sampling location was indicated by {CLIENT.name_en}.")
        b.para("- Performance content and code of sample:", keep_next=True)
        hdr = ["Sampling location", "Sample type", "Code of sample", "Performance content"]
        row = [SAMPLE_POS_EN, SAMPLE_TYPE_EN, SAMPLE_CODE_EN, "Sampling and testing of water quality"]
    else:
        b.para(f"- Vị trí lấy mẫu theo chỉ định của {CLIENT.name_vn}.")
        b.para("- Nội dung thực hiện và ký hiệu mẫu:", keep_next=True)
        hdr = ["Vị trí lấy mẫu", "Loại mẫu", "Ký hiệu mẫu", "Nội dung thực hiện"]
        row = [SAMPLE_POS_VN, SAMPLE_TYPE_VN, SAMPLE_CODE, "Lấy mẫu và thử nghiệm chất lượng nước"]
    t = b.table(2, 4, [28, 22, 20, 30], header_rows=1, size=13, row_height=RES_ROW_TWIPS)
    for j, h in enumerate(hdr):
        cell_text(t, t.rows[0].cells[j], h)
    for j, v in enumerate(row):
        cell_text(t, t.rows[1].cells[j], v)
    _keep_table_head(t._tbl, 2)
    b.para("- Photo of sampling process:" if is_en else "- Hình ảnh quá trình lấy mẫu:", keep_next=True, space_before=6)
    photo_grid(b)

    # 9.2 Kết quả
    head("Inspection result" if is_en else "Kết quả giám định")
    b.para(f"Sample: {SAMPLE_CODE_EN}" if is_en else f"Mẫu: {SAMPLE_CODE}", bold=True, keep_next=True)
    hdr = (["No.", "Parameter", "Unit", "Test method", "Result"] if is_en else
           ["Stt", "Chỉ tiêu", "Đơn vị", "Phương pháp thử", "Kết quả"])
    rows = rows_chi_tieu(is_en)
    t = b.table(1 + len(rows), 5, [8, 36, 14, 24, 18], header_rows=1, size=13, row_height=RES_ROW_TWIPS)
    for j, h in enumerate(hdr):
        cell_text(t, t.rows[0].cells[j], h)
    for i, r in enumerate(rows, start=1):
        cells = t.rows[i].cells
        cell_text(t, cells[0], str(i))
        cell_text(t, cells[1], r[0], align="left")
        for j in range(1, 4):
            cell_text(t, cells[1 + j], r[j])
    _keep_table_head(t._tbl, 5)
    _keep_table_tail(t._tbl, 2)
    # Nước không theo QCVN …/BYT -> câu VIMCERTS 078 dòng đầu, không câu VIAS 004 (AGENTS mục 3, 16/11/2026) — tô vàng
    # chờ người dùng xác nhận cho nước siêu sạch
    notes = [("●", "Parameters are accredited to VIMCERTS 078." if is_en else
                   "Các thông số được chứng nhận VIMCERTS 078.")]
    nt = note_table(b, note_proto, notes)
    for r_el in nt.iter(qn("w:r")):
        if "VIMCERTS" in "".join(x.text or "" for x in r_el.iter(qn("w:t"))):
            rpr = r_el.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r_el.insert(0, rpr)
            h = OxmlElement("w:highlight")
            h.set(qn("w:val"), "yellow")
            rpr.append(h)
    for tr in nt.findall(qn("w:tr"))[:-1]:          # không có Đánh giá sau Ghi chú -> dòng cuối không keepNext
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)

    add_document_ending_line(doc)
    _appendix_total_from_numpages(doc)
    _highlight_header_placeholders(doc)
    from nuoc_common import highlight_text          # điểm người dùng chốt giữ nhưng tô vàng (17/11/2026)
    highlight_text(doc, [CLIENT.name_en, SCOPE_EN, SAMPLE_CODE_EN] if is_en else [SCOPE_VN, SAMPLE_CODE])
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - nuoc - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)          # tên handler phòng sạch (file do code tạo)
    print(f"[build_141835] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


# ---------------------------------------------------------------------------
# BCKQ
# ---------------------------------------------------------------------------
def build_bckq(d: Dossier) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - nuoc - {CODE}.docx")
    doc = Document(out)
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    h1 = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Nội dung thực hiện"))
    h2 = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Kết quả đo đạc"))
    gdv = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Báo cáo của GĐV"))
    text_proto = copy.deepcopy(h1.getnext())
    i1, i2, ig = kids.index(h1), kids.index(h2), kids.index(gdv)
    for el in kids[i1 + 1:i2] + kids[i2 + 1:ig]:
        body_el.remove(el)

    b = Body(doc, text_proto, 13)

    def at(anchor):
        b.sect = anchor
    at(h2)
    b.para(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm được chỉ định bởi đại diện {CLIENT.name_vn}, theo báo giá số "
           "QONV26012970, cụ thể như sau:", keep_next=True)
    t = b.table(2, 4, [30, 22, 20, 28], header_rows=1, size=13)
    for j, h in enumerate(["Vị trí lấy mẫu", "Loại mẫu", "Ký hiệu mẫu", "Thời điểm lấy mẫu"]):
        cell_text(t, t.rows[0].cells[j], h)
    for j, v in enumerate([SAMPLE_POS_VN, SAMPLE_TYPE_VN, SAMPLE_CODE, START]):
        cell_text(t, t.rows[1].cells[j], v)
    _keep_table_head(t._tbl, 2)
    b.para("- Dụng cụ chứa mẫu: túi vô trùng Whirl-Pak; số lượng: 01 mẫu.", space_before=6)
    b.para("- Hình ảnh do nhân viên nhà máy chụp và gửi lại (theo quy định của Công ty, ghi trên BBLM).")
    b.para("Hình ảnh quá trình thực hiện:", keep_next=True)
    photo_grid(b)

    at(gdv)
    kqtn = KQTN_SO or PLACEHOLDER
    kqtn_ngay = KQTN_NGAY or PLACEHOLDER
    b.para(f"- KQTN mẫu {SAMPLE_CODE}: xem tại phiếu KQTN số {kqtn} ngày {kqtn_ngay}.")
    b.para("- Kết quả đo tại hiện trường: -")
    b.para("Gửi mẫu thử nghiệm:", bold=True, keep_next=True, space_before=6)
    b.para(f"Phiếu KQTN số {kqtn} ngày {kqtn_ngay}", keep_next=True)
    b.para(f"Ngày gửi mẫu: {NGAY_NHAN_MAU or PLACEHOLDER}", keep_next=True)
    b.para(f"Ngày yêu cầu cấp KQTN: {NGAY_YEU_CAU_KQTN or PLACEHOLDER}", keep_next=True)
    b.para(f"Ngày nhận KQTN: {NGAY_NHAN_KQTN or PLACEHOLDER}", keep_next=True)
    from Core.Tools.kqtn_late import tre_han_text   # ≤ 10 ngày làm việc từ ngày gửi mẫu: không trễ (AGENTS 13)
    tre = tre_han_text(KQTN_NGAY, NGAY_NHAN_MAU, NGAY_YEU_CAU_KQTN, PLACEHOLDER)
    b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre}")

    _keep(h2)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "Hình ảnh")):
            p.paragraph_format.keep_with_next = True
        _hl_only_placeholders(p)
    from nuoc_common import highlight_text          # người dùng chốt 17/11/2026: tô vàng
    highlight_text(doc, [SCOPE_VN, SAMPLE_CODE, END, PLACEHOLDER])
    doc.save(out)
    print(f"[build_141835] BCKQ -> {out}")
    return out


# ---------------------------------------------------------------------------
def split_edoc() -> list:
    """BBLM = scan 1 trang (AGENTS 8A). Không có scan phiếu ghi chép -> không tạo Ghi Chep.pdf."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    name = f"{CODE6} - BBLM.pdf"
    src = fitz.open(SCAN_BBLM)
    dst = fitz.open()
    dst.insert_pdf(src)
    finish_edoc(dst, INPUT_DIR, name)
    p = os.path.join(edoc, name)
    dst.save(p)
    print(f"[build_141835] Edoc -> {p}")
    return [p]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word COM) + chèn giấy tiêu đề UpChungThu")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--chi-vn", action="store_true", help="Chỉ dựng TBKQ tiếng Việt")
    g.add_argument("--chi-en", action="store_true", help="Chỉ dựng TBKQ tiếng Anh")
    args = ap.parse_args()

    print(f"[build_141835] KQTN: {KQTN_SO or 'CHƯA CÓ'} {KQTN_NGAY} | chỉ tiêu: {len(CHI_TIEU) or 'CHƯA CÓ'}")
    d = build_dossier()
    d.ensure_output_directories()
    split_edoc()
    docs = [build_bckq(d)]
    if not args.chi_en:
        docs.append(build_tbkq(d, is_en=False))
    if not args.chi_vn:
        docs.append(build_tbkq(d, is_en=True))
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(d).export_files(docs)
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=d.code).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
