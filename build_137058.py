# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D137058 - Công ty TNHH Hon Chuan Vietnam — vi sinh bề mặt (swab), 1 chứng thư (mã trần).

Nguồn: Input/137058/20261105160213304.pdf (18 trang: PKSCV 1, điều kiện chung 2, PYC 3–6 (2 bản), báo giá 7–8,
       đơn hàng KH 9, ghi chép BM3030.01 11–12, BBLM BM3030.02 13–14, bảng vị trí swab của KH 15 (phôi P1–P12)
       và 17 (nắp C1–C9); trang 10, 16, 18 trắng), Input/137058/Photo/1..21.jpg (ảnh n = mẫu ký hiệu n).
- BBLM: lấy mẫu 05/11/2026 09:00–11:00; DMCA01 09 mẫu ký hiệu 1 → 9, DMPA01 12 mẫu ký hiệu 10 → 21;
  chỉ tiêu tổng số vi sinh vật hiếu khí, tổng số nấm men - nấm mốc; quy trình QUATEST 3 3030:2025; N2.TC.09;
  người thực hiện Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam, người báo cáo Phát + Minh + Nam.
- Ghi chép: diện tích quét 10 cm × 10 cm, dung dịch BPW, nhà xưởng không sản xuất (nghỉ), "22 mẫu Blank".
- PYC mục 15: chứng thư chỉ chọn Tiếng Anh -> mặc định chỉ dựng TBKQ EN (thêm --vn để dựng bản tiếng Việt).

Người dùng (05/11/2026): kết quả chờ KQTN vi sinh; không hiệu chỉnh Feller (swab bề mặt, không phải lấy mẫu khí chủ động).
-> Không có phiếu tính toán Excel: kết quả CFU/ 100 cm² ghi theo KQTN (AGENTS mục 13).
   Có KQTN: điền KQTN_SO, KQTN_NGAY, NGAY_NHAN_MAU, NGAY_NHAN_KQTN, KET_QUA, BLANK rồi chạy lại.

Địa chỉ (lệch nguồn): BBLM/ghi chép ghi "…phường Hòa Phú, tỉnh Bình Dương" và địa điểm "…phường Vĩnh Tân,
tỉnh Bình Dương"; PYC ghi "P. Bình Dương, TP.Hồ Chí Minh" và "Phường Vĩnh Tân, Thành phố Hồ Chí Minh"
(trùng chứng thư Q326D054991 cùng khách). Người dùng chốt 05/11/2026: GIỮ THEO PYC.

Người dùng (14/11/2026): "Tham khảo các nội dung trong hồ sơ này [TBKQ - 1 Q326D124796, K Pack — vi sinh không khí +
swab], và thực hiện dựng lại hồ sơ Honchuan, Duy Tan" -> phụ lục TBKQ dựng theo khuôn K Pack: 9.1 = 3 dòng dẫn + bảng
"Sampling area | Sampling location | Code of location | Performance content" + ảnh; 9.2 = bảng kết quả chỉ ghi ký hiệu
vị trí; Ghi chú theo thứ tự "< 10" -> VILAS -> VIAS. Bỏ các dòng chỉ tiêu / số lượng mẫu / diện tích quét ở 9.1 (K Pack
không có). Giữ: 21 ảnh theo từng vị trí có nhãn C1…P12 (quyết định 05/11, 12/11), cột khu vực (2 dây chuyền — suy luận
của AI, K Pack không có cột này), "QUATEST 3 3030:2025, N2.TC.09". TBKQ EN 5 trang (trước 4). BCKQ giữ bố cục cũ.

Chạy:  python build_137058.py          (DOCX + Edoc)
       python build_137058.py --pdf    (thêm PDF qua Word COM + UpChungThu, chỉ Windows)
       python build_137058.py --vn     (dựng thêm TBKQ tiếng Việt)
Không đồng bộ ổ mạng Z: trong code (network_dir để trống) — chép lên Z bằng robocopy (AGENTS mục 0).
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

CODE = "Q326D137058"
CODE6 = "137058"
REG_DATE = "28/10/2026"            # PYC + BBLM mục 1
SAMPLE_DATE = "05/11/2026"         # BBLM, ghi chép
START, END = "09:00", "11:00"      # BBLM mục 6
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN = os.path.join(INPUT_DIR, "20261105160213304.pdf")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photo")

# --- KQTN vi sinh: phiếu KQ-Q326D137058-VSTN26028919.pdf (Z TBKQ\, chép về Input 12/11/2026) ----------------
# 22 phiếu con VSTN26028919.01 … .22, ngày 12/11/2026, nhận mẫu 05/11/2026, thử nghiệm 05/11 – 12/11/2026;
# mọi mẫu: nấm men, nấm mốc (ISO 21527-1:2008), VSV hiếu khí (ISO 4833-1:2013/Amd 1:2022) đều "< 10 CFU/mẫu".
# Phiếu chỉ ghi tên mẫu "Vi sinh bề mặt", không ghi vị trí -> ghép .01–.21 ↔ ký hiệu mẫu 1–21 (C1…P12) theo thứ tự
# BBLM, .22 ↔ blank (ký hiệu 22) — NGƯỜI DÙNG XÁC NHẬN 12/11/2026 ("Đúng, giữ nguyên"). Mẫu .14 (P5) KQTN ghi dạng
# mẫu "(1 que + 10 mL) + (1 que + 8 mL)" — giữ nguyên, diện tích quét vẫn ghi 100 cm², không thêm ghi chú.
KQTN_SO = "VSTN26028919"
KQTN_NGAY = "12/11/2026"  # dd/mm/yyyy — cũng là ngày ban hành TBKQ (AGENTS 2)
NGAY_NHAN_MAU = "05/11/2026"   # "Ngày nhận mẫu" trên phiếu KQTN (= ngày gửi mẫu, AGENTS 13)
NGAY_YEU_CAU_KQTN = "13/11/2026"   # PYC để trống "Ngày dự kiến trả kết quả" -> tạm ngày lấy mẫu + 8 (AGENTS 13)
NGAY_NHAN_KQTN = "12/11/2026"
LOD = "10"                # giới hạn phát hiện trên KQTN -> "< 10 CFU/ 100 cm²"
# Kết quả CFU/ 100 cm² ghi như KQTN (vd "< 10", "2,1 × 10²"), mẫu 1..21: (tổng VSV hiếu khí, tổng nấm men, tổng nấm mốc)
# — nấm men và nấm mốc tách 2 cột riêng (người dùng chốt 05/11/2026, như mẫu Q326D079195)
KET_QUA = {n: ("< 10", "< 10", "< 10") for n in range(1, 22)}
BLANK = "< 10 CFU/ mẫu (tổng số vi sinh vật hiếu khí, tổng số nấm men, tổng số nấm mốc)"   # VSTN26028919.22
CERT_DATE = KQTN_NGAY or PLACEHOLDER

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Hon Chuan Vietnam",
    name_en="Hon Chuan Vietnam Company Limited",
    # Theo PYC (người dùng chốt 05/11/2026) — BBLM ghi "phường Hòa Phú, tỉnh Bình Dương"
    address_vn="Số 29 VSIP II, đường số 7, Khu công nghiệp Việt Nam - Singapore II, phường Bình Dương, "
               "Thành phố Hồ Chí Minh",
    address_en="No. 29 VSIP II, street No. 7, Vietnam - Singapore II Industrial Park, Binh Duong ward, "
               "Ho Chi Minh City",
)
LOCATION_VN = ("Nhà máy của Công ty TNHH Hon Chuan Vietnam, số 28 VSIP II-A, đường số 30, Khu công nghiệp "
               "Việt Nam - Singapore II-A, phường Vĩnh Tân, Thành phố Hồ Chí Minh")
LOCATION_EN = ("Factory of Hon Chuan Vietnam Company Limited, No. 28 VSIP II-A, street No. 30, Vietnam - Singapore "
               "II-A Industrial Park, Vinh Tan ward, Ho Chi Minh City")
SCOPE_VN = "Lấy mẫu và xác định mức nhiễm bẩn vi sinh vật trên bề mặt"   # câu chữ theo TBKQ mẫu Q326D124796
SCOPE_EN = "Sampling and determining microbial contamination on surface"
PROCEDURE = "QUATEST 3 3030:2025; N2.TC.09"           # BBLM mục 8 (code chuẩn hóa dấu phẩy)
EQUIP_VN = "Tăm bông tiệt trùng và dung dịch đệm pepton (BPW)"   # ghi chép: dung dịch BPW (như Q326D079195)
EQUIP_EN = "Sterile swab and buffered peptone water (BPW)"

LINES = [  # (khu vực VN, EN, [(ký hiệu, vị trí VN, vị trí EN)]) — bảng vị trí swab của KH (scan tr. 15, 17)
    ("Dây chuyền sản xuất nắp DMCA01", "Closure production line DMCA01", [
        (1, "C1 - Core của khuôn nắp", "C1 - Closure mold core"),
        (2, "C2 - Băng tải sau ép nắp", "C2 - Post-injection closure conveyor"),
        (3, "C3 - Bồn làm mát", "C3 - Cooling hopper"),
        (4, "C4 - Bên trong bồn Hopper", "C4 - Inside hopper"),
        (5, "C5 - Máy gấp nắp Onlytex", "C5 - Onlytex closure folding machine"),
        (6, "C6 - Thanh dẫn xuống camera", "C6 - The bar leads to camera machine"),
        (7, "C7 - Băng tải trong camera tới thùng đóng gói", "C7 - Conveyor inside camera leads to packing machine"),
        (8, "C8 - Thùng giấy đựng nắp thành phẩm", "C8 - Carton box for finished caps"),
        (9, "C9 - Máy lạnh", "C9 - Air-conditioner"),
    ]),
    ("Dây chuyền sản xuất phôi DMPA01", "Preform production line DMPA01", [
        (10, "P1 - Core của khuôn phôi", "P1 - Preform mold core"),
        (11, "P2 - Ống Robot", "P2 - Robot tube"),
        (12, "P3 - Ống thổi gió", "P3 - Cool pick"),
        (13, "P4 - Băng tải sau ép phôi", "P4 - Post-injection preform conveyor"),
        (14, "P5 - Bên trong thùng hopper", "P5 - Inside hopper"),
        (15, "P6 - Con lăn (vị trí máy đảo phôi)", "P6 - Roller (preform aligner position)"),
        (16, "P7 - Thanh dẫn xuống camera", "P7 - The bar leads to the inspection machine"),
        (17, "P8 - Đầu hít phôi", "P8 - Vacuum inhalation head"),
        (18, "P9 - Băng tải tới thùng đóng gói", "P9 - Conveyor leads to packing machine"),
        (19, "P10 - Gàu nhả phôi", "P10 - Funnel for packing"),
        (20, "P11 - Lồng đựng thành phẩm phôi", "P11 - Metal cage for finished preform"),
        (21, "P12 - Máy lạnh", "P12 - Air-conditioner"),
    ]),
]
BLANK_CODE = 22
# Ký hiệu vị trí C1…C9, P1…P12 theo bảng của KH (người dùng chốt 05/11/2026): TBKQ (cột ký hiệu, chú thích ảnh)
# ghi C1…P12; BCKQ giữ thêm ký hiệu mẫu 1 → 21 theo BBLM để đối chiếu phiếu KQTN.
LABEL = {code: pos_vn.split(" - ", 1)[0] for _, _, ss in LINES for code, pos_vn, _ in ss}


def pos_name(pos: str) -> str:
    """'C1 - Core của khuôn nắp' -> 'Core của khuôn nắp'."""
    return pos.split(" - ", 1)[1]
N_SAMPLES = sum(len(s) for _, _, s in LINES)
UNIT = "CFU/ 100 cm²"
# Bố cục theo TBKQ mẫu Q326D124796 (K Pack) — chỉnh theo PDF
LOC_WIDTHS = [16, 40, 14, 30]      # bảng 9.1: khu vực | vị trí | ký hiệu vị trí | nội dung thực hiện (%)
LOC_ROW_TWIPS = 400                # hàng bảng 9.1 (0,7 cm) để bảng 21 vị trí nằm trọn trang phụ lục 1
RES_ROW_TWIPS = 400                # hàng bảng kết quả


# ---------------------------------------------------------------------------
# Tiện ích
# ---------------------------------------------------------------------------
def en_num(s: str) -> str:
    """Dấu thập phân EN (AGENTS mục 6): '2,1 × 10²' -> '2.1 × 10²'."""
    return re.sub(r"(\d),(\d)", r"\1.\2", s)


def result(code: int, idx: int, is_en: bool) -> str:
    v = (KET_QUA.get(code) or (None, None, None))[idx]
    if v in (None, ""):
        return PLACEHOLDER
    v = str(v).strip()
    return en_num(v) if is_en else v


# Ngoại lệ AGENTS mục 3 "không crop" — NGƯỜI DÙNG CHO PHÉP 12/11/2026: ảnh 1–6 (C1–C6; C6 thêm sau) chụp khung đứng
# được crop sang khung ngang 4:3 (thay vì xoay 90°), giữ trọn tay cầm que swab. Giá trị = mép trên vùng cắt (tỉ lệ chiều
# cao ảnh sau xoay EXIF); vùng cắt rộng hết ảnh, cao = rộng × 3/4 (~56 % ảnh dọc). Đã soát từng ảnh 12/11/2026.
CROP_LANDSCAPE_TOP = {1: 0.30, 2: 0.27, 3: 0.25, 4: 0.29, 5: 0.22, 6: 0.05}


def _optimized(path: str, max_w: int = 1000) -> str:
    """Ảnh nén trong <thư mục ảnh>/_optimized (AGENTS mục 1, 3): xoay EXIF, ảnh dọc xoay 90°, không crop
    (trừ ảnh trong CROP_LANDSCAPE_TOP — người dùng cho phép)."""
    from PIL import Image, ImageOps
    out_dir = os.path.join(os.path.dirname(path), "_optimized")
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join(out_dir, stem + ".jpg")
    im = ImageOps.exif_transpose(Image.open(path))
    top = CROP_LANDSCAPE_TOP.get(int(stem)) if stem.isdigit() else None
    if top is not None and im.height > im.width:
        h = round(im.width * 3 / 4)
        y = min(round(top * im.height), im.height - h)
        im = im.crop((0, y, im.width, y + h))
    if im.height > im.width:
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=85)
    return out


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


def photo_grid(body: Body, is_en: bool, caption_size=13):
    """Ảnh 1..21 = mẫu ký hiệu 1..21 (C1…P12); lưới 3 cột 4,8 cm (AGENTS mục 3)."""
    photos = [(n, _optimized(os.path.join(PHOTO_DIR, f"{n}.jpg"))) for n in range(1, N_SAMPLES + 1)]
    ncol = 3
    nrow = (len(photos) + ncol - 1) // ncol
    t = body.doc.add_table(rows=nrow, cols=ncol)
    body.emit(t._tbl)
    set_table_autofit_window(t)
    for i, (n, path) in enumerate(photos):
        cell = t.rows[i // ncol].cells[i % ncol]
        cap = f"Sampling location {LABEL[n]}" if is_en else f"Vị trí lấy mẫu {LABEL[n]}"
        insert_image_to_cell(cell, path, width=PHOTO_WIDTH_3COL, caption=cap, font_size_caption=caption_size)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for row in t.rows:
        add_cant_split(row)
    return t


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
        standards_vn="-", standards_en="-",                         # không có mức quy định
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),                    # BBLM: Nguyễn Tiến Phát -> GĐV Nguyễn Xuân Nam
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

    # Bố cục phụ lục theo TBKQ mẫu Q326D124796 (K Pack) — người dùng yêu cầu 14/11/2026:
    # 9.1 = 3 dòng dẫn + bảng vị trí / ký hiệu vị trí / nội dung thực hiện (thêm cột khu vực: 2 dây chuyền) + ảnh;
    # 9.2 = bảng kết quả chỉ ghi ký hiệu vị trí; Ghi chú: '< 10' -> VILAS -> VIAS.
    head("Performance content" if is_en else "Nội dung thực hiện")
    if is_en:
        b.para(f"- Sampling location was indicated by {CLIENT.name_en}.")
        b.para("- Status of factory: Not in operation.")
        b.para("- Performance content and code of location:", keep_next=True)
        hdr = ["Sampling area", "Sampling location", "Code of location", "Performance content"]
        content = ("Sampling and determining total aerobic plate count, total yeast and total mold on surface "
                   "of equipment")
    else:
        b.para(f"- Khu vực lấy mẫu theo chỉ định của {CLIENT.name_vn}.")
        b.para("- Trạng thái nhà xưởng: không hoạt động sản xuất.")
        b.para("- Nội dung thực hiện và ký hiệu vị trí:", keep_next=True)
        hdr = ["Khu vực lấy mẫu", "Vị trí", "Ký hiệu vị trí", "Nội dung thực hiện"]
        content = ("Lấy mẫu và xác định tổng số vi sinh vật hiếu khí, tổng số nấm men và tổng số nấm mốc trên "
                   "bề mặt thiết bị")
    t = b.table(1 + N_SAMPLES, 4, LOC_WIDTHS, header_rows=1, size=13, row_height=LOC_ROW_TWIPS)
    for j, h in enumerate(hdr):
        cell_text(t, t.rows[0].cells[j], h)
    i = 1
    for line_vn, line_en, samples in LINES:
        first = i
        for code, pos_vn, pos_en in samples:
            cell_text(t, t.rows[i].cells[1], pos_name(pos_en if is_en else pos_vn), align="left")
            cell_text(t, t.rows[i].cells[2], LABEL[code])
            i += 1
        cell_text(t, t.rows[first].cells[0].merge(t.rows[i - 1].cells[0]), line_en if is_en else line_vn)
        cell_text(t, t.rows[first].cells[3].merge(t.rows[i - 1].cells[3]), content)
    _keep_table_head(t._tbl, 4)
    b.para("- Photo of sampling process:" if is_en else "- Hình ảnh quá trình lấy mẫu:", keep_next=True, space_before=6)
    photo_grid(b, is_en)

    # 9.2 Kết quả
    head("Inspection result" if is_en else "Kết quả giám định")
    b.para("Microbiological contamination on equipment surface" if is_en else
           "Độ nhiễm bẩn vi sinh trên bề mặt thiết bị", bold=True, keep_next=True)
    hdr = (["Sampling location", "Colony-forming units on surface, ",
            "Total aerobic plate count", "Total yeast", "Total mold"] if is_en else
           ["Vị trí lấy mẫu", "Số khuẩn lạc có trên bề mặt, ",
            "Tổng số vi sinh vật hiếu khí", "Tổng số nấm men", "Tổng số nấm mốc"])
    t = b.table(2 + N_SAMPLES, 4, [25, 25, 25, 25], header_rows=2, size=13, row_height=RES_ROW_TWIPS)   # 13pt (AGENTS 3, 15)
    r0 = t.rows[0].cells
    cell_text(t, r0[0].merge(t.rows[1].cells[0]), hdr[0])
    m = r0[1].merge(r0[3])
    cell_text(t, m, hdr[1])
    add_runs(m.paragraphs[0], UNIT, t._size, italic=True)        # đơn vị in nghiêng như TBKQ mẫu
    for j in (1, 2, 3):
        c = t.rows[1].cells[j]
        cell_text(t, c, hdr[1 + j], italic=True)
        tcPr = c._tc.get_or_add_tcPr()                           # ô tên chỉ tiêu không nền xám (TBKQ mẫu)
        for shd in tcPr.findall(qn("w:shd")):
            tcPr.remove(shd)
    i = 2
    for _line_vn, _line_en, samples in LINES:
        for code, _pos_vn, _pos_en in samples:
            cells = t.rows[i].cells
            cell_text(t, cells[0], LABEL[code])
            for j in range(3):
                cell_text(t, cells[1 + j], result(code, j, is_en))
            i += 1
    _keep_table_head(t._tbl, 5)
    _keep_table_tail(t._tbl, 3)
    lod = LOD or PLACEHOLDER
    if is_en:
        notes = [("●", f"The result is expressed as “< {en_num(lod)} {UNIT}” when no target organism is detected "
                       "in the sampled surface area."),
                 ("●", "The test method is accredited to ISO 17025 - VILAS 004."),
                 ("●", "The performance field is accredited to ISO 17020 - VIAS 004.")]
    else:
        notes = [("●", f"Kết quả được biểu thị “< {lod} {UNIT}” khi không phát hiện vi sinh vật mục tiêu trên "
                       "diện tích bề mặt được lấy mẫu."),
                 ("●", "Phương pháp thử nghiệm được công nhận ISO 17025 - VILAS 004."),
                 ("●", "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.")]
    nt = note_table(b, note_proto, notes)
    for tr in nt.findall(qn("w:tr"))[:-1]:          # sau bảng Ghi chú không có Đánh giá -> dòng cuối không keepNext
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)

    add_document_ending_line(doc)
    _appendix_total_from_numpages(doc)
    _highlight_header_placeholders(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - vsbm - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)          # tên handler phòng sạch "TBKQ - en - 1 - phong sach - …" (file do code tạo)
    print(f"[build_137058] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


# ---------------------------------------------------------------------------
# BCKQ
# ---------------------------------------------------------------------------
def build_bckq(d: Dossier) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - vsbm - {CODE}.docx")
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
    # Ghi tuần tự: đặt "điểm chèn" tạm trước h2, rồi trước gdv
    def at(anchor):
        b.sect = anchor
    at(h2)
    b.para(f"Vị trí lấy mẫu và chỉ tiêu được chỉ định bởi đại diện {CLIENT.name_vn}, cụ thể như sau:")
    b.para("+ Chỉ tiêu: tổng số vi sinh vật hiếu khí, tổng số nấm men - nấm mốc;", indent_cm=0.5)
    b.para(f"+ Số lượng mẫu: {N_SAMPLES} mẫu và 01 mẫu blank (ký hiệu {BLANK_CODE}); diện tích quét bề mặt mỗi mẫu "
           "10 cm × 10 cm (100 cm²), dung dịch BPW;", indent_cm=0.5)
    b.para("+ Tên gọi, ký hiệu các vị trí lấy mẫu:", indent_cm=0.5, keep_next=True)
    t = b.table(1 + N_SAMPLES, 4, [24, 12, 48, 16], header_rows=1, size=13)
    for j, h in enumerate(["Khu vực lấy mẫu", "Ký hiệu", "Vị trí lấy mẫu", "Ký hiệu mẫu (BBLM)"]):
        cell_text(t, t.rows[0].cells[j], h)
    i = 1
    for line_vn, _, samples in LINES:
        first = i
        for code, pos_vn, _ in samples:
            cell_text(t, t.rows[i].cells[1], LABEL[code])
            cell_text(t, t.rows[i].cells[2], pos_name(pos_vn), align="left")
            cell_text(t, t.rows[i].cells[3], str(code))
            i += 1
        m = t.rows[first].cells[0].merge(t.rows[i - 1].cells[0])
        cell_text(t, m, line_vn)
    _keep_table_head(t._tbl, 4)
    b.para("- Tại thời điểm thực hiện, nhà máy không hoạt động sản xuất.", space_before=6)
    b.para("Hình ảnh quá trình thực hiện:", keep_next=True)
    photo_grid(b, is_en=False)

    at(gdv)
    kqtn = KQTN_SO or PLACEHOLDER
    kqtn_ngay = KQTN_NGAY or PLACEHOLDER
    t = b.table(4, 4, [30, 20, 30, 20], header_rows=1, size=13)
    for j, h in enumerate(["Khu vực lấy mẫu", "Ký hiệu mẫu", "KQTN (CFU/ mẫu)", "Diện tích quét bề mặt"]):
        cell_text(t, t.rows[0].cells[j], h)
    rows = [(LINES[0][0], "1 → 9\n(C1 → C9)", "100 cm²"), (LINES[1][0], "10 → 21\n(P1 → P12)", "100 cm²"), ("Mẫu Blank", str(BLANK_CODE), "-")]
    for k, (area, codes, s) in enumerate(rows, start=1):
        cell_text(t, t.rows[k].cells[0], area)
        cell_text(t, t.rows[k].cells[1], codes)
        cell_text(t, t.rows[k].cells[3], s)
    m = t.rows[1].cells[2].merge(t.rows[3].cells[2])
    cell_text(t, m, f"Xem tại phiếu KQTN số {kqtn}\nngày {kqtn_ngay}")
    b.para(f"- Kết quả (CFU/ 100 cm²) ghi theo KQTN, không hiệu chỉnh (lấy mẫu swab bề mặt, không áp dụng bảng "
           f"hiệu chỉnh Feller).", space_before=6)
    b.para(f"- KQTN biểu thị “< {LOD or PLACEHOLDER} CFU/ mẫu” khi không phát hiện khuẩn lạc trên diện tích quét bề mặt.")
    b.para(f"- KQTN của mẫu Blank: {BLANK or PLACEHOLDER}.")
    b.para("Gửi mẫu thử nghiệm:", bold=True, keep_next=True, space_before=6)
    b.para(f"Phiếu KQTN số {kqtn} ngày {kqtn_ngay}", keep_next=True)
    b.para(f"Ngày gửi mẫu: {NGAY_NHAN_MAU or PLACEHOLDER}", keep_next=True)
    b.para(f"Ngày yêu cầu cấp KQTN: {NGAY_YEU_CAU_KQTN or PLACEHOLDER}", keep_next=True)
    b.para(f"Ngày nhận KQTN: {NGAY_NHAN_KQTN or PLACEHOLDER}", keep_next=True)
    from Core.Tools.kqtn_late import tre_han_text   # ≤ 10 ngày làm việc từ ngày gửi mẫu: không trễ (AGENTS 13)
    tre = tre_han_text(KQTN_NGAY, NGAY_NHAN_MAU, NGAY_YEU_CAU_KQTN, PLACEHOLDER)
    b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre}")

    # Báo cáo GĐV: ngày kết thúc / ngày báo cáo = ngày ban hành (ngày KQTN) — mẫu đã điền CERT_DATE
    _keep(h2)
    for tr in t._tbl.findall(qn("w:tr"))[:-1]:       # bảng KQTN 4 dòng đi liền tiêu đề 4.6
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "Hình ảnh")):
            p.paragraph_format.keep_with_next = True
        _hl_only_placeholders(p)
    doc.save(out)
    print(f"[build_137058] BCKQ -> {out}")
    return out


# ---------------------------------------------------------------------------
def split_edoc() -> list:
    """BBLM = scan tr. 13–14; Ghi chép = tr. 11–12 (BM3030.01) + tr. 15, 17 (bảng vị trí swab của KH) (AGENTS 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(SCAN)
    outs = []
    for name, pages in ((f"{CODE6} - BBLM.pdf", [12, 13]), (f"{CODE6} - Ghi Chep.pdf", [10, 11, 14, 16])):
        dst = fitz.open()
        for pg in pages:
            dst.insert_pdf(src, from_page=pg, to_page=pg)
        p = os.path.join(edoc, name)
        finish_edoc(dst, INPUT_DIR, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_137058] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--vn", action="store_true", help="Dựng thêm TBKQ tiếng Việt (PYC chỉ chọn tiếng Anh)")
    args = ap.parse_args()

    missing = [n for n in range(1, N_SAMPLES + 1) if n not in KET_QUA]
    print(f"[build_137058] KQTN: {KQTN_SO or 'CHƯA CÓ'} {KQTN_NGAY} | thiếu kết quả {len(missing)}/{N_SAMPLES} mẫu")
    d = build_dossier()
    d.ensure_output_directories()
    split_edoc()
    docs = [build_bckq(d), build_tbkq(d, is_en=True)]
    if args.vn:
        docs.append(build_tbkq(d, is_en=False))
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(d).export_files(docs)
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=d.code).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
