# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D143038 - Công ty Cổ phần Opella Healthcare Việt Nam (17/11/2026) — 1 chứng thư (mã trần), Việt - Anh:
  Vi sinh không khí 9 vị trí (Buck B30120, 100 L/phút x 2 phút = 200 L, đĩa tổng số) + vi sinh bề mặt 6 vị trí
  (đĩa tiếp xúc Rodac), blank không khí (đĩa 10), blank bề mặt (đĩa 17); chỉ tiêu: tổng số vi sinh vật hiếu khí.
  3 phòng pha chế hóa chất DC 10, DC 35, DC 38; nhà xưởng không sản xuất (nghỉ).
Dựng theo khuôn bộ /2 vi sinh của build_135843.py (khuôn TBKQ K Pack Q326D124796, AGENTS 8C) + phần EN theo
build_137058.py.

Nguồn (Input/143038 — chép từ Z "Q326D143038 - Opella Healthcare Việt Nam - vskk - swab" 17/11/2026, chỉ đọc):
  20261117173611826.pdf — scan hiện trường 6 trang:
      tr. 1–2 BBLM BM3030.02 ngày 17/11/2026: PYC Q326D143038 ngày 09/11/2026; Đơn vị "Công ty Cổ phần Opella
          Healthcare Việt Nam", địa chỉ Lô I-8-2, đường D8, Khu công nghệ cao, phường Tăng Nhơn Phú, TP HCM; địa điểm
          (viết tay) "Nhà máy của Công ty Cổ phần Opella Health Việt Nam - <địa chỉ>"; thực hiện 10:00 – 11:30
          17/11/2026; thiết bị Buck B30120; quy trình QUATEST 3 3030:2025; N2.TC.09; STT 1 vi sinh không khí
          "Phòng pha chế hóa chất DC10, DC35, DC38" 09 mẫu, ký hiệu 1 → 9, 10 Blank; STT 2 vi sinh bề mặt 06 mẫu,
          11 → 16, 17 Blank; chỉ tiêu tổng số vi sinh vật hiếu khí; "Tại thời điểm thực hiện, khu vực nêu trên ở trạng
          thái nghỉ"; ký hiệu phiếu ghi chép Q326D143038/1; người thực hiện Nguyễn Tiến Phát; đại diện KH Nguyễn Phương Thảo.
      tr. 3–6 Ghi chép BM3030.01: thời gian 10:00 – 14:30 (khác BBLM); thiết bị Buck + "Dụng cụ lấy mẫu vi sinh bề
          mặt chuyên dụng"; mỗi phòng: HVAC cấp trên xuống, hồi tường, không sản xuất (nghỉ); APBuck 100 L/p, 2 phút;
          TS lấy mẫu DC 10: 1, 2, 3; DC 35: 4, 5, 6; DC 38: 7, 8, 9, TS blank 10; bề mặt "Đĩa Rodac" (gạch
          10 cm x 10 cm ở DC 10): DC 10 đĩa 11, 12 vách trái, vách phải; DC 35 đĩa 13, 14 vách trước, vách sau;
          DC 38 đĩa 15, 16 (số 15 viết đè, nét giống "10"; sơ đồ ghi 15) vách ngay cửa, vách trước mặt; blank 17;
          sơ đồ tay vị trí +1 +2 +3 mỗi phòng.
  Scan/… Scan PYC.pdf — 16 trang (nhiều trang trắng): phiếu kiểm soát công việc; PYC (tên in "CÔNG TY CỔ PHẦN
      SANOFI VIỆT NAM", mục 15 ngôn ngữ để trống, "Ngày dự kiến trả kết quả" trống); điều kiện chung; báo giá
      QONV26016777 (27/10/2026: 9 vị trí không khí, 6 vị trí bề mặt, DC 10, DC 35, DC 38); PO E00213784 Opella.
  Photos-1-001 (7)/Use — 4 ảnh người dùng chọn (1 bề mặt Rodac, 3 không khí Buck).
  Thuc hien/ — biểu mẫu BM3030.01/.02 in sẵn (file người dùng, không dùng làm nguồn số liệu).
  CHƯA CÓ KQTN -> kết quả, số / ngày KQTN, ngày ban hành để '………' tô vàng.

Quyết định của người dùng (17/11/2026):
  - "ngôn ngữ Việt - Anh nha" -> dựng TBKQ VN + EN.

Suy luận của AI (cần người dùng xác nhận):
  - Ký hiệu TBKQ: không khí K1–K9 (BBLM 1–9), bề mặt BM1–BM6 (BBLM 11–16); BCKQ có cột "Ký hiệu mẫu (BBLM)".
  - Thời gian: BBLM 10:00 – 11:30, ghi chép 10:00 – 14:30 -> BCKQ đang ghi theo BBLM.
  - Tên KH theo BBLM (Opella; PYC in tên cũ Sanofi — đổi tên hiệu lực 09/11/2026, xem 139329); tên EN
    "Opella Healthcare Vietnam Joint Stock Company" như 139329. Địa điểm: BBLM viết "Opella Health", ghi chép
    "Opella Healthcare" -> ghi "Nhà máy của Công ty Cổ phần Opella Healthcare Việt Nam - <địa chỉ>".
  - Bề mặt lấy bằng đĩa tiếp xúc Rodac: đơn vị kết quả UNIT_BM (tạm "CFU/ đĩa") và dòng thiết bị bề mặt
    (tạm theo ghi chép "Dụng cụ lấy mẫu vi sinh bề mặt chuyên dụng") — chờ KQTN / người dùng.
  - Tên vị trí không khí "Vị trí 1/2/3" theo sơ đồ tay (+1 +2 +3); bề mặt theo chữ trên ghi chép.
  - Ngày yêu cầu cấp KQTN = ngày lấy mẫu + 8 = 25/11/2026 (PYC để trống — AGENTS mục 13).
  - 1 báo cáo -> số chứng thư mã trần (BBLM ghi ký hiệu phiếu ghi chép "Q326D143038/1").
  - Ảnh: lưới 2 cột, không chú thích (AGENTS mục 3; chú thích nhóm kiểu K Pack chưa được duyệt).

Chạy:  python build_143038.py          (DOCX + XLSX + Edoc)
       python build_143038.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu)
Khi có KQTN: điền KQTN_VSKK, KQTN_VSBM, giá trị trong AIR / SWAB (CFU/đĩa như phiếu), CERT_DATE = ngày KQTN muộn nhất.
Không đồng bộ Z trong code (network_dir trống) — chép lên Z bằng Core/Tools/ai_files.py (AGENTS mục 0).
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
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_ton_tai_Z"))

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm
from docx.text.paragraph import Paragraph

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, style_note_tables, appendix_total_field
from Core.Styling.elements import add_document_ending_line
from Core.Styling.styles import set_cell_shading
from Core.Styling.tables import (format_table_structure, set_table_autofit_window, insert_image_to_cell,
                                 add_tbl_header, add_cant_split, PHOTO_WIDTH_3COL)
from Core.Tools.cleanroom_calc import CleanroomCalcBook, LEFT
from Core.Tools.kqtn_late import tre_han_text
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D143038"
CODE6 = "143038"
REG_DATE = "09/11/2026"            # PYC / BBLM mục 1
SAMPLE_DATE = "17/11/2026"         # BBLM, ghi chép
START, END = "10:00", "11:30"      # BBLM mục 6 (ghi chép ghi kết thúc 14:30 — chờ người dùng)
CERT_DATE = PLACEHOLDER            # chưa có KQTN -> ngày KQTN muộn nhất khi có (AGENTS mục 2)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
PHOTO_DIR = os.path.join(INPUT_DIR, "Photos-1-001 (7)", "Use")
SCAN_FIELD = os.path.join(INPUT_DIR, "20261117173611826.pdf")          # tr. 1–2 BBLM, 3–6 ghi chép
SCAN_PYC = os.path.join(INPUT_DIR, "Scan", "Q326D143038 - Opella Healthcare Việt Nam - Scan PYC.pdf")
NGAY_YC_KQTN = "25/11/2026"        # TẠM: ngày lấy mẫu + 8 (PYC trống "Ngày dự kiến trả kết quả" — AGENTS mục 13)

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần Opella Healthcare Việt Nam",
    name_en="Opella Healthcare Vietnam Joint Stock Company",          # như 139329 (suy luận)
    address_vn="Lô I-8-2, đường D8, Khu công nghệ cao, phường Tăng Nhơn Phú, Thành phố Hồ Chí Minh",
    address_en="Lot I-8-2, D8 Street, Saigon Hi-Tech Park, Tang Nhon Phu Ward, Ho Chi Minh City",
)
LOCATION_VN = f"Nhà máy của {CLIENT.name_vn} - {CLIENT.address_vn}"
LOCATION_EN = f"Factory of {CLIENT.name_en} - {CLIENT.address_en}"
EXECUTOR = "Phát"                  # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam; người báo cáo Phát + Minh + Nam

# ---------------------------------------------------------------------------
# Số liệu
# ---------------------------------------------------------------------------
ROOMS = {"DC 10": ("Phòng pha chế hóa chất DC 10", "Chemical preparation room DC 10"),
         "DC 35": ("Phòng pha chế hóa chất DC 35", "Chemical preparation room DC 35"),
         "DC 38": ("Phòng pha chế hóa chất DC 38", "Chemical preparation room DC 38")}
VS_FLOW_LPM, VS_MINUTES, BUCK_HOLES = 100, 2, 380    # ghi chép; bảng Conversion Buck B30120 (như 135843)
VS_LIT = VS_FLOW_LPM * VS_MINUTES                    # 200 L
AIR_LOD = round(1000 / VS_LIT)                       # KQTN '< 1' CFU/đĩa -> '< 5' CFU/m³
KQTN_VSKK = {"so": PLACEHOLDER, "ngay": PLACEHOLDER, "ngay_nhan_mau": PLACEHOLDER}
KQTN_VSBM = {"so": PLACEHOLDER, "ngay": PLACEHOLDER, "ngay_nhan_mau": PLACEHOLDER}
# (ký hiệu TBKQ, số mẫu BBLM, phòng, vị trí VN, vị trí EN, KQTN CFU/đĩa — None = chưa có)
AIR = [
    ("K1", 1, "DC 10", "Vị trí 1", "Location 1", None),
    ("K2", 2, "DC 10", "Vị trí 2", "Location 2", None),
    ("K3", 3, "DC 10", "Vị trí 3", "Location 3", None),
    ("K4", 4, "DC 35", "Vị trí 1", "Location 1", None),
    ("K5", 5, "DC 35", "Vị trí 2", "Location 2", None),
    ("K6", 6, "DC 35", "Vị trí 3", "Location 3", None),
    ("K7", 7, "DC 38", "Vị trí 1", "Location 1", None),
    ("K8", 8, "DC 38", "Vị trí 2", "Location 2", None),
    ("K9", 9, "DC 38", "Vị trí 3", "Location 3", None),
]
AIR_BLANK = 10
SWAB = [  # kết quả ghi như KQTN (AGENTS mục 13)
    ("BM1", 11, "DC 10", "Vách trái", "Left wall", None),
    ("BM2", 12, "DC 10", "Vách phải", "Right wall", None),
    ("BM3", 13, "DC 35", "Vách trước", "Front wall", None),
    ("BM4", 14, "DC 35", "Vách sau", "Rear wall", None),
    ("BM5", 15, "DC 38", "Vách ngay cửa ra vào", "Wall next to the door", None),
    ("BM6", 16, "DC 38", "Vách đối diện cửa ra vào", "Wall opposite the door", None),
]
SWAB_BLANK = 17
UNIT_BM, UNIT_BM_EN = "CFU/ đĩa", "CFU/ plate"         # TẠM — đĩa tiếp xúc Rodac; chờ KQTN
BM_LOD = "1"                                          # TẠM — LOD theo KQTN
ITEM_VN, ITEM_EN = "Tổng số vi sinh vật hiếu khí", "Total aerobic plate count"
CONTENT = {"air": ("Lấy mẫu và xác định tổng số vi sinh vật hiếu khí trong không khí",
                   "Sampling and determining total aerobic plate count in workplace air"),
           "bm": ("Lấy mẫu và xác định tổng số vi sinh vật hiếu khí trên bề mặt",
                  "Sampling and determining total aerobic plate count on surface")}
PHOTOS = ["20261117_111021.jpg", "20261117_113458.jpg", "20261117_113506.jpg", "20261117_110630.jpg"]
NOTES = {False: ("Phương pháp thử nghiệm được công nhận ISO 17025 - VILAS 004.",
                 "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004."),
         True: ("The test method is accredited to ISO 17025 - VILAS 004.",
                "The performance field is accredited to ISO 17020 - VIAS 004.")}
EQUIP_BM_VN = "Dụng cụ lấy mẫu vi sinh bề mặt chuyên dụng"            # ghi chép (TẠM)
EQUIP_BM_EN = "Dedicated surface microbial sampling kit"
BCKQ_PHOTO_W2 = Cm(7.0)


def buck_corrected(r: int) -> int:
    """Số khuẩn lạc sau hiệu chỉnh theo bảng Conversion Buck B30120 (N lỗ) — như phiếu 135843."""
    s = 0.0
    for i in range(1, r + 1):
        s = round(s + 1 / (BUCK_HOLES - i + 1), 6)
    return round(s * BUCK_HOLES)


def air_value(cfu, is_en=False) -> str:
    """CFU/đĩa -> CFU/m³; '< 1' -> '< 5'; None -> '………'."""
    if cfu is None:
        return PLACEHOLDER
    if isinstance(cfu, str) or cfu == 0:
        return f"< {AIR_LOD}"
    return str(round(buck_corrected(cfu) * 1000 / VS_LIT))


def bm_value(v, is_en=False) -> str:
    if v is None:
        return PLACEHOLDER
    return v.replace(",", ".") if is_en else v


# ---------------------------------------------------------------------------
# Tiện ích DOCX (theo build_137035 / build_137058)
# ---------------------------------------------------------------------------
def _optimized(path: str, max_w: int = 1200) -> str:
    """Ảnh nén trong Input/134179/Photo/_optimized (AGENTS mục 1, 3): xoay EXIF, ảnh dọc xoay 90°, không crop."""
    from PIL import Image, ImageOps
    out_dir = os.path.join(INPUT_DIR, "Photo", "_optimized")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0].replace(" ", "_") + ".jpg")
    im = ImageOps.exif_transpose(Image.open(path))
    if im.height > im.width:
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=88)
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


def _highlight_placeholders(doc):
    """Tô vàng mọi '………' còn sót (thân văn bản, bảng) để người dùng thấy ô chờ dữ liệu."""
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


def _highlight_header_placeholders(doc):
    for sec in doc.sections:
        for part in (sec.header, sec.first_page_header, sec.footer):
            if part.is_linked_to_previous:
                continue
            for r in part._element.iter(qn("w:r")):
                if any(PLACEHOLDER in (t.text or "") for t in r.findall(qn("w:t"))):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr")
                        r.insert(0, rpr)
                    if rpr.find(qn("w:highlight")) is None:
                        h = OxmlElement("w:highlight")
                        h.set(qn("w:val"), "yellow")
                        rpr.append(h)


def _keep(p_el):
    pPr = p_el.get_or_add_pPr()
    if pPr.find(qn("w:keepNext")) is None:
        pPr.insert(0, OxmlElement("w:keepNext"))


def _keep_rows(rows):
    for tr in rows:
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)


def _keep_table_head(tbl_el, n=4):
    rows = tbl_el.findall(qn("w:tr"))
    _keep_rows(rows[:min(n, len(rows) - 1)])


def _keep_table_tail(tbl_el, n=2):
    _keep_rows(tbl_el.findall(qn("w:tr"))[-n:])


def _is_note_table(el) -> bool:
    return el is not None and el.tag == qn("w:tbl") and _ptext(el).strip().startswith(("Ghi chú", "Note"))


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
        if prev is not None and prev.tag == qn("w:p") and _ptext(prev).strip():
            _keep(prev)
            rows = el.findall(qn("w:tr"))
            _keep_rows(rows[:min(5, len(rows) - 1)])


def _superscript_note_marks(doc):
    """'(n)' đầu dòng trong bảng Ghi chú -> chỉ số trên (AGENTS mục 3)."""
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


def _replace_in_paragraphs(root, mapping: dict) -> int:
    """Thay chuỗi trong từng đoạn (kể cả khi chuỗi nằm vắt qua nhiều run)."""
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


def _set_text(p_el, text: str, superscript_tail: str = ""):
    """Đặt lại chữ của đoạn, giữ định dạng run đầu; `superscript_tail` = phần cuối in chỉ số trên (vd '(1)')."""
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
    """Ô bảng (chép từ bảng mẫu): giữ đoạn đầu có run, bỏ đoạn thừa, ghi chữ mới (mỗi '\\n' một đoạn)."""
    ps = tc.findall(qn("w:p"))
    base = next((p for p in ps if p.findall(qn("w:r"))), ps[0])
    for p in ps:
        if p is not base:
            tc.remove(p)
    lines = text.split("\n")
    cur = base
    for i, line in enumerate(lines):
        if i:
            nxt = copy.deepcopy(base)
            cur.addnext(nxt)
            cur = nxt
        _set_text(cur, line, superscript_tail if i == len(lines) - 1 else "")


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


class Body:
    """Ghi tuần tự các khối vào trước một mốc (mặc định: sectPr cuối thân văn bản)."""

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
            pf.left_indent = Cm(indent_cm)
            pf.first_line_indent = Cm(0)
        if keep_next:
            pf.keep_with_next = True
        if space_before is not None:
            pf.space_before = Pt(space_before)
        return p

    def table(self, rows, cols, widths_pct, header_rows=1, size=13, row_height=440):
        t = self.doc.add_table(rows=rows, cols=cols)
        self.emit(t._tbl)
        format_table_structure(t, row_height=row_height)
        for mar in list(t._tbl.iter(qn("w:tcMar"))):     # E51: tcMar 40/40 làm Word cộng thêm ~4,5 pt/dòng
            mar.getparent().remove(mar)
        total = 9355
        for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths_pct):
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
        t._size = size
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


def photo_grid(body: Body, files, width=PHOTO_WIDTH_3COL):
    """Ảnh hiện trường: lưới 3 cột 4–5 cm (chuẩn 4,8 cm), không chú thích (AGENTS mục 3)."""
    photos = [_optimized(os.path.join(PHOTO_DIR, f)) for f in files]
    ncol = 3
    nrow = (len(photos) + ncol - 1) // ncol
    t = body.doc.add_table(rows=nrow, cols=ncol)
    body.emit(t._tbl)
    set_table_autofit_window(t)
    for i, path in enumerate(photos):
        cell = t.rows[i // ncol].cells[i % ncol]
        insert_image_to_cell(cell, path, width=width)
        for p in cell.paragraphs:
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(3)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for row in t.rows:
        add_cant_split(row)
    return t


def photo_groups(body: Body, groups, width2=Cm(7.8), width3=Cm(5.0)):
    """Ảnh gom theo nhóm như TBKQ mẫu Q326D124796 (K Pack): mỗi nhóm một lưới ảnh + một dòng chú thích nhóm
    (13pt nghiêng, giữa). Nhóm 3 ảnh hoặc từ 5 ảnh: lưới 3 cột (4–5 cm); còn lại lưới 2 cột (7–8 cm) — AGENTS mục 3."""
    for caption, files in groups:
        photos = [_optimized(os.path.join(PHOTO_DIR, f)) for f in files]
        ncol = 3 if (len(photos) == 3 or len(photos) >= 5) else min(2, len(photos))
        width = width3 if ncol == 3 else width2
        t = body.doc.add_table(rows=(len(photos) + ncol - 1) // ncol, cols=ncol)
        body.emit(t._tbl)
        set_table_autofit_window(t)
        for i, path in enumerate(photos):
            cell = t.rows[i // ncol].cells[i % ncol]
            insert_image_to_cell(cell, path, width=width)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(3)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        for row in t.rows:
            add_cant_split(row)
        _keep_rows([t.rows[-1]._tr])            # hàng ảnh cuối đi liền dòng chú thích nhóm
        cap = body.para(caption, italic=True)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = cap.paragraph_format
        pf.left_indent = pf.first_line_indent = Cm(0)
        pf.space_before, pf.space_after = Pt(0), Pt(6)


def unshade(cell):
    """Bỏ nền xám của ô tiêu đề con (tên chỉ tiêu) — như bảng kết quả của TBKQ mẫu Q326D124796."""
    tcPr = cell._tc.get_or_add_tcPr()
    for shd in tcPr.findall(qn("w:shd")):
        tcPr.remove(shd)


def unit_header(table, cell, text, unit):
    """Ô tiêu đề 'Số khuẩn lạc có …, <đơn vị>' — đơn vị in nghiêng như TBKQ mẫu."""
    cell_text(table, cell, text)
    add_runs(cell.paragraphs[0], unit, table._size, italic=True)



def photo_grid2(body: Body, files, width=Cm(7.5)):
    """Ảnh hiện trường: lưới 2 cột 7–8 cm, không chú thích (AGENTS mục 3)."""
    photos = [_optimized(os.path.join(PHOTO_DIR, f)) for f in files]
    t = body.doc.add_table(rows=(len(photos) + 1) // 2, cols=2)
    body.emit(t._tbl)
    set_table_autofit_window(t)
    for i, path in enumerate(photos):
        cell = t.rows[i // 2].cells[i % 2]
        insert_image_to_cell(cell, path, width=width)
        for p in cell.paragraphs:
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(3)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for row in t.rows:
        add_cant_split(row)
    return t


def location_table(b: Body, is_en=False, bckq=False, row_height=440):
    """Bảng 'Khu vực lấy mẫu | Vị trí | Ký hiệu vị trí | Nội dung thực hiện' (khuôn K Pack, có cột khu vực như
    137058); BCKQ thêm cột ký hiệu mẫu theo BBLM."""
    rows = [("air", *r) for r in AIR] + [("bm", *r) for r in SWAB]
    if bckq:
        hdr, widths = ["Khu vực lấy mẫu", "Vị trí", "Ký hiệu vị trí", "Ký hiệu mẫu (BBLM)", "Nội dung thực hiện"], \
            [22, 18, 13, 14, 33]
    elif is_en:
        hdr, widths = ["Sampling area", "Sampling location", "Code of location", "Performance content"], [27, 22, 15, 36]
    else:
        hdr, widths = ["Khu vực lấy mẫu", "Vị trí", "Ký hiệu vị trí", "Nội dung thực hiện"], [27, 22, 15, 36]
    t = b.table(1 + len(rows), len(hdr), widths, header_rows=1, size=13, row_height=row_height)
    for j, h in enumerate(hdr):
        cell_text(t, t.rows[0].cells[j], h)
    ci = len(hdr) - 1
    for i, (kind, sym, no, room, pos_vn, pos_en, _v) in enumerate(rows, start=1):
        cell_text(t, t.rows[i].cells[1], pos_en if is_en else pos_vn, align="left")
        cell_text(t, t.rows[i].cells[2], sym)
        if bckq:
            cell_text(t, t.rows[i].cells[3], str(no))
    # gộp dọc: khu vực theo phòng liên tiếp, nội dung theo loại mẫu
    def merge_runs(key, col, text_of):
        start = 1
        for i in range(2, len(rows) + 2):
            if i == len(rows) + 1 or key(rows[i - 1]) != key(rows[start - 1]):
                c = t.rows[start].cells[col]
                if i - 1 > start:
                    c = c.merge(t.rows[i - 1].cells[col])
                cell_text(t, c, text_of(rows[start - 1]), align="left" if col == 0 else "center")
                start = i
    merge_runs(lambda r: (r[0], r[3]), 0, lambda r: ROOMS[r[3]][1 if is_en else 0])
    merge_runs(lambda r: r[0], ci, lambda r: CONTENT[r[0]][1 if is_en else 0])
    _keep_table_head(t._tbl, 4)
    return t


# ---------------------------------------------------------------------------
def build_dossier() -> Dossier:
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False, dossier_type="cleanroom",   # khung bìa phòng sạch, phụ lục dựng lại
        scope_vn=("- Lấy mẫu và xác định mức nhiễm bẩn vi sinh vật trong không khí;\n"
                  "- Lấy mẫu và xác định mức nhiễm bẩn vi sinh vật trên bề mặt."),
        scope_en=("- Sampling and determining microbial contamination in workplace air;\n"
                  "- Sampling and determining microbial contamination on surface."),
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.09", procedure_en="QUATEST 3 3030:2025, N2.TC.09",
        standards_vn="-", standards_en="-",                         # không có mức quy định
        client=CLIENT, personnel=PersonnelInfo(executor=EXECUTOR),
        equipments=[eq("buck_b30120"), Equipment(name_vn=EQUIP_BM_VN, name_en=EQUIP_BM_EN)],
        locations=[], photos=[], results_data={"bckq_label": "vi sinh"},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


def build_tbkq(d: Dossier, is_en: bool) -> str:
    src = CleanroomHandler(d).build_tbkq(is_en=is_en)
    doc = Document(src)
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    sect_i = next(i for i, el in enumerate(kids)
                  if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None)
    after = kids[sect_i + 1:]
    heading = next(el for el in after if el.tag == qn("w:p")
                   and _ptext(el).strip() in ("Performance content", "Nội dung thực hiện"))
    text_proto = copy.deepcopy(heading.getnext())
    note_proto = copy.deepcopy(next(el for el in after if _is_note_table(el)))
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

    # 9.1 (khuôn K Pack Q326D124796 — AGENTS 8C)
    head("Performance content" if is_en else "Nội dung thực hiện")
    if is_en:
        b.para(f"- Sampling location was indicated by {CLIENT.name_en}.")
        b.para("- Status of factory: Not in operation.")
        b.para("- Performance content and code of location:", keep_next=True)
    else:
        b.para(f"- Khu vực lấy mẫu theo chỉ định của {CLIENT.name_vn}.")
        b.para("- Trạng thái nhà xưởng: không hoạt động sản xuất.")
        b.para("- Nội dung thực hiện và ký hiệu vị trí:", keep_next=True)
    location_table(b, is_en=is_en)
    b.para("- Photo of sampling process:" if is_en else "- Hình ảnh quá trình lấy mẫu:", keep_next=True,
           space_before=6)
    photo_grid2(b, PHOTOS)

    # 9.2
    head("Inspection result" if is_en else "Kết quả giám định")
    item = ITEM_EN if is_en else ITEM_VN
    blocks = [
        ("air", AIR, "Microbiological contamination in workplace air" if is_en else
         "9.2.1 Độ nhiễm bẩn vi sinh trong không khí",
         ("Colony-forming units in air, " if is_en else "Số khuẩn lạc có trong không khí, "), "CFU/m³",
         (f"The result is expressed as “< {AIR_LOD} CFU/m³” when no target organism is detected in the sampled "
          "air volume." if is_en else
          f"Kết quả được biểu thị “< {AIR_LOD} CFU/m³” khi không phát hiện vi sinh vật mục tiêu trên thể tích không "
          "khí được lấy mẫu.")),
        ("bm", SWAB, "Microbiological contamination on surface" if is_en else
         "9.2.2 Độ nhiễm bẩn vi sinh trên bề mặt",
         ("Colony-forming units on surface, " if is_en else "Số khuẩn lạc có trên bề mặt, "),
         UNIT_BM_EN if is_en else UNIT_BM,
         (f"The result is expressed as “< {BM_LOD} {UNIT_BM_EN}” when no target organism is detected on the "
          "sampled surface." if is_en else
          f"Kết quả được biểu thị “< {BM_LOD} {UNIT_BM}” khi không phát hiện vi sinh vật mục tiêu trên bề mặt "
          "được lấy mẫu.")),
    ]
    for n, (kind, data, title, hdr, unit, lod_note) in enumerate(blocks):
        if is_en:
            title = f"9.2.{n + 1} {title}"
        b.para(title, bold=True, keep_next=True, space_before=6 if n else None)
        t = b.table(2 + len(data), 2, [40, 60], header_rows=2, size=13)
        r0 = t.rows[0].cells
        cell_text(t, r0[0].merge(t.rows[1].cells[0]), "Sampling location" if is_en else "Vị trí lấy mẫu")
        unit_header(t, r0[1], hdr, unit)
        cell_text(t, t.rows[1].cells[1], item, italic=True)
        unshade(t.rows[1].cells[1])
        for k, (sym, _no, _room, _pv, _pe, v) in enumerate(data, start=2):
            cell_text(t, t.rows[k].cells[0], sym)
            cell_text(t, t.rows[k].cells[1], air_value(v, is_en) if kind == "air" else bm_value(v, is_en))
        _keep_rows(t._tbl.findall(qn("w:tr")))
        vilas, vias = NOTES[is_en]
        nt = b.emit(note_rows(note_proto, [("●", lod_note), ("●", vilas), ("●", vias)]))
        _keep_rows(nt.findall(qn("w:tr"))[:-1])

    _fix_cover_lists(doc)
    add_document_ending_line(doc)
    appendix_total_field(doc)
    style_note_tables(doc)
    _highlight_placeholders(doc)
    _highlight_header_placeholders(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - vi sinh - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)              # tên handler phòng sạch (file do code tạo)
    print(f"[build_{CODE6}] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


def build_bckq(d: Dossier) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - vi sinh - {CODE}.docx")
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
    b.sect = h2
    n_air, n_bm = len(AIR), len(SWAB)
    b.para(f"- Các vị trí lấy mẫu và chỉ tiêu phân tích được chỉ định bởi đại diện {CLIENT.name_vn}, cụ thể như sau:",
           keep_next=True)
    location_table(b, bckq=True, row_height=400)
    b.para(f"- Vi sinh không khí: thiết bị Buck Bio-Culture B30120, lưu lượng {VS_FLOW_LPM} lít/phút, thời gian "
           f"{VS_MINUTES:02d} phút/đĩa ({VS_LIT} lít; mỗi vị trí 01 đĩa tổng số); 01 mẫu blank (đĩa {AIR_BLANK}). "
           f"Vi sinh bề mặt: đĩa tiếp xúc Rodac, mỗi vị trí 01 đĩa; 01 mẫu blank (đĩa {SWAB_BLANK}).", space_before=6)
    b.para("- Tại thời điểm thực hiện, các khu vực lấy mẫu ở trạng thái nghỉ (không sản xuất); hệ thống HVAC cấp gió "
           "trên xuống, hồi tường.")
    b.para("- Hình ảnh các vị trí lấy mẫu:", keep_next=True)
    photo_grid2(b, PHOTOS, width=BCKQ_PHOTO_W2)

    b.sect = gdv
    b.para("4.6.1 Kết quả lấy mẫu vi sinh không khí", bold=True, keep_next=True)
    t = b.table(2 + n_air, 4, [34, 16, 22, 28], header_rows=1, size=13)
    for j, h in enumerate(["Vị trí lấy mẫu", "Ký hiệu mẫu", "KQTN (CFU/ đĩa)", "KQTT (CFU/ m3)"]):
        cell_text(t, t.rows[0].cells[j], h)
    for k, (sym, no, room, pos, _pe, _v) in enumerate(AIR, start=1):
        cell_text(t, t.rows[k].cells[0], f"{room} - {pos}", align="left")
        cell_text(t, t.rows[k].cells[1], sym)
    cell_text(t, t.rows[n_air + 1].cells[0], "Mẫu Blank", align="left")
    cell_text(t, t.rows[n_air + 1].cells[1], str(AIR_BLANK))
    cell_text(t, t.rows[1].cells[2].merge(t.rows[n_air + 1].cells[2]),
              f"Xem tại phiếu KQTN số {KQTN_VSKK['so']} ngày {KQTN_VSKK['ngay']}")
    cell_text(t, t.rows[1].cells[3].merge(t.rows[n_air].cells[3]),
              f"Xem tại phiếu Tính toán kết quả vi sinh vật trong không khí số {CODE}/Vi sinh ngày {SAMPLE_DATE}")
    _keep_rows(t._tbl.findall(qn("w:tr"))[:-1])
    b.para("- KQTT (CFU/m³) = KQTN sau hiệu chỉnh theo bảng Conversion của thiết bị Buck Bio-Culture B30120 chia thể "
           f"tích không khí được lấy mẫu ({VS_LIT} lít).", space_before=6)
    b.para("- KQTN được biểu thị < 1 CFU/ đĩa → không phát hiện vi sinh vật mục tiêu trong thể tích không khí được "
           f"lấy mẫu ({VS_LIT} lít) ~ < {AIR_LOD} CFU/ m³.")
    b.para(f"- KQTN của mẫu Blank (đĩa {AIR_BLANK}): {PLACEHOLDER}")
    b.para("4.6.2 Kết quả lấy mẫu vi sinh trên bề mặt", bold=True, keep_next=True, space_before=6)
    t = b.table(3, 3, [34, 26, 40], header_rows=1, size=13)
    for j, h in enumerate(["Nội dung", "Ký hiệu mẫu", f"KQTN ({UNIT_BM})"]):
        cell_text(t, t.rows[0].cells[j], h)
    cell_text(t, t.rows[1].cells[0], "Vi sinh bề mặt")
    cell_text(t, t.rows[1].cells[1], f"BM1 → BM{n_bm}")
    cell_text(t, t.rows[2].cells[0], "Mẫu Blank")
    cell_text(t, t.rows[2].cells[1], str(SWAB_BLANK))
    cell_text(t, t.rows[1].cells[2].merge(t.rows[2].cells[2]),
              f"Xem tại phiếu KQTN số {KQTN_VSBM['so']} ngày {KQTN_VSBM['ngay']}")
    _keep_rows(t._tbl.findall(qn("w:tr")))
    b.para(f"- KQTN của mẫu Blank (đĩa {SWAB_BLANK}): {PLACEHOLDER}", space_before=6)
    b.para("Gửi mẫu thử nghiệm:", bold=True, keep_next=True, space_before=6)
    for stt, k in zip("ab", (KQTN_VSKK, KQTN_VSBM)):
        b.para(f"{stt}. Phiếu KQTN số {k['so']} ngày {k['ngay']}", keep_next=True)
        b.para(f"Ngày gửi mẫu: {k['ngay_nhan_mau']}", keep_next=True)
        b.para(f"Ngày yêu cầu cấp KQTN: {NGAY_YC_KQTN}", keep_next=True)
        b.para(f"Ngày nhận KQTN: {k['ngay']}", keep_next=True)
        tre = (PLACEHOLDER if PLACEHOLDER in (k["ngay"], k["ngay_nhan_mau"])
               else tre_han_text(k["ngay"], k["ngay_nhan_mau"], NGAY_YC_KQTN, PLACEHOLDER))
        b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre}")

    _keep(h2)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "- Hình ảnh")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    print(f"[build_{CODE6}] BCKQ -> {out}")
    return out


def build_excel() -> str:
    """Phiếu tính toán vi sinh không khí (chỉ bảng số liệu — AGENTS mục 11); bảng Conversion Buck ở sheet ẩn."""
    sheet = "BuckConversion"
    book = CleanroomCalcBook()
    ws = book.extra_sheet("Vi sinh", so=f"{CODE}/Vi sinh", date=SAMPLE_DATE,
                          widths=(5, 44, 9, 8, 9, 8, 7, 12, 14, 11),
                          title="TÍNH TOÁN KẾT QUẢ VI SINH VẬT TRONG KHÔNG KHÍ", title_rows="1:2")
    for c, lbl in enumerate(("STT", "Vị trí lấy mẫu", "Ký hiệu mẫu", "Số đĩa (BBLM)", "Lưu lượng, L/phút",
                             "Thời gian, phút", "Thể tích, L",
                             "KQTN tổng số VSV hiếu khí, CFU/đĩa",
                             "Sau hiệu chỉnh theo bảng Conversion Buck B30120, CFU/đĩa", "Kết quả, CFU/m³"), 1):
        book.put(ws, 1, c, lbl, bold=True)
    ws.row_dimensions[1].height = 78
    tbl = f"{sheet}!$A$2:$B${BUCK_HOLES + 1}"
    for i, (sym, no, room, pos, _pe, cfu) in enumerate(AIR):
        r = 2 + i
        book.put(ws, r, 1, i + 1)
        book.put(ws, r, 2, f"{ROOMS[room][0]} - {pos}", align=LEFT)
        book.put(ws, r, 3, sym)
        book.put(ws, r, 4, no)
        book.put(ws, r, 5, VS_FLOW_LPM)
        book.put(ws, r, 6, VS_MINUTES)
        book.put(ws, r, 7, f"=E{r}*F{r}")
        book.put(ws, r, 8, cfu if cfu is not None else "")
        book.put(ws, r, 9, f'=IF(H{r}="","",IF(OR(H{r}="< 1",H{r}=0),"< 1",VLOOKUP(H{r},{tbl},2,FALSE)))')
        book.put(ws, r, 10, f'=IF(I{r}="","",IF(I{r}="< 1","< "&ROUND(1000/$G{r},0),ROUND(I{r}*1000/$G{r},0)))',
                 fmt="0")
    wb_ = book.wb.create_sheet(sheet)
    wb_["A1"], wb_["B1"], wb_["C1"], wb_["D1"], wb_["E1"] = "r (CFU/đĩa)", "Pr (sau hiệu chỉnh)", "Σ 1/(N-i+1)", "N (số lỗ)", BUCK_HOLES
    for k in range(1, BUCK_HOLES + 1):
        r = k + 1
        wb_[f"A{r}"] = k
        wb_[f"C{r}"] = "=ROUND(1/$E$1,6)" if k == 1 else f"=ROUND(C{r - 1}+1/($E$1-A{r}+1),6)"
        wb_[f"B{r}"] = f"=ROUND(C{r}*$E$1,0)"
    wb_.sheet_state = "hidden"
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - TinhToanViSinh.xlsx")
    book.save(out)
    print(f"[build_{CODE6}] Excel -> {out}")
    return out


def split_edoc() -> list:
    """BBLM = scan hiện trường tr. 1–2; Ghi chép = tr. 3–6; Khác = scan PYC (PKSCV, PYC, điều kiện chung, báo giá,
    PO; trang trắng bỏ — finish_edoc)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    outs = []
    for name, src, a, z in ((f"{CODE6} - BBLM.pdf", SCAN_FIELD, 0, 1),
                            (f"{CODE6} - Ghi Chep.pdf", SCAN_FIELD, 2, 5),
                            (f"{CODE6} - Khac.pdf", SCAN_PYC, 0, None)):
        s = fitz.open(src)
        dst = fitz.open()
        dst.insert_pdf(s, from_page=a, to_page=(s.page_count - 1 if z is None else z))
        finish_edoc(dst, INPUT_DIR, name)
        p = os.path.join(edoc, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_{CODE6}] Edoc -> {p} ({dst.page_count} tr.)")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()
    d = build_dossier()
    d.ensure_output_directories()
    split_edoc()
    docs = [build_bckq(d), build_tbkq(d, is_en=False), build_tbkq(d, is_en=True)]
    xlsx = build_excel()
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(d).export_files(docs)
        PDFExporter(d).export_excel_files([xlsx])
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
