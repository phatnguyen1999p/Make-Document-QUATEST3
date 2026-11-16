# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D131671 - Công ty Cổ phần P/S (phòng sạch, 01/11/2026) — 1 chứng thư (mã trần).

Khu vực: "Phòng đóng gói cấp 1" (trạng thái nghỉ) — GMP Grade D theo Phụ lục 2, Thông tư 14/2012/TT-BYT:
  - Nồng độ tiểu phân (hạt bụi): 09 vị trí (AeroTrak 9350, file 520–528; file 519 là zeroing, không tính).
  - Vi sinh vật trong không khí (tổng số vi sinh vật hiếu khí): 05 mẫu (ký hiệu 1–5) + 01 mẫu blank (6),
    Buck B30120 100 L/phút x 1 phút.

Nguồn (chỉ đọc):
  - Input/131671/131671 _ ScanAll.pdf: tr.1–2 BBLM, tr.3–5 phiếu ghi chép (ISO Class 8 in sẵn; mức vi sinh
    ≤ 200 CFU/m³), tr.6 trắng.
  - Input/131671/131671 - Data Do/6ChannelFixed.xlsx (AeroTrak, hạt/m³, Σ).
  - Input/131671/Layout/*.jpg (sơ đồ khách hàng vẽ: 5,2 m x (8,45 - 2,0) m = 33,5 m², cao 2,45 m).
  - Input/131671/Photo/Use/*.jpg (ảnh chèn chứng thư — người dùng chọn 01/11/2026).
  - Input/131671/KQTN/KQ-Q326D131671-VSTN26028555.01 … .06.pdf (chép từ Z 12/11/2026): phiếu KQTN ngày
    11/11/2026, nhận mẫu 01/11/2026, "Vi sinh không khí", tổng số VSV hiếu khí CFU/đĩa:
    .01 = 2,0 x 10⁰; .02 = < 1; .03 = < 1; .04 = 1,0 x 10⁰; .05 = 1,0 x 10⁰; .06 = < 1.
  - Input/131671/_Z_sua_tay_10-11/ (bản người dùng sửa tay trên Z 10/11/2026, chép về 12/11/2026):
    "131671 - TBKQ.docx" (TBKQ VN), "BCKQ - 1 - phong sach - Q326D131671.docx",
    "Q326D131671 - Tinh Toan VSKK.xlsx" (form vi sinh: hiệu chỉnh theo bảng Conversion Buck B30120, 380 lỗ).

Quyết định của người dùng:
  - 01/11/2026: diện tích 33,5 m²; ảnh Photo/Use.
  - 10/11/2026 (sửa tay trên Z — đưa vào script 12/11/2026, AGENTS mục 0):
      * Cấp sạch GMP cấp D, mức theo "Phụ lục 2: Khái niệm và phân loại cấp sạch theo GMP ban hành kèm theo
        Thông tư số 14/2012/TT-BYT ngày 31/8/2012 của Bộ Y tế" (bụi trạng thái nghỉ d ≥ 0,5 µm 3 520 000,
        d ≥ 5,0 µm 29 000 tiểu phân/m³; vi sinh ≤ 200 CFU/m³) — thay ISO 8 / mức khách hàng; Căn cứ thêm TT 14,
        giữ ISO 14644-1:2015 (số vị trí theo bảng A1). Bảng bụi bỏ cột d ≥ 1 µm.
      * TBKQ VN dùng chữ "tiểu phân" (bảng bụi, tiêu đề, câu đánh giá); tiêu đề vi sinh "Kết quả xác định độ nhiễm
        bẩn vi sinh vật trong không khí"; bảng khu vực "- Nồng độ tiểu phân: 09 vị trí", "- Độ nhiễm bẩn vi sinh vật:
        05 vị trí"; tên phòng trong bảng không ngoặc kép; dòng "Khu vực thực hiện …: “Phòng đóng gói cấp 1”" 1 dòng.
      * Nội dung yêu cầu "Quan trắc chất lượng phòng sạch"; Địa điểm "Nhà máy của Công ty Cổ phần P/S, 513C …".
      * Sơ đồ có chú thích "Sơ đồ vị trí đo tiểu phân" / "Sơ đồ vị trí lấy mẫu vi sinh".
      * BCKQ: bảng viện dẫn vi sinh (Vị trí | Ký hiệu mẫu | KQTN | KQTT), dòng "+ KQ < 1 CFU/đĩa ~ < 10 CFU/m³ …",
        "+ KQTN của mẫu Blank …", khối "Gửi mẫu thử nghiệm", Ngày yêu cầu cấp KQTN 08/11/2026.
      * Vi sinh quy đổi theo bảng Conversion Buck B30120 (380 lỗ, Feller) rồi chia thể tích 0,1 m³.
    Khác bản Z có chủ ý (theo AGENTS): "GMP Grade D" (mục 8C, 10/11/2026 19:40) thay "GMP D"/"cấp D"; thời gian
    "Từ ngày 01/11/2026" (mục 2); tên máy chuẩn (mục 14); BCKQ ghi chú "(1)" và Địa điểm ghi đủ tên Công ty Cổ phần
    P/S (bản Z còn sót "Công ty Cổ phần Phòng khám Gia đình …" của hồ sơ khác, Địa điểm chỉ còn địa chỉ — mục 15);
    BCKQ cấp sạch GMP Grade D (bản Z còn ISO 8).
  - 12/11/2026: người dùng xác nhận giữ 3 điểm suy luận cũ —
      (1) KQTN .01 … .05 = mẫu 1 … 5 (vị trí 1 … 5), .06 = blank (6) (phiếu KQTN không ghi ký hiệu mẫu);
      (2) Ngày nhận KQTN = 11/11/2026; KQTN 11/11 = 8 ngày làm việc từ ngày gửi mẫu 01/11 -> "KQTN có trễ hạn: [X] không"
          (người dùng 12/11/2026: "Không đánh trễ kết quả thử nghiệm dưới 10 ngày làm việc" — trước đó "có");
      (3) dựng TBKQ EN tương ứng bản VN.

Suy luận của AI (cần người dùng xác nhận):
  Không còn (đã xác nhận 12/11/2026).

Chạy:  python build_131671.py          (DOCX + XLSX + Edoc)
       python build_131671.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
       python build_131671.py --chi-excel [--pdf]   (chỉ dựng lại phiếu tính toán + sửa dòng viện dẫn BCKQ)
Không đồng bộ ổ mạng Z trong code (network_dir để trống). Z của hồ sơ này dùng thư mục người dùng đặt tên
"Q326D131671 - ThucHien", "Q326D131671 - TBKQ", "pdf" -> chép từng file (AGENTS mục 0).
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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, within, verdict, set_assessment_paragraph, style_note_tables
from Core.Styling.styles import set_cell_shading
from Core.Styling.tables import format_table_structure, add_tbl_header
from Core.Tools.qa_check import check_single_docx
from Core.Tools.cleanroom_calc import CleanroomCalcBook, read_aerotrak, replace_reference_lines, LEFT
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D131671"
CODE6 = "131671"
REG_DATE = "16/10/2026"
SAMPLE_DATE = "01/11/2026"
START, END = "08:15", "09:30"
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN = os.path.join(INPUT_DIR, f"{CODE6} _ ScanAll.pdf")
RAW_XLSX = os.path.join(INPUT_DIR, f"{CODE6} - Data Do", "6ChannelFixed.xlsx")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photo", "Use")
LAYOUT_DIR = os.path.join(INPUT_DIR, "Layout")

# --- KQTN vi sinh (phiếu VSTN26028555.01 … .06) ---
KQTN_SO = "VSTN26028555"
KQTN_NGAY = "11/11/2026"          # ngày phiếu KQTN -> ngày ban hành TBKQ (AGENTS mục 2)
NGAY_GUI_MAU = "01/11/2026"       # "Ngày nhận mẫu" trên phiếu KQTN (AGENTS mục 13)
NGAY_YC_KQTN = "08/11/2026"       # người dùng ghi trên BCKQ Z 10/11/2026
NGAY_NHAN_KQTN = "11/11/2026"     # suy luận (2)
# Số khuẩn lạc CFU/đĩa theo KQTN, mẫu 1..5 (suy luận (1)); "< 1" = không phát hiện
VI_SINH_CFU_DIA = {1: 2, 2: "< 1", 3: "< 1", 4: 1, 5: 1}
VI_SINH_BLANK = "< 1"             # mẫu 6
BLANK_SAMPLE = 6
VS_FLOW_LPM, VS_MINUTES = 100, 1  # Buck 100 L/phút x 1 phút = 0,1 m³ (phiếu ghi chép)
BUCK_HOLES = 380                  # bảng Conversion Buck B30120 (form người dùng "Tinh Toan VSKK.xlsx")

CERT_DATE = KQTN_NGAY or PLACEHOLDER

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần P/S",
    name_en="P/S Joint Stock Company",
    address_vn="513C đường Nam Hòa, phường Phước Long, Thành phố Hồ Chí Minh",
    address_en="513C Nam Hoa Street, Phuoc Long Ward, Ho Chi Minh City",
)
LOCATION_VN = "Nhà máy của Công ty Cổ phần P/S, 513C đường Nam Hòa, phường Phước Long, Thành phố Hồ Chí Minh"
LOCATION_EN = "Factory of P/S Joint Stock Company, 513C Nam Hoa Street, Phuoc Long Ward, Ho Chi Minh City"

ROOM_VN, ROOM_EN = "Phòng đóng gói cấp 1", "Primary packaging room"
ROOM_W, ROOM_D, ROOM_H = 5.2, 6.45, 2.45      # m — theo sơ đồ khách hàng (8 450 - 2 000 mm); 33,5 m² người dùng chốt 01/11/2026
GRADE = "D"
GMP_VN = GMP_EN = f"GMP Grade {GRADE}"        # AGENTS mục 8C (10/11/2026)
LIM05, LIM5 = 3520000, 29000                  # TT 14/2012/TT-BYT Phụ lục 2, cấp D trạng thái nghỉ, tiểu phân/m³
VS_LIMIT = 200                                # CFU/m³, cấp D
BASIS_VN = "Phụ lục 2 kèm theo Thông tư số 14/2012/TT-BYT ngày 31/08/2012"
BASIS_EN = "Appendix 2 attached to Circular No. 14/2012/TT-BYT dated 31/08/2012"
ZERO_FILE = 519
POINT_FILES = list(range(520, 529))           # vị trí 1..9 (phiếu ghi chép trang 2/3)
TRANG_SO_DO = "1"                             # trang phụ lục chứa sơ đồ — kiểm tra lại trên PDF

LABELS, RECS = read_aerotrak(RAW_XLSX)
_miss = [f for f in [ZERO_FILE] + POINT_FILES if f not in RECS]
if _miss:
    raise SystemExit(f"[FATAL] 6ChannelFixed.xlsx thiếu file đo {_miss}")


def dust_ok() -> bool:
    return all(RECS[f]["Ch2"] <= LIM05 and RECS[f]["Ch6"] <= LIM5 for f in POINT_FILES)


def buck_corrected(r: int) -> int:
    """Số khuẩn lạc sau hiệu chỉnh theo bảng Conversion Buck B30120 (Feller, N lỗ) — như form người dùng:
    Pr = ROUND(N × Σ 1/(N - i + 1), 0), tổng cộng dồn làm tròn 6 số lẻ."""
    s = 0.0
    for i in range(1, r + 1):
        s = round(s + 1 / (BUCK_HOLES - i + 1), 6)
    return round(s * BUCK_HOLES)


def vs_value(sample: int):
    """Kết quả CFU/m³ của mẫu (số hoặc chuỗi '< 10'); chưa có -> PLACEHOLDER."""
    vol_m3 = VS_FLOW_LPM * VS_MINUTES / 1000.0
    cfu = VI_SINH_CFU_DIA.get(sample)
    if cfu is None:
        return PLACEHOLDER
    if isinstance(cfu, str) or cfu == 0:
        return f"< {round(1 / vol_m3)}"
    return round(buck_corrected(cfu) / vol_m3)


def vs_ok():
    vals = [vs_value(s) for s in range(1, 6)]
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


def build_dossier() -> Dossier:
    photos = []
    for f in ("20261101_084517.jpg", "20261101_091001.jpg"):
        p = os.path.join(PHOTO_DIR, f)
        if os.path.isfile(p):
            photos.append({"path": _optimized(p), "caption_vn": "", "caption_en": ""})
    # Sơ đồ khách hàng vẽ; chú thích theo bản người dùng sửa trên Z (10/11/2026)
    layouts = [{"path": os.path.join(LAYOUT_DIR, "sơ đồ đo bụi.jpg"),
                "caption_vn": "Sơ đồ vị trí đo tiểu phân", "caption_en": "Particle measurement location layout"},
               {"path": os.path.join(LAYOUT_DIR, "sơ đồ lấy mẫu vi sinh.jpg"),
                "caption_vn": "Sơ đồ vị trí lấy mẫu vi sinh", "caption_en": "Microbial sampling location layout"}]

    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False,
        dossier_type="cleanroom",
        scope_vn="Quan trắc chất lượng phòng sạch",
        scope_en="Cleanroom quality monitoring",
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01, N2.TC.09",
        procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01, N2.TC.09",
        standards_vn=("Phụ lục 2: Khái niệm và phân loại cấp sạch theo GMP ban hành kèm theo Thông tư số "
                      "14/2012/TT-BYT ngày 31/8/2012 của Bộ Y tế\nISO 14644-1:2015"),
        standards_en=("Appendix 2: Concepts and classification of GMP cleanliness grades, issued together with "
                      "Circular No. 14/2012/TT-BYT dated 31/8/2012 of the Ministry of Health\nISO 14644-1:2015"),
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),     # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam
        equipments=[eq("aerotrak_9350"), eq("buck_b30120")],
        locations=[MonitoringLocation(
            name_vn=ROOM_VN, name_en=ROOM_EN, area_m2=f"{ROOM_W * ROOM_D:.1f}", clean_class=GRADE,
            state_vn="Trạng thái nghỉ", state_en="At rest",
            points_summary_vn=["- Nồng độ tiểu phân: 09 vị trí", "- Độ nhiễm bẩn vi sinh vật: 05 vị trí"],
            points_summary_en=["- Particle concentration: 09 locations", "- Microbial contamination: 05 locations"],
        )],
        photos=photos, layouts=layouts,
        results_data={
            "bui_3kenh": [{"cap": GRADE, "gioi_han": [LIM05, "", LIM5], "trang_so_do": TRANG_SO_DO,
                           "diem": [{"khu_vuc": ROOM_VN, "vi_tri": i + 1,
                                     "gia_tri": [RECS[f]["Ch2"], RECS[f]["Ch3"], RECS[f]["Ch6"]]}
                                    for i, f in enumerate(POINT_FILES)],
                           "ket_luan_vn": verdict(dust_ok(), False), "ket_luan_en": verdict(dust_ok(), True)}],
            "vi_sinh": [{"cap": GRADE, "gioi_han": VS_LIMIT, "trang_so_do": TRANG_SO_DO,
                         "diem": [{"khu_vuc": ROOM_VN, "vi_tri": s, "gia_tri": vs_value(s)} for s in range(1, 6)]}],
        },
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX (helper chung theo build_127215.py)
# ---------------------------------------------------------------------------
def _ptext(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _replace_in_paragraph(p_el, old: str, new: str) -> bool:
    """Thay chuỗi trong một đoạn (kể cả khi trải qua nhiều run): chữ mới nằm ở run chứa ký tự đầu."""
    ts = list(p_el.iter(qn("w:t")))
    full = "".join(t.text or "" for t in ts)
    i = full.find(old)
    if i < 0:
        return False
    j = i + len(old)
    pos, done = 0, False
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
    for p in doc.element.body.iter(qn("w:p")):
        guard = 0
        while old in _ptext(p) and _replace_in_paragraph(p, old, new) and guard < 10:
            n += 1
            guard += 1
            if old in new:
                break
    return n


def _set_paragraph_text(p_el, text: str):
    """Ghi lại chữ của đoạn: giữ định dạng run đầu, bỏ các run sau."""
    runs = p_el.findall(qn("w:r"))
    ts = [t for r in runs for t in r.findall(qn("w:t"))]
    if not ts:
        return
    ts[0].text = text
    ts[0].set(qn("xml:space"), "preserve")
    for t in ts[1:]:
        t.text = ""


def _set_tcw(tcW, widths, c0, span):
    w = sum(widths[c0:c0 + span])
    if tcW.get(qn("w:type")) == "pct":
        tcW.set(qn("w:w"), str(round(5000 * w / sum(widths))))
    elif tcW.get(qn("w:type")) == "dxa":
        tcW.set(qn("w:w"), str(w))


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


def _drop_grid_column(tbl, gcol: int):
    """Xóa cột lưới `gcol`; độ rộng cột bỏ: nửa cho cột Khu vực, nửa chia các cột số liệu còn lại."""
    grid = tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    w_drop = int(cols[gcol].get(qn("w:w")))
    grid.remove(cols[gcol])
    cols = grid.findall(qn("w:gridCol"))
    cols[0].set(qn("w:w"), str(int(cols[0].get(qn("w:w"))) + w_drop // 2))
    for k in range(2, len(cols)):
        cols[k].set(qn("w:w"), str(int(cols[k].get(qn("w:w"))) + w_drop // 2 // (len(cols) - 2)))
    for tr in tbl.findall(qn("w:tr")):
        for tc, c0, span in _grid_cells(tr):
            if not (c0 <= gcol < c0 + span):
                continue
            tcPr = tc.find(qn("w:tcPr"))
            if span > 1:
                gs = tcPr.find(qn("w:gridSpan"))
                if span - 1 == 1:
                    tcPr.remove(gs)
                else:
                    gs.set(qn("w:val"), str(span - 1))
            else:
                tr.remove(tc)
            break
    widths = [int(c.get(qn("w:w"))) for c in grid.findall(qn("w:gridCol"))]
    for tr in tbl.findall(qn("w:tr")):
        for tc, c0, span in _grid_cells(tr):
            tcPr = tc.find(qn("w:tcPr"))
            tcW = tcPr.find(qn("w:tcW")) if tcPr is not None else None
            if tcW is not None:
                _set_tcw(tcW, widths, c0, span)


def _keep(p_el, on=True):
    ppr = p_el.get_or_add_pPr()
    for x in ppr.findall(qn("w:keepNext")):
        ppr.remove(x)
    if on:
        ppr.insert(0, OxmlElement("w:keepNext"))


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


def _merge_area_line(doc):
    """'Khu vực thực hiện theo sự chỉ định của …:' + đoạn '“Phòng …”' -> 1 dòng (bản người dùng sửa trên Z)."""
    body = doc.element.body
    for p in list(body.iterchildren(qn("w:p"))):
        t = _ptext(p).strip()
        if not (t.startswith("Khu vực thực hiện theo sự chỉ định") or t.startswith("The performance location was indicated")):
            continue
        nxt = p.getnext()
        if nxt is None or nxt.tag != qn("w:p"):
            continue
        t2 = _ptext(nxt).strip()
        if not t2.startswith("“"):
            continue
        runs = [r for r in nxt.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
        if runs:
            sp = copy.deepcopy(runs[0])
            for x in sp.findall(qn("w:t")):
                sp.remove(x)
            tt = OxmlElement("w:t")
            tt.text = " "
            tt.set(qn("xml:space"), "preserve")
            sp.append(tt)
            p.append(sp)
            for r in runs:
                p.append(r)
        body.remove(nxt)


def _strip_room_quotes(doc):
    """Tên phòng trong ô bảng không ngoặc kép (bản người dùng sửa trên Z)."""
    for tbl in doc.element.body.iter(qn("w:tbl")):
        for tc in tbl.iter(qn("w:tc")):
            for p in tc.iter(qn("w:p")):
                for name in (ROOM_VN, ROOM_EN):
                    _replace_in_paragraph(p, f"“{name}”", name)


def _area_table(doc, is_en: bool):
    """Bảng khu vực: cấp sạch 'GMP Grade D' (template ghi 'ISO D' / 'D')."""
    for t in doc.tables:
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if not ("Cấp sạch" in hdr or "Class" in hdr):
            continue
        for row in t.rows[1:]:
            for c in row.cells:
                for p in c.paragraphs:
                    if p.text.strip() in (f"ISO {GRADE}", GRADE):
                        _set_paragraph_text(p._p, GMP_EN)
            for tr in [row._tr]:
                for h in tr.iter(qn("w:trHeight")):
                    if int(h.get(qn("w:val"), "0")) > 760:
                        h.set(qn("w:val"), "560")


def _fix_dust_table(doc, is_en: bool) -> bool:
    """Bảng bụi: bỏ cột d ≥ 1 µm (TT 14 chỉ có 0,5 và 5 µm); tiêu đề/dòng giới hạn theo bản người dùng sửa."""
    for t in doc.tables:
        last = t._tbl.findall(qn("w:tr"))[-1]
        txt = _ptext(last)
        if "ISO 14644-1:2015" not in txt or not re.search(r"(ISO Class|cấp ISO) \w", txt):
            continue
        lim_txt = (f"Maximum permissible particle count (particles/m3) for {GMP_EN} cleanroom (at rest) specified in "
                   f"{BASIS_EN}" if is_en else
                   f"Số lượng tiểu phân tối đa cho phép (tiểu phân/m3) trong khu vực sạch {GMP_VN} (trạng thái nghỉ) "
                   f"quy định tại {BASIS_VN}")
        first_p = next(iter(last.iter(qn("w:p"))))
        _set_paragraph_text(first_p, lim_txt)
        _drop_grid_column(t._tbl, 3)          # cột lưới: 0 khu vực, 1 vị trí, 2 d≥0,5, 3 d≥1, 4 d≥5
        if not is_en:
            for tr in t._tbl.findall(qn("w:tr"))[:2]:
                for p in tr.iter(qn("w:p")):
                    _replace_in_paragraph(p, "Nồng độ hạt bụi trong không khí, hạt/m3",
                                          "Số lượng tiểu phân có trong 1 m3 không khí, tiểu phân/m3")
                    _replace_in_paragraph(p, "Kích thước hạt bụi", "Kích thước tiểu phân")
        for tr in t._tbl.findall(qn("w:tr"))[:2]:
            for p in tr.iter(qn("w:p")):
                _replace_in_paragraph(p, "d ≥ 5 ", "d ≥ 5,0 " if not is_en else "d ≥ 5.0 ")
        return True
    return False


def _fix_texts(doc, is_en: bool):
    """Tiêu đề mục, câu đánh giá bụi, chữ 'tiểu phân' (VN) theo bản người dùng sửa trên Z."""
    if is_en:
        _replace_in_doc(doc, " (3 particle sizes)", "")
        _replace_in_doc(doc, "Determination result of microbial contamination level",
                        "Determination result of airborne microbial contamination")
        _replace_in_doc(doc, f"with the specified limit for ISO Class {GRADE} cleanroom according to ISO 14644-1:2015",
                        f"with the specified limit for {GMP_EN} cleanroom (at rest) according to {BASIS_EN}")
        _replace_in_doc(doc, "The cleanroom specifications provided by",
                        "The cleanroom specifications provided by the representative of")
        # EN gọi "location(s)", không "point(s)" (AGENTS mục 6) — template còn chữ cũ
        for old, new in (("number of monitoring point", "number of monitoring locations"),
                         ("measurement points determined", "measurement locations determined"),
                         ("Measurement point", "Measurement location"), ("Sampling point", "Sampling location"),
                         ("See measuring points", "See measurement locations"),
                         ("See sampling points", "See sampling locations")):
            _replace_in_doc(doc, old, new)
    else:
        _replace_in_doc(doc, "Kết quả đo nồng độ hạt bụi (3 kênh hạt)", "Kết quả đo nồng độ tiểu phân")
        _replace_in_doc(doc, "Kết quả xác định mức độ nhiễm bẩn vi sinh vật",
                        "Kết quả xác định độ nhiễm bẩn vi sinh vật trong không khí")
        _replace_in_doc(doc, "Kết quả đo nồng độ hạt bụi trong không khí tại",
                        "Kết quả đo nồng độ tiểu phân trong không khí tại")
        _replace_in_doc(doc, f"với mức quy định đối với khu vực cấp sạch ISO {GRADE} theo ISO 14644-1:2015",
                        f"với mức quy định đối với khu vực sạch {GMP_VN} (trạng thái nghỉ) nêu tại {BASIS_VN}")
        _replace_in_doc(doc, "Thông số kỹ thuật phòng sạch do Công ty", "Thông số kỹ thuật phòng sạch do đại diện Công ty")


def _fix_vi_sinh(doc, is_en: bool):
    """Dòng mức vi sinh theo TT 14 cấp D; thêm đoạn Đánh giá sau bảng Ghi chú (template không có — E44c)."""
    vs_tbl = next((t for t in doc.tables
                   if any(k in t.rows[-1].cells[0].text for k in ("WHO TRS 961", "microbial contamination", "vi sinh"))
                   and "Ghi chú" not in t.rows[0].cells[0].text and "Note" not in t.rows[0].cells[0].text), None)
    if vs_tbl is None:
        print("[Warning] build_131671: không thấy bảng vi sinh trong TBKQ")
        return
    lim_txt = (f"Microbial limit for {GMP_EN} cleanroom specified in {BASIS_EN}" if is_en else
               f"Giới hạn vi sinh vật đối với khu vực sạch {GMP_VN} nêu tại {BASIS_VN}")
    _set_paragraph_text(next(iter(vs_tbl.rows[-1]._tr.iter(qn("w:p")))), lim_txt)
    if not is_en:     # "Khu vực / lấy mẫu" (2 đoạn) -> 1 dòng như bản người dùng sửa trên Z
        ps = vs_tbl.rows[0].cells[0].paragraphs
        if len(ps) == 2 and ps[0].text.strip() == "Khu vực" and ps[1].text.strip() == "lấy mẫu":
            _set_paragraph_text(ps[0]._p, "Khu vực lấy mẫu")
            ps[1]._p.getparent().remove(ps[1]._p)
    _keep_table_together(vs_tbl)
    _superscript_header_marks(vs_tbl)

    note_el = vs_tbl._tbl.getnext()
    while note_el is not None and note_el.tag != qn("w:tbl"):
        note_el = note_el.getnext()
    paras = list(doc.paragraphs)
    head = next(p for p in paras if p.text.strip() in ("Đánh giá:", "Assessment:"))
    sent = paras[paras.index(head) + 1]
    new_head, new_sent = copy.deepcopy(head._p), copy.deepcopy(sent._p)
    note_el.addnext(new_head)
    new_head.addnext(new_sent)
    chain = [vs_tbl._tbl.findall(qn("w:tr"))[-1]]
    gap = vs_tbl._tbl.getnext()
    while gap is not None and gap is not note_el:
        chain.append(gap)
        gap = gap.getnext()
    for el in chain + [note_el, new_head]:
        for p_el in el.iter(qn("w:p")):
            _keep(p_el)
    kl = verdict(vs_ok(), is_en)
    if is_en:
        txt = (f"The analysis result of total aerobic microbial count in the air of “{ROOM_EN}” {kl} with the "
               f"specified limit for {GMP_EN} cleanroom according to {BASIS_EN}.")
    else:
        txt = (f"Kết quả xác định tổng số vi sinh vật hiếu khí trong không khí tại “{ROOM_VN}” {kl} với mức quy định "
               f"đối với khu vực sạch {GMP_VN} nêu tại {BASIS_VN}.")
    set_assessment_paragraph(Paragraph(new_sent, head._parent), txt, font_size=13, keep_next=False)


def _keep_table_together(tbl):
    rows = tbl.rows
    for i, row in enumerate(rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        for tc in row._tr.tc_lst:
            for p_el in tc.iter(qn("w:p")):
                _keep(p_el, i < len(rows) - 1)


def _superscript_header_marks(tbl):
    """'(2)' trong tiêu đề cột -> chỉ số trên (AGENTS mục 3)."""
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
        prev = el.getprevious()
        if prev is not None and prev.tag == qn("w:p") and _ptext(prev).strip():
            _keep(prev)
            rows = el.findall(qn("w:tr"))
            for tr in rows[:min(5, len(rows) - 1)]:
                for p_el in tr.iter(qn("w:p")):
                    _keep(p_el)


def _keep_result_blocks(doc):
    """Bảng kết quả ngắn + bảng Ghi chú + 'Đánh giá:' + câu kết luận cùng trang; câu kết luận chỉ keepLines."""
    body = doc.element.body
    for tbl in list(body.iterchildren(qn("w:tbl")))[1:]:
        nxt = tbl.getnext()
        gaps = []
        while nxt is not None and nxt.tag == qn("w:p") and not _ptext(nxt).strip():
            gaps.append(nxt)
            nxt = nxt.getnext()
        if nxt is None or nxt.tag != qn("w:tbl"):
            continue
        if not _ptext(nxt).strip().startswith(("Ghi chú", "Note")):
            continue
        rows = tbl.findall(qn("w:tr"))
        for tr in (rows if len(rows) <= 12 else rows[-2:]):
            for p_el in tr.iter(qn("w:p")):
                _keep(p_el)
        for g in gaps:
            _keep(g)
        head = nxt.getnext()
        if head is None or head.tag != qn("w:p"):
            continue
        if _ptext(head).strip() not in ("Đánh giá:", "Assessment:"):
            for p_el in nxt.iter(qn("w:p")):
                _keep(p_el, False)      # Ghi chú không có Đánh giá theo sau -> không keepNext (E32)
            continue
        for p_el in nxt.iter(qn("w:p")):
            _keep(p_el)
        _keep(head)
        sent = head.getnext()
        if sent is not None and sent.tag == qn("w:p"):
            _keep(sent, False)
            ppr = sent.get_or_add_pPr()
            if ppr.find(qn("w:keepLines")) is None:
                ppr.insert(0, OxmlElement("w:keepLines"))


def _highlight_placeholders(doc):
    """Tô vàng riêng phần '………' (thân văn bản, bảng, header/footer)."""
    roots = [doc.element.body]
    for s in doc.sections:
        roots += [s.header._element, s.footer._element, s.first_page_header._element]
    for root in roots:
        for r in list(root.iter(qn("w:r"))):
            t = r.find(qn("w:t"))
            if t is None or PLACEHOLDER not in (t.text or "") or len(r.findall(qn("w:t"))) != 1:
                continue
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            if rpr.find(qn("w:highlight")) is None:
                h = OxmlElement("w:highlight")
                h.set(qn("w:val"), "yellow")
                rpr.append(h)


def finish_tbkq(src: str, is_en: bool) -> str:
    doc = Document(src)
    _fix_cover_lists(doc)
    _merge_area_line(doc)
    _area_table(doc, is_en)
    if not _fix_dust_table(doc, is_en):
        print("[Warning] build_131671: không thấy bảng bụi")
    _fix_texts(doc, is_en)
    _fix_vi_sinh(doc, is_en)
    _strip_room_quotes(doc)
    _keep_headings_with_tables(doc)
    _keep_result_blocks(doc)
    style_note_tables(doc)      # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3)
    _highlight_placeholders(doc)
    lang = "en" if is_en else "viet"
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {lang} - phong sach - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)   # tên handler đặt ("TBKQ - vn - 1 - ...") khác quy ước AGENTS mục 1 (E41)
    return dst


# ---------------------------------------------------------------------------
# BCKQ
# ---------------------------------------------------------------------------
SO_BUI = f"{CODE}/Bụi"
SO_VI_SINH = f"{CODE}/Vi sinh"
VS_LINE = "- Kết quả lấy mẫu vi sinh vật trong không khí:"


def bckq_ref_lines(kqtn, vs_date):
    return [
        f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {SO_BUI} ngày {SAMPLE_DATE}.",
        VS_LINE,
    ]


def _cell(cell, text, size=13, align="center"):
    for p in list(cell.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = cell.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    for i, line in enumerate(text.split("\n")):
        p = p0 if i == 0 else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before, pf.space_after, pf.line_spacing = Pt(1), Pt(1), 1.0
        run = p.add_run(line)
        run.font.size = Pt(size)
        run.font.name = "Times New Roman"
        if PLACEHOLDER in line:
            run.font.highlight_color = 7
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _vs_ref_table(doc, anchor_p):
    """Bảng viện dẫn vi sinh (theo bản người dùng sửa trên Z): Vị trí | Ký hiệu mẫu | KQTN | KQTT."""
    kqtn, ngay = KQTN_SO or PLACEHOLDER, KQTN_NGAY or PLACEHOLDER
    t = doc.add_table(rows=3, cols=4)
    anchor_p.addnext(t._tbl)
    format_table_structure(t, row_height=440)
    for mar in list(t._tbl.iter(qn("w:tcMar"))):
        mar.getparent().remove(mar)
    widths = [24, 14, 30, 32]
    grid = t._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(round(9355 * w / 100)))
    for row in t.rows:
        for c, w in zip(row._tr.findall(qn("w:tc")), widths):
            tcPr = c.get_or_add_tcPr()
            old = tcPr.find(qn("w:tcW"))
            if old is not None:
                tcPr.remove(old)
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(round(50 * w)))
            tcW.set(qn("w:type"), "pct")
            tcPr.insert(0, tcW)
    add_tbl_header(t.rows[0])
    for j, h in enumerate(["Vị trí lấy mẫu", "Ký hiệu mẫu", "KQTN (CFU/ đĩa)", "KQTT (CFU/ m3)"]):
        _cell(t.rows[0].cells[j], h)
        set_cell_shading(t.rows[0].cells[j], "D9D9D9")
    _cell(t.rows[1].cells[0], ROOM_VN)
    _cell(t.rows[1].cells[1], "1 → 5")
    _cell(t.rows[2].cells[0], "Blank")
    _cell(t.rows[2].cells[1], str(BLANK_SAMPLE))
    _cell(t.rows[2].cells[3], "-")
    m = t.rows[1].cells[2].merge(t.rows[2].cells[2])
    _cell(m, f"Xem tại phiếu số {kqtn}\nngày {ngay}")
    _cell(t.rows[1].cells[3], f"Xem tại phiếu Tính toán vi sinh không khí số {SO_VI_SINH} ngày {ngay}")
    for tr in t._tbl.findall(qn("w:tr")):
        trPr = tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
    for tr in t._tbl.findall(qn("w:tr"))[:-1]:
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)
    return t


def _clone_para(proto_el, text):
    el = copy.deepcopy(proto_el)
    _set_paragraph_text(el, text)
    return el


def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    lod = round(1000 / (VS_FLOW_LPM * VS_MINUTES))
    ctx["tham_chieu"] = [
        *bckq_ref_lines(KQTN_SO or PLACEHOLDER, KQTN_NGAY or PLACEHOLDER),
        f"+ KQ được biểu thị < 1 CFU/đĩa ~ < {lod} CFU/m³ → không phát hiện vi sinh vật mục tiêu trong thể tích "
        f"không khí được lấy mẫu.",
        f"+ KQTN của mẫu Blank: {VI_SINH_BLANK or PLACEHOLDER} CFU/ mẫu → môi trường lấy mẫu đảm bảo kết quả.",
    ]
    dok, vok = dust_ok(), vs_ok()
    dat = (["Nồng độ hạt bụi"] if dok else []) + (["độ nhiễm bẩn vi sinh vật"] if vok else [])
    khong = (["Nồng độ hạt bụi"] if dok is False else []) + (["độ nhiễm bẩn vi sinh vật"] if vok is False else [])
    if vok is None:
        dat.append(f"độ nhiễm bẩn vi sinh vật: {PLACEHOLDER}")
    ctx["danh_gia"] = [{
        "khu_vuc": ROOM_VN,
        "dat": ", ".join(dat) if dat else "Không",
        "khong_dat": ", ".join(khong) if khong else "Không",
        "can_cu": f"Mức quy định đối với khu vực sạch {GMP_VN} nêu tại {BASIS_VN}",
    }]
    out = b.build(ctx=ctx, out_name=f"BCKQ - 1 - phong sach - {CODE}.docx")
    doc = Document(out)
    _merge_area_line(doc)
    _area_table(doc, False)
    _strip_room_quotes(doc)
    _replace_in_doc(doc, "- Nồng độ tiểu phân: 09 vị trí", "- Nồng độ hạt bụi: 09 vị trí;")
    _replace_in_doc(doc, "- Độ nhiễm bẩn vi sinh vật: 05 vị trí", "- Độ nhiễm bẩn vi sinh vật: 05 vị trí.")
    body = doc.element.body
    vs_p = next(p for p in body.iter(qn("w:p")) if _ptext(p).strip() == VS_LINE)
    _keep(vs_p)
    _vs_ref_table(doc, vs_p)
    # "Gửi mẫu thử nghiệm" trước "Báo cáo của GĐV" (theo bản người dùng sửa trên Z)
    gdv = next(p for p in body.iterchildren(qn("w:p")) if _ptext(p).strip().startswith("Báo cáo của GĐV"))
    proto = next(p for p in body.iterchildren(qn("w:p")) if _ptext(p).strip().startswith("Ngày kết thúc hồ sơ"))
    from Core.Tools.kqtn_late import tre_han_text   # ≤ 10 ngày làm việc từ ngày gửi mẫu: không trễ (AGENTS 13)
    tre = tre_han_text(KQTN_NGAY, NGAY_GUI_MAU, NGAY_YC_KQTN, PLACEHOLDER)
    head = copy.deepcopy(gdv)
    _set_paragraph_text(head, "Gửi mẫu thử nghiệm:")
    lines = [head] + [_clone_para(proto, s) for s in (
        f"Phiếu KQTN số {KQTN_SO or PLACEHOLDER} ngày {KQTN_NGAY or PLACEHOLDER}",
        f"Ngày gửi mẫu: {NGAY_GUI_MAU or PLACEHOLDER}",
        f"Ngày yêu cầu cấp KQTN: {NGAY_YC_KQTN or PLACEHOLDER}",
        f"Ngày nhận KQTN: {NGAY_NHAN_KQTN or PLACEHOLDER}",
        f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre}")]
    for el in lines:
        gdv.addprevious(el)
    for el in lines[:-1]:
        _keep(el)
    _keep_headings_with_tables(doc)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Hình ảnh", "Sơ đồ", "Báo cáo của GĐV", "Gửi mẫu thử nghiệm")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    return out


def _d(s):
    import datetime
    return datetime.datetime.strptime(s, "%d/%m/%Y").date()


# ---------------------------------------------------------------------------
# Excel phiếu tính toán (form phòng sạch chung, AGENTS mục 8C) — "Bụi 0,3" + "Vi sinh" (+ sheet ẩn bảng Buck)
# ---------------------------------------------------------------------------
BUCK_SHEET = "BuckConversion"


def build_excel() -> str:
    from openpyxl.styles import PatternFill
    yellow = PatternFill("solid", fgColor="FFFF00")
    book = CleanroomCalcBook()
    book.dust_aerotrak([{"room": ROOM_VN, "cap": GMP_VN, "basis": BASIS_VN,
                         "limits": {"Ch2": LIM05, "Ch6": LIM5},
                         "zero": RECS[ZERO_FILE], "points": [RECS[f] for f in POINT_FILES]}],
                       so=SO_BUI, date=SAMPLE_DATE, labels=LABELS)

    vs_date = KQTN_NGAY or PLACEHOLDER
    wv = book.extra_sheet("Vi sinh", so=SO_VI_SINH, date=vs_date,
                          widths=(6, 30, 9, 11, 10, 10, 14, 16, 13, 12, 11),
                          title="TÍNH TOÁN KẾT QUẢ VI SINH VẬT TRONG KHÔNG KHÍ")
    hdr = ("STT", "Vị trí lấy mẫu", "Ký hiệu mẫu", "Lưu lượng, L/phút", "Thời gian, phút", "Thể tích, L",
           "Tổng số VSV hiếu khí theo KQTN, CFU/đĩa", "Sau hiệu chỉnh theo bảng Conversion Buck B30120, CFU/đĩa",
           "Kết quả, CFU/m³", "Mức quy định, CFU/m³", "Đánh giá")
    for c, lbl in enumerate(hdr, 1):
        book.put(wv, 1, c, lbl, bold=True)
    wv.row_dimensions[1].height = 62
    tbl = f"{BUCK_SHEET}!$A$2:$B${BUCK_HOLES + 1}"
    for i in range(6):
        r = 2 + i
        s = i + 1
        blank = (s == BLANK_SAMPLE)
        book.put(wv, r, 1, s)
        book.put(wv, r, 2, "Mẫu blank" if blank else f"{ROOM_VN} - vị trí {s}", align=LEFT)
        book.put(wv, r, 3, s)
        cfu = VI_SINH_BLANK if blank else VI_SINH_CFU_DIA.get(s)
        c_ = book.put(wv, r, 7, cfu)
        if cfu is None:
            c_.fill = yellow
        if blank:
            for c in (4, 5, 6, 8, 9, 10, 11):
                book.put(wv, r, c, "-")
            continue
        book.put(wv, r, 4, VS_FLOW_LPM)
        book.put(wv, r, 5, VS_MINUTES)
        book.put(wv, r, 6, f"=D{r}*E{r}")
        book.put(wv, r, 8, f'=IF(G{r}="","",IF(OR(G{r}="< 1",G{r}=0),"< 1",VLOOKUP(G{r},{tbl},2,FALSE)))')
        book.put(wv, r, 9, f'=IF(H{r}="","",IF(H{r}="< 1","< "&ROUND(1000/F{r},0),ROUND(H{r}*1000/F{r},0)))',
                 fmt="0")
        book.put(wv, r, 10, f"≤ {VS_LIMIT}")
        book.put(wv, r, 11, f'=IF(I{r}="","",IF(ISNUMBER(I{r}),IF(I{r}<={VS_LIMIT},"Đạt","Không đạt"),"Đạt"))')
    # Bảng Conversion Buck B30120 (Feller, N lỗ) — sheet ẩn, công thức động như form người dùng
    wb_ = book.wb.create_sheet(BUCK_SHEET)
    wb_["A1"], wb_["B1"], wb_["C1"], wb_["D1"], wb_["E1"] = "r (CFU/đĩa)", "Pr (sau hiệu chỉnh)", "Σ 1/(N-i+1)", "N (số lỗ)", BUCK_HOLES
    for k in range(1, BUCK_HOLES + 1):
        r = k + 1
        wb_[f"A{r}"] = k
        wb_[f"C{r}"] = "=ROUND(1/$E$1,6)" if k == 1 else f"=ROUND(C{r - 1}+1/($E$1-A{r}+1),6)"
        wb_[f"B{r}"] = f"=ROUND(C{r}*$E$1,0)"
    wb_.sheet_state = "hidden"

    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - TinhToanPhongSach.xlsx")
    book.save(out)
    print(f"[build_131671] Excel -> {out}")
    return out


def patch_bckq_refs() -> str:
    """Chỉ dựng lại phiếu: sửa dòng viện dẫn bụi trong BCKQ đã có (giữ phần còn lại)."""
    path = os.path.join(INPUT_DIR, "ThucHien", f"BCKQ - 1 - phong sach - {CODE}.docx")
    new = bckq_ref_lines(KQTN_SO or PLACEHOLDER, KQTN_NGAY or PLACEHOLDER)
    a = replace_reference_lines(path, "- Kết quả đo nồng độ hạt bụi", new[:1])
    print(f"[build_131671] BCKQ viện dẫn bụi: {'đã sửa' if a else 'không thấy'}")
    return path


def split_edoc() -> list:
    """BBLM = trang 1–2, Ghi chép = trang 3–5 (trang 6 trắng, bỏ) -> TBKQ/Edoc (AGENTS mục 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(SCAN)
    outs = []
    for name, pages in ((f"{CODE6} - BBLM.pdf", (0, 1)), (f"{CODE6} - Ghi Chep.pdf", (2, 4))):
        dst = fitz.open()
        dst.insert_pdf(src, from_page=pages[0], to_page=pages[1])
        p = os.path.join(edoc, name)
        finish_edoc(dst, INPUT_DIR, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_131671] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--chi-excel", action="store_true",
                    help="Chỉ dựng lại phiếu tính toán (+ sửa dòng viện dẫn BCKQ); không đụng TBKQ")
    args = ap.parse_args()
    for s in range(1, 6):
        print(f"[build_131671] Vi sinh mẫu {s}: {VI_SINH_CFU_DIA.get(s)} CFU/đĩa -> {vs_value(s)} CFU/m³")
    print(f"[build_131671] Bụi {'phù hợp' if dust_ok() else 'KHÔNG phù hợp'} {GMP_VN}; vi sinh "
          f"{ {True: 'phù hợp', False: 'KHÔNG phù hợp', None: 'chờ KQTN'}[vs_ok()] }")
    if args.chi_excel:
        xlsx = build_excel()
        bckq = patch_bckq_refs()
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            ex = PDFExporter(build_dossier())
            ex.export_files([bckq])
            ex.export_excel_files([xlsx])
        return

    d = build_dossier()
    d.ensure_output_directories()
    split_edoc()
    h = CleanroomHandler(d)
    vn = finish_tbkq(h.build_tbkq(is_en=False), is_en=False)
    en = finish_tbkq(h.build_tbkq(is_en=True), is_en=True)
    bckq = build_bckq(d)
    xlsx = build_excel()
    docs = [bckq, vn, en]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(d)
        ex.export_files(docs)
        ex.export_excel_files([xlsx])
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=d.code).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
