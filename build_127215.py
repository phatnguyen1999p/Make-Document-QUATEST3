# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D127215 - Công ty Cổ phần Ánh Sáng Châu Á (phòng sạch "Phòng LAF", 10/11/2026) — 1 chứng thư (mã trần).

Nguồn (Input/127215, chép từ Z 10/11/2026 — chỉ đọc):
  - 20261110181415649.pdf (scan 6 trang): tr.1–2 BBLM (BM3030.02); tr.3–5 phiếu ghi chép (BM3030.01);
    tr.6 trang trắng. Không có PYC / báo giá (BBLM viện dẫn báo giá QONV26015453).
      BBLM: PYC Q326D127215 ngày 08/10/2026; địa chỉ "336 Nguyễn Trọng Tuyển, phường Tân Sơn Hòa, thành phố Hồ Chí Minh";
      địa điểm "Khu C - lô 3D-1, đường số 10, KCN Long Hậu, Cần Giuộc, Long An"; ngày 10/11/2026 14h30 – 15h15;
      "Phòng LAF": đo bụi 09 vị trí, vi sinh không khí 05 mẫu (1 → 5), chỉ tiêu tổng số vi sinh vật hiếu khí;
      quy trình QUATEST 3 3030:2025, N2.TC.PS.01, N2.TC.09; người thực hiện Nguyễn Tiến Phát.
      Ghi chép: trạng thái nghỉ; AeroTrak 9350, 1 phút (50 L), cao 1 m; file "Zero 963"; P. thay đồ (P4) 965, 964;
      P. Air Lock 1 (P3) 966, 967; P. Air lock nguyên liệu (P2) 968, 969; Buồng cân mẫu (P1) 973, 974; Pass box 971.
      Vi sinh: Buồng cân mẫu (P1) -> mẫu 1; Phòng Air-lock nguyên liệu (P2) -> 2; Phòng Air-lock 1 (P3) -> 4;
      Phòng thay đồ (P4) -> 5; Pass box -> 3; blank 6. Lưu lượng / thời gian lấy mẫu vi sinh để trống.
  - 6ChannelFixed.xlsx / .pdf: AeroTrak 9350-02 S/N 93501302003, zone "127215 i8", file 563 – 574 (hạt/m³, Σ).
  - Layout.png: sơ đồ người dùng vẽ (bụi 1 – 9, vi sinh VS1 – VS5) -> dùng nguyên, không chú thích.
  - Photos-1-001 (5)/Use/*.jpg: 2 ảnh người dùng chọn (143458, 150826) — chỉ dùng 2 ảnh này.
  - Old test/29887 - TBKQ(.pdf, - en.pdf): TBKQ Q326D029887 ngày 13/04/2026 (Công ty TNHH MTV Hóa dược Sài Gòn,
    cùng khu nhà máy Long Hậu): P3, P4 GMP D; P1, P2, Passbox GMP C; 2,04 / 1,00 m²; WHO TRS 961 – Phụ lục 6.
  - Hồ sơ cũ cùng khách: Z:\\...\\2025\\982N2.25DG - Ánh Sáng Châu Á (TBKQ 982, 29/10/2025).

Quyết định của người dùng (10/11/2026):
  - Cấp sạch A/B/C/D là theo GMP -> ghi "GMP Grade C" / "GMP Grade D" (bảng khu vực, dòng giới hạn, câu đánh giá,
    BCKQ, Excel; VN và EN) — hàm GMP().
  - Ảnh chèn: chỉ 2 ảnh trong Photos-1-001 (5)/Use.
  - "Giữ như hiện tại": đồng ý cả 6 điểm suy luận dưới đây (10/11/2026).

Suy luận của AI (người dùng đã đồng ý 10/11/2026):
  (1) File trên phiếu ghi "963 … 974" = file máy 563 … 574 (chữ 5 viết giống 9; 2 số cuối khớp, thời điểm khớp thứ tự đo).
      File 570 (14:58) và 572 (15:02) không có trên phiếu -> không dùng.
  (2) Vị trí bụi theo sơ đồ: 1, 2 = P4 (564, 565); 3, 4 = P3 (566, 567); 5, 6 = P2 (568, 569); 7 = Pass box (571);
      8, 9 = P1 (573, 574).
  (3) Cấp sạch, diện tích, mức đánh giá theo hồ sơ 2025 (982N2.25DG): P1, P2, Pass box cấp C; P3, P4 cấp D theo
      WHO TRS 961 (2011) – Phụ lục 6 (trạng thái nghỉ: C 352 000 / 2 900; D 3 520 000 / 29 000 hạt/m³ d ≥ 0,5 / 5 µm);
      vi sinh C ≤ 100, D ≤ 200 CFU/m³; P1 – P4 2,04 m², Pass box 1,00 m². Máy đặt zone "127215 i8" (ISO 8) — không dùng.
  (4) Vi sinh: VS1 = P4 (mẫu 5), VS2 = P3 (mẫu 4), VS3 = P2 (mẫu 2), VS4 = P1 (mẫu 1), VS5 = Pass box (mẫu 3);
      lấy mẫu 100 L/phút × 2 phút như 2025 (phiếu để trống).
  (5) Tên EN "Asia Shine Joint Stock Company" theo TBKQ 2025 (chưa có MST để tra tên quốc tế).
  (6) Địa chỉ theo BBLM (336 Nguyễn Trọng Tuyển; hồ sơ 2025 ghi 388); địa điểm chuẩn hóa 2 cấp "xã Cần Giuộc,
      tỉnh Tây Ninh" như TBKQ 2025.

Vi sinh CHƯA có KQTN: điền VI_SINH_CFU_DIA (theo ký hiệu mẫu 1..5), VI_SINH_BLANK, KQTN_SO, KQTN_NGAY rồi chạy lại.
Ngày ban hành TBKQ = ngày KQTN (AGENTS mục 2) -> chưa có thì để "………" tô vàng.

Chạy:  python build_127215.py          (DOCX + XLSX + Edoc)
       python build_127215.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
       python build_127215.py --chi-excel [--pdf]   (chỉ dựng lại phiếu tính toán + sửa dòng viện dẫn BCKQ)
Không đồng bộ Z trong code — chép bằng `python Core/Tools/ai_files.py chep 127215 --ghi` (AGENTS mục 0).
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
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, within, verdict, set_assessment_paragraph, style_note_tables
from Core.Tools.qa_check import check_single_docx
from Core.Tools.cleanroom_calc import CleanroomCalcBook, read_aerotrak, replace_reference_lines, LEFT
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D127215"
CODE6 = "127215"
REG_DATE = "08/10/2026"
SAMPLE_DATE = "10/11/2026"
START, END = "14:30", "15:15"
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN = os.path.join(INPUT_DIR, "20261110181415649.pdf")
RAW_XLSX = os.path.join(INPUT_DIR, "6ChannelFixed.xlsx")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photos-1-001 (5)", "Use")   # người dùng chọn 2 ảnh (10/11/2026)
LAYOUT = os.path.join(INPUT_DIR, "Layout.png")

# --- Vi sinh: chưa có KQTN. Điền khi có kết quả rồi chạy lại ---
KQTN_SO = ""             # vd "MTTN26/xxxx"
KQTN_NGAY = ""           # dd/mm/yyyy -> cũng là ngày ban hành TBKQ
# Số khuẩn lạc CFU/đĩa theo KQTN, theo KÝ HIỆU MẪU trên phiếu ghi chép (mẫu 1..5); None = chưa có.
VI_SINH_CFU_DIA = {1: None, 2: None, 3: None, 4: None, 5: None}
VI_SINH_BLANK = None      # mẫu 6
# Hoặc ghi thẳng kết quả CFU/m³ theo KQTN (ưu tiên hơn CFU/đĩa), vd {1: "< 5", 2: 10}
VI_SINH_KQ = {}
VS_FLOW_LPM, VS_MINUTES = 100, 2          # suy luận (4): như hồ sơ 2025 — phiếu để trống
BLANK_SAMPLE = 6

CERT_DATE = KQTN_NGAY or PLACEHOLDER

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần Ánh Sáng Châu Á",
    name_en="Asia Shine Joint Stock Company",
    address_vn="336 Nguyễn Trọng Tuyển, phường Tân Sơn Hòa, Thành phố Hồ Chí Minh",
    address_en="336 Nguyen Trong Tuyen Street, Tan Son Hoa Ward, Ho Chi Minh City",
)
SITE_VN = "Khu C, Lô 3D-1, đường số 10, Khu công nghiệp Long Hậu, xã Cần Giuộc, tỉnh Tây Ninh"
SITE_EN = "Area C, Lot 3D-1, Road No. 10, Long Hau Industrial Park, Can Giuoc Commune, Tay Ninh Province"
LOCATION_VN = f"Phòng sạch của {CLIENT.name_vn} - {SITE_VN}"
LOCATION_EN = f"Cleanroom of {CLIENT.name_en} - {SITE_EN}"

BASIS_VN, BASIS_EN = "WHO TRS 961 (2011) – Phụ lục 6", "WHO TRS 961 (2011) – Annex 6"
# Giới hạn trạng thái nghỉ (WHO TRS 961 (2011) – Annex 6), hạt/m³: d ≥ 0,5 µm; d ≥ 5 µm
GRADE_LIMITS = {"C": (352000, 2900), "D": (3520000, 29000)}
GRADE_VS = {"C": 100, "D": 200}           # CFU/m³


def GMP(g: str) -> str:
    """Cấp sạch A/B/C/D theo GMP -> 'GMP Grade C' (người dùng chốt 10/11/2026, VN và EN)."""
    return f"GMP Grade {g}"

ZERO_FILE = 563
# Phòng: (tên VN, tên EN, cấp, diện tích m², [(vị trí bụi, file)], vị trí vi sinh, ký hiệu mẫu vi sinh trên phiếu)
ROOMS = [
    ("Buồng cân mẫu (P1)", "Sample weighing booth (P1)", "C", "2,04", [(8, 573), (9, 574)], "VS4", 1),
    ("Phòng Air lock nguyên liệu (P2)", "Material air lock room (P2)", "C", "2,04", [(5, 568), (6, 569)], "VS3", 2),
    ("Pass box", "Pass box", "C", "1,00", [(7, 571)], "VS5", 3),
    ("Phòng Air lock 1 (P3)", "Air lock room 1 (P3)", "D", "2,04", [(3, 566), (4, 567)], "VS2", 4),
    ("Phòng thay đồ (P4)", "Changing room (P4)", "D", "2,04", [(1, 564), (2, 565)], "VS1", 5),
]
GROUP_ORDER = ("D", "C")                   # bảng kết quả: cấp D (vị trí 1–4) rồi cấp C (5–9)
TRANG_SO_DO = "2"                          # trang phụ lục chứa sơ đồ — kiểm tra lại trên PDF
EN_NAME = {r[0]: r[1] for r in ROOMS}

LABELS, RECS = read_aerotrak(RAW_XLSX)
_need = [ZERO_FILE] + [f for r in ROOMS for _, f in r[4]]
if [f for f in _need if f not in RECS]:
    raise SystemExit(f"[FATAL] 6ChannelFixed.xlsx thiếu file {[f for f in _need if f not in RECS]}")


def rooms_of(grade):
    """Phòng của một cấp, sắp theo số vị trí bụi nhỏ nhất."""
    return sorted((r for r in ROOMS if r[2] == grade), key=lambda r: min(p for p, _ in r[4]))


def dust_ok(grade):
    lim05, lim5 = GRADE_LIMITS[grade]
    return all(RECS[f]["Ch2"] <= lim05 and RECS[f]["Ch6"] <= lim5 for r in rooms_of(grade) for _, f in r[4])


def vs_value(sample):
    """Kết quả CFU/m³ của một ký hiệu mẫu; chưa có -> PLACEHOLDER."""
    vol_m3 = VS_FLOW_LPM * VS_MINUTES / 1000.0
    kq = VI_SINH_KQ.get(sample)
    if kq not in (None, ""):
        return kq
    cfu = VI_SINH_CFU_DIA.get(sample)
    if cfu is None:
        return PLACEHOLDER
    if cfu == 0:
        return f"< {round(1 / vol_m3)}"
    return round(cfu / vol_m3)


def vs_ok(grade):
    vals = [vs_value(r[6]) for r in rooms_of(grade)]
    if any(v == PLACEHOLDER for v in vals):
        return None
    return all(within(str(v).replace(".", ","), f"≤ {GRADE_VS[grade]}") for v in vals)


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
    photos = [{"path": _optimized(os.path.join(PHOTO_DIR, f)), "caption_vn": "", "caption_en": ""}
              for f in sorted(os.listdir(PHOTO_DIR)) if f.lower().endswith(".jpg")]
    layouts = [{"path": LAYOUT, "caption_vn": "", "caption_en": ""}]   # sơ đồ người dùng vẽ -> không chú thích

    locations = []
    for vn, en, grade, area, pts, vs, _ in ROOMS:
        nums = ", ".join(str(p) for p, _ in pts)
        n = len(pts)
        locations.append(MonitoringLocation(
            name_vn=vn, name_en=en, area_m2=area, clean_class=grade,
            state_vn="Trạng thái nghỉ", state_en="At rest",
            points_summary_vn=[f"Nồng độ hạt bụi: {n:02d} vị trí ({nums});",
                               f"Vi sinh vật trong không khí: 01 vị trí ({vs})."],
            points_summary_en=[f"Particle concentration: {n:02d} point{'s' if n > 1 else ''} ({nums});",
                               f"Microorganisms: 01 point ({vs})."],   # EN gọn cho cùng số dòng VN (AGENTS mục 3)
        ))

    bui, vi_sinh = [], []
    for grade in GROUP_ORDER:
        lim05, lim5 = GRADE_LIMITS[grade]
        diem = [{"khu_vuc": r[0], "vi_tri": p, "gia_tri": [RECS[f]["Ch2"], RECS[f]["Ch3"], RECS[f]["Ch6"]]}
                for r in rooms_of(grade) for p, f in r[4]]
        diem.sort(key=lambda x: x["vi_tri"])
        ok = dust_ok(grade)
        bui.append({"cap": grade, "gioi_han": [lim05, "", lim5], "trang_so_do": TRANG_SO_DO, "diem": diem,
                    "ket_luan_vn": verdict(ok, False), "ket_luan_en": verdict(ok, True)})
        vi_sinh.append({"cap": grade, "gioi_han": GRADE_VS[grade], "trang_so_do": TRANG_SO_DO,
                        "diem": [{"khu_vuc": r[0], "vi_tri": r[5], "gia_tri": vs_value(r[6])}
                                 for r in sorted(rooms_of(grade), key=lambda r: r[5])]})

    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False,
        dossier_type="cleanroom",
        scope_vn="Giám định chất lượng phòng sạch",
        scope_en="Inspection of cleanroom quality",
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01, N2.TC.09",
        procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01, N2.TC.09",
        standards_vn="ISO 14644-1:2015; WHO TRS 961 (2011) – Annex 6",
        standards_en="ISO 14644-1:2015; WHO TRS 961 (2011) – Annex 6",
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),     # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam
        equipments=[eq("aerotrak_9350"), eq("buck_b30120")],
        locations=locations,
        photos=photos, layouts=layouts,
        results_data={"bui_3kenh": bui, "vi_sinh": vi_sinh},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX
# ---------------------------------------------------------------------------
def _replace_in_paragraph(p_el, old: str, new: str) -> bool:
    """Thay chuỗi trong một đoạn (kể cả khi trải qua nhiều run): chữ mới nằm ở run chứa ký tự đầu."""
    ts = [t for t in p_el.iter(qn("w:t"))]
    full = "".join(t.text or "" for t in ts)
    i = full.find(old)
    if i < 0:
        return False
    j = i + len(old)
    pos = 0
    done = False
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
        while _replace_in_paragraph(p, old, new):
            n += 1
    return n


def _set_tcw(tcW, widths, c0, span):
    """tcW theo lưới mới: dxa = tổng twips; pct = phần của 5000 (bảng AutoFit to Window)."""
    w = sum(widths[c0:c0 + span])
    if tcW.get(qn("w:type")) == "pct":
        tcW.set(qn("w:w"), str(round(5000 * w / sum(widths))))
    elif tcW.get(qn("w:type")) == "dxa":
        tcW.set(qn("w:w"), str(w))


def _grid_cells(tr):
    """[(tc, cột lưới bắt đầu, số cột gộp)] của một hàng."""
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
    """Xóa cột lưới `gcol` (bỏ ô / giảm gridSpan); độ rộng cột bỏ chia cho các cột số liệu còn lại."""
    grid = tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    w_drop = int(cols[gcol].get(qn("w:w")))
    grid.remove(cols[gcol])
    cols = grid.findall(qn("w:gridCol"))
    # độ rộng cột bỏ: nửa cho cột "Khu vực" (tên phòng dài), nửa chia 2 cột số liệu (tiêu đề cột VN 1 dòng)
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
    # độ rộng ô theo lưới mới
    widths = [int(c.get(qn("w:w"))) for c in grid.findall(qn("w:gridCol"))]
    for tr in tbl.findall(qn("w:tr")):
        for tc, c0, span in _grid_cells(tr):
            tcPr = tc.find(qn("w:tcPr"))
            tcW = tcPr.find(qn("w:tcW")) if tcPr is not None else None
            if tcW is not None:
                _set_tcw(tcW, widths, c0, span)


def _shift_width(tbl, src: int, dst: int, frac: float):
    """Chuyển `frac` độ rộng cột lưới `src` sang cột `dst`, cập nhật tcW các ô."""
    grid = tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    d = int(int(cols[src].get(qn("w:w"))) * frac)
    cols[src].set(qn("w:w"), str(int(cols[src].get(qn("w:w"))) - d))
    cols[dst].set(qn("w:w"), str(int(cols[dst].get(qn("w:w"))) + d))
    widths = [int(c.get(qn("w:w"))) for c in cols]
    for tr in tbl.findall(qn("w:tr")):
        for tc, c0, span in _grid_cells(tr):
            tcPr = tc.find(qn("w:tcPr"))
            tcW = tcPr.find(qn("w:tcW")) if tcPr is not None else None
            if tcW is not None:
                _set_tcw(tcW, widths, c0, span)


def _fix_dust_tables(doc, is_en: bool) -> int:
    """Bảng bụi: bỏ cột d ≥ 1 µm (WHO TRS 961 chỉ có giới hạn 0,5 và 5 µm), dòng giới hạn + câu Đánh giá theo cấp GMP."""
    pat = (re.compile(r"The maximum permissible limit for ISO Class (\w) cleanroom specified in ISO 14644-1:2015") if is_en
           else re.compile(r"Giới hạn tối đa cho phép đối với phòng sạch cấp ISO (\w) quy định tại ISO 14644-1:2015"))
    n = 0
    for t in doc.tables:
        last = t._tbl.findall(qn("w:tr"))[-1]
        txt = "".join(x.text or "" for x in last.iter(qn("w:t")))
        m = pat.search(txt)
        if not m:
            continue
        g = m.group(1)
        new = (f"Maximum permissible limit for {GMP(g)} cleanroom (at rest) specified in {BASIS_EN}" if is_en
               else f"Giới hạn tối đa cho phép đối với khu vực sạch {GMP(g)} (trạng thái nghỉ) theo {BASIS_VN}")
        for p in last.iter(qn("w:p")):
            _replace_in_paragraph(p, m.group(0), new)
        _drop_grid_column(t._tbl, 3)          # cột lưới: 0 khu vực, 1 vị trí, 2 d≥0,5, 3 d≥1, 4 d≥5
        for src in (2, 3):                    # dòng giới hạn VN dài ('GMP Grade C', gộp cột 0+1) -> bớt 1 dòng
            _shift_width(t._tbl, src=src, dst=0, frac=0.15)
        n += 1
    for g in GRADE_LIMITS:
        if is_en:
            _replace_in_doc(doc, f"with the specified limit for ISO Class {g} cleanroom according to ISO 14644-1:2015",
                            f"with the specified limit for {GMP(g)} cleanroom (at rest) according to {BASIS_EN}")
            # dòng giới hạn bảng vi sinh (template: "… for Grade C specified in WHO TRS 961 …")
            _replace_in_doc(doc, f"for Grade {g} specified in", f"for {GMP(g)} cleanroom specified in")
        else:
            _replace_in_doc(doc, f"với mức quy định đối với khu vực cấp sạch ISO {g} theo ISO 14644-1:2015",
                            f"với mức quy định đối với khu vực sạch {GMP(g)} (trạng thái nghỉ) theo {BASIS_VN}")
            # dòng giới hạn bảng vi sinh (template: "… đối với cấp sạch C nêu tại WHO TRS 961 …")
            _replace_in_doc(doc, f"đối với cấp sạch {g} nêu tại", f"đối với khu vực sạch {GMP(g)} nêu tại")
    return n


def _fix_vi_sinh(doc, is_en: bool):
    """Mỗi bảng vi sinh: thêm đoạn Đánh giá sau bảng Ghi chú (template không có — E44c)."""
    key = "Annex 6" if is_en else "Phụ lục 6"
    vs_tbls = [t for t in doc.tables
               if key in t.rows[-1].cells[0].text and ("microbial" in t.rows[-1].cells[0].text.lower()
                                                        or "vi sinh" in t.rows[-1].cells[0].text)]
    if len(vs_tbls) != len(GROUP_ORDER):
        print(f"[Warning] build_127215: thấy {len(vs_tbls)} bảng vi sinh (cần {len(GROUP_ORDER)})")
    paras = list(doc.paragraphs)
    head = next(p for p in paras if p.text.strip() in ("Đánh giá:", "Assessment:"))
    sent = paras[paras.index(head) + 1]
    for t, grade in zip(vs_tbls, GROUP_ORDER):
        _shift_width(t._tbl, src=2, dst=0, frac=0.5)     # cột "Khu vực lấy mẫu" hẹp -> tên phòng 2–3 dòng
        _keep_table_together(t)
        _superscript_header_marks(t)
        note_el = t._tbl.getnext()
        while note_el is not None and note_el.tag != qn("w:tbl"):
            note_el = note_el.getnext()
        new_head, new_sent = copy.deepcopy(head._p), copy.deepcopy(sent._p)
        note_el.addnext(new_head)
        new_head.addnext(new_sent)
        # bảng vi sinh (cả dòng cuối) + bảng Ghi chú + "Đánh giá:" đi liền câu kết luận (AGENTS mục 3) —
        # VN từng rớt bảng Ghi chú + Đánh giá cấp C sang trang sau, tách khỏi bảng
        chain = [t._tbl.findall(qn("w:tr"))[-1]]
        gap = t._tbl.getnext()
        while gap is not None and gap is not note_el:     # đoạn trống ngăn bảng kết quả – bảng Ghi chú
            chain.append(gap)
            gap = gap.getnext()
        for el in chain + [note_el, new_head]:
            for p_el in el.iter(qn("w:p")):
                ppr = p_el.get_or_add_pPr()
                if ppr.find(qn("w:keepNext")) is None:
                    ppr.insert(0, OxmlElement("w:keepNext"))
        rooms = "; ".join(f"“{(r[1] if is_en else r[0])}”" for r in rooms_of(grade))
        kl = verdict(vs_ok(grade), is_en)
        if is_en:
            txt = (f"The analysis result of total aerobic plate count in the air of {rooms} {kl} with the specified "
                   f"limit for {GMP(grade)} cleanroom according to {BASIS_EN}.")
        else:
            txt = (f"Kết quả phân tích tổng số vi sinh vật hiếu khí trong không khí tại {rooms} {kl} với mức quy định "
                   f"đối với khu vực sạch {GMP(grade)} theo {BASIS_VN}.")
        set_assessment_paragraph(Paragraph(new_sent, head._parent), txt, font_size=13, keep_next=False)


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


def _keep_table_together(tbl):
    """Bảng nhỏ: mọi dòng cantSplit + keepNext trừ dòng cuối."""
    rows = tbl.rows
    for i, row in enumerate(rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        for tc in row._tr.tc_lst:
            for p_el in tc.iter(qn("w:p")):
                pPr = p_el.get_or_add_pPr()
                for el in pPr.findall(qn("w:keepNext")):
                    pPr.remove(el)
                if i < len(rows) - 1:
                    pPr.insert(0, OxmlElement("w:keepNext"))


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
    """Tiêu đề mục ngay trước bảng kết quả -> keepNext; tiêu đề cột + vài dòng đầu đi liền (AGENTS mục 3)."""
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
            parts = t.text.split(PLACEHOLDER)
            pieces = []
            for i, part in enumerate(parts):
                if part:
                    pieces.append((part, False))
                if i < len(parts) - 1:
                    pieces.append((PLACEHOLDER, True))
            parent, idx = r.getparent(), r.getparent().index(r)
            extras = [c for c in r if c.tag not in (qn("w:rPr"), qn("w:t"))]
            if extras:
                lead = copy.deepcopy(r)
                for c in list(lead):
                    if c.tag == qn("w:t"):
                        lead.remove(c)
                parent.insert(idx + 1, lead)
                idx += 1
            for k, (txt, hl) in enumerate(pieces):
                nr = copy.deepcopy(r)
                for c in list(nr):
                    if c.tag not in (qn("w:rPr"), qn("w:t")):
                        nr.remove(c)
                nt = nr.find(qn("w:t"))
                nt.text = txt
                nt.set(qn("xml:space"), "preserve")
                if hl:
                    rpr = nr.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr")
                        nr.insert(0, rpr)
                    h = OxmlElement("w:highlight")
                    h.set(qn("w:val"), "yellow")
                    rpr.append(h)
                parent.insert(idx + 1 + k, nr)
            parent.remove(r)


def _fix_misc(doc, is_en: bool):
    """Cấp sạch GMP ở bảng khu vực (template ghi 'ISO <cấp>'); tiêu đề mục bụi bỏ '(3 kênh hạt)' (còn 2 cỡ hạt);
    ghi chú '< 10 CFU/m³' theo thể tích lấy mẫu thực tế."""
    for t in doc.tables:
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if ("Cấp sạch" in hdr or "Class" in hdr) and len(t.columns) >= 3:
            # cột Diện tích / Cấp sạch hẹp ('GMP Grade C' 3 dòng ở EN) -> lấy bớt từ cột chỉ tiêu; VN = EN cùng ngắt trang
            if len(t._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))) == 4:
                _shift_width(t._tbl, src=3, dst=2, frac=0.07)
            # template ép hàng 1339 twips -> hàng 2 dòng 560 twips atLeast (AGENTS mục 9, 15 [Row Too Tall])
            for tr in t._tbl.findall(qn("w:tr"))[1:]:
                for h in tr.iter(qn("w:trHeight")):
                    if int(h.get(qn("w:val"), "0")) > 760:
                        h.set(qn("w:val"), "560")
            for row in t.rows[1:]:
                for c in row.cells:
                    for p in c.paragraphs:
                        for g in GRADE_LIMITS:
                            if p.text.strip() == f"ISO {g}":
                                _replace_in_paragraph(p._p, f"ISO {g}", GMP(g))
    _replace_in_doc(doc, " (3 particle sizes)" if is_en else " (3 kênh hạt)", "")
    lod = round(1000 / (VS_FLOW_LPM * VS_MINUTES))
    _replace_in_doc(doc, "“< 10 CFU/m³”", f"“< {lod} CFU/m³”")


def _keep_result_blocks(doc):
    """Bảng kết quả ngắn (≤ 10 hàng) + đoạn trống + bảng Ghi chú + 'Đánh giá:' + câu kết luận cùng trang;
    câu kết luận chỉ keepLines, không keepNext (AGENTS mục 3, 15)."""
    def kn(el, on=True):
        for p_el in ([el] if el.tag == qn("w:p") else list(el.iter(qn("w:p")))):
            ppr = p_el.get_or_add_pPr()
            for x in ppr.findall(qn("w:keepNext")):
                ppr.remove(x)
            if on:
                ppr.insert(0, OxmlElement("w:keepNext"))

    body = doc.element.body
    tbls = list(body.iterchildren(qn("w:tbl")))[1:]          # bỏ Table 0 (bìa)
    for tbl in tbls:
        nxt = tbl.getnext()
        gaps = []
        while nxt is not None and nxt.tag == qn("w:p") and not "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip():
            gaps.append(nxt)
            nxt = nxt.getnext()
        if nxt is None or nxt.tag != qn("w:tbl"):
            continue
        first = "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip()
        if not first.startswith(("Ghi chú", "Note")):
            continue
        rows = tbl.findall(qn("w:tr"))
        for tr in (rows if len(rows) <= 12 else rows[-2:]):
            kn(tr)
        for g in gaps:
            kn(g)
        head = nxt.getnext()
        if head is None or head.tag != qn("w:p"):
            continue
        if "".join(t.text or "" for t in head.iter(qn("w:t"))).strip() not in ("Đánh giá:", "Assessment:"):
            kn(nxt, False)      # Ghi chú không có Đánh giá theo sau -> không keepNext (E32)
            continue
        kn(nxt)
        kn(head)
        sent = head.getnext()
        if sent is not None and sent.tag == qn("w:p"):
            kn(sent, False)
            ppr = sent.get_or_add_pPr()
            if ppr.find(qn("w:keepLines")) is None:
                ppr.insert(0, OxmlElement("w:keepLines"))


def finish_tbkq(src: str, is_en: bool) -> str:
    doc = Document(src)
    _fix_cover_lists(doc)
    _fix_misc(doc, is_en)
    n = _fix_dust_tables(doc, is_en)
    if n != len(GROUP_ORDER):
        print(f"[Warning] build_127215: sửa {n} bảng bụi (cần {len(GROUP_ORDER)})")
    _fix_vi_sinh(doc, is_en)
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


def bckq_ref_lines(kqtn, vs_date):
    mapping = "; ".join(f"mẫu {r[6]} – {r[5]} ({r[0]})" for r in sorted(ROOMS, key=lambda r: r[6]))
    return [
        f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {SO_BUI} ngày {SAMPLE_DATE}.",
        f"- Kết quả lấy mẫu vi sinh vật trong không khí (05 mẫu và 01 mẫu blank): xem tại phiếu Kết quả thử nghiệm "
        f"số {kqtn} ngày {vs_date} và phiếu Tính toán kết quả vi sinh vật trong không khí số {SO_VI_SINH} "
        f"ngày {vs_date}.",
        f"- Ký hiệu mẫu vi sinh (phiếu ghi chép) – vị trí trên sơ đồ: {mapping}; mẫu blank: {BLANK_SAMPLE}.",
    ]


def _th(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    vs_date = KQTN_NGAY or PLACEHOLDER
    ctx["tham_chieu"] = [
        *bckq_ref_lines(KQTN_SO or PLACEHOLDER, vs_date),
        f"- KQTN của mẫu Blank: {VI_SINH_BLANK if VI_SINH_BLANK is not None else PLACEHOLDER}.",
    ]
    vs_name = "Vi sinh vật trong không khí (tổng số vi sinh vật hiếu khí)"
    ctx["danh_gia"] = []
    for grade in GROUP_ORDER:
        dok, vok = dust_ok(grade), vs_ok(grade)
        dat = (["Nồng độ hạt bụi"] if dok else []) + ([vs_name] if vok else [])
        khong = (["Nồng độ hạt bụi"] if dok is False else []) + ([vs_name] if vok is False else [])
        if vok is None:
            dat.append(f"{vs_name}: {PLACEHOLDER}")
        ctx["danh_gia"].append({
            "khu_vuc": "; ".join(f"“{r[0]}”" for r in rooms_of(grade)),
            "dat": ";\n".join(dat),
            "khong_dat": ";\n".join(khong) if khong else ("Không" if vok is not None else PLACEHOLDER),
            "can_cu": (f"{BASIS_VN} (khu vực sạch {GMP(grade)}, trạng thái nghỉ: d ≥ 0,5 µm ≤ "
                       f"{_th(GRADE_LIMITS[grade][0])} hạt/m³; d ≥ 5 µm ≤ {_th(GRADE_LIMITS[grade][1])} hạt/m³; "
                       f"vi sinh vật ≤ {GRADE_VS[grade]} CFU/m³)"),
        })
    out = b.build(ctx=ctx, out_name=f"BCKQ - 1 - phong sach - {CODE}.docx")
    doc = Document(out)
    _fix_misc(doc, False)
    _keep_headings_with_tables(doc)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Hình ảnh", "Sơ đồ", "Báo cáo của GĐV")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    return out


# ---------------------------------------------------------------------------
# Excel phiếu tính toán (form phòng sạch chung, AGENTS mục 8C) — "Bụi 0,3" + "Vi sinh"
# ---------------------------------------------------------------------------
def build_excel() -> str:
    from openpyxl.styles import PatternFill
    yellow = PatternFill("solid", fgColor="FFFF00")
    book = CleanroomCalcBook()
    groups, first = [], True
    for grade in GROUP_ORDER:
        lim05, lim5 = GRADE_LIMITS[grade]
        for r in rooms_of(grade):
            pts = sorted(r[4])
            groups.append({"room": r[0], "cap": GMP(grade), "basis": BASIS_VN,
                           "limits": {"Ch2": lim05, "Ch6": lim5},
                           "zero": RECS[ZERO_FILE] if first else None,
                           "points": [RECS[f] for _, f in pts], "labels": [p for p, _ in pts]})
            first = False
    book.dust_aerotrak(groups, so=SO_BUI, date=SAMPLE_DATE, labels=LABELS)

    vs_date = KQTN_NGAY or PLACEHOLDER
    wv = book.extra_sheet("Vi sinh", so=SO_VI_SINH, date=vs_date, widths=(6, 34, 9, 11, 10, 10, 15, 13, 12, 11),
                          title="TÍNH TOÁN KẾT QUẢ VI SINH VẬT TRONG KHÔNG KHÍ")
    hdr = ("STT", "Vị trí lấy mẫu", "Ký hiệu mẫu", "Lưu lượng, L/phút", "Thời gian, phút", "Thể tích, L",
           "Số khuẩn lạc, CFU/đĩa (theo KQTN)", "Kết quả, CFU/m³", "Mức quy định, CFU/m³", "Đánh giá")
    for c, lbl in enumerate(hdr, 1):
        book.put(wv, 1, c, lbl, bold=True)
    wv.row_dimensions[1].height = 45
    order = sorted(ROOMS, key=lambda r: r[5])               # VS1..VS5
    for i, room in enumerate(order + [None]):
        r = 2 + i
        book.put(wv, r, 1, i + 1)
        if room is None:
            book.put(wv, r, 2, "Mẫu blank (TS)", align=LEFT)
            book.put(wv, r, 3, BLANK_SAMPLE)
            for c in (4, 5, 6, 8, 9, 10):
                book.put(wv, r, c, "-")
            c_ = book.put(wv, r, 7, VI_SINH_BLANK)
            if VI_SINH_BLANK is None:
                c_.fill = yellow
            continue
        vn, _, grade, _, _, vs, sample = room
        book.put(wv, r, 2, f"{vn} - {vs}", align=LEFT)
        book.put(wv, r, 3, sample)
        book.put(wv, r, 4, VS_FLOW_LPM)
        book.put(wv, r, 5, VS_MINUTES)
        book.put(wv, r, 6, f"=D{r}*E{r}")
        cfu = VI_SINH_CFU_DIA.get(sample)
        c_ = book.put(wv, r, 7, cfu)
        if cfu is None:
            c_.fill = yellow
        kq = VI_SINH_KQ.get(sample)
        if kq not in (None, ""):
            book.put(wv, r, 8, kq)
        else:
            book.put(wv, r, 8, f'=IF(G{r}="","",IF(G{r}=0,"< "&ROUND(1000/F{r},0),ROUND(G{r}*1000/F{r},0)))', fmt="0")
        lim = GRADE_VS[grade]
        book.put(wv, r, 9, f"≤ {lim}")
        book.put(wv, r, 10, f'=IF(H{r}="","",IF(ISNUMBER(H{r}),IF(H{r}<={lim},"Đạt","Không đạt"),"Đạt"))')

    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - TinhToanPhongSach.xlsx")
    book.save(out)
    for n in book.notes:
        print(f"[build_127215] {n}")
    print(f"[build_127215] Excel -> {out}")
    return out


def patch_bckq_refs() -> str:
    path = os.path.join(INPUT_DIR, "ThucHien", f"BCKQ - 1 - phong sach - {CODE}.docx")
    new = bckq_ref_lines(KQTN_SO or PLACEHOLDER, KQTN_NGAY or PLACEHOLDER)
    a = replace_reference_lines(path, "- Kết quả đo nồng độ hạt bụi", new[:1])
    b = replace_reference_lines(path, "- Kết quả lấy mẫu vi sinh vật", new[1:2])
    print(f"[build_127215] BCKQ viện dẫn: bụi {'đã sửa' if a else 'không thấy'}, vi sinh {'đã sửa' if b else 'không thấy'}")
    return path


def split_edoc() -> list:
    """BBLM = trang 1–2, Ghi chép = trang 3–5; trang 6 trắng bỏ; không có trang 'Khác' (AGENTS mục 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(SCAN)
    outs = []
    for name, (a, b) in ((f"{CODE6} - BBLM.pdf", (0, 1)), (f"{CODE6} - Ghi Chep.pdf", (2, 4))):
        dst = fitz.open()
        dst.insert_pdf(src, from_page=a, to_page=b)
        p = os.path.join(edoc, name)
        finish_edoc(dst, INPUT_DIR, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_127215] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--chi-excel", action="store_true",
                    help="Chỉ dựng lại phiếu tính toán (+ sửa dòng viện dẫn BCKQ); không đụng TBKQ")
    args = ap.parse_args()
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
