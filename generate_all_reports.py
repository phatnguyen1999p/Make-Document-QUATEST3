# -*- coding: utf-8 -*-
"""
Full Document Generator for Cleanroom Dossier Q326D132115.
Generates:
1. TBKQ - viet - 1 - phong sach - Q326D132115.docx
2. TBKQ - en - 1 - phong sach - Q326D132115.docx
3. BCKQ - 1 - phong sach - Q326D132115.docx
"""

import os
import re
import copy
import math
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml import parse_xml
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from Core.Styling import (
    W_NS,
    set_cell_content,
    set_paragraph_text,
    set_assessment_paragraph,
    set_table_autofit_window,
    format_table_structure,
    strip_all_highlights,
    add_document_ending_line,
    update_headers,
    populate_cleanroom_table0,
    PHOTO_WIDTH_2COL,
    strip_image_crops,
)
from Core.Styling.tables import (
    add_cant_split,
    add_tbl_header,
    A_NS,
    set_cell_margins,
    set_table_col_widths
)
from Core.Models.dossier import Dossier, ClientInfo, MonitoringLocation, PersonnelInfo, Equipment
from Core.Builders.bckq_builder import populate_bckq_table0
from build_cleanroom_132115 import (
    ROOMS,
    PRESSURE_PAIRS,
    ISO8_LIMITS,
    eval_dp,
    eval_ach,
    temp_rh_points,
    fmt_dec,
    PARTICLE_DATA,
    fmt_int_vn,
    fmt_int_en,
    assess_dust,
    assess_temp_rh,
    assess_dp,
    assess_ach,
    room_summary,
    points_detail,
    ISO_CLASS,
    SHEET_NO,
    TEMP_RH_SPECS,
    temp_rh_room_avg,
    trh_room_fail,
    trh_spec_groups,
    dp_value,
    fmt_flow,
    ach_value,
    spec_txt,
)
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

# 0. Can cu danh gia: chi liet ke cac chi tieu CO muc quy dinh trong file Excel so lieu
def _customer_spec_items():
    vn, en = [], []
    if any(sp[0] is not None or sp[1] is not None for pts in TEMP_RH_SPECS.values() for sp in pts):
        vn.append("nhiệt độ - độ ẩm"); en.append("temperature - humidity")
    if any(p["spec"] is not None for p in PRESSURE_PAIRS):
        vn.append("chênh áp"); en.append("differential pressure")
    if any(r.get("ach_spec") is not None for r in ROOMS):
        vn.append("lưu lượng khí cấp và số lần trao đổi khí"); en.append("supply airflow rate and air change rate")
    return vn, en


_sv, _se = _customer_spec_items()
STANDARDS_VN = "- ISO 14644-1:2015 Cleanrooms and associated controlled environments — Part 1: Classification of air cleanliness by particle concentration"
STANDARDS_EN = "- ISO 14644-1:2015 Cleanrooms and associated controlled environments - Part 1: Classification of air cleanliness by particle concentration"
if _sv:
    STANDARDS_VN += ";\n- Mức quy định đối với " + ", ".join(_sv) + " được cung cấp bởi Công ty TNHH Lâm Quang Đại."
    STANDARDS_EN += ";\n- Technical specifications for " + ", ".join(_se) + " provided by Lam Quang Dai Company Limited."
else:
    STANDARDS_VN += "."
    STANDARDS_EN += "."

# 1. Dossier Metadata
d = Dossier(
    code="Q326D132115",
    dossier_type="phong sach",
    sub_no="1",
    is_multi_cert=False,
    cert_date="23/10/2026",
    sample_date="23/10/2026",
    reg_date="19/10/2026",
    sample_time_vn="Từ ngày 23/10/2026",
    sample_time_en="From 23/10/2026",
    client=ClientInfo(
        name_vn="Công ty TNHH Lâm Quang Đại",
        name_en="Lam Quang Dai Company Limited",
        address_vn="89 Lê Thị Riêng, Phường Thới An, Thành phố Hồ Chí Minh",
        address_en="89 Le Thi Rieng, Thoi An Ward, Ho Chi Minh City"
    ),
    location_desc_vn="Công ty TNHH Thực phẩm Farina - Chi nhánh Hồ Chí Minh (Dự án Farina - Khu R0A, KCN Việt Nam - Singapore III, phường Tân Uyên, TP. Tân Uyên, Bình Dương)",
    location_desc_en="Farina Food Company Limited - Ho Chi Minh Branch (Farina Project - Zone R0A, Vietnam - Singapore III Industrial Park, Tan Uyen Ward, Tan Uyen City, Binh Duong Province)",
    scope_vn="Quan trắc chất lượng phòng sạch",
    scope_en="Cleanroom environmental monitoring",
    procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01",
    procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01",
    equipments=[
        eq("aerotrak_9350"),
        eq("testo_625"),
        eq("testo_420")
    ],
    standards_vn=STANDARDS_VN,
    standards_en=STANDARDS_EN,
    personnel=PersonnelInfo(
        executor="Nguyễn Văn Phát",
        sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
        sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
        sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
        sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen",
        reporter="Phát + Minh + Nam"
    ),
    input_dir=os.path.join("Input", "132115"),
    thuchien_dir=os.path.join("Input", "132115", "ThucHien"),
    tbkq_dir=os.path.join("Input", "132115", "TBKQ"),
    edoc_dir=os.path.join("Input", "132115", "TBKQ", "Edoc"),
    pdf_dir=os.path.join("Input", "132115", "pdf")
)

def optimized_image(path, max_width_px, as_jpeg=True, landscape=False):
    """Down-scaled COPY of a field photo / diagram (same aspect ratio, NO crop) in %TEMP%. The original 12 MP photos
    made the DOCX ~14 MB and the Word PDF export very slow; 7.5 cm at ~400 dpi needs only ~1200 px."""
    try:
        from PIL import Image, ImageOps
        import tempfile, hashlib
        cache = os.path.join(tempfile.gettempdir(), "lamkq_img_cache")
        os.makedirs(cache, exist_ok=True)
        st = os.stat(path)
        key = hashlib.md5(f"{os.path.abspath(path)}|{st.st_mtime_ns}|{max_width_px}|{landscape}".encode()).hexdigest()[:12]
        out = os.path.join(cache, f"{os.path.splitext(os.path.basename(path))[0]}_{key}.{'jpg' if as_jpeg else 'png'}")
        if os.path.exists(out):
            return out
        with Image.open(path) as im:
            im = ImageOps.exif_transpose(im)
            if landscape and im.height > im.width:
                # Anh doc -> xoay 90 do nguoc chieu kim dong ho de chen theo phuong ngang (khong crop, giu ti le)
                im = im.rotate(90, expand=True)
            if im.width > max_width_px:
                im = im.resize((max_width_px, round(im.height * max_width_px / im.width)), Image.LANCZOS)
            if as_jpeg:
                im.convert("RGB").save(out, "JPEG", quality=88, optimize=True)
            else:
                im.save(out, "PNG", optimize=True)
        return out
    except ImportError:
        try:  # khong co Pillow -> dung PyMuPDF (da co san cho QA)
            import fitz, tempfile, hashlib
            cache = os.path.join(tempfile.gettempdir(), "lamkq_img_cache")
            os.makedirs(cache, exist_ok=True)
            st = os.stat(path)
            key = hashlib.md5(f"{os.path.abspath(path)}|{st.st_mtime_ns}|{max_width_px}".encode()).hexdigest()[:12]
            out = os.path.join(cache, f"{os.path.splitext(os.path.basename(path))[0]}_{key}.{'jpg' if as_jpeg else 'png'}")
            if not os.path.exists(out):
                pix = fitz.Pixmap(path)
                if pix.width > max_width_px:
                    pix = fitz.Pixmap(pix, max_width_px, round(pix.height * max_width_px / pix.width), None)
                if pix.alpha or pix.colorspace.n > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                pix.save(out, jpg_quality=88) if as_jpeg else pix.save(out)
            return out
        except Exception as e:
            print(f"[Warning] Khong thu nho duoc anh {path}: {type(e).__name__}: {e} (dung anh goc)")
            return path
    except Exception as e:
        print(f"[Warning] Khong thu nho duoc anh {path}: {type(e).__name__}: {e} (dung anh goc)")
        return path


# 2. Photos & Layouts
photo_paths = [os.path.join("Input", "132115", "Photo", f) for f in (
    "20261023_091512.jpg",
    "20261023_093846.jpg",
    "20261023_094449.jpg",
    "20261023_144716.jpg",
)]
# Anh hien truong chen theo phuong NGANG (landscape) theo yeu cau nguoi dung 26/10/2026
d.photos = [
    {"path": optimized_image(p, 1200, landscape=True), "caption_vn": f"Hình {i+1}: Quá trình đo đạc tại hiện trường", "caption_en": f"Figure {i+1}: Field inspection process"}
    for i, p in enumerate(photo_paths) if os.path.exists(p)
]

# So do vi tri do theo tung chi tieu, nguoi dung ve san trong Input\132115\Layout (30/10/2026)
LAYOUT_DIR = os.path.join("Input", "132115", "Layout")
d.layouts = [
    {"path": optimized_image(os.path.abspath(os.path.join(LAYOUT_DIR, f)), 2000, as_jpeg=False),
     "caption_vn": f"Sơ đồ {i + 1}: {vn}", "caption_en": f"Diagram {i + 1}: {en}"}
    for i, (f, vn, en) in enumerate((
        ("Bui_NhietDo.png", "Vị trí đo nồng độ hạt bụi trong không khí và nhiệt độ, độ ẩm",
         "Airborne particle concentration and temperature – humidity measurement points"),
        ("ChenhAp.png", "Vị trí đo chênh lệch áp suất không khí",
         "Air pressure differential measurement points"),
        ("H13 Supply Filter.png", "Vị trí đo lưu lượng khí cấp qua lọc HEPA",
         "HEPA-filtered supply airflow measurement locations"),
    ))
]
# Cau cong nhan bat buoc o moi bang Ghi chu duoi bang ket qua (AGENTS muc 3, qa_check kiem tra)
ACCREDIT_VN = "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004."
ACCREDIT_EN = "The performance field is accredited to ISO 17020 - VIAS 004."
LAYOUT_WIDTH = Cm(13)               # 13 cm: TBKQ VN va EN cung so trang (14,5 cm thi VN 12, EN 11)


def fill_layout_table(t4, is_en=False):
    """Bang so do: moi so do 1 dong (khong tach dong qua trang), anh chi dat width (khong crop).
    Khong ghi chu thich duoi so do (nguoi dung chot 30/10/2026: so do da co chu giai trong anh)."""
    while len(t4.rows) < len(d.layouts):
        t4._tbl.append(copy.deepcopy(t4.rows[0]._tr))
    for row, lay in zip(t4.rows, d.layouts):
        cell = row.cells[0]
        for p in list(cell.paragraphs):
            p._p.getparent().remove(p._p)
        p_img = cell.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(0)
        r_img = p_img.add_run()
        if os.path.isfile(lay["path"]):
            r_img.add_picture(lay["path"], width=LAYOUT_WIDTH)
            for sr in r_img._r.findall(f".//{{{A_NS}}}srcRect"):
                sr.attrib.clear()
        else:
            print(f"[WARNING] Khong thay so do {lay['path']}")
        add_cant_split(row)
    set_table_autofit_window(t4)

for r in ROOMS:
    d.locations.append(MonitoringLocation(
        name_vn=r["name_vn"],
        name_en=r["name_en"],
        area_m2=r["area_vn"],
        clean_class=r["class"],
        points_summary_vn=r["points_summary_vn"],
        points_summary_en=r["points_summary_en"],
        state_vn="Trạng thái hoạt động",
        state_en="Operational"
    ))

def remove_table(doc, table):
    """Safely removes a table from docx Document."""
    p = table._tbl.getparent()
    p.remove(table._tbl)

def remove_paragraph(doc, p):
    """Safely removes a paragraph from docx Document, preserving section breaks."""
    sectPr = p._p.find(f"{{{W_NS}}}pPr/{{{W_NS}}}sectPr")
    if sectPr is not None:
        return
    parent = p._p.getparent()
    if parent is not None:
        parent.remove(p._p)

def _followed_by_assessment(tbl):
    """True khi phan tu ngay sau bang la doan 'Đánh giá:' / 'Assessment:'."""
    nxt = tbl._tbl.getnext()
    if nxt is None or nxt.tag != f"{{{W_NS}}}p":
        return False
    txt = "".join(nxt.itertext()).strip().lower()
    return txt.startswith(("đánh giá", "assessment")) or bool(re.search(r"phù hợp|complied|comply", txt))


def keep_tail_together(result_tbl, note_tbl, last_rows=2):
    """
    Keeps the last rows of a result table, its 'Ghi chú/Note' table and the assessment paragraph that
    follows on the SAME page (keepNext chain), so the conclusion never drops onto a new page alone.
    Bang Ghi chu khong co cau Danh gia theo sau (nhom khu vuc khong co muc) -> dong cuoi bang Ghi chu
    khong dat keepNext, tranh keo ca muc sau sang trang moi (E32, AGENTS muc 3).
    """
    note_rows = note_tbl.rows
    chain_last = _followed_by_assessment(note_tbl)
    for tbl, rows in ((result_tbl, result_tbl.rows[-last_rows:]), (note_tbl, note_rows)):
        for row in rows:
            keep = chain_last or tbl is not note_tbl or row._tr is not note_rows[-1]._tr
            # moi o cua dong, ke ca o tiep noi cua o gop doc (row.cells tra ve o goc nen bo sot): thieu keepNext
            # o 1 o la Word van ngat trang giua cac dong nay
            for p_el in row._tr.iter(f"{{{W_NS}}}p"):
                Paragraph(p_el, None).paragraph_format.keep_with_next = keep


# ----------------------------------------------------------------------------------------------
# Helpers: fixed table layout, assessment paragraphs, "out of specification" test for bold values
# ----------------------------------------------------------------------------------------------
def apply_fixed_layout(table, widths):
    """Column widths that Word really honours (autofit squeezed the last columns to ~1 cm and wrapped numbers)."""
    set_table_col_widths(table, widths)
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(f"{{{W_NS}}}tblLayout"):
        tblPr.remove(old)
    lay = parse_xml(f'<w:tblLayout xmlns:w="{W_NS}" w:type="fixed"/>')
    anchor = next((c for c in tblPr if c.tag in (f"{{{W_NS}}}tblCellMar", f"{{{W_NS}}}tblLook")), None)
    if anchor is not None:
        anchor.addprevious(lay)
    else:
        tblPr.append(lay)


def set_assessment_paras(p, paras):
    """Writes the assessment into `p`; when there are two paragraphs (compliant areas on top, non-compliant
    areas below) the second one is inserted right after the first with identical formatting.
    paras rong (chi tieu khong danh gia) -> xoa luon doan Danh gia khoi van ban."""
    if not paras:
        p._p.getparent().remove(p._p)
        return
    cur = p
    for i, txt in enumerate(paras):
        if i > 0:
            new_p = copy.deepcopy(p._p)
            cur._p.addnext(new_p)
            cur = Paragraph(new_p, p._parent)
        set_assessment_paragraph(cur, txt, font_size=13, keep_next=False)
        pf = cur.paragraph_format
        last = (i == len(paras) - 1)
        pf.space_before = Pt(2)
        pf.space_after = Pt(2 if last else 4)
        pf.line_spacing = 1.0
        pf.page_break_before = False
        pf.keep_with_next = not last


def trim_empty_paragraphs(cell):
    """After a merge the cell keeps one empty paragraph per merged cell: that made 'Mức quy định' rows 2-3 lines
    taller than needed. Keep the first paragraph, drop the empty ones after it."""
    for extra in cell.paragraphs[1:]:
        if not extra.text.strip():
            extra._p.getparent().remove(extra._p)


def dust_fail(value, key):
    """Particle count above the ISO 14644-1 limit (class from the Excel 'Cấp sạch yêu cầu') of that channel."""
    lim = ISO8_LIMITS[key]
    return lim is not None and (value or 0) > lim


# Tong so trang phu luc (footer 'Trang X/Y'): 30/10/2026 tach bang nhiet am theo nhom muc + dong Trung binh -> 11 trang phu luc (TBKQ 12 trang).
# Doi so nay neu so trang thuc te thay doi (xem log 'Pages:' o buoc xuat PDF; UpChungThu bao 'Phu luc trang' neu lech).
TBKQ_APPENDIX_PAGES = 11


_MARK_RE = re.compile(r"\((\d)\)")


def superscript_note_markers(doc):
    """Moi chi so ghi chu (1)...(9) trong than van ban (tieu de cot, dong Ghi chu, doan van) -> chu nho dang superscript.
    Tach run chua '(n)' thanh nhieu run, giu nguyen dinh dang, chi them vertAlign=superscript cho '(n)'."""
    targets = []
    for root in [doc.element.body]:          # toan bo than van ban: bang, dong ghi chu, doan van
        for r_el in root.iter(f"{{{W_NS}}}r"):
            t_els = r_el.findall(f"{{{W_NS}}}t")
            if len(t_els) == 1 and _MARK_RE.search(t_els[0].text or "") and t_els[0].text.strip() not in [f"({i})" for i in range(10)]:
                targets.append((r_el, t_els))
            elif len(t_els) == 1 and _MARK_RE.fullmatch((t_els[0].text or "").strip()):
                targets.append((r_el, t_els))
    for r_el, t_els in targets:
        if True:
            text = t_els[0].text
            parts = [x for x in _MARK_RE.split(text)]
            # split -> [txt, digit, txt, digit, ...]
            pieces = []
            for i, part in enumerate(parts):
                if i % 2 == 1:
                    pieces.append((f"({part})", True))
                elif part:
                    pieces.append((part, False))
            anchor = r_el
            for txt, sup in pieces:
                new_r = copy.deepcopy(r_el)
                for el in new_r.findall(f"{{{W_NS}}}t"):
                    new_r.remove(el)
                rpr = new_r.find(f"{{{W_NS}}}rPr")
                if rpr is None:
                    rpr = parse_xml(f'<w:rPr xmlns:w="{W_NS}"/>')
                    new_r.insert(0, rpr)
                for old in rpr.findall(f"{{{W_NS}}}vertAlign"):
                    rpr.remove(old)
                if sup:
                    rpr.append(parse_xml(f'<w:vertAlign xmlns:w="{W_NS}" w:val="superscript"/>'))
                t = parse_xml(f'<w:t xmlns:w="{W_NS}" xml:space="preserve"></w:t>')
                t.text = txt
                new_r.append(t)
                anchor.addnext(new_r)
                anchor = new_r
            r_el.getparent().remove(r_el)


def _na_start():
    """So NA dau tien cua tung phong: NA danh so lien tuc toan ho so theo thu tu ROOMS (khop so do Bui_NhietDo),
    khong doi khi phong duoc xep vao bang nhom muc khac."""
    out, n = {}, 1
    for r in ROOMS:
        out[r["name_vn"]] = n
        n += len(temp_rh_points(r["name_vn"]))
    return out


def fill_trh_table(tbl, rooms, spec, is_en):
    """Bang nhiet do - do am cua 1 nhom phong cung muc quy dinh: moi diem 1 dong (TB 3 lan do sau hieu chuan),
    moi phong them dong 'Trung bình' (gia tri dem so voi muc, in dam neu khong dat); dong cuoi ghi muc cua nhom,
    nhom khong co muc (spec None) thi khong co dong muc (nguoi dung chot 30/10/2026)."""
    while len(tbl.rows) > 1:
        tbl.rows[-1]._tr.getparent().remove(tbl.rows[-1]._tr)
    hdr = ("Measurement room", "Measurement point (2)", "Temperature, °C", "Humidity, %RH") if is_en else \
          ("Khu vực đo", "Vị trí đo (2)", "Nhiệt độ, °C", "Độ ẩm, %RH")
    for c, txt in enumerate(hdr):
        set_cell_content(tbl.rows[0].cells[c], txt, bold=True, font_size=12, fill_hex="D9D9D9")

    na0 = _na_start()
    room_rows = []
    for r in rooms:
        name = r["name_vn"]
        rname = r["name_en"] if is_en else name
        start_row_idx = len(tbl.rows)
        for k, (t_avg, rh_avg) in enumerate(temp_rh_points(name)):
            row = tbl.add_row()
            set_cell_content(row.cells[0], rname if k == 0 else "", bold=True, font_size=12)
            set_cell_content(row.cells[1], f"NA{na0[name] + k}", align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            set_cell_content(row.cells[2], fmt_dec(t_avg, is_en), align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            set_cell_content(row.cells[3], fmt_dec(rh_avg, is_en), align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        t_room, h_room = temp_rh_room_avg(name)
        t_fail, h_fail = trh_room_fail(name)
        row = tbl.add_row()
        set_cell_content(row.cells[0], "", font_size=12)
        set_cell_content(row.cells[1], "Average" if is_en else "Trung bình", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[2], fmt_dec(t_room, is_en), bold=t_fail, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[3], fmt_dec(h_room, is_en), bold=h_fail, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        room_rows.append((start_row_idx, len(tbl.rows) - 1))

    if spec is not None:
        t_sp, h_sp = spec
        lim = tbl.add_row()
        set_cell_content(lim.cells[0], "Specifications provided by Lam Quang Dai Company Limited" if is_en
                         else "Mức quy định cung cấp bởi Công ty TNHH Lâm Quang Đại", bold=True, font_size=12)
        set_cell_content(lim.cells[1], "", font_size=12)  # merged into cells[0]; do not repeat the label
        set_cell_content(lim.cells[2], spec_txt(t_sp, is_en) if t_sp else "-", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(lim.cells[3], spec_txt(h_sp, is_en) if h_sp else "-", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        lim.cells[0].merge(lim.cells[1])
        trim_empty_paragraphs(lim.cells[0])

    # Vertical merge of the room name column over each room's points + its average row
    for s_idx, e_idx in room_rows:
        if e_idx > s_idx:
            cell_top = tbl.rows[s_idx].cells[0]
            for r_i in range(s_idx + 1, e_idx + 1):
                cell_top.merge(tbl.rows[r_i].cells[0])
            trim_empty_paragraphs(cell_top)

    format_table_structure(tbl, is_note=False, widths=[2800, 2500, 2050, 2050], row_height=340)  # col 2 wide enough for the EN header 'Measurement point (2)' on ONE line, so VN/EN tables have equal row heights and break at the same row
    apply_fixed_layout(tbl, [2800, 2500, 2050, 2050])
    add_tbl_header(tbl.rows[0])
    # So dong cuoi phai di lien Ghi chu + Danh gia: 3 dong cuoi cua phong cuoi (gom dong Trung bình) + dong muc.
    # Chi giu 2 dong thi Word ngat ngay truoc dong muc (o ten phong gop doc), dong muc roi mot minh sang trang moi.
    s_idx, e_idx = room_rows[-1]
    return min(e_idx - s_idx + 1, 3) + (1 if spec is not None else 0)


def fill_note(tbl, text, is_en):
    while len(tbl.rows) > 1:
        tbl.rows[-1]._tr.getparent().remove(tbl.rows[-1]._tr)
    set_cell_content(tbl.rows[0].cells[0], "Note:" if is_en else "Ghi chú:", bold=True, italic=True, font_size=10)
    set_cell_content(tbl.rows[0].cells[1], text, italic=True, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_table_structure(tbl, is_note=True)


def center_table_cells(table):
    """Căn giữa toàn bộ ô bảng kết quả: ngang (paragraph CENTER) + dọc (vAlign center)."""
    for tr in table._tbl.tr_lst:
        for tc in tr.tc_lst:
            tcPr = tc.get_or_add_tcPr()
            for old in tcPr.findall(f"{{{W_NS}}}vAlign"):
                tcPr.remove(old)
            tcPr.append(parse_xml(f'<w:vAlign xmlns:w="{W_NS}" w:val="center"/>'))
            for p_el in tc.iter(f"{{{W_NS}}}p"):
                Paragraph(p_el, None).alignment = WD_ALIGN_PARAGRAPH.CENTER


def build_tbkq(is_en=False, total_pages=TBKQ_APPENDIX_PAGES):
    template_file = os.path.join("Core", "Templates", "_goc_truoc_TAG", "TBKQ - en - phong sach.docx" if is_en else "TBKQ - viet - phong sach.docx")  # form gốc (không TAG): script này điền theo vị trí
    doc = docx.Document(template_file)

    # 1. Update Headers & Table 0
    cert_no = d.cert_no_en if is_en else d.cert_no_vn
    update_headers(doc, cert_no=cert_no, cert_date=d.cert_date, is_en=is_en)
    populate_cleanroom_table0(doc, dossier=d, is_en=is_en)

    # 2. Section 1 Body Paragraphs (Page 2)
    client_name = d.client.name_en if is_en else d.client.name_vn
    room_list_str = "; ".join([f"“{r['name_en'] if is_en else r['name_vn']}”" for r in ROOMS])
    site_str = d.location_desc_en if is_en else d.location_desc_vn

    if is_en:
        set_paragraph_text(doc.paragraphs[1], "Performance content", bold=True, font_size=13)
        set_paragraph_text(doc.paragraphs[2], f"Measurement area was indicated by {client_name}:", font_size=13)
        set_paragraph_text(doc.paragraphs[3], f"{room_list_str} at {site_str}", font_size=13)
        set_paragraph_text(doc.paragraphs[4], "Cleanroom occupancy state\t: “Operational”", font_size=13)
    else:
        set_paragraph_text(doc.paragraphs[1], "Nội dung thực hiện", bold=True, font_size=13)
        set_paragraph_text(doc.paragraphs[2], f"Khu vực thực hiện theo sự chỉ định của {client_name}:", font_size=13)
        set_paragraph_text(doc.paragraphs[3], f"{room_list_str} tại {site_str}", font_size=13)
        set_paragraph_text(doc.paragraphs[4], "Trạng thái phòng sạch\t: “Trạng thái hoạt động”", font_size=13)

    # 3. Table 1: Area Info
    t1 = doc.tables[1]
    while len(t1.rows) > 1:
        row = t1.rows[-1]
        row._tr.getparent().remove(row._tr)

    set_cell_content(t1.rows[0].cells[0], "Measurement location" if is_en else "Khu vực thực hiện", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t1.rows[0].cells[1], "Area (1)" if is_en else "Diện tích (1)", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t1.rows[0].cells[2], "Class (1)" if is_en else "Cấp sạch (1)", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t1.rows[0].cells[3], "Number of monitoring point" if is_en else "Số lượng vị trí quan trắc", bold=True, font_size=12, fill_hex="D9D9D9")

    for r in ROOMS:
        row = t1.add_row()
        rname = r["name_en"] if is_en else r["name_vn"]
        rarea = r["area_en"] if is_en else r["area_vn"]
        rsum = "\n".join(points_detail(r["name_vn"], is_en))  # moi chi tieu + diem do cua phong
        set_cell_content(row.cells[0], f"“{rname}”", align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[1], rarea, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[2], r["class"], align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[3], rsum, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=12)

    format_table_structure(t1, is_note=False, widths=[1900, 1400, 1450, 4650], row_height=425)
    apply_fixed_layout(t1, [1900, 1400, 1450, 4650])

    # 4. Table 2: Note 1
    t2 = doc.tables[2]
    while len(t2.rows) > 1:
        row = t2.rows[-1]
        row._tr.getparent().remove(row._tr)
    note1_txt = "(1) The cleanroom specifications provided by the representative of Lam Quang Dai Company Limited." if is_en else "(1) Thông số kỹ thuật phòng sạch do đại diện Công ty TNHH Lâm Quang Đại chỉ định."
    set_cell_content(t2.rows[0].cells[0], "Note:" if is_en else "Ghi chú:", bold=True, italic=True, font_size=10)
    set_cell_content(t2.rows[0].cells[1], note1_txt, italic=True, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_table_structure(t2, is_note=True)

    # 5. Table 3: Photos (2 columns, 7.5 cm width, strictly no crop)
    t3 = doc.tables[3]
    num_cols = 2
    num_rows = math.ceil(len(d.photos) / num_cols)
    img_w = PHOTO_WIDTH_2COL

    while len(t3.columns) < num_cols:
        t3.add_column(img_w)
    while len(t3.rows) < num_rows:
        t3.add_row()

    for idx, pinfo in enumerate(d.photos):
        r_idx = idx // num_cols
        c_idx = idx % num_cols
        cell = t3.rows[r_idx].cells[c_idx]
        for p in list(cell.paragraphs):
            p._p.getparent().remove(p._p)
        p_img = cell.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(0)
        r_img = p_img.add_run()
        if os.path.isfile(pinfo["path"]):
            r_img.add_picture(pinfo["path"], width=img_w)
            for sr in r_img._r.findall(f".//{{{A_NS}}}srcRect"):
                sr.attrib.clear()

    for row in t3.rows:
        add_cant_split(row)
    set_table_autofit_window(t3)

    # 6. Table 4: Layout diagrams (Sơ đồ 1–3)
    t4 = doc.tables[4]
    fill_layout_table(t4, is_en)

    # 7. Table 5: Airborne Particle Concentration (All 59 points)
    t5 = doc.tables[5]
    while len(t5.rows) > 2:
        row = t5.rows[-1]
        row._tr.getparent().remove(row._tr)

    # Setup headers
    set_cell_content(t5.rows[0].cells[0], "Measurement room" if is_en else "Khu vực đo", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[0].cells[1], "Measurement point (2)" if is_en else "Vị trí đo (2)", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[0].cells[2], "Airborne particle concentration, particles/m³" if is_en else "Nồng độ hạt bụi trong không khí, hạt/m³", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[0].cells[3], "Airborne particle concentration, particles/m³" if is_en else "Nồng độ hạt bụi trong không khí, hạt/m³", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[0].cells[4], "Airborne particle concentration, particles/m³" if is_en else "Nồng độ hạt bụi trong không khí, hạt/m³", bold=True, font_size=12, fill_hex="D9D9D9")

    set_cell_content(t5.rows[1].cells[0], "Measurement room" if is_en else "Khu vực đo", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[1].cells[1], "Measurement point (2)" if is_en else "Vị trí đo (2)", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[1].cells[2], "Particle size d ≥ 0.5 µm" if is_en else "Kích thước hạt bụi d ≥ 0,5 µm", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[1].cells[3], "Particle size d ≥ 1 µm" if is_en else "Kích thước hạt bụi d ≥ 1 µm", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[1].cells[4], "Particle size d ≥ 5 µm" if is_en else "Kích thước hạt bụi d ≥ 5 µm", bold=True, font_size=12, fill_hex="D9D9D9")

    # Populate 59 points
    room_start_row = {}
    for r in ROOMS:
        rname = r["name_en"] if is_en else r["name_vn"]
        start_sid = r["start_sample"]
        cnt = r["points_count"]
        start_row_idx = len(t5.rows)
        for pt_idx in range(1, cnt + 1):
            row = t5.add_row()
            cur_r = len(t5.rows) - 1
            sid = start_sid + pt_idx - 1
            c05, c10, c50 = PARTICLE_DATA.get(sid, (0, 0, 0))
            if pt_idx == 1:
                set_cell_content(row.cells[0], rname, bold=True, font_size=12)
            else:
                set_cell_content(row.cells[0], "", font_size=12)
            set_cell_content(row.cells[1], str(pt_idx), align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            # Values above the ISO 14644-1 Class 8 limit are printed in bold (non-compliant)
            set_cell_content(row.cells[2], fmt_int_en(c05) if is_en else fmt_int_vn(c05), bold=dust_fail(c05, "0.5"), align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            set_cell_content(row.cells[3], fmt_int_en(c10) if is_en else fmt_int_vn(c10), bold=dust_fail(c10, "1"), align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            set_cell_content(row.cells[4], fmt_int_en(c50) if is_en else fmt_int_vn(c50), bold=dust_fail(c50, "5"), align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        end_row_idx = len(t5.rows) - 1
        room_start_row[rname] = (start_row_idx, end_row_idx)

    # Standard Limit Row
    limit_row = t5.add_row()
    lim_lbl = f"Specified limits in ISO 14644-1:2015 for ISO Class {ISO_CLASS} cleanroom" if is_en else f"Mức quy định nêu tại ISO 14644-1:2015 đối với khu vực cấp sạch ISO {ISO_CLASS}"
    set_cell_content(limit_row.cells[0], lim_lbl, bold=True, font_size=12)
    set_cell_content(limit_row.cells[1], "", font_size=12)
    for _ci, _k in ((2, "0.5"), (3, "1"), (4, "5")):
        _lim = ISO8_LIMITS[_k]
        set_cell_content(limit_row.cells[_ci], f"≤ {_lim:,}".replace(",", " ") if _lim else "-", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
    limit_row.cells[0].merge(limit_row.cells[1])
    trim_empty_paragraphs(limit_row.cells[0])

    # Vertical merge Room Name column for each room
    for rname, (s_idx, e_idx) in room_start_row.items():
        if e_idx > s_idx:
            cell_top = t5.rows[s_idx].cells[0]
            for r_i in range(s_idx + 1, e_idx + 1):
                cell_top.merge(t5.rows[r_i].cells[0])
            trim_empty_paragraphs(cell_top)

    format_table_structure(t5, is_note=False, widths=[2300, 1300, 1950, 1950, 1900], row_height=340)
    apply_fixed_layout(t5, [2300, 1300, 1950, 1950, 1900])
    add_tbl_header(t5.rows[0])
    add_tbl_header(t5.rows[1])

    # 8. Note 2 under Table 5
    t6 = doc.tables[6]
    while len(t6.rows) > 1:
        row = t6.rows[-1]
        row._tr.getparent().remove(row._tr)
    note2_txt = f"{ACCREDIT_EN}\n(2) See measuring points in the layout of monitoring points." if is_en else f"{ACCREDIT_VN}\n(2) Vị trí đo xem sơ đồ các vị trí quan trắc."
    set_cell_content(t6.rows[0].cells[0], "Note:" if is_en else "Ghi chú:", bold=True, italic=True, font_size=10)
    set_cell_content(t6.rows[0].cells[1], note2_txt, italic=True, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_table_structure(t6, is_note=True)

    # 9. Delete 4 tables at index 7: 6-channel particle table, its note, microbes, and its note
    for _ in range(4):
        remove_table(doc, doc.tables[7])

    # 10. Table 7 (now Temp & Humidity) + 11. its note (index 8).
    # Danh gia theo trung binh ca phong; moi nhom phong cung muc quy dinh = 1 bang rieng (dong cuoi = muc cua nhom,
    # roi Ghi chu, roi Danh gia); nhom khong co muc: bang khong co dong muc va khong co Danh gia. Bang nhom 1 dung
    # bang cua form; cac nhom sau chen o buoc 16b (sau khi da lay het bang theo chi so doc.tables).
    t7 = doc.tables[7]
    t8 = doc.tables[8]
    trh_groups = trh_spec_groups()
    t7_blank, t8_blank = copy.deepcopy(t7._tbl), copy.deepcopy(t8._tbl)
    note_trh_txt = f"{ACCREDIT_EN}\n(2) See measuring points in the layout of monitoring points." if is_en else f"{ACCREDIT_VN}\n(2) Vị trí đo xem sơ đồ các vị trí quan trắc."
    trh_tails = [fill_trh_table(t7, trh_groups[0][1], trh_groups[0][0], is_en)]
    fill_note(t8, note_trh_txt, is_en)

    # 12. Table 9 (now Differential Pressure)
    t9 = doc.tables[9]
    while len(t9.rows) > 1:
        row = t9.rows[-1]
        row._tr.getparent().remove(row._tr)

    set_cell_content(t9.rows[0].cells[0], "Area 1" if is_en else "Khu vực 1", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t9.rows[0].cells[1], "", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t9.rows[0].cells[2], "Area 2" if is_en else "Khu vực 2", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t9.rows[0].cells[3], "Measuring point (2)" if is_en else "Vị trí đo (2)", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t9.rows[0].cells[4], "Differential pressure, Pa" if is_en else "Mức chênh áp, Pa", bold=True, font_size=12, fill_hex="D9D9D9")

    dp_eval = eval_dp()
    for pair in PRESSURE_PAIRS:
        row = t9.add_row()
        v1 = pair["v1_en" if is_en else "v1_vn"]
        v2 = pair["v2_en" if is_en else "v2_vn"]
        kh = pair["kh"]
        val = fmt_dec(dp_value(pair), is_en)
        set_cell_content(row.cells[0], v1, font_size=12)
        set_cell_content(row.cells[1], "→", align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[2], v2, font_size=12)
        set_cell_content(row.cells[3], kh, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[4], val, bold=(dp_eval[kh]["status"] == "fail"), align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)

    # Pressure Standard Limit Row: bo dong nay khi khong co muc quy dinh (Excel ghi N/A cho moi vi tri)
    if any(p["spec"] is not None for p in PRESSURE_PAIRS):
        dp_lim = t9.add_row()
        dp_lbl = "Specifications provided by Lam Quang Dai Company Limited" if is_en else "Mức quy định cung cấp bởi Công ty TNHH Lâm Quang Đại"
        set_cell_content(dp_lim.cells[0], dp_lbl, bold=True, font_size=12)
        for _ci in (1, 2, 3):
            set_cell_content(dp_lim.cells[_ci], "", font_size=12)
        _dp_specs = {p["spec"] for p in PRESSURE_PAIRS}
        if _dp_specs == {None}:
            _dp_txt = "Not specified" if is_en else "Không quy định"
        elif len(_dp_specs) == 1:
            _dp_txt = spec_txt(next(iter(_dp_specs)), is_en)
        else:
            _dp_txt = "According to each location" if is_en else "Theo từng vị trí"
        set_cell_content(dp_lim.cells[4], _dp_txt, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        dp_lim.cells[0].merge(dp_lim.cells[1])
        dp_lim.cells[0].merge(dp_lim.cells[2])
        dp_lim.cells[0].merge(dp_lim.cells[3])
        trim_empty_paragraphs(dp_lim.cells[0])

    format_table_structure(t9, is_note=False, widths=[2500, 450, 2450, 1800, 2200], row_height=340)
    apply_fixed_layout(t9, [2500, 450, 2450, 1800, 2200])
    add_tbl_header(t9.rows[0])

    # 13. Note Differential Pressure (now at index 10)
    t10 = doc.tables[10]
    while len(t10.rows) > 1:
        row = t10.rows[-1]
        row._tr.getparent().remove(row._tr)
    note_dp_txt = f"{ACCREDIT_EN}\n(2) See measuring points in the layout of monitoring points." if is_en else f"{ACCREDIT_VN}\n(2) Vị trí đo xem sơ đồ các vị trí quan trắc."
    set_cell_content(t10.rows[0].cells[0], "Note:" if is_en else "Ghi chú:", bold=True, italic=True, font_size=10)
    set_cell_content(t10.rows[0].cells[1], note_dp_txt, italic=True, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_table_structure(t10, is_note=True)

    # 14. Delete 5 tables at index 11: HEPA leak, note, placeholder 17, HEPA leak 2, note
    for _ in range(5):
        remove_table(doc, doc.tables[11])

    # 15. Table 11 (now Airflow Rate & ACH)
    t11 = doc.tables[11]
    while len(t11.rows) > 1:
        row = t11.rows[-1]
        row._tr.getparent().remove(row._tr)

    set_cell_content(t11.rows[0].cells[0], "Measurement room" if is_en else "Khu vực đo", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t11.rows[0].cells[1], "Measurement point (2)" if is_en else "Vị trí đo (2)", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t11.rows[0].cells[2], "Airflow rate, m³/h" if is_en else "Lưu lượng gió, m³/h", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t11.rows[0].cells[3], "Total airflow rate, m³/h" if is_en else "Tổng lưu lượng, m³/h", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t11.rows[0].cells[4], "Room volume (3), m³" if is_en else "Thể tích phòng (3), m³", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t11.rows[0].cells[5], "Air change rate, ACH" if is_en else "Số lần trao đổi khí, ACH", bold=True, font_size=12, fill_hex="D9D9D9")

    flow_rooms = [r for r in ROOMS if r.get("grilles")]
    ach_eval = eval_ach()
    flow_room_rows = {}
    for r in flow_rooms:
        rname = r["name_en"] if is_en else r["name_vn"]
        flow_room_rows[rname] = []
        for g_idx, g in enumerate(r["grilles"]):
            row = t11.add_row()
            cur_r = len(t11.rows) - 1
            flow_room_rows[rname].append(cur_r)
            gname = g["name_en" if is_en else "name_vn"]
            gflow = fmt_flow(g["flow"])
            # Values shared by all grilles of a room are written ONLY in the first row; the cells of the
            # following rows stay empty and are vertically merged below (no duplicated text after merge).
            first = (g_idx == 0)
            set_cell_content(row.cells[0], rname if first else "", bold=True, font_size=12)
            set_cell_content(row.cells[1], gname, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            set_cell_content(row.cells[2], gflow, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            set_cell_content(row.cells[3], fmt_flow(r["flow_tot"]) if first else "", align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            set_cell_content(row.cells[4], fmt_dec(r["vol"], is_en) if first else "", align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
            set_cell_content(row.cells[5], fmt_dec(ach_value(r), is_en) if first else "", bold=(first and ach_eval[r["name_vn"]]["status"] == "fail"), align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)

    # Airflow Standard Limit Row
    # Airflow limit row: bo dong nay khi khong co muc quy dinh ACH
    if any(r.get("ach_spec") is not None for r in flow_rooms):
        flow_lim = t11.add_row()
        fl_lbl = "Specifications provided by Lam Quang Dai Company Limited" if is_en else "Mức quy định cung cấp bởi Công ty TNHH Lâm Quang Đại"
        set_cell_content(flow_lim.cells[0], fl_lbl, bold=True, font_size=12)
        for _ci in (1, 2, 3, 4):
            set_cell_content(flow_lim.cells[_ci], "", font_size=12)
        _ach_specs = {r.get("ach_spec") for r in flow_rooms}
        _ach_txt = spec_txt(next(iter(_ach_specs)), is_en) if len(_ach_specs) == 1 else ("According to each room" if is_en else "Theo từng phòng")
        set_cell_content(flow_lim.cells[5], _ach_txt, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        flow_lim.cells[0].merge(flow_lim.cells[1])
        flow_lim.cells[0].merge(flow_lim.cells[2])
        flow_lim.cells[0].merge(flow_lim.cells[3])
        flow_lim.cells[0].merge(flow_lim.cells[4])
        trim_empty_paragraphs(flow_lim.cells[0])

    # Vertical merge Room Name, Total flow, Volume, ACH for multiple grilles
    for rname, r_idxs in flow_room_rows.items():
        if len(r_idxs) > 1:
            top_0 = t11.rows[r_idxs[0]].cells[0]
            top_3 = t11.rows[r_idxs[0]].cells[3]
            top_4 = t11.rows[r_idxs[0]].cells[4]
            top_5 = t11.rows[r_idxs[0]].cells[5]
            for r_i in r_idxs[1:]:
                top_0.merge(t11.rows[r_i].cells[0])
                top_3.merge(t11.rows[r_i].cells[3])
                top_4.merge(t11.rows[r_i].cells[4])
                top_5.merge(t11.rows[r_i].cells[5])
            for _c in (top_0, top_3, top_4, top_5):
                trim_empty_paragraphs(_c)

    format_table_structure(t11, is_note=False, widths=[2000, 1700, 1400, 1500, 1500, 1300], row_height=360)
    apply_fixed_layout(t11, [2000, 1700, 1400, 1500, 1500, 1300])
    add_tbl_header(t11.rows[0])

    # 16. Note Airflow (now at index 12)
    t12 = doc.tables[12]
    while len(t12.rows) > 1:
        row = t12.rows[-1]
        row._tr.getparent().remove(row._tr)
    note_fl_txt = f"{ACCREDIT_EN}\n(2) See measuring points in the layout of monitoring points.\n(3) Room volume is calculated based on floor area and ceiling height H = 2.80 m." if is_en else f"{ACCREDIT_VN}\n(2) Vị trí đo xem sơ đồ các vị trí quan trắc.\n(3) Thể tích phòng được tính toán trên cơ sở diện tích và chiều cao trần H = 2,80 m."
    set_cell_content(t12.rows[0].cells[0], "Note:" if is_en else "Ghi chú:", bold=True, italic=True, font_size=10)
    set_cell_content(t12.rows[0].cells[1], note_fl_txt, italic=True, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_table_structure(t12, is_note=True)

    # 16b. Nhiet do - do am: chen bang + Ghi chu + doan Danh gia cua cac nhom muc con lai ngay sau doan Danh gia cua
    # nhom truoc. Doan Danh gia mau cua form nam sau bang Ghi chu (t8).
    el = t8._tbl.getnext()
    trh_p = None
    while el is not None and el.tag != f"{{{W_NS}}}tbl":
        _t = "".join(x.text or "" for x in el.iter(f"{{{W_NS}}}t")).lower()
        if ("nhiệt độ" in _t or "temperature" in _t) and ("phù hợp" in _t or "complied" in _t):
            trh_p = el
            break
        el = el.getnext()
    if trh_p is None:
        raise ValueError("Form TBKQ: khong thay doan Danh gia nhiet do - do am sau bang Ghi chu")
    p_blank = copy.deepcopy(trh_p)
    trh_pairs = [(t7, t8)]
    trh_assess = [(trh_p, trh_groups[0])]
    anchor = trh_p
    for grp in trh_groups[1:]:
        tb, nb, pb = copy.deepcopy(t7_blank), copy.deepcopy(t8_blank), copy.deepcopy(p_blank)
        anchor.addnext(tb)
        tb.addnext(nb)
        nb.addnext(pb)
        anchor = pb
        tt, nt = Table(tb, t7._parent), Table(nb, t8._parent)
        trh_tails.append(fill_trh_table(tt, grp[1], grp[0], is_en))
        fill_note(nt, note_trh_txt, is_en)
        trh_pairs.append((tt, nt))
        trh_assess.append((pb, grp))
    for p_el, (spec, rooms) in trh_assess:
        set_assessment_paras(Paragraph(p_el, t8._parent), assess_temp_rh(is_en, rooms) if spec is not None else [])
    trh_tables = [x for x, _ in trh_pairs]
    for tbl in trh_tables[1:]:          # cach bang nhom sau khoi cau Danh gia cua nhom truoc
        prev = tbl._tbl.getprevious()
        if prev is not None and prev.tag == f"{{{W_NS}}}p":
            Paragraph(prev, None).paragraph_format.space_after = Pt(8)

    # Apply compact padding & tight line spacing on result tables
    for tbl in [t5, *trh_tables, t9, t11]:
        for r in tbl.rows:
            for c in r.cells:
                set_cell_margins(c, top=10, bottom=10, left=60, right=60)
                for cp in c.paragraphs:
                    cp.paragraph_format.space_before = Pt(0)
                    cp.paragraph_format.space_after = Pt(0)
                    cp.paragraph_format.line_spacing = 1.0
                    # form goc co thut le am (vd -244 twips o 'Measuring point') -> chu tran ra ngoai o, bi cat
                    if cp._p.pPr is not None:
                        for _ind in cp._p.pPr.findall(f"{{{W_NS}}}ind"):
                            cp._p.pPr.remove(_ind)
        center_table_cells(tbl)

    # 17. Body Paragraphs & Assessments cleanup
    seen_particle_assessment = False

    # Pass 1: Remove obsolete sections & duplicate assessments
    for p in list(doc.paragraphs):
        txt = p.text.strip()
        txt_lower = txt.lower()

        # Remove standalone "Đánh giá:" / "Assessment:" label paragraphs
        if txt_lower in ["đánh giá:", "assessment:"]:
            remove_paragraph(doc, p)
            continue

        # Remove template option b
        if "mẫu b - bảng sơ đồ" in txt_lower or "template option b" in txt_lower:
            remove_paragraph(doc, p)
            continue

        # Remove 6-channel heading
        if "6 kênh hạt" in txt_lower or "6 particle sizes" in txt_lower:
            remove_paragraph(doc, p)
            continue

        # Remove microbes section
        if "vi sinh vật" in txt_lower or "microbial" in txt_lower or "microorganisms" in txt_lower:
            remove_paragraph(doc, p)
            continue

        # Remove HEPA leaks & scan sections (preserving airflow)
        if ("rò rỉ" in txt_lower or "leak" in txt_lower or "hepa" in txt_lower) and \
           not any(k in txt_lower for k in ["lưu lượng", "airflow", "trao đổi khí", "air change"]):
            remove_paragraph(doc, p)
            continue

        # Check for particle assessments: keep first (3-channel), delete second (6-channel)
        if "nồng độ hạt bụi trong không khí" in txt_lower or "airborne particle concentration in" in txt_lower:
            if seen_particle_assessment:
                remove_paragraph(doc, p)
                continue
            else:
                seen_particle_assessment = True

    # Pass 2: Update remaining headings & assessments with exact logic
    for p in list(doc.paragraphs):
        txt = p.text.strip()
        txt_lower = txt.lower()

        # Particle heading
        if "3 kênh hạt" in txt_lower or "3 particle sizes" in txt_lower:
            set_paragraph_text(p, "Measurement result of airborne particle concentration" if is_en else "Kết quả đo nồng độ hạt bụi trong không khí", bold=True, font_size=13)
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0

        # Particle assessment
        elif ("nồng độ hạt bụi trong không khí" in txt_lower or "airborne particle concentration in" in txt_lower) and \
             ("phù hợp" in txt_lower or "complied" in txt_lower):
            set_assessment_paras(p, assess_dust(is_en))

        # Temp/RH assessment (MUST check before heading because assessment text has "nhiệt độ, độ ẩm"):
        # da dien theo tung nhom muc o buoc 16b, khong ghi lai
        elif ("nhiệt độ" in txt_lower or "temperature" in txt_lower) and \
             ("phù hợp" in txt_lower or "complied" in txt_lower or "comply" in txt_lower):
            pass

        # Temp/RH heading
        elif ("nhiệt độ" in txt_lower or "temperature" in txt_lower) and \
             ("kết quả" in txt_lower or "measurement result" in txt_lower):
            set_paragraph_text(p, "Measurement result of temperature and humidity" if is_en else "Kết quả đo nhiệt độ - độ ẩm", bold=True, font_size=13)
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0

        # DP assessment (MUST check before heading because assessment text has "chênh lệch áp suất")
        elif ("áp suất" in txt_lower or "pressure" in txt_lower) and \
             ("phù hợp" in txt_lower or "complied" in txt_lower):
            set_assessment_paras(p, assess_dp(is_en))

        # DP heading
        elif ("áp suất" in txt_lower or "pressure" in txt_lower) and \
             ("kết quả" in txt_lower or "measurement result" in txt_lower):
            set_paragraph_text(p, "Measurement result of pressure differential between areas" if is_en else "Kết quả đo chênh lệch áp suất giữa các khu vực", bold=True, font_size=13)
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0

        # Airflow & ACH assessment (check before heading)
        elif ("trao đổi khí" in txt_lower or "air change" in txt_lower or "lưu lượng" in txt_lower) and \
             ("phù hợp" in txt_lower or "complied" in txt_lower):
            set_assessment_paras(p, assess_ach(is_en))

        # Airflow heading
        elif ("lưu lượng" in txt_lower or "airflow" in txt_lower) and \
             ("kết quả" in txt_lower or "measurement result" in txt_lower):
            set_paragraph_text(p, "Measurement result of supply airflow rate and air change rate" if is_en else "Kết quả đo lưu lượng khí cấp và số lần trao đổi khí", bold=True, font_size=13)
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0

    # Pass 3: Section headings (khong ngat trang cung)
    headings_to_break = [
        "hình ảnh quá trình", "photos of the",
        "sơ đồ các vị trí", "layout of monitoring",
        "kết quả đo nồng độ hạt bụi trong không khí", "measurement result of airborne particle",
        "kết quả đo nhiệt độ - độ ẩm", "measurement result of temperature and humidity",
        "kết quả đo chênh lệch áp suất", "measurement result of pressure differential",
        "kết quả đo lưu lượng khí cấp", "measurement result of supply airflow"
    ]
    for p in list(doc.paragraphs):
        txt_lower = p.text.strip().lower()
        # The 2nd assessment paragraph (non-compliant areas) starts with the same words as a section heading:
        # it is recognised by its evaluation keyword and must NOT get a page break.
        is_assessment = bool(re.search(r"phù hợp|complied|comply", txt_lower))
        if not txt_lower.startswith("đánh giá") and not txt_lower.startswith("assessment") and not is_assessment:
            if any(h in txt_lower for h in headings_to_break):
                # Khong ep sang trang moi (de tiet kiem giay): cac muc chay lien nhau, tieu de luon di cung bang/anh ngay sau
                p.paragraph_format.page_break_before = False
                p.paragraph_format.keep_with_next = True
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(2)

    # Pass 4: Remove empty paragraphs between sections
    for p in list(doc.paragraphs):
        if not p.text.strip():
            remove_paragraph(doc, p)

    # Dong cuoi bang + bang Ghi chu + cau Danh gia nam cung trang (khong de cau Danh gia roi mot minh sang trang moi)
    for _res, _note in ((t5, t6), (t9, t10), (t11, t12)):
        keep_tail_together(_res, _note, last_rows=2)
    for (_res, _note), _n in zip(trh_pairs, trh_tails):
        keep_tail_together(_res, _note, last_rows=_n)
    # Dau bang: tieu de muc + dong tieu de + vai dong dau di lien (khong de tieu de/1 dong le loi o cuoi trang)
    for _res, _n in ((t5, 5), *[(x, 4) for x in trh_tables], (t9, 4), (t11, 4)):
        for _row in _res.rows[:_n]:
            for _c in _row.cells:
                for _p in _c.paragraphs:
                    _p.paragraph_format.keep_with_next = True
    # Bo hard page break con sot trong template (giu section break cua trang bia)
    for _br in list(doc.element.body.iter(f"{{{W_NS}}}br")):
        if _br.get(f"{{{W_NS}}}type") == "page":
            _br.getparent().remove(_br)
    for _pb in list(doc.element.body.iter(f"{{{W_NS}}}pageBreakBefore")):
        _pb.getparent().remove(_pb)

    superscript_note_markers(doc)

    # 18. Add Ending Line (TBKQ only)
    add_document_ending_line(doc)
    strip_all_highlights(doc)
    strip_image_crops(doc)

    # 19. Update Section 1 footer total appendix pages (cover page excluded)
    for s in doc.sections:
        for p in s.footer.paragraphs:
            for i, r in enumerate(p.runs):
                if r.text.strip() == "/" and i + 1 < len(p.runs):
                    next_r = p.runs[i + 1]
                    if next_r.text.strip().isdigit():
                        next_r.text = str(total_pages)

    out_dir = os.path.join(d.input_dir, "TBKQ")
    os.makedirs(out_dir, exist_ok=True)
    out_file = f"TBKQ - {'en' if is_en else 'viet'} - 1 - phong sach - {d.code}.docx"
    target_path = os.path.join(out_dir, out_file)
    doc.save(target_path)
    print(f"Successfully generated TBKQ -> {target_path}")
    return target_path

def set_header_page_fields(doc):
    """Header BCKQ trong form goc ghi cung 'Trang 1/3' -> thay bang truong Word PAGE / NUMPAGES."""
    for sec in doc.sections:
        for part in (sec.header, sec.first_page_header):
            if part is None:
                continue
            for p_el in part._element.iter(f"{{{W_NS}}}p"):
                para = Paragraph(p_el, part)
                if not re.fullmatch(r"\s*Trang\s*\d+\s*/\s*\d+\s*", para.text or ""):
                    continue
                runs = p_el.findall(f"{{{W_NS}}}r")
                rpr = runs[0].find(f"{{{W_NS}}}rPr") if runs else None
                rpr_xml = rpr.xml if rpr is not None else ""
                for r_el in runs:
                    p_el.remove(r_el)
                def run(text):
                    return parse_xml(f'<w:r xmlns:w="{W_NS}">{rpr_xml}<w:t xml:space="preserve">{text}</w:t></w:r>')
                def field(code, shown):
                    return parse_xml(f'<w:fldSimple xmlns:w="{W_NS}" w:instr=" {code} \\* MERGEFORMAT "><w:r>{rpr_xml}<w:t>{shown}</w:t></w:r></w:fldSimple>')
                for el in (run("Trang "), field("PAGE", "1"), run("/"), field("NUMPAGES", "1")):
                    p_el.append(el)


def build_bckq():
    template_file = os.path.join("Core", "Templates", "_goc_truoc_TAG", "BCKQ - phong sach.docx")  # form gốc (không TAG): script này điền theo vị trí
    doc = docx.Document(template_file)

    # 1. Update Table 0 (aligned top, 3pt space before/after)
    populate_bckq_table0(doc.tables[0], d)

    # 2. Section 2: Nội dung thực hiện
    client_name = d.client.name_vn
    room_list_str = "; ".join([f"“{r['name_vn']}”" for r in ROOMS])
    site_str = d.location_desc_vn

    for p in doc.paragraphs:
        txt = p.text.strip()
        if "Khu vực thực hiện theo sự chỉ định của" in txt:
            set_paragraph_text(p, f"Khu vực thực hiện theo sự chỉ định của {client_name}:", font_size=13)
        elif "“………”; “………”" in txt:
            set_paragraph_text(p, f"{room_list_str} tại {site_str}", font_size=13)
        elif "Trạng thái phòng sạch" in txt:
            set_paragraph_text(p, "Trạng thái phòng sạch\t: “Trạng thái hoạt động”", font_size=13)

    # 3. Table 1: Area Info
    t1 = doc.tables[1]
    while len(t1.rows) > 1:
        row = t1.rows[-1]
        row._tr.getparent().remove(row._tr)

    set_cell_content(t1.rows[0].cells[0], "Khu vực thực hiện", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t1.rows[0].cells[1], "Diện tích (1)", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t1.rows[0].cells[2], "Cấp sạch (1)", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t1.rows[0].cells[3], "Số lượng vị trí quan trắc", bold=True, font_size=12, fill_hex="D9D9D9")

    for r in ROOMS:
        row = t1.add_row()
        rname = r["name_vn"]
        rarea = r["area_vn"]
        rsum = "\n".join(points_detail(r["name_vn"], False))
        set_cell_content(row.cells[0], f"“{rname}”", align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[1], rarea, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[2], r["class"], align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[3], rsum, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=12)

    format_table_structure(t1, is_note=False, widths=[1900, 1400, 1450, 4595], row_height=425)

    # 4. Table 2: Note 1
    t2 = doc.tables[2]
    while len(t2.rows) > 1:
        row = t2.rows[-1]
        row._tr.getparent().remove(row._tr)
    note1_txt = "(1) Thông số kỹ thuật phòng sạch do đại diện Công ty TNHH Lâm Quang Đại chỉ định."
    set_cell_content(t2.rows[0].cells[0], "Ghi chú:", bold=True, italic=True, font_size=10)
    set_cell_content(t2.rows[0].cells[1], note1_txt, italic=True, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_table_structure(t2, is_note=True)

    # 5. Table 3: Photos (2 columns, 7.5 cm width, strictly no crop)
    t3 = doc.tables[3]
    num_cols = 2
    num_rows = math.ceil(len(d.photos) / num_cols)
    img_w = PHOTO_WIDTH_2COL

    while len(t3.columns) < num_cols:
        t3.add_column(img_w)
    while len(t3.rows) < num_rows:
        t3.add_row()

    for idx, pinfo in enumerate(d.photos):
        r_idx = idx // num_cols
        c_idx = idx % num_cols
        cell = t3.rows[r_idx].cells[c_idx]
        for p in list(cell.paragraphs):
            p._p.getparent().remove(p._p)
        p_img = cell.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(0)
        r_img = p_img.add_run()
        if os.path.isfile(pinfo["path"]):
            r_img.add_picture(pinfo["path"], width=img_w)
            for sr in r_img._r.findall(f".//{{{A_NS}}}srcRect"):
                sr.attrib.clear()

    for row in t3.rows:
        add_cant_split(row)
    set_table_autofit_window(t3)

    # 6. Table 4: Layout diagrams (Sơ đồ 1–3)
    t4 = doc.tables[4]
    fill_layout_table(t4)

    # 7. Table 5: Summary Evaluation Table
    t5 = doc.tables[5]
    while len(t5.rows) > 1:
        row = t5.rows[-1]
        row._tr.getparent().remove(row._tr)

    set_cell_content(t5.rows[0].cells[0], "Khu vực đo", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[0].cells[1], "Chỉ tiêu đạt", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[0].cells[2], "Chỉ tiêu không đạt", bold=True, font_size=12, fill_hex="D9D9D9")
    set_cell_content(t5.rows[0].cells[3], "Căn cứ đánh giá", bold=True, font_size=12, fill_hex="D9D9D9")

    # Chi tieu dat / khong dat theo tung phong: tinh tu cung ham eval_* voi TBKQ (khop ket luan TBKQ, AGENTS.md muc 7)
    summary = room_summary()
    for r in ROOMS:
        row = t5.add_row()
        sm = summary[r["name_vn"]]
        good_txt = "\n".join(f"- {x}" for x in sm["good"])
        if sm["na"]:
            good_txt += ("\n" if good_txt else "") + "Không đánh giá (không có mức quy định): " + ", ".join(x[:1].lower() + x[1:] for x in sm["na"])
        good_txt = good_txt or "Không"
        bad_txt = "\n".join(f"- {x}" for x in sm["bad"]) or "Không"
        set_cell_content(row.cells[0], f"“{r['name_vn']}”", align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
        set_cell_content(row.cells[1], good_txt, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=12)
        set_cell_content(row.cells[2], bad_txt, align=WD_ALIGN_PARAGRAPH.CENTER if not sm["bad"] else WD_ALIGN_PARAGRAPH.LEFT, font_size=12)
        basis = sm["basis"].replace("mức quy định do Công ty TNHH Lâm Quang Đại cung cấp", "mức quy định của khách hàng")
        set_cell_content(row.cells[3], basis, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)

    format_table_structure(t5, is_note=False, widths=[1650, 2650, 3000, 2000], row_height=425)
    for tr in t5._tbl.tr_lst:  # can giua theo chieu doc; giu can ngang tung o (danh sach chi tieu can trai)
        for tc in tr.tc_lst:
            tcPr = tc.get_or_add_tcPr()
            for old in tcPr.findall(f"{{{W_NS}}}vAlign"):
                tcPr.remove(old)
            tcPr.append(parse_xml(f'<w:vAlign xmlns:w="{W_NS}" w:val="center"/>'))
    for c in t5.rows[0].cells:
        for p_ in c.paragraphs:
            p_.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 8. Section 4: Results & Section 5: Assessment & Inspector Report
    for p in list(doc.paragraphs):
        txt = p.text.strip()
        if "- kết quả xác định mức độ nhiễm bẩn vi sinh vật" in txt.lower():
            remove_paragraph(doc, p)
        elif "- kết quả đo nồng độ hạt bụi" in txt.lower():
            res_text = (
                f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {d.code} ngày {d.cert_date}\n"
                f"- Kết quả đo nhiệt độ, độ ẩm: xem tại phiếu Tính toán hiệu chuẩn nhiệt độ, độ ẩm số {d.code} ngày {d.cert_date}\n"
                f"- Kết quả đo chênh lệch áp suất: xem tại phiếu Tính toán chênh áp số {d.code} ngày {d.cert_date}\n"
                f"- Kết quả tính toán lưu lượng và số lần trao đổi khí: xem tại phiếu Tính toán lưu lượng và số lần trao đổi gió số {d.code} ngày {d.cert_date}"
            )
            set_paragraph_text(p, res_text, font_size=13, align=WD_ALIGN_PARAGRAPH.LEFT)
        elif "- kết quả đo nhiệt độ" in txt.lower():
            remove_paragraph(doc, p)
        elif txt.startswith("Đánh giá:"):
            # Danh gia BCKQ = dung cac cau danh gia cua TBKQ (assess_*), moi chi tieu 1 dong '- ...'
            items = []
            for fn in (assess_dust, assess_temp_rh, assess_dp, assess_ach):
                for para in fn(False):
                    para = para[len("Đánh giá: "):] if para.startswith("Đánh giá: ") else para
                    items.append("- " + para)
            set_paragraph_text(p, "Đánh giá:", bold=True, font_size=13)
            new_p = copy.deepcopy(p._p)
            p._p.addnext(new_p)
            set_assessment_paras(Paragraph(new_p, p._parent), items)
        elif "ngày kết thúc hồ sơ thực tế" in txt.lower():
            set_paragraph_text(p, f"Ngày kết thúc hồ sơ thực tế: {d.cert_date}", font_size=13)
        elif "thời gian trễ hạn:" in txt.lower():
            set_paragraph_text(p, "Thời gian trễ hạn: Không", font_size=13)
        elif "nguyên nhân trễ hạn:" in txt.lower():
            set_paragraph_text(p, "Nguyên nhân trễ hạn: Không", font_size=13)
        elif "trễ hạn:" in txt.lower():
            set_paragraph_text(p, "Trễ hạn: \tcó\t\tkhông [x]", font_size=13)
        elif "các bất thường trong quá trình" in txt.lower():
            set_paragraph_text(p, "Các bất thường trong quá trình giám định: Không", font_size=13)
        elif "báo cáo:" in txt.lower():
            if "ngày báo cáo" in txt.lower() or "ngày báo cáo" in txt.lower():
                set_paragraph_text(p, f"\t\t\t\t\t\tNgày báo cáo: {d.cert_date}", font_size=13)
            elif "người báo cáo" in txt.lower():
                set_paragraph_text(p, f"\t\t\t\t\t\tNgười báo cáo: {d.personnel.reporter}", font_size=13)

    # 'So do cac vi tri do:' khong bi tach khoi so do o trang sau
    for p in doc.paragraphs:
        if p.text.strip().lower().startswith(("sơ đồ các vị trí đo", "hình ảnh quá trình")):
            p.paragraph_format.keep_with_next = True
    set_header_page_fields(doc)

    superscript_note_markers(doc)

    # Note: NO ending line in BCKQ (per AGENTS.md Rule 5)
    strip_all_highlights(doc)
    strip_image_crops(doc)

    out_dir = os.path.join(d.input_dir, "ThucHien")
    os.makedirs(out_dir, exist_ok=True)
    out_file = f"BCKQ - 1 - phong sach - {d.code}.docx"
    target_path = os.path.join(out_dir, out_file)
    doc.save(target_path)
    print(f"Successfully generated BCKQ -> {target_path}")
    return target_path

def _pdf_pages(docx_path):
    """Number of pages of the PDF exported from `docx_path` (None if it cannot be read)."""
    pdf = os.path.join(d.input_dir, "pdf", os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    try:
        import fitz
        with fitz.open(pdf) as _doc:
            return _doc.page_count
    except Exception as e:
        print(f"[Warning] Khong doc duoc so trang cua {pdf}: {type(e).__name__}: {e}")
        return None


if __name__ == "__main__":
    tbkq_vn = build_tbkq(is_en=False)
    tbkq_en = build_tbkq(is_en=True)
    bckq = build_bckq()
    print("All documents generated successfully!")

    # 4. Export PDFs & UpChungThu integration
    from Core.Exporters.pdf_exporter import PDFExporter
    exporter = PDFExporter(d)
    print("\n--- Exporting DOCX to PDF ---")
    exporter.export_files([tbkq_vn, tbkq_en, bckq])

    # The footer 'Trang X/Y' needs the real number of appendix pages: read it from the exported PDFs and, if it
    # differs from the value used, rebuild + re-export the two TBKQ once.
    n_vn, n_en = _pdf_pages(tbkq_vn), _pdf_pages(tbkq_en)
    if n_vn and n_en:
        if n_vn != n_en:
            print(f"[WARNING] TBKQ tieng Viet {n_vn} trang, tieng Anh {n_en} trang: khac nhau (AGENTS.md muc 9 yeu cau bang nhau).")
        elif n_vn - 1 != TBKQ_APPENDIX_PAGES:
            print(f"[Auto] So trang phu luc thuc te = {n_vn - 1} (da dung {TBKQ_APPENDIX_PAGES}): dung lai TBKQ va xuat PDF lai.")
            tbkq_vn = build_tbkq(is_en=False, total_pages=n_vn - 1)
            tbkq_en = build_tbkq(is_en=True, total_pages=n_vn - 1)
            exporter.export_files([tbkq_vn, tbkq_en])

    print("\n--- Exporting Excel Calculation to PDF ---")
    calc_xlsx = os.path.join(d.input_dir, "ThucHien", f"{d.code} - TinhToanPhongSach.xlsx")
    if os.path.isfile(calc_xlsx):
        exporter.export_excel_files([calc_xlsx])
