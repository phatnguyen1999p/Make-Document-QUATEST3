# -*- coding: utf-8 -*-
"""
Dựng dữ liệu (context) cho các template dạng TAG (Core/Templates/*.docx, *.xlsx).

Mỗi template dùng chung tên TAG cho bản tiếng Việt và tiếng Anh; context tự chọn ngôn ngữ.
Danh mục TAG đầy đủ: Core/Templates/TAGS.md.

Quy tắc nghiệp vụ được giữ nguyên như code cũ (Styling/elements.py: update_headers, populate_cleanroom_table0):
- Số chứng thư có /1, /2, /E chỉ ở header; thân văn bản dùng mã trần (AGENTS.md mục 2).
- Thời gian giám định: VN "Từ ngày dd/mm/yyyy", EN "From dd/mm/yyyy" (không ghi giờ).
- Thiết bị: mỗi dòng "- <tên>;", dòng cuối kết thúc bằng ".".
- Chữ ký: GIÁM ĐỊNH VIÊN / INSPECTOR nằm ở dòng 2, ngang hàng PHÓ GIÁM ĐỐC / VICE DIRECTOR.
"""

import math
import os
import re
from typing import Any, Dict, Optional

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm, Pt

from ..Models.dossier import Dossier, inspection_scope, inspection_procedure, bsc_class_of
from ..Styling import (
    W_NS,
    PHOTO_WIDTH_2COL,
    PHOTO_WIDTH_3COL,
    add_cant_split,
    set_paragraph_text,
    set_assessment_paragraph,
    set_table_autofit_window,
)
from ..Styling.elements import normalize_company_name
from ..Tools.standards import cite_standards
from ..Tools.tag_engine import TagBlock, render_docx

PLACEHOLDER = "………"

# Quy trình mặc định theo loại hồ sơ (khi Dossier.procedure_* để trống)
DEFAULT_PROCEDURE = {
    "cleanroom": "QUATEST 3 3030:2025, N2.TC.PS.01",
    "compressed_air": "QUATEST 3 3030:2025, N2.TC.KN.01",
    "bsc": "QUATEST 3 3030:2025, N2.TC.BSC.02",   # 1 mã theo cấp tủ (người dùng chốt 06/11/2026); cấp I -> BSC.01 qua bsc_class_of
    "laf": "QUATEST 3 3030:2025, N2.TC.LAF.01",
    "laf_bsc": "QUATEST 3 3030:2025, N2.TC.LAF.01",
    "fumehood": "QUATEST 3 3030:2025, N2.TC.T.01",
}


def _default_procedure(d, dtype: str, is_en: bool) -> str:
    """Quy trình theo BBLM (Dossier.procedure_*); trống thì mặc định theo loại — tủ ATSH cấp I là BSC.01."""
    own = (d.procedure_en if is_en else d.procedure_vn)
    if own:
        return own
    if dtype == "bsc" and bsc_class_of(d) == 1:
        return "QUATEST 3 3030:2025, N2.TC.BSC.01"
    return DEFAULT_PROCEDURE.get(dtype, DEFAULT_PROCEDURE["cleanroom"])


def norm_type(dossier_type: str) -> str:
    t = (dossier_type or "").lower().strip()
    if t in ("compressed_air", "khi nen", "khí nén"):
        return "compressed_air"
    if t in ("fumehood", "fume_hood", "tu hut", "tu_hut", "tủ hút"):
        return "fumehood"
    if t in ("bsc", "biological_safety_cabinet", "biosafety_cabinet"):
        return "bsc"
    if t in ("laf", "laf_bsc"):
        return "laf"
    return "cleanroom"


# ---------------------------------------------------------------------------
# Định dạng số / chữ
# ---------------------------------------------------------------------------
def dec(value: Any, is_en: bool, ndigits: Optional[int] = None) -> str:
    """Số thập phân theo ngôn ngữ: VN dấu phẩy, EN dấu chấm (AGENTS.md mục 6). Giữ nguyên chữ không phải số."""
    if value is None:
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        s = f"{value:.{ndigits}f}" if ndigits is not None else (str(int(value)) if float(value).is_integer() else str(value))
    else:
        s = str(value)
    return s.replace(",", ".") if is_en else s.replace(".", ",")


def thousands(value: Any) -> str:
    """Số nguyên có khoảng trắng ngăn cách hàng nghìn (vd 70 388). Giữ 0 là '0' (không ghi '-')."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{int(round(value)):,}".replace(",", " ")
    return "" if value is None else str(value)


def quoted(name: str) -> str:
    name = (name or "").strip().strip('"“”')
    return f"“{name}”" if name else ""


def date_of(text: str) -> str:
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text or "")
    return m.group(1) if m else ""


def _date_str(v) -> str:
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    return str(v or "").strip()


def scope_text(d: Dossier, is_en: bool) -> str:
    """Mục 3 'Nội dung yêu cầu giám định': form tủ (BSC, tủ hút, tủ cấy/clean bench) ghi 'Giám định tủ ...'."""
    s = (d.scope_en if is_en else d.scope_vn) or ""
    if norm_type(d.dossier_type) in ("bsc", "laf", "fumehood"):
        s = inspection_scope(s, is_en)
    return s


def inspection_date(d: Dossier, is_en: bool = False) -> str:
    """Ngày giám định dùng cho 'Từ ngày ...' / 'From ...'. Nhiều ngày (Dossier.inspection_dates) -> liệt kê
    'd1 và d2' / 'd1, d2, và d3' (người dùng chốt 07/11/2026, AGENTS mục 2)."""
    from ..Models.dossier import dates_phrase
    if len(getattr(d, "inspection_dates", None) or []) > 1:
        return dates_phrase(d.inspection_dates, is_en)
    return date_of(d.sample_time_vn) or date_of(d.sample_time_en) or _date_str(d.sample_date) or _date_str(d.cert_date)


def equipment_text(d: Dossier, is_en: bool, empty: str = PLACEHOLDER) -> str:
    lines = []
    for eq in d.equipments:
        name = (eq.name_en if is_en else eq.name_vn) or ""
        name = name.strip()
        if not name:
            continue
        s = name if name.startswith("-") else f"- {name}"
        s = s.rstrip(";. ") + ";"
        lines.append(s)
    if not lines:
        return empty
    lines[-1] = lines[-1].rstrip(";") + "."
    return "\n".join(lines)


def _parse_sig(txt: str):
    lines = [l.strip() for l in (txt or "").split("\n") if l.strip()]
    if not lines:
        return "", "", ""
    if len(lines) == 1:
        return lines[0], "", ""
    if len(lines) == 2:
        return lines[0], "", lines[1]
    return lines[0], lines[1], lines[-1]


def signature_context(d: Dossier, is_en: bool) -> Dict[str, str]:
    t1_l, _t2_l, n_l = _parse_sig(d.personnel.sign_gdv_en if is_en else d.personnel.sign_gdv_vn)
    if not n_l:  # chỉ có 1 dòng -> coi là họ tên
        n_l, t1_l = t1_l, ""
    t1_r, t2_r, n_r = _parse_sig(d.personnel.sign_leader_en if is_en else d.personnel.sign_leader_vn)
    if not n_r:
        n_r, t1_r = t1_r, ""
    return {
        "gdv_chuc_danh": t1_l or ("INSPECTOR" if is_en else "GIÁM ĐỊNH VIÊN"),
        "gdv_ten": n_l or ("Nguyen Xuan Nam" if is_en else "Nguyễn Xuân Nam"),
        "ld_chuc_danh_1": t1_r or ("FOR DIRECTOR" if is_en else "KT. GIÁM ĐỐC"),
        "ld_chuc_danh_2": t2_r if (t1_r or t2_r) else ("VICE DIRECTOR" if is_en else "PHÓ GIÁM ĐỐC"),
        "ld_ten": n_r or ("Le Dang Huyen" if is_en else "Lê Đăng Huyền"),
    }


def default_standards(d: Dossier, is_en: bool, client_name: str) -> str:
    if norm_type(d.dossier_type) == "cleanroom":
        if is_en:
            return ("- ISO 14644-1:2015 Cleanrooms and associated controlled environments - Part 1: Classification of air "
                    "cleanliness by particle concentration;\n- Technical specifications for temperature - humidity, "
                    f"differential pressure, supply airflow rate and air change rate provided by {client_name}.")
        return ("- ISO 14644-1:2015 Cleanrooms and associated controlled environments — Part 1: Classification of air "
                "cleanliness by particle concentration;\n- Mức quy định đối với nhiệt độ - độ ẩm, chênh áp, lưu lượng khí "
                f"cấp và số lần trao đổi khí được cung cấp bởi {client_name}.")
    return (d.basis_en if is_en else d.basis_vn) or PLACEHOLDER


# ---------------------------------------------------------------------------
# Context trang bìa (header + Table 0) — dùng cho mọi TBKQ
# ---------------------------------------------------------------------------
def cover_context(d: Dossier, is_en: bool) -> Dict[str, Any]:
    dtype = norm_type(d.dossier_type)
    client_name = normalize_company_name(d.client.name_en if is_en else d.client.name_vn)
    client_addr = d.client.address_en if is_en else d.client.address_vn
    loc_desc = normalize_company_name(d.location_desc_en if is_en else d.location_desc_vn)
    if not loc_desc:
        loc_desc = f"{client_name} - {client_addr}"
    reg_date = _date_str(getattr(d, "reg_date", "") or d.sample_date or d.cert_date)
    results_pointer = (getattr(d, "results_pointer_en", None) if is_en else getattr(d, "results_pointer_vn", None)) \
        or ("See the attached Appendix" if is_en else "Xem Phụ lục đính kèm")
    ctx = {
        "so_ct": d.cert_no_en if is_en else d.cert_no_vn,
        "ngay_ct": _date_str(d.cert_date),
        "ma_hs": d.code.strip(),
        "ngay_dk": reg_date,
        "kh_ten": client_name,
        "kh_dia_chi": client_addr,
        "noi_dung_yc": scope_text(d, is_en) or PLACEHOLDER,
        "ngay_gd": inspection_date(d, is_en),
        "dia_diem": loc_desc,
        # Quy trình: chỉ QUATEST 3 3030:2025 + mã N2.TC.* (người dùng chốt 02/11/2026)
        "quy_trinh": inspection_procedure(_default_procedure(d, dtype, is_en), bsc_class_of(d)),
        "thiet_bi": equipment_text(d, is_en),
        # Căn cứ: trích dẫn đầy đủ tên nguyên bản của tiêu chuẩn (người dùng chốt 01/11/2026)
        "can_cu": cite_standards((d.standards_en if is_en else d.standards_vn) or default_standards(d, is_en, client_name)),
        "ket_qua_gd": results_pointer,
        "is_en": is_en,
    }
    ctx.update(signature_context(d, is_en))
    return ctx


# ---------------------------------------------------------------------------
# Context chung BCKQ (BM3030.03)
# ---------------------------------------------------------------------------
def bckq_common_context(d: Dossier) -> Dict[str, Any]:
    reporter = d.personnel.reporter or ""
    tabs = "\t" * (6 if len(reporter) > 18 else 7)
    return {
        "so_bc": d.cert_no_vn.strip(),   # số báo cáo = số chứng thư: có /1, /2 khi hồ sơ nhiều báo cáo (AGENTS mục 7)
        "ma_hs": d.code.strip(),
        "kh_ten": d.client.name_vn,
        "kh_dia_chi": d.client.address_vn,
        "noi_dung_yc": scope_text(d, False) or PLACEHOLDER,
        "dia_diem": d.location_desc_vn or f"{d.client.name_vn} - {d.client.address_vn}",
        "tg_th": d.sample_time_vn or (f"Ngày {d.sample_date}" if d.sample_date else PLACEHOLDER),
        "thiet_bi": equipment_text(d, False),
        "nguoi_th": d.personnel.executor or PLACEHOLDER,
        "nguoi_bc": reporter or PLACEHOLDER,
        "tab_bc": tabs,
        # Ngày kết thúc hồ sơ thực tế & Ngày báo cáo = ngày hiện tại (ngày dựng), kể cả hồ sơ có KQTN và khi dựng lại
        # (người dùng chốt 09/11/2026, AGENTS mục 7) — khác ngày ban hành TBKQ (theo KQTN, mục 2)
        "ngay_kt": __import__("datetime").date.today().strftime("%d/%m/%Y"),
        "ngay_bc": __import__("datetime").date.today().strftime("%d/%m/%Y"),
        "ngay_ct": _date_str(d.cert_date),
        "tre_han": "[  ] có\t\t[X] không",
        "tg_tre_han": "Không",
        "nn_tre_han": "Không",
        "bat_thuong": "[  ] Có\t\t[X] không.",
    }


# ---------------------------------------------------------------------------
# Khối lưới ảnh (TagBlock) — AGENTS.md mục 3
# ---------------------------------------------------------------------------
def _process_photo(path: str, max_w: int = 1600) -> str:
    """Xoay theo EXIF + thu nhỏ (KHÔNG crop). Lỗi thì dùng ảnh gốc."""
    try:
        from PIL import Image, ImageOps
        im = Image.open(path)
        im = ImageOps.exif_transpose(im)
        w, h = im.size
        if w <= max_w and im.format in ("JPEG", "PNG"):
            return path
        if w > max_w:
            im = im.resize((max_w, int(h * max_w / w)), Image.Resampling.LANCZOS)
        out_dir = os.path.join(os.path.dirname(path), "_optimized")
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir, f"opt_{os.path.splitext(os.path.basename(path))[0]}.jpg")
        im.convert("RGB").save(out, format="JPEG", quality=90)
        return out
    except Exception as _err:
        print(f"[Warning] tag_context photo: {type(_err).__name__}: {_err}")
        return path


def _grid_items(items, is_en: bool):
    out = []
    for it in items or []:
        if isinstance(it, dict):
            path = it.get("path", "")
            cap = it.get("caption_en" if is_en else "caption_vn", "") or ""
        else:
            path, cap = str(it), ""
        if path and os.path.isfile(path):
            out.append((path, cap))
    return out


def _build_grid(doc, anchor_p, ctx, items, is_en: bool, cols: Optional[int], optimize: bool,
                single_width_cm: float, captions: bool = True, width_cm: Optional[float] = None):
    entries = _grid_items(items, is_en)
    if not captions:
        entries = [(path, "") for path, _ in entries]
    if not entries:
        return []
    n = len(entries)
    if cols is None:
        cols = 3 if n >= 5 else (2 if n >= 2 else 1)
    if cols == 3:
        width = PHOTO_WIDTH_3COL
    elif cols == 2:
        width = PHOTO_WIDTH_2COL
    else:
        width = Cm(single_width_cm)
    if width_cm and cols in (2, 3):
        # chỉnh cỡ ảnh trong khung cho phép để bỏ trang thừa / VN = EN (AGENTS mục 3: 2 cột 7–8 cm, 3 cột 4–5 cm)
        lo, hi = (7.0, 8.0) if cols == 2 else (4.0, 5.0)
        width = Cm(min(max(float(width_cm), lo), hi))
    rows = math.ceil(n / cols)
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for idx, (path, cap) in enumerate(entries):
        cell = tbl.rows[idx // cols].cells[idx % cols]
        p_img = cell.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(0)
        p_img.paragraph_format.line_spacing = 1.0
        run = p_img.add_run()
        run.add_picture(_process_photo(path) if optimize else path, width=width)  # chỉ đặt width, không crop
        for sr in run._r.iter(f"{{{a_ns}}}srcRect"):
            sr.attrib.clear()
        if cap:
            p_cap = cell.add_paragraph()
            set_paragraph_text(p_cap, cap, italic=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
    for row in tbl.rows:
        add_cant_split(row)
    set_table_autofit_window(tbl)
    el = tbl._tbl
    el.getparent().remove(el)  # doc.add_table() gắn ở cuối thân văn bản -> tách ra để đặt vào chỗ TAG
    return [el]


def photo_grid(items, is_en: bool, cols: Optional[int] = None, optimize: bool = False,
               single_width_cm: float = 13.0, captions: bool = True, width_cm: Optional[float] = None) -> TagBlock:
    """
    TAG khối lưới ảnh: 1 ảnh -> 1 cột 13 cm; 2-4 ảnh -> 2 cột 7,5 cm; >= 5 ảnh -> 3 cột 4,8 cm.
    Đặt {{ bang_anh }} riêng một đoạn trong template; không có ảnh -> đoạn bị xóa.
    `captions=False`: ảnh hiện trường không chú thích (người dùng chốt 06/11/2026, AGENTS mục 3);
    sơ đồ giữ chú thích.
    """
    return TagBlock(_build_grid, items, is_en, cols, optimize, single_width_cm, captions, width_cm)


INTRO_PREFIXES = ("Hình ảnh", "Sơ đồ", "Photos", "Photo of", "Layout", "Diagram")


def keep_intros_with_next(doc) -> int:
    """Đoạn dẫn 'Hình ảnh quá trình thực hiện:' / 'Sơ đồ các vị trí …:' (EN 'Photos…', 'Layout…') đứng ngay trước
    bảng ảnh/sơ đồ -> keepNext để không rớt lẻ cuối trang (AGENTS mục 3). Chỉ đặt cho đoạn dẫn, không cho bảng ảnh."""
    from docx.oxml.ns import qn as _qn
    n = 0
    body = doc.element.body
    for el in list(body):
        if not el.tag.endswith("}p"):
            continue
        txt = "".join(t.text or "" for t in el.iter(_qn("w:t"))).strip()
        if not txt.startswith(INTRO_PREFIXES) or len(txt) > 80:
            continue
        nxt = el.getnext()
        if nxt is None or not nxt.tag.endswith("}tbl"):
            continue
        pPr = el.find(_qn("w:pPr"))
        if pPr is None:
            pPr = el.makeelement(_qn("w:pPr"), {})
            el.insert(0, pPr)
        if pPr.find(_qn("w:keepNext")) is None:
            pPr.append(pPr.makeelement(_qn("w:keepNext"), {}))
            n += 1
    return n


# ---------------------------------------------------------------------------
# Hậu xử lý chung sau khi điền TAG
# ---------------------------------------------------------------------------
ASSESS_WORDS = ("phù hợp", "không phù hợp", "complied", "comply", "did not comply")


def format_assessments(doc):
    """'Đánh giá:'/'Assessment:' giữ cùng trang (keepNext); cụm đánh giá in đậm nghiêng (AGENTS.md mục 3)."""
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.startswith("Đánh giá:") or txt.startswith("Assessment:"):
            pPr = p._p.get_or_add_pPr()
            if pPr.find(f"{{{W_NS}}}keepNext") is None:
                pPr.append(parse_xml(f'<w:keepNext xmlns:w="{W_NS}"/>'))
        elif any(w in txt for w in ASSESS_WORDS) and len(txt) > 20:
            # câu kết luận: chỉ keepLines, KHÔNG keepNext (AGENTS mục 3) — keepNext kéo cả mục sau theo,
            # câu 'Đánh giá:' rớt sang trang mới (khí nén 133266, 07/11/2026)
            set_assessment_paragraph(p, txt, font_size=13, keep_next=False)


def merge_same_vertical(table, col: int, start_row: int, end_row: Optional[int] = None):
    """Gộp dọc các ô liên tiếp cùng nội dung ở cột `col` (vd tên khu vực lặp lại trên nhiều điểm đo)."""
    rows = table.rows
    end_row = len(rows) - 1 if end_row is None else end_row
    r = start_row
    while r <= end_row:
        try:
            txt = rows[r].cells[col].text.strip()
        except IndexError:
            r += 1
            continue
        s = r
        while r + 1 <= end_row:
            try:
                nxt = rows[r + 1].cells[col].text.strip()
            except IndexError:
                break
            if nxt != txt or not txt:
                break
            r += 1
        if r > s:
            keep = rows[s].cells[col]
            for rr in range(s + 1, r + 1):
                c = rows[rr].cells[col]
                for p in list(c.paragraphs):
                    for t in p._p.iter(f"{{{W_NS}}}t"):
                        t.text = ""
            merged = keep.merge(rows[r].cells[col])
            # chỉ giữ đoạn đầu (có chữ)
            ps = merged.paragraphs
            for p in ps[1:]:
                if not p.text.strip():
                    p._p.getparent().remove(p._p)
            tcPr = merged._tc.get_or_add_tcPr()
            old = tcPr.find(f"{{{W_NS}}}vAlign")
            if old is not None:
                tcPr.remove(old)
            tcPr.append(parse_xml(f'<w:vAlign xmlns:w="{W_NS}" w:val="center"/>'))
        r += 1


def fix_docpr_ids(doc):
    """Đánh số lại id của ảnh/hình (wp:docPr) cho duy nhất trong thân văn bản (tránh lỗi Bẫy 34 khi chèn nhiều ảnh)."""
    wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    seen = set()
    els = list(doc.element.body.iter(f"{{{wp}}}docPr"))
    next_id = max([int(e.get("id", "0")) for e in els if (e.get("id") or "").isdigit()] + [0]) + 1
    for e in els:
        i = e.get("id")
        if i in seen:
            e.set("id", str(next_id))
            next_id += 1
        seen.add(e.get("id"))


def _is_dash_bullet(part, pPr) -> bool:
    """Đoạn có numPr mà định dạng số là gạch ('-', '–') — gạch đầu dòng tự động của form."""
    num_id = pPr.find(f"{{{W_NS}}}numPr/{{{W_NS}}}numId")
    if num_id is None:
        return False
    try:
        numbering = part.numbering_part.element
    except (KeyError, NotImplementedError):
        return False
    nid = num_id.get(f"{{{W_NS}}}val")
    ilvl_el = pPr.find(f"{{{W_NS}}}numPr/{{{W_NS}}}ilvl")
    ilvl = ilvl_el.get(f"{{{W_NS}}}val") if ilvl_el is not None else "0"
    for num in numbering.findall(f"{{{W_NS}}}num"):
        if num.get(f"{{{W_NS}}}numId") != nid:
            continue
        aid = num.find(f"{{{W_NS}}}abstractNumId").get(f"{{{W_NS}}}val")
        for an in numbering.findall(f"{{{W_NS}}}abstractNum"):
            if an.get(f"{{{W_NS}}}abstractNumId") != aid:
                continue
            for lvl in an.findall(f"{{{W_NS}}}lvl"):
                if lvl.get(f"{{{W_NS}}}ilvl") == ilvl:
                    txt = lvl.find(f"{{{W_NS}}}lvlText")
                    return txt is not None and (txt.get(f"{{{W_NS}}}val") or "").strip() in ("-", "–")
    return False


def drop_auto_dash(table):
    """Ô Table 0 có gạch đầu dòng tự động (numPr) mà nội dung đã tự ghi '- ' (Căn cứ nhiều tiêu chuẩn)
    -> bỏ numPr và thụt treo để không ra '- - ' (E23, AGENTS mục 14); mọi dòng '- ' thẳng hàng.
    Ô Thiết bị của template TAG đã bỏ numPr nhưng còn hanging, nhiều thiết bị nằm chung 1 đoạn (ngắt dòng)
    -> dòng đầu lùi ra lề, các dòng sau thụt vào (132488 tủ hút, 02/11/2026): bỏ hanging kể cả khi không có numPr.
    Căn cứ 1 tiêu chuẩn (vd SEFA 1-2010) dùng gạch tự động nằm sát lề, lệch với gạch Thiết bị ngay trên
    -> ô trải hết dòng (Thiết bị, Căn cứ) có gạch tự động thì đổi thành '- ' viết bằng chữ."""
    seen = set()
    for row in table.rows:
        cells = []
        for cell in row.cells:
            if cell._tc not in seen:
                seen.add(cell._tc)
                cells.append(cell)
        for cell in cells:
            for p in cell.paragraphs:
                pPr = p._p.pPr
                if pPr is None:
                    continue
                if len(cells) == 1 and p.text.strip() and not p.text.lstrip().startswith("-") \
                        and _is_dash_bullet(table.part, pPr):
                    runs = [r for r in p.runs if r.text]
                    runs[0].text = "- " + runs[0].text.lstrip()
                if not p.text.lstrip().startswith("-"):
                    continue
                for num in pPr.findall(f"{{{W_NS}}}numPr"):
                    pPr.remove(num)
                for ind in pPr.findall(f"{{{W_NS}}}ind"):
                    for k in ("hanging", "firstLine", "hangingChars", "firstLineChars"):
                        ind.attrib.pop(f"{{{W_NS}}}{k}", None)


def align_table0_values(table) -> int:
    """Ô giá trị Table 0 (hàng 'nhãn | : | giá trị') thẳng hàng: thụt lề trái 0, bỏ hanging / firstLine.
    Template khí nén, phòng sạch để ô 'Thời gian giám định' thụt left 23 / hanging 136 (dòng đầu lùi −113 twips),
    các ô khác 0 hoặc −104 / −113 -> chữ lệch cột (137035, người dùng báo 09/11/2026). Không đụng ô nhãn, ô ':'
    và ô trải hết dòng (Thiết bị, Căn cứ — đã do drop_auto_dash xử lý)."""
    n = 0
    for row in table.rows:
        cells, seen = [], set()
        for cell in row.cells:
            if cell._tc not in seen:
                seen.add(cell._tc)
                cells.append(cell)
        if len(cells) != 3 or cells[1].text.strip() != ":":
            continue
        for p in cells[2].paragraphs:
            pPr = p._p.pPr
            if pPr is None:
                continue
            for ind in pPr.findall(f"{{{W_NS}}}ind"):
                for k in ("hanging", "firstLine", "hangingChars", "firstLineChars", "leftChars", "startChars"):
                    ind.attrib.pop(f"{{{W_NS}}}{k}", None)
                for k in ("left", "start"):
                    if ind.get(f"{{{W_NS}}}{k}") not in (None, "0"):
                        ind.set(f"{{{W_NS}}}{k}", "0")
                n += 1
    return n


def _exact_spacing(p, line_twips: int):
    pPr = p.find(f"{{{W_NS}}}pPr")
    if pPr is None:
        pPr = parse_xml(f'<w:pPr xmlns:w="{W_NS}"/>')
        p.insert(0, pPr)
    sp = pPr.find(f"{{{W_NS}}}spacing")
    if sp is None:
        sp = parse_xml(f'<w:spacing xmlns:w="{W_NS}"/>')
        pPr.append(sp)
    for k, v in (("before", "0"), ("after", "0"), ("line", str(line_twips)), ("lineRule", "exact")):
        sp.set(f"{{{W_NS}}}{k}", v)
    return pPr


def shrink_cover_break(doc, compact: bool = False):
    """Giữ trang bìa TBKQ trong 1 trang khi Table 0 dài (vd Căn cứ trích dẫn đủ tên tiêu chuẩn — 132488/1, 01/11/2026,
    trang bìa tràn làm trang 2 trống / chữ ký sang trang 2 và UpChungThu báo thiếu số trang phụ lục):
    - dòng trống ngăn giữa mục 'Kết quả giám định' và khối chữ ký -> cao 6pt (không đổi khoảng ký 3 cm);
    - đoạn trống chứa ngắt section ngay sau Table 0 -> cỡ 1pt;
    - compact=True (Căn cứ từ 2 tiêu chuẩn): khoảng cách trên/dưới đoạn trong các ô 3pt -> 1,5pt (trừ khối chữ ký);
      compact=2 (Căn cứ từ 4 mục, vd 132488/1 thêm CAG-003:2022 — 02/11/2026, bìa EN tràn): 3pt -> 0."""
    tbl = doc.tables[0]._tbl
    rows = tbl.findall(f"{{{W_NS}}}tr")
    if compact:
        new_sp = "0" if compact == 2 else "30"
        for tr in rows[:-2]:
            for sp in tr.iter(f"{{{W_NS}}}spacing"):
                for k in ("before", "after"):
                    if sp.get(f"{{{W_NS}}}{k}") == "60":
                        sp.set(f"{{{W_NS}}}{k}", new_sp)
    if len(rows) >= 2:
        spacer = rows[-2]
        if not "".join(t.text or "" for t in spacer.iter(f"{{{W_NS}}}t")).strip() \
                and spacer.find(f".//{{{W_NS}}}drawing") is None:
            for p in spacer.iter(f"{{{W_NS}}}p"):
                _exact_spacing(p, 120)
    nxt = tbl.getnext()
    while nxt is not None and nxt.tag != f"{{{W_NS}}}p":
        if nxt.tag == f"{{{W_NS}}}tbl":
            return
        nxt = nxt.getnext()
    if nxt is None or nxt.find(f".//{{{W_NS}}}sectPr") is None:
        return
    if "".join(t.text or "" for t in nxt.iter(f"{{{W_NS}}}t")).strip():
        return
    pPr = _exact_spacing(nxt, 20)
    targets =[pPr.find(f"{{{W_NS}}}rPr")] + nxt.findall(f"{{{W_NS}}}r/{{{W_NS}}}rPr")
    for rPr in targets:
        if rPr is None:
            continue
        for tag in ("sz", "szCs"):
            el = rPr.find(f"{{{W_NS}}}{tag}")
            if el is None:
                el = parse_xml(f'<w:{tag} xmlns:w="{W_NS}"/>')
                rPr.append(el)
            el.set(f"{{{W_NS}}}val", "2")


def render(doc, ctx: Dict[str, Any], label: str = "") -> set:
    """Điền TAG; in cảnh báo nếu thiếu dữ liệu."""
    missing = render_docx(doc, ctx, warn=False)
    fix_docpr_ids(doc)
    if "can_cu" in ctx and doc.tables:
        drop_auto_dash(doc.tables[0])
        align_table0_values(doc.tables[0])
        n_can_cu = str(ctx.get("can_cu") or "").count("\n") + 1
        shrink_cover_break(doc, compact=2 if n_can_cu >= 4 else n_can_cu >= 2)
    if missing:
        print(f"[TAG]{(' ' + label) if label else ''} thiếu dữ liệu (để trống): " + ", ".join(sorted(missing)))
    return missing


def ensure_result_tables(doc, min_height: int = 400, skip_first: bool = True):
    """
    Bảng kết quả sau khi điền: dòng tiêu đề lặp lại (tblHeader), không tách dòng qua trang (cantSplit),
    chiều cao dòng tối thiểu 0,7 cm (AGENTS.md mục 3). Không đổi viền/độ rộng của template.
    Bỏ qua Table 0 (trang bìa), bảng ghi chú ('Ghi chú'/'Note') và bảng ảnh.
    """
    from ..Styling.tables import add_tbl_header, set_table_row_height
    for i, tbl in enumerate(doc.tables):
        if skip_first and i == 0:
            continue
        if not tbl.rows:
            continue
        first = tbl.rows[0]
        first_txt = "".join(c.text for c in first.cells).strip()
        if first_txt.startswith("Ghi chú") or first_txt.startswith("Note") or tbl._tbl.find(f".//{{{W_NS}}}drawing") is not None:
            for row in tbl.rows:
                add_cant_split(row)
            continue
        add_tbl_header(first)
        for row in tbl.rows:
            add_cant_split(row)
            trPr = row._tr.get_or_add_trPr()
            h = trPr.find(f"{{{W_NS}}}trHeight")
            try:
                cur = int(h.get(f"{{{W_NS}}}val")) if h is not None else 0
            except (TypeError, ValueError):
                cur = 0
            if cur < min_height:
                set_table_row_height(row, height_twips=440)


def number_of(text: Any) -> Optional[float]:
    """Số đầu tiên trong chuỗi ('< 0,01' -> 0.01, '< 27 615 hạt/m3' -> 27615, '- 41,9' -> -41.9)."""
    if isinstance(text, (int, float)) and not isinstance(text, bool):
        return float(text)
    s = str(text or "").strip()
    if not s:
        return None
    # Class / Cấp không kèm toán tử so sánh số -> không phải số
    if re.search(r"^(?:iso\s*)?(?:class|cấp)\s*\d+", s, re.IGNORECASE) and not re.search(r"[<>≤≥=]", s):
        return None
    m = re.search(r"(-\s*)?\d[\d  .,]*\d|(?:\b|-)\d\b", s)
    if not m:
        return None
    raw = m.group(0).replace(" ", "").replace(" ", "")
    neg = raw.startswith("-")
    raw = raw.lstrip("-")
    # Định dạng dấu phẩy nghìn kiểu Anh: 10,200 hoặc 1,234,567 hoặc 10,200.5
    if re.search(r"^\d{1,3}(,\d{3})+(\.\d+)?$", raw):
        raw = raw.replace(",", "")
    # Định dạng dấu chấm nghìn kiểu Việt/Đức: 10.200,5
    elif re.search(r"^\d{1,3}(\.\d{3})+(,\d+)?$", raw):
        raw = raw.replace(".", "").replace(",", ".")
    else:
        # Số thập phân thông thường: 0,01 -> 0.01; 24,5 -> 24.5
        raw = raw.replace(",", ".")
    try:
        v = float(raw)
    except ValueError:
        return None
    return -v if neg else v


def requirement_text(text: Any, is_en: bool) -> str:
    """'< 1mg/m3' -> '< 1'; '< 27 615 hạt/m3' -> '< 27 615'; giữ dấu so sánh, bỏ đơn vị, đổi dấu thập phân."""
    s = str(text or "").strip()
    if not s:
        return ""
    m = re.match(r"^\s*([<>≤≥=]+)?\s*(-?\s*\d[\d  ]*(?:[.,]\d+)?)", s)
    if not m:
        return dec(s, is_en)
    op = (m.group(1) or "").strip()
    num = m.group(2).strip()
    v = number_of(num)
    if v is not None and v.is_integer() and abs(v) >= 1000:
        num = thousands(v)
    num = dec(num, is_en)
    return f"{op} {num}".strip()


def complies(values, limit_text: str, mode: str = "max") -> Optional[bool]:
    """So sánh giá trị đo với mức yêu cầu. mode='max': đo <= mức; 'min': đo >= mức. Không đọc được -> None."""
    lim = number_of(limit_text)
    if lim is None:
        return None
    strict = str(limit_text).strip().startswith(("<", ">")) and not str(limit_text).strip().startswith(("<=", ">=", "≤", "≥"))
    ok = True
    for v in values:
        x = number_of(v)
        if x is None:
            return None
        v_str = str(v).strip()
        less_than = v_str.startswith("<")
        greater_than = v_str.startswith(">")
        if mode == "max":
            if less_than:
                ok &= x <= lim
            elif greater_than:
                ok = False
            else:
                ok &= (x < lim) if strict else (x <= lim)
        else:
            if less_than:
                ok = False
            elif greater_than:
                ok &= x >= lim
            else:
                ok &= (x > lim) if strict else (x >= lim)
    return ok


def verdict(ok: Optional[bool], is_en: bool, explicit: Optional[str] = None) -> str:
    """Cụm đánh giá: 'phù hợp' / 'không phù hợp' / 'complied' / 'did not comply'; không xác định -> '………'."""
    if explicit:
        return explicit
    if ok is None:
        return PLACEHOLDER
    if is_en:
        return "complied" if ok else "did not comply"
    return "phù hợp" if ok else "không phù hợp"


def open_tag_template(path: str):
    """Mở template; báo lỗi rõ ràng nếu template không phải dạng TAG."""
    import docx as _docx
    from ..Tools.tag_engine import has_tags
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Không tìm thấy template: {path}")
    doc = _docx.Document(path)
    if not has_tags(doc):
        raise ValueError(
            f"Template {os.path.basename(path)} không có TAG {{{{ }}}}. "
            "Chạy: python -m Core.Tools.build_tag_templates để dựng lại template dạng TAG.")
    return doc


def save_outputs(doc, d: Dossier, folder: str, file_name: str, net_sub: str = "", label: str = "") -> str:
    """Lưu vào thư mục hồ sơ (ThucHien/ hoặc TBKQ/) và sao sang output_dir / ổ mạng nếu được cấu hình."""
    import shutil
    d.ensure_output_directories()
    target = os.path.join(folder, file_name)
    doc.save(target)
    if d.output_dir:
        try:
            shutil.copy2(target, os.path.join(d.output_dir, file_name))
        except Exception as _err:
            print(f"[Warning] {label} copy output_dir {type(_err).__name__}: {_err}")
    if d.network_dir and os.path.isdir(d.network_dir):
        try:
            net = os.path.join(d.network_dir, net_sub) if net_sub else d.network_dir
            os.makedirs(net, exist_ok=True)
            shutil.copy2(target, os.path.join(net, file_name))
        except Exception as _err:
            print(f"[Warning] {label} copy network {type(_err).__name__}: {_err}")
    return target


def within(value: Any, limit: Any) -> Optional[bool]:
    """
    Giá trị có đạt mức quy định không. Hiểu các dạng: '≤ 0,01', '< 1', '≥ 10', '> 5', '18 - 25', '(40 ÷ 60)'.
    Mức '-' / 'Không quy định' / rỗng -> None (không đánh giá).
    """
    lim = str(limit or "").strip()
    if not lim or lim in ("-", "–") or "quy định" in lim.lower() and "không" in lim.lower() or "not specified" in lim.lower():
        return None
    if lim.startswith(("≤", "<", "≥", ">")):
        return complies([value], lim, mode="max" if lim[0] in "≤<" else "min")
    lim2 = re.sub(r"(?<=\d)\s*[-–÷~]\s*(?=-?\d)", " ~ ", lim)
    nums = re.findall(r"-?\d+(?:[.,]\d+)?", lim2)
    if len(nums) >= 2:
        lo, hi = (float(n.replace(",", ".")) for n in nums[:2])
        x = number_of(value)
        return None if x is None else (min(lo, hi) <= x <= max(lo, hi))
    return complies([value], lim, mode="max")


def all_ok(results) -> Optional[bool]:
    """Gộp nhiều kết quả: có False -> False; toàn True -> True; còn None -> None."""
    vals = list(results)
    if any(v is False for v in vals):
        return False
    if vals and all(v is True for v in vals):
        return True
    return None


def find_tagged_tables(doc, name: str):
    """Các bảng có w:tblDescription = name (đặt trong template để code tìm lại sau khi điền)."""
    out = []
    for t in doc.tables:
        tblPr = t._tbl.tblPr
        d_el = tblPr.find(f"{{{W_NS}}}tblDescription") if tblPr is not None else None
        if d_el is not None and d_el.get(f"{{{W_NS}}}val") == name:
            out.append(t)
    return out


def clear_table_tags(doc):
    for t in doc.tables:
        tblPr = t._tbl.tblPr
        if tblPr is None:
            continue
        for d_el in tblPr.findall(f"{{{W_NS}}}tblDescription"):
            if (d_el.get(f"{{{W_NS}}}val") or "").startswith("tag:"):
                tblPr.remove(d_el)


def merge_by_groups(table, start_row: int, end_row: int, key_col: int = 0, cols=(0,)):
    """Gộp dọc các ô ở `cols` theo nhóm dòng liên tiếp có cùng nội dung ở cột key_col (vd cùng khu vực)."""
    rows = table.rows
    groups = []
    r = start_row
    while r <= end_row:
        key = rows[r].cells[key_col].text.strip()
        s = r
        while r + 1 <= end_row and rows[r + 1].cells[key_col].text.strip() == key and key:
            r += 1
        if r > s:
            groups.append((s, r))
        r += 1
    for s, e in groups:
        for c in cols:
            top = table.cell(s, c)
            for rr in range(s + 1, e + 1):
                for t in table.cell(rr, c)._tc.iter(f"{{{W_NS}}}t"):
                    t.text = ""
            merged = top.merge(table.cell(e, c))
            for p in merged.paragraphs[1:]:
                if not p.text.strip():
                    p._p.getparent().remove(p._p)
            tcPr = merged._tc.get_or_add_tcPr()
            old = tcPr.find(f"{{{W_NS}}}vAlign")
            if old is not None:
                tcPr.remove(old)
            tcPr.append(parse_xml(f'<w:vAlign xmlns:w="{W_NS}" w:val="center"/>'))


def clear_appendix_body(doc):
    """Xóa toàn bộ nội dung phần Phụ lục (sau đoạn ngắt section đầu tiên) — dùng khi phần Phụ lục do code dựng."""
    body = doc.element.body
    sect_p = None
    for el in body:
        if el.tag == f"{{{W_NS}}}p" and el.find(f".//{{{W_NS}}}sectPr") is not None:
            sect_p = el
            break
    if sect_p is None:
        return
    started = False
    for el in list(body):
        if el is sect_p:
            started = True
            continue
        if started and el.tag != f"{{{W_NS}}}sectPr":
            body.remove(el)


def render_cover(doc, d: Dossier, is_en: bool, clear_body: bool = True, extra: Optional[Dict[str, Any]] = None) -> set:
    """
    Điền trang bìa (header 'Số/Ngày' + Table 0) từ template TAG. Dùng cho BSC / LAF / tủ hút,
    nơi phần Phụ lục vẫn do code dựng. clear_body=True: xóa nội dung Phụ lục mẫu trước khi điền.
    """
    from ..Tools.tag_engine import has_tags
    if clear_body:
        clear_appendix_body(doc)
    if not has_tags(doc):
        raise ValueError("Template trang bìa không có TAG. Chạy: python -m Core.Tools.build_tag_templates")
    ctx = cover_context(d, is_en)
    if extra:
        ctx.update(extra)
    return render(doc, ctx, label=f"trang bìa {d.dossier_type} {'EN' if is_en else 'VN'}")


def next_year(date_text: str) -> str:
    """'22/10/2026' -> '22/10/2027' (ngày kiểm tra tiếp theo khuyến nghị)."""
    parts = (date_text or "").split("/")
    if len(parts) == 3:
        try:
            return f"{parts[0]}/{parts[1]}/{int(parts[2]) + 1}"
        except ValueError:
            pass
    return date_text


def _done(res: Dict[str, Any], key: str, default: bool = True) -> bool:
    """Chỉ tiêu `key` có thực hiện không: mặc định theo `default`; {"skip": True} -> không; dữ liệu rỗng/None -> theo default."""
    v = res.get(key)
    if isinstance(v, dict) and v.get("skip"):
        return False
    if default:
        return True
    return bool(v)


def seal_content_lines(results_data: Dict[str, Any], kind: str = "bsc") -> list:
    """
    Các dòng "Nội dung kiểm tra" trên tem tủ (E25, 30/10/2026) — CHỈ liệt kê chỉ tiêu đã thực hiện (AGENTS mục 8C).
    Mỗi dòng tối đa 3 chỉ tiêu, cách nhau bằng tab (tab stop 932/1759 twips của template), giữ bố cục tem cũ:
      BSC/LAF:  Inflow | Downflow | HEPA filter leak
                Lighting level | Noise level | Airflow pattern
                UVC intensity | Particle count          (chỉ khi có đo)
      Tủ hút:   Face velocity
                Lighting level | Noise level | Airflow pattern
    Ghi đè thủ công: results_data["seal_lines"] = ["-A\t-B", "-C"].
    """
    res = results_data or {}
    if res.get("seal_lines"):
        return list(res["seal_lines"])
    if kind == "fumehood":
        rows = [
            [("Face velocity", _done(res, "inflow"))],
            [("Lighting level", _done(res, "lux")), ("Noise level", _done(res, "noise")),
             ("Airflow pattern", _done(res, "airflow_pattern"))],
        ]
    else:
        rows = [
            [("Inflow", _done(res, "inflow")), ("Downflow", _done(res, "downflow")),
             ("HEPA filter leak", _done(res, "hepa"))],
            [("Lighting level", _done(res, "lux")), ("Noise level", _done(res, "noise")),
             ("Airflow pattern", _done(res, "airflow_pattern"))],
            [("UVC intensity", _done(res, "uvc", default=False)),
             ("Particle count", _done(res, "particle", default=False))],
        ]
    lines = []
    for row in rows:
        names = [f"-{name}" for name, ok in row if ok]
        if names:
            lines.append("\t".join(names))
    return lines


def seal_standard(results_data: Dict[str, Any], kind: str = "bsc") -> str:
    """
    Căn cứ in sau "Nội dung kiểm tra:" trên tem tủ.
      BSC/LAF: "EN 12469:2000"; có đo nồng độ bụi -> "EN 12469:2000;" + xuống dòng + "ISO 14644-1:2015" (người dùng chốt 30/10/2026).
      Tủ hút:  "ASHRAE 110, SEFA 1".
    Ghi đè: results_data["seal_standard"].
    """
    res = results_data or {}
    if res.get("seal_standard"):
        return str(res["seal_standard"])
    if kind == "fumehood":
        return "ASHRAE 110, SEFA 1"
    std = "EN 12469:2000"
    if _done(res, "particle", default=False):
        # Xuống dòng (w:br) để "ISO 14644-1:2015" không bị ngắt giữa dấu gạch (ô tem rộng ~5,2 cm, chữ 6pt)
        std += ";\nISO 14644-1:2015"
    return std


def seal_serial(cab: Dict[str, Any]) -> str:
    """Dòng SN/ID trên tem: "<serial>/ <mã thiết bị>" (vd "39809392/ VS0012"); không có mã thiết bị thì chỉ serial."""
    empty = ("", "N/A", "NA", "-", "NONE", "KHÔNG CÓ")
    serial = str(cab.get("serial", "") or "").strip()
    dev = str(cab.get("device_id", "") or "").strip()
    has_sn = serial.upper() not in empty
    has_dev = dev.upper() not in empty
    if has_sn and has_dev:
        return f"{serial}/ {dev}"
    if has_dev:            # vd tủ hút 132488: serial N/A -> chỉ ghi mã thiết bị
        return dev
    return serial


def build_tem_tu_doc(d: Dossier, template_path: str, out_name: str, so_tem: str = None,
                     ten_thiet_bi: str = "", label: str = "tem tủ", noi_dung_kt: Optional[list] = None) -> str:
    """Sinh tem tủ từ template TAG (tem tu.docx / tem tu - fumehood.docx) và lưu vào thư mục TBKQ."""
    doc = open_tag_template(template_path)
    cab = (d.results_data or {}).get("cabinet", {}) or {}
    ngay = _date_str(d.sample_date or d.cert_date)
    kind = "fumehood" if "fumehood" in os.path.basename(template_path).lower() else "bsc"
    if noi_dung_kt is None:
        noi_dung_kt = seal_content_lines(d.results_data, kind)
    ctx = {
        "noi_dung_kt": noi_dung_kt,
        "can_cu_tem": seal_standard(d.results_data, kind),
        "so_tem": so_tem or d.code,
        "ma_hs": d.code,
        "ten_thiet_bi": ten_thiet_bi,
        "serial": seal_serial(cab),
        "model": cab.get("model", ""),
        "nhan_hieu": cab.get("brand", ""),
        "ngay_kiem_tra": ngay,
        "ngay_kiem_tra_tiep": next_year(ngay),
    }
    render(doc, ctx, label=label)
    from ..Styling import strip_all_highlights
    strip_all_highlights(doc)
    target = save_outputs(doc, d, d.tbkq_dir, out_name, label=label)
    print(f"[{label}] Generated Validation Seal -> {target}")
    return target


SEALS_PER_PAGE = 3   # 3 tủ / trang A4 ngang, mỗi tủ 1 cột (2 tem xếp dọc)


def build_tem_tu_sheet(seal_files: list, out_path: str, per_page: int = SEALS_PER_PAGE) -> str:
    """
    Gom tem của nhiều tủ vào chung 1 file (người dùng chốt 30/10/2026): mỗi trang 3 tủ, tủ thứ k ở cột k
    (2 tem xếp dọc), đủ 3 tủ thì sang trang mới. `seal_files` = các file tem 1 tủ do build_tem_tu() sinh ra
    (BSC/LAF và tủ hút trộn được; mọi tem cùng kích thước bảng). File tem lẻ từng tủ vẫn giữ nguyên.
    Ảnh / liên kết (logo, email) của từng tem được gắn lại quan hệ (rId) vào file gom.
    """
    import copy
    import docx as _docx
    from docx.oxml.ns import qn
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    if not seal_files:
        return ""
    W_TBL, W_TR, W_TC, W_P = qn("w:tbl"), qn("w:tr"), qn("w:tc"), qn("w:p")
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    base = _docx.Document(seal_files[0])
    body = base.element.body
    tpl_tbl = copy.deepcopy([el for el in body if el.tag == W_TBL][0])
    # bảng khung trống làm mẫu cho mỗi trang
    for tr in tpl_tbl.findall(W_TR):
        for tc in tr.findall(W_TC):
            for ch in list(tc):
                if ch.tag != qn("w:tcPr"):
                    tc.remove(ch)
            tc.append(parse_xml(f'<w:p xmlns:w="{W_NS}"/>'))
    sect = body.find(qn("w:sectPr"))
    tail_p = [el for el in body if el.tag == W_P][-1]   # đoạn 1pt cuối (đã thu nhỏ trong template)
    for el in list(body):
        if el is not sect:
            body.remove(el)

    def _remap(el, src_doc):
        """Gắn lại rId (ảnh, liên kết) của phần tử chép từ src_doc sang file gom."""
        if src_doc is docs[0]:   # cùng file với `base` -> rId giữ nguyên
            return
        for node in el.iter():
            for attr, val in list(node.attrib.items()):
                if not attr.startswith("{%s}" % R_NS):
                    continue
                rel = src_doc.part.rels.get(val)
                if rel is None:
                    continue
                if rel.is_external:
                    new_id = base.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif rel.reltype == RT.IMAGE:
                    # chép ảnh thành part mới (tên part tự đặt, trùng nội dung thì dùng lại) — tránh trùng
                    # tên /word/media/imageN trong file zip
                    import io
                    new_id, _img = base.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                else:
                    continue
                node.set(attr, new_id)

    docs = [_docx.Document(f) for f in seal_files]   # đọc lại file 1 vì `base` đã bị xóa thân văn bản
    pages = [docs[i:i + per_page] for i in range(0, len(docs), per_page)]
    for pi, page_docs in enumerate(pages):
        if pi > 0:   # ngắt trang bằng đoạn 1pt (không đẩy thêm trang trống)
            br_p = copy.deepcopy(tail_p)
            r = parse_xml(f'<w:r xmlns:w="{W_NS}"><w:rPr><w:sz w:val="2"/></w:rPr><w:br w:type="page"/></w:r>')
            br_p.append(r)
            body.insert(len(body) - 1, br_p)
        tbl = copy.deepcopy(tpl_tbl)
        rows = tbl.findall(W_TR)
        for col, sdoc in enumerate(page_docs):
            src_rows = [el for el in sdoc.element.body if el.tag == W_TBL][0].findall(W_TR)
            for ri, tr in enumerate(rows):
                dst_tc = tr.findall(W_TC)[col]
                src_tc = src_rows[ri].findall(W_TC)[0]
                for ch in list(dst_tc):
                    if ch.tag != qn("w:tcPr"):
                        dst_tc.remove(ch)
                for ch in src_tc:
                    if ch.tag == qn("w:tcPr"):
                        continue
                    new = copy.deepcopy(ch)
                    _remap(new, sdoc)
                    dst_tc.append(new)
        body.insert(len(body) - 1, tbl)
    body.insert(len(body) - 1, copy.deepcopy(tail_p))
    # đánh lại id hình vẽ cho khỏi trùng (Word báo lỗi nội dung nếu trùng wp:docPr id)
    WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    for i, dp in enumerate(body.iter("{%s}docPr" % WP_NS), start=1):
        dp.set("id", str(i))
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    base.save(out_path)
    print(f"[tem tủ] Gom {len(seal_files)} tủ ({len(pages)} trang) -> {out_path}")
    return out_path


def appendix_total_field(doc) -> int:
    """Footer phụ lục TBKQ: trường SECTIONPAGES -> {= {NUMPAGES} - 1} (E46e, AGENTS mục 4).
    Mọi run mới (mã trường, NUMPAGES lồng, kết quả) chép định dạng chữ (rPr) của run cũ — Word lấy định dạng ký tự đầu
    mã trường cho kết quả, run thiếu cỡ chữ làm số tổng to hơn 'Trang x/' (132809, người dùng báo 09/11/2026)."""
    import copy as _copy
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _qn

    n = 0
    for sec in doc.sections:
        for ftr in (sec.footer, sec.first_page_footer, sec.even_page_footer):
            if ftr.is_linked_to_previous:
                continue
            for it in list(ftr._element.iter(_qn("w:instrText"))):
                if "SECTIONPAGES" not in (it.text or ""):
                    continue
                r_old = it.getparent()
                rpr = r_old.find(_qn("w:rPr"))
                if rpr is None:      # lấy định dạng của run fldChar 'begin' ngay trước (cùng trường)
                    prev = r_old.getprevious()
                    rpr = prev.find(_qn("w:rPr")) if prev is not None else None

                def run(child):
                    r = OxmlElement("w:r")
                    if rpr is not None:
                        r.append(_copy.deepcopy(rpr))
                    r.append(child)
                    return r

                def instr(text):
                    el = OxmlElement("w:instrText")
                    el.set(_qn("xml:space"), "preserve")
                    el.text = text
                    return run(el)

                def fld(kind):
                    el = OxmlElement("w:fldChar")
                    el.set(_qn("w:fldCharType"), kind)
                    return run(el)

                t = OxmlElement("w:t")
                t.text = "3"
                for r_new in [instr(" = "), fld("begin"), instr(" NUMPAGES "), fld("separate"), run(t), fld("end"),
                              instr(r" - 1 \* MERGEFORMAT ")]:
                    r_old.addprevious(r_new)
                r_old.getparent().remove(r_old)
                n += 1
    # run kết quả / fldChar của trường ngoài (separate, end) cũng thống nhất cỡ chữ với chữ 'Trang'
    return n


NOTE_PT = 11          # bảng 'Ghi chú:' / 'Note:' dưới bảng kết quả: 11pt in nghiêng (người dùng chốt 09/11/2026)


def is_note_table(tbl_el) -> bool:
    from docx.oxml.ns import qn as _qn
    txt = "".join(t.text or "" for t in tbl_el.iter(_qn("w:t"))).strip()
    return txt.startswith(("Ghi chú", "Note"))


def fix_note_label(p) -> bool:
    """Nhãn 'Ghi chú:' / 'Note:' gạch chân đủ cả nhãn, dấu ':' không gạch (template khí nén VN từng gạch
    'Ghi ' mà sót 'chú:' — người dùng báo 09/11/2026). Gộp các run của nhãn thành 2 run: nhãn (u=single) + ':'."""
    import copy
    from docx.oxml.ns import qn as _qn
    runs = [r for r in p.findall(_qn("w:r"))]
    full = "".join(t.text or "" for r in runs for t in r.findall(_qn("w:t")))
    label = next((lb for lb in ("Ghi chú", "Note") if full.lstrip().startswith(lb)), None)
    if not label or ":" not in full[:len(label) + 4]:
        return False
    used, acc = [], ""
    for r in runs:
        if any(ch.tag not in (_qn("w:rPr"), _qn("w:t")) for ch in r):
            return False            # run nhãn có tab / ký hiệu lạ -> không đụng
        used.append(r)
        acc += "".join(t.text or "" for t in r.findall(_qn("w:t")))
        if ":" in acc:
            break
    head, rest = acc.split(":", 1)
    if head.strip() != label:
        return False
    if not any((r.find(_qn("w:rPr") + "/" + _qn("w:u")) is not None
                and r.find(_qn("w:rPr") + "/" + _qn("w:u")).get(_qn("w:val")) not in ("none", None)) for r in used):
        return False                # form không gạch chân nhãn (BSC, tủ hút…) -> giữ nguyên
    first = used[0]
    ts = first.findall(_qn("w:t")) or [None]
    if ts[0] is None:
        return False
    for t in ts[1:]:
        first.remove(t)
    ts[0].text = label
    rpr = first.find(_qn("w:rPr"))
    if rpr is None:
        from docx.oxml import OxmlElement
        rpr = OxmlElement("w:rPr")
        first.insert(0, rpr)
    u = rpr.find(_qn("w:u"))
    if u is None:
        from docx.oxml import OxmlElement
        u = OxmlElement("w:u")
        rpr.append(u)
    u.set(_qn("w:val"), "single")
    colon = copy.deepcopy(first)
    cu = colon.find(_qn("w:rPr")).find(_qn("w:u"))
    colon.find(_qn("w:rPr")).remove(cu)
    colon.find(_qn("w:t")).text = ":"
    first.addnext(colon)
    last = used[-1]
    for r in used[1:]:
        if r is last and rest:
            lt = r.findall(_qn("w:t"))
            for t in lt[1:]:
                r.remove(t)
            lt[0].text = rest
            lt[0].set(_qn("xml:space"), "preserve")
        else:
            p.remove(r)
    if last is first and rest:
        tail = copy.deepcopy(colon)
        tail.find(_qn("w:t")).text = rest
        tail.find(_qn("w:t")).set(_qn("xml:space"), "preserve")
        colon.addnext(tail)
    return True


def superscript_note_markers(doc) -> int:
    """Ký hiệu ghi chú '(1)' … '(9)' đứng đầu dòng / đầu ô của bảng Ghi chú / Note -> chỉ số trên (AGENTS mục 3).
    Template phòng sạch (TBKQ VN/EN, BCKQ) để '(1) …' chung một run chữ thường — người dùng báo 16/11/2026 (067460):
    tách run thành '(n)' (superscript, giữ font / nghiêng) + phần còn lại (' …', câu ghi chú). Dùng cho TBKQ
    (trong `style_note_tables`) và BCKQ (`BCKQBuilder`, không đổi cỡ chữ). Trả số ký hiệu đã sửa."""
    import copy
    import re as _re
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _qn

    n = 0
    for tbl in doc.element.body.iter(_qn("w:tbl")):
        if not is_note_table(tbl):
            continue
        for p in tbl.iter(_qn("w:p")):
            runs = [r for r in p.findall(_qn("w:r")) if "".join(t.text or "" for t in r.findall(_qn("w:t"))).strip()]
            if not runs:
                continue
            r = runs[0]
            ts = r.findall(_qn("w:t"))
            m = _re.match(r"^(\s*)(\(\d\))(.*)$", ts[0].text or "", _re.S) if len(ts) == 1 else None
            if not m:
                continue
            rpr = r.find(_qn("w:rPr"))
            va = rpr.find(_qn("w:vertAlign")) if rpr is not None else None
            if va is not None and va.get(_qn("w:val")) == "superscript":
                continue
            if m.group(3):
                rest = copy.deepcopy(r)
                rest.find(_qn("w:t")).text = m.group(3)
                rest.find(_qn("w:t")).set(_qn("xml:space"), "preserve")
                r.addnext(rest)
            ts[0].text = m.group(1) + m.group(2)
            ts[0].set(_qn("xml:space"), "preserve")
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            if va is None:
                va = OxmlElement("w:vertAlign")
                rpr.append(va)
            va.set(_qn("w:val"), "superscript")
            n += 1
    return n


def style_note_tables(doc, size_pt: int = NOTE_PT, italic: bool = True) -> int:
    """Toàn bộ bảng Ghi chú / Note (nhãn, câu công nhận, ghi chú ⁽ⁿ⁾) -> `size_pt` + in nghiêng (AGENTS mục 3).
    Giữ chỉ số trên, gạch chân nhãn. Gọi ở bước cuối dựng TBKQ (handler) và sau mọi hậu xử lý thêm ghi chú (script)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _qn

    def _set(rpr):
        for tag, val in (("w:sz", str(size_pt * 2)), ("w:szCs", str(size_pt * 2))):
            el = rpr.find(_qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rpr.append(el)
            el.set(_qn("w:val"), val)
        if italic:
            for tag in ("w:i", "w:iCs"):
                el = rpr.find(_qn(tag))
                if el is None:
                    rpr.append(OxmlElement(tag))
                else:
                    el.attrib.pop(_qn("w:val"), None)

    superscript_note_markers(doc)       # '(1) …' của template -> '(1)' chỉ số trên
    n = 0
    for tbl in doc.element.body.iter(_qn("w:tbl")):
        if not is_note_table(tbl):
            continue
        n += 1
        for p in tbl.iter(_qn("w:p")):
            fix_note_label(p)
            ppr = p.find(_qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                p.insert(0, ppr)
            prpr = ppr.find(_qn("w:rPr"))
            if prpr is None:
                prpr = OxmlElement("w:rPr")
                ppr.append(prpr)
            _set(prpr)
            for r in p.iter(_qn("w:r")):
                rpr = r.find(_qn("w:rPr"))
                if rpr is None:
                    rpr = OxmlElement("w:rPr")
                    r.insert(0, rpr)
                _set(rpr)
    return n
