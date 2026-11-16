# -*- coding: utf-8 -*-
"""
Internal Inspection Report Builder (BCKQ - BM3030.03) for QUATEST 3.

Từ 26/10/2026: sinh BCKQ từ template dạng TAG (Core/Templates/BCKQ - <loại>.docx, danh mục TAG: TAGS.md).
populate_bckq_table0() giữ lại để các script cũ (generate_all_reports.py) và bộ dựng template dùng.
Implements:
- Standard BM3030.03 form (2 pages, NO ending line) per AGENTS.md Section 7.
- Supports both Cleanroom/LAF/BSC Form (4-column Table 1) and Compressed Air Form (3-column Table 1).
- Automatic Reporter calculation:
    Toàn / Nghĩa -> <Tên> + Minh
    Other (Phát...) -> <Tên> + Minh + Nam
- AutoFit to Window for Photo and Layout tables (AGENTS.md Section 3).
- Borderless Notes (font size 10).
- Automatic file dispatch to Input/<profile>/TBKQ, Output/, and network drive.
"""

import copy
import os
import re
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from ..Models.dossier import Dossier, inspection_scope, inspection_procedure
from ..Styling import (
    W_NS,
    set_cell_content,
    format_bckq_table0,
    strip_all_highlights,
    strip_image_crops,
)
from ..Tools.tag_engine import TagBlock
from .tag_context import (
    PLACEHOLDER,
    bckq_common_context,
    equipment_text,
    photo_grid,
    quoted,
    open_tag_template,
    save_outputs,
    render,
)


def populate_bckq_table0(t0, d, is_kn=False):
    """Populate Table 0 dynamically by row labels across all BM3030.03 form variants."""
    set_cell_content(t0.rows[0].cells[4], d.code, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=13, valign="top", space_before=3, space_after=3)
    loc_str = d.location_desc_vn or f"{d.client.name_vn} - {d.client.address_vn}"
    time_str = d.sample_time_vn or f"Ngày {d.sample_date}"
    eq_lines = [f"- {eq.name_vn};" for eq in d.equipments]
    if eq_lines:
        eq_lines[-1] = eq_lines[-1].rstrip(";") + "."
    eq_text = "\n".join(eq_lines) if eq_lines else "………"

    for r_idx, r in enumerate(t0.rows):
        label = r.cells[0].text.strip()
        if "Đơn vị yêu cầu" in label:
            set_cell_content(r.cells[3], d.client.name_vn, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13, valign="top", space_before=3, space_after=3)
        elif "Địa chỉ" in label:
            set_cell_content(r.cells[3], d.client.address_vn, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13, valign="top", space_before=3, space_after=3)
        elif "Nội dung yêu cầu" in label:
            scope_str = d.scope_vn or ("Giám định tủ hút khí độc" if "fume" in d.dossier_type else ("Giám định tủ an toàn sinh học cấp 2" if "bsc" in d.dossier_type else f"Quan trắc và đánh giá chất lượng {d.dossier_type.upper()}"))
            if is_cabinet(d):
                scope_str = inspection_scope(scope_str)
            set_cell_content(r.cells[3], scope_str, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13, valign="top", space_before=3, space_after=3)
        elif "Địa điểm thực hiện" in label:
            set_cell_content(r.cells[3], loc_str, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13, valign="top", space_before=3, space_after=3)
        elif "Thời gian thực hiện" in label:
            set_cell_content(r.cells[3], time_str, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13, valign="top", space_before=3, space_after=3)
        elif "Thiết bị sử dụng" in label:
            if r_idx + 1 < len(t0.rows) and "Người thực hiện" not in t0.rows[r_idx + 1].cells[0].text:
                eq_cell = t0.rows[r_idx + 1].cells[0]
                set_cell_content(eq_cell, eq_text, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13, valign="top", space_before=3, space_after=3)
                for p in eq_cell.paragraphs:
                    if p._p.pPr is not None:
                        numPr = p._p.pPr.find(f"{{{W_NS}}}numPr")
                        if numPr is not None:
                            p._p.pPr.remove(numPr)
            else:
                eq_col = 0 if is_kn else 3
                set_cell_content(r.cells[eq_col], eq_text, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13, valign="top", space_before=3, space_after=3)
                for p in r.cells[eq_col].paragraphs:
                    if p._p.pPr is not None:
                        numPr = p._p.pPr.find(f"{{{W_NS}}}numPr")
                        if numPr is not None:
                            p._p.pPr.remove(numPr)
        elif "Người thực hiện" in label:
            set_cell_content(r.cells[3], d.personnel.executor, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13, valign="top", space_before=3, space_after=3)

    format_bckq_table0(t0)


# ---------------------------------------------------------------------------
# Đánh giá BCKQ tủ (BSC/LAF, tủ cấy, tủ hút): 3 phần — người dùng chốt 30/10/2026
#   + Chỉ tiêu phù hợp / + Chỉ tiêu không phù hợp / + Chỉ tiêu không đánh giá
# Kết luận lấy từ chính các dòng bảng kết quả TBKQ (BSCHandler.result_rows) để BCKQ khớp TBKQ (AGENTS mục 7).
# ---------------------------------------------------------------------------
CABINET_TYPES = ("bsc", "biological_safety_cabinet", "biosafety_cabinet", "laf", "laf_bsc",
                 "fumehood", "fume_hood", "tu hut", "tu_hut", "tủ hút")
FUMEHOOD_TYPES = ("fumehood", "fume_hood", "tu hut", "tu_hut", "tủ hút")
ASSESS_LABELS = (("phu_hop", "Chỉ tiêu phù hợp"), ("khong_phu_hop", "Chỉ tiêu không phù hợp"),
                 ("khong_danh_gia", "Chỉ tiêu không đánh giá"))
_NOT_SPEC = ("không quy định", "not specified", "")
_MARK_RE = re.compile(r"\s*(⁽[⁰¹²³⁴⁵⁶⁷⁸⁹]⁾|\(\*{1,2}\))")


def is_cabinet(d: Dossier) -> bool:
    return d.dossier_type.lower() in CABINET_TYPES


def cabinet_title(d: Dossier) -> str:
    """'Tủ cấy ESCO AVC-4D8 (TBKN46-01)' — bỏ các phần N/A."""
    rd = d.results_data or {}
    cab = rd.get("cabinet") or (rd.get("cabinets") or [{}])[0] or {}
    default = "Tủ hút khí độc" if d.dossier_type.lower() in FUMEHOOD_TYPES else "Tủ an toàn sinh học cấp 2"
    ok = lambda v: v and str(v).strip().upper() not in ("N/A", "NA", "-")
    parts = [str(cab.get("cabinet_name") or default).strip().strip('"“”')]
    parts += [str(cab[k]).strip() for k in ("brand", "model") if ok(cab.get(k))]
    title = " ".join(parts)
    tag = cab.get("device_id") if ok(cab.get("device_id")) else (cab.get("serial") if ok(cab.get("serial")) else "")
    return f"{title} ({tag})" if tag else title


def _verdict(text) -> str:
    t = (text or "").strip().lower()
    if t in ("đạt", "pass", "phù hợp", "comply"):
        return "phu_hop"
    if t in ("không đạt", "fail", "không phù hợp", "did not comply"):
        return "khong_phu_hop"
    return "khong_danh_gia"


def _lower_first(s: str) -> str:
    return s[:1].lower() + s[1:] if s else s


def _fumehood_rows(d: Dossier) -> list:
    """Dòng kết quả tủ hút giống FumeHoodHandler.build_tbkq (bản VN, không có dòng ảnh sơ đồ cửa)."""
    rd = d.results_data or {}
    has_ass = getattr(d, "has_assessment", True)
    ok_txt = "Đạt" if has_ass else ""
    rows = []
    ops = rd.get("openings") or []
    for i, op in enumerate(ops):
        name = "Tốc độ trung bình dòng khí qua cửa làm việc"
        if len(ops) > 1:
            pct = f" ({op['opening_pct']}%)" if op.get("opening_pct") else ""
            name += f" - Khoảng mở cửa: {op.get('height_mm', '')} mm{pct}"
        try:
            v = f"{float(op.get('avg')):.2f}".replace(".", ",")
        except (TypeError, ValueError):
            v = str(op.get("avg", "")).replace(".", ",")
        rows.append(("1" if i == 0 else "", name, f"{v} m/s", op.get("limit_vn", "≥ 0,5 m/s (*)"),
                     op.get("assessment_vn", ok_txt) if has_ass else "", False, False))
    mk = "(**)" if has_ass else "(*)"
    lux = rd.get("lux") or {}
    if lux and not lux.get("skip"):
        lux_v = lux.get("avg_rounded", int(round(lux.get("avg", 0) or 0)))
        rows.append(("2", "Cường độ sáng trung bình bên trong tủ", f"{lux_v} lux", f"≥ 861 lux {mk}",
                     lux.get("assessment_vn", ok_txt) if has_ass else "", False, False))
    noise = rd.get("noise") or {}
    rows.append(("3", "Đo độ ồn", "", "", "", True, False))
    rows.append(("", "Độ ồn nền", f"{str(noise.get('ambient', '-')).replace('.', ',')} dBA", "Không quy định", "", False, False))
    op_n = noise.get("corrected_dba", noise.get("operating", "-"))
    rows.append(("", "Độ ồn khi vận hành", f"{str(op_n).replace('.', ',')} dBA", f"≤ 85,0 dBA {mk}",
                 noise.get("assessment_vn", ok_txt) if has_ass else "", False, False))
    if not (rd.get("airflow_pattern") or {}).get("skip"):
        rows.append(("4", "Kiểm tra hình thái dòng khí", "", "", "", True, False))
        rows.append(("", "Bên trong không gian tủ", "", f"Theo SEFA 1-2010 {mk}", ok_txt, False, False))
        rows.append(("", "Khu vực dọc theo chu vi phía bên ngoài của vùng mở cửa tủ", "", f"Theo SEFA 1-2010 {mk}", ok_txt, False, False))
    return rows


def cabinet_assessment(d: Dossier) -> dict:
    """{'tieu_de', 'phu_hop': [...], 'khong_phu_hop': [...], 'khong_danh_gia': [...]} cho 1 tủ.

    Ghi đè bằng tay: d.assessment['phu_hop_vn' / 'khong_phu_hop_vn' / 'khong_danh_gia_vn'] (chuỗi hoặc danh sách).
    """
    out = {"tieu_de": cabinet_title(d), "so": d.cert_no_vn, "phu_hop": [], "khong_phu_hop": [], "khong_danh_gia": []}
    manual = {k: (d.assessment or {}).get(f"{k}_vn") for k, _ in ASSESS_LABELS}
    if any(v is not None for v in manual.values()):
        for k, v in manual.items():
            out[k] = [v] if isinstance(v, str) else list(v or [])
        return out

    is_hood = d.dossier_type.lower() in FUMEHOOD_TYPES
    has_ass = getattr(d, "has_assessment", True) if is_hood else True
    if is_hood:
        rows = _fumehood_rows(d)
        stds = {"(*)": "tài liệu kỹ thuật của nhà sản xuất" if has_ass else "SEFA 1-2010", "(**)": "SEFA 1-2010"}
    else:
        from .handlers.bsc import BSCHandler
        h = BSCHandler(d)
        rows = h.result_rows(False)
        stds = {"⁽¹⁾": h.basis_std(), "⁽²⁾": "ISO 14644-1:2015"}
        if h.nsx_mark():
            brand = ((d.results_data or {}).get('cabinet') or {}).get('brand', '')
            stds[h.nsx_mark()] = f'công bố của nhà sản xuất {brand}'.strip()
        if h.airflow_mark():
            stds[h.airflow_mark()] = d.results_data["airflow_pattern"]["basis_vn"]
    lim_word = "mức quy định" if has_ass else "mức khuyến nghị"

    def clean(lim: str):
        found = [m.strip() for m in _MARK_RE.findall(lim or "")]
        txt = " ".join(x.strip().lstrip("-").strip() for x in _MARK_RE.sub("", lim or "").splitlines() if x.strip())
        std = next((stds[m] for m in found if m in stds), "")
        return txt, std

    def detail(res: str, lim: str) -> str:
        txt, std = clean(lim)
        if txt.lower() in _NOT_SPEC:
            return f"{res} (không có mức quy định)" if res else "(không có mức quy định)"
        ref = f"{lim_word} {txt}" + (f", {std}" if std else "")
        return f"{res} ({ref})" if res else f"({ref})"

    # Gom dòng thành nhóm: dòng có STT là chỉ tiêu; dòng cha (is_parent) kéo theo các dòng con
    groups = []
    for r in rows:
        stt, param, res_v, lim_v, ass_v, is_parent, is_header = r[:7]
        if is_header:
            continue
        if is_parent:
            groups.append([param, []])
        elif stt or not groups or groups[-1][0] is None:
            groups.append([None, [(param, res_v, lim_v, ass_v)]])
        else:
            groups[-1][1].append((param, res_v, lim_v, ass_v))

    for parent, kids in groups:
        if parent is None:
            name, res_v, lim_v, ass_v = kids[0]
            if clean(lim_v)[0].lower() in _NOT_SPEC and _verdict(ass_v) != "khong_phu_hop":
                key = "khong_danh_gia"     # không có mức quy định thì không đánh giá (AGENTS mục 3)
            else:
                key = _verdict(ass_v) if has_ass else "khong_danh_gia"
            out[key].append(f"{name} {detail(res_v, lim_v)}")
            continue
        # dòng con có mức (bỏ "Độ ồn nền — Không quy định"); không dòng nào có mức thì lấy hết
        real = [k for k in kids if clean(k[2])[0].lower() not in _NOT_SPEC] or kids
        pl = parent.lower()
        if "độ ồn" in pl or "rò rỉ" in pl:
            for name, res_v, lim_v, ass_v in real:
                nm = name if "độ ồn" in pl else f"Mức rò rỉ {_lower_first(name)}"
                key = _verdict(ass_v) if has_ass else "khong_danh_gia"
                out[key].append(f"{nm} {detail(res_v, lim_v)}")
            continue
        keys = {(_verdict(k[3]) if has_ass else "khong_danh_gia") for k in real}
        if len(keys) == 1:   # cả nhóm cùng kết luận -> 1 dòng (hình thái dòng khí, nồng độ hạt bụi)
            key = keys.pop()
            _, std = clean(real[0][2])
            cls = ((d.results_data or {}).get("particle") or {}).get("clean_class", "") if "bụi" in pl else ""
            cls = re.sub(r"(?i)^iso\s*class\s*", "ISO ", cls)
            ref = ", ".join(x for x in (f"{lim_word} cấp {cls}" if cls else "", std) if x)
            name = re.sub(r"^Kiểm tra\s+", "", parent)
            out[key].append(name[:1].upper() + name[1:] + (f" ({ref})" if ref else ""))
        else:                # kết luận khác nhau -> liệt kê từng dòng con
            for name, res_v, lim_v, ass_v in real:
                out[_verdict(ass_v)].append(f"{parent} - {_lower_first(name.splitlines()[0])} {detail(res_v, lim_v)}")
    return out


def _clone_p(anchor_p, text_parts, left_cm: float = 0.0, hanging_cm: float = 0.0, keep_next: bool = True):
    """Đoạn mới cùng định dạng anchor_p; text_parts = [(chữ, in đậm)]."""
    p = copy.deepcopy(anchor_p)
    for ch in list(p):
        if ch.tag != qn("w:pPr"):
            p.remove(ch)
    ppr = p.find(qn("w:pPr"))
    if ppr is not None:
        for tag in ("w:numPr", "w:ind"):
            for x in ppr.findall(qn(tag)):
                ppr.remove(x)
    para = Paragraph(p, None)
    pf = para.paragraph_format
    pf.left_indent = Cm(left_cm + hanging_cm)
    pf.first_line_indent = Cm(-hanging_cm) if hanging_cm else Cm(0)
    pf.keep_with_next = keep_next
    for txt, bold in text_parts:
        r = para.add_run(txt)
        r.bold = bold
    return p


def assessment_block(doc, anchor_p, ctx, items):
    """TagBlock {{ danh_gia }}: '- Đánh giá:' rồi 3 phần cho mỗi tủ (nhiều tủ thì có dòng tên tủ + số)."""
    els = [_clone_p(anchor_p, [("- Đánh giá:", False)])]
    multi = len(items) > 1
    for i, a in enumerate(items):
        base = 0.0
        if multi:
            head = f"{i + 1}) {a['tieu_de']}" + (f" - số {a['so']}" if a.get("so") else "") + ":"
            els.append(_clone_p(anchor_p, [(head, True)], left_cm=0.3))
            base = 0.6
        for j, (key, label) in enumerate(ASSESS_LABELS):
            vals = [v.strip().rstrip(";.") for v in a.get(key) or [] if str(v).strip()]
            vals = [v if v[1:2].isupper() else _lower_first(v) for v in vals]   # "độ ồn ...", giữ "UVC ..."
            body = ("; ".join(vals) + ".") if vals else "Không."
            last = (i == len(items) - 1) and j == len(ASSESS_LABELS) - 1
            els.append(_clone_p(anchor_p, [(f"+ {label}: ", True), (body, False)],
                                left_cm=base, hanging_cm=0.4, keep_next=not last))
    return els


def _split_br_paragraphs(cell):
    """Tách đoạn có ngắt dòng (w:br) thành nhiều đoạn — đoạn căn đều có ngắt dòng làm Word dãn chữ ("Tủ ... cấy")."""
    for p in list(cell.paragraphs):
        if p._p.find(f".//{{{W_NS}}}br") is None:
            continue
        runs = p._p.findall(qn("w:r"))
        rpr = runs[0].find(qn("w:rPr")) if runs else None
        lines = p.text.split("\n")
        for line in lines:
            q = copy.deepcopy(p._p)
            for ch in list(q):
                if ch.tag != qn("w:pPr"):
                    q.remove(ch)
            para = Paragraph(q, p._parent)
            r = para.add_run(line)
            if rpr is not None:
                r._r.insert(0, copy.deepcopy(rpr))
            p._p.addprevious(q)
        p._p.getparent().remove(p._p)


def tidy_cabinet_table(table):
    """Bảng 'Thông tin thiết bị | Chỉ tiêu thực hiện': căn trái, mỗi dòng 1 đoạn, chỉ tiêu đánh số 1, 2, 3..."""
    for row in table.rows[1:]:
        cells = row.cells
        for c_idx, cell in enumerate(cells[:2]):
            _split_br_paragraphs(cell)
            paras = [p for p in cell.paragraphs if p.text.strip()]
            for n, p in enumerate(paras, start=1):
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if c_idx == 1:
                    ppr = p._p.pPr
                    if ppr is not None:
                        for x in ppr.findall(qn("w:numPr")):
                            ppr.remove(x)
                    txt = re.sub(r"^\s*(\d+[.)]|-)\s*", "", p.text).strip()
                    for r in p.runs[1:]:
                        r._r.getparent().remove(r._r)
                    p.runs[0].text = f"{n}. {txt}"
                    p.paragraph_format.left_indent = Cm(0.6)
                    p.paragraph_format.first_line_indent = Cm(-0.6)


class BCKQBuilder:
    """Builder for BM3030.03 internal inspection reports (template dạng TAG)."""

    def __init__(self, dossier: Dossier, template_path: str = None):
        self.dossier = dossier
        templates_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Templates")
        if not template_path:
            dtype = dossier.dossier_type.lower()
            if dtype in ("compressed_air", "khi nen"):
                template_path = os.path.join(templates_dir, "BCKQ - khi nen.docx")
            elif dtype in ("fumehood", "fume_hood", "tu hut", "tu_hut", "tủ hút"):
                template_path = os.path.join(templates_dir, "BCKQ - fumehood.docx")
            elif dtype in ("bsc", "biological_safety_cabinet", "biosafety_cabinet", "laf", "laf_bsc"):
                template_path = os.path.join(templates_dir, "BCKQ - bsc.docx")
            else:
                template_path = os.path.join(templates_dir, "BCKQ - phong sach.docx")
        self.template_path = template_path

    # ------------------------------------------------------------------
    @staticmethod
    def _cabinet_info(cab: dict, default_name: str) -> str:
        name = cab.get("cabinet_name", default_name).strip().strip('"')
        return (f'"{name}"\n'
                f"- Hãng\t: {cab.get('brand', '')}\n"
                f"- Model\t: {cab.get('model', '')}\n"
                f"- Serial No.\t: {cab.get('serial', '')}\n"
                f"- Mã quản lý\t: {cab.get('device_id', 'N/A')}")

    @staticmethod
    def reference_line(d: Dossier) -> str:
        """Viện dẫn phiếu tính toán của 1 tủ (không viện dẫn TBKQ — AGENTS mục 7)."""
        return (f"- Kết quả đo {cabinet_title(d)}: xem chi tiết tại phiếu Tính toán các thông số tủ "
                f"số {d.cert_no_vn} ngày {d.cert_date}.")

    def build_context(self) -> dict:
        d = self.dossier
        dtype = d.dossier_type.lower()
        rd = d.results_data or {}
        ctx = bckq_common_context(d)
        photos = [p for p in (d.photos or []) if os.path.isfile(p.get("path", "") if isinstance(p, dict) else str(p))]
        layouts = [p for p in (d.layouts or []) if os.path.isfile(p.get("path", "") if isinstance(p, dict) else str(p))]
        # Ảnh hiện trường không chú thích (AGENTS mục 3); khí nén giữ nhãn "Vị trí lấy mẫu KN1" (ảnh định danh vị trí)
        ctx["co_anh"], ctx["bang_anh"] = bool(photos), photo_grid(photos, False, captions=(dtype in ("compressed_air", "khi nen")))
        ctx["co_so_do"], ctx["bang_so_do"] = bool(layouts), photo_grid(layouts, False, cols=2 if len(layouts) >= 2 else 1)

        if dtype in ("compressed_air", "khi nen"):
            from .handlers.compressed_air import _chain_block
            no_data = not any(k in rd for k in ("oil", "dewpoint", "particles"))
            ctx["co_dau"] = no_data or bool((rd.get("oil") or {}).get("rows"))
            ctx["co_dp"] = no_data or bool((rd.get("dewpoint") or {}).get("rows"))
            ctx["co_hat"] = no_data or bool((rd.get("particles") or {}).get("rows"))
            params = [x for x, on in (("Hàm lượng tổng dầu", ctx["co_dau"]), ("Nhiệt độ điểm sương", ctx["co_dp"]),
                                      ("Nồng độ hạt bụi có kích thước (0,1-5,0) µm", ctx["co_hat"])) if on]
            kqtn_by_loc = rd.get("kqtn_by_location", {}) or {}
            ctx["vi_tri"] = []
            for loc in d.locations:
                sign = loc.points_summary_vn[0] if loc.points_summary_vn else loc.name_vn
                kqtt = (f"Xem tại phiếu Kết quả tính toán số {d.code} - {sign} ngày {d.cert_date}" if sign != d.code
                        else f"Xem tại phiếu Kết quả tính toán số {d.code} ngày {d.cert_date}")
                ctx["vi_tri"].append({"ten": loc.name_vn, "ky_hieu": sign, "chi_tieu_dong": "; ".join(params) + ".",
                                      "kqtn": kqtn_by_loc.get(sign, "MTTN………\nHDTN………"), "kqtt": kqtt})
            chain = rd.get("system_chain_vn")
            if isinstance(chain, str):
                chain = [c.strip() for c in chain.replace("→", ">").split(">") if c.strip()]
            ctx["so_do_he_thong"] = TagBlock(_chain_block, chain) if chain else PLACEHOLDER
            kq = rd.get("kqtn") or ([{}] if ctx["co_dau"] else [])
            ctx["kqtn"] = []
            for i, k in enumerate(kq):
                ctx["kqtn"].append({
                    "stt": "abcdefghijklmnopqrstuvwxyz"[i] if i < 26 else str(i + 1),
                    "so": k.get("so", PLACEHOLDER), "ngay": k.get("ngay", PLACEHOLDER),
                    "ngay_gui": k.get("ngay_gui", PLACEHOLDER), "ngay_yc": k.get("ngay_yc", PLACEHOLDER),
                    "ngay_nhan": k.get("ngay_nhan", PLACEHOLDER),
                    "tre_han": k.get("tre_han", "[  ] có\t\t[X] không"),
                })
        elif dtype in ("fumehood", "fume_hood", "tu hut", "tu_hut", "tủ hút"):
            cabs = rd.get("cabinets") or [rd.get("cabinet", {})]
            tasks = ["Đo tốc độ dòng khí qua cửa làm việc;"]
            if not (rd.get("lux") or {}).get("skip"):
                tasks.append("Đo cường độ sáng bên trong tủ;")
            tasks.append("Đo độ ồn của tủ khi vận hành;")
            if not (rd.get("airflow_pattern") or {}).get("skip"):
                tasks.append("Kiểm tra hình thái dòng khí của tủ;")
            ctx["tu"] = [{"thong_tin": self._cabinet_info(c or {}, "Tủ hút khí độc"),
                          "chi_tieu": (c or {}).get("chi_tieu_vn") or "\n".join(tasks)} for c in cabs]
            ctx["tham_chieu"] = [self.reference_line(d)]
            ctx["danh_gia"] = TagBlock(assessment_block, [cabinet_assessment(d)])
        elif dtype in ("bsc", "biological_safety_cabinet", "biosafety_cabinet", "laf", "laf_bsc"):
            cabs = rd.get("cabinets") or [rd.get("cabinet", {})]
            if rd.get("tasks_vn"):
                tasks = list(rd["tasks_vn"])
            else:
                # Theo chỉ tiêu thực hiện như mục 10.2 TBKQ (BSCHandler): bỏ inflow / downflow / lọc / hình thái
                # dòng khí khi skip (tủ cấy, BSC cấp I) — 139329/2 từng liệt kê nhầm inflow cho tủ cấy (10/11/2026)
                hepa = rd.get("hepa") or {}
                both_filters = not hepa.get("skip_exhaust") and not hepa.get("skip_downflow")
                tasks = []
                if not (rd.get("inflow") or {}).get("skip"):
                    tasks.append("Đo tốc độ dòng khí qua cửa làm việc;")
                if not (rd.get("downflow") or {}).get("skip"):
                    tasks.append("Đo tốc độ dòng khí cấp qua lọc HEPA;")
                tasks += ["Đo cường độ sáng bên trong tủ;", "Đo độ ồn của tủ khi vận hành;",
                          "Kiểm tra độ rò rỉ của các lọc HEPA;" if both_filters else "Kiểm tra độ rò rỉ của lọc HEPA;"]
                if not (rd.get("airflow_pattern") or {}).get("skip"):
                    tasks.append("Kiểm tra hình thái dòng khí của tủ;")
                if rd.get("uvc"):
                    tasks.append("Đo cường độ tia UV dải C trong khoang làm việc;")
                if rd.get("particle"):
                    tasks.append("Đo nồng độ hạt bụi trong tủ;")
            if tasks:
                tasks[-1] = tasks[-1].rstrip(".;") + "."
            ctx["tu"] = [{"thong_tin": self._cabinet_info(c or {}, "Tủ an toàn sinh học cấp 2"), "chi_tieu": "\n".join(tasks)}
                         for c in cabs]
            ctx["tham_chieu"] = [self.reference_line(d)]
            ctx["danh_gia"] = TagBlock(assessment_block, [cabinet_assessment(d)])
        else:  # phòng sạch
            ctx["ds_khu_vuc"] = "; ".join(quoted(loc.name_vn) for loc in d.locations) or PLACEHOLDER
            ctx["trang_thai"] = (d.locations[0].state_vn if d.locations else "") or PLACEHOLDER
            ctx["khu_vuc"] = []
            for loc in d.locations:
                area = (loc.area_m2 or "").replace("m²", "").replace("m2", "").strip()
                cap = (loc.clean_class or "").strip()
                cap = cap[4:].strip() if cap.upper().startswith("ISO ") else cap
                ctx["khu_vuc"].append({
                    "ten": loc.name_vn, "dien_tich": area.replace(".", ",") or PLACEHOLDER, "cap": cap or PLACEHOLDER,
                    "chi_tieu": [ln if ln.strip().startswith("-") else f"- {ln.strip()}" for ln in loc.points_summary_vn if ln.strip()] or [PLACEHOLDER],
                })
            # Chỉ viện dẫn phiếu tính toán / ghi chép, KHÔNG viện dẫn TBKQ (AGENTS.md mục 7)
            lines = []
            for item in d.assessment.get("items_list_vn", ["nồng độ hạt bụi", "tốc độ gió qua lọc HEPA"]):
                it = item.strip().rstrip(";")
                il = it.lower()
                if "bụi" in il:
                    sheet = "Tính toán kết quả nồng độ hạt bụi"
                elif "gió" in il or "hepa" in il or "lưu lượng" in il:
                    sheet = "Tính toán vận tốc gió và lưu lượng qua lọc HEPA"
                elif any(k in il for k in ("chênh áp", "vi khí hậu", "nhiệt độ", "độ ẩm", "độ rọi", "cường độ sáng")):
                    sheet = "Ghi chép giám định tại hiện trường"
                else:
                    sheet = f"Tính toán kết quả đo {it}"
                lines.append(f"- Kết quả đo {it}: xem tại phiếu {sheet} số {d.cert_no_vn} ngày {d.cert_date}.")
            ctx["tham_chieu"] = lines
            rows = d.assessment.get("rows")
            if not rows:
                rows = [{"khu_vuc": d.locations[0].name_vn if d.locations else "",
                         "dat": d.assessment.get("pass_vn", "Nồng độ hạt bụi;\nTốc độ gió qua lọc HEPA"),
                         "khong_dat": d.assessment.get("fail_vn", "-"),
                         "can_cu": d.assessment.get("basis_vn", d.basis_vn)}]
            ctx["danh_gia"] = rows
        return ctx

    def build(self, ctx: dict = None, out_name: str = None) -> str:
        """Builds and saves the BCKQ document, returning the saved path."""
        d = self.dossier
        doc = open_tag_template(self.template_path)
        render(doc, ctx if ctx is not None else self.build_context(), label=f"BCKQ {d.dossier_type}")
        format_bckq_table0(doc.tables[0])
        if is_cabinet(d) and len(doc.tables) > 1:
            tidy_cabinet_table(doc.tables[1])

        # Bỏ đoạn trống thừa (giữ đoạn có ảnh)
        for p in list(doc.paragraphs):
            if not p.text.strip() and "w:drawing" not in p._p.xml and p._p.find(f".//{{{W_NS}}}sectPr") is None:
                p._p.getparent().remove(p._p)
        strip_all_highlights(doc)
        for p in doc.paragraphs:  # 'Đánh giá:' đi liền bảng đánh giá
            if p.text.strip().startswith("Đánh giá"):
                p.paragraph_format.keep_with_next = True

        # Font Times New Roman 13 cho toàn bộ BCKQ (AGENTS.md mục 7)
        def _tnr13(paragraphs):
            for p in paragraphs:
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(13)
                    rFonts = r._r.get_or_add_rPr().get_or_add_rFonts()
                    rFonts.set(qn("w:ascii"), "Times New Roman")
                    rFonts.set(qn("w:hAnsi"), "Times New Roman")
                    rFonts.set(qn("w:cs"), "Times New Roman")
        _tnr13(doc.paragraphs)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    _tnr13(c.paragraphs)
        strip_image_crops(doc)

        # BCKQ KHÔNG có đường kết thúc (AGENTS.md mục 5)
        for p in list(doc.element.body.findall(f"{{{W_NS}}}p")):
            p_xml = p.xml
            if "thinThick" in p_xml or "Straight Connector" in p_xml:
                p.getparent().remove(p)

        out_name = out_name or f"BCKQ - {d.sub_no or 1} - {(d.results_data or {}).get('bckq_label', d.dossier_type)} - {d.code}.docx"
        target = save_outputs(doc, d, d.thuchien_dir, out_name, net_sub="ThucHien", label="bckq_builder.py")
        print(f"[BCKQBuilder] Successfully generated -> {target}")
        return target

    # ------------------------------------------------------------------
    @classmethod
    def build_merged(cls, dossiers, out_name: str = None) -> str:
        """Gộp BCKQ các tủ (BSC/LAF, tủ cấy, tủ hút) của cùng hồ sơ thành 1 file (người dùng cho phép 30/10/2026).

        Bảng 0 dùng chung (số = mã hồ sơ trần vì file bao gồm mọi số phụ), bảng thiết bị 1 dòng / tủ,
        viện dẫn phiếu tính toán từng tủ (/1, /2...), Đánh giá chia theo tủ, mỗi tủ 3 phần.
        Tên file mặc định: BCKQ - tu - <mã>.docx.
        """
        dossiers = [d for d in dossiers if is_cabinet(d)]
        if not dossiers:
            raise ValueError("build_merged: không có báo cáo tủ nào")
        if len({d.code for d in dossiers}) > 1:
            raise ValueError("build_merged: chỉ gộp báo cáo cùng mã hồ sơ")
        d0 = dossiers[0]
        ctxs = [cls(d).build_context() for d in dossiers]
        ctx = dict(ctxs[0])
        ctx["so_bc"] = d0.code.strip()

        # Nội dung yêu cầu: 'Giám định 01 tủ cấy, 03 tủ hút' (người dùng chốt 01/11/2026)
        prefix = "Giám định "
        scopes = [inspection_scope(d.scope_vn) for d in dossiers]
        if all(sc.startswith(prefix) for sc in scopes):
            counts = {}
            for sc in scopes:
                key = sc[len(prefix):].strip()
                counts[key] = counts.get(key, 0) + 1
            ctx["noi_dung_yc"] = prefix + ", ".join(f"{n:02d} {k}" for k, n in counts.items())
        else:
            ctx["noi_dung_yc"] = "; ".join(dict.fromkeys(sc for sc in scopes if sc)) or ctx["noi_dung_yc"]
        ctx["tg_th"] = "; ".join(dict.fromkeys(c["tg_th"] for c in ctxs))
        ctx["nguoi_th"] = ", ".join(dict.fromkeys(c["nguoi_th"] for c in ctxs))
        ctx["quy_trinh"] = inspection_procedure(", ".join(c.get("quy_trinh", "") for c in ctxs), bsc_class=0)

        # Thiết bị: hợp danh sách các tủ, giữ thứ tự, không trùng
        merged = copy.copy(d0)
        seen, eqs = set(), []
        for d in dossiers:
            for eq in d.equipments:
                key = (eq.name_vn or "").strip()
                if key and key not in seen:
                    seen.add(key)
                    eqs.append(eq)
        merged.equipments = eqs
        ctx["thiet_bi"] = equipment_text(merged, False)

        ctx["tu"] = [t for c in ctxs for t in c["tu"]]
        ctx["tham_chieu"] = [cls.reference_line(d) for d in dossiers]
        ctx["danh_gia"] = TagBlock(assessment_block, [cabinet_assessment(d) for d in dossiers])
        return cls(d0).build(ctx=ctx, out_name=out_name or f"BCKQ - tu - {d0.code}.docx")

