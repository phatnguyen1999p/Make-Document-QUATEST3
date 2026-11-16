# -*- coding: utf-8 -*-
"""
Compressed Air Document Handler (ISO 8573-1:2010) for QUATEST 3.

Sinh TBKQ khí nén (VN/EN) bằng template dạng TAG (Core/Templates/TBKQ - vn/en - khi nen.docx):
1. Dựng context từ Dossier (trang bìa + kết quả: dầu, điểm sương, hạt bụi, mức yêu cầu của khách hàng).
2. Điền TAG (Core/Tools/tag_engine.py).
3. Hậu xử lý chung: in đậm nghiêng cụm đánh giá, lặp tiêu đề bảng, đường kết thúc, bỏ highlight.

Danh mục TAG: Core/Templates/TAGS.md (mục "TBKQ khí nén").
"""

import copy
import os
import re

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

from ...Models.dossier import Dossier
from ...Styling import W_NS, strip_all_highlights, add_document_ending_line, drop_section_gaps, strip_image_crops, fix_unit_symbols
from ...Tools.tag_engine import TagBlock
from ...Tools.iso8573_cal import classify_iso8573_oil
from ..tag_context import (
    PLACEHOLDER,
    cover_context,
    dec,
    thousands,
    photo_grid,
    format_assessments,
    keep_intros_with_next,
    style_note_tables,
    ensure_result_tables,
    requirement_text,
    number_of,
    complies,
    verdict,
    open_tag_template,
    save_outputs,
    render,
)

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _chain_block(doc, anchor_p, ctx, items):
    """Sơ đồ hệ thống khí nén: các thiết bị nối bằng mũi tên Wingdings (như form gốc)."""
    p = copy.deepcopy(anchor_p)
    runs = [r for r in p.findall(qn("w:r"))]
    rpr = runs[0].find(qn("w:rPr")) if runs else None
    for r in runs:
        p.remove(r)
    for i, name in enumerate(items):
        if i:
            r_sym = OxmlElement("w:r")
            if rpr is not None:
                r_sym.append(copy.deepcopy(rpr))
            r_sym.append(parse_xml(f'<w:sym xmlns:w="{W_NS}" w:font="Wingdings" w:char="F0E0"/>'))
            p.append(r_sym)
        r = OxmlElement("w:r")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = OxmlElement("w:t")
        t.text = f" {name} " if i else f"{name} "
        t.set(XML_SPACE, "preserve")
        r.append(t)
        p.append(r)
    return [p]


def _tidy_note_marks(doc):
    """'(1) ...' của form gốc (mục hạt bụi) -> '(1) …' như các ghi chú khác; qa_check coi '...' là ô chưa điền."""
    for p in doc.element.body.iter(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        if re.fullmatch(r"\s*\(\d\)\s*\.\.\.\s*", "".join(t.text or "" for t in ts)):
            for t in ts:
                if t.text and "..." in t.text:
                    t.text = t.text.replace("...", "…")


def _style_result_cells(doc, ctx):
    """Ô kết quả (AGENTS mục 3): '(3)' cuối ô 'Không phân loại' -> chỉ số trên; giá trị không đạt mức yêu cầu -> in đậm."""
    from docx.table import Table
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in list(p.runs):
                        m = re.search(r"\(\d\)$", r.text)
                        if not m or not r.text[:m.start()].strip() or r.font.superscript:
                            continue
                        r.text = r.text[:m.start()]
                        sup = copy.deepcopy(r._r)
                        r._r.addnext(sup)
                        from docx.text.run import Run
                        Run(sup, p).text = m.group(0)
                        Run(sup, p).font.superscript = True
    # (từ khóa tiêu đề bảng, cột giá trị, mức yêu cầu)
    checks = []
    if ctx.get("dau") and ctx["dau"].get("yc"):
        checks.append((("Hàm lượng tổng dầu", "Total oil content"), 1, ctx["dau"]["yc"]))
    if ctx.get("dp") and ctx["dp"].get("yc_co"):
        checks.append((("Điều kiện đo", "Measurement condition"), 5, ctx["dp"]["yc"][4]))
    if ctx.get("hat") and ctx["hat"].get("yc"):
        checks.append((("d > 0,5", "d > 0.5"), 1, ctx["hat"]["yc"]))
    for keys, col, lim in checks:
        for tbl in doc.tables:
            head = " ".join(c.text for c in tbl.rows[0].cells)
            if not any(k in head for k in keys):
                continue
            for row in tbl.rows[1:]:
                cells = row.cells
                if len(cells) <= col or not re.match(r"KN\d*", cells[0].text.strip()):
                    continue
                if complies([cells[col].text.strip()], lim) is False:
                    for p in cells[col].paragraphs:
                        for r in p.runs:
                            r.font.bold = True


def _renumber_result_headings(doc):
    """Bản EN ghi cứng '9.2.1 / 9.2.2 / 9.2.3' (VN dùng đánh số tự động): bỏ mục dầu / điểm sương thì EN lệch
    (137035 chỉ đo bụi ra '9.2.3') -> đánh lại theo thứ tự còn lại."""
    import re
    n = 0
    for p in doc.paragraphs:
        txt = p.text.strip()
        m = re.match(r"^9\.2\.(\d)\s", txt)
        if not m:
            continue
        n += 1
        if int(m.group(1)) != n:          # số có thể vắt qua nhiều run ('9' '.2.' '3 ')
            runs = [r for r in p.runs if r.text]
            rest = len(p.text) - len(p.text.lstrip()) + len("9.2.") + 1   # khoảng trắng đầu + "9.2.N"
            for r in runs:                # xóa tiền tố cũ khỏi các run
                k = min(rest, len(r.text))
                r.text = r.text[k:]
                rest -= k
                if rest <= 0:
                    break
            runs[0].text = f"9.2.{n}" + runs[0].text


def keep_result_heading_with_next(doc):
    """Tiêu đề '9.2. Kết quả giám định' + đoạn trống sau nó -> keepNext. Không gọi mặc định: 133266 bản EN dài hơn
    VN ~3 dòng, giữ liền thì EN ngắt khác VN — chỉ gọi khi đã soát PDF VN = EN (vd build_137035)."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    for p in doc.paragraphs:          # VN: tiêu đề "Kết quả giám định" đánh số tự động (không có chữ "9.2")
        if p.text.strip() not in ("Kết quả giám định", "Inspection result"):
            continue
        p.paragraph_format.keep_with_next = True
        nxt = p._p.getnext()          # đoạn trống còn lại sau khi bỏ khối {%p if dau %} -> cũng keepNext
        while nxt is not None and nxt.tag == qn("w:p") and not "".join(
                t.text or "" for t in nxt.iter(qn("w:t"))).strip():
            ppr = nxt.get_or_add_pPr()
            if ppr.find(qn("w:keepNext")) is None:
                ppr.insert(0, OxmlElement("w:keepNext"))
            nxt = nxt.getnext()


def _first_requirement(rd: dict, key_rd: str, key_pp: str) -> str:
    req = (rd.get(key_rd) or {}).get("requirement") if isinstance(rd.get(key_rd), dict) else None
    if req:
        return str(req)
    for pp in rd.get("parsed_points", []) or []:
        val = (pp.get(key_pp) or {}).get("requirement")
        if val:
            return str(val)
    return ""


def build_context(d: Dossier, is_en: bool) -> dict:
    """Context đầy đủ cho TBKQ khí nén (xem TAGS.md)."""
    ctx = cover_context(d, is_en)
    rd = d.results_data or {}
    client = ctx["kh_ten"]
    n = len(d.locations)
    ctx["so_vi_tri"] = (f"{n:02d} location" + ("s" if n > 1 else "")) if is_en else f"{n:02d} vị trí"
    ctx["trang_thai_he_thong"] = rd.get("system_status_en" if is_en else "system_status_vn") or ("In operation" if is_en else "Đang hoạt động")
    ctx["trang_thai_nha_xuong"] = rd.get("factory_status_en" if is_en else "factory_status_vn") or ("In operation" if is_en else "Đang hoạt động")
    chain = rd.get("system_chain_en" if is_en else "system_chain_vn")
    if isinstance(chain, str):
        chain = [c.strip() for c in chain.replace("→", ">").split(">") if c.strip()]
    ctx["so_do_he_thong"] = TagBlock(_chain_block, chain) if chain else PLACEHOLDER

    has_oil = bool((rd.get("oil") or {}).get("rows"))
    has_dp = bool((rd.get("dewpoint") or {}).get("rows"))
    has_part = bool((rd.get("particles") or {}).get("rows"))
    params = []
    if has_oil:
        params.append("Total oil content;" if is_en else "Hàm lượng tổng dầu;")
    if has_dp:
        params.append("Pressure dew point;" if is_en else "Nhiệt độ điểm sương;")
    if has_part:
        params.append("Concentration of particle sizes (0.1-5.0) µm." if is_en else "Nồng độ hạt bụi có kích thước (0,1-5,0) µm.")
    if params:
        params[-1] = params[-1].rstrip(";.") + "."

    ctx["vi_tri"] = []
    for loc in d.locations:
        sym = (loc.points_summary_en[0] if is_en and loc.points_summary_en else
               (loc.points_summary_vn[0] if loc.points_summary_vn else loc.name_vn))
        ctx["vi_tri"].append({"ten": (loc.name_en if is_en else loc.name_vn), "ky_hieu": sym, "chi_tieu": params})

    photos = [p for p in (d.photos or []) if os.path.isfile(p.get("path", "") if isinstance(p, dict) else str(p))]
    ctx["co_anh"] = bool(photos)
    ctx["bang_anh"] = photo_grid(photos, is_en)
    many = n > 1
    at_loc = ("at the locations mentioned above" if many else "at the location mentioned above") if is_en else \
             ("tại các vị trí nêu trên" if many else "tại vị trí nêu trên")

    # --- Hàm lượng tổng dầu ---
    if has_oil:
        oil = rd["oil"]
        dong = [{"vi_tri": r.get("loc", ""), "gia_tri": dec(r.get("val", ""), is_en), "phan_loai": r.get("class", "")}
                for r in oil["rows"]]
        req_raw = _first_requirement(rd, "oil", "total_oil")
        yc = requirement_text(req_raw, is_en)
        yc_cls = oil.get("requirement_class") or ""
        if yc and not yc_cls:
            v = number_of(yc)
            yc_cls = classify_iso8573_oil(v) if v is not None else "-"
        ok = complies([r["gia_tri"] for r in dong], yc) if yc else None
        kl = verdict(ok, is_en, oil.get("assessment_en" if is_en else "assessment_vn"))
        ctx["dau"] = {
            "dong": dong, "yc": yc, "yc_phan_loai": yc_cls,
            "danh_gia": (f"The test result of total oil content in compressed air {at_loc} {kl} with the corresponding "
                         f"requirement provided by {client}." if is_en else
                         f"Kết quả đo hàm lượng tổng dầu trong khí nén {at_loc} {kl} với mức yêu cầu tương ứng được "
                         f"cung cấp bởi {client}."),
        }
    else:
        ctx["dau"] = None

    # --- Điểm sương ---
    if has_dp:
        dp = rd["dewpoint"]
        dong = []
        for r in dp["rows"]:
            vals = [dec(v, is_en) for v in (r.get("values") or [])]
            vals += [""] * (6 - len(vals))
            dong.append({"vi_tri": r.get("loc", ""), "v": vals[:6]})
        req_raw = _first_requirement(rd, "dewpoint", "dewpoint")
        yc_water = requirement_text(req_raw, is_en)
        yc = ["-", "-", "-", "-", yc_water or "-", "-"]
        ok = complies([r["v"][4] for r in dong], yc_water) if yc_water else None
        kl = verdict(ok, is_en, dp.get("assessment_en" if is_en else "assessment_vn"))
        ctx["dp"] = {
            "dong": dong, "yc": yc, "yc_co": bool(yc_water),
            "danh_gia": (f"The test result of water content in compressed air {at_loc} {kl} with the corresponding "
                         f"requirement provided by {client}." if is_en else
                         f"Kết quả đo hàm lượng nước trong khí nén {at_loc} {kl} với mức yêu cầu tương ứng được "
                         f"cung cấp bởi {client}."),
        }
    else:
        ctx["dp"] = None

    # --- Nồng độ hạt bụi ---
    if has_part:
        pr = rd["particles"]
        dong = []
        any_gt5 = False
        for r in pr["rows"]:
            vals = list(r.get("values") or [])
            counts = [dec(v, is_en) for v in vals[:3]] + [""] * max(0, 3 - len(vals[:3]))
            gt5 = r.get("d_gt_50", 0) or 0
            try:
                gt5_num = int(float(str(gt5).replace(" ", "").replace(",", ".")))
            except ValueError:
                gt5_num = 0
            if gt5_num > 0:
                any_gt5 = True
                cls = "Not classified\n(3)" if is_en else "Không\nphân loại\n(3)"
                mo_ta = (f"Airborne particle concentration of size d > 5.0 µm: {thousands(gt5_num)} "
                         f"particle{'s' if gt5_num > 1 else ''}/m3" if is_en
                         else f"Nồng độ hạt bụi có kích thước d > 5,0 µm: {thousands(gt5_num)} hạt/m3")
            else:
                cls = (r.get("class_en") if is_en else r.get("class_vi")) or next(
                    (str(v) for v in reversed(vals) if "class" in str(v).lower() or "phân loại" in str(v).lower()),
                    str(vals[3]) if len(vals) > 3 else "")
                mo_ta = ("Airborne particle concentration of size d > 5.0 µm: 0 particle/m3" if is_en
                         else "Nồng độ hạt bụi có kích thước d > 5,0 µm: 0 hạt/m3")
            dong.append({"vi_tri": r.get("loc", ""), "v": counts, "phan_loai": cls, "mo_ta_5um": mo_ta})
        req_raw = _first_requirement(rd, "particles", "particles")
        yc = requirement_text(req_raw, is_en)
        yc_dong = []
        for pp in rd.get("parsed_points", []) or []:
            v = (pp.get("particles") or {}).get("d_gt_05_total", 0)
            yc_dong.append({"vi_tri": pp.get("location_symbol", ""), "d05": thousands(v)})
        if yc and not yc_dong:
            yc = ""  # không có số liệu d > 0,5 µm thì không lập bảng so sánh mức yêu cầu
        ok = complies([r["d05"] for r in yc_dong], yc) if yc else None
        kl = verdict(ok, is_en, pr.get("assessment_en" if is_en else "assessment_vn"))
        ctx["hat"] = {
            "dong": dong, "co_khong_phan_loai": any_gt5, "yc": yc, "yc_dong": yc_dong,
            "danh_gia": (f"The measurement result of concentration of particle size d > 0.5 µm in compressed air {at_loc} "
                         f"{kl} with the corresponding requirement provided by {client}." if is_en else
                         f"Kết quả đo nồng độ hạt bụi có kích thước d > 0,5 µm trong khí nén {at_loc} {kl} với mức "
                         f"yêu cầu tương ứng được cung cấp bởi {client}."),
        }
    else:
        ctx["hat"] = None
    return ctx


class CompressedAirHandler:
    """Builder handler for ISO 8573-1 Compressed Air inspection reports (template dạng TAG)."""

    def __init__(self, dossier: Dossier, template_path_vn: str = None, template_path_en: str = None):
        self.dossier = dossier
        templates_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "Templates")
        self.template_vn = template_path_vn or os.path.join(templates_dir, "TBKQ - vn - khi nen.docx")
        self.template_en = template_path_en or os.path.join(templates_dir, "TBKQ - en - khi nen.docx")

    def build_tbkq(self, is_en: bool = False, has_assessment: bool = True, **kwargs) -> str:
        """Builds complete TBKQ document (VN or EN) from the tagged Compressed Air template."""
        d = self.dossier
        doc = open_tag_template(self.template_en if is_en else self.template_vn)
        ctx = build_context(d, is_en)
        render(doc, ctx, label=f"TBKQ khí nén {'EN' if is_en else 'VN'}")

        ensure_result_tables(doc)
        format_assessments(doc)
        keep_intros_with_next(doc)
        style_note_tables(doc)          # bảng Ghi chú 11pt in nghiêng (09/11/2026)
        _renumber_result_headings(doc)
        fix_unit_symbols(doc)
        _tidy_note_marks(doc)
        _style_result_cells(doc, ctx)
        drop_section_gaps(doc)          # đoạn trống sót giữa các mục (mục bị xóa khỏi template) — AGENTS mục 3
        add_document_ending_line(doc)
        strip_all_highlights(doc)
        strip_image_crops(doc)

        out_name = f"TBKQ - {'en' if is_en else 'vn'} - {d.sub_no or 2} - khi nen - {d.code}.docx"
        target = save_outputs(doc, d, d.tbkq_dir, out_name, label="compressed_air.py")
        print(f"[CompressedAirHandler] Successfully generated TBKQ {'EN' if is_en else 'VN'} -> {target}")
        return target
