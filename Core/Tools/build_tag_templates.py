# -*- coding: utf-8 -*-
"""
Dựng lại toàn bộ template dạng TAG trong Core/Templates từ bản gốc ở Core/Templates/_goc_truoc_TAG/.

    python -m Core.Tools.build_tag_templates            # dựng tất cả
    python -m Core.Tools.build_tag_templates --only khi_nen,bckq

Nguyên tắc:
- Bản gốc (chứa dữ liệu hồ sơ cũ) giữ nguyên trong _goc_truoc_TAG/, KHÔNG sửa.
- Bản TAG giữ nguyên bố cục/định dạng của bản gốc; dữ liệu khách hàng/hồ sơ cũ được thay bằng TAG {{ ... }}.
- Trang bìa (header + Table 0) được dựng bằng chính hàm định dạng cũ (update_headers, populate_cleanroom_table0)
  với giá trị là TAG, nên định dạng (gộp ô, khoảng ký 3 cm, căn chữ ký) giống hệt bản do code cũ sinh ra.
- Danh mục TAG: Core/Templates/TAGS.md.
"""

import argparse
import copy
import os
import re
import shutil

import docx
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Pt

from ..Models.dossier import Dossier, ClientInfo, Equipment, PersonnelInfo
from ..Styling import W_NS, update_headers, populate_cleanroom_table0, set_paragraph_text
from .tag_engine import _p_text_nodes, _p_text, xlsx_set_cells_raw

CORE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL_DIR = os.path.join(CORE_DIR, "Templates")
SRC_DIR = os.path.join(TPL_DIR, "_goc_truoc_TAG")

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_P, _TBL, _TR, _TC, _R, _T = qn("w:p"), qn("w:tbl"), qn("w:tr"), qn("w:tc"), qn("w:r"), qn("w:t")


# ===========================================================================
# Tiện ích sửa template
# ===========================================================================
def src(name: str) -> str:
    p = os.path.join(SRC_DIR, name)
    if not os.path.isfile(p):
        raise FileNotFoundError(f"Thiếu bản gốc: {p}")
    return p


def dst(name: str) -> str:
    return os.path.join(TPL_DIR, name)


def _el(x):
    """Phần tử lxml từ đối tượng python-docx (Paragraph/_Cell/_Row/Table) hoặc trả nguyên nếu đã là lxml."""
    from lxml import etree
    if isinstance(x, etree._Element):
        return x
    for attr in ("_p", "_tc", "_tr", "_tbl"):
        if hasattr(x, attr):
            return getattr(x, attr)
    return x


def strip_highlight(el):
    for h in list(el.iter(qn("w:highlight"))):
        h.getparent().remove(h)


def set_p(p, text: str, keep_run: int = 0):
    """Đặt chữ cho đoạn, GIỮ định dạng run thứ keep_run (pPr giữ nguyên), xóa các run còn lại."""
    p = _el(p)
    runs = [r for r in p.iter(_R) if r.getparent() is p or r.getparent().tag in (qn("w:hyperlink"), qn("w:smartTag"))]
    text_runs = [r for r in runs if r.find(_T) is not None]
    base = text_runs[min(keep_run, len(text_runs) - 1)] if text_runs else (runs[0] if runs else None)
    if base is None:
        r = OxmlElement("w:r")
        p.append(r)
        base = r
    for r in runs:
        if r is not base:
            r.getparent().remove(r)
    for child in list(base):
        if child.tag != qn("w:rPr"):
            base.remove(child)
    t = OxmlElement("w:t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    base.append(t)
    strip_highlight(base)
    return p


def replace_p(p, old: str, new: str, count: int = 1) -> bool:
    """Thay chuỗi `old` -> `new` trong đoạn (kể cả khi `old` bị tách qua nhiều run). Giữ định dạng run đầu."""
    p = _el(p)
    done = 0
    while done < count:
        nodes = _p_text_nodes(p)
        texts = [n.text or "" for n in nodes]
        full = "".join(texts)
        i = full.find(old)
        if i < 0:
            break
        starts, pos = [], 0
        for t in texts:
            starts.append(pos)
            pos += len(t)

        def loc(ch):
            for k in range(len(nodes)):
                if starts[k] <= ch < starts[k] + len(texts[k]):
                    return k
            return len(nodes) - 1
        a, b = loc(i), loc(i + len(old) - 1)
        oa, ob = i - starts[a], i + len(old) - 1 - starts[b]
        if a == b:
            nodes[a].text = texts[a][:oa] + new + texts[a][ob + 1:]
        else:
            nodes[a].text = texts[a][:oa] + new
            for k in range(a + 1, b):
                nodes[k].text = ""
            nodes[b].text = texts[b][ob + 1:]
        for k in range(a, b + 1):
            nodes[k].set(XML_SPACE, "preserve")
        done += 1
    return done > 0


def fill_hl(p, tags):
    """Thay lần lượt các chỗ trống tô vàng '………' trong đoạn bằng các TAG trong danh sách."""
    p = _el(p)
    tags = list(tags)
    for r in [r for r in p.iter(_R)]:
        if not tags:
            break
        rpr = r.find(qn("w:rPr"))
        hl = rpr is not None and rpr.find(qn("w:highlight")) is not None
        t = r.find(_T)
        if t is None or not hl:
            continue
        if "…" in (t.text or "") or (t.text or "").strip() == "":
            t.text = re.sub(r"…+", tags.pop(0), t.text or "", count=1) if "…" in (t.text or "") else tags.pop(0)
            t.set(XML_SPACE, "preserve")
    # Còn '………' không tô vàng (vd trong ngoặc kép) -> thay nốt
    while tags:
        if not replace_p(p, "………", tags[0]):
            break
        tags.pop(0)
    strip_highlight(p)
    if tags:
        raise ValueError(f"Không đủ chỗ trống để điền {tags} vào đoạn: {_p_text(p)[:80]!r}")


def set_cell(cell_el, text: str):
    """Đặt chữ cho ô: giữ đoạn đầu (định dạng run đầu), xóa các đoạn còn lại."""
    tc = _el(cell_el)
    ps = tc.findall(_P)
    for p in ps[1:]:
        tc.remove(p)
    set_p(ps[0], text)


def cell(tbl_el, r: int, c: int):
    """Ô thực (w:tc) thứ c của dòng r (không tính gridSpan)."""
    tr = tbl_el.findall(_TR)[r]
    return tr.findall(_TC)[c]


def rows(tbl_el):
    return tbl_el.findall(_TR)


def new_p_like(ref_p, text: str, bold=None):
    """Đoạn mới, sao chép pPr của ref_p (bỏ numPr), một run Times New Roman 13."""
    ref_p = _el(ref_p)
    p = OxmlElement("w:p")
    ppr = ref_p.find(qn("w:pPr"))
    if ppr is not None:
        ppr = copy.deepcopy(ppr)
        for tag in ("w:numPr", "w:sectPr", "w:pStyle"):
            for x in ppr.findall(qn(tag)):
                ppr.remove(x)
        p.append(ppr)
    r = OxmlElement("w:r")
    rpr = parse_xml(f'<w:rPr xmlns:w="{W_NS}"><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
                    f'{"<w:b/>" if bold else ""}<w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>')
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    r.append(t)
    p.append(r)
    return p


def marker_p(text: str):
    """Đoạn chỉ chứa TAG khối {%p ... %} (bị xóa khi điền)."""
    p = OxmlElement("w:p")
    ppr = parse_xml(f'<w:pPr xmlns:w="{W_NS}"><w:spacing w:before="0" w:after="0"/></w:pPr>')
    p.append(ppr)
    r = OxmlElement("w:r")
    r.append(parse_xml(f'<w:rPr xmlns:w="{W_NS}"><w:color w:val="7F7F7F"/><w:sz w:val="18"/></w:rPr>'))
    t = OxmlElement("w:t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    r.append(t)
    p.append(r)
    return p


def wrap_p(first_el, last_el, open_tag: str, close_tag: str = None):
    """Bọc các phần tử thân văn bản từ first_el đến last_el bằng {%p ... %} ... {%p end... %}."""
    first_el, last_el = _el(first_el), _el(last_el)
    if close_tag is None:
        kw = open_tag.split()[0]
        close_tag = "endfor" if kw == "for" else "endif"
    first_el.addprevious(marker_p("{%p " + open_tag + " %}"))
    last_el.addnext(marker_p("{%p " + close_tag + " %}"))


def marker_row(ref_tr, text: str):
    """Dòng chỉ chứa TAG khối {%tr ... %}: gộp toàn bộ dòng thành 1 ô."""
    ref_tr = _el(ref_tr)
    tr = copy.deepcopy(ref_tr)
    tcs = tr.findall(_TC)
    total = 0
    for tc in tcs:
        tcpr = tc.find(qn("w:tcPr"))
        gs = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
        total += int(gs.get(qn("w:val"))) if gs is not None else 1
    first = tcs[0]
    for tc in tcs[1:]:
        tr.remove(tc)
    tcpr = first.find(qn("w:tcPr"))
    if tcpr is None:
        tcpr = OxmlElement("w:tcPr")
        first.insert(0, tcpr)
    for tag in ("w:vMerge", "w:gridSpan", "w:hMerge"):
        for x in tcpr.findall(qn(tag)):
            tcpr.remove(x)
    if total > 1:
        gs = OxmlElement("w:gridSpan")
        gs.set(qn("w:val"), str(total))
        # gridSpan phải đứng sau tcW
        tcw = tcpr.find(qn("w:tcW"))
        if tcw is not None:
            tcw.addnext(gs)
        else:
            tcpr.insert(0, gs)
    trpr = tr.find(qn("w:trPr"))
    if trpr is not None:
        for tag in ("w:tblHeader", "w:trHeight"):
            for x in trpr.findall(qn(tag)):
                trpr.remove(x)
    for p in first.findall(_P)[1:]:
        first.remove(p)
    p0 = first.find(_P)
    set_p(p0, text)
    for r in p0.iter(_R):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        for tag in ("w:b", "w:i"):
            for x in rpr.findall(qn(tag)):
                rpr.remove(x)
    return tr


def wrap_rows(first_tr, last_tr, open_tag: str, close_tag: str = None, ref_tr=None):
    first_tr, last_tr = _el(first_tr), _el(last_tr)
    if close_tag is None:
        kw = open_tag.split()[0]
        close_tag = "endfor" if kw == "for" else "endif"
    ref = _el(ref_tr) if ref_tr is not None else first_tr
    first_tr.addprevious(marker_row(ref, "{%tr " + open_tag + " %}"))
    last_tr.addnext(marker_row(ref, "{%tr " + close_tag + " %}"))


def remove(el):
    el = _el(el)
    el.getparent().remove(el)


def strip_connectors(body):
    """Gỡ các đường kẻ kết thúc (Straight Connector/thinThick) mồ côi trong template — code thêm lại ở cuối khi điền."""
    MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
    for r in list(body.iter(_R)):
        xml = r.xml if hasattr(r, "xml") else ""
        if r.find(f".//{MC}") is not None or r.find(qn("w:drawing")) is not None or r.find(qn("w:pict")) is not None:
            from lxml import etree
            s = etree.tostring(r, encoding="unicode")
            if "thinThick" in s or "Straight Connector" in s or "v:line" in s:
                r.getparent().remove(r)


def body_children(doc):
    return list(doc.element.body)


def _nfc(s: str) -> str:
    import unicodedata
    return unicodedata.normalize("NFC", (s or "").replace("\u00a0", " "))


def find_p(doc, prefix: str, start_el=None, contains: bool = False):
    """Đoạn thân văn bản (cấp body) đầu tiên có chữ bắt đầu bằng / chứa `prefix` (sau start_el nếu có)."""
    started = start_el is None
    for el in doc.element.body:
        if not started:
            if el is _el(start_el):
                started = True
            continue
        if el.tag != _P:
            continue
        txt = _nfc(_p_text(el)).strip()
        pre = _nfc(prefix)
        if (pre in txt) if contains else txt.startswith(pre):
            return el
    raise KeyError(f"Không tìm thấy đoạn: {prefix!r}")


def next_tbl(el):
    el = _el(el).getnext()
    while el is not None and el.tag != _TBL:
        el = el.getnext()
    if el is None:
        raise KeyError("Không có bảng phía sau")
    return el


def body_tables(doc):
    return [el for el in doc.element.body if el.tag == _TBL]


def remove_between(a, b):
    """Xóa mọi phần tử nằm giữa a và b (kể cả ảnh cũ của hồ sơ trước), trừ đoạn ngắt section."""
    a, b = _el(a), _el(b)
    el = a.getnext()
    while el is not None and el is not b:
        nxt = el.getnext()
        if el.find(f".//{qn('w:sectPr')}") is None and el.tag in (_P, _TBL):
            el.getparent().remove(el)
        el = nxt


def _field_runs(rpr, instr: str, cached: str):
    """Các run của một trường Word: begin / instrText / separate / giá trị tạm / end."""
    from lxml import etree
    rpr_xml = "" if rpr is None else etree.tostring(rpr, encoding="unicode")
    out = []
    for kind in ("begin", "instr", "separate", "num", "end"):
        nr = parse_xml(f'<w:r xmlns:w="{W_NS}">{rpr_xml}</w:r>')
        if kind == "instr":
            it = OxmlElement("w:instrText")
            it.set(XML_SPACE, "preserve")
            it.text = f" {instr}  \\* Arabic  \\* MERGEFORMAT "
            nr.append(it)
        elif kind == "num":
            nt = OxmlElement("w:t")
            nt.text = cached
            nr.append(nt)
        else:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), kind)
            nr.append(fc)
        out.append(nr)
    return out


def use_page_fields(doc, total_field: str = "SECTIONPAGES", kinds=("footer",)):
    """
    'Trang X/Y' ở header/footer: Y cố định -> trường tự tính (lỗi E15).
    TBKQ: phần Phụ lục đánh số lại từ 1 -> SECTIONPAGES. BCKQ (1 section): NUMPAGES.
    Nếu cả X cũng là chữ cố định ('Trang 1/3') thì X -> trường PAGE.
    """
    for rel in doc.part.rels.values():
        if rel.is_external:
            continue
        part = rel.target_part
        ct = getattr(part, "content_type", "")
        if not any(ct.endswith(k + "+xml") for k in kinds):
            continue
        for p in part.element.iter(_P):
            txt = _p_text(p)
            if not re.search(r"(Trang|Page)\s*\d*\s*/\s*\d+", txt):
                continue
            instr = "".join(x.text or "" for x in p.iter(qn("w:instrText")))
            if re.search(r"\b(NUMPAGES|SECTIONPAGES)\b", instr):
                continue   # đã có trường tổng số trang (form gốc BCKQ) -> không thêm trường thứ hai (lỗi E20 "Trang 1/22")
            has_page = "PAGE" in instr
            runs = p.findall(_R)
            for i, r in enumerate(runs):
                t = r.find(_T)
                if t is None or "/" not in (t.text or ""):
                    continue
                head, _, tail = t.text.partition("/")
                rpr = r.find(qn("w:rPr"))
                anchor = r
                if not has_page:
                    m = re.match(r"^(.*?(?:Trang|Page)\s*)(\d+)\s*$", head)
                    if not m:
                        break
                    t.text = m.group(1)
                    t.set(XML_SPACE, "preserve")
                    for fr in _field_runs(rpr, "PAGE", m.group(2)):
                        anchor.addnext(fr)
                        anchor = fr
                    slash = parse_xml(f'<w:r xmlns:w="{W_NS}"></w:r>')
                    if rpr is not None:
                        slash.append(copy.deepcopy(rpr))
                    st = OxmlElement("w:t")
                    st.text = "/"
                    slash.append(st)
                    anchor.addnext(slash)
                    anchor = slash
                else:
                    t.text = head + "/"
                    t.set(XML_SPACE, "preserve")
                cached = re.sub(r"\D", "", tail)
                for r2 in runs[i + 1:]:
                    t2 = r2.find(_T)
                    if t2 is not None and re.fullmatch(r"\s*\d+\s*", t2.text or ""):
                        cached = cached or t2.text.strip()
                        p.remove(r2)
                for fr in _field_runs(rpr, total_field, cached or "1"):
                    anchor.addnext(fr)
                    anchor = fr
                break


def use_sectionpages(doc):
    use_page_fields(doc, "SECTIONPAGES", ("footer",))


def remove_empty_p_between(a, b):
    """Xóa các đoạn trống (không ảnh/không chữ) nằm giữa a và b (không gồm a, b)."""
    a, b = _el(a), _el(b)
    el = a.getnext()
    while el is not None and el is not b:
        nxt = el.getnext()
        if el.tag == _P and not _p_text(el).strip() and el.find(f".//{qn('w:drawing')}") is None \
                and el.find(f".//{qn('w:sectPr')}") is None:
            el.getparent().remove(el)
        el = nxt


# ===========================================================================
# Trang bìa (header + Table 0) — chung cho mọi TBKQ
# ===========================================================================
def tag_dossier(dossier_type: str = "cleanroom") -> Dossier:
    d = Dossier(
        code="{{ ma_hs }}",
        cert_date="{{ ngay_ct }}",
        sample_date="{{ ngay_gd }}",
        reg_date="{{ ngay_dk }}",
        sample_time_vn="",
        sample_time_en="",
        dossier_type=dossier_type,
        scope_vn="{{ noi_dung_yc }}", scope_en="{{ noi_dung_yc }}",
        location_desc_vn="{{ dia_diem }}", location_desc_en="{{ dia_diem }}",
        standards_vn="{{ can_cu }}", standards_en="{{ can_cu }}",
        procedure_vn="{{ quy_trinh }}", procedure_en="{{ quy_trinh }}",
        client=ClientInfo(name_vn="{{ kh_ten }}", name_en="{{ kh_ten }}",
                          address_vn="{{ kh_dia_chi }}", address_en="{{ kh_dia_chi }}"),
        personnel=PersonnelInfo(
            executor="{{ nguoi_th }}",
            sign_gdv_vn="{{ gdv_chuc_danh }}\n{{ gdv_ten }}", sign_gdv_en="{{ gdv_chuc_danh }}\n{{ gdv_ten }}",
            sign_leader_vn="{{ ld_chuc_danh_1 }}\n{{ ld_chuc_danh_2 }}\n{{ ld_ten }}",
            sign_leader_en="{{ ld_chuc_danh_1 }}\n{{ ld_chuc_danh_2 }}\n{{ ld_ten }}",
            reporter="{{ nguoi_bc }}",
        ),
        equipments=[Equipment(name_vn="{{ thiet_bi }}", name_en="{{ thiet_bi }}")],
    )
    d.results_pointer_vn = "{{ ket_qua_gd }}"
    d.results_pointer_en = "{{ ket_qua_gd }}"
    return d


def tag_cover(doc, is_en: bool):
    """Gắn TAG cho header (Số chứng thư/Ngày) và Table 0 của TBKQ (form 16 dòng)."""
    update_headers(doc, cert_no="{{ so_ct }}", cert_date="{{ ngay_ct }}", is_en=is_en)
    populate_cleanroom_table0(doc, tag_dossier(), is_en=is_en)
    t0 = doc.tables[0]
    eq_cell = t0.rows[10].cells[0]
    for p in eq_cell.paragraphs:
        replace_p(p, "- {{ thiet_bi }}.", "{{ thiet_bi }}")
        ppr = p._p.pPr
        if ppr is not None and ppr.find(qn("w:numPr")) is not None:
            ppr.remove(ppr.find(qn("w:numPr")))  # dòng thiết bị đã có "- " -> bỏ gạch đầu dòng tự động (tránh "- - ...")
    strip_highlight(t0._tbl)


# ===========================================================================
# TBKQ KHÍ NÉN (VN/EN)
# ===========================================================================
def build_khi_nen(lang: str):
    is_en = lang == "en"
    src_name = "TBKQ - en - khi nen.docx" if is_en else "TBKQ - vn - khi nen.docx"
    doc = docx.Document(src(src_name))
    strip_connectors(doc.element.body)
    tag_cover(doc, is_en)
    body = doc.element.body
    T = body_tables(doc)

    # --- Nội dung thực hiện ---
    p_num = find_p(doc, "The number of performance location" if is_en else "Số lượng vị trí thực hiện")
    from docx.text.paragraph import Paragraph
    set_paragraph_text(Paragraph(p_num, doc._body),
                       "The number of performance locations was indicated by {{ kh_ten }}\t: {{ so_vi_tri }}" if is_en
                       else "Số lượng vị trí thực hiện theo chỉ định của {{ kh_ten }}: {{ so_vi_tri }}", font_size=13)
    p_st1 = find_p(doc, "Status of compressed air system" if is_en else "Trạng thái hệ thống khí nén")
    set_p(p_st1, ("Status of compressed air system" if is_en else "Trạng thái hệ thống khí nén") + "\t: {{ trang_thai_he_thong }}")
    p_st2 = find_p(doc, "Status of the factory" if is_en else "Trạng thái nhà xưởng")
    set_p(p_st2, ("Status of the factory" if is_en else "Trạng thái nhà xưởng") + "\t: {{ trang_thai_nha_xuong }}")
    p_desc = find_p(doc, "Description of compressed air system" if is_en else "Mô tả hệ thống khí nén")
    p_chain = p_desc.getnext()
    while p_chain.tag != _P:
        p_chain = p_chain.getnext()
    # Sơ đồ hệ thống: TagBlock (các thiết bị nối bằng mũi tên Wingdings như bản gốc)
    for s in list(p_chain.iter(qn("w:sym"))):
        s.getparent().getparent().remove(s.getparent())
    for b in list(p_chain.iter(qn("w:bookmarkStart"))) + list(p_chain.iter(qn("w:bookmarkEnd"))):
        b.getparent().remove(b)
    set_p(p_chain, "{{ so_do_he_thong }}")
    wrap_p(p_desc, p_chain, "if so_do_he_thong")

    # --- Bảng 1: vị trí ---
    t1 = T[1]
    r1 = rows(t1)
    for tr in r1[2:]:
        t1.remove(tr)
    set_cell(cell(t1, 1, 0), "{{ vt.ten }}")
    set_cell(cell(t1, 1, 1), "{{ vt.ky_hieu }}")
    tc_ct = cell(t1, 1, 2)
    ps_ct = tc_ct.findall(_P)
    for x in ps_ct[1:]:
        tc_ct.remove(x)
    set_p(ps_ct[0], "{{ ct }}")
    wrap_p(ps_ct[0], ps_ct[0], "for ct in vt.chi_tieu")
    wrap_rows(r1[1], r1[1], "for vt in vi_tri")

    # --- Ảnh vị trí: TagBlock ---
    p_photo = find_p(doc, "Photo of performance location" if is_en else "Hình ảnh vị trí thực hiện")
    p_kq = find_p(doc, "Inspection result" if is_en else "Kết quả giám định", start_el=p_photo)
    remove_between(p_photo, p_kq)
    p_grid = new_p_like(p_photo, "{{ bang_anh }}")
    p_photo.addnext(p_grid)
    wrap_p(p_photo, p_grid, "if co_anh")

    # --- Hàm lượng tổng dầu (T3, T4) ---
    t3, t4 = T[3], T[4]
    r3 = rows(t3)
    set_cell(cell(t3, 1, 0), "{{ r.vi_tri }}")
    set_cell(cell(t3, 1, 1), "{{ r.gia_tri }}")
    set_cell(cell(t3, 1, 2), "{{ r.phan_loai }}")
    wrap_rows(r3[1], r3[1], "for r in dau.dong")
    set_cell(cell(t3, len(rows(t3)) - 1, 1), "{{ dau.yc }}")
    set_cell(cell(t3, len(rows(t3)) - 1, 2), "{{ dau.yc_phan_loai }}")
    last = rows(t3)[-1]
    wrap_rows(last, last, "if dau.yc", ref_tr=r3[1])
    r4 = rows(t4)
    set_cell(cell(t4, 2, 1), ("The requirement provided by {{ kh_ten }}." if is_en else "Mức yêu cầu được cung cấp bởi {{ kh_ten }}."))
    wrap_rows(r4[2], r4[2], "if dau.yc")
    h_oil = find_p(doc, "9.2.1 Total oil content" if is_en else "Hàm lượng tổng dầu", start_el=p_kq)
    a_oil = find_p(doc, "Assessment" if is_en else "Đánh giá", start_el=t4)
    s_oil = a_oil.getnext()
    set_p(s_oil, "{{ dau.danh_gia }}")
    wrap_p(a_oil, s_oil, "if dau.yc")
    wrap_p(h_oil, s_oil, "if dau")

    # --- Điểm sương (T5, T6) ---
    t5, t6 = T[5], T[6]
    r5 = rows(t5)
    set_cell(cell(t5, 2, 0), "{{ r.vi_tri }}")
    for c in range(1, 7):
        set_cell(cell(t5, 2, c), "{{ r.v[%d] }}" % (c - 1))
    wrap_rows(r5[2], r5[2], "for r in dp.dong")
    lastr = rows(t5)[-1]
    lt = lastr.findall(_TC)
    for c in range(1, len(lt)):
        set_cell(lt[c], "{{ dp.yc[%d] }}" % (c - 1))
    wrap_rows(lastr, lastr, "if dp.yc_co", ref_tr=r5[2])
    r6 = rows(t6)
    set_cell(cell(t6, 2, 1), ("The requirement provided by {{ kh_ten }}." if is_en else "Mức yêu cầu được cung cấp bởi {{ kh_ten }}."))
    wrap_rows(r6[2], r6[2], "if dp.yc_co")
    h_dp = find_p(doc, "9.2.2 Pressure dew point" if is_en else "Nhiệt độ điểm sương", start_el=s_oil)
    a_dp = find_p(doc, "Assessment" if is_en else "Đánh giá", start_el=t6)
    s_dp = a_dp.getnext()
    set_p(s_dp, "{{ dp.danh_gia }}")
    wrap_p(a_dp, s_dp, "if dp.yc_co")
    wrap_p(h_dp, s_dp, "if dp")

    # --- Nồng độ hạt (T7, T8) + yêu cầu d > 0,5 µm (T9, T10) ---
    t7, t8, t9, t10 = T[7], T[8], T[9], T[10]
    r7 = rows(t7)
    set_cell(cell(t7, 2, 0), "{{ r.vi_tri }}")
    for c in range(1, 4):
        set_cell(cell(t7, 2, c), "{{ r.v[%d] }}" % (c - 1))
    set_cell(cell(t7, 2, 4), "{{ r.phan_loai }}")
    # dòng 2 của mỗi vị trí: ô Vị trí và ô Phân loại là phần gộp dọc (vMerge) của dòng 1
    set_cell(cell(t7, 3, 0), "")
    set_cell(cell(t7, 3, 1), "{{ r.mo_ta_5um }}")
    set_cell(cell(t7, 3, 2), "")
    wrap_rows(r7[2], r7[3], "for r in hat.dong", ref_tr=r7[2])
    r8 = rows(t8)
    # (3) ... Không phân loại do xuất hiện hạt bụi > 5 µm: chỉ khi có
    wrap_rows(r8[2], r8[2], "if hat.co_khong_phan_loai")
    r9 = rows(t9)
    set_cell(cell(t9, 1, 0), "{{ r.vi_tri }}")
    set_cell(cell(t9, 1, 1), "{{ r.d05 }}")
    wrap_rows(r9[1], r9[1], "for r in hat.yc_dong")
    set_cell(cell(t9, len(rows(t9)) - 1, 1), "{{ hat.yc }}")
    r10 = rows(t10)
    set_cell(cell(t10, 1, 1), ("The requirement provided by {{ kh_ten }}." if is_en else "Mức yêu cầu được cung cấp bởi {{ kh_ten }}."))
    if len(r10) > 2:
        c10_2 = cell(t10, 2, 1)
        if "Nextern" in c10_2.text:
            set_cell(c10_2, c10_2.text.replace("Công ty TNHH Nextern Việt Nam", "{{ kh_ten }}").replace("Nextern Vietnam Limited Liability Company", "{{ kh_ten }}"))
    h_p = find_p(doc, "9.2.3 Concentration of particles" if is_en else "Nồng độ hạt bụi", start_el=s_dp)
    a_p = find_p(doc, "Assessment" if is_en else "Đánh giá", start_el=t10)
    s_p = a_p.getnext()
    set_p(s_p, "{{ hat.danh_gia }}")
    # khối yêu cầu: từ đoạn trống trước T9 đến câu đánh giá
    wrap_p(t9, s_p, "if hat.yc")
    wrap_p(h_p, s_p, "if hat")
    remove_empty_p_between(s_dp, h_p)

    # Bảng 13pt, Ghi chú 10pt, Ghi chú sát bảng (AGENTS mục 15; bản gốc khí nén lẫn 12 / 12,5 / 10,5 / 11pt — 07/11/2026)
    tighten_note_gaps(doc)
    normalize_table_font_sizes(doc)
    clear_negative_cell_indents(doc)
    fit_dewpoint_table(doc, lang)
    use_sectionpages(doc)
    out = dst("TBKQ - en - khi nen.docx" if is_en else "TBKQ - vn - khi nen.docx")
    doc.save(out)
    if not is_en:
        shutil.copy2(out, dst("TBKQ - viet - khi nen.docx"))
    return out


BUILDERS = {}


def register(name):
    def deco(fn):
        BUILDERS[name] = fn
        return fn
    return deco


@register("khi_nen")
def _b_khi_nen():
    return [build_khi_nen("vn"), build_khi_nen("en")]


# ===========================================================================
# TBKQ PHÒNG SẠCH (VN/EN) — form tổng hợp đầy đủ (master)
# ===========================================================================
def set_tbl_desc(tbl, name: str):
    """Đánh dấu bảng (w:tblDescription) để code tìm lại bảng sau khi điền (gộp ô khu vực...)."""
    tblPr = tbl.find(qn("w:tblPr"))
    for x in tblPr.findall(qn("w:tblDescription")):
        tblPr.remove(x)
    el = OxmlElement("w:tblDescription")
    el.set(qn("w:val"), name)
    tblPr.append(el)


def clear_vmerge(tr):
    for tc in _el(tr).findall(_TC):
        tcpr = tc.find(qn("w:tcPr"))
        if tcpr is not None:
            for x in tcpr.findall(qn("w:vMerge")):
                tcpr.remove(x)


def loop_rows(tbl, first_idx: int, n_rows_template: int, loop: str, cells: list):
    """Dòng first_idx thành dòng lặp; xóa các dòng mẫu thừa phía sau (n_rows_template - 1 dòng)."""
    rs = rows(tbl)
    row = rs[first_idx]
    for extra in rs[first_idx + 1:first_idx + n_rows_template]:
        tbl.remove(extra)
    clear_vmerge(row)
    tcs = row.findall(_TC)
    for i, text in enumerate(cells):
        if text is None:
            continue
        set_cell(tcs[i], text)
    wrap_rows(row, row, loop)
    return row


def replace_all(p, pairs):
    for old, new in pairs:
        if not replace_p(p, old, new):
            raise ValueError(f"Không thấy {old!r} trong đoạn {_p_text(_el(p))[:80]!r}")
    strip_highlight(_el(p))


DATA_ROW_TWIPS = 440   # hàng dữ liệu 1 dòng ~0,78 cm (AGENTS mục 9; người dùng yêu cầu 04/11/2026, hồ sơ 132809)
LIMIT_ROW_TWIPS = 580  # hàng giới hạn 2 dòng ~1,0 cm, atLeast (AGENTS mục 9; 127215 10/11/2026)


def set_row_height(tr, twips: int):
    """Đặt chiều cao tối thiểu (atLeast) cho một hàng bảng."""
    trpr = tr.find(qn("w:trPr"))
    if trpr is None:
        trpr = OxmlElement("w:trPr")
        tr.insert(0, trpr)
    th = trpr.find(qn("w:trHeight"))
    if th is None:
        th = OxmlElement("w:trHeight")
        trpr.append(th)
    th.set(qn("w:val"), str(twips))
    th.set(qn("w:hRule"), "atLeast")


def tighten_note_gaps(doc):
    """Đoạn trống ngăn cách bảng kết quả và bảng 'Ghi chú:'/'Note:' -> cao 1 pt, khoảng cách 0 (bảng Ghi chú nằm sát
    bảng kết quả; người dùng yêu cầu 04/11/2026, hồ sơ 132809). Không xóa đoạn: 2 bảng liền nhau Word gộp làm một."""
    body = doc.element.body
    for tbl in body.iterchildren(qn("w:tbl")):
        if not "".join(t.text or "" for t in tbl.iter(qn("w:t"))).strip().startswith(("Ghi chú", "Note")):
            continue
        gap = tbl.getprevious()
        if gap is None or gap.tag != qn("w:p") or _p_text(gap).strip() or gap.getprevious() is None                 or gap.getprevious().tag != qn("w:tbl"):
            continue
        ppr = gap.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            gap.insert(0, ppr)
        for x in ppr.findall(qn("w:spacing")):
            ppr.remove(x)
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:before"), "0")
        sp.set(qn("w:after"), "0")
        sp.set(qn("w:line"), "20")
        sp.set(qn("w:lineRule"), "exact")
        anchor_el = next((x for x in ppr if x.tag in (qn("w:ind"), qn("w:jc"), qn("w:rPr"))), None)  # đúng thứ tự schema
        if anchor_el is not None:
            anchor_el.addprevious(sp)
        else:
            ppr.append(sp)


_AFTER_SZ = {qn("w:" + t) for t in ("highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
                                     "em", "lang", "eastAsianLayout", "specVanish", "oMath")}


def _set_rpr_size(rpr, half_pts: int):
    for tag in ("w:sz", "w:szCs"):
        for x in rpr.findall(qn(tag)):
            rpr.remove(x)
    anchor = next((x for x in rpr if x.tag in _AFTER_SZ), None)   # đúng thứ tự schema CT_RPr
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(half_pts))
        if anchor is not None:
            anchor.addprevious(el)
        else:
            rpr.append(el)


def normalize_table_font_sizes(doc, body_pt: int = 13, note_pt: int = 11):
    """Mọi chữ trong bảng thân văn bản 13pt, bảng 'Ghi chú:'/'Note:' 10pt (AGENTS mục 3). Bản gốc phòng sạch lẫn
    12 / 12,5 / 11pt (người dùng QC 05/11/2026, hồ sơ 132809). Ghi sz tường minh cho cả run và dấu đoạn."""
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        is_note = "".join(t.text or "" for t in tbl.iter(qn("w:t"))).strip().startswith(("Ghi chú", "Note"))
        hp = (note_pt if is_note else body_pt) * 2
        for r in tbl.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            _set_rpr_size(rpr, hp)
        for ppr in tbl.iter(qn("w:pPr")):
            rpr = ppr.find(qn("w:rPr"))
            if rpr is not None:
                _set_rpr_size(rpr, hp)


def set_table_col_pct(tbl, pct: list):
    """Đặt độ rộng từng cột lưới (đơn vị 1/50 %, tổng 5000) cho tblGrid và mọi tcW (tính theo gridSpan)."""
    assert sum(pct) == 5000, pct
    gcols = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    assert len(gcols) == len(pct), (len(gcols), pct)
    total = sum(int(g.get(qn("w:w"))) for g in gcols)
    for g, p in zip(gcols, pct):
        g.set(qn("w:w"), str(round(total * p / 5000)))
    for tr in tbl.findall(_TR):
        pos = 0
        for tc in tr.findall(_TC):
            tcpr = tc.find(qn("w:tcPr"))
            gs = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
            span = int(gs.get(qn("w:val"))) if gs is not None else 1
            tcw = tcpr.find(qn("w:tcW")) if tcpr is not None else None
            if tcw is not None:
                tcw.set(qn("w:w"), str(sum(pct[pos:pos + span])))
                tcw.set(qn("w:type"), "pct")
            pos += span


# Bảng điểm sương khí nén ở 13pt: cột 'Vị trí thực hiện (Thời gian đo)' ~20 % để '(15:08 – 15:21)' nằm 1 dòng
DEWPOINT_COL_PCT = {"vn": [1020, 593, 576, 745, 713, 727, 626], "en": [1020, 823, 682, 640, 606, 608, 621]}


# Bảng tổng dầu: cột 1 ~20 % để 'Requirement ⁽²⁾' / 'Mức yêu cầu ⁽²⁾' nằm 1 dòng ở 13pt
OIL_COL_PCT = [1000, 1900, 2100]


def fit_dewpoint_table(doc, lang: str):
    """Chia lại cột bảng điểm sương và bảng tổng dầu khí nén (07/11/2026, sau khi lên 13pt)."""
    done = False
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        txt = "".join(t.text or "" for t in tbl.iter(qn("w:t")))
        if txt.strip().startswith(("Ghi chú", "Note")):
            continue
        if "Điều kiện đo" in txt or "Measurement condition" in txt:
            set_table_col_pct(tbl, DEWPOINT_COL_PCT["en" if lang == "en" else "vn"])
            done = True
        elif ("Hàm lượng tổng dầu" in txt or "Total oil content" in txt) and \
                len(tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))) == 3:
            set_table_col_pct(tbl, OIL_COL_PCT)
        elif ("d > 0,5" in txt or "d > 0.5" in txt) and len(tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))) == 2:
            set_table_col_pct(tbl, [1000, 4000])           # bảng hạt d > 0,5 µm theo mức yêu cầu
    return done


def clear_negative_cell_indents(doc):
    """Đoạn trong ô bảng kết quả có thụt lề âm (form gốc khí nén -56 … -207 twips để nhét chữ 12pt) -> 0: lên 13pt chữ
    lấn lề ô, chạm viền ('Nhiệt độ điểm', 'Purity class'). Không áp bảng Ghi chú (07/11/2026, khí nén 133266)."""
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        if "".join(t.text or "" for t in tbl.iter(qn("w:t"))).strip().startswith(("Ghi chú", "Note")):
            continue
        for ind in tbl.iter(qn("w:ind")):
            for side in ("left", "right", "start", "end"):
                v = ind.get(qn("w:" + side))
                if v is not None and v.lstrip("-").isdigit() and int(v) < 0:
                    ind.set(qn("w:" + side), "0")


def fit_bui6_table(tbl, first_pct: int = 860, cell_mar: int = 57):
    """Bảng bụi 6 kênh ở 13pt: tiêu đề cỡ hạt 'd ≥ 0,1µm' xuống 2 dòng ('d ≥' / '0,1µm'), chia lại độ rộng cột
    (2 cột đầu first_pct/5000, 6 cột cỡ hạt chia đều phần còn lại), lề ô trái/phải cell_mar twips. Layout cố định
    12 %/cột làm chữ tràn viền và EN gãy 'Measuremen t' (người dùng QC 05/11/2026, hồ sơ 132809)."""
    trs = tbl.findall(_TR)
    # 1) Tiêu đề cỡ hạt -> 2 dòng
    for tc in trs[1].findall(_TC):
        for t in tc.iter(_T):
            if t.text and "≥" in t.text:
                i = t.text.index("≥") + 1
                rest = t.text[i:].lstrip()
                t.text = t.text[:i]
                r = t.getparent()
                nr = copy.deepcopy(r)
                for x in list(nr):
                    if x.tag != qn("w:rPr"):
                        nr.remove(x)
                nr.append(OxmlElement("w:br"))
                nt = OxmlElement("w:t")
                nt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                nt.text = rest
                nr.append(nt)
                r.addnext(nr)
                break
    # 2) Độ rộng cột
    grid = tbl.find(qn("w:tblGrid"))
    gcols = grid.findall(qn("w:gridCol"))
    total = sum(int(g.get(qn("w:w"))) for g in gcols)
    rest_pct = 5000 - 2 * first_pct
    pct = [first_pct, first_pct] + [rest_pct // 6 + (1 if k < rest_pct % 6 else 0) for k in range(6)]
    for g, p in zip(gcols, pct):
        g.set(qn("w:w"), str(round(total * p / 5000)))
    for tr in trs:
        pos = 0
        for tc in tr.findall(_TC):
            tcpr = tc.find(qn("w:tcPr"))
            gs = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
            span = int(gs.get(qn("w:val"))) if gs is not None else 1
            tcw = tcpr.find(qn("w:tcW")) if tcpr is not None else None
            if tcw is not None:
                tcw.set(qn("w:w"), str(sum(pct[pos:pos + span])))
                tcw.set(qn("w:type"), "pct")
            pos += span
    # 3) Lề ô trái/phải
    tblpr = tbl.find(qn("w:tblPr"))
    for x in tblpr.findall(qn("w:tblCellMar")):
        tblpr.remove(x)
    mar = OxmlElement("w:tblCellMar")
    for side in ("left", "right"):
        m = OxmlElement("w:" + side)
        m.set(qn("w:w"), str(cell_mar))
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    after = {qn("w:tblLook"), qn("w:tblCaption"), qn("w:tblDescription")}
    anchor = next((x for x in tblpr if x.tag in after), None)
    if anchor is not None:
        anchor.addprevious(mar)
    else:
        tblpr.append(mar)


def build_phong_sach(lang: str):
    is_en = lang == "en"
    doc = docx.Document(src("TBKQ - en - phong sach.docx" if is_en else "TBKQ - viet - phong sach.docx"))
    strip_connectors(doc.element.body)
    tag_cover(doc, is_en)
    T = body_tables(doc)
    E = (lambda en, vn: en if is_en else vn)
    PH = "………"

    # --- Nội dung thực hiện ---
    p_kv = find_p(doc, E("The performance location was indicated by", "Khu vực thực hiện theo sự chỉ định của"))
    fill_hl(p_kv, ["{{ kh_ten }}"])
    p_ds = p_kv.getnext()
    while p_ds.tag != _P or not _p_text(p_ds).strip():
        p_ds = p_ds.getnext()
    set_p(p_ds, "{{ ds_khu_vuc }}")
    p_tt = find_p(doc, E("Cleanroom occupancy state", "Trạng thái phòng sạch"))
    fill_hl(p_tt, ["{{ trang_thai }}"])

    t1 = T[1]
    row = rows(t1)[1]
    tcs = row.findall(_TC)
    fill_hl(tcs[0].find(_P), ["{{ kv.ten }}"])
    fill_hl(tcs[1].find(_P), ["{{ kv.dien_tich }}"])
    fill_hl(tcs[2].find(_P), ["{{ kv.cap }}"])
    ps = tcs[3].findall(_P)
    for x in ps[1:]:
        tcs[3].remove(x)
    set_p(ps[0], "{{ ct }}")
    wrap_p(ps[0], ps[0], "for ct in kv.chi_tieu")
    wrap_rows(row, row, "for kv in khu_vuc")
    fill_hl(cell(T[2], 1, 1).find(_P), ["{{ kh_ten }}"])

    # --- Ảnh + sơ đồ (TagBlock) ---
    p_photo = find_p(doc, E("Photos of the performance process", "Hình ảnh quá trình thực hiện"))
    p_lay = find_p(doc, E("Layout of monitoring points", "Sơ đồ các vị trí quan trắc"))
    remove_between(p_photo, p_lay)
    g1 = new_p_like(p_photo, "{{ bang_anh }}")
    p_photo.addnext(g1)
    wrap_p(p_photo, g1, "if co_anh")
    p_res = find_p(doc, E("Measurement result of particle concentration (3", "Kết quả đo nồng độ hạt bụi (3"))
    remove_between(p_lay, p_res)
    g2 = new_p_like(p_lay, "{{ bang_so_do }}")
    p_lay.addnext(g2)
    wrap_p(p_lay, g2, "if co_so_do")

    def section(head_prefix, tbl, note, loop_var, data, has_assess=True, after=None):
        """Bọc 1 mục kết quả: {%p if data %} tiêu đề {%p for g in data %} bảng ... đánh giá {%p endfor %} {%p endif %}"""
        head = find_p(doc, head_prefix, start_el=after)
        if has_assess:
            a = find_p(doc, E("Assessment", "Đánh giá"), start_el=note)
            last = a.getnext()
        else:
            last = note
        return head, last

    # --- Bụi 3 kênh (T5/T6) & 6 kênh (T7/T8) ---
    for (tbl, note, nval, key, head_txt) in (
        (T[5], T[6], 3, "bui3", E("Measurement result of particle concentration (3", "Kết quả đo nồng độ hạt bụi (3")),
        (T[7], T[8], 6, "bui6", E("Measurement result of particle concentration (6", "Kết quả đo nồng độ hạt bụi (6")),
    ):
        set_tbl_desc(tbl, "tag:" + key)
        loop_rows(tbl, 2, 2, "for r in g.diem", ["{{ r.khu_vuc }}", "{{ r.vi_tri }}"] + ["{{ r.v[%d] }}" % i for i in range(nval)])
        if nval == 6:
            fit_bui6_table(tbl)
        lim = rows(tbl)[-1]
        set_row_height(lim, LIMIT_ROW_TWIPS)          # bản gốc ép 1020/1247 twips -> khoảng trắng cuối trang (127215)
        ltcs = lim.findall(_TC)
        fill_hl(ltcs[0].find(_P), ["{{ g.cap }}"])
        for i in range(nval):
            set_cell(ltcs[1 + i], "{{ g.gioi_han[%d] }}" % i)
        fill_hl(cell(note, 1, 1).find(_P), ["{{ g.trang_so_do }}"])
        head, last = section(head_txt, tbl, note, "g", key)
        replace_all(last, [("“………”", "{{ g.ds_khu_vuc }}"), (E(" complied ", " phù hợp "), " {{ g.ket_luan }} "), ("ISO Class ………" if is_en else "ISO ………", "ISO Class {{ g.cap }}" if is_en else "ISO {{ g.cap }}")])
        wrap_p(tbl, last, "for g in " + key)
        wrap_p(head, last.getnext(), "if " + key)

    # --- Vi sinh (T9/T10) ---
    tbl, note = T[9], T[10]
    set_tbl_desc(tbl, "tag:vi_sinh")
    loop_rows(tbl, 2, 2, "for r in g.diem", ["{{ r.khu_vuc }}", "{{ r.vi_tri }}", "{{ r.gia_tri }}"])
    set_row_height(rows(tbl)[-1], LIMIT_ROW_TWIPS)    # bản gốc ép 1247 twips cho hàng giới hạn 2 dòng (127215)
    lim = rows(tbl)[-1].findall(_TC)
    fill_hl(lim[0].find(_P), ["{{ g.cap }}"])
    fill_hl(lim[1].find(_P), ["{{ g.gioi_han }}"])
    fill_hl(cell(note, 1, 1).find(_P), ["{{ g.trang_so_do }}"])
    head = find_p(doc, E("Determination result of microbial", "Kết quả xác định mức độ nhiễm bẩn vi sinh vật"))
    wrap_p(tbl, note, "for g in vi_sinh")
    wrap_p(head, note.getnext(), "if vi_sinh")

    # --- Nhiệt độ - độ ẩm - độ ồn (T11/T12) ---
    tbl, note = T[11], T[12]
    loop_rows(tbl, 1, 1, "for r in nhiet_am.dong", ["{{ r.khu_vuc }}", "{{ r.nhiet_do }}", "{{ r.do_am }}", "{{ r.do_on }}"])
    set_row_height(rows(tbl)[2], DATA_ROW_TWIPS)      # bản gốc ép 1361 twips cho hàng 1 dòng -> cao thừa
    lim = rows(tbl)[-1].findall(_TC)
    set_cell(lim[1], "{{ nhiet_am.muc_nhiet_do }}")
    set_cell(lim[2], "{{ nhiet_am.muc_do_am }}")
    set_cell(lim[3], "{{ nhiet_am.muc_do_on }}")
    fill_hl(cell(note, 1, 1).find(_P), ["{{ kh_ten }}"])
    head = find_p(doc, E("Measurement result of temperature", "Kết quả đo nhiệt độ - độ ẩm"))
    a = find_p(doc, E("Assessment", "Đánh giá"), start_el=note)
    last = a.getnext()
    replace_all(last, [("“………”", "{{ nhiet_am.ds_khu_vuc }}"), (E(" complied ", " phù hợp "), " {{ nhiet_am.ket_luan }} "), ("………", "{{ kh_ten }}")])
    wrap_p(head, last, "if nhiet_am")

    # --- Chênh áp (T13/T14) ---
    tbl, note = T[13], T[14]
    loop_rows(tbl, 1, 1, "for r in chenh_ap.dong", ["{{ r.khu_vuc_1 }}", None, "{{ r.khu_vuc_2 }}", "{{ r.gia_tri }}", "{{ r.muc }}"])
    set_row_height(rows(tbl)[2], DATA_ROW_TWIPS)      # bản gốc ép 1247/1338 twips
    fill_hl(cell(note, 1, 1).find(_P), ["{{ kh_ten }}"])
    head = find_p(doc, E("Measurement result of pressure differential", "Kết quả đo chênh lệch áp suất"))
    a = find_p(doc, E("Assessment", "Đánh giá"), start_el=note)
    last = a.getnext()
    replace_all(last, [(E(" complied ", " phù hợp "), " {{ chenh_ap.ket_luan }} "), ("………", "{{ kh_ten }}")])
    wrap_p(head, last, "if chenh_ap")

    # --- Độ rò rỉ HEPA (T15/T16) ---
    tbl, note = T[15], T[16]
    set_tbl_desc(tbl, "tag:hepa")
    loop_rows(tbl, 1, 2, "for r in g.diem", ["{{ r.khu_vuc }}", "{{ r.vi_tri }}", "{{ r.gia_tri }}"])
    # Mức rò rỉ HEPA theo nhóm (mặc định ≤ 0,01; lọc H13 ≤ 0,1 — CleanroomHandler.hepa_leak_limit, chốt 30/10/2026)
    for tc in tbl.iter(_TC):
        if "".join(t.text or "" for t in tc.iter(_T)).strip() in ("≤ 0,01", "≤ 0.01"):
            set_cell(tc, "{{ g.muc }}")
    fill_hl(cell(note, 1, 1).find(_P), ["{{ g.trang_so_do }}"])
    head = find_p(doc, E("Measurement result of HEPA filter leakage", "Kết quả đo độ rò rỉ lọc HEPA"))
    a = find_p(doc, E("Assessment", "Đánh giá"), start_el=note)
    last = a.getnext()
    replace_all(last, [("“………”", "{{ g.ds_khu_vuc }}"), (E(" complied ", " phù hợp "), " {{ g.ket_luan }} ")])
    wrap_p(tbl, last, "for g in hepa")
    wrap_p(head, last.getnext(), "if hepa")

    # --- Quét rò rỉ HEPA (Scanloc) (T17 sơ đồ, T18, T19) ---
    t17, tbl, note = T[17], T[18], T[19]
    head = find_p(doc, E("Measurement result of HEPA filter scan", "Kết quả quét rò rỉ màn lọc HEPA"))
    p_sd = find_p(doc, E("Layout of HEPA filter locations", "Sơ đồ vị trí các lọc HEPA"), start_el=head)
    fill_hl(p_sd, ["{{ scan.khu_vuc }}"])
    c17 = cell(t17, 0, 0)
    set_cell(c17, "{{ scan.so_do }}")
    set_tbl_desc(tbl, "tag:scan")
    loop_rows(tbl, 1, 2, "for r in scan.diem", ["{{ r.khu_vuc }}", "{{ r.vi_tri }}", "{{ r.hepa }}", "{{ r.gasket }}"])
    fill_hl(cell(note, 1, 1).find(_P), ["{{ scan.trang_so_do }}"])
    wrap_p(head, note, "if scan")

    # --- Lưu lượng khí cấp / ACH (T20/T21) ---
    tbl, note = T[20], T[21]
    set_tbl_desc(tbl, "tag:luu_luong")
    loop_rows(tbl, 1, 2, "for r in g.diem", ["{{ r.khu_vuc }}", "{{ r.vi_tri }}", "{{ r.luu_luong }}", "{{ r.tong }}", "{{ r.the_tich }}", "{{ r.ach }}"])
    lim = rows(tbl)[-1].findall(_TC)
    fill_hl(lim[0].find(_P), ["{{ kh_ten }}"])
    fill_hl(lim[1].find(_P), ["{{ g.muc_ach }}"])
    fill_hl(cell(note, 1, 1).find(_P), ["{{ g.trang_so_do }}"])
    fill_hl(cell(note, 2, 1).find(_P), ["{{ kh_ten }}"])
    head = find_p(doc, E("Measurement result of airflow rate", "Kết quả đo lưu lượng khí cấp"))
    a = find_p(doc, E("Assessment", "Đánh giá"), start_el=note)
    last = a.getnext()
    replace_all(last, [("“………”", "{{ g.ds_khu_vuc }}"), (E(" complied ", " phù hợp "), " {{ g.ket_luan }} "), ("………", "{{ kh_ten }}")])
    wrap_p(tbl, last, "for g in luu_luong")
    wrap_p(head, last.getnext(), "if luu_luong")

    # còn sót '………' nào trong thân văn bản -> báo lỗi để sửa builder
    left = [(_p_text(p)[:60]) for p in doc.element.body.iter(_P) if "………" in _p_text(p) and "{{" not in _p_text(p)]
    if left:
        print(f"[Cảnh báo] phong sach {lang}: còn chỗ trống chưa gắn TAG: {left}")
    tighten_note_gaps(doc)
    normalize_table_font_sizes(doc)
    use_sectionpages(doc)
    strip_highlight(doc.element.body)
    out = dst("TBKQ - en - phong sach.docx" if is_en else "TBKQ - viet - phong sach.docx")
    doc.save(out)
    if not is_en:
        shutil.copy2(out, dst("TBKQ - vn - phong sach.docx"))
    return out


@register("phong_sach")
def _b_phong_sach():
    return [build_phong_sach("vn"), build_phong_sach("en")]


# ===========================================================================
# TEM TỦ (tem kiểm tra dán tủ) — chữ nằm trong textbox
# ===========================================================================
def _set_runs(p, texts):
    """Ghi lần lượt các chuỗi vào các run có chữ của đoạn; run thừa để trống (giữ định dạng từng run)."""
    runs = [r for r in p.iter(_R) if r.find(_T) is not None and _owner_p_of(r) is p]
    for i, r in enumerate(runs):
        ts = r.findall(_T)
        for t in ts:
            t.text = ""
        if i < len(texts):
            ts[0].text = texts[i]
            ts[0].set(XML_SPACE, "preserve")
    if len(texts) > len(runs) and runs:
        ts = runs[-1].findall(_T)
        ts[0].text = (ts[0].text or "") + "".join(texts[len(runs):])


def _owner_p_of(node):
    el = node.getparent()
    while el is not None and el.tag != _P:
        el = el.getparent()
    return el


def _tag_seal_content_lines(doc):
    """
    Dòng "Nội dung kiểm tra" của tem (E25, 30/10/2026): thay các dòng chỉ tiêu ghi cứng
    ("-Inflow  -Downflow  -HEPA filter leak", "-Lighting level ...", "-Face velocity") bằng vòng lặp
        {%p for dong in noi_dung_kt %} / {{ dong }} / {%p endfor %}
    Mỗi phần tử `noi_dung_kt` là một dòng, các chỉ tiêu cách nhau bằng tab (tab stop của template giữ nguyên).
    Danh sách do code dựng theo chỉ tiêu đã thực hiện (tag_context.seal_content_lines).
    """
    body = doc.element.body
    for head in [p for p in body.iter(_P) if _p_text(p).strip().startswith("Nội dung kiểm tra")]:
        tc = head.getparent()
        sibs = list(tc)
        idx = sibs.index(head)
        items = []
        for el in sibs[idx + 1:]:
            if el.tag != _P:
                break
            txt = _p_text(el).strip()
            if not txt.startswith("-"):
                break
            items.append(el)
        if not items:
            continue
        base = items[0]
        # Giữ run đầu tiên (định dạng chữ 7pt thường), xóa các run/thẻ còn lại trong đoạn
        runs = [r for r in base if r.tag == _R and r.find(_T) is not None]
        keep = runs[0]
        for ch in list(base):
            if ch.tag != qn("w:pPr") and ch is not keep:
                base.remove(ch)
        for ch in list(keep):
            if ch.tag != qn("w:rPr"):
                keep.remove(ch)
        t = OxmlElement("w:t")
        t.text = "{{ dong }}"
        t.set(XML_SPACE, "preserve")
        keep.append(t)
        for extra in items[1:]:
            tc.remove(extra)

        def _marker(text):
            m = copy.deepcopy(base)
            for r in [r for r in m if r.tag == _R]:
                for t2 in r.findall(_T):
                    t2.text = text
            return m
        base.addprevious(_marker("{%p for dong in noi_dung_kt %}"))
        base.addnext(_marker("{%p endfor %}"))
    # Căn cứ sau "Nội dung kiểm tra:" -> TAG {{ can_cu_tem }} (30/10/2026: đo bụi thì thêm ISO 14644-1:2015).
    # Mặc định do code điền: "EN 12469:2000" (BSC/LAF), "ASHRAE 110, SEFA 1" (tủ hút — sửa lỗi chính tả ASHREA).
    for head in [p for p in body.iter(_P) if _p_text(p).strip().startswith("Nội dung kiểm tra")]:
        for t in head.iter(_T):
            if t.text and ("EN 12469" in t.text or "SEFA" in t.text or "ASHREA" in t.text or "ASHRAE" in t.text):
                t.text = "{{ can_cu_tem }}"
                t.set(XML_SPACE, "preserve")
                # Căn cứ (EN 12469, ISO 14644, ASHRAE...) cỡ 8pt như giá trị Tên thiết bị/SN/Chủng loại
                # (người dùng chốt 30/10/2026; tem gốc để 6pt)
                r = t.getparent()
                rpr = r.find(qn("w:rPr"))
                if rpr is None:
                    rpr = OxmlElement("w:rPr")
                    r.insert(0, rpr)
                for tag in ("w:sz", "w:szCs"):
                    el = rpr.find(qn(tag))
                    if el is None:
                        el = OxmlElement(tag)
                        rpr.append(el)
                    el.set(qn("w:val"), "16")
                break
    for t in body.iter(_T):
        if t.text and "ASHREA" in t.text:
            t.text = t.text.replace("ASHREA", "ASHRAE")


def _unwrap_content_controls(body):
    """
    Gỡ toàn bộ content control (w:sdt) trong tem, giữ lại nội dung bên trong; xóa run rỗng (30/10/2026).
    Lý do: bản gốc để serial / chủng loại / tên thiết bị trong content control; sau khi gắn TAG các control này
    rỗng -> Word hiện ô xám "Click or tap here to enter text." đè lên tem (lỗi người dùng báo ở tem 132918).
    """
    _SDT_CONTENT = qn("w:sdtContent")
    changed = True
    while changed:
        changed = False
        for sdt in list(body.iter(qn("w:sdt"))):
            parent = sdt.getparent()
            if parent is None:
                continue
            content = sdt.find(_SDT_CONTENT)
            idx = parent.index(sdt)
            if content is not None:
                for i, ch in enumerate(list(content)):
                    parent.insert(idx + i, ch)
            parent.remove(sdt)
            changed = True
            break
    for r in list(body.iter(_R)):
        kids = [c for c in r if c.tag != qn("w:rPr")]
        if kids and all(c.tag == _T and not (c.text or "") for c in kids):
            r.getparent().remove(r)


def _keep_two_seals(doc):
    """
    1 tủ = 2 tem, xếp DỌC ở cột 1 của lưới 2 dòng x 3 cột (người dùng chốt 30/10/2026) — giống tem tủ hút.
    Trang A4 ngang chứa tối đa 6 tem = 3 tủ (mỗi cột 1 tủ). Bản gốc `tem tu.docx` có 4 tem (2 tủ 059003/1, /2
    nằm ngang 2 cột): chỉ giữ tem ở cột đầu, các ô tem khác để trống (1 đoạn rỗng).
    """
    body = doc.element.body
    for tbl in [t for t in body if t.tag == _TBL]:
        for tr in tbl.findall(_TR):
            for ci, tc in enumerate(tr.findall(_TC)):
                has_seal = any(_p_text(p).strip().startswith("Số/ No") for p in tc.iter(_P))
                if has_seal and ci > 0:
                    for ch in list(tc):
                        if ch.tag != qn("w:tcPr"):
                            tc.remove(ch)
                    tc.append(OxmlElement("w:p"))


def _seal_page_margins(doc):
    """
    Lề trang tem nhỏ lại để 2 dòng tem (6 tem / 3 tủ) nằm gọn 1 trang A4 ngang (30/10/2026):
    lề trên/dưới 0,5 cm, trái/phải 1,0 cm; đoạn rỗng bắt buộc sau bảng thu về 1pt để không đẩy sang trang 2.
    """
    body = doc.element.body
    for sp in body.iter(qn("w:sectPr")):
        mar = sp.find(qn("w:pgMar"))
        if mar is not None:
            for k, v in (("top", "284"), ("bottom", "284"), ("left", "567"), ("right", "567"),
                         ("header", "284"), ("footer", "284")):
                mar.set(qn(f"w:{k}"), v)
    last = [el for el in body if el.tag == _P]
    if last and body[-2] is last[-1]:
        p = last[-1]
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            p.insert(0, ppr)
        for tag in ("w:spacing", "w:rPr"):
            old = ppr.find(qn(tag))
            if old is not None:
                ppr.remove(old)
        sp_el = OxmlElement("w:spacing")
        sp_el.set(qn("w:before"), "0"); sp_el.set(qn("w:after"), "0")
        sp_el.set(qn("w:line"), "20"); sp_el.set(qn("w:lineRule"), "exact")
        ppr.append(sp_el)
        rpr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "2"); rpr.append(sz)
        ppr.append(rpr)


SEAL_CONTENT_ROW_TWIPS = 1250   # ô "Nội dung kiểm tra": đủ 2 dòng căn cứ (EN + ISO) + 3 dòng chỉ tiêu


def _fix_seal_row_heights(doc, height: int = SEAL_CONTENT_ROW_TWIPS):
    """
    Mọi tem (BSC/LAF, tủ hút) cùng kích thước bảng (người dùng chốt 30/10/2026, để gom 3 tủ / trang):
    ô "Nội dung kiểm tra" đặt chiều cao CỐ ĐỊNH (exact) thay vì "tối thiểu" — không còn co giãn theo số dòng
    chỉ tiêu; bỏ đoạn rỗng thừa cuối ô (template tủ hút có 1 đoạn rỗng làm tem tủ hút cao hơn tem BSC).
    """
    for tr in doc.element.body.iter(_TR):
        # chỉ dòng của bảng tem bên trong (đoạn "Nội dung kiểm tra" là con trực tiếp của ô), không phải bảng khung ngoài
        if not any(_p_text(p).strip().startswith("Nội dung kiểm tra")
                   for tc in tr.findall(_TC) for p in tc.findall(_P)):
            continue
        for tc in tr.findall(_TC):
            ps = [c for c in tc if c.tag == _P]
            while len(ps) > 1 and not _p_text(ps[-1]).strip():
                tc.remove(ps.pop())
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = OxmlElement("w:trPr")
            tr.insert(0, trpr)
        th = trpr.find(qn("w:trHeight"))
        if th is None:
            th = OxmlElement("w:trHeight")
            trpr.append(th)
        th.set(qn("w:val"), str(height))
        th.set(qn("w:hRule"), "exact")


def build_tem_tu(src_name: str, dst_name: str):
    doc = docx.Document(src(src_name))
    _keep_two_seals(doc)
    _seal_page_margins(doc)
    _fix_seal_row_heights(doc)
    _tag_seal_content_lines(doc)
    date_re = re.compile(r"^\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{4}$")
    code_re = re.compile(r"^Q3\d{2}[A-Z]\d+")
    for p in doc.element.body.iter(_P):
        txt = _p_text(p).strip()
        if not txt:
            continue
        if txt.startswith("Số/ No:"):
            _set_runs(p, ["Số/ No: ", "{{ so_tem }}"])
        elif txt.startswith("Tên thiết bị:"):
            # Giữ cỡ chữ như tem gốc: nhãn "Tên thiết bị:" 6pt (run 1), tên thiết bị 8pt (run 3).
            # Bản trước (26/10) dồn cả vào run 2 (10pt) -> tên dài bị rớt xuống dòng 2 (sửa 30/10/2026).
            n_runs = len([r for r in p.iter(_R) if r.find(_T) is not None])
            _set_runs(p, ["Tên thiết bị: {{ ten_thiet_bi }}"] if n_runs == 1
                      else ["Tên thiết bị: ", "{{ ten_thiet_bi }}"] if n_runs == 2
                      else ["Tên thiết bị: ", "", "{{ ten_thiet_bi }}"])
        elif txt.startswith("SN/ID"):
            # Nhãn 6pt (run 1), giá trị 8pt (run 2) như tem gốc (người dùng chốt 30/10/2026)
            _set_runs(p, ["SN/ID         : ", "{{ serial }}"])
        elif txt.startswith("Chủng loại"):
            _set_runs(p, ["Chủng loại : ", "{{ model }}"])
        elif txt.startswith("Ngày kiểm tra:"):
            _set_runs(p, ["Ngày kiểm tra: ", "{{ ngay_kiem_tra }}"])
        elif date_re.match(txt):
            _set_runs(p, ["{{ ngay_kiem_tra_tiep }}"])
        elif code_re.match(txt):
            _set_runs(p, ["{{ so_tem }}"])
    _unwrap_content_controls(doc.element.body)
    strip_highlight(doc.element.body)
    out = dst(dst_name)
    doc.save(out)
    return out


@register("tem_tu")
def _b_tem_tu():
    return [build_tem_tu("tem tu.docx", "tem tu.docx"), build_tem_tu("tem tu - fumehood.docx", "tem tu - fumehood.docx")]


# ===========================================================================
# BCKQ (BM3030.03) — 4 form: khí nén, phòng sạch, BSC/LAF, tủ hút
# ===========================================================================
def _para(doc, el):
    from docx.text.paragraph import Paragraph
    return Paragraph(el, doc._body)


def tag_bckq_table0(doc, is_kn: bool):
    from ..Builders.bckq_builder import populate_bckq_table0
    td = tag_dossier()
    td.code = "{{ so_bc }}"
    td.sample_time_vn = "{{ tg_th }}"
    td.location_desc_vn = "{{ dia_diem }}"
    populate_bckq_table0(doc.tables[0], td, is_kn=is_kn)
    for p in doc.tables[0]._tbl.iter(_P):
        replace_p(p, "- {{ thiet_bi }}.", "{{ thiet_bi }}")
    strip_highlight(doc.tables[0]._tbl)


def tag_bckq_tail(doc):
    """Phần 'Báo cáo của GĐV về thời gian thực hiện' (chung cho mọi BCKQ)."""
    pairs = [
        ("Ngày kết thúc hồ sơ thực tế:", "Ngày kết thúc hồ sơ thực tế: {{ ngay_kt }}", {}),
        ("Trễ hạn:", "Trễ hạn: \t{{ tre_han }}", {}),
        ("Thời gian trễ hạn:", "Thời gian trễ hạn: {{ tg_tre_han }}", {}),
        ("Nguyên nhân trễ hạn:", "Nguyên nhân trễ hạn: {{ nn_tre_han }}", {}),
        ("Các bất thường trong quá trình giám định:", "Các bất thường trong quá trình giám định: {{ bat_thuong }}", {}),
        ("Ngày báo cáo:", "{{ tab_bc }}Ngày báo cáo: {{ ngay_bc }}", {"italic": True}),
        ("báo cáo:", "{{ tab_bc }}Người báo cáo: {{ nguoi_bc }}", {"bold": True}),
    ]
    start = find_p(doc, "Báo cáo của GĐV", contains=True)
    for key, text, fmt in pairs:
        el = find_p(doc, key, start_el=start, contains=True) if key != "báo cáo:" else None
        if key == "báo cáo:":
            el = next(e for e in doc.element.body.iter(_P) if "Người báo cáo" in _p_text(e).replace("  ", " "))
        if key == "Trễ hạn:":
            el = next(e for e in doc.element.body.iter(_P) if _p_text(e).strip().startswith("Trễ hạn"))
        para = _para(doc, el)
        set_paragraph_text(para, text, font_size=13, **fmt)
        if "báo cáo" in key.lower():
            para.paragraph_format.left_indent = Pt(0)
            para.paragraph_format.right_indent = Pt(0)


def build_bckq_khi_nen():
    doc = docx.Document(src("BCKQ - khi nen.docx"))
    strip_connectors(doc.element.body)
    tag_bckq_table0(doc, is_kn=True)
    T = body_tables(doc)
    p_cd = find_p(doc, "- Vị trí lấy mẫu, đo đạc khí nén", contains=True)
    replace_p(p_cd, "Công ty TNHH NEXTERN Việt Nam", "{{ kh_ten }}")
    # Bảng vị trí
    t1 = T[1]
    for tr in rows(t1)[2:]:
        t1.remove(tr)
    loop_rows(t1, 1, 1, "for vt in vi_tri", ["{{ vt.ten }}", "{{ vt.ky_hieu }}", "{{ vt.chi_tieu_dong }}"])
    # Sơ đồ hệ thống khí nén
    p_sd = find_p(doc, "- Sơ đồ hệ thống khí nén", contains=True)
    p_chain = p_sd.getnext()
    for x in list(p_chain.iter(qn("w:sym"))):
        x.getparent().getparent().remove(x.getparent())
    for b in list(p_chain.iter(qn("w:bookmarkStart"))) + list(p_chain.iter(qn("w:bookmarkEnd"))):
        b.getparent().remove(b)
    set_p(p_chain, "{{ so_do_he_thong }}")
    wrap_p(p_sd, p_chain, "if so_do_he_thong")
    # Ảnh vị trí lấy mẫu
    p_img = find_p(doc, "- Hình ảnh vị trí lấy mẫu", contains=True)
    p_kq = find_p(doc, "Kết quả đo đạc/ phân tích", contains=True, start_el=p_img)
    remove_between(p_img, p_kq)
    g = new_p_like(p_img, "{{ bang_anh }}")
    p_img.addnext(g)
    wrap_p(p_img, g, "if co_anh")
    T = body_tables(doc)
    # Bảng tổng hợp dầu (KQTN/KQTT)
    t3 = T[-1]
    loop_rows(t3, 1, 1, "for vt in vi_tri", ["{{ vt.ten }}", "{{ vt.ky_hieu }}", "{{ vt.kqtn }}", "{{ vt.kqtt }}"])
    h_oil = find_p(doc, "Hàm lượng tổng dầu", start_el=p_kq)
    wrap_p(h_oil, t3, "if co_dau")
    h_b = find_p(doc, "Nồng độ hạt bụi", start_el=t3)
    s_b = h_b.getnext()
    set_p(s_b, "Xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {{ ma_hs }} ngày {{ ngay_ct }}")
    wrap_p(h_b, s_b, "if co_hat")
    h_d = find_p(doc, "Nhiệt độ điểm sương", start_el=s_b)
    s_d = h_d.getnext()
    set_p(s_d, "Xem tại phiếu Tính toán kết quả nhiệt độ điểm sương số {{ ma_hs }} ngày {{ ngay_ct }}")
    wrap_p(h_d, s_d, "if co_dp")
    # Gửi mẫu thử nghiệm: một khối cho mỗi phiếu KQTN
    h_g = find_p(doc, "Gửi mẫu thử nghiệm", contains=True)
    p_a = find_p(doc, "a. Phiếu KQTN", contains=True, start_el=h_g)
    p_b = find_p(doc, "b. Phiếu KQTN", contains=True, start_el=p_a)
    h_bc = find_p(doc, "Báo cáo của GĐV", contains=True, start_el=p_b)
    remove_between(p_b.getprevious(), h_bc)  # bỏ khối b. (dùng vòng lặp)
    blk = [p_a]
    el = p_a.getnext()
    while el is not None and el is not h_bc:
        blk.append(el)
        el = el.getnext()
    texts = ["{{ k.stt }}. Phiếu KQTN số {{ k.so }} ngày {{ k.ngay }}", "Ngày gửi mẫu\t\t: {{ k.ngay_gui }}",
             "Ngày yêu cầu cấp KQTN\t: {{ k.ngay_yc }}", "Ngày nhận KQTN\t\t: {{ k.ngay_nhan }}",
             "KQTN có trễ hạn so với yêu cầu của GĐV hay không? \t{{ k.tre_han }}"]
    real = [e for e in blk if e.tag == _P and _p_text(e).strip()]
    for e, t in zip(real, texts):
        set_p(e, t)
    for e in blk:
        if e.tag == _P and not _p_text(e).strip():
            e.getparent().remove(e)
    wrap_p(real[0], real[len(texts) - 1], "for k in kqtn")
    wrap_p(h_g, real[len(texts) - 1].getnext(), "if kqtn")
    tag_bckq_tail(doc)
    use_page_fields(doc, "NUMPAGES", ("header",))
    out = dst("BCKQ - khi nen.docx")
    doc.save(out)
    return out


def build_bckq_phong_sach():
    doc = docx.Document(src("BCKQ - phong sach.docx"))
    strip_connectors(doc.element.body)
    tag_bckq_table0(doc, is_kn=False)
    T = body_tables(doc)
    p_kv = find_p(doc, "Khu vực thực hiện theo sự chỉ định của")
    fill_hl(p_kv, ["{{ kh_ten }}"])
    p_ds = p_kv.getnext()
    set_p(p_ds, "{{ ds_khu_vuc }}")
    fill_hl(find_p(doc, "Trạng thái phòng sạch"), ["{{ trang_thai }}"])
    row = rows(T[1])[1]
    tcs = row.findall(_TC)
    fill_hl(tcs[0].find(_P), ["{{ kv.ten }}"])
    fill_hl(tcs[1].find(_P), ["{{ kv.dien_tich }}"])
    fill_hl(tcs[2].find(_P), ["{{ kv.cap }}"])
    ps = tcs[3].findall(_P)
    for x in ps[1:]:
        tcs[3].remove(x)
    set_p(ps[0], "{{ ct }}")
    wrap_p(ps[0], ps[0], "for ct in kv.chi_tieu")
    wrap_rows(row, row, "for kv in khu_vuc")
    note = cell(T[2], 0, 1)
    for p in note.findall(_P):
        if "………" in _p_text(p):
            fill_hl(p, ["{{ kh_ten }}"])
    p_photo = find_p(doc, "Hình ảnh quá trình thực hiện")
    p_lay = find_p(doc, "Sơ đồ các vị trí đo")
    remove_between(p_photo, p_lay)
    g1 = new_p_like(p_photo, "{{ bang_anh }}")
    p_photo.addnext(g1)
    wrap_p(p_photo, g1, "if co_anh")
    p_kq = find_p(doc, "Kết quả đo đạc/ phân tích", contains=True)
    remove_between(p_lay, p_kq)
    g2 = new_p_like(p_lay, "{{ bang_so_do }}")
    p_lay.addnext(g2)
    wrap_p(p_lay, g2, "if co_so_do")
    # Tham chiếu kết quả: chỉ viện dẫn phiếu tính toán/ghi chép, KHÔNG viện dẫn TBKQ (AGENTS.md mục 7)
    p_dg = find_p(doc, "Đánh giá", start_el=p_kq)
    lines = []
    el = p_kq.getnext()
    while el is not None and el is not p_dg:
        lines.append(el)
        el = el.getnext()
    first = lines[0]
    for e in lines[1:]:
        e.getparent().remove(e)
    set_p(first, "{{ kq }}")
    strip_highlight(first)
    wrap_p(first, first, "for kq in tham_chieu")
    T = body_tables(doc)
    t5 = T[-1]
    row = rows(t5)[1]
    tcs = row.findall(_TC)
    set_cell(tcs[0], "{{ r.khu_vuc }}")
    set_cell(tcs[1], "{{ r.dat }}")
    set_cell(tcs[2], "{{ r.khong_dat }}")
    set_cell(tcs[3], "{{ r.can_cu }}")
    wrap_rows(row, row, "for r in danh_gia")
    tag_bckq_tail(doc)
    strip_highlight(doc.element.body)
    use_page_fields(doc, "NUMPAGES", ("header",))
    out = dst("BCKQ - phong sach.docx")
    doc.save(out)
    return out


def _bckq_cabinet_body(doc, kind: str):
    """Phần thân BCKQ tủ (BSC/LAF hoặc tủ hút): bảng thông tin tủ + chỉ tiêu, dòng tham chiếu, đánh giá."""
    T = body_tables(doc)
    t1 = T[1]
    for tr in rows(t1)[2:]:
        t1.remove(tr)
    row = loop_rows(t1, 1, 1, "for t in tu", ["{{ t.thong_tin }}", "{{ t.chi_tieu }}"])
    p0 = row.findall(_TC)[0].find(_P)
    ppr = p0.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p0.insert(0, ppr)
    for x in ppr.findall(qn("w:tabs")):
        ppr.remove(x)
    ppr.insert(0, parse_xml(f'<w:tabs xmlns:w="{W_NS}"><w:tab w:val="left" w:pos="1800"/></w:tabs>'))
    p_ref = find_p(doc, "xem chi tiết tại phiếu", contains=True) if kind == "bsc" else \
        next(e for e in doc.element.body.iter(_P) if "Xem chi tiết tại phiếu" in _p_text(e))
    set_p(p_ref, "{{ kq }}")
    wrap_p(p_ref, p_ref, "for kq in tham_chieu")
    # Đánh giá: khối dựng bằng code (phù hợp / không phù hợp / không đánh giá — người dùng chốt 30/10/2026),
    # BSC/LAF và tủ hút cùng một kiểu để gộp được BCKQ nhiều tủ vào 1 file.
    p_dg = find_p(doc, "- Đánh giá:", contains=True)
    set_p(p_dg, "{{ danh_gia }}")
    if kind != "bsc":
        remove(body_tables(doc)[2])   # bảng "Thông tin tủ / Chỉ tiêu đạt / không đạt / Căn cứ" của form gốc


def build_bckq_bsc():
    doc = docx.Document(src("BCKQ - bsc.docx"))
    strip_connectors(doc.element.body)
    tag_bckq_table0(doc, is_kn=False)
    _bckq_cabinet_body(doc, "bsc")
    tag_bckq_tail(doc)
    use_page_fields(doc, "NUMPAGES", ("header",))
    out = dst("BCKQ - bsc.docx")
    doc.save(out)
    return out


def build_bckq_fumehood():
    doc = docx.Document(src("BCKQ - fumehood.docx"))
    strip_connectors(doc.element.body)
    tag_bckq_table0(doc, is_kn=False)
    _bckq_cabinet_body(doc, "fumehood")
    tag_bckq_tail(doc)
    use_page_fields(doc, "NUMPAGES", ("header",))
    out = dst("BCKQ - fumehood.docx")
    doc.save(out)
    return out


@register("bckq")
def _b_bckq():
    return [build_bckq_khi_nen(), build_bckq_phong_sach(), build_bckq_bsc(), build_bckq_fumehood()]


# ===========================================================================
# EXCEL TÍNH TOÁN — chỉ gắn TAG vào ô thông tin hồ sơ/khách hàng (giữ nguyên công thức, ảnh, định dạng)
# ===========================================================================
XLSX_HEADER_SUBS = [
    (r"Q3\d{2}[A-Z]\d{6}(?:/\d+)?", "{{ so_ct }}"),
    (r"\d{1,2}/\d{1,2}/\d{4}", "{{ ngay_ct }}"),
]

XLSX_TAGS = {
    "TinhToanBSC_template.xlsx": {
        "Bảng tính": {"C1": "{{ ma_hs }}", "G1": "{{ so_phu }}", "I1": "{{ ngay_dk }}", "C2": "{{ ngay_gd }}",
                      "F2": "{{ gio_bd }}", "I2": "{{ gio_kt }}", "B3": "{{ ten_tu }}", "G3": "{{ nhan_hieu }}",
                      "B4": "{{ cap_tu }}", "G4": "{{ model }}", "B5": "{{ serial }}", "H5": "{{ ma_tb }}"},
        "TBKQ-VN": {"T11": "{{ kh_ten }}", "T13": "{{ kh_dia_chi }}", "T15": "{{ noi_dung_yc }}",
                    "U19": "{{ dia_diem }}", "T33": "{{ ngay_ct }}", "T40": "{{ gdv_ten }}"},
        "TBKQ - ENG": {"T9": "{{ kh_ten_en }}", "T11": "{{ kh_dia_chi_en }}", "T13": "{{ noi_dung_yc_en }}",
                       "U17": "{{ dia_diem_en }}", "T41": "{{ gdv_ten_en }}", "T57": "{{ ten_tu_en }}"},
    },
    "TinhToanFumehood_template.xlsx": {
        "Bảng tính": {"C1": "{{ ma_hs }}", "G1": "{{ so_phu }}", "I1": "{{ ngay_dk }}", "C2": "{{ ngay_gd }}",
                      "F2": "{{ gio_bd }}", "I2": "{{ gio_kt }}", "B3": "{{ ten_tu }}", "G3": "{{ nhan_hieu }}",
                      "B4": "{{ loai_tu }}", "G4": "{{ model }}", "B5": "{{ serial }}", "H5": "{{ ma_tb }}"},
        "Thong Tin Chung Thu": {"C9": "{{ kh_ten }}", "Q9": "{{ kh_ten_en }}", "C11": "{{ kh_dia_chi }}",
                                "Q11": "{{ kh_dia_chi_en }}", "C13": "{{ noi_dung_yc }}", "Q13": "{{ noi_dung_yc_en }}",
                                "D17": "{{ dia_diem }}", "Q17": "{{ dia_diem_en }}", "C28": "{{ ngay_ct }}",
                                "C40": "{{ gdv_ten }}", "Q56": "{{ ten_tu_en }}"},
        "TBKQ - ENG": {"AD18": "{{ dia_diem_en }}"},
        "TBKQ_EN_KhongDanhGia": {"AD18": "{{ dia_diem_en }}"},
    },
    "TinhToanKhiNen_template.xlsx": {
        "Tinh_Toan_Khi_Co_Ap_V6": {"H2": "{{ vi_tri_ten }}", "H3": "{{ ky_hieu }}", "I6": "{{ yc_bui }}",
                                   "I8": "{{ yc_nuoc }}", "I9": "{{ yc_dau }}"},
    },
}


@register("excel")
def _b_excel():
    outs = []
    for name, cells in XLSX_TAGS.items():
        out = dst(name)
        xlsx_set_cells_raw(src(name), out, cells, header_subs=XLSX_HEADER_SUBS)
        outs.append(out)
    return outs


def main(argv=None):
    ap = argparse.ArgumentParser(description="Dựng template dạng TAG từ bản gốc (_goc_truoc_TAG).")
    ap.add_argument("--only", help="Danh sách nhóm, phân cách dấu phẩy: " + ", ".join(BUILDERS))
    args = ap.parse_args(argv)
    names = [n.strip() for n in args.only.split(",")] if args.only else list(BUILDERS)
    for n in names:
        outs = BUILDERS[n]()
        for o in outs:
            print(f"[TAG template] {n}: {os.path.relpath(o, CORE_DIR)}")


if __name__ == "__main__":
    main()
