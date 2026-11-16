"""
DOCX Patcher & Structural Editor for QUATEST 3 (Room N2).
Replaces apply-all.ps1, set-cells.ps1, dup-row.ps1, del-row.ps1, del-tables.ps1,
overwrite-images.ps1, and fix-docpr-ids.ps1.
"""
import os
import sys
import copy
import json
import shutil
import zipfile
import argparse
from lxml import etree
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding="utf-8")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def set_run_highlight(r_el, color: str = "yellow"):
    """Set highlight on a <w:r> element, avoiding duplicate <w:highlight> tags (Bẫy 52)."""
    rPr = r_el.find(w_tag("rPr"))
    if rPr is None:
        rPr = etree.Element(w_tag("rPr"))
        r_el.insert(0, rPr)
    for old_hl in rPr.findall(w_tag("highlight")):
        rPr.remove(old_hl)
    hl = etree.SubElement(rPr, w_tag("highlight"))
    hl.set(w_tag("val"), color)


def remove_run_highlight(r_el):
    """Remove highlight from <w:r> and its parent <w:p><w:pPr><w:rPr> (Bẫy 23 & Bẫy 56)."""
    rPr = r_el.find(w_tag("rPr"))
    if rPr is not None:
        for old_hl in rPr.findall(w_tag("highlight")):
            rPr.remove(old_hl)
    p_el = r_el.getparent()
    if p_el is not None and p_el.tag == w_tag("p"):
        pPr = p_el.find(w_tag("pPr"))
        if pPr is not None:
            pr_rPr = pPr.find(w_tag("rPr"))
            if pr_rPr is not None:
                for old_hl in pr_rPr.findall(w_tag("highlight")):
                    pr_rPr.remove(old_hl)


def set_run_bold(r_el, bold: bool = True):
    """Set or remove bold (<w:b/> and <w:bCs/>) on a <w:r> element."""
    rPr = r_el.find(w_tag("rPr"))
    if rPr is None:
        if not bold:
            return
        rPr = etree.Element(w_tag("rPr"))
        r_el.insert(0, rPr)
    for tag in ("b", "bCs"):
        for old_b in rPr.findall(w_tag(tag)):
            rPr.remove(old_b)
    if bold:
        etree.SubElement(rPr, w_tag("b"))
        etree.SubElement(rPr, w_tag("bCs"))


def apply_plan_to_docx(docx_path: str, jobs: list[dict]):
    """
    Apply run-index patches to a .docx archive.
    Each job in jobs:
      {
        "entry": "word/document.xml" | "word/header1.xml" ...,
        "set": { "12": "new text", ... },
        "hl": [12, 15],
        "unhl": [20],
        "bold": [30],
        "unbold": [31]
      }
    """
    tmp_path = docx_path + ".tmp_patch"
    jobs_by_entry = {}
    for j in jobs:
        ent = j.get("entry", "word/document.xml")
        jobs_by_entry.setdefault(ent, []).append(j)

    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            raw = zin.read(item.filename)
            if item.filename in jobs_by_entry:
                root = etree.fromstring(raw)
                t_elems = root.findall(f".//{{{W_NS}}}t")
                for job in jobs_by_entry[item.filename]:
                    set_map = job.get("set", {})
                    hl_list = set(int(x) for x in job.get("hl", []))
                    unhl_list = set(int(x) for x in job.get("unhl", []))
                    bold_list = set(int(x) for x in job.get("bold", []))
                    unbold_list = set(int(x) for x in job.get("unbold", []))

                    set_ok = 0
                    for k_str, val in set_map.items():
                        idx = int(k_str)
                        if 0 <= idx < len(t_elems):
                            t_el = t_elems[idx]
                            t_el.text = str(val)
                            if str(val).startswith(" ") or str(val).endswith(" ") or "  " in str(val):
                                t_el.set(XML_SPACE, "preserve")
                            set_ok += 1
                        else:
                            print(f"  [WARN] {item.filename}: run index {idx} out of range (max {len(t_elems)-1})")
                    for idx in hl_list:
                        if 0 <= idx < len(t_elems):
                            r_el = t_elems[idx].getparent()
                            if r_el is not None:
                                set_run_highlight(r_el, "yellow")
                    for idx in unhl_list:
                        if 0 <= idx < len(t_elems):
                            r_el = t_elems[idx].getparent()
                            if r_el is not None:
                                remove_run_highlight(r_el)
                    for idx in bold_list:
                        if 0 <= idx < len(t_elems):
                            r_el = t_elems[idx].getparent()
                            if r_el is not None:
                                set_run_bold(r_el, True)
                    for idx in unbold_list:
                        if 0 <= idx < len(t_elems):
                            r_el = t_elems[idx].getparent()
                            if r_el is not None:
                                set_run_bold(r_el, False)
                    print(
                        f"  Patched {os.path.basename(docx_path)} [{item.filename}]: "
                        f"set={set_ok}/{len(set_map)}, hl={len(hl_list)}, unhl={len(unhl_list)}"
                    )
                raw = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            zout.writestr(item, raw)

    shutil.move(tmp_path, docx_path)


def set_cell_bullet_lines(tc_elm, lines: list[str], font_size_pt: float = 13.0):
    """
    Replace paragraphs in a cell with plain '- ' lines
    (no w:numPr, 0cm indent, Times New Roman 13pt, bold=False) per QUY-TAC-FONT-VA-THIET-BI.md Phần 2 & 3.
    """
    for p in list(tc_elm.findall(qn("w:p"))):
        tc_elm.remove(p)
    sz_val = str(int(round(font_size_pt * 2)))
    for line in lines:
        new_p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "0")
        ind.set(qn("w:right"), "0")
        ind.set(qn("w:firstLine"), "0")
        pPr.append(ind)

        rPr_p = OxmlElement("w:rPr")
        rFonts_p = OxmlElement("w:rFonts")
        rFonts_p.set(qn("w:ascii"), "Times New Roman")
        rFonts_p.set(qn("w:hAnsi"), "Times New Roman")
        rFonts_p.set(qn("w:cs"), "Times New Roman")
        sz_p = OxmlElement("w:sz")
        sz_p.set(qn("w:val"), sz_val)
        szCs_p = OxmlElement("w:szCs")
        szCs_p.set(qn("w:val"), sz_val)
        rPr_p.extend([rFonts_p, sz_p, szCs_p])
        pPr.append(rPr_p)
        new_p.append(pPr)

        new_r = OxmlElement("w:r")
        new_r.append(copy.deepcopy(rPr_p))
        t = OxmlElement("w:t")
        s_line = line.strip()
        t.text = s_line if s_line.startswith("- ") else f"- {s_line}"
        new_r.append(t)
        new_p.append(new_r)
        tc_elm.append(new_p)


def set_monitoring_points_cell(tc_elm, items: list[list[str]], tab_pos: int = 2150, font_size_half_pt: int = 26):
    """Set multiple left-aligned tab-separated lines in Column 3 of Table #1."""
    for p in list(tc_elm.findall(qn("w:p"))):
        tc_elm.remove(p)
    sz_str = str(font_size_half_pt)
    for label, val in items:
        p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "left")
        tab.set(qn("w:pos"), str(tab_pos))
        tabs.append(tab)
        pPr.append(tabs)

        rPr_p = OxmlElement("w:rPr")
        rFonts_p = OxmlElement("w:rFonts")
        rFonts_p.set(qn("w:ascii"), "Times New Roman")
        rFonts_p.set(qn("w:hAnsi"), "Times New Roman")
        rFonts_p.set(qn("w:cs"), "Times New Roman")
        sz_p = OxmlElement("w:sz")
        sz_p.set(qn("w:val"), sz_str)
        szCs_p = OxmlElement("w:szCs")
        szCs_p.set(qn("w:val"), sz_str)
        rPr_p.extend([rFonts_p, sz_p, szCs_p])
        pPr.append(rPr_p)
        p.append(pPr)

        r1 = OxmlElement("w:r")
        r1.append(copy.deepcopy(rPr_p))
        t1 = OxmlElement("w:t")
        t1.text = str(label)
        r1.append(t1)
        p.append(r1)

        r2 = OxmlElement("w:r")
        r2.append(copy.deepcopy(rPr_p))
        r2.append(OxmlElement("w:tab"))
        t2 = OxmlElement("w:t")
        t2.text = f": {val}"
        r2.append(t2)
        p.append(r2)

        tc_elm.append(p)


def set_cells_in_docx(docx_path: str, cells_spec: list[dict]):
    """
    Patch table cells by 0-based (table, row, col).
    Each item in cells_spec:
      {
        "table": 0, "row": 2, "col": 1,
        "text": "new value",           # optional: single/multi-line string
        "bullets": ["line 1", ...],    # optional: equipment/standards '- ' list
        "monitoring_points": [["Nồng độ hạt bụi", "21 vị trí"], ...], # optional
        "b": true,                     # optional: bold
        "unbold": true,                # optional: remove bold
        "hl": true,                    # optional: yellow highlight
        "unhl": true                   # optional: remove highlight
      }
    """
    doc = docx.Document(docx_path)
    for spec in cells_spec:
        t_idx = int(spec["table"])
        r_idx = int(spec["row"])
        c_idx = int(spec["col"])
        tbl = doc.tables[t_idx]
        cell = tbl.rows[r_idx].cells[c_idx]
        tc = cell._tc

        if "bullets" in spec:
            set_cell_bullet_lines(tc, spec["bullets"], spec.get("font_size_pt", 13.0))
        elif "monitoring_points" in spec:
            set_monitoring_points_cell(tc, spec["monitoring_points"])
        elif "text" in spec:
            new_text = str(spec["text"])
            t_elems = tc.findall(f".//{{{W_NS}}}t")
            # Remove stray tabs/breaks inside cell runs (Bẫy 37)
            for tag_rm in ("tab", "br", "cr"):
                for el_rm in tc.findall(f".//{{{W_NS}}}{tag_rm}"):
                    el_rm.getparent().remove(el_rm)

            if t_elems:
                t_elems[0].text = new_text
                if new_text.startswith(" ") or new_text.endswith(" "):
                    t_elems[0].set(XML_SPACE, "preserve")
                for extra_t in t_elems[1:]:
                    extra_t.text = ""
            else:
                # Empty cell: find first paragraph or create one, insert run with cloned or standard rPr
                p_el = tc.find(qn("w:p"))
                if p_el is None:
                    p_el = OxmlElement("w:p")
                    tc.append(p_el)
                r_el = OxmlElement("w:r")
                rPr = OxmlElement("w:rPr")
                # Check if paragraph has pPr/rPr to inherit font/size
                pPr = p_el.find(qn("w:pPr"))
                inherited = pPr.find(qn("w:rPr")) if pPr is not None else None
                if inherited is not None:
                    rPr = copy.deepcopy(inherited)
                else:
                    rFonts = OxmlElement("w:rFonts")
                    rFonts.set(qn("w:ascii"), "Times New Roman")
                    rFonts.set(qn("w:hAnsi"), "Times New Roman")
                    rFonts.set(qn("w:cs"), "Times New Roman")
                    sz = OxmlElement("w:sz")
                    sz.set(qn("w:val"), "26")
                    szCs = OxmlElement("w:szCs")
                    szCs.set(qn("w:val"), "26")
                    rPr.extend([rFonts, sz, szCs])
                r_el.append(rPr)
                t_el = OxmlElement("w:t")
                t_el.text = new_text
                if new_text.startswith(" ") or new_text.endswith(" "):
                    t_el.set(XML_SPACE, "preserve")
                r_el.append(t_el)
                p_el.append(r_el)

        # Apply bold / unbold / hl / unhl across all runs in the cell
        for r_el in tc.findall(f".//{{{W_NS}}}r"):
            if spec.get("b"):
                set_run_bold(r_el, True)
            if spec.get("unbold"):
                set_run_bold(r_el, False)
            if spec.get("hl"):
                set_run_highlight(r_el, "yellow")
            if spec.get("unhl"):
                remove_run_highlight(r_el)

    doc.save(docx_path)
    print(f"Updated {len(cells_spec)} cell(s) in {os.path.basename(docx_path)}")


def fix_docpr_ids_in_tree(root, start_id: int = 900100) -> int:
    """Ensure every wp:docPr and pic:cNvPr in root has a unique integer ID (fixes Bẫy 34)."""
    cur_id = start_id
    count = 0
    for drw in root.findall(f".//{{{W_NS}}}drawing"):
        for dp in drw.findall(f".//{{{WP_NS}}}docPr"):
            dp.set("id", str(cur_id))
            cur_id += 1
            count += 1
        for cnv in drw.findall(f".//{{{PIC_NS}}}cNvPr"):
            cnv.set("id", str(cur_id))
            cur_id += 1
    return count


def ensure_p_prop(p_elm, prop_name: str):
    pPr = p_elm.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elm.insert(0, pPr)
    if pPr.find(qn(prop_name)) is None:
        pPr.append(OxmlElement(prop_name))


def ensure_tr_cantsplit(tr_elm):
    trPr = tr_elm.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr_elm.insert(0, trPr)
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))


def ensure_tr_tblheader(tr_elm):
    trPr = tr_elm.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr_elm.insert(0, trPr)
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def set_table_rows_keep_next(tbl_elm, last_n_rows: int = None):
    rows = tbl_elm.findall(qn("w:tr"))
    for tr in rows:
        ensure_tr_cantsplit(tr)
    target_rows = rows if last_n_rows is None else rows[-last_n_rows:]
    for tr in target_rows:
        for p in tr.findall(f".//{qn('w:p')}"):
            ensure_p_prop(p, "w:keepNext")
            ensure_p_prop(p, "w:keepLines")


def apply_keep_with_next_rules(doc):
    """
    Enforce QUY-TAC-FONT-VA-THIET-BI.md Phần 5:
    Any result table with an accompanying conclusion ('Đánh giá:' / 'Assessment:' + conclusion text)
    must keep its conclusion on the same page as the table.
    Also repeat header rows (<w:tblHeader/>) for data tables.
    """
    body = doc.element.body
    children = list(body)
    n = len(children)

    tbl_count = 0
    for el in children:
        if el.tag == qn("w:tbl"):
            if tbl_count > 0:
                rows = el.findall(qn("w:tr"))
                for tr in rows:
                    ensure_tr_cantsplit(tr)
                # First cell text check
                first_cell_txt = "".join(rows[0].itertext()).strip() if rows else ""
                is_note_tbl = first_cell_txt.startswith("Ghi chú") or first_cell_txt.startswith("Note")
                has_pic = rows[0].find(f".//{qn('w:drawing')}") is not None if rows else False
                if not is_note_tbl and not has_pic and len(rows) > 1:
                    ensure_tr_tblheader(rows[0])
            tbl_count += 1

    for el in children:
        if el.tag == qn("w:p"):
            txt = "".join(el.itertext()).strip()
            has_num = el.find(f".//{qn('w:numPr')}") is not None
            # Tiêu đề "Kết quả …" (không kết thúc bằng dấu chấm). Câu kết luận VN "Kết quả đo … phù hợp …." KHÔNG
            # được keepNext (AGENTS mục 3: chỉ keepLines) — trước 07/11/2026 nó kéo cả mục sau, câu Đánh giá rớt trang.
            if has_num or txt.startswith("9.") or (txt.startswith("Kết quả") and not txt.endswith(".")):
                ensure_p_prop(el, "w:keepNext")
                ensure_p_prop(el, "w:keepLines")

    for idx, el in enumerate(children):
        if el.tag != qn("w:p"):
            continue
        txt = "".join(el.itertext()).strip()
        if txt.startswith("Đánh giá") or txt.startswith("Assessment"):
            ensure_p_prop(el, "w:keepNext")
            ensure_p_prop(el, "w:keepLines")
            j = idx + 1
            while j < n and children[j].tag == qn("w:p"):
                j_txt = "".join(children[j].itertext()).strip()
                if not j_txt:
                    ensure_p_prop(children[j], "w:keepNext")
                    j += 1
                    continue
                ensure_p_prop(children[j], "w:keepLines")
                break

            k = idx - 1
            seen_result_tbl = False
            while k >= 0:
                prev_el = children[k]
                if prev_el.tag == qn("w:p"):
                    p_txt = "".join(prev_el.itertext()).strip()
                    if not p_txt:
                        ensure_p_prop(prev_el, "w:keepNext")
                        k -= 1
                        continue
                    break
                elif prev_el.tag == qn("w:tbl"):
                    first_cell_txt = "".join(prev_el.itertext()).strip()
                    is_note_tbl = first_cell_txt.startswith("Ghi chú") or first_cell_txt.startswith("Note")
                    if is_note_tbl and not seen_result_tbl:
                        set_table_rows_keep_next(prev_el, last_n_rows=None)
                        k -= 1
                        continue
                    else:
                        rows = prev_el.findall(qn("w:tr"))
                        last_n = None if len(rows) <= 4 else 2
                        set_table_rows_keep_next(prev_el, last_n_rows=last_n)
                        seen_result_tbl = True
                        break
                else:
                    break


def dup_table_row(docx_path: str, table_idx: int, row_idx: int, count: int = 1):
    """Duplicate a row in table_idx count times, then re-index docPr IDs (Bẫy 34) and enforce keepNext."""
    doc = docx.Document(docx_path)
    tbl = doc.tables[table_idx]._tbl
    rows = tbl.findall(qn("w:tr"))
    target_tr = rows[row_idx]
    insert_pos = tbl.index(target_tr) + 1
    for i in range(count):
        new_tr = copy.deepcopy(target_tr)
        tbl.insert(insert_pos + i, new_tr)
    fix_docpr_ids_in_tree(doc.element.body)
    apply_keep_with_next_rules(doc)
    doc.save(docx_path)
    print(f"Duplicated Table #{table_idx} Row #{row_idx} x{count} in {os.path.basename(docx_path)}")


def del_table_rows(docx_path: str, table_idx: int, row_indices: list[int]):
    """Delete specified 0-based rows from table_idx."""
    doc = docx.Document(docx_path)
    tbl = doc.tables[table_idx]._tbl
    rows = tbl.findall(qn("w:tr"))
    for r_i in sorted(set(row_indices), reverse=True):
        if 0 <= r_i < len(rows):
            tbl.remove(rows[r_i])
    apply_keep_with_next_rules(doc)
    doc.save(docx_path)
    print(f"Deleted rows {row_indices} from Table #{table_idx} in {os.path.basename(docx_path)}")


def del_tables(docx_path: str, table_indices: list[int], with_heading: bool = False):
    """Delete specified 0-based tables (and optionally the heading paragraph right before each)."""
    doc = docx.Document(docx_path)
    body = doc.element.body
    tables = [doc.tables[i]._tbl for i in sorted(set(table_indices), reverse=True) if 0 <= i < len(doc.tables)]
    for tbl_el in tables:
        if with_heading:
            prev = tbl_el.getprevious()
            while prev is not None and prev.tag == qn("w:p"):
                txt = "".join(prev.itertext()).strip()
                to_del = prev
                prev = prev.getprevious()
                body.remove(to_del)
                if txt:
                    break
        body.remove(tbl_el)
    apply_keep_with_next_rules(doc)
    doc.save(docx_path)
    print(f"Deleted tables {table_indices} (with_heading={with_heading}) in {os.path.basename(docx_path)}")


def overwrite_images(docx_path: str, replacements: dict[str, str], strip_crop: bool = True):
    """
    Replace embedded images in docx_path with original image files without resampling (Bẫy 35).
    replacements keys can be:
      - media path: 'word/media/image1.jpeg' or 'image1.jpeg'
      - rId: 'rId8'
      - drawing index: 'drawing:0'
    Also verifies [Content_Types].xml has matching extension (Bẫy 6) and optionally strips <a:srcRect> crop (4a).
    """
    with zipfile.ZipFile(docx_path, "r") as zin:
        rels_map = {}
        if "word/_rels/document.xml.rels" in zin.namelist():
            rels_root = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
            for rel in rels_root:
                rels_map[rel.get("Id")] = rel.get("Target")

        doc_root = etree.fromstring(zin.read("word/document.xml"))
        drawings = doc_root.findall(f".//{{{W_NS}}}drawing")

        target_media_to_src = {}
        rIds_to_uncrop = set()
        for key, src_img_path in replacements.items():
            if not os.path.isfile(src_img_path):
                raise FileNotFoundError(f"Source image not found: {src_img_path}")
            if key.startswith("drawing:"):
                d_idx = int(key.split(":", 1)[1])
                blip = drawings[d_idx].find(f".//{{{A_NS}}}blip")
                r_id = blip.get(f"{{{R_NS}}}embed")
                t_rel = rels_map[r_id]
                m_path = f"word/{t_rel}" if not t_rel.startswith("word/") else t_rel
                target_media_to_src[m_path] = src_img_path
                rIds_to_uncrop.add(r_id)
            elif key.startswith("rId"):
                t_rel = rels_map[key]
                m_path = f"word/{t_rel}" if not t_rel.startswith("word/") else t_rel
                target_media_to_src[m_path] = src_img_path
                rIds_to_uncrop.add(key)
            else:
                m_path = key if key.startswith("word/media/") else f"word/media/{key}"
                target_media_to_src[m_path] = src_img_path
                for r_id, t_rel in rels_map.items():
                    full_t = f"word/{t_rel}" if not t_rel.startswith("word/") else t_rel
                    if full_t == m_path:
                        rIds_to_uncrop.add(r_id)

        # Strip <a:srcRect> crop on replaced images if requested
        doc_modified = False
        if strip_crop and rIds_to_uncrop:
            for drw in drawings:
                blip = drw.find(f".//{{{A_NS}}}blip")
                if blip is not None and blip.get(f"{{{R_NS}}}embed") in rIds_to_uncrop:
                    for sr in drw.findall(f".//{{{A_NS}}}srcRect"):
                        if sr.attrib:
                            sr.attrib.clear()
                            doc_modified = True

        # Ensure unique docPr IDs (Bẫy 34)
        fix_docpr_ids_in_tree(doc_root)
        doc_modified = True

        # Check [Content_Types].xml extensions (Bẫy 6)
        ct_root = etree.fromstring(zin.read("[Content_Types].xml"))
        existing_exts = {
            el.get("Extension", "").lower()
            for el in ct_root.findall(f"{{{CT_NS}}}Default")
        }
        ct_modified = False
        for m_path in target_media_to_src:
            ext = os.path.splitext(m_path)[1].lstrip(".").lower()
            if ext and ext not in existing_exts:
                mime = "image/png" if ext == "png" else "image/jpeg"
                d_el = etree.SubElement(ct_root, f"{{{CT_NS}}}Default")
                d_el.set("Extension", ext)
                d_el.set("ContentType", mime)
                existing_exts.add(ext)
                ct_modified = True

        tmp_path = docx_path + ".tmp_img"
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in target_media_to_src:
                    with open(target_media_to_src[item.filename], "rb") as f_img:
                        raw = f_img.read()
                    print(f"  Overwrote {item.filename} <- {os.path.basename(target_media_to_src[item.filename])} ({len(raw):,}B)")
                elif item.filename == "word/document.xml" and doc_modified:
                    raw = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
                elif item.filename == "[Content_Types].xml" and ct_modified:
                    raw = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
                else:
                    raw = zin.read(item.filename)
                zout.writestr(item, raw)

    shutil.move(tmp_path, docx_path)


def enforce_all_rules(docx_path: str):
    """Apply keepNext/keepLines/cantSplit rules, strip leftover Word comments (Bẫy 50), and fix docPr IDs (Bẫy 34)."""
    doc = docx.Document(docx_path)
    apply_keep_with_next_rules(doc)
    body = doc.element.body
    fix_docpr_ids_in_tree(body)
    # Strip commentRangeStart, commentRangeEnd, commentReference (Bẫy 50)
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for el in body.findall(f".//{{{W_NS}}}{tag}"):
            el.getparent().remove(el)
    doc.save(docx_path)
    print(f"Enforced keepNext/cantSplit, docPr IDs, and comment cleanup on {os.path.basename(docx_path)}")


def main():
    parser = argparse.ArgumentParser(description="DOCX Patcher for QUATEST 3")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_plan = sub.add_parser("apply-plan", help="Apply run-index plan JSON")
    p_plan.add_argument("plan_json", help="Path to plan.json")

    p_cells = sub.add_parser("set-cells", help="Patch table cells from JSON")
    p_cells.add_argument("docx", help="Path to .docx file")
    p_cells.add_argument("cells_json", help="Path to cells JSON file")

    p_dup = sub.add_parser("dup-row", help="Duplicate a table row")
    p_dup.add_argument("docx", help="Path to .docx file")
    p_dup.add_argument("--table", type=int, required=True)
    p_dup.add_argument("--row", type=int, required=True)
    p_dup.add_argument("--count", type=int, default=1)

    p_delr = sub.add_parser("del-rows", help="Delete table rows")
    p_delr.add_argument("docx", help="Path to .docx file")
    p_delr.add_argument("--table", type=int, required=True)
    p_delr.add_argument("--rows", required=True, help="Comma-separated 0-based row indices")

    p_delt = sub.add_parser("del-tables", help="Delete tables")
    p_delt.add_argument("docx", help="Path to .docx file")
    p_delt.add_argument("--tables", required=True, help="Comma-separated 0-based table indices")
    p_delt.add_argument("--with-heading", action="store_true")

    p_img = sub.add_parser("overwrite-images", help="Overwrite embedded images with raw bytes")
    p_img.add_argument("docx", help="Path to .docx file")
    p_img.add_argument("map_json", help="Path to JSON mapping media/rId/drawing:N -> image file path")

    p_enf = sub.add_parser("enforce-rules", help="Enforce keepNext/cantSplit, unique docPr IDs, strip comments")
    p_enf.add_argument("docx", help="Path to .docx file")

    args = parser.parse_args()
    if args.cmd == "apply-plan":
        with open(args.plan_json, "r", encoding="utf-8") as f:
            plan = json.load(f)
        by_file = {}
        for job in plan.get("jobs", []):
            by_file.setdefault(job["file"], []).append(job)
        for fpath, jobs in by_file.items():
            apply_plan_to_docx(fpath, jobs)
    elif args.cmd == "set-cells":
        with open(args.cells_json, "r", encoding="utf-8") as f:
            data = json.load(f)
        cells = data["cells"] if isinstance(data, dict) and "cells" in data else data
        set_cells_in_docx(args.docx, cells)
    elif args.cmd == "dup-row":
        dup_table_row(args.docx, args.table, args.row, args.count)
    elif args.cmd == "del-rows":
        rows = [int(x.strip()) for x in args.rows.split(",") if x.strip()]
        del_table_rows(args.docx, args.table, rows)
    elif args.cmd == "del-tables":
        tbls = [int(x.strip()) for x in args.tables.split(",") if x.strip()]
        del_tables(args.docx, tbls, args.with_heading)
    elif args.cmd == "overwrite-images":
        with open(args.map_json, "r", encoding="utf-8") as f:
            mp = json.load(f)
        overwrite_images(args.docx, mp)
    elif args.cmd == "enforce-rules":
        enforce_all_rules(args.docx)


if __name__ == "__main__":
    main()
