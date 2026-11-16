import os
import copy
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"
dir_2026 = os.path.join(base_dir, "New From 2026")
dir_2025 = os.path.join(base_dir, "Old Form 2025")

MC_AC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"

def strip_drawings(elm):
    """Remove w:drawing, w:pict, mc:AlternateContent (containing drawings/lines) from an element copy."""
    for ac in elm.findall(f".//{MC_AC}"):
        if ac.find(f".//{qn('w:drawing')}") is not None or ac.find(f".//{qn('w:pict')}") is not None:
            ac.getparent().remove(ac)
    for d in elm.findall(f".//{qn('w:drawing')}"):
        d.getparent().remove(d)
    for p in elm.findall(f".//{qn('w:pict')}"):
        p.getparent().remove(p)

def set_num_id(p_elm, num_id_str, ilvl_str=None):
    """Ensure paragraph numPr uses the specified numId and optionally ilvl."""
    numPr = p_elm.find(f".//{qn('w:numPr')}")
    if numPr is not None:
        numId = numPr.find(qn('w:numId'))
        if numId is not None:
            numId.set(qn('w:val'), str(num_id_str))
        if ilvl_str is not None:
            ilvl = numPr.find(qn('w:ilvl'))
            if ilvl is not None:
                ilvl.set(qn('w:val'), str(ilvl_str))

def replace_text_in_element(elm, old_str, new_str):
    for t in elm.findall(f".//{qn('w:t')}"):
        if t.text and old_str in t.text:
            t.text = t.text.replace(old_str, new_str)

def set_cell_bullet_lines(tc_elm, lines):
    """Replace paragraphs in a cell with plain '- ' lines (no w:numPr, 0cm indent, Times New Roman 13pt, bold=False)."""
    ps = tc_elm.findall(qn('w:p'))
    for p in ps:
        tc_elm.remove(p)
    for line in lines:
        new_p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), '0')
        ind.set(qn('w:right'), '0')
        ind.set(qn('w:firstLine'), '0')
        pPr.append(ind)
        rPr_p = OxmlElement('w:rPr')
        rFonts_p = OxmlElement('w:rFonts')
        rFonts_p.set(qn('w:ascii'), 'Times New Roman')
        rFonts_p.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts_p.set(qn('w:cs'), 'Times New Roman')
        sz_p = OxmlElement('w:sz')
        sz_p.set(qn('w:val'), '26') # 13pt
        szCs_p = OxmlElement('w:szCs')
        szCs_p.set(qn('w:val'), '26')
        rPr_p.extend([rFonts_p, sz_p, szCs_p])
        pPr.append(rPr_p)
        new_p.append(pPr)

        new_r = OxmlElement('w:r')
        new_r.append(copy.deepcopy(rPr_p))
        t = OxmlElement('w:t')
        text_val = line if line.startswith("- ") else f"- {line}"
        t.text = text_val
        new_r.append(t)
        new_p.append(new_r)
        tc_elm.append(new_p)

def set_monitoring_points_cell(tc_elm, items):
    """Set multiple left-aligned tab-separated lines in Column 3 of Table #1."""
    ps = tc_elm.findall(qn('w:p'))
    for p in ps:
        tc_elm.remove(p)
    for label, val in items:
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        tabs = OxmlElement('w:tabs')
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'left')
        tab.set(qn('w:pos'), '2150')
        tabs.append(tab)
        pPr.append(tabs)
        rPr_p = OxmlElement('w:rPr')
        sz_p = OxmlElement('w:sz')
        sz_p.set(qn('w:val'), '25')
        szCs_p = OxmlElement('w:szCs')
        szCs_p.set(qn('w:val'), '25')
        rPr_p.extend([sz_p, szCs_p])
        pPr.append(rPr_p)
        p.append(pPr)

        r1 = OxmlElement('w:r')
        rPr1 = copy.deepcopy(rPr_p)
        r1.append(rPr1)
        t1 = OxmlElement('w:t')
        t1.text = label
        r1.append(t1)
        p.append(r1)

        r2 = OxmlElement('w:r')
        rPr2 = copy.deepcopy(rPr_p)
        r2.append(rPr2)
        r2.append(OxmlElement('w:tab'))
        t2 = OxmlElement('w:t')
        t2.text = f": {val}"
        r2.append(t2)
        p.append(r2)

        tc_elm.append(p)

def make_label_paragraph(text):
    """Create a clean italic note/separator paragraph to label alternative template tables."""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '120')
    spacing.set(qn('w:after'), '60')
    pPr.append(spacing)
    p.append(pPr)
    
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    b = OxmlElement('w:b')
    i = OxmlElement('w:i')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '1F4E79')
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '22') # 11pt
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), '22')
    rPr.extend([b, i, color, sz, szCs])
    r.append(rPr)
    
    t = OxmlElement('w:t')
    t.text = text
    r.append(t)
    p.append(r)
    return p

def extract_end_line_run(p_elm):
    """Extract the run containing the thinThick end-of-report line from a paragraph."""
    for r in p_elm.findall(qn('w:r')):
        if r.find(f".//{MC_AC}") is not None or r.find(f".//{qn('w:drawing')}") is not None:
            return copy.deepcopy(r)
    return None

def ensure_p_prop(p_elm, prop_name):
    """Ensure <w:keepNext/> or <w:keepLines/> exists inside <w:pPr> of a paragraph."""
    pPr = p_elm.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p_elm.insert(0, pPr)
    if pPr.find(qn(prop_name)) is None:
        pPr.append(OxmlElement(prop_name))

def ensure_tr_cantsplit(tr_elm):
    """Ensure <w:cantSplit/> exists inside <w:trPr> of a table row."""
    trPr = tr_elm.find(qn('w:trPr'))
    if trPr is None:
        trPr = OxmlElement('w:trPr')
        tr_elm.insert(0, trPr)
    if trPr.find(qn('w:cantSplit')) is None:
        trPr.append(OxmlElement('w:cantSplit'))

def set_table_rows_keep_next(tbl_elm, last_n_rows=None):
    """Set cantSplit on all rows, and keepNext + keepLines on paragraphs in the last_n_rows (or all rows if None)."""
    rows = tbl_elm.findall(qn('w:tr'))
    for tr in rows:
        ensure_tr_cantsplit(tr)
    target_rows = rows if last_n_rows is None else rows[-last_n_rows:]
    for tr in target_rows:
        for p in tr.findall(f".//{qn('w:p')}"):
            ensure_p_prop(p, 'w:keepNext')
            ensure_p_prop(p, 'w:keepLines')

def apply_keep_with_next_rules(doc):
    """
    Enforce page-break rule:
    Any table that has an accompanying conclusion ('Đánh giá:' / 'Assessment:' + conclusion text)
    must keep its conclusion on the same page as the table (linking last rows of the result table
    -> Note table -> 'Đánh giá:'/'Assessment:' -> conclusion paragraph via w:keepNext & w:keepLines).
    Also keep section headings with their following table.
    """
    body = doc.element.body
    children = list(body)
    n = len(children)

    # 1. Set cantSplit on all tables (except cover table #0)
    tbl_count = 0
    for el in children:
        if el.tag == qn('w:tbl'):
            if tbl_count > 0:
                for tr in el.findall(qn('w:tr')):
                    ensure_tr_cantsplit(tr)
            tbl_count += 1

    # 2. Keep numbered headings, image/layout titles, and '[Mẫu B...' / '[Template Option B...' labels with next element
    keep_next_prefixes = (
        "[Mẫu B", "[Template Option B", "9.2.",
        "Hình ảnh", "Sơ đồ", "Mô tả hệ thống", "Khu vực thực hiện theo",
        "Photo", "Layout of", "Description of compressed", "Measurement area was"
    )
    for idx, el in enumerate(children):
        if el.tag == qn('w:p'):
            txt = "".join(el.itertext()).strip()
            has_num = el.find(f".//{qn('w:numPr')}") is not None
            if has_num or txt.startswith(keep_next_prefixes):
                ensure_p_prop(el, 'w:keepNext')
                ensure_p_prop(el, 'w:keepLines')

    # 3. Glue every Note table ('Ghi chú:' / 'Note:') with the last 2 rows of its preceding Result table
    for idx, el in enumerate(children):
        if el.tag == qn('w:tbl'):
            first_cell_txt = "".join(el.itertext()).strip()
            if first_cell_txt.startswith("Ghi chú") or first_cell_txt.startswith("Note"):
                # Find the preceding table
                k = idx - 1
                while k >= 0:
                    if children[k].tag not in (qn('w:p'), qn('w:tbl')):
                        k -= 1
                        continue
                    if children[k].tag == qn('w:p') and not "".join(children[k].itertext()).strip():
                        ensure_p_prop(children[k], 'w:keepNext')
                        k -= 1
                        continue
                    break
                if k >= 0 and children[k].tag == qn('w:tbl'):
                    rows = children[k].findall(qn('w:tr'))
                    last_n = None if len(rows) <= 4 else 2
                    set_table_rows_keep_next(children[k], last_n_rows=last_n)

    # 4. Link every 'Đánh giá:' / 'Assessment:' block with its preceding Note table & Result table,
    #    and link 'Đánh giá:' / 'Assessment:' with the conclusion paragraph(s) after it.
    for idx, el in enumerate(children):
        if el.tag != qn('w:p'):
            continue
        txt = "".join(el.itertext()).strip()
        if txt.startswith("Đánh giá") or txt.startswith("Assessment"):
            # Keep 'Đánh giá:' / 'Assessment:' with the conclusion text below it
            ensure_p_prop(el, 'w:keepNext')
            ensure_p_prop(el, 'w:keepLines')

            # Find conclusion paragraph(s) immediately following (skip bookmarkStart/End etc.)
            j = idx + 1
            while j < n:
                if children[j].tag not in (qn('w:p'), qn('w:tbl')):
                    j += 1
                    continue
                if children[j].tag != qn('w:p'):
                    break
                j_txt = "".join(children[j].itertext()).strip()
                if not j_txt:
                    # Empty paragraph between Đánh giá and conclusion text
                    ensure_p_prop(children[j], 'w:keepNext')
                    j += 1
                    continue
                # This is the conclusion text paragraph
                ensure_p_prop(children[j], 'w:keepLines')
                break

            # Walk backward from 'Đánh giá:' / 'Assessment:' through empty paragraphs, Note tables, and Result table(s)
            k = idx - 1
            seen_result_tbl = False
            while k >= 0:
                prev_el = children[k]
                if prev_el.tag not in (qn('w:p'), qn('w:tbl')):
                    k -= 1
                    continue
                if prev_el.tag == qn('w:p'):
                    p_txt = "".join(prev_el.itertext()).strip()
                    if not p_txt:
                        # Glue empty paragraph so keepNext chain is not broken
                        ensure_p_prop(prev_el, 'w:keepNext')
                        k -= 1
                        continue
                    else:
                        break
                elif prev_el.tag == qn('w:tbl'):
                    first_cell_txt = "".join(prev_el.itertext()).strip()
                    is_note_tbl = first_cell_txt.startswith("Ghi chú") or first_cell_txt.startswith("Note")
                    if is_note_tbl and not seen_result_tbl:
                        # All rows of the Note table must keepNext to glue to 'Đánh giá:'
                        set_table_rows_keep_next(prev_el, last_n_rows=None)
                        k -= 1
                        continue
                    else:
                        # This is the Result table preceding the Note/Assessment
                        rows = prev_el.findall(qn('w:tr'))
                        # Keep last 2 rows (or all rows if small table <= 4 rows) with the Note + Conclusion
                        last_n = None if len(rows) <= 4 else 2
                        set_table_rows_keep_next(prev_el, last_n_rows=last_n)
                        seen_result_tbl = True
                        break

# ==============================================================================
# 1. BUILD CLEANROOM VIETNAMESE MASTER TEMPLATE (2026)
# ==============================================================================
def build_cleanroom_vi():
    base_path = os.path.join(dir_2026, "TBKQ - viet - 1 - phong sach - Q326D108033.docx")
    kn_vi_path = os.path.join(dir_2026, "TBKQ - vn - 2 - khi nen - Q326D108033.docx")
    d1_path = os.path.join(dir_2025, "985 - TBKQ - viet - 1 - phong sach - đợt 1.docx")
    d2_path = os.path.join(dir_2025, "985 - TBKQ - viet - 2 - phong sach - đợt 2.docx")
    d3_path = os.path.join(dir_2025, "985 - TBKQ - viet - 3 - phong sach - đợt 3.docx")

    doc = docx.Document(base_path)
    doc_kn = docx.Document(kn_vi_path)
    doc_d1 = docx.Document(d1_path)
    doc_d2 = docx.Document(d2_path)
    doc_d3 = docx.Document(d3_path)

    # 1. Update Table #0 (Cover page) rows 9..10 with multi-line rows 9..12 from kn_vi
    t0 = doc.tables[0]._tbl
    t0_kn = doc_kn.tables[0]._tbl
    tr9_orig = t0.findall(qn('w:tr'))[9]
    tr10_orig = t0.findall(qn('w:tr'))[10]
    tr9_new = copy.deepcopy(t0_kn.findall(qn('w:tr'))[9])
    tr10_new = copy.deepcopy(t0_kn.findall(qn('w:tr'))[10])
    tr11_new = copy.deepcopy(t0_kn.findall(qn('w:tr'))[11])
    tr12_new = copy.deepcopy(t0_kn.findall(qn('w:tr'))[12])

    set_cell_bullet_lines(tr10_new.findall(qn('w:tc'))[0], [
        "Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;",
        "Máy đo nồng độ hạt bụi PMS Lasair III 110, USA;",
        "Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;",
        "Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;",
        "Thiết bị đo chênh lệch áp suất TSI 8375, USA;",
        "Thiết bị đo lưu lượng TSI 8375, USA."
    ])
    set_cell_bullet_lines(tr12_new.findall(qn('w:tc'))[0], [
        "ISO 14644-1:2015 Cleanrooms and associated controlled environments — Part 1: Classification of air cleanliness by particle concentration;",
        "ISO 14644-3:2019 Cleanrooms and associated controlled environments — Part 3: Test methods;",
        "Mức quy định được cung cấp bởi đại diện Công ty TNHH Nextern Việt Nam."
    ])
    t0.insert(t0.index(tr9_orig), tr9_new)
    t0.insert(t0.index(tr9_orig), tr10_new)
    t0.insert(t0.index(tr9_orig), tr11_new)
    t0.insert(t0.index(tr9_orig), tr12_new)
    t0.remove(tr9_orig)
    t0.remove(tr10_orig)

    # 2. Update Table #1 (Monitoring locations & points) with all 5 parameters
    t1 = doc.tables[1]
    set_monitoring_points_cell(t1.rows[1].cells[3]._tc, [
        ("Nồng độ hạt bụi", "21 vị trí"),
        ("Độ rò rỉ lọc HEPA", "12 vị trí"),
        ("Nhiệt độ - độ ẩm", "01 vị trí"),
        ("Chênh lệch áp suất", "01 vị trí"),
        ("Lưu lượng gió", "12 vị trí"),
    ])
    set_monitoring_points_cell(t1.rows[2].cells[3]._tc, [
        ("Nồng độ hạt bụi", "02 vị trí"),
        ("Độ rò rỉ lọc HEPA", "01 vị trí"),
        ("Nhiệt độ - độ ẩm", "01 vị trí"),
        ("Chênh lệch áp suất", "01 vị trí"),
        ("Lưu lượng gió", "01 vị trí"),
    ])
    set_monitoring_points_cell(t1.rows[3].cells[3]._tc, [
        ("Nồng độ hạt bụi", "08 vị trí"),
        ("Độ rò rỉ lọc HEPA", "02 vị trí"),
        ("Nhiệt độ - độ ẩm", "01 vị trí"),
        ("Chênh lệch áp suất", "01 vị trí"),
        ("Lưu lượng gió", "02 vị trí"),
    ])

    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))

    # Fix P[18] ("Kết quả đo nồng độ hạt bụi") numId from 27 -> 10 so all 9.x headings auto-number
    set_num_id(body[18], "10", "1")

    # Extract the end-of-report thinThick line from body[30] and strip it from body[30]
    end_line_run = extract_end_line_run(body[30])
    strip_drawings(body[30])

    # Insert 2x3 layout table from d3 (index 17) right after body[14] (1x1 layout table)
    layout_2x3 = copy.deepcopy(doc_d3.element.body[17])
    strip_drawings(layout_2x3)
    for tc in layout_2x3.findall(qn('w:tr'))[0].findall(qn('w:tc')):
        p_tc = tc.find(qn('w:p'))
        if p_tc is not None:
            r_tc = OxmlElement('w:r')
            t_tc = OxmlElement('w:t')
            t_tc.text = "[Chèn sơ đồ vị trí đo]"
            r_tc.append(t_tc)
            p_tc.append(r_tc)
    pos_after_layout = body.index(body[15])
    body.insert(pos_after_layout, make_label_paragraph("[Mẫu B - Bảng sơ đồ vị trí đo nhiều chỉ tiêu (trích xuất Form 2025)]:"))
    body.insert(pos_after_layout + 1, layout_2x3)

    blocks_to_append = []

    # --- 9.3. Kết quả đo độ rò rỉ lọc HEPA (from d1 indices 32..46) ---
    for idx in range(32, 47):
        el = copy.deepcopy(doc_d1.element.body[idx])
        strip_drawings(el)
        set_num_id(el, "10", "1")
        replace_text_in_element(el, "trang 3", "trang 1")
        replace_text_in_element(el, "Công ty Cổ phần Kỹ thuật INTECH Sài Gòn", "Công ty TNHH Nextern Việt Nam")
        blocks_to_append.append(el)

    # --- 9.4. Kết quả đo nhiệt độ - độ ẩm (from d2 indices 32..39) ---
    for idx in range(32, 40):
        el = copy.deepcopy(doc_d2.element.body[idx])
        strip_drawings(el)
        set_num_id(el, "10", "1")
        replace_text_in_element(el, "trang 3", "trang 1")
        replace_text_in_element(el, "Công ty Cổ phần Kỹ thuật INTECH Sài Gòn", "Công ty TNHH Nextern Việt Nam")
        replace_text_in_element(el, "Cổ phần Kỹ thuật INTECH Sài Gòn", "TNHH Nextern Việt Nam")
        blocks_to_append.append(el)

    # --- 9.5. Kết quả đo chênh lệch áp suất giữa hai khu vực (from d3 indices 32..38) ---
    for idx in range(32, 39):
        el = copy.deepcopy(doc_d3.element.body[idx])
        strip_drawings(el)
        set_num_id(el, "10", "1")
        replace_text_in_element(el, "trang 3", "trang 1")
        replace_text_in_element(el, "công ty", "Công ty")
        replace_text_in_element(el, "Công ty Cổ phần Kỹ thuật INTECH Sài Gòn", "Công ty TNHH Nextern Việt Nam")
        replace_text_in_element(el, "Cổ phần Kỹ thuật INTECH Sài Gòn", "TNHH Nextern Việt Nam")
        blocks_to_append.append(el)

    # --- 9.6. Kết quả đo lưu lượng khí cấp qua lọc HEPA (from d3 indices 39..44) ---
    for idx in range(39, 45):
        el = copy.deepcopy(doc_d3.element.body[idx])
        strip_drawings(el)
        set_num_id(el, "10", "1")
        replace_text_in_element(el, "trang 3", "trang 1")
        replace_text_in_element(el, "công ty", "Công ty")
        replace_text_in_element(el, "Công ty CP Kỹ thuật INTECH Sài Gòn", "Công ty TNHH Nextern Việt Nam")
        replace_text_in_element(el, "Công ty Cổ phần Kỹ thuật INTECH Sài Gòn", "Công ty TNHH Nextern Việt Nam")
        replace_text_in_element(el, "Cổ phần Kỹ thuật INTECH Sài Gòn", "TNHH Nextern Việt Nam")
        blocks_to_append.append(el)

    if end_line_run is not None and blocks_to_append:
        blocks_to_append[-1].insert(0, end_line_run)

    for b in blocks_to_append:
        sectPr.addprevious(b)

    apply_keep_with_next_rules(doc)

    dir_core = os.path.join(base_dir, "Core")
    os.makedirs(dir_core, exist_ok=True)
    out_file = os.path.join(dir_core, "TEMPLATE - viet - 1 - phong sach - 2026.docx")
    doc.save(out_file)
    print("Saved:", out_file)


# ==============================================================================
# 2. BUILD CLEANROOM ENGLISH MASTER TEMPLATE (2026)
# ==============================================================================
def build_cleanroom_en():
    base_path = os.path.join(dir_2026, "TBKQ - en - 1 - phong sach - Q326D108033.docx")
    kn_en_path = os.path.join(dir_2026, "TBKQ - en - 2 - khi nen - Q326D108033.docx")
    d1_path = os.path.join(dir_2025, "985 - TBKQ - eng - 1 - phong sach - đợt 1.docx")
    d2_path = os.path.join(dir_2025, "985 - TBKQ - eng - 2 - phong sach - đợt 2.docx")
    d3_path = os.path.join(dir_2025, "985 - TBKQ - eng - 3 - phong sach - đợt 3.docx")

    doc = docx.Document(base_path)
    doc_kn = docx.Document(kn_en_path)
    doc_d1 = docx.Document(d1_path)
    doc_d2 = docx.Document(d2_path)
    doc_d3 = docx.Document(d3_path)

    # 1. Update Table #0 (Cover page) rows 9..10 with multi-line rows 9..12 from kn_en
    t0 = doc.tables[0]._tbl
    t0_kn = doc_kn.tables[0]._tbl
    tr9_orig = t0.findall(qn('w:tr'))[9]
    tr10_orig = t0.findall(qn('w:tr'))[10]
    tr9_new = copy.deepcopy(t0_kn.findall(qn('w:tr'))[9])
    tr10_new = copy.deepcopy(t0_kn.findall(qn('w:tr'))[10])
    tr11_new = copy.deepcopy(t0_kn.findall(qn('w:tr'))[11])
    tr12_new = copy.deepcopy(t0_kn.findall(qn('w:tr'))[12])

    set_cell_bullet_lines(tr10_new.findall(qn('w:tc'))[0], [
        "Portable particle counter PMS Lasair III 110, USA / TSI AeroTrak 9350, USA;",
        "HEPA leakage test ATI 2i and Aerosol generator ATI 6D, USA;",
        "Thermohygrometer Testo 625, Germany;",
        "Airflow-pressure differential meter TSI AccuBalance 8380, USA."
    ])
    set_cell_bullet_lines(tr12_new.findall(qn('w:tc'))[0], [
        "ISO 14644-1:2015 Cleanrooms and associated controlled environments — Part 1: Classification of air cleanliness by particle concentration;",
        "ISO 14644-3:2019 Cleanrooms and associated controlled environments — Part 3: Test methods;",
        "Permissible limit provided by the representative of Nextern Vietnam Limited Liability Company."
    ])
    t0.insert(t0.index(tr9_orig), tr9_new)
    t0.insert(t0.index(tr9_orig), tr10_new)
    t0.insert(t0.index(tr9_orig), tr11_new)
    t0.insert(t0.index(tr9_orig), tr12_new)
    t0.remove(tr9_orig)
    t0.remove(tr10_orig)

    # 2. Update Table #1 (Monitoring locations & points) with all 5 parameters
    t1 = doc.tables[1]
    set_monitoring_points_cell(t1.rows[1].cells[3]._tc, [
        ("Particle concentration", "21 points"),
        ("HEPA filter leakage", "12 points"),
        ("Temperature - humidity", "01 point"),
        ("Pressure differential", "01 point"),
        ("Airflow rate", "12 points"),
    ])
    set_monitoring_points_cell(t1.rows[2].cells[3]._tc, [
        ("Particle concentration", "02 points"),
        ("HEPA filter leakage", "01 point"),
        ("Temperature - humidity", "01 point"),
        ("Pressure differential", "01 point"),
        ("Airflow rate", "01 point"),
    ])
    set_monitoring_points_cell(t1.rows[3].cells[3]._tc, [
        ("Particle concentration", "08 points"),
        ("HEPA filter leakage", "02 points"),
        ("Temperature - humidity", "01 point"),
        ("Pressure differential", "01 point"),
        ("Airflow rate", "02 points"),
    ])

    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))

    # Extract end-of-report line from body[33] and strip it from body[33]
    end_line_run = extract_end_line_run(body[33])
    strip_drawings(body[33])

    # Insert 2x3 layout table from d3 (index 18) after body[16] (1x1 layout table)
    layout_2x3 = copy.deepcopy(doc_d3.element.body[18])
    strip_drawings(layout_2x3)
    for tc in layout_2x3.findall(qn('w:tr'))[0].findall(qn('w:tc')):
        p_tc = tc.find(qn('w:p'))
        if p_tc is not None:
            r_tc = OxmlElement('w:r')
            t_tc = OxmlElement('w:t')
            t_tc.text = "[Insert layout diagram]"
            r_tc.append(t_tc)
            p_tc.append(r_tc)
    pos_after_layout = body.index(body[17])
    body.insert(pos_after_layout, make_label_paragraph("[Template Option B - Multi-parameter layout diagram table (extracted from 2025 Form)]:"))
    body.insert(pos_after_layout + 1, layout_2x3)

    blocks_to_append = []

    # --- 9.3. Measurement result of HEPA filter leakage (from d1 indices 32..44) ---
    for idx in range(32, 45):
        el = copy.deepcopy(doc_d1.element.body[idx])
        strip_drawings(el)
        set_num_id(el, "10", "1")
        replace_text_in_element(el, "page 3", "page 1")
        replace_text_in_element(el, "INTECH Saigon Engineering Joint Stock Company", "Nextern Vietnam Limited Liability Company")
        blocks_to_append.append(el)

    # --- 9.4. Measurement result of temperature – humidity (from d2 indices 31..36) ---
    for idx in range(31, 37):
        el = copy.deepcopy(doc_d2.element.body[idx])
        strip_drawings(el)
        set_num_id(el, "10", "1")
        replace_text_in_element(el, "page 3", "page 1")
        replace_text_in_element(el, "INTECH Saigon Engineering Joint Stock Company", "Nextern Vietnam Limited Liability Company")
        blocks_to_append.append(el)

    # --- 9.5. Measurement result of pressure differentials between areas (from d3 indices 31..37) ---
    for idx in range(31, 38):
        el = copy.deepcopy(doc_d3.element.body[idx])
        strip_drawings(el)
        set_num_id(el, "10", "1")
        replace_text_in_element(el, "page 3", "page 1")
        replace_text_in_element(el, "INTECH Saigon Engineering", "Nextern Vietnam")
        replace_text_in_element(el, "Joint Stock Company", "Limited Liability Company")
        blocks_to_append.append(el)

    # --- 9.6. Measurement result of airflow rate supply through HEPA filter (from d3 indices 38..46) ---
    for idx in range(38, 47):
        el = copy.deepcopy(doc_d3.element.body[idx])
        strip_drawings(el)
        set_num_id(el, "10", "1")
        replace_text_in_element(el, "page 3", "page 1")
        replace_text_in_element(el, "Intech Saigon Engineering JSC", "Nextern Vietnam Limited Liability Company")
        replace_text_in_element(el, "INTECH Saigon Engineering Joint Stock Company", "Nextern Vietnam Limited Liability Company")
        replace_text_in_element(el, "INTECH Saigon Engineering", "Nextern Vietnam")
        replace_text_in_element(el, "Joint Stock Company", "Limited Liability Company")
        blocks_to_append.append(el)

    if end_line_run is not None and blocks_to_append:
        blocks_to_append[-1].insert(0, end_line_run)

    for b in blocks_to_append:
        sectPr.addprevious(b)

    apply_keep_with_next_rules(doc)

    dir_core = os.path.join(base_dir, "Core")
    os.makedirs(dir_core, exist_ok=True)
    out_file = os.path.join(dir_core, "TEMPLATE - en - 1 - phong sach - 2026.docx")
    doc.save(out_file)
    print("Saved:", out_file)


# ==============================================================================
# 3. BUILD COMPRESSED AIR VIETNAMESE MASTER TEMPLATE (2026)
# ==============================================================================
def build_khinen_vi():
    base_path = os.path.join(dir_2026, "TBKQ - vn - 2 - khi nen - Q326D108033.docx")
    d4_path = os.path.join(dir_2025, "985 - TBKQ - viet - 4 - khi nen.docx")

    doc = docx.Document(base_path)
    doc_d4 = docx.Document(d4_path)

    # Normalize Table #0 Equipment (row 10) and Standards (row 12) according to PHẦN 2 & PHẦN 3
    t0 = doc.tables[0]._tbl
    set_cell_bullet_lines(t0.findall(qn('w:tr'))[10].findall(qn('w:tc'))[0], [
        "Dụng cụ lấy mẫu tổng dầu (Regulator, Charcoal tube, Cassette);",
        "Máy đo nồng độ hạt bụi Lasair III 110 và phụ kiện HPD II, USA;",
        "Thiết bị đo nhiệt độ điểm sương."
    ])
    set_cell_bullet_lines(t0.findall(qn('w:tr'))[12].findall(qn('w:tc'))[0], [
        "ISO 8573 - 1:2010 Compressed air - Part 1: Contaminants and purity classes;",
        "Mức yêu cầu đối với khí nén cung cấp bởi Công ty TNHH Nextern Việt Nam."
    ])

    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))

    # Clean accidental 7 connectors in body[34] ("Đánh giá:" under Nhiệt độ điểm sương)
    strip_drawings(body[34])

    # Extract end-of-report line from body[52] and strip it from body[52]
    end_line_run = extract_end_line_run(body[52])
    strip_drawings(body[52])

    # Standardize sub-headings 9.2.1, 9.2.2, 9.2.3 to numId=10, ilvl=2
    for idx in [23, 30, 43]:
        set_num_id(body[idx], "10", "2")

    # Grab reference anchor elements in body BEFORE inserting anything
    anchor_after_loc_tbl = body[13]   # right after 2026 Table #2 (Note for location table)
    anchor_after_oil = body[29]       # empty paragraph after Oil Assessment
    anchor_after_dew = body[36]       # empty paragraph after Dewpoint Assessment
    anchor_end = sectPr               # end of body (after Particle Assessment)

    # 1. Insert 2025 Multi-location Table (KN1, KN2) + Note + 1x2 Photo Table
    t_loc_25 = copy.deepcopy(doc_d4.element.body[13])
    n_loc_25 = copy.deepcopy(doc_d4.element.body[15])
    t_photo_25 = copy.deepcopy(doc_d4.element.body[18])
    strip_drawings(t_photo_25)
    for tc in t_photo_25.findall(qn('w:tr'))[0].findall(qn('w:tc')):
        p_tc = tc.find(qn('w:p'))
        if p_tc is not None:
            r_tc = OxmlElement('w:r')
            t_tc = OxmlElement('w:t')
            t_tc.text = "[Chèn ảnh vị trí lấy mẫu/đo đạc]"
            r_tc.append(t_tc)
            p_tc.append(r_tc)

    anchor_after_loc_tbl.addprevious(make_label_paragraph("[Mẫu B - Bảng vị trí & khung ảnh trường hợp nhiều vị trí đo (trích xuất Form 2025)]:"))
    anchor_after_loc_tbl.addprevious(t_loc_25)
    anchor_after_loc_tbl.addprevious(n_loc_25)
    anchor_after_loc_tbl.addprevious(t_photo_25)

    # 2. Insert 2025 Total Oil Content Table (ISO classification only, no client limit row)
    t_oil_25 = copy.deepcopy(doc_d4.element.body[26])
    n_oil_25 = copy.deepcopy(doc_d4.element.body[28])
    anchor_after_oil.addprevious(make_label_paragraph("[Mẫu B - Bảng Hàm lượng tổng dầu trường hợp nhiều vị trí / chỉ phân loại ISO 8573-1 (trích xuất Form 2025)]:"))
    anchor_after_oil.addprevious(t_oil_25)
    anchor_after_oil.addprevious(n_oil_25)

    # 3. Insert 2025 Pressure Dewpoint Table (6 columns, Class inside T2 cell, no client limit row)
    t_dew_25 = copy.deepcopy(doc_d4.element.body[30])
    n_dew_25 = copy.deepcopy(doc_d4.element.body[32])
    replace_text_in_element(n_dew_25, "Công  \n          ty INTECH Sài Gòn", "Công ty TNHH Nextern Việt Nam")
    replace_text_in_element(n_dew_25, "INTECH Sài Gòn", "TNHH Nextern Việt Nam")
    anchor_after_dew.addprevious(make_label_paragraph("[Mẫu B - Bảng Nhiệt độ điểm sương 6 cột trường hợp nhiều vị trí / chỉ phân loại ISO 8573-1 (trích xuất Form 2025)]:"))
    anchor_after_dew.addprevious(t_dew_25)
    anchor_after_dew.addprevious(n_dew_25)

    # 4. Insert 2025 Particle Concentration Table (multi-location KN1, KN2) before sectPr
    t_part_25 = copy.deepcopy(doc_d4.element.body[34])
    n_part_25 = copy.deepcopy(doc_d4.element.body[36])
    anchor_end.addprevious(make_label_paragraph("[Mẫu B - Bảng Nồng độ hạt bụi trường hợp nhiều vị trí KN1, KN2 (trích xuất Form 2025)]:"))
    anchor_end.addprevious(t_part_25)
    anchor_end.addprevious(n_part_25)

    # Add final paragraph with end_line_run
    final_p = copy.deepcopy(doc_d4.element.body[27]) # clean empty paragraph
    if end_line_run is not None:
        final_p.insert(0, end_line_run)
    anchor_end.addprevious(final_p)

    apply_keep_with_next_rules(doc)

    dir_core = os.path.join(base_dir, "Core")
    os.makedirs(dir_core, exist_ok=True)
    out_file = os.path.join(dir_core, "TEMPLATE - vn - 2 - khi nen - 2026.docx")
    doc.save(out_file)
    print("Saved:", out_file)


# ==============================================================================
# 4. BUILD COMPRESSED AIR ENGLISH MASTER TEMPLATE (2026)
# ==============================================================================
def build_khinen_en():
    base_path = os.path.join(dir_2026, "TBKQ - en - 2 - khi nen - Q326D108033.docx")
    d4_path = os.path.join(dir_2025, "985 - TBKQ - eng - 4 - khi nen.docx")

    doc = docx.Document(base_path)
    doc_d4 = docx.Document(d4_path)

    # Normalize Table #0 Equipment (row 10) and Standards (row 12) according to PHẦN 2 & PHẦN 3
    t0 = doc.tables[0]._tbl
    set_cell_bullet_lines(t0.findall(qn('w:tr'))[10].findall(qn('w:tc'))[0], [
        "Glass fiber filter Pall life Sciences, carbon activated tube ORBO®;",
        "Airborne particle counter PMS Lasair III 110 and accessory PMS HPD II-100, USA;",
        "Dewpoint meter CS Instrument GmbH DP510 and pressure meter CS16, Germany."
    ])
    set_cell_bullet_lines(t0.findall(qn('w:tr'))[12].findall(qn('w:tc'))[0], [
        "ISO 8573 - 1:2010 Compressed air - Part 1: Contaminants and purity classes;",
        "Requirements for compressed air provided by Nextern Vietnam Limited Liability Company."
    ])

    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))

    # Extract end-of-report line from body[44] and strip it from body[44]
    end_line_run = extract_end_line_run(body[44])
    strip_drawings(body[44])

    anchor_after_loc_tbl = body[12]   # right after 2026 Table #2
    anchor_after_oil = body[26]       # empty paragraph after Oil Assessment
    anchor_after_dew = body[33]       # empty paragraph after Dewpoint Assessment
    anchor_end = sectPr

    # 1. Insert 2025 Multi-location Table (KN1, KN2) + Note + 1x2 Photo Table
    t_loc_25 = copy.deepcopy(doc_d4.element.body[9])
    n_loc_25 = copy.deepcopy(doc_d4.element.body[11])
    t_photo_25 = copy.deepcopy(doc_d4.element.body[13])
    strip_drawings(t_photo_25)
    for tc in t_photo_25.findall(qn('w:tr'))[0].findall(qn('w:tc')):
        p_tc = tc.find(qn('w:p'))
        if p_tc is not None:
            r_tc = OxmlElement('w:r')
            t_tc = OxmlElement('w:t')
            t_tc.text = "[Insert photo of sampling/measuring location]"
            r_tc.append(t_tc)
            p_tc.append(r_tc)

    anchor_after_loc_tbl.addprevious(make_label_paragraph("[Template Option B - Multi-location sampling table & photo grid (extracted from 2025 Form)]:"))
    anchor_after_loc_tbl.addprevious(t_loc_25)
    anchor_after_loc_tbl.addprevious(n_loc_25)
    anchor_after_loc_tbl.addprevious(t_photo_25)

    # 2. Insert 2025 Total Oil Content Table (ISO classification only)
    t_oil_25 = copy.deepcopy(doc_d4.element.body[21])
    n_oil_25 = copy.deepcopy(doc_d4.element.body[23])
    anchor_after_oil.addprevious(make_label_paragraph("[Template Option B - Total oil content table for multi-location / ISO 8573-1 classification only (extracted from 2025 Form)]:"))
    anchor_after_oil.addprevious(t_oil_25)
    anchor_after_oil.addprevious(n_oil_25)

    # 3. Insert 2025 Pressure Dewpoint Table (6 columns)
    t_dew_25 = copy.deepcopy(doc_d4.element.body[25])
    n_dew_25 = copy.deepcopy(doc_d4.element.body[27])
    replace_text_in_element(n_dew_25, "INTECH Saigon Company", "Nextern Vietnam Limited Liability Company")
    anchor_after_dew.addprevious(make_label_paragraph("[Template Option B - Pressure dewpoint 6-column table for multi-location / ISO 8573-1 classification only (extracted from 2025 Form)]:"))
    anchor_after_dew.addprevious(t_dew_25)
    anchor_after_dew.addprevious(n_dew_25)

    # 4. Insert 2025 Particle Concentration Table (multi-location KN1, KN2)
    t_part_25 = copy.deepcopy(doc_d4.element.body[29])
    n_part_25 = copy.deepcopy(doc_d4.element.body[31])
    anchor_end.addprevious(make_label_paragraph("[Template Option B - Particle concentration table for multi-location KN1, KN2 (extracted from 2025 Form)]:"))
    anchor_end.addprevious(t_part_25)
    anchor_end.addprevious(n_part_25)

    final_p = copy.deepcopy(doc_d4.element.body[22])
    if end_line_run is not None:
        final_p.insert(0, end_line_run)
    anchor_end.addprevious(final_p)

    apply_keep_with_next_rules(doc)

    dir_core = os.path.join(base_dir, "Core")
    os.makedirs(dir_core, exist_ok=True)
    out_file = os.path.join(dir_core, "TEMPLATE - en - 2 - khi nen - 2026.docx")
    doc.save(out_file)
    print("Saved:", out_file)

if __name__ == "__main__":
    build_cleanroom_vi()
    build_cleanroom_en()
    build_khinen_vi()
    build_khinen_en()
