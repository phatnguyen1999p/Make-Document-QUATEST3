# -*- coding: utf-8 -*-
"""
Table formatting and OpenXML layout utilities for QUATEST 3 Documents.
Enforces:
- AutoFit to Window for photo and layout tables (AGENTS Section 3)
- Row height >= 0.6cm (optimal 0.7 - 0.8cm = 440 twips) for result tables (AGENTS Section 3)
- Repeat header rows (<w:tblHeader/>) for data tables spanning across pages (AGENTS Section 3)
- Prevent split rows (<w:cantSplit/>) and enforce keepNext coupling (AGENTS Section 3)
"""

import os
from docx.oxml import parse_xml
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

from .styles import W_NS, set_cell_margins, set_cell_content
from .paragraphs import set_paragraph_text

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Standard width for 2-column image grid (7 - 8 cm, strictly no crop)
PHOTO_WIDTH_2COL = Cm(7.5)

# Standard width for 3-column image grid (4 - 5 cm, strictly no crop)
PHOTO_WIDTH_3COL = Cm(4.8)


def strip_image_crops(doc_or_element):
    """
    Strips all crop attributes from <a:srcRect> elements in the given doc or element.
    Strictly enforces rule: 'tuyệt đối không crop ảnh, Áp dụng tất cả trường hợp'.
    """
    root = getattr(doc_or_element, "element", doc_or_element)
    if hasattr(root, "body"):
        root = root.body
    for sr in root.findall(f".//{{{A_NS}}}srcRect"):
        if sr.attrib:
            sr.attrib.clear()


def add_tbl_header(tr):
    """Marks a table row as a repeating header across page splits (<w:tblHeader/>)."""
    trPr = tr._tr.get_or_add_trPr()
    if trPr.find(f"{{{W_NS}}}tblHeader") is None:
        trPr.append(parse_xml(f'<w:tblHeader xmlns:w="{W_NS}"/>'))


def add_cant_split(tr):
    """Prevents a table row from splitting across a page break (<w:cantSplit/>)."""
    trPr = tr._tr.get_or_add_trPr()
    if trPr.find(f"{{{W_NS}}}cantSplit") is None:
        trPr.append(parse_xml(f'<w:cantSplit xmlns:w="{W_NS}"/>'))


def set_table_row_height(tr_or_tbl, row_idx: int = None, height_twips: int = 440, h_rule: str = "atLeast"):
    """
    Sets minimum row height in twips (440 twips = 0.78cm, meeting the 0.7cm - 0.8cm rule).
    If a table is passed without row_idx, applies to all rows.
    """
    if hasattr(tr_or_tbl, "rows"):
        # Tr_or_tbl is a docx.Table
        target_rows = tr_or_tbl.rows if row_idx is None else [tr_or_tbl.rows[row_idx]]
    else:
        # Tr_or_tbl is a single TableRow
        target_rows = [tr_or_tbl]

    for r in target_rows:
        trPr = r._tr.get_or_add_trPr()
        old_h = trPr.find(f"{{{W_NS}}}trHeight")
        if old_h is not None:
            trPr.remove(old_h)
        trPr.append(parse_xml(f'<w:trHeight xmlns:w="{W_NS}" w:val="{height_twips}" w:hRule="{h_rule}"/>'))


def set_table_col_widths(table, widths_dxa: list[int]):
    """Sets explicit column widths on table cells and table grid."""
    tblGrid = table._tbl.find(f"{{{W_NS}}}tblGrid")
    if tblGrid is not None:
        table._tbl.remove(tblGrid)
    new_grid = parse_xml(f'<w:tblGrid xmlns:w="{W_NS}"/>')
    for w in widths_dxa:
        new_grid.append(parse_xml(f'<w:gridCol xmlns:w="{W_NS}" w:w="{w}"/>'))
    table._tbl.insert(1, new_grid)

    # Walk the real <w:tc> elements of every row and accumulate the grid position through
    # gridSpan, so horizontally merged cells (header / footer rows) receive the SUM of the grid
    # columns they span. Using row.cells[idx] is wrong for merged rows (python-docx returns the
    # same cell for every spanned grid column) and used to distort the table layout.
    for tr in table._tbl.tr_lst:
        grid_pos = 0
        for tc in tr.tc_lst:
            tcPr = tc.get_or_add_tcPr()
            gs = tcPr.find(f"{{{W_NS}}}gridSpan")
            span = int(gs.get(f"{{{W_NS}}}val")) if gs is not None else 1
            w = sum(widths_dxa[grid_pos:grid_pos + span])
            grid_pos += span
            if w <= 0:
                continue
            old_w = tcPr.find(f"{{{W_NS}}}tcW")
            if old_w is not None:
                tcPr.remove(old_w)
            # tcW must be the first child of tcPr (schema order)
            tcPr.insert(0, parse_xml(f'<w:tcW xmlns:w="{W_NS}" w:w="{w}" w:type="dxa"/>'))


def set_table_autofit_window(table):
    """
    Strictly configures table to AutoFit to Window:
    - Width: 5000 (100% in pct mode)
    - Layout: autofit (<w:tblLayout w:type="autofit"/>)
    - Borders: none
    - Column widths: equally distributed percentage
    - Alignment: Center
    Mandatory for photo and layout tables per AGENTS.md Section 3.
    """
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr

    # Width 100% (5000 pct)
    tblW = tblPr.find(f"{{{W_NS}}}tblW")
    if tblW is not None:
        tblPr.remove(tblW)
    tblPr.append(parse_xml(f'<w:tblW xmlns:w="{W_NS}" w:w="5000" w:type="pct"/>'))

    # Table layout autofit
    tblLayout = tblPr.find(f"{{{W_NS}}}tblLayout")
    if tblLayout is not None:
        tblPr.remove(tblLayout)
    tblPr.append(parse_xml(f'<w:tblLayout xmlns:w="{W_NS}" w:type="autofit"/>'))

    # Remove all borders (borderless)
    tblBorders = tblPr.find(f"{{{W_NS}}}tblBorders")
    if tblBorders is not None:
        tblPr.remove(tblBorders)
    tblPr.append(
        parse_xml(
            f'<w:tblBorders xmlns:w="{W_NS}">'
            f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'</w:tblBorders>'
        )
    )

    # Set each cell to equal percentage width
    num_cols = len(table.columns)
    if num_cols > 0:
        pct_per_col = 5000 // num_cols
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                old_w = tcPr.find(f"{{{W_NS}}}tcW")
                if old_w is not None:
                    tcPr.remove(old_w)
                tcPr.append(parse_xml(f'<w:tcW xmlns:w="{W_NS}" w:w="{pct_per_col}" w:type="pct"/>'))
                set_cell_margins(cell, top=20, bottom=20, left=40, right=40)


def format_table_structure(table, is_note: bool = False, widths: list[int] = None, row_height=440):
    """
    Enforces standard QUATEST 3 table structure:
    - If is_note=True: borderless table for notes directly below result tables
    - If is_note=False: standard result table with row height (440 twips = 0.78cm),
      header repeat (<w:tblHeader/>), no split (<w:cantSplit/>), center alignment
    - row_height can be:
        + an int: uniform row height applied to all rows (e.g. 440)
        + a list of ints / None: specific row height per row index. If None, row height is untouched.
    """
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr

    tblW = tblPr.find(f"{{{W_NS}}}tblW")
    if tblW is not None:
        tblPr.remove(tblW)
    tblPr.append(parse_xml(f'<w:tblW xmlns:w="{W_NS}" w:w="5000" w:type="pct"/>'))

    tblBorders = tblPr.find(f"{{{W_NS}}}tblBorders")
    if tblBorders is not None:
        tblPr.remove(tblBorders)

    if is_note:
        # Borderless note table
        tblPr.append(
            parse_xml(
                f'<w:tblBorders xmlns:w="{W_NS}">'
                f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'</w:tblBorders>'
            )
        )
    else:
        # Standard thin single border (4/8 pt)
        tblPr.append(
            parse_xml(
                f'<w:tblBorders xmlns:w="{W_NS}">'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                f'</w:tblBorders>'
            )
        )
        if len(table.rows) > 0:
            add_tbl_header(table.rows[0])

    for r_idx, row in enumerate(table.rows):
        add_cant_split(row)
        if not is_note:
            if isinstance(row_height, (list, tuple)):
                h = row_height[r_idx] if r_idx < len(row_height) else 440
                if h is not None:
                    set_table_row_height(row, height_twips=h)
            elif row_height:
                set_table_row_height(row, height_twips=row_height)
        for cell in row.cells:
            set_cell_margins(cell, top=40, bottom=40, left=80, right=80)

    if widths:
        set_table_col_widths(table, widths)


def get_bsc_row_heights(has_particle: bool = False, has_uvc: bool = False) -> list:
    """
    Returns standardized row heights (in twips) for BSC / LAF result tables.
    Synchronizes row heights between VN and EN to guarantee equal table length,
    matching page breaks (clean split after Item 5), and identical page counts across versions.
    """
    heights = [
        440,   # 0: Header
        760,   # 1: Average inflow velocity (equivalent to 3 lines in VN)
        760,   # 2: Average downflow velocity (equivalent to 3 lines in VN)
        580,   # 3: Average illuminance (equivalent to 2 lines in VN)
        440,   # 4: Sound pressure level parent
        440,   #    Background noise
        440,   #    Operation noise
        440,   # 5: HEPA leakage test parent
        580,   #    Downflow filter (equivalent to 2 lines in VN)
        580,   #    Exhaust filter (equivalent to 2 lines in VN)
        440,   # 6: Airflow pattern test parent
        1000,  #    Inside cabinet workspace (equivalent to 4 lines in VN)
        600,   #    Perimeter of work opening (equivalent to 2 lines in VN)
    ]
    if has_uvc:
        heights.append(440)  # 7: UVC radiation intensity
    if has_particle:
        heights.append(440)  # Airborne particle concentration parent
        heights.extend([560, 560, 560, 560])  # 4 particle sizes (wraps to 2 lines for threshold)
    return heights


def get_fumehood_row_heights(num_openings: int = 1) -> list:
    """
    Returns standardized row heights (in twips) for Fume Hood result tables.
    Synchronizes row heights between VN and EN to guarantee equal table length
    and matching page breaks across versions.
    """
    heights = [440]  # Header
    for _ in range(num_openings):
        heights.append(760)   # Face velocity text (matches 3-4 lines in VN)
        heights.append(None)  # Sash diagram illustration (auto-sized by image)
    heights.append(580)   # Average illuminance (matches 2 lines in VN)
    heights.append(440)   # Sound pressure level parent
    heights.append(440)   # Background noise
    heights.append(440)   # Operation noise
    heights.append(440)   # Smoke pattern test parent
    heights.append(1000)  # Inside workspace (4 lines)
    heights.append(580)   # Perimeter of work opening (2 lines)
    return heights


def format_bckq_table0(table):
    """
    Formats Table 0 of BCKQ (Báo cáo kết quả) strictly adhering to QUATEST 3 rules:
    - Vertical alignment: Align Top (vAlign="top") for ALL cells.
    - Paragraph spacing: layout/spacing Before 3pt, After 3pt (w:before="60" w:after="60").
    - Typography: Font size 13 (Times New Roman 13pt) for all text and labels.
    """
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            old_v = tcPr.find(f"{{{W_NS}}}vAlign")
            if old_v is not None:
                tcPr.remove(old_v)
            tcPr.append(parse_xml(f'<w:vAlign xmlns:w="{W_NS}" w:val="top"/>'))
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.0
                pPr = p._p.get_or_add_pPr()
                sp = pPr.find(f"{{{W_NS}}}spacing")
                if sp is not None:
                    pPr.remove(sp)
                pPr.append(parse_xml(f'<w:spacing xmlns:w="{W_NS}" w:before="60" w:after="60" w:line="240" w:lineRule="auto"/>'))
                for r_run in p.runs:
                    r_run.font.name = "Times New Roman"
                    r_run.font.size = Pt(13)


def insert_image_to_cell(
    cell,
    image_path: str,
    width=PHOTO_WIDTH_2COL,
    caption: str = None,
    font_size_caption: int = 13,
    align=WD_ALIGN_PARAGRAPH.CENTER,
):
    """
    Inserts an image into a table cell adhering strictly to project rules:
    - Width is 7 - 8 cm (default PHOTO_WIDTH_2COL = Cm(7.5)) for 2-column grids
    - Strictly NO crop: height is NOT set (auto-scaled by python-docx to maintain aspect ratio)
    - Any <a:srcRect> crop attributes are cleared
    - Paragraph is center-aligned with 0 space before/after
    - Optional caption in font size 13 italic, centered
    """
    for p in list(cell.paragraphs):
        p._p.getparent().remove(p._p)

    p_img = cell.add_paragraph()
    p_img.alignment = align
    p_img.paragraph_format.space_before = Pt(0)
    p_img.paragraph_format.space_after = Pt(0)
    p_img.paragraph_format.line_spacing = 1.0

    if image_path and os.path.isfile(image_path):
        r_img = p_img.add_run()
        # Strictly set ONLY width to preserve aspect ratio without crop
        r_img.add_picture(image_path, width=width)
        # Strip any crop attributes from <a:srcRect>
        for sr in r_img._r.findall(f".//{{{A_NS}}}srcRect"):
            sr.attrib.clear()

    if caption:
        p_cap = cell.add_paragraph()
        set_paragraph_text(
            p_cap,
            caption,
            italic=True,
            font_size=font_size_caption,
            align=align,
        )



def merge_stt_column_child_rows(
    table,
    stt_col_idx: int = 0,
    start_row: int = 1,
    font_size: int = 13,
    set_uniform_spacing: bool = True,
):
    """
    Quy chuẩn tổng quát: 'Với chỉ tiêu có nhiều dòng con, thì merge cell STT cho các dòng con'
    Tự động quét các hàng của bảng bắt đầu từ start_row (sau hàng tiêu đề) tại cột stt_col_idx.
    Khi một chỉ tiêu có số STT và có các dòng con bên dưới (ô STT bị rỗng), tự động thực hiện
    gộp ô dọc (vertical merge <w:vMerge>) từ dòng đầu đến dòng con cuối cùng của chỉ tiêu đó.
    Căn giữa hoàn toàn theo cả chiều ngang (center) và chiều đứng (<w:vAlign w:val="center"/>).
    Các ô dữ liệu bình thường khác giữ nguyên valign='top'.
    """
    total_rows = len(table.rows)
    if total_rows <= start_row:
        return

    # 1. Identify merge groups (start_idx, end_idx, stt_text)
    groups = []
    curr_start = None
    curr_stt = ""

    for r_idx in range(start_row, total_rows):
        row = table.rows[r_idx]
        if len(row.cells) <= stt_col_idx:
            if curr_start is not None:
                if (r_idx - 1) > curr_start:
                    groups.append((curr_start, r_idx - 1, curr_stt))
                curr_start = None
            continue

        # Check if row is horizontally merged across all columns (e.g. diagram/image banner)
        if len(row.cells) > 1 and row.cells[0]._tc == row.cells[-1]._tc:
            if curr_start is not None:
                if (r_idx - 1) > curr_start:
                    groups.append((curr_start, r_idx - 1, curr_stt))
                curr_start = None
            continue

        # Check if entire row is empty
        row_has_text = any(c.text.strip() for c in row.cells)
        if not row_has_text:
            if curr_start is not None:
                if (r_idx - 1) > curr_start:
                    groups.append((curr_start, r_idx - 1, curr_stt))
                curr_start = None
            continue

        cell = row.cells[stt_col_idx]
        cell_txt = cell.text.strip()

        if cell_txt:
            # Starts a new potential group
            if curr_start is not None and (r_idx - 1) > curr_start:
                groups.append((curr_start, r_idx - 1, curr_stt))
            curr_start = r_idx
            curr_stt = cell_txt
        else:
            # Empty STT cell: if curr_start is active, this row is a child row
            pass

    if curr_start is not None and (total_rows - 1) > curr_start:
        groups.append((curr_start, total_rows - 1, curr_stt))

    # 2. Merge each group vertically
    for s_idx, e_idx, stt in groups:
        c_merged = table.cell(s_idx, stt_col_idx).merge(table.cell(e_idx, stt_col_idx))
        set_cell_content(
            c_merged,
            stt,
            bold=False,
            font_size=font_size,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            valign="center",
            space_before=3,
            space_after=3,
        )

    # 3. Ensure vertically merged cells have vAlign='center' and other cells retain vAlign='top'
    for tr in table._tbl.tr_lst[start_row:]:
        for tc in tr.tc_lst:
            vMerge = tc.find(f".//{{{W_NS}}}vMerge")
            is_restart = vMerge is not None and vMerge.get(f"{{{W_NS}}}val") == "restart"
            tcPr = tc.get_or_add_tcPr()
            old_v = tcPr.find(f"{{{W_NS}}}vAlign")
            if old_v is not None:
                tcPr.remove(old_v)
            val_str = "center" if is_restart else "top"
            tcPr.append(parse_xml(f'<w:vAlign xmlns:w="{W_NS}" w:val="{val_str}"/>'))

    # 4. Uniform snug spacing (30 twips = 1.5pt) to prevent accidental overflow
    if set_uniform_spacing:
        for tr in table._tbl.tr_lst[start_row:]:
            for p_elem in tr.xpath(".//w:p"):
                pPr = p_elem.get_or_add_pPr()
                old_sp = pPr.find(f"{{{W_NS}}}spacing")
                if old_sp is not None:
                    pPr.remove(old_sp)
                pPr.append(parse_xml(f'<w:spacing xmlns:w="{W_NS}" w:before="30" w:after="30" w:line="240" w:lineRule="auto"/>'))
