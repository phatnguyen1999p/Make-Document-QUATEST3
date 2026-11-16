# -*- coding: utf-8 -*-
"""
Cell and text styling utilities strictly adhering to QUATEST 3 Typography Rules.
Enforces Times New Roman, specified font size (13 for body, 10 for table notes),
exact padding and alignment.
"""

from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def set_cell_shading(cell, fill_hex: str):
    """Sets background shading fill color in hex (e.g. 'D9D9D9' or 'E7E6E6')."""
    tcPr = cell._tc.get_or_add_tcPr()
    existing_shd = tcPr.find(f"{{{W_NS}}}shd")
    if existing_shd is not None:
        tcPr.remove(existing_shd)
    if fill_hex:
        shd = parse_xml(f'<w:shd xmlns:w="{W_NS}" w:fill="{fill_hex}"/>')
        tcPr.append(shd)


def set_cell_margins(cell, top: int = 40, bottom: int = 40, left: int = 80, right: int = 80):
    """Sets cell internal padding (margins) in dxa (twips)."""
    tcPr = cell._tc.get_or_add_tcPr()
    old_tcMar = tcPr.find(f"{{{W_NS}}}tcMar")
    if old_tcMar is not None:
        tcPr.remove(old_tcMar)
    tcMar = parse_xml(
        f'<w:tcMar xmlns:w="{W_NS}">'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def set_cell_content(
    cell,
    text: str,
    bold: bool = False,
    italic: bool = False,
    align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER,
    font_size: float = 11,
    fill_hex: str = None,
    keep_next: bool = False,
    valign: str = "top",
    space_before: float = 0,
    space_after: float = 0,
    line_spacing: float = 1.0
):
    """
    Sets text in a table cell, enforcing Times New Roman, specified vertical alignment ('top'/'center'), and exact styling.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    old_v = tcPr.find(f"{{{W_NS}}}vAlign")
    if old_v is not None:
        tcPr.remove(old_v)
    if valign:
        tcPr.append(parse_xml(f'<w:vAlign xmlns:w="{W_NS}" w:val="{valign}"/>'))

    # Clean previous paragraphs
    for p in cell.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p = cell.paragraphs[0]
    p.text = ""
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing

    t_clean = text.strip()
    if keep_next or t_clean.startswith("Đánh giá") or t_clean.startswith("Assessment"):
        pPr = p._p.get_or_add_pPr()
        if pPr.find(f"{{{W_NS}}}keepNext") is None:
            pPr.append(parse_xml(f'<w:keepNext xmlns:w="{W_NS}"/>'))

    lines = text.split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            p.add_run().add_break()
        run = p.add_run(line)
        run.bold = bold
        run.italic = italic
        run.font.name = "Times New Roman"
        run.font.size = Pt(font_size)
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn("w:ascii"), "Times New Roman")
        rFonts.set(qn("w:hAnsi"), "Times New Roman")
        rFonts.set(qn("w:cs"), "Times New Roman")

    if fill_hex:
        set_cell_shading(cell, fill_hex)
