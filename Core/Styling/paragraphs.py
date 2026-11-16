# -*- coding: utf-8 -*-
"""
Paragraph styling and assessment formatting for QUATEST 3 Documents.
Enforces:
- Font size 13 Times New Roman for body paragraphs and headings
- Bold + Italic formatting for assessment keywords (phù hợp, không phù hợp, comply, complied, did not comply)
- keepNext / keepLines to prevent orphan conclusions
"""

import re
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from .styles import W_NS


def set_paragraph_text(p, text: str, bold: bool = False, italic: bool = False, font_size: float = 13, align=None, keep_next: bool = False):
    """Sets text in a body paragraph, enforcing Times New Roman 13pt and spacing."""
    p.text = ""
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15

    t_clean = text.strip()
    if keep_next or t_clean.startswith("Đánh giá") or t_clean.startswith("Assessment"):
        pPr = p._p.get_or_add_pPr()
        if pPr.find(f"{{{W_NS}}}keepNext") is None:
            pPr.append(parse_xml(f'<w:keepNext xmlns:w="{W_NS}"/>'))
        if pPr.find(f"{{{W_NS}}}keepLines") is None:
            pPr.append(parse_xml(f'<w:keepLines xmlns:w="{W_NS}"/>'))

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


def set_assessment_paragraph(p, text: str, font_size: float = 13, align=None, keep_next: bool = True):
    """
    Sets assessment paragraph text with Times New Roman 13pt and automatically formats
    assessment keywords in bold AND italic (bold=True, italic=True) per AGENTS.md Section 3:
      - Tiếng Việt: ***phù hợp***, ***không phù hợp***
      - Tiếng Anh: ***comply***, ***complied***, ***did not comply***
    """
    p.text = ""
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15

    pPr = p._p.get_or_add_pPr()
    if pPr.find(f"{{{W_NS}}}keepNext") is None and keep_next:
        pPr.append(parse_xml(f'<w:keepNext xmlns:w="{W_NS}"/>'))
    if pPr.find(f"{{{W_NS}}}keepLines") is None:
        pPr.append(parse_xml(f'<w:keepLines xmlns:w="{W_NS}"/>'))

    # Markdown-style emphasis markers (***phù hợp***) are only hints: never print them literally.
    text = text.replace("***", "")

    # Regex pattern to match evaluation keywords
    kw_pattern = re.compile(
        r"(không phù hợp|phù hợp|did not comply|complied|comply)",
        re.IGNORECASE
    )

    tokens = kw_pattern.split(text)
    for token in tokens:
        if not token:
            continue
        is_kw = bool(kw_pattern.fullmatch(token.strip()))
        run = p.add_run(token)
        run.bold = is_kw
        run.italic = is_kw
        run.font.name = "Times New Roman"
        run.font.size = Pt(font_size)
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn("w:ascii"), "Times New Roman")
        rFonts.set(qn("w:hAnsi"), "Times New Roman")
        rFonts.set(qn("w:cs"), "Times New Roman")
