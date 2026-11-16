# -*- coding: utf-8 -*-
"""
Document elements utilities for QUATEST 3 Documents:
- Ending Line: Standard QUATEST 3 thinThick 4.5pt centered double line (TBKQ only)
- Header Synchronization: Synchronizes cert number and issue date across all sections
- Table 0 Cover Page: Standardized customer, dossier, equipment, and signature blocks
- Highlight Removal: Cleans all leftover yellow highlights
"""

import re
from docx.shared import Pt, Cm
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .styles import W_NS, set_cell_content
from ..Models.dossier import Dossier, inspection_scope, inspection_procedure, bsc_class_of
from ..Tools.standards import cite_standards


def strip_all_highlights(doc):
    """Strips all yellow highlights (<w:highlight>) from body, headers, and footers."""
    # Body highlights
    for h in doc.element.body.findall(f".//{{{W_NS}}}highlight"):
        h.getparent().remove(h)

    # Headers and footers highlights
    for section in doc.sections:
        for h in section.header._element.findall(f".//{{{W_NS}}}highlight"):
            h.getparent().remove(h)
        for h in section.footer._element.findall(f".//{{{W_NS}}}highlight"):
            h.getparent().remove(h)


_GRAPHIC_TAGS = (qn("w:drawing"), qn("w:pict"), qn("w:object"),
                 "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent")


def is_blank_paragraph(p_el) -> bool:
    """Đoạn trống thật: không chữ, không hình/đường kẻ, không ngắt section/trang.
    Không xét chuỗi `p.xml` ('drawing' có sẵn trong khai báo namespace `drawingml` của mọi đoạn — lỗi khiến
    `_drop_trailing_empty` của các script hồ sơ không xóa được đoạn nào, 127215/120423 10/11/2026)."""
    if p_el is None or p_el.tag != qn("w:p"):
        return False
    if "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip():
        return False
    for tag in _GRAPHIC_TAGS:
        if next(p_el.iter(tag), None) is not None:
            return False
    if p_el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
        return False
    if any(br.get(qn("w:type")) == "page" for br in p_el.iter(qn("w:br"))):
        return False
    return True


def drop_trailing_empty_paragraphs(doc) -> int:
    """Xóa mọi đoạn trống ở cuối thân văn bản (sau câu kết luận cuối) — đường kết thúc phải sát dòng cuối
    (AGENTS mục 3, 5). Trả số đoạn đã xóa."""
    body = doc.element.body
    n = 0
    el = body[-1] if len(body) else None
    if el is not None and el.tag == qn("w:sectPr"):
        el = el.getprevious()
    while is_blank_paragraph(el):
        prev = el.getprevious()
        body.remove(el)
        n += 1
        el = prev
    return n


def drop_section_gaps(doc, space_before_pt: float = 6) -> int:
    """Đoạn trống sót giữa 2 đoạn chữ (vd giữa câu Đánh giá và tiêu đề mục kế, chỗ template xóa mục nhiệt ẩm /
    HEPA / scan…) -> bỏ, đoạn chữ sau đặt space_before 6pt (AGENTS mục 3, mẫu `_drop_section_gaps` 138143).
    Đoạn trống đứng cạnh bảng (giữa bảng kết quả và bảng Ghi chú, sau bảng ảnh/sơ đồ) giữ nguyên.
    Chỉ xét phần sau ngắt section đầu tiên (phụ lục) — trang bìa không đụng tới."""
    body = doc.element.body
    started = body.find(f".//{qn('w:pPr')}/{qn('w:sectPr')}") is None
    n = 0
    for el in list(body):
        if not started:
            if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
                started = True
            continue
        if el.getparent() is None or not is_blank_paragraph(el):
            continue
        prev = el.getprevious()
        while is_blank_paragraph(prev):
            prev = prev.getprevious()
        nxt = el.getnext()
        while is_blank_paragraph(nxt):
            nxt = nxt.getnext()
        if prev is None or nxt is None or prev.tag != qn("w:p") or nxt.tag != qn("w:p"):
            continue
        if prev.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            continue
        if not "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip():
            continue
        body.remove(el)
        n += 1
        from docx.text.paragraph import Paragraph
        pf = Paragraph(nxt, None).paragraph_format
        if pf.space_before is None or pf.space_before < Pt(space_before_pt):
            pf.space_before = Pt(space_before_pt)
    return n


def add_document_ending_line(doc, docpr_id: int = 900999):
    """
    Appends the official QUATEST 3 ending double line (thinThick, 4.5pt strokeweight, black)
    at the very end of the TBKQ document, strictly adhering to AGENTS.md Section 5:
    - Pure horizontal line (cy="0")
    - Centered between margins (<wp:align>center</wp:align>)
    - Placed in a dedicated empty paragraph right under the last line (~6pt gap, user 30/10/2026), never overlapping text
    - Mọi đoạn trống trước đó bị xóa (127215/120423 10/11/2026: 6 đoạn trống của template đẩy đường kẻ xa dòng cuối)
    - TBKQ ONLY: Never add to BCKQ!
    """
    body = doc.element.body
    # Remove misplaced ending lines, then trailing empty paragraphs
    for p in list(body.findall(f"{{{W_NS}}}p")):
        p_xml = p.xml
        if "thinThick" in p_xml or "Straight Connector" in p_xml:
            p.getparent().remove(p)
    drop_trailing_empty_paragraphs(doc)

    # Create dedicated paragraph for ending line
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Đường kẻ nằm sát dòng cuối (người dùng 30/10/2026): không space_before, offset 6pt dưới đỉnh đoạn.
    # Đoạn chứa đường kẻ cao đúng 1pt (exact) -> không bao giờ tự rớt sang trang mới chỉ có đường kết thúc.
    p_end.paragraph_format.space_before = Pt(0)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.paragraph_format.line_spacing = Pt(1)

    ending_line_xml = f"""<w:r xmlns:w="{W_NS}"
xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:w10="urn:schemas-microsoft-com:office:word"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:AlternateContent>
    <mc:Choice Requires="wps">
      <w:drawing>
        <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="margin">
            <wp:align>center</wp:align>
          </wp:positionH>
          <wp:positionV relativeFrom="paragraph">
            <wp:posOffset>76200</wp:posOffset>
          </wp:positionV>
          <wp:extent cx="3696970" cy="0"/>
          <wp:effectExtent l="0" t="19050" r="55880" b="38100"/>
          <wp:wrapNone/>
          <wp:docPr id="{docpr_id}" name="Straight Connector {docpr_id}"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
          </wp:cNvGraphicFramePr>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:cNvCnPr>
                  <a:cxnSpLocks noChangeShapeType="1"/>
                </wps:cNvCnPr>
                <wps:spPr bwMode="auto">
                  <a:xfrm flipV="1">
                    <a:off x="0" y="0"/>
                    <a:ext cx="3696970" cy="0"/>
                  </a:xfrm>
                  <a:prstGeom prst="line">
                    <a:avLst/>
                  </a:prstGeom>
                  <a:noFill/>
                  <a:ln w="57150" cmpd="thinThick">
                    <a:solidFill>
                      <a:srgbClr val="000000"/>
                    </a:solidFill>
                    <a:round/>
                    <a:headEnd/>
                    <a:tailEnd/>
                  </a:ln>
                </wps:spPr>
                <wps:bodyPr/>
              </wps:wsp>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing>
    </mc:Choice>
    <mc:Fallback>
      <w:pict>
        <v:line from="0pt,6pt" to="291.1pt,6pt" strokeweight="4.5pt" style="position:absolute;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:paragraph;mso-position-vertical-relative:paragraph;z-index:251659264;">
          <v:stroke linestyle="thinThick" color="black"/>
        </v:line>
      </w:pict>
    </mc:Fallback>
  </mc:AlternateContent>
</w:r>"""
    p_end._p.append(parse_xml(ending_line_xml))


def update_headers(doc, cert_no: str, cert_date: str, is_en: bool = False):
    """
    Synchronizes certificate number and issue date across Section 0 and Section 1 headers,
    enforcing Times New Roman 12pt, Bold, Centered, Non-italic per AGENTS.md Section 2.
    """
    cert_text = f"No: {cert_no} – Date: {cert_date}" if is_en else f"Số: {cert_no} - Ngày {cert_date}"

    for section in doc.sections:
        # Header paragraphs
        for p in section.header.paragraphs:
            txt = p.text
            if "Số:" in txt or "No:" in txt or "Ngày" in txt or "Date" in txt or "……" in txt or "..." in txt:
                p.text = ""
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pPr = p._p.get_or_add_pPr()
                jc = pPr.find(f"{{{W_NS}}}jc")
                if jc is not None:
                    pPr.remove(jc)
                pPr.append(parse_xml(f'<w:jc xmlns:w="{W_NS}" w:val="center"/>'))

                run = p.add_run(cert_text)
                run.bold = True
                run.italic = False
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
                rPr = run._r.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                rFonts.set(qn("w:ascii"), "Times New Roman")
                rFonts.set(qn("w:hAnsi"), "Times New Roman")
                rFonts.set(qn("w:cs"), "Times New Roman")

        # Header tables (if any)
        for tbl in section.header.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        txt = p.text
                        if "Số:" in txt or "No:" in txt or "Ngày" in txt or "Date" in txt or "……" in txt or "..." in txt:
                            p.text = ""
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            pPr = p._p.get_or_add_pPr()
                            jc = pPr.find(f"{{{W_NS}}}jc")
                            if jc is not None:
                                pPr.remove(jc)
                            pPr.append(parse_xml(f'<w:jc xmlns:w="{W_NS}" w:val="center"/>'))

                            run = p.add_run(cert_text)
                            run.bold = True
                            run.italic = False
                            run.font.name = "Times New Roman"
                            run.font.size = Pt(12)
                            rPr = run._r.get_or_add_rPr()
                            rFonts = rPr.get_or_add_rFonts()
                            rFonts.set(qn("w:ascii"), "Times New Roman")
                            rFonts.set(qn("w:hAnsi"), "Times New Roman")
                            rFonts.set(qn("w:cs"), "Times New Roman")

        # Strip highlights in header
        for h in section.header._element.findall(f".//{{{W_NS}}}highlight"):
            h.getparent().remove(h)


def normalize_company_name(name: str) -> str:
    """
    Ensures company name is not fully capitalized, conforming to proper Vietnamese casing
    (e.g., 'Công ty TNHH Giải pháp Y sinh ABT' instead of 'CÔNG TY TNHH GIẢI PHÁP Y SINH ABT').
    """
    if not name:
        return name
    name = name.replace("CÔNG TY TNHH GIẢI PHÁP Y SINH ABT", "Công ty TNHH Giải pháp Y sinh ABT")
    name = name.replace("ABT BIOMEDICAL SOLUTIONS COMPANY LIMITED", "ABT Biomedical Solutions Company Limited")
    if name.startswith("CÔNG TY"):
        name = name.replace("CÔNG TY TNHH", "Công ty TNHH")
        name = name.replace("CÔNG TY CỔ PHẦN", "Công ty Cổ phần")
        name = name.replace("CÔNG TY CP", "Công ty CP")
        name = name.replace("CÔNG TY", "Công ty")
    return name


def populate_cleanroom_table0(doc, dossier: Dossier, is_en: bool = False):
    """
    Populates 16-row Table 0 for Cleanroom certificates strictly conforming to AGENTS.md:
    - Preserves all 16 rows and exact OpenXML formatting.
    - Row 4 contains bare application code (NO suffix /1, /2, /E), bold=False.
    - Row 8 preserves inspection procedure.
    - Row 10 contains complete equipment list.
    - Row 12 contains applicable standards and customer criteria.
    - Row 13 results pointer ('Xem Phụ lục đính kèm' / 'See the attached Appendix').
    - Row 15 signatures with 3cm space and horizontal title alignment.
    """
    t0 = doc.tables[0]
    client_name = normalize_company_name(dossier.client.name_en if is_en else dossier.client.name_vn)
    client_addr = dossier.client.address_en if is_en else dossier.client.address_vn

    scope = dossier.scope_en if is_en else dossier.scope_vn
    if dossier.dossier_type.lower() in ("bsc", "laf", "laf_bsc", "fumehood"):
        scope = inspection_scope(scope, is_en)  # form tủ: 'Giám định tủ ...'
    loc_desc = normalize_company_name(dossier.location_desc_en if is_en else dossier.location_desc_vn)
    if not loc_desc:
        loc_desc = f"{client_name} - {client_addr}"

    if is_en:
        sample_time = dossier.sample_time_en or f"From {dossier.sample_date}"
    else:
        sample_time = dossier.sample_time_vn or f"Từ ngày {dossier.sample_date}"
        m_date = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', sample_time)
        if "từ" in sample_time.lower() and "đến" in sample_time.lower() and m_date:
            sample_time = f"Từ ngày {m_date.group(1)}"
        elif sample_time.lower().startswith("từ ") and not sample_time.lower().startswith("từ ngày "):
            sample_time = sample_time.replace("Từ ", "Từ ngày ").replace("từ ", "từ ngày ")
        elif not sample_time.lower().startswith("từ"):
            if m_date:
                sample_time = f"Từ ngày {m_date.group(1)}"
            else:
                sample_time = f"Từ ngày {dossier.sample_date}"

    if len(getattr(dossier, "inspection_dates", None) or []) > 1:     # nhiều ngày hiện trường (AGENTS mục 2)
        from ..Models.dossier import dates_phrase
        sample_time = f"From {dates_phrase(dossier.inspection_dates, True)}" if is_en else             f"Từ ngày {dates_phrase(dossier.inspection_dates, False)}"

    procedure = inspection_procedure((dossier.procedure_en if is_en else dossier.procedure_vn) or "QUATEST 3 3030:2025, N2.TC.PS.01", bsc_class_of(dossier))

    val_col = 3 if is_en else 2

    # Row 1: Client name
    set_cell_content(t0.rows[1].cells[val_col], client_name, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 2: Client address
    set_cell_content(t0.rows[2].cells[val_col], client_addr, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 4: Bare application code (MUST NOT be bold, MUST NOT have /1, /2, /E per AGENTS Rule 2)
    # Formatted strictly as "<code bare> ngày <date>" (VN) or "<code bare> dated on <date>" (EN) with exactly 1 space
    r4 = t0.rows[4]
    bare_code = dossier.code.strip()
    raw_date = getattr(dossier, "reg_date", None) or dossier.sample_date or dossier.cert_date
    if hasattr(raw_date, "strftime"):
        date_str = raw_date.strftime("%d/%m/%Y")
    else:
        date_str = str(raw_date or "").strip()
    if is_en:
        if len(r4.cells) > 1 and r4.cells[1]._tc != r4.cells[-1]._tc:
            c_target = r4.cells[1].merge(r4.cells[-1])
        else:
            c_target = r4.cells[1] if len(r4.cells) > 1 else r4.cells[0]
        full_text = f"{bare_code} dated on {date_str}"
        set_cell_content(c_target, full_text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    else:
        start_idx = 3 if len(r4.cells) > 3 else (1 if len(r4.cells) > 1 else 0)
        if len(r4.cells) > start_idx and r4.cells[start_idx]._tc != r4.cells[-1]._tc:
            c_target = r4.cells[start_idx].merge(r4.cells[-1])
        else:
            c_target = r4.cells[start_idx]
        full_text = f"{bare_code} ngày {date_str}"
        set_cell_content(c_target, full_text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 5: Request / Scope
    set_cell_content(t0.rows[5].cells[val_col], scope, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 6: Testing time
    set_cell_content(t0.rows[6].cells[val_col], sample_time, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 7: Testing location
    set_cell_content(t0.rows[7].cells[val_col], loc_desc, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 8: Procedure / Standard
    set_cell_content(t0.rows[8].cells[val_col], procedure, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 10: Equipment (Font size 13)
    eq_lines = []
    for eq in dossier.equipments:
        eq_name = eq.name_en if is_en else eq.name_vn
        if eq_name:
            s = f"- {eq_name}"
            if not s.endswith(";"):
                s += ";"
            eq_lines.append(s)
    if eq_lines:
        eq_lines[-1] = eq_lines[-1].rstrip(";") + "."
        eq_content = "\n".join(eq_lines)
    else:
        eq_content = "………"
    set_cell_content(t0.rows[10].cells[0], eq_content, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 12: Applicable standards & evaluation criteria
    st_content = (dossier.standards_en if is_en else dossier.standards_vn) or (
        f"- ISO 14644-1:2015 Cleanrooms and associated controlled environments - Part 1: Classification of air cleanliness by particle concentration;\n- Technical specifications for temperature - humidity, differential pressure, supply airflow rate and air change rate provided by {client_name}."
        if is_en else
        f"- ISO 14644-1:2015 Cleanrooms and associated controlled environments — Part 1: Classification of air cleanliness by particle concentration;\n- Mức quy định đối với nhiệt độ - độ ẩm, chênh áp, lưu lượng khí cấp và số lần trao đổi khí được cung cấp bởi {client_name}."
    )
    st_content = cite_standards(st_content)  # tên đầy đủ của tiêu chuẩn
    set_cell_content(t0.rows[12].cells[0], st_content, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Form goc danh dau dong '-' tu dong (numPr) cho 2 o nay; noi dung da tu ghi '- ' -> bo numPr, khong de '- - ' (E23).
    # Bo luon thut treo (hanging) cua numPr de dong '-' dau thang hang voi cac dong '-' sau (cung thut trai).
    for ri, content in ((10, eq_content), (12, st_content)):
        if content.lstrip().startswith("-"):
            for p in t0.rows[ri].cells[0].paragraphs:
                pPr = p._p.pPr
                if pPr is not None:
                    for num in pPr.findall(qn("w:numPr")):
                        pPr.remove(num)
                    for ind in pPr.findall(qn("w:ind")):
                        for k in ("w:hanging", "w:firstLine", "w:hangingChars", "w:firstLineChars"):
                            ind.attrib.pop(qn(k), None)

    # Row 13: Results pointer
    results_text = (getattr(dossier, "results_pointer_en", None) if is_en else getattr(dossier, "results_pointer_vn", None)) or ("See the attached Appendix" if is_en else "Xem Phụ lục đính kèm")
    set_cell_content(t0.rows[13].cells[val_col], results_text, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 15: Signatures (minimum 3cm gap for physical signatures, perfectly aligned horizontally)
    last_row = t0.rows[15] if len(t0.rows) > 15 else t0.rows[-1]
    u_sig_cells = []
    for c in last_row.cells:
        if not u_sig_cells or c._tc != u_sig_cells[-1]._tc:
            u_sig_cells.append(c)

    cell_l = u_sig_cells[0]
    cell_r = u_sig_cells[-1]

    # Parse GDV title and name
    gdv_raw = dossier.personnel.sign_gdv_en if is_en else dossier.personnel.sign_gdv_vn
    lines_l = [l.strip() for l in gdv_raw.split("\n") if l.strip()]
    if len(lines_l) >= 2:
        title_l = lines_l[0]
        name_l = lines_l[-1]
    elif len(lines_l) == 1:
        title_l = "INSPECTOR" if is_en else "GIÁM ĐỊNH VIÊN"
        name_l = lines_l[0]
    else:
        title_l = "INSPECTOR" if is_en else "GIÁM ĐỊNH VIÊN"
        name_l = "Nguyen Xuan Nam" if is_en else "Nguyễn Xuân Nam"

    # Parse Leader title and name
    leader_raw = dossier.personnel.sign_leader_en if is_en else dossier.personnel.sign_leader_vn
    lines_r = [l.strip() for l in leader_raw.split("\n") if l.strip()]
    if len(lines_r) >= 3:
        title_r1 = lines_r[0]
        title_r2 = lines_r[1]
        name_r = lines_r[-1]
    elif len(lines_r) == 2:
        title_r1 = lines_r[0]
        title_r2 = ""
        name_r = lines_r[1]
    elif len(lines_r) == 1:
        title_r1 = "FOR DIRECTOR" if is_en else "KT. GIÁM ĐỐC"
        title_r2 = "VICE DIRECTOR" if is_en else "PHÓ GIÁM ĐỐC"
        name_r = lines_r[0]
    else:
        title_r1 = "FOR DIRECTOR" if is_en else "KT. GIÁM ĐỐC"
        title_r2 = "VICE DIRECTOR" if is_en else "PHÓ GIÁM ĐỐC"
        name_r = "Le Dang Huyen" if is_en else "Lê Đăng Huyền"

    # Setup Left Cell
    for p in cell_l.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p0_l = cell_l.paragraphs[0]
    p0_l.text = ""
    p0_l.paragraph_format.space_before = Pt(0)
    p0_l.paragraph_format.space_after = Pt(0)
    p0_l.paragraph_format.line_spacing = 1.0

    p1_l = cell_l.add_paragraph()
    p1_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1_l.paragraph_format.space_before = Pt(0)
    p1_l.paragraph_format.space_after = Pt(0)
    p1_l.paragraph_format.line_spacing = 1.0
    r1_l = p1_l.add_run(title_l)
    r1_l.bold = True
    r1_l.font.name = "Times New Roman"
    r1_l.font.size = Pt(13)

    p_name_l = cell_l.add_paragraph()
    p_name_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name_l.paragraph_format.space_before = Cm(3.0)
    p_name_l.paragraph_format.space_after = Pt(0)
    p_name_l.paragraph_format.line_spacing = 1.0
    r_nl = p_name_l.add_run(name_l)
    r_nl.bold = True
    r_nl.font.name = "Times New Roman"
    r_nl.font.size = Pt(13)

    # Setup Right Cell
    for p in cell_r.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p0_r = cell_r.paragraphs[0]
    p0_r.text = ""
    p0_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0_r.paragraph_format.space_before = Pt(0)
    p0_r.paragraph_format.space_after = Pt(0)
    p0_r.paragraph_format.line_spacing = 1.0
    r0_r = p0_r.add_run(title_r1)
    r0_r.bold = True
    r0_r.font.name = "Times New Roman"
    r0_r.font.size = Pt(13)

    if title_r2:
        p1_r = cell_r.add_paragraph()
        p1_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1_r.paragraph_format.space_before = Pt(0)
        p1_r.paragraph_format.space_after = Pt(0)
        p1_r.paragraph_format.line_spacing = 1.0
        r1_r = p1_r.add_run(title_r2)
        r1_r.bold = True
        r1_r.font.name = "Times New Roman"
        r1_r.font.size = Pt(13)

    p_name_r = cell_r.add_paragraph()
    p_name_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name_r.paragraph_format.space_before = Cm(3.0)
    p_name_r.paragraph_format.space_after = Pt(0)
    p_name_r.paragraph_format.line_spacing = 1.0
    r_nr = p_name_r.add_run(name_r)
    r_nr.bold = True
    r_nr.font.name = "Times New Roman"
    r_nr.font.size = Pt(13)

    if len(u_sig_cells) >= 3:
        for p in u_sig_cells[1].paragraphs[1:]:
            p._p.getparent().remove(p._p)
        u_sig_cells[1].paragraphs[0].text = ""
        u_sig_cells[1].paragraphs[0].paragraph_format.space_before = Pt(0)
        u_sig_cells[1].paragraphs[0].paragraph_format.space_after = Pt(0)

    # Uniform 3pt top/bottom paragraph spacing for every line of Table 0 (even text rows)
    apply_table0_spacing(t0)


def apply_table0_spacing(table, before_pt: float = 3, after_pt: float = 3):
    """
    Makes the text rows of Table 0 (TBKQ cover table) evenly spaced:
    Layout/Spacing Before = After = 3pt, single line spacing, for EVERY paragraph of the table.
    The only exception is the signature-name paragraph, whose "before" gap (3 cm) is mandated
    by the signing-space rule and is therefore preserved (its "after" is still 3pt, symmetric
    on both sides so the signature titles / names stay perfectly aligned).
    Empty spacer paragraphs get a 13pt paragraph mark so that they have exactly the same line
    height as the 13pt titles next to them (keeps INSPECTOR level with VICE DIRECTOR).
    """
    seen = []
    for row in table.rows:
        for cell in row.cells:
            if any(cell._tc is s for s in seen):
                continue
            seen.append(cell._tc)
            for p in cell.paragraphs:
                pf = p.paragraph_format
                cur_before = pf.space_before
                if not (cur_before is not None and cur_before >= Cm(2.9)):
                    pf.space_before = Pt(before_pt)
                pf.space_after = Pt(after_pt)
                pf.line_spacing = 1.0
                if not p.text.strip():
                    pPr = p._p.get_or_add_pPr()
                    rPr = pPr.find(qn("w:rPr"))
                    if rPr is None:
                        rPr = parse_xml(f'<w:rPr xmlns:w="{W_NS}"/>')
                        pPr.append(rPr)
                    for tag in ("w:sz", "w:szCs"):
                        el = rPr.find(qn(tag))
                        if el is None:
                            el = parse_xml(f'<{tag} xmlns:w="{W_NS}" w:val="26"/>')
                            rPr.append(el)
                        else:
                            el.set(qn("w:val"), "26")


def update_table0(doc, dossier: Dossier, is_en: bool = False):
    """
    Updates cover page certificate table (Table 0) strictly conforming to AGENTS.md:
    - Row 4 contains bare application code (NO suffix /1, /2, /E), bold=False.
    - Synchronized customer name, address, time, scope, equipment list, standards, signatures.
    """
    t0 = doc.tables[0]
    if dossier.dossier_type.lower() in ("phong sach", "cleanroom") or len(t0.rows) >= 16:
        return populate_cleanroom_table0(doc, dossier, is_en=is_en)
    client_name = normalize_company_name(dossier.client.name_en if is_en else dossier.client.name_vn)
    client_addr = dossier.client.address_en if is_en else dossier.client.address_vn

    scope = dossier.scope_en if is_en else dossier.scope_vn
    if dossier.dossier_type.lower() in ("bsc", "laf", "laf_bsc", "fumehood"):
        scope = inspection_scope(scope, is_en)  # form tủ: 'Giám định tủ ...'
    loc_desc = normalize_company_name(dossier.location_desc_en if is_en else dossier.location_desc_vn)
    if not loc_desc:
        loc_desc = f"{client_name} - {client_addr}"

    if is_en:
        sample_time = dossier.sample_time_en or f"From {dossier.sample_date}"
    else:
        sample_time = dossier.sample_time_vn or f"Từ ngày {dossier.sample_date}"
        m_date = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', sample_time)
        if "từ" in sample_time.lower() and "đến" in sample_time.lower() and m_date:
            sample_time = f"Từ ngày {m_date.group(1)}"
        elif sample_time.lower().startswith("từ ") and not sample_time.lower().startswith("từ ngày "):
            sample_time = sample_time.replace("Từ ", "Từ ngày ").replace("từ ", "từ ngày ")
        elif not sample_time.lower().startswith("từ"):
            if m_date:
                sample_time = f"Từ ngày {m_date.group(1)}"
            else:
                sample_time = f"Từ ngày {dossier.sample_date}"

    # Build equipment list lines
    eq_lines = []
    for eq in dossier.equipments:
        eq_name = eq.name_en if is_en else eq.name_vn
        s = f"- {eq_name}"
        if not s.endswith(";"):
            s += ";"
        eq_lines.append(s)
    if eq_lines:
        eq_lines[-1] = eq_lines[-1].rstrip(";") + "."
    eq_content = "\n".join(eq_lines) if eq_lines else ("………" if not is_en else "………")

    standards = cite_standards(dossier.standards_en if is_en else dossier.standards_vn)  # tên đầy đủ của tiêu chuẩn
    basis = dossier.basis_en if is_en else dossier.basis_vn

    if hasattr(dossier, "results_pointer_vn") and (dossier.results_pointer_vn or dossier.results_pointer_en):
        results_text = dossier.results_pointer_en if is_en else dossier.results_pointer_vn
    elif dossier.dossier_type.lower() in ("bsc", "fumehood", "laf_bsc"):
        results_text = "See the attached Appendix" if is_en else "Xem Phụ lục đính kèm"
    else:
        results_text = "See from page 2/4 to page 4/4" if is_en else "Xem từ trang 2/4 đến trang 4/4"

    sign_left = dossier.personnel.sign_gdv_en if is_en else dossier.personnel.sign_gdv_vn
    sign_right = dossier.personnel.sign_leader_en if is_en else dossier.personnel.sign_leader_vn

    val_col = 3 if is_en else 2

    # Row 1: Client name
    set_cell_content(t0.rows[1].cells[val_col], client_name, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 2: Client address
    set_cell_content(t0.rows[2].cells[val_col], client_addr, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 4: Bare application code (MUST NOT be bold, MUST NOT have /1, /2, /E per AGENTS Rule 2)
    # Formatted strictly as "<code bare> ngày <date>" (VN) or "<code bare> dated on <date>" (EN) with exactly 1 space
    r4 = t0.rows[4]
    bare_code = dossier.code.strip()
    raw_date = getattr(dossier, "reg_date", None) or dossier.sample_date or dossier.cert_date
    if hasattr(raw_date, "strftime"):
        date_str = raw_date.strftime("%d/%m/%Y")
    else:
        date_str = str(raw_date or "").strip()
    if is_en:
        if len(r4.cells) > 1 and r4.cells[1]._tc != r4.cells[-1]._tc:
            c_target = r4.cells[1].merge(r4.cells[-1])
        else:
            c_target = r4.cells[1] if len(r4.cells) > 1 else r4.cells[0]
        # Re-balance pct width so Word COM never freezes on repair
        tcPr = c_target._tc.get_or_add_tcPr()
        tcW = tcPr.find(f"{{{W_NS}}}tcW")
        if tcW is not None:
            tcW.attrib[f"{{{W_NS}}}w"] = "3182"
        full_text = f"{bare_code} dated on {date_str}"
        set_cell_content(c_target, full_text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    else:
        start_idx = 3 if len(r4.cells) > 3 else (1 if len(r4.cells) > 1 else 0)
        if len(r4.cells) > start_idx and r4.cells[start_idx]._tc != r4.cells[-1]._tc:
            c_target = r4.cells[start_idx].merge(r4.cells[-1])
        else:
            c_target = r4.cells[start_idx]
        # Re-balance pct width so Word COM never freezes on repair
        tcPr = c_target._tc.get_or_add_tcPr()
        tcW = tcPr.find(f"{{{W_NS}}}tcW")
        if tcW is not None:
            tcW.attrib[f"{{{W_NS}}}w"] = "2764"
        full_text = f"{bare_code} ngày {date_str}"
        set_cell_content(c_target, full_text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 5: Request / Scope
    set_cell_content(t0.rows[5].cells[val_col], scope, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 6: Testing time
    set_cell_content(t0.rows[6].cells[val_col], sample_time, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 7: Testing location
    set_cell_content(t0.rows[7].cells[val_col], loc_desc, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 8: Procedure / Standard
    set_cell_content(t0.rows[8].cells[val_col], standards, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
    # Row 10: Equipment (Font size 13)
    set_cell_content(t0.rows[10].cells[0], eq_content, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 13: Results pointer
    set_cell_content(t0.rows[13].cells[val_col], results_text, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Row 11 / 12: Assessment basis
    # For BSC / LAF / single-line basis, basis belongs on Row 11 on the right of colon (val_col), and Row 12 is deleted
    if dossier.dossier_type.lower() in ("bsc", "fumehood", "laf", "laf_bsc") or "\n" not in basis.strip():
        set_cell_content(t0.rows[11].cells[val_col], basis, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)
        if len(t0.rows) > 13:
            t0._tbl.remove(t0.rows[12]._tr)
            # Remove redundant empty spacing row before signatures if present
            if len(t0.rows) > 13 and not any(c.text.strip() for c in t0.rows[13].cells):
                t0._tbl.remove(t0.rows[13]._tr)
    else:
        set_cell_content(t0.rows[12].cells[0], basis, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=13)

    # Last row: Signatures (minimum 3cm gap for physical signatures, perfectly aligned horizontally)
    last_row = t0.rows[-1]
    u_sig_cells = []
    for c in last_row.cells:
        if not u_sig_cells or c._tc != u_sig_cells[-1]._tc:
            u_sig_cells.append(c)

    def _setup_sig_cell(cell, title_line1, title_line2, name):
        for p in cell.paragraphs[1:]:
            p._p.getparent().remove(p._p)
        p0 = cell.paragraphs[0]
        p0.text = ""
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        p0.paragraph_format.line_spacing = 1.0
        r0 = p0.add_run(title_line1)
        r0.bold = True
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(13)

        if title_line2:
            p1 = cell.add_paragraph()
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.paragraph_format.space_before = Pt(0)
            p1.paragraph_format.space_after = Pt(0)
            p1.paragraph_format.line_spacing = 1.0
            r1 = p1.add_run(title_line2)
            r1.bold = True
            r1.font.name = "Times New Roman"
            r1.font.size = Pt(13)

        p_name = cell.add_paragraph()
        p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_name.paragraph_format.space_before = Cm(3.0)  # Exactly 3.0 cm physical blank space for signing (1701 twips)
        p_name.paragraph_format.space_after = Pt(0)
        p_name.paragraph_format.line_spacing = 1.0
        r_name = p_name.add_run(name)
        r_name.bold = True
        r_name.font.name = "Times New Roman"
        r_name.font.size = Pt(13)

    # Parse names and titles
    def _parse_sig(txt):
        lines = [l.strip() for l in txt.split("\n") if l.strip()]
        if not lines:
            return "", "", ""
        if len(lines) == 1:
            return lines[0], "", ""
        if len(lines) == 2:
            return lines[0], "", lines[1]
        return lines[0], lines[1], lines[-1]

    t1_l, t2_l, n_l = _parse_sig(sign_left)
    t1_r, t2_r, n_r = _parse_sig(sign_right)

    # When leader has 2 title lines (e.g. KT. GIÁM ĐỐC + PHÓ GIÁM ĐỐC / FOR DIRECTOR + VICE DIRECTOR),
    # GDV / INSPECTOR must align horizontally with PHÓ GIÁM ĐỐC / VICE DIRECTOR on line 2.
    if t2_r and not t2_l:
        t2_l = t1_l
        t1_l = " "  # Blank spacer on line 1 so line 2 aligns horizontally with PHÓ GIÁM ĐỐC / VICE DIRECTOR

    if len(u_sig_cells) >= 3:
        _setup_sig_cell(u_sig_cells[0], t1_l, t2_l, n_l)
        _setup_sig_cell(u_sig_cells[2], t1_r, t2_r, n_r)
        # Middle cell spacer
        for p in u_sig_cells[1].paragraphs[1:]:
            p._p.getparent().remove(p._p)
        p_mid = u_sig_cells[1].paragraphs[0]
        p_mid.text = ""
        p_mid.paragraph_format.space_before = Pt(0)
        p_mid.paragraph_format.space_after = Pt(0)
    else:
        _setup_sig_cell(last_row.cells[0], t1_l, t2_l, n_l)
        _setup_sig_cell(last_row.cells[-1], t1_r, t2_r, n_r)

    # Enforce vAlign='top' across all cells in Table 0
    for tc in t0._tbl.xpath(".//w:tc"):
        tcPr = tc.get_or_add_tcPr()
        old_v = tcPr.find(f"{{{W_NS}}}vAlign")
        if old_v is not None:
            tcPr.remove(old_v)
        tcPr.append(parse_xml(f'<w:vAlign xmlns:w="{W_NS}" w:val="top"/>'))

    # Enforce balanced paragraph spacing on Table 0 (Row 10 equipment line-height single, others 2pt)
    for r_idx in range(len(t0.rows) - 1):
        row = t0.rows[r_idx]
        for c in row.cells:
            for p in c.paragraphs:
                pPr = p._p.get_or_add_pPr()
                old_sp = pPr.find(f"{{{W_NS}}}spacing")
                if old_sp is not None:
                    pPr.remove(old_sp)
                if r_idx == 10:  # Equipment multi-line row
                    pPr.append(parse_xml(f'<w:spacing xmlns:w="{W_NS}" w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'))
                else:
                    pPr.append(parse_xml(f'<w:spacing xmlns:w="{W_NS}" w:before="40" w:after="40" w:line="260" w:lineRule="auto"/>'))


_UNIT_FIXES = (
    (re.compile(r"(?<=[\s\d,(])oC\b"), "°C"),     # '20 oC', 'Temperature, oC' -> °C
    (re.compile(r"\bm3\b"), "m³"),                 # 'mg/m3', 'hạt/m3', 'm3/h' -> m³
)


def fix_unit_symbols(doc):
    """Sửa ký hiệu đơn vị gõ thường trong thân văn bản: 'oC' -> '°C', 'm3' -> 'm³' (sửa theo từng run,
    không đụng run đã định dạng chỉ số trên)."""
    for t in doc.element.body.iter(f"{{{W_NS}}}t"):
        txt = t.text or ""
        new = txt
        for rx, rep in _UNIT_FIXES:
            new = rx.sub(rep, new)
        if new != txt:
            t.text = new
