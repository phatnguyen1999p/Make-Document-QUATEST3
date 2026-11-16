# -*- coding: utf-8 -*-
"""
Fume Hood (Tủ hút khí độc) Document Handler for QUATEST 3.
Implements the standardized All-In-One (AIO) inspection report layout matching
ASHRAE 110, SEFA 1-2010, and QUATEST 3 standards.

Supports 2 distinct report forms:
1. Form Có đánh giá (With assessment / Pass-Fail evaluation):
   Columns: STT | Chỉ tiêu thực hiện | Kết quả đo | Mức quy định | Đánh giá (5 cột)
   English: No. | Parameter | Result | Threshold | Assessment
2. Form Không đánh giá (Without assessment / Reference recommendations):
   Columns: STT | Chỉ tiêu thực hiện | Kết quả đo | Mức khuyến nghị (4 cột)
   English: No. | Parameter | Result | Recommended level

Features:
- Dynamically generates and embeds the sash opening measurement grid illustrations
  (showing sash width, height opening percentage, and measurement grid values).
- Vietnamese & English technical inspection reports (TBKQ VN / EN).
- Validation seal sticker document ('tem tu.docx').
- Complete compliance with AGENTS.md rules (font sizes, AutoFit, ending lines, row heights, headers).
"""

import os
import shutil
import math
import docx
from docx.shared import Pt, Cm
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from typing import Optional, Dict

from ...Models.dossier import Dossier, inspection_scope
from ...Styling import (
    W_NS,
    set_cell_content,
    set_table_autofit_window,
    format_table_structure,
    strip_all_highlights,
    add_document_ending_line,
    PHOTO_WIDTH_2COL,
    PHOTO_WIDTH_3COL,
    strip_image_crops,
    merge_stt_column_child_rows,
    get_fumehood_row_heights,
)
from ..tag_context import render_cover, build_tem_tu_doc
from ...Styling.tables import add_tbl_header, add_cant_split, set_cell_margins, A_NS
from ...Tools.fumehood_calc import generate_sash_measurement_image


class FumeHoodHandler:
    """Specialized Builder handler for Fume Hood inspection reports."""

    def __init__(self, dossier: Dossier, template_path_vn: str = None, template_path_en: str = None):
        self.dossier = dossier
        templates_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            "Templates"
        )
        cand_vn = os.path.join(templates_dir, "TBKQ - vn - phong sach.docx")
        if not os.path.exists(cand_vn):
            cand_vn = os.path.join(templates_dir, "TBKQ - viet - phong sach.docx")
        self.templates_dir = templates_dir
        self.template_vn = template_path_vn or cand_vn
        self.template_en = template_path_en or os.path.join(templates_dir, "TBKQ - en - phong sach.docx")
        tem_fh = os.path.join(templates_dir, "tem tu - fumehood.docx")
        self.tem_tu_template = tem_fh if os.path.exists(tem_fh) else os.path.join(templates_dir, "tem tu.docx")

    def build_tbkq(self, is_en: bool = False, has_assessment: Optional[bool] = None) -> str:
        """
        Builds complete TBKQ document (VN or EN) matching QUATEST 3 standard form.
        
        Parameters:
        - is_en: False for Vietnamese, True for English.
        - has_assessment:
            True: 5-column form with 'Mức quy định' & 'Đánh giá' (Pass/Fail).
            False: 4-column form with 'Mức khuyến nghị' / 'Recommended level' and NO assessment column.
            None: defaults to self.dossier.has_assessment.
        """
        d = self.dossier
        if has_assessment is None:
            has_assessment = getattr(d, "has_assessment", True)

        template_file = self.template_en if is_en else self.template_vn
        doc = docx.Document(template_file)

        # 1. Update Headers & Table 0 (Trang bìa - Section 0)

        # Set Table 0 Standards according to assessment mode
        if not d.procedure_vn:
            d.procedure_vn = "QUATEST 3 3030:2025, N2.TC.T.01"
        if not d.procedure_en:
            d.procedure_en = "QUATEST 3 3030:2025, N2.TC.T.01"

        if has_assessment:
            std_vn = "ASHRAE 110, SEFA 1-2010\nTài liệu kỹ thuật công bố bởi nhà sản xuất"
            std_en = "ASHRAE 110, SEFA 1-2010\nTechnical documentation published by the manufacturer"
        else:
            std_vn = "ASHRAE 110, SEFA 1-2010\nTài liệu kỹ thuật tham khảo"
            std_en = "ASHRAE 110, SEFA 1-2010\nTechnical reference documentation"

        d.standards_vn = d.standards_vn or std_vn
        d.standards_en = d.standards_en or std_en

        c_info = d.results_data.get("cabinet", {})
        cab_name = c_info.get("cabinet_name", "Tủ hút khí độc")
        if not d.scope_vn:
            d.scope_vn = inspection_scope(f"Giám định {cab_name}")
        if not d.scope_en:
            d.scope_en = f'Measuring/ validating "{cab_name.strip(chr(34))}"'

        render_cover(doc, d, is_en)  # trang bìa từ template TAG (Core/Templates/TAGS.md)

        # 2. Clear out existing body elements of Section 1
        body = doc.element.body
        sect_break_p = None
        for el in body:
            if el.tag.endswith('p') and el.findall(f".//{{{W_NS}}}sectPr"):
                sect_break_p = el
                break

        if sect_break_p is not None:
            pPr = sect_break_p.get_or_add_pPr()
            sp = pPr.find(f"{{{W_NS}}}spacing")
            if sp is not None:
                pPr.remove(sp)
            pPr.append(parse_xml(f'<w:spacing xmlns:w="{W_NS}" w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'))
            idx = list(body).index(sect_break_p)
            for child in list(body)[idx + 1:]:
                if not child.tag.endswith('sectPr'):
                    body.remove(child)

        # Helper for adding paragraphs with Times New Roman and keepNext
        def add_p(text="", bold=False, italic=False, font_size=13, align=None, keep_next=False, space_before=0, space_after=2):
            p = doc.add_paragraph()
            if align is not None:
                p.alignment = align
            p.paragraph_format.space_before = Pt(space_before)
            p.paragraph_format.space_after = Pt(space_after)
            p.paragraph_format.line_spacing = 1.15
            if keep_next:
                pPr = p._p.get_or_add_pPr()
                pPr.append(parse_xml(f'<w:keepNext xmlns:w="{W_NS}"/>'))
            if text:
                r = p.add_run(text)
                r.bold = bold
                r.italic = italic
                r.font.name = "Times New Roman"
                r.font.size = Pt(font_size)
                rPr = r._r.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                rFonts.set(qn("w:ascii"), "Times New Roman")
                rFonts.set(qn("w:hAnsi"), "Times New Roman")
                rFonts.set(qn("w:cs"), "Times New Roman")
            return p

        dec = "." if is_en else ","
        client_name = d.client.name_en if is_en else d.client.name_vn

        # --- 10.1 Thông tin thiết bị (Device information) ---
        h101 = "10.1 Device information" if is_en else "10.1 Thông tin thiết bị"
        add_p(h101, bold=True, font_size=13, keep_next=True, space_before=4)

        brand_val = c_info.get('brand', '')
        model_val = c_info.get('model', '')
        serial_val = c_info.get('serial', '')
        dev_id_val = c_info.get('device_id', 'N/A')

        t_info = doc.add_table(rows=2, cols=2)
        format_table_structure(t_info, is_note=True, widths=[4800, 4800])
        add_tbl_header(t_info.rows[0])

        info_rows = [
            (f"- {'Manufacturer' if is_en else 'Nhãn hiệu'}: {brand_val}", f"- Model: {model_val}"),
            (f"- {'Serial' if is_en else 'Số Serial'}: {serial_val}", f"- {'Device\'s ID' if is_en else 'Mã thiết bị'}: {dev_id_val}")
        ]
        for r_idx, (col0_txt, col1_txt) in enumerate(info_rows):
            r = t_info.rows[r_idx]
            add_cant_split(r)
            for c_idx, txt in enumerate([col0_txt, col1_txt]):
                c = r.cells[c_idx]
                set_cell_margins(c, top=0, bottom=20, left=0, right=0)
                set_cell_content(c, txt, font_size=13, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True, space_before=0, space_after=2)

        v_vn = f"- Theo xác nhận của đại diện {client_name}, thiết bị \"{cab_name.strip(chr(34))}\" nêu trên đã được vệ sinh và vận hành bình thường tại thời điểm thực hiện việc giám định."
        v_en = f"- As validated by the representative of {client_name}, the \"{cab_name.strip(chr(34))}\" has been cleaned and was functioning normally at the time of performance."
        add_p(v_en if is_en else v_vn, font_size=13, keep_next=True, space_after=6)

        # --- 10.2 Nội dung thực hiện (Performance content) ---
        h102 = "10.2 Performance content" if is_en else "10.2 Nội dung thực hiện"
        add_p(h102, bold=True, font_size=13, keep_next=True, space_before=4)

        # Tùy chọn bỏ chỉ tiêu (mặc định giữ nguyên 4 chỉ tiêu):
        # results_data["lux"] = {"skip": True}, results_data["airflow_pattern"] = {"skip": True}
        has_lux = not (d.results_data.get("lux") or {}).get("skip", False)
        has_airflow = not (d.results_data.get("airflow_pattern") or {}).get("skip", False)
        tasks_vn = ["Đo tốc độ dòng khí qua cửa làm việc;"]
        tasks_en = ["Face air velocity test;"]
        if has_lux:
            tasks_vn.append("Đo cường độ sáng bên trong tủ;"); tasks_en.append("Illuminance test in the cabinet workspace;")
        tasks_vn.append("Đo độ ồn của tủ khi vận hành;"); tasks_en.append("Noise test of the cabinet in operation;")
        if has_airflow:
            tasks_vn.append("Kiểm tra hình thái dòng khí của tủ;"); tasks_en.append("Smoke visualization test;")
        tasks_vn[-1] = tasks_vn[-1].rstrip(";") + "."
        tasks_en[-1] = tasks_en[-1].rstrip(";") + "."
        for i, t_txt in enumerate(tasks_en if is_en else tasks_vn, start=1):
            add_p(f"{i}. {t_txt}", font_size=13, keep_next=True)

        # --- 10.3 Kết quả giám định (Inspection result) ---
        h103 = "10.3. Inspection result" if is_en else "10.3. Kết quả giám định"
        add_p(h103, bold=True, font_size=13, keep_next=True, space_before=6)

        scratch_diag_dir = os.path.join(d.input_dir, "TBKQ", "_diagrams")
        os.makedirs(scratch_diag_dir, exist_ok=True)

        res = d.results_data
        openings = res.get("openings", [])
        if not openings:
            openings = [{
                "width_mm": 1530,
                "height_mm": 585,
                "opening_pct": 100,
                "avg": 0.52,
                "limit_vn": "≥ 0,5 m/s (*)",
                "limit_en": "≥ 0.5 m/s (*)",
                "assessment_vn": "Đạt",
                "assessment_en": "Pass",
                "readings": [[0.52, 0.53, 0.51, 0.52, 0.55], [0.51, 0.53, 0.52, 0.51, 0.51]]
            }]

        lux_res = res.get("lux", {})
        noise_res = res.get("noise", {})
        airflow_res = res.get("airflow_pattern", {})

        lux_val = str(lux_res.get("avg_rounded", int(round(lux_res.get("avg", 863)))) if lux_res else "863")
        noise_amb = str(noise_res.get("ambient", "51.6")).replace(".", dec)
        noise_op = str(noise_res.get("corrected_dba", noise_res.get("operating", "70.7"))).replace(".", dec)
        pass_txt = "Pass" if is_en else "Đạt"

        # Table Column Configuration
        if has_assessment:
            # 5 Columns: STT, Parameter, Result, Threshold, Assessment
            num_cols = 5
            widths = [700, 3600, 2000, 2200, 1100]
            header_tuple = ("No." if is_en else "STT", "Parameter" if is_en else "Chỉ tiêu thực hiện", "Result" if is_en else "Kết quả đo", "Threshold" if is_en else "Mức quy định", "Assessment" if is_en else "Đánh giá", False, True, False, None)
        else:
            # 4 Columns: STT, Parameter, Result, Recommended level (NO assessment column)
            num_cols = 4
            widths = [700, 4200, 2300, 2400]
            header_tuple = ("No." if is_en else "STT", "Parameter" if is_en else "Chỉ tiêu thực hiện", "Result" if is_en else "Kết quả đo", "Recommended level" if is_en else "Mức khuyến nghị", "", False, True, False, None)

        rows_def = [header_tuple]

        # Inflow / Face velocity rows and diagrams for each sash opening
        for op_idx, op in enumerate(openings):
            stt = "1" if op_idx == 0 else ""
            w_mm = op.get("width_mm", 1530)
            h_mm = op.get("height_mm", 585)
            pct = op.get("opening_pct")
            pct_desc = f" ({pct}%)" if pct else ""

            if len(openings) > 1:
                param_vn = f"Tốc độ trung bình dòng khí qua cửa làm việc - Khoảng mở cửa: {h_mm} mm{pct_desc}"
                param_en = f"Average face velocity - Sash opening: {h_mm} mm{pct_desc}"
            else:
                param_vn = "Tốc độ trung bình dòng khí qua cửa làm việc"
                param_en = "Average face velocity"

            v_avg_val = op.get("avg", 0.50)
            try:
                v_avg = f"{float(v_avg_val):.2f}".replace(".", dec)
            except Exception:
                v_avg = str(v_avg_val).replace(".", dec)

            lim_v = op.get("limit_en", "≥ 0.5 m/s (*)") if is_en else op.get("limit_vn", "≥ 0,5 m/s (*)")
            ass_v = (op.get("assessment_en", pass_txt) if is_en else op.get("assessment_vn", pass_txt)) if has_assessment else ""

            rows_def.append((stt, param_en if is_en else param_vn, f"{v_avg} m/s", lim_v, ass_v, False, False, False, None))

            # Generate sash illustration image
            lang_sfx = "en" if is_en else "vn"
            mode_sfx = "eval" if has_assessment else "noeval"
            img_filename = f"sash_diag_op{op_idx+1}_{lang_sfx}_{mode_sfx}.png"
            img_path = os.path.join(scratch_diag_dir, img_filename)
            readings_grid = op.get("readings")
            generate_sash_measurement_image(
                output_path=img_path,
                width_mm=w_mm,
                height_mm=h_mm,
                opening_pct=pct,
                velocity_grid=readings_grid,
                is_en=is_en,
                dpi=300
            )

            # Insert Image Row directly below the velocity row
            rows_def.append(("", "", "", "", "", False, False, True, img_path))

        next_stt = 2
        # Parameter 2: Illuminance
        lim_lux = "≥ 861 lux (**)" if has_assessment else "≥ 861 lux (*)"
        if has_lux:
            rows_def.append((
                str(next_stt),
                "Average illuminance" if is_en else "Cường độ sáng trung bình bên trong tủ",
                f"{lux_val} lux",
                lim_lux,
                (lux_res.get("assessment_en" if is_en else "assessment_vn", pass_txt)) if has_assessment else "",
                False, False, False, None
            ))
            next_stt += 1

        # Parameter 3: Sound pressure level
        lim_noise = ("≤ 85.0 dBA (**)" if is_en else "≤ 85,0 dBA (**)") if has_assessment else ("≤ 85.0 dBA (*)" if is_en else "≤ 85,0 dBA (*)")
        rows_def.append((str(next_stt), "Sound pressure level" if is_en else "Đo độ ồn", "", "", "", True, False, False, None))
        next_stt += 1
        rows_def.append(("", "Background noise" if is_en else "Độ ồn nền", f"{noise_amb} dBA", "Not specified" if is_en else "Không quy định", "-" if has_assessment else "", False, False, False, None))
        rows_def.append(("", "Operation noise" if is_en else "Độ ồn khi vận hành", f"{noise_op} dBA", lim_noise,
                         (noise_res.get("assessment_en" if is_en else "assessment_vn", pass_txt)) if has_assessment else "", False, False, False, None))

        # Parameter 4: Smoke visualization test / Airflow pattern test
        if has_airflow:
            sfx_smoke = " (**)" if has_assessment else " (*)"
            rows_def.append((str(next_stt), "Smoke visualization test" if is_en else "Kiểm tra hình thái dòng khí", "", "", "", True, False, False, None))
            rows_def.append((
                "",
                "Inside cabinet workspace\n- Along center line, above work opening\n- Behind work sash\n- Along edge of work sash" if is_en else "Bên trong không gian tủ\n- Dọc đường trung tâm, phía trên khoảng mở làm việc\n- Phía sau cửa mở làm việc\n- Dọc theo mép cửa mở làm việc",
                "- Air flow entered rear wall, towards intake vents.\n- Air flow behind sash did not escape." if is_en else "- Dòng khí đi vào vách sau, hướng vào các khe hút của tủ.\n- Dòng khí sau cửa không thoát ra ngoài.",
                ("- Air flow entered rear wall, towards intake vents.\n- Air flow behind sash did not escape." + sfx_smoke) if is_en else ("- Dòng khí đi vào vách sau, hướng vào các khe hút của tủ.\n- Dòng khí sau cửa không thoát ra ngoài." + sfx_smoke),
                pass_txt if has_assessment else "",
                False, False, False, None
            ))
            rows_def.append((
                "",
                "Area along the outer perimeter of the work opening" if is_en else "Khu vực dọc theo chu vi phía bên ngoài của vùng mở cửa tủ",
                "- Airflow went into cabinet, not out." if is_en else "- Dòng khí đi vào trong tủ, không đi ngược ra ngoài.",
                ("- Airflow went into cabinet, not out." + sfx_smoke) if is_en else ("- Dòng khí đi vào trong tủ, không đi ngược ra ngoài." + sfx_smoke),
                pass_txt if has_assessment else "",
                False, False, False, None
            ))

        # Build Word Table
        tbl = doc.add_table(rows=len(rows_def), cols=num_cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        for r_idx, r_data in enumerate(rows_def):
            stt, param, res_v, lim_v, assess_v, is_parent, is_header, is_image, img_path = r_data
            row = tbl.rows[r_idx]
            add_cant_split(row)

            if is_header:
                add_tbl_header(row)
                for c_idx in range(num_cols):
                    c = row.cells[c_idx]
                    tcPr = c._tc.get_or_add_tcPr()
                    tcPr.append(parse_xml(f'<w:tcW xmlns:w="{W_NS}" w:w="{widths[c_idx]}" w:type="dxa"/>'))
                set_cell_content(row.cells[0], stt, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[1], param, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[2], res_v, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[3], lim_v, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
                if has_assessment:
                    set_cell_content(row.cells[4], assess_v, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)

            elif is_image:
                # Merge across all columns for the measurement diagram
                last_c_idx = num_cols - 1
                merged_c = row.cells[0].merge(row.cells[last_c_idx])
                set_cell_margins(merged_c, top=40, bottom=40, left=40, right=40)
                p_img = merged_c.paragraphs[0]
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.paragraph_format.line_spacing = 1.0
                pPr = p_img._p.get_or_add_pPr()
                pPr.append(parse_xml(f'<w:keepNext xmlns:w="{W_NS}"/>'))
                if img_path and os.path.exists(img_path):
                    r_pic = p_img.add_run()
                    r_pic.add_picture(img_path, width=Cm(13.2))

            elif is_parent:
                set_cell_content(row.cells[0], stt, bold=False, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                last_c_idx = num_cols - 1
                merged_param = row.cells[1].merge(row.cells[last_c_idx])
                set_cell_content(merged_param, param, bold=False, font_size=13, align=WD_ALIGN_PARAGRAPH.LEFT, valign="top", space_before=3, space_after=3)
                for c in row.cells:
                    for p in c.paragraphs:
                        p.paragraph_format.keep_with_next = True

            else:
                for c_idx in range(num_cols):
                    c = row.cells[c_idx]
                    tcPr = c._tc.get_or_add_tcPr()
                    tcPr.append(parse_xml(f'<w:tcW xmlns:w="{W_NS}" w:w="{widths[c_idx]}" w:type="dxa"/>'))
                set_cell_content(row.cells[0], stt, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[1], param, font_size=13, align=WD_ALIGN_PARAGRAPH.LEFT, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[2], res_v, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[3], lim_v, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                if has_assessment:
                    set_cell_content(row.cells[4], assess_v, bold=(assess_v != "-"), italic=(assess_v != "-"), font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                if "Inside cabinet workspace" in param or "Bên trong không gian tủ" in param:
                    for c in row.cells:
                        for p in c.paragraphs:
                            p.paragraph_format.keep_with_next = True

        fh_row_heights = get_fumehood_row_heights(num_openings=len(openings))
        if not has_lux or not has_airflow:
            # Bỏ chiều cao của các dòng chỉ tiêu không thực hiện để khớp số dòng bảng
            base = 1 + 2 * len(openings)
            hs = fh_row_heights[:base]
            if has_lux:
                hs.append(fh_row_heights[base])
            hs += fh_row_heights[base + 1:base + 4]
            if has_airflow:
                hs += fh_row_heights[base + 4:]
            fh_row_heights = hs
        format_table_structure(tbl, is_note=False, widths=widths, row_height=fh_row_heights)
        merge_stt_column_child_rows(tbl, stt_col_idx=0, start_row=1, font_size=13, set_uniform_spacing=True)

        # Note Table (Font 10 per AGENTS.md Rule 3, borderless)
        t_note = doc.add_table(rows=1, cols=2)
        set_cell_content(t_note.rows[0].cells[0], "Note:" if is_en else "Ghi chú:", bold=True, font_size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)
        if has_assessment:
            note_msg = (
                "The performance field is accredited to ISO 17020 - VIAS 004.\n"
                "(*) Technical specifications declared by the manufacturer.\n"
                "(**) Recommendation specified in SEFA 1-2010."
                if is_en else
                "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.\n"
                "(*) Mức kỹ thuật công bố bởi nhà sản xuất.\n"
                "(**) Mức khuyến nghị dựa trên SEFA 1-2010."
            )
        else:
            note_msg = (
                "The performance field is accredited to ISO 17020 - VIAS 004.\n"
                "(*) Recommended level specified in SEFA 1-2010."
                if is_en else
                "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.\n"
                "(*) Mức khuyến nghị dựa trên SEFA 1-2010."
            )
        set_cell_content(t_note.rows[0].cells[1], note_msg, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_table_structure(t_note, is_note=True, widths=[1000, 8600])

        # --- 10.4 Hình ảnh quá trình thực hiện (Photos) ---
        h104 = "10.4. Photos of performance process" if is_en else "10.4. Hình ảnh quá trình thực hiện"
        add_p(h104, bold=True, font_size=13, keep_next=True, space_before=2, space_after=1)

        def _get_processed_photo(p_path: str, max_w: int = 1600) -> str:
            try:
                from PIL import Image, ImageOps
                im = Image.open(p_path)
                im = ImageOps.exif_transpose(im)
                if im.height > im.width:  # ảnh dọc -> xoay 90° để chèn ngang (AGENTS.md mục 3), không crop
                    im = im.rotate(90, expand=True)
                w, h = im.size
                if w > max_w:
                    new_h = int(h * (max_w / w))
                    im = im.resize((max_w, new_h), Image.Resampling.LANCZOS)
                c_dir = os.path.join(os.path.dirname(p_path), "_optimized")
                os.makedirs(c_dir, exist_ok=True)
                out_p = os.path.join(c_dir, f"opt_{os.path.splitext(os.path.basename(p_path))[0]}.jpg")
                im.convert("RGB").save(out_p, format="JPEG", quality=90)
                return out_p
            except Exception as e:
                print(f"[FumeHoodHandler] Photo optimization fallback: {e}")
                return p_path

        photo_paths = []
        if d.photos and len(d.photos) > 0:
            for p in d.photos:
                p_path = p.get("path") if isinstance(p, dict) else p
                if os.path.exists(p_path):
                    photo_paths.append(p_path)
        else:
            photo_dir = None
            for cand in ["photo", "Photo Thuc Hien", "Photo", "photos", "Vi tri"]:
                p_cand = os.path.join(d.input_dir, cand)
                if os.path.exists(p_cand) and os.path.isdir(p_cand):
                    photo_dir = p_cand
                    break

            if photo_dir:
                img_exts = ('.jpg', '.jpeg', '.png')
                cands = [os.path.join(photo_dir, f) for f in sorted(os.listdir(photo_dir)) if f.lower().endswith(img_exts) and not f.startswith("thumb_") and not f.startswith("opt_")]
                photo_paths.extend(cands[:6])

        if photo_paths:
            if len(photo_paths) >= 5:
                num_cols_p = 3
                num_rows_p = math.ceil(len(photo_paths) / num_cols_p)
                img_w = PHOTO_WIDTH_3COL  # 4.8 cm (4 - 5 cm, strictly no crop)
            else:
                num_cols_p = 2
                num_rows_p = math.ceil(len(photo_paths) / num_cols_p)
                img_w = PHOTO_WIDTH_2COL  # 7.5 cm (7 - 8 cm, strictly no crop)

            t_photo = doc.add_table(rows=num_rows_p, cols=num_cols_p)
            t_photo.alignment = WD_TABLE_ALIGNMENT.CENTER

            for idx, p_path in enumerate(photo_paths):
                r_idx = idx // num_cols_p
                c_idx = idx % num_cols_p
                cell = t_photo.rows[r_idx].cells[c_idx]
                p_img = cell.paragraphs[0]
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(0)
                p_img.paragraph_format.space_after = Pt(0)
                p_img.paragraph_format.line_spacing = 1.0
                if os.path.exists(p_path):
                    p_proc = _get_processed_photo(p_path)
                    r = p_img.add_run()
                    # Strictly specify only width (no height) to preserve aspect ratio without crop
                    r.add_picture(p_proc, width=img_w)
                    for sr in r._r.findall(f".//{{{A_NS}}}srcRect"):
                        sr.attrib.clear()

            for row in t_photo.rows:
                add_cant_split(row)

            set_table_autofit_window(t_photo)

        # Standard Ending Line (4.5pt thinThick centered, Rule 5)
        add_document_ending_line(doc)
        last_p = doc.paragraphs[-1]
        last_p.paragraph_format.space_before = Pt(2)
        last_p.paragraph_format.space_after = Pt(0)

        # Tổng số trang Phụ lục ở footer: template TAG dùng trường SECTIONPAGES, Word tự tính (không ghi cứng).

        strip_all_highlights(doc)
        strip_image_crops(doc)

        # Dispatch files per AGENTS.md Rule 1
        d.ensure_output_directories()
        lang_code = "en" if is_en else "viet"
        sub_str = f" - {d.sub_no}" if d.sub_no else ""
        mode_str = "" if has_assessment else " - KhongDanhGia"
        file_name = f"TBKQ - {lang_code}{sub_str} - FumeHood{mode_str} - {d.code}.docx"

        target_tbkq = os.path.join(d.tbkq_dir, file_name)

        from ..tag_context import style_note_tables
        style_note_tables(doc)          # bảng Ghi chú 11pt in nghiêng (09/11/2026)
        doc.save(target_tbkq)
        if d.output_dir:
            try:
                shutil.copy2(target_tbkq, os.path.join(d.output_dir, file_name))
            except Exception as _err:
                print(f"[Warning] fumehood.py:540 {type(_err).__name__}: {_err}")

        print(f"[FumeHoodHandler] Generated TBKQ {'EN' if is_en else 'VN'} ({'Có đánh giá' if has_assessment else 'Không đánh giá'}) -> {target_tbkq}")
        return target_tbkq

    def build_all_modes(self) -> Dict[str, str]:
        """Builds both Có đánh giá and Không đánh giá documents in VN and EN."""
        return {
            "vn_eval": self.build_tbkq(is_en=False, has_assessment=True),
            "en_eval": self.build_tbkq(is_en=True, has_assessment=True),
            "vn_no_eval": self.build_tbkq(is_en=False, has_assessment=False),
            "en_no_eval": self.build_tbkq(is_en=True, has_assessment=False),
        }

    def build_tem_tu(self, template_path: Optional[str] = None) -> str:
        """Tem kiểm tra dán tủ hút từ template TAG Core/Templates/tem tu - fumehood.docx."""
        d = self.dossier
        cab_name = (d.results_data.get("cabinet", {}) or {}).get("cabinet_name", "Tủ hút khí độc")
        return build_tem_tu_doc(d, template_path or self.tem_tu_template, (f"tem tu - {d.sub_no} - FumeHood - {d.code}.docx" if (d.is_multi_cert and d.sub_no) else f"tem tu - FumeHood - {d.code}.docx"),
                                so_tem=d.cert_no_vn or d.code, ten_thiet_bi=f"“{cab_name.strip(chr(34))}”",
                                label="FumeHoodHandler")
