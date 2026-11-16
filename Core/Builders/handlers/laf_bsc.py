# -*- coding: utf-8 -*-
"""
LAF and Biological Safety Cabinet (BSC) Document Handler for QUATEST 3.
Implements the All-In-One (AIO) inspection report layout matching TBKQ_AIO / TBKQ_AIO_EN
and standard EN 12469:2000 / ISO 14644-1:2015.

Outputs:
- TBKQ Vietnamese (.docx)
- TBKQ English (.docx)
- Tem kiểm tra / Validation Seal (.docx)
"""

import os
import shutil
import re
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from typing import Optional

from ...Models.dossier import Dossier
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
)
from ..tag_context import render_cover, build_tem_tu_doc
from ...Styling.tables import add_tbl_header, add_cant_split, set_cell_margins, A_NS


class LAFBSCHandler:
    """Builder handler for LAF / BSC inspection reports adhering to TBKQ_AIO format."""

    def __init__(self, dossier: Dossier, template_path_vn: str = None, template_path_en: str = None):
        self.dossier = dossier
        templates_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            "Templates"
        )
        self.templates_dir = templates_dir
        vn_cand = os.path.join(templates_dir, "TBKQ - viet - phong sach.docx")
        if not os.path.isfile(vn_cand):
            vn_cand = os.path.join(templates_dir, "TBKQ - vn - phong sach.docx")
        self.template_vn = template_path_vn or vn_cand
        self.template_en = template_path_en or os.path.join(templates_dir, "TBKQ - en - phong sach.docx")

    def build_tbkq(self, is_en: bool = False, **kwargs) -> str:
        """Builds complete TBKQ document (VN or EN) matching QUATEST 3 AIO standard form."""
        d = self.dossier
        template_file = self.template_en if is_en else self.template_vn
        doc = docx.Document(template_file)

        # 1. Update Headers & Table 0 (Trang bìa - Section 0)
        render_cover(doc, d, is_en)  # trang bìa từ template TAG (Core/Templates/TAGS.md)

        # 2. Clear out existing body elements of Section 1
        body = doc.element.body
        sect_break_p = None
        for el in body:
            if el.tag.endswith('p') and el.findall(f".//{{{W_NS}}}sectPr"):
                sect_break_p = el
                break

        if sect_break_p is not None:
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
        c_info = d.results_data.get("cabinet", {})
        client_name = d.client.name_en if is_en else d.client.name_vn

        # --- 10.1 Thông tin thiết bị ---
        h101 = "10.1 Device information" if is_en else "10.1 Thông tin thiết bị"
        add_p(h101, bold=True, font_size=13, keep_next=True, space_before=4)

        brand_val = c_info.get('brand', 'ABT')
        model_val = c_info.get('model', 'ISA2 - 1000')
        serial_val = c_info.get('serial', '26080302-01')
        dev_id_val = c_info.get('device_id', 'N/A')
        cab_name = c_info.get('cabinet_name', 'Tủ an toàn sinh học cấp 2')

        # 2-column borderless table for device info so bullets line up vertically and balanced
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

        v_vn = f"- Theo xác nhận của đại diện {client_name}, thiết bị \"{cab_name}\" nêu trên đã được vệ sinh và vận hành bình thường tại thời điểm thực hiện việc giám định."
        v_en = f"- As validated by the representative of {client_name}, the Biosafety Cabinet class II has been cleaned and operated under normal conditions during the inspection."
        add_p(v_en if is_en else v_vn, font_size=13, keep_next=True, space_after=6)

        # --- 10.2 Nội dung thực hiện ---
        h102 = "10.2 Performance content" if is_en else "10.2 Nội dung thực hiện"
        add_p(h102, bold=True, font_size=13, keep_next=True, space_before=4)

        tasks_vn = [
            "1. Đo tốc độ dòng khí qua cửa làm việc;",
            "2. Đo tốc độ dòng khí cấp qua lọc HEPA;",
            "3. Đo cường độ sáng bên trong tủ;",
            "4. Đo độ ồn của tủ khi vận hành;",
            "5. Kiểm tra độ rò rỉ của các lọc HEPA;",
            "6. Kiểm tra hình thái dòng khí của tủ;",
        ]
        tasks_en = [
            "1. Inflow air velocity test;",
            "2. Downflow air velocity test;",
            "3. Illuminance test in the cabinet workspace;",
            "4. Noise test of the cabinet in operation;",
            "5. HEPA filter leakage test;",
            "6. Airflow patterns test;",
        ]
        if d.results_data.get("uvc"):
            tasks_vn.append("7. Đo cường độ tia UV dải C trong khoang làm việc.")
            tasks_en.append("7. UVC intensity measurement in the workspace.")
        if d.results_data.get("particle"):
            tasks_vn.append("8. Đo nồng độ hạt bụi trong tủ.")
            tasks_en.append("8. Particle concentration measurement in the workspace.")

        for t_txt in (tasks_en if is_en else tasks_vn):
            add_p(t_txt, font_size=13, keep_next=True)

        # --- 10.3 Kết quả giám định ---
        h103 = "10.3. Inspection result" if is_en else "10.3. Kết quả giám định"
        add_p(h103, bold=True, font_size=13, keep_next=True, space_before=6)

        # Dòng bảng và đánh giá dùng chung BSCHandler.result_rows (ngưỡng EN 12469: inflow ≥ 0,40 m/s, downflow
        # 0,25-0,50 m/s, lux ≥ 750, ồn ≤ 65 dBA, rò rỉ HEPA ≤ 0,01 %); trước 30/10/2026 handler này ghi cứng "Đạt".
        from .bsc import BSCHandler
        rows_full = BSCHandler(d).result_rows(is_en)
        rows_def = [r[:7] for r in rows_full]

        tbl = doc.add_table(rows=len(rows_def), cols=5)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        widths = [700, 3600, 2000, 2200, 1100]

        for r_idx, r_data in enumerate(rows_def):
            stt, param, res_v, lim_v, assess_v, is_parent, is_header = r_data
            row = tbl.rows[r_idx]
            add_cant_split(row)
            if is_header:
                add_tbl_header(row)

            for c_idx in range(5):
                c = row.cells[c_idx]
                tcPr = c._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:tcW xmlns:w="{W_NS}" w:w="{widths[c_idx]}" w:type="dxa"/>'))

            if is_header:
                set_cell_content(row.cells[0], stt, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[1], param, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[2], res_v, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[3], lim_v, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[4], assess_v, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
            elif is_parent:
                parent_bold = False   # dòng cha không in đậm (như BSCHandler)
                set_cell_content(row.cells[0], stt, bold=parent_bold, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                merged_param = row.cells[1].merge(row.cells[4])
                set_cell_content(merged_param, param, bold=parent_bold, font_size=13, align=WD_ALIGN_PARAGRAPH.LEFT, valign="top", space_before=3, space_after=3)
                if rows_full[r_idx][7]:   # keep_next_parent (hình thái dòng khí, nồng độ hạt bụi)
                    for c in row.cells:
                        for p in c.paragraphs:
                            p.paragraph_format.keep_with_next = True
                else:
                    for c in row.cells:
                        for p in c.paragraphs:
                            p.paragraph_format.keep_with_next = False
            else:
                set_cell_content(row.cells[0], stt, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[1], param, font_size=13, align=WD_ALIGN_PARAGRAPH.LEFT, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[2], res_v, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[3], lim_v, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[4], assess_v, bold=(assess_v != "-"), italic=(assess_v != "-"), font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                if "Inside cabinet workspace" in param or "Bên trong không gian tủ" in param:
                    for c in row.cells:
                        for p in c.paragraphs:
                            p.paragraph_format.keep_with_next = True

        has_part = bool(d.results_data.get("particle"))
        bsc_row_heights = [r[8] for r in rows_full]
        format_table_structure(tbl, is_note=False, widths=widths, row_height=bsc_row_heights)
        merge_stt_column_child_rows(tbl, stt_col_idx=0, start_row=1, font_size=13, set_uniform_spacing=True)

        # Note Table (Font 10 per AGENTS.md Rule 3, borderless)
        t_note = doc.add_table(rows=1, cols=2)
        set_cell_content(t_note.rows[0].cells[0], "Note:" if is_en else "Ghi chú:", bold=True, font_size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)
        note_msg = (
            "The performance field is accredited to ISO 17020 - VIAS 004.\n⁽¹⁾ The threshold specified in EN 12469:2000."
            + ("\n⁽²⁾ The threshold specified for ISO Class 5 in ISO 14644-1:2015." if has_part else "")
            if is_en else
            "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.\n⁽¹⁾ Mức quy định tại EN 12469:2000."
            + ("\n⁽²⁾ Mức quy định đối với cấp độ sạch ISO 5 tại ISO 14644-1:2015." if has_part else "")
        )
        note_msg += BSCHandler(d).extra_notes(is_en)
        set_cell_content(t_note.rows[0].cells[1], note_msg, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_table_structure(t_note, is_note=True, widths=[1000, 8600])

        # --- 10.4 Hình ảnh quá trình thực hiện ---
        h104 = "10.4. Photos of performance process" if is_en else "10.4. Hình ảnh quá trình thực hiện"
        add_p(h104, bold=True, font_size=13, keep_next=True, space_before=6)

        # Helper to optimize and fix EXIF orientation for Word
        def _get_processed_photo(p_path: str, max_w: int = 1600) -> str:
            try:
                from PIL import Image, ImageOps
                im = Image.open(p_path)
                im = ImageOps.exif_transpose(im)
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
                print(f"[Warning] Failed to optimize photo {p_path}: {e}")
                return p_path

        # Standard default captions map for BSC photos
        default_captions_vn = [
            "Hình 1: Vị trí đo và bố trí thiết bị kiểm tra",
            "Hình 2: Bố trí thiết bị đo trong khoang làm việc",
            "Hình 3: Nhãn thông số kỹ thuật của tủ",
            "Hình 4: Khảo sát hình thái dòng khí bên trong tủ",
            "Hình 5: Khảo sát hình thái dòng khí tại cửa làm việc",
            "Hình 6: Kiểm tra độ rò rỉ màng lọc HEPA",
        ]
        default_captions_en = [
            "Figure 1: Measurement location and equipment setup",
            "Figure 2: Measurement instruments inside workspace",
            "Figure 3: Biosafety cabinet nameplate",
            "Figure 4: Airflow pattern inside workspace",
            "Figure 5: Airflow pattern at work opening",
            "Figure 6: HEPA filter leak test",
        ]

        # Gather photos
        photo_items = []
        if d.photos and len(d.photos) > 0:
            for i, p in enumerate(d.photos):
                p_path = p.get("path") if isinstance(p, dict) else p
                p_cap = (p.get("caption_en" if is_en else "caption_vn") or p.get("caption") or (default_captions_en[i] if is_en else default_captions_vn[i])) if isinstance(p, dict) else (default_captions_en[i] if is_en else default_captions_vn[i])
                photo_items.append((p_path, p_cap))
        else:
            photo_dir = None
            for cand in ["photo", "Photo Thuc Hien", "Photo", "photos"]:
                p_cand = os.path.join(d.input_dir, cand)
                if os.path.exists(p_cand) and os.path.isdir(p_cand):
                    photo_dir = p_cand
                    break

            if photo_dir:
                img_exts = ('.jpg', '.jpeg', '.png')
                cands = [os.path.join(photo_dir, f) for f in sorted(os.listdir(photo_dir)) if f.lower().endswith(img_exts) and not f.startswith("thumb_") and not f.startswith("opt_")]

                for i, c_path in enumerate(cands[:6]):
                    base_f = os.path.basename(c_path)
                    cap_vn = default_captions_vn[i] if i < len(default_captions_vn) else f"Hình {i+1}: Quá trình thực hiện kiểm tra"
                    cap_en = default_captions_en[i] if i < len(default_captions_en) else f"Figure {i+1}: Performance process"
                    if "workspace" in base_f:
                        cap_vn = "Hình 2: Bố trí thiết bị đo trong khoang làm việc"
                        cap_en = "Figure 2: Measurement instruments inside workspace"
                    elif "093155" in base_f:
                        cap_vn = "Hình 3: Nhãn thông số kỹ thuật của tủ"
                        cap_en = "Figure 3: Biosafety cabinet nameplate"
                    elif "105528" in base_f:
                        cap_vn = "Hình 4: Khảo sát hình thái dòng khí bên trong tủ"
                        cap_en = "Figure 4: Airflow pattern inside workspace"
                    elif "110228" in base_f:
                        cap_vn = "Hình 5: Khảo sát hình thái dòng khí tại cửa làm việc"
                        cap_en = "Figure 5: Airflow pattern at work opening"
                    elif "111307" in base_f:
                        cap_vn = "Hình 6: Kiểm tra độ rò rỉ màng lọc HEPA"
                        cap_en = "Figure 6: HEPA filter leak test"
                    photo_items.append((c_path, cap_en if is_en else cap_vn))

        # Re-number captions sequentially from 1 to len(photo_items)
        numbered_photos = []
        for i, (p_path, cap) in enumerate(photo_items):
            clean_cap = re.sub(r'^(Hình \d+: |Figure \d+: )', '', cap or '')
            prefix = f"Figure {i+1}: " if is_en else f"Hình {i+1}: "
            numbered_photos.append((p_path, prefix + clean_cap))

        if numbered_photos:
            import math
            if len(numbered_photos) >= 5:
                num_cols = 3
                num_rows = 2
                img_w = PHOTO_WIDTH_3COL  # 4.8 cm (4 - 5 cm, strictly no crop)
                cap_sz = 11
            else:
                num_cols = 2
                num_rows = math.ceil(len(numbered_photos) / num_cols)
                img_w = PHOTO_WIDTH_2COL  # 7.5 cm (7 - 8 cm, strictly no crop)
                cap_sz = 13  # Font size 13 italic for captions per AGENTS.md Section 3

            t_photo = doc.add_table(rows=num_rows, cols=num_cols)
            t_photo.alignment = WD_TABLE_ALIGNMENT.CENTER

            for idx, (p_path, p_cap) in enumerate(numbered_photos):
                r_idx = idx // num_cols
                c_idx = idx % num_cols
                cell = t_photo.rows[r_idx].cells[c_idx]
                p_img = cell.paragraphs[0]
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(0)
                p_img.paragraph_format.space_after = Pt(2)
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

        # Standard Ending Line (4.5pt thinThick centered)
        add_document_ending_line(doc)
        last_p = doc.paragraphs[-1]
        last_p.paragraph_format.space_before = Pt(4)
        last_p.paragraph_format.space_after = Pt(0)

        # Tổng số trang Phụ lục ở footer: template TAG dùng trường SECTIONPAGES, Word tự tính (không ghi cứng).

        strip_all_highlights(doc)
        strip_image_crops(doc)

        # Dispatch files per AGENTS.md Rule 1
        d.ensure_output_directories()
        lang_code = "en" if is_en else "viet"
        sub_str = f" - {d.sub_no}" if d.sub_no else ""
        file_name = f"TBKQ - {lang_code}{sub_str} - BSC - {d.code}.docx"

        target_tbkq = os.path.join(d.tbkq_dir, file_name)
        

        from ..tag_context import style_note_tables
        style_note_tables(doc)          # bảng Ghi chú 11pt in nghiêng (09/11/2026)
        doc.save(target_tbkq)

        # Word COM normalization: ensure Table 1 paragraph spacing (3pt before & 3pt after) is recognized uniformly by Word ribbon
        try:
            import win32com.client
            from ...Exporters.pdf_exporter import kill_orphaned_word_processes
            kill_orphaned_word_processes()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                wdoc = word.Documents.Open(os.path.abspath(target_tbkq))
                if wdoc.Tables.Count >= 2:
                    t1_w = wdoc.Tables(2)
                    t1_w.Range.ParagraphFormat.SpaceBefore = 3
                    t1_w.Range.ParagraphFormat.SpaceAfter = 3
                wdoc.Save()
                wdoc.Close(False)
            finally:
                word.Quit()
        except Exception as e:
            print(f"[LAFBSCHandler] COM normalization skipped/failed: {e}")

        if d.output_dir:
            try:
                shutil.copy2(target_tbkq, os.path.join(d.output_dir, file_name))
            except Exception as _err:
                print(f"[Warning] laf_bsc.py:503 {type(_err).__name__}: {_err}")

        print(f"[LAFBSCHandler] Generated TBKQ {'EN' if is_en else 'VN'} -> {target_tbkq}")
        return target_tbkq

    def build_tem_tu(self, template_path: Optional[str] = None) -> str:
        """Tem kiểm tra dán tủ LAF/BSC từ template TAG Core/Templates/tem tu.docx."""
        d = self.dossier
        tpl = template_path or os.path.join(getattr(self, "templates_dir", os.path.join("Core", "Templates")), "tem tu.docx")
        return build_tem_tu_doc(d, tpl, f"{d.code.replace('Q326D', '')} - tem tu.docx", so_tem=d.code,
                                ten_thiet_bi=(d.results_data.get("cabinet", {}) or {}).get("seal_item_name", "“Tủ cấy an toàn sinh học cấp 2”"),
                                label="LAFBSCHandler")
