# -*- coding: utf-8 -*-
"""
Biological Safety Cabinet (BSC) Document Handler for QUATEST 3.
Implements the standardized All-In-One (AIO) inspection report layout matching
EN 12469:2000 and QUATEST 3 standards.

Outputs:
- TBKQ Vietnamese (.docx & .pdf)
- TBKQ English (.docx & .pdf)
- Validation Seal Sticker / Tem kiểm tra (.docx)
"""

import os
import shutil
import math
import docx
from docx.shared import Pt, Cm
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from typing import Optional, Any, Tuple

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
    get_bsc_row_heights,
)
from ..tag_context import render_cover, build_tem_tu_doc
from ...Styling.tables import add_tbl_header, add_cant_split, set_cell_margins, A_NS


class BSCHandler:
    """Specialized Builder handler for Biosafety Cabinet Class II inspection reports."""

    def __init__(self, dossier: Dossier, template_path_vn: str = None, template_path_en: str = None):
        self.dossier = dossier
        templates_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            "Templates"
        )
        cand_vn = os.path.join(templates_dir, "TBKQ - vn - phong sach.docx")
        if not os.path.exists(cand_vn):
            cand_vn = os.path.join(templates_dir, "TBKQ - viet - phong sach.docx")
        self.template_vn = template_path_vn or cand_vn
        self.template_en = template_path_en or os.path.join(templates_dir, "TBKQ - en - phong sach.docx")
        self.tem_tu_template = os.path.join(templates_dir, "tem tu.docx")

    def nsx_mark(self) -> str:
        """Ký hiệu ghi chú cho mức công bố của nhà sản xuất (inflow/downflow/lux có limit_source="nsx"):
        ⁽²⁾, hoặc ⁽³⁾ khi ⁽²⁾ đã dùng cho nồng độ hạt bụi; "" nếu không dùng."""
        res = self.dossier.results_data or {}
        if not any((res.get(k) or {}).get("limit_source") == "nsx" for k in ("inflow", "downflow", "lux")):
            return ""
        return "⁽³⁾" if res.get("particle") else "⁽²⁾"

    def nsx_note(self, is_en: bool = False) -> str:
        """Dòng ghi chú dưới bảng kết quả cho mức NSX (results_data["nsx"]["basis_vn" / "basis_en"])."""
        mark = self.nsx_mark()
        if not mark:
            return ""
        nsx = (self.dossier.results_data or {}).get("nsx") or {}
        if is_en:
            return f"\n{mark} The threshold declared by the manufacturer" + (f": {nsx['basis_en']}." if nsx.get("basis_en") else ".")
        return f"\n{mark} Mức công bố của nhà sản xuất" + (f": {nsx['basis_vn']}." if nsx.get("basis_vn") else ".")

    def basis_std(self) -> str:
        """Tài liệu của mức ⁽¹⁾: mặc định EN 12469:2000; results_data["limit_basis"] = "IEST-RP-CC002.4" khi người dùng
        chọn mức tham chiếu khác (tủ cấy 139329/2 theo chứng thư cũ 045837/2 — người dùng chốt 10/11/2026)."""
        return (self.dossier.results_data or {}).get("limit_basis") or "EN 12469:2000"

    def basis_note(self, is_en: bool = False) -> str:
        std = self.basis_std()
        if std == "EN 12469:2000":
            return "The threshold specified in EN 12469:2000." if is_en else "Mức quy định tại EN 12469:2000."
        return f"The threshold referenced in {std}." if is_en else f"Mức quy định tham chiếu tại {std}."

    def airflow_mark(self) -> str:
        """Ký hiệu ghi chú cho mức hình thái dòng khí tham chiếu tài liệu riêng
        (results_data["airflow_pattern"]["basis_vn" / "basis_en"], vd CAG-003:2022): số kế tiếp sau
        ⁽¹⁾ EN 12469, ⁽²⁾ bụi, mức NSX; "" nếu không dùng. Dòng "rows" tự chọn ghi "{ref}" ở chỗ đặt ký hiệu."""
        res = self.dossier.results_data or {}
        af = res.get("airflow_pattern") or {}
        if af.get("skip") or not af.get("basis_vn"):
            return ""
        n = 2 + bool(res.get("particle")) + bool(self.nsx_mark())
        return "⁽" + str(n).translate(str.maketrans("0123456789", "⁰¹²³⁴⁵⁶⁷⁸⁹")) + "⁾"

    def extra_notes(self, is_en: bool = False) -> str:
        """Các dòng ghi chú sau ⁽¹⁾/⁽²⁾: mức NSX rồi mức tham chiếu hình thái dòng khí."""
        txt = self.nsx_note(is_en)
        mark = self.airflow_mark()
        if mark:
            af = self.dossier.results_data["airflow_pattern"]
            if is_en:
                txt += f"\n{mark} The threshold referenced in {af.get('basis_en') or af['basis_vn']}."
            else:
                txt += f"\n{mark} Mức quy định tham chiếu tại {af['basis_vn']}."
        return txt

    def result_rows(self, is_en: bool = False) -> list:
        """Dòng bảng kết quả TBKQ (mục 10.3), dùng chung cho TBKQ và phần Đánh giá BCKQ.

        Mỗi dòng: (stt, chỉ tiêu, kết quả, mức quy định, đánh giá, is_parent, is_header, keep_next_parent, height_twips).
        """
        d = self.dossier
        dec = "." if is_en else ","
        res = d.results_data
        inflow_res = res.get("inflow", {}) or {}
        df_res = res.get("downflow", {}) or {}
        lux_res = res.get("lux", {}) or {}
        noise_res = res.get("noise", {}) or {}
        hepa_res = res.get("hepa", {}) or {}
        uvc_res = res.get("uvc", {})
        airflow_res = res.get("airflow_pattern", {}) or {}
        has_inflow = not inflow_res.get("skip", False)
        has_downflow = not df_res.get("skip", False)
        has_supply = not hepa_res.get("skip_downflow", False)
        has_exhaust = not hepa_res.get("skip_exhaust", False)
        has_airflow = not airflow_res.get("skip", False)
        has_part = bool(res.get("particle"))

        v_in_val = inflow_res.get("velocity", 0.40)
        try:
            v_in = f"{float(v_in_val):.2f}".replace(".", dec)
        except Exception:
            v_in = str(v_in_val).replace(".", dec)

        v_df_val = df_res.get("avg", 0.28)
        try:
            v_df = f"{float(v_df_val):.2f}".replace(".", dec)
        except Exception:
            v_df = str(v_df_val).replace(".", dec)
        lux_val = str(lux_res.get("avg_rounded", int(round(lux_res.get("avg", 750)))) if lux_res else "750")

        noise_amb = str(noise_res.get("ambient", "-")).replace(".", dec)
        noise_op = str(noise_res.get("corrected_dba", noise_res.get("operating", "-"))).replace(".", dec)

        def _format_leak(val: Any) -> Tuple[str, str]:
            """Formats HEPA leak value and assessment.
            Quy tắc: Với BSC quét rò rỉ lọc là 0,0000% thì thể hiện trên văn bản là:
            - Tiếng Việt: '< 0,0001%'
            - Tiếng Anh: '< 0.0001%'
            Nếu >= 0.0001%: thể hiện 4 chữ số thập phân (vd: '0,0006 %', '0,2578 %').
            """
            if val is None:
                fmt = "< 0.0001%" if is_en else "< 0,0001%"
                ass = "Pass" if is_en else "Đạt"
                return fmt, ass
            s_val = str(val).strip()
            if "<" in s_val:
                fmt = "< 0.0001%" if is_en else "< 0,0001%"
                ass = "Pass" if is_en else "Đạt"
                return fmt, ass
            try:
                num = float(s_val.replace("%", "").replace(",", ".").strip())
                if num < 0.0001:
                    fmt = "< 0.0001%" if is_en else "< 0,0001%"
                    ass = "Pass" if is_en else "Đạt"
                else:
                    fmt = f"{num:.4f}".replace(".", dec) + " %"
                    ass = ("Pass" if is_en else "Đạt") if num <= 0.01 else ("Fail" if is_en else "Không đạt")
                return fmt, ass
            except Exception:
                if s_val in ("0", "0.0", "0,0", "0.00", "0,00", "0.000", "0,000", "0.0000", "0,0000", "0%", "0.0%", "0,0%", "0.0000%", "0,0000%"):
                    return ("< 0.0001%" if is_en else "< 0,0001%"), ("Pass" if is_en else "Đạt")
                return f"{s_val.replace('.', dec)} %", ("Pass" if is_en else "Đạt")

        df_leak, df_leak_assess = _format_leak(hepa_res.get('downflow_pct', 0.0))
        ex_leak, ex_assess = _format_leak(hepa_res.get('exhaust_pct', 0.0))
        lim_leak = "≤ 0.01% ⁽¹⁾" if is_en else "≤ 0,01% ⁽¹⁾"
        not_spec = "Not specified" if is_en else "Không quy định"

        # Dynamic assessments for all parameters
        # Mức inflow: mặc định ≥ 0,40 m/s (BSC cấp II); inflow["limit_range"] = (min, max) cho BSC cấp I (EN 12469: 0,7 - 1,0 m/s)
        # inflow["limit_source"] = "nsx" -> mức công bố của nhà sản xuất (vd nhãn ESCO 0,45 ± 0,025 m/s — 136247)
        in_range = inflow_res.get("limit_range")
        in_mark = self.nsx_mark() if inflow_res.get("limit_source") == "nsx" else "⁽¹⁾"

        def _num(x: float) -> str:   # giữ chữ số có nghĩa (0,425), tối thiểu 2 số lẻ
            s = f"{x:.3f}".rstrip("0")
            return s if len(s.split(".")[1]) >= 2 else f"{x:.2f}"
        if in_range:
            in_lim = f"({_num(in_range[0])}-{_num(in_range[1])}) m/s {in_mark}".replace(".", dec)
        else:
            in_lim = "≥ 0.40 m/s ⁽¹⁾" if is_en else "≥ 0,40 m/s ⁽¹⁾"
        v_in_ok = True
        try:
            v_in_f = float(str(v_in_val).replace(",", ".").strip())
            v_in_ok = (in_range[0] <= v_in_f <= in_range[1]) if in_range else v_in_f >= 0.40
        except Exception as _err:
            print(f"[Warning] bsc.py inflow {type(_err).__name__}: {_err}")
        inflow_assess = ("Pass" if is_en else "Đạt") if v_in_ok else ("Fail" if is_en else "Không đạt")

        # Mức downflow / lux: mặc định EN 12469 ⁽¹⁾; downflow["limit_range"] = (min, max), lux["limit_min"] = x,
        # kèm limit_source="nsx" thì ghi theo mức công bố của nhà sản xuất (ghi chú self.nsx_mark()).
        nsx_mark = self.nsx_mark()
        df_lo, df_hi = df_res.get("limit_range") or (0.25, 0.50)
        df_mark = nsx_mark if df_res.get("limit_source") == "nsx" else "⁽¹⁾"
        lux_min = float(lux_res.get("limit_min", 750))
        lux_mark = nsx_mark if lux_res.get("limit_source") == "nsx" else "⁽¹⁾"
        lux_min_txt = f"{int(lux_min):,}".replace(",", "," if is_en else " ") if lux_min == int(lux_min) else str(lux_min).replace(".", dec)

        v_df_ok = True
        try:
            df_f = float(str(v_df_val).replace(",", ".").strip())
            v_df_ok = (df_lo <= df_f <= df_hi)
        except Exception as _err:
            print(f"[Warning] bsc.py downflow {type(_err).__name__}: {_err}")
        df_assess = ("Pass" if is_en else "Đạt") if v_df_ok else ("Fail" if is_en else "Không đạt")

        lux_ok = True
        try:
            lux_ok = float(str(lux_val).replace(",", ".").strip()) >= lux_min
        except Exception as _err:
            print(f"[Warning] bsc.py lux {type(_err).__name__}: {_err}")
        lux_assess = ("Pass" if is_en else "Đạt") if lux_ok else ("Fail" if is_en else "Không đạt")

        # Noise evaluation: mặc định <= 65.0 dBA (EN 12469); noise["limit_max"] ghi đè (vd IEST-RP-CC002.4: 70)
        noise_max = float(noise_res.get("limit_max", 65.0))
        noise_lim_txt = (f"≤ {noise_max:.1f} dBA ⁽¹⁾" if noise_max != int(noise_max) or noise_max == 65.0
                         else f"≤ {int(noise_max)} dBA ⁽¹⁾").replace(".", dec)
        noise_val_raw = noise_res.get("corrected_dba", noise_res.get("operating"))
        noise_ok = True
        if noise_val_raw is not None:
            try:
                noise_num = float(str(noise_val_raw).replace(",", ".").strip())
                noise_ok = noise_num <= noise_max
            except (ValueError, TypeError):
                pass

        if "assessment" in noise_res and noise_res["assessment"] in ("Đạt", "Không đạt"):
            noise_assess_vn = noise_res["assessment"]
            noise_assess_en = noise_res.get("assessment_en", "Pass" if noise_assess_vn == "Đạt" else "Fail")
        else:
            noise_assess_vn = "Đạt" if noise_ok else "Không đạt"
            noise_assess_en = "Pass" if noise_ok else "Fail"
        noise_assess = noise_assess_en if is_en else noise_assess_vn

        pass_txt = "Pass" if is_en else "Đạt"

        # rows_def: (stt, param, result, limit, assessment, is_parent, is_header, keep_next_parent, height_twips)
        rows_def = [("No." if is_en else "STT", "Parameter" if is_en else "Chỉ tiêu thực hiện", "Result" if is_en else "Kết quả đo",
                     "Threshold ⁽¹⁾" if is_en else "Mức quy định", "Assessment" if is_en else "Đánh giá", False, True, False, 440)]
        stt_no = 0

        def _next_stt() -> str:
            nonlocal stt_no
            stt_no += 1
            return str(stt_no)

        if has_inflow:
            rows_def.append((_next_stt(), "Average inflow velocity" if is_en else "Tốc độ trung bình dòng khí qua cửa làm việc", f"{v_in} m/s",
                             in_lim, inflow_assess, False, False, False, 760))
        if df_res.get("no_eval"):
            df_lim, df_ass = not_spec, "-"
        else:
            df_lim, df_ass = f"({df_lo:.2f}-{df_hi:.2f}) m/s {df_mark}".replace(".", dec), df_assess
        if has_downflow:
            rows_def.append((_next_stt(), "Average downflow velocity" if is_en else "Tốc độ trung bình dòng khí cấp qua lọc HEPA", f"{v_df} m/s",
                             df_lim, df_ass, False, False, False, 760))
        if lux_res.get("no_eval"):
            lux_lim, lux_ass = not_spec, "-"
        else:
            lux_lim, lux_ass = f"≥ {lux_min_txt} lux {lux_mark}", lux_assess
        rows_def.append((_next_stt(), "Average illuminance" if is_en else "Cường độ sáng trung bình bên trong tủ", f"{lux_val} lux",
                         lux_lim, lux_ass, False, False, False, 580))
        rows_def.append((_next_stt(), "Sound pressure level" if is_en else "Đo độ ồn", "", "", "", True, False, False, 440))
        rows_def.append(("", "Background noise" if is_en else "Độ ồn nền", f"{noise_amb} dBA", not_spec, "-", False, False, False, 440))
        rows_def.append(("", "Operation noise" if is_en else "Độ ồn khi vận hành", f"{noise_op} dBA", noise_lim_txt,
                         noise_assess, False, False, False, 440))
        hepa_vn = "Mức rò rỉ của các lọc HEPA" if (has_supply and has_exhaust) else "Mức rò rỉ của lọc HEPA"
        rows_def.append((_next_stt(), "HEPA filter leakage test" if is_en else hepa_vn, "", "", "", True, False, False, 440))
        if has_supply:
            rows_def.append(("", "Downflow filter" if is_en else "Lọc khí cấp", df_leak, lim_leak, df_leak_assess, False, False, False, 580))
        if has_exhaust:
            rows_def.append(("", "Exhaust filter" if is_en else "Lọc khí thải", ex_leak, lim_leak, ex_assess, False, False, False, 580))

        if has_airflow:
            rows_def.append((_next_stt(), "Airflow pattern test" if is_en else "Kiểm tra hình thái dòng khí", "", "", "", True, False, True, 440))
            custom_af = airflow_res.get("rows")
            if custom_af:
                af_mark = self.airflow_mark()
                for af in custom_af:
                    sfx = "en" if is_en else "vn"
                    rows_def.append(("", af.get(f"param_{sfx}", ""), af.get(f"result_{sfx}", ""),
                                     af.get(f"limit_{sfx}", "").replace("{ref}", af_mark).rstrip(),
                                     af.get(f"assess_{sfx}", pass_txt), False, False, False, af.get("height", 1000)))
            else:
                rows_def.append((
                    "",
                    "Inside cabinet workspace\n- Along center line, above work opening\n- Behind work sash\n- Along edge of work sash" if is_en else "Bên trong không gian tủ\n- Dọc đường trung tâm, phía trên khoảng mở làm việc\n- Phía sau cửa mở làm việc\n- Dọc theo mép cửa mở làm việc",
                    "- Smooth downward flow, no dead spots;\n- No smoke escaped." if is_en else "- Dòng khói hướng xuống, không có điểm tù;\n- Khói không thoát ra ngoài tủ.",
                    "- Smooth downward flow, no dead spots;\n- No smoke escaped. ⁽¹⁾" if is_en else "- Dòng khói hướng xuống, không có điểm tù;\n- Khói không thoát ra ngoài tủ. ⁽¹⁾",
                    pass_txt, False, False, False, 1000
                ))
                rows_def.append((
                    "",
                    "Area along the outer perimeter of the work opening" if is_en else "Khu vực dọc theo chu vi phía bên ngoài của vùng mở cửa tủ",
                    "- Smoke is drawn into grilles, not spreading." if is_en else "- Dòng khói hướng vào các lưới đục lỗ trên tủ, không lan trên bề mặt làm việc.",
                    "- Smoke is drawn into grilles, not spreading. ⁽¹⁾" if is_en else "- Dòng khói hướng vào các lưới đục lỗ trên tủ, không lan trên bề mặt làm việc. ⁽¹⁾",
                    pass_txt, False, False, False, 600
                ))

        if uvc_res:
            uvc_val = str(uvc_res.get("avg_rounded", int(round(uvc_res.get("avg", 100)))))
            rows_def.append((_next_stt(), "UVC radiation intensity" if is_en else "Cường độ tia UV dải C", f"{uvc_val} µW/cm²", not_spec, "-", False, False, False, 440))   # không có mức quy định -> "-" (người dùng chốt 30/10/2026)

        if has_part:
            part_res = d.results_data.get("particle", {})
            p_title = "Airborne particle concentration" if is_en else "Nồng độ hạt bụi trong tủ"
            rows_def.append((_next_stt(), p_title, "", "", "", True, False, True, 440))
            sizes = part_res.get("sizes")
            if not sizes:
                sizes = [
                    ("Kích thước hạt d ≥ 0,3 µm", "Particle size d ≥ 0.3 µm", "0 hạt/m³", "0 particles/m³", "≤ 10 200\nhạt/m³ ⁽²⁾", "≤ 10,200\nparticles/m³ ⁽²⁾", pass_txt, pass_txt),
                    ("Kích thước hạt d ≥ 0,5 µm", "Particle size d ≥ 0.5 µm", "0 hạt/m³", "0 particles/m³", "≤ 3 520\nhạt/m³ ⁽²⁾", "≤ 3,520\nparticles/m³ ⁽²⁾", pass_txt, pass_txt),
                    ("Kích thước hạt d ≥ 1,0 µm", "Particle size d ≥ 1.0 µm", "0 hạt/m³", "0 particles/m³", "≤ 832\nhạt/m³ ⁽²⁾", "≤ 832\nparticles/m³ ⁽²⁾", pass_txt, pass_txt),
                    ("Kích thước hạt d ≥ 5,0 µm", "Particle size d ≥ 5.0 µm", "0 hạt/m³", "0 particles/m³", "≤ 29\nhạt/m³ ⁽²⁾", "≤ 29\nparticles/m³ ⁽²⁾", pass_txt, pass_txt),
                ]
            for s_info in sizes:
                lbl = s_info[1] if is_en else s_info[0]
                val = s_info[3] if is_en else s_info[2]
                lim = s_info[5] if is_en else s_info[4]
                # Ensure unit (particles/m³ or hạt/m³) consistently wraps to line 2 for neat alignment
                if "particles/m³" in lim and "\n" not in lim:
                    lim = lim.replace(" particles/m³", "\nparticles/m³")
                elif "hạt/m³" in lim and "\n" not in lim:
                    lim = lim.replace(" hạt/m³", "\nhạt/m³")
                ass = s_info[7] if is_en else s_info[6]
                rows_def.append(("", lbl, val, lim, ass, False, False, False, 560))

        # Vi sinh không khí trong tủ (kết quả theo phiếu KQTN): results_data["microbial"] =
        # {"title_vn", "title_en", "points": [{"name_vn", "name_en", "result_vn", "result_en"}]} — không có mức, không đánh giá
        mic = res.get("microbial") or {}
        if mic.get("points"):
            rows_def.append((_next_stt(), mic.get("title_en" if is_en else "title_vn", ""), "", "", "", True, False, True, 440))
            sub_t = mic.get("subtitle_en" if is_en else "subtitle_vn")
            if sub_t:
                rows_def.append(("", sub_t, "", not_spec, "-", True, False, True, 440))
            for pt in mic["points"]:
                sfx = "en" if is_en else "vn"
                rows_def.append(("", pt.get(f"name_{sfx}", ""), pt.get(f"result_{sfx}", ""), not_spec, "-", False, False, False, 440))
        return rows_def

    def build_tbkq(self, is_en: bool = False, **kwargs) -> str:
        """Builds complete TBKQ document (VN or EN) matching QUATEST 3 standard form."""
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
        c_info = d.results_data.get("cabinet", {})
        client_name = d.client.name_en if is_en else d.client.name_vn

        # --- 10.1 Thông tin thiết bị ---
        h101 = "10.1 Device information" if is_en else "10.1 Thông tin thiết bị"
        add_p(h101, bold=True, font_size=13, keep_next=True, space_before=4)

        brand_val = c_info.get('brand', '')
        model_val = c_info.get('model', '')
        serial_val = c_info.get('serial', '')
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
        cab_name_en = c_info.get('cabinet_name_en')
        v_en = (f"- As validated by the representative of {client_name}, the \"{cab_name_en}\" has been cleaned and operated under normal conditions during the inspection."
                if cab_name_en else
                f"- As validated by the representative of {client_name}, the Biosafety Cabinet class II has been cleaned and operated under normal conditions during the inspection.")
        add_p(v_en if is_en else v_vn, font_size=13, keep_next=True, space_after=6)

        # --- 10.2 Nội dung thực hiện ---
        h102 = "10.2 Performance content" if is_en else "10.2 Nội dung thực hiện"
        add_p(h102, bold=True, font_size=13, keep_next=True, space_before=4)

        # Build Results Table
        res = d.results_data
        inflow_res = res.get("inflow", {}) or {}
        df_res = res.get("downflow", {}) or {}
        lux_res = res.get("lux", {}) or {}
        noise_res = res.get("noise", {}) or {}
        hepa_res = res.get("hepa", {}) or {}
        uvc_res = res.get("uvc", {})
        airflow_res = res.get("airflow_pattern", {}) or {}

        # Tùy chọn bỏ chỉ tiêu (tương thích ngược: mặc định giữ nguyên bố cục BSC cấp 2)
        # - inflow: {"skip": True}  -> không đo inflow (vd tủ cấy LAF)
        # - hepa: {"skip_exhaust": True} -> tủ không có lọc khí thải
        # - downflow: {"skip": True}, hepa: {"skip_downflow": True} -> tủ không có dòng khí cấp / lọc khí cấp (BSC cấp I)
        # - downflow/lux: {"no_eval": True} -> ghi "Không quy định" và không đánh giá
        # - airflow_pattern: {"rows": [...]} -> nội dung kiểm tra hình thái dòng khí riêng; {"skip": True} -> bỏ
        has_inflow = not inflow_res.get("skip", False)
        has_downflow = not df_res.get("skip", False)
        has_supply = not hepa_res.get("skip_downflow", False)
        has_exhaust = not hepa_res.get("skip_exhaust", False)
        has_airflow = not airflow_res.get("skip", False)
        has_part = bool(res.get("particle"))

        tasks_vn, tasks_en = [], []
        if has_inflow:
            tasks_vn.append("Đo tốc độ dòng khí qua cửa làm việc;"); tasks_en.append("Inflow air velocity test;")
        if has_downflow:
            tasks_vn.append("Đo tốc độ dòng khí cấp qua lọc HEPA;"); tasks_en.append("Downflow air velocity test;")
        tasks_vn += ["Đo cường độ sáng bên trong tủ;", "Đo độ ồn của tủ khi vận hành;",
                     "Kiểm tra độ rò rỉ của các lọc HEPA;" if has_supply and has_exhaust else "Kiểm tra độ rò rỉ của lọc HEPA;"]
        tasks_en += ["Illuminance test in the cabinet workspace;", "Noise test of the cabinet in operation;",
                     "HEPA filter leakage test;"]
        if has_airflow:
            tasks_vn.append("Kiểm tra hình thái dòng khí của tủ;"); tasks_en.append("Airflow patterns test;")
        if uvc_res:
            tasks_vn.append("Đo cường độ tia UV dải C trong khoang làm việc;"); tasks_en.append("UVC intensity measurement in the workspace;")
        if has_part:
            tasks_vn.append("Đo nồng độ hạt bụi trong tủ;"); tasks_en.append("Particle concentration measurement in the workspace;")
        if (res.get("microbial") or {}).get("points"):
            mic_info = res.get("microbial") or {}
            t_vn = mic_info.get("task_vn", "Lấy mẫu vi sinh không khí trong tủ;")
            t_en = mic_info.get("task_en", "Airborne microbial sampling in the workspace;")
            tasks_vn.append(t_vn); tasks_en.append(t_en)
        if res.get("tasks_vn"):
            tasks_vn = list(res["tasks_vn"])
        if res.get("tasks_en"):
            tasks_en = list(res["tasks_en"])
        for lst in (tasks_vn, tasks_en):
            if lst:
                lst[-1] = lst[-1].rstrip(".;") + "."
        for i, t_txt in enumerate(tasks_en if is_en else tasks_vn, start=1):
            add_p(f"{i}. {t_txt}", font_size=13, keep_next=True)

        # --- 10.3 Kết quả giám định ---
        h103 = "10.3. Inspection result" if is_en else "10.3. Kết quả giám định"
        add_p(h103, bold=True, font_size=13, keep_next=True, space_before=6)

        rows_def = self.result_rows(is_en)

        # Khách hàng yêu cầu không thể hiện đánh giá (dossier.has_assessment=False, vd 134792):
        # bảng 3 cột STT | Chỉ tiêu | Kết quả đo — bỏ cả cột mức và cột đánh giá
        has_ass = getattr(d, "has_assessment", True)
        ncols = 5 if has_ass else 3
        tbl = doc.add_table(rows=len(rows_def), cols=ncols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        widths = [700, 3600, 2000, 2200, 1100] if has_ass else [700, 5400, 3500]

        for r_idx, r_data in enumerate(rows_def):
            stt, param, res_v, lim_v, assess_v, is_parent, is_header, kn_parent, _h = r_data
            row = tbl.rows[r_idx]
            add_cant_split(row)
            if is_header:
                add_tbl_header(row)

            for c_idx in range(ncols):
                c = row.cells[c_idx]
                tcPr = c._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:tcW xmlns:w="{W_NS}" w:w="{widths[c_idx]}" w:type="dxa"/>'))

            if is_header:
                for c_idx, txt in enumerate((stt, param, res_v, lim_v, assess_v)[:ncols]):
                    set_cell_content(row.cells[c_idx], txt, bold=True, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex="E7E6E6", valign="top", space_before=3, space_after=3)
            elif is_parent:
                # Parent rows are explicitly NOT BOLD per user requirement
                set_cell_content(row.cells[0], stt, bold=False, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                merged_param = row.cells[1].merge(row.cells[ncols - 1])
                set_cell_content(merged_param, param, bold=False, font_size=13, align=WD_ALIGN_PARAGRAPH.LEFT, valign="top", space_before=3, space_after=3)
                for c in row.cells:
                    for p in c.paragraphs:
                        p.paragraph_format.keep_with_next = bool(kn_parent)
            else:
                set_cell_content(row.cells[0], stt, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[1], param, font_size=13, align=WD_ALIGN_PARAGRAPH.LEFT, valign="top", space_before=3, space_after=3)
                set_cell_content(row.cells[2], res_v, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                if has_ass:
                    set_cell_content(row.cells[3], lim_v, font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                    set_cell_content(row.cells[4], assess_v, bold=(assess_v != "-"), italic=(assess_v != "-"), font_size=13, align=WD_ALIGN_PARAGRAPH.CENTER, valign="top", space_before=3, space_after=3)
                if "Inside cabinet workspace" in param or "Bên trong không gian tủ" in param:
                    for c in row.cells:
                        for p in c.paragraphs:
                            p.paragraph_format.keep_with_next = True

        has_uvc = bool(uvc_res)
        if has_inflow and has_downflow and has_supply and has_exhaust and has_airflow and not airflow_res.get("rows"):
            bsc_row_heights = get_bsc_row_heights(has_particle=has_part, has_uvc=has_uvc)
        else:
            bsc_row_heights = [r[8] for r in rows_def]
        if not has_ass:   # bảng 3 cột: cột chỉ tiêu rộng, các dòng đều 1 dòng chữ (trừ dòng nhiều dòng tự đặt)
            bsc_row_heights = [440 if h <= 760 else h for h in bsc_row_heights]
        format_table_structure(tbl, is_note=False, widths=widths, row_height=bsc_row_heights)
        merge_stt_column_child_rows(tbl, stt_col_idx=0, start_row=1, font_size=13, set_uniform_spacing=True)
        if not has_ass:
            # giữ trọn mỗi chỉ tiêu nhiều dòng con trên 1 trang để VN/EN ngắt cùng chỗ (AGENTS mục 9)
            for g_start, r in enumerate(rows_def):
                if not r[5]:
                    continue
                g_end = next((i for i in range(g_start + 1, len(rows_def)) if rows_def[i][0] or rows_def[i][5]), len(rows_def))
                for i in range(g_start, g_end - 1):
                    for p in tbl.rows[i]._tr.iter(qn("w:p")):
                        Paragraph(p, None).paragraph_format.keep_with_next = True
        # Hình thái dòng khí: giữ trọn chỉ tiêu trên 1 trang (keepNext mọi đoạn, kể cả ô STT merge dọc, trừ dòng con cuối)
        # để VN/EN ngắt cùng chỗ (AGENTS mục 9); đặt sau merge vì merge ghi lại đoạn ô STT.
        if has_airflow:
            af_start = next(i for i, r in enumerate(rows_def) if r[5] and r[1] in ("Airflow pattern test", "Kiểm tra hình thái dòng khí"))
            af_end = next((i for i in range(af_start + 1, len(rows_def)) if rows_def[i][0] or rows_def[i][5]), len(rows_def))
            for i in range(af_start, af_end - 1):
                for p in tbl.rows[i]._tr.iter(qn("w:p")):
                    Paragraph(p, None).paragraph_format.keep_with_next = True

        # Note Table (Font 10 per AGENTS.md Rule 3, borderless)
        t_note = doc.add_table(rows=1, cols=2)
        set_cell_content(t_note.rows[0].cells[0], "Note:" if is_en else "Ghi chú:", bold=True, font_size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)
        note_msg = (
            "The performance field is accredited to ISO 17020 - VIAS 004.\n⁽¹⁾ " + self.basis_note(True)
            + ("\n⁽²⁾ The threshold specified for ISO Class 5 in ISO 14644-1:2015." if has_part else "")
            if is_en else
            "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.\n⁽¹⁾ " + self.basis_note(False)
            + ("\n⁽²⁾ Mức quy định đối với cấp độ sạch ISO 5 tại ISO 14644-1:2015." if has_part else "")
        )
        note_msg += self.extra_notes(is_en)
        if not has_ass:   # không có cột mức -> chỉ giữ câu công nhận
            note_msg = note_msg.split("\n")[0]
        set_cell_content(t_note.rows[0].cells[1], note_msg, font_size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_table_structure(t_note, is_note=True, widths=[1000, 8600])

        # Sơ đồ vị trí lấy mẫu vi sinh (results_data["microbial"]["diagram"] / ["diagram_en"]: ảnh tự sinh)
        mic = res.get("microbial") or {}
        mic_diag = (mic.get("diagram_en") if is_en else None) or mic.get("diagram")
        if mic_diag and os.path.isfile(mic_diag):
            diag_h = mic.get("heading_en" if is_en else "heading_vn",
                             "Microbial sampling locations diagram:" if is_en else "Sơ đồ vị trí lấy mẫu vi sinh:")
            add_p(diag_h, bold=True, font_size=13, keep_next=True, space_before=6, space_after=2)
            t_diag = doc.add_table(rows=1, cols=1)
            t_diag.alignment = WD_TABLE_ALIGNMENT.CENTER
            p_d = t_diag.rows[0].cells[0].paragraphs[0]
            p_d.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_d.paragraph_format.keep_with_next = True
            p_d.add_run().add_picture(mic_diag, width=Cm(6.5))
            cap = mic.get("caption_en" if is_en else "caption_vn", "")
            if cap:
                p_c = t_diag.rows[0].cells[0].add_paragraph()
                p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_c = p_c.add_run(cap)
                r_c.italic = True
                r_c.font.name = "Times New Roman"
                r_c.font.size = Pt(13)
            add_cant_split(t_diag.rows[0])
            set_table_autofit_window(t_diag)

        # --- 10.4 Hình ảnh quá trình thực hiện ---
        h104 = "10.4. Photos of performance process" if is_en else "10.4. Hình ảnh quá trình thực hiện"
        # Không ép ngắt trang trước mục ảnh (AGENTS mục 3): ảnh chạy tiếp sau bảng kết quả nếu còn chỗ
        # (người dùng yêu cầu 30/10/2026 — 132488/1 ảnh nằm riêng trang 4 dù trang 3 còn trống).
        add_p(h104, bold=True, font_size=13, keep_next=True, space_before=2, space_after=1)

        # Helper to optimize and fix EXIF orientation for Word
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
                print(f"[BSCHandler] Photo optimization fallback: {e}")
                return p_path

        # Gather photos
        photo_paths = []
        if d.photos and len(d.photos) > 0:
            for p in d.photos:
                p_path = p.get("path") if isinstance(p, dict) else p
                if os.path.exists(p_path):
                    photo_paths.append(p_path)
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
                photo_paths.extend(cands[:6])

        # Lay out photos cleanly: lưới 2 cột (7 - 8 cm) cho mọi số ảnh
        # Quy tắc: Grid 2 cột thì kích thước ngang của ảnh 7 - 8cm (PHOTO_WIDTH_2COL = 7.5cm), tuyệt đối không crop ảnh
        # NO captions per user request: "không cần ghi chú nội dung hình"
        if photo_paths:
            # lưới 2 cột cho mọi bảng ảnh; đúng 3 ảnh -> 1 hàng 3 cột (người dùng chốt 20/11/2026, AGENTS mục 3)
            num_cols = 3 if len(photo_paths) == 3 else 2
            num_rows = math.ceil(len(photo_paths) / num_cols)
            img_w = PHOTO_WIDTH_3COL if num_cols == 3 else PHOTO_WIDTH_2COL  # 4,8 / 7,5 cm, không crop

            t_photo = doc.add_table(rows=num_rows, cols=num_cols)
            t_photo.alignment = WD_TABLE_ALIGNMENT.CENTER

            for idx, p_path in enumerate(photo_paths):
                r_idx = idx // num_cols
                c_idx = idx % num_cols
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

        # Standard Ending Line (4.5pt thinThick centered)
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
        file_label = d.results_data.get("file_label", "BSC")
        file_name = f"TBKQ - {lang_code}{sub_str} - {file_label} - {d.code}.docx"

        target_tbkq = os.path.join(d.tbkq_dir, file_name)

        from ..tag_context import style_note_tables
        style_note_tables(doc)          # bảng Ghi chú 11pt in nghiêng (09/11/2026)
        doc.save(target_tbkq)
        if d.output_dir:
            try:
                shutil.copy2(target_tbkq, os.path.join(d.output_dir, file_name))
            except Exception as _err:
                print(f"[Warning] bsc.py:521 {type(_err).__name__}: {_err}")

        print(f"[BSCHandler] Generated TBKQ {'EN' if is_en else 'VN'} -> {target_tbkq}")
        return target_tbkq

    def build_tem_tu(self, template_path: Optional[str] = None) -> str:
        """Tem kiểm tra dán tủ BSC từ template TAG Core/Templates/tem tu.docx (TAG: so_tem, serial, model, ngày kiểm tra...)."""
        d = self.dossier
        return build_tem_tu_doc(d, template_path or self.tem_tu_template,
                                (f"{d.code.replace('Q326D', '')} - {d.sub_no} - tem tu.docx" if (d.is_multi_cert and d.sub_no)
                                 else f"{d.code.replace('Q326D', '')} - tem tu.docx"),
                                so_tem=(d.cert_no_vn if (d.is_multi_cert and d.sub_no) else d.code),
                                ten_thiet_bi=(d.results_data.get("cabinet", {}) or {}).get("seal_item_name", "“Tủ cấy an toàn sinh học cấp 2”"),
                                label="BSCHandler")
