# -*- coding: utf-8 -*-
"""
BSC (Biosafety Cabinet Class II) Calculation Engine & Automated Excel Generator
Compliant with EN 12469:2000 and QUATEST 3 standards.

Features:
- Inflow calculation via DIM (Q/A/3600) or Kanomax sash edge (V_tb * k)
- Downflow calculation: Grid average, Min, Max, Uniformity tolerance (10% or 20%), limits check
- Lux calculation: Workspace grid average (threshold >= 750 lux)
- Noise level calculation: Ambient noise, operating noise, delta N, table-based correction (<= 65 dBA)
- HEPA filter leakage evaluation (Downflow & Exhaust <= 0.01%)
- Airflow pattern assessment (work space and front sash)
- UVC radiation intensity calculation (average in uW/cm2)
- Automated generation of standardized Excel calculation workbook (.xlsx)
"""

import os
import math
from typing import List, Dict, Any, Optional
import openpyxl
from .tag_excel import fill_workbook_tags
from .noise_corr import background_noise_correction, noise_correction_value, noise_correction_formula
from .excel_borders import strike_box, thin_outside_borders
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from openpyxl.drawing.image import Image as OpenpyxlImage
    from PIL import Image as PILImage, ImageOps
    HAS_PIL_IMAGE = True
except ImportError:
    HAS_PIL_IMAGE = False

try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

# Font chữ hình vẽ: Times New Roman (AGENTS mục 3). Máy không có Times New Roman (vd Linux)
# thì dùng font có số đo tương đương (Liberation Serif / Tinos / Nimbus Roman), không rơi về DejaVu Sans.
_DIAGRAM_FONT_CANDIDATES = ["Times New Roman", "Liberation Serif", "Tinos", "Nimbus Roman", "DejaVu Serif"]
DIAGRAM_FONT = "Times New Roman"
if HAS_MATPLOTLIB:
    from matplotlib import font_manager as _fm
    _installed = {f.name for f in _fm.fontManager.ttflist}
    DIAGRAM_FONT = next((f for f in _DIAGRAM_FONT_CANDIDATES if f in _installed), "serif")
    if DIAGRAM_FONT != "Times New Roman":
        print(f"[Warning] Không có font Times New Roman, hình vẽ dùng font thay thế: {DIAGRAM_FONT}")


def generate_bsc_downflow_grid_image(
    output_path: str,
    width_mm: int = 930,
    depth_mm: int = 450,
    velocity_grid=None,
    is_en: bool = False,
    dpi: int = 300
) -> str:
    """
    Generates a standardized QUATEST 3 computerized measurement layout illustration
    for BSC Downflow air velocity grid (Step 2 - Layout vi tính).
    """
    if not HAS_MATPLOTLIB:
        raise RuntimeError("matplotlib is required to generate BSC downflow grid images.")

    if velocity_grid is None or len(velocity_grid) == 0:
        velocity_grid = [
            [0.32, 0.35, 0.35, 0.36],
            [0.33, 0.34, 0.33, 0.32]
        ]

    num_rows = len(velocity_grid)
    num_cols = len(velocity_grid[0]) if num_rows > 0 else 1
    dec = "." if is_en else ","

    fig, ax = plt.subplots(figsize=(6.5, 2.8), dpi=dpi)
    ax.set_xlim(-0.08, 1.35)
    ax.set_ylim(-0.35, 1.25)
    ax.axis("off")

    rect_x, rect_y, rect_w, rect_h = 0.0, 0.0, 1.0, 1.0

    # 1. Main cabinet work surface
    rect = patches.Rectangle(
        (rect_x, rect_y), rect_w, rect_h,
        linewidth=1.2, edgecolor="black", facecolor="#FBFBFB"
    )
    ax.add_patch(rect)

    # 2. Grid lines
    for c in range(1, num_cols):
        gx = rect_x + c * (rect_w / num_cols)
        ax.plot([gx, gx], [rect_y, rect_y + rect_h], color="#999999", linestyle="--", lw=0.8)
    for r in range(1, num_rows):
        gy = rect_y + r * (rect_h / num_rows)
        ax.plot([rect_x, rect_x + rect_w], [gy, gy], color="#999999", linestyle="--", lw=0.8)

    # 3. Title
    title_text = "Downflow Velocity Measuring Grid" if is_en else "Sơ đồ bố trí các vị trí đo tốc độ gió cấp (Downflow)"
    ax.text(
        rect_x + rect_w / 2.0, rect_y + rect_h + 0.08,
        title_text,
        ha="center", va="bottom",
        fontfamily=DIAGRAM_FONT, fontsize=12, fontweight="bold", color="black"
    )

    # 4. Grid values
    for r_idx, row in enumerate(velocity_grid):
        y_pos = rect_y + rect_h - (r_idx + 0.5) * (rect_h / num_rows)
        for c_idx, val in enumerate(row):
            x_pos = rect_x + (c_idx + 0.5) * (rect_w / num_cols)
            val_str = f"{float(val):.2f}".replace(".", dec)
            ax.plot(x_pos, y_pos + 0.08, marker="o", markersize=4, color="#1A5276")
            ax.text(
                x_pos, y_pos - 0.05,
                val_str,
                ha="center", va="center",
                fontfamily=DIAGRAM_FONT, fontsize=11, fontweight="bold", color="black"
            )

    # 5. Bottom Dimension (Width)
    dim_y = rect_y - 0.20
    tick_len = 0.06
    ax.plot([rect_x, rect_x], [dim_y - tick_len / 2, dim_y + tick_len / 2], color="black", lw=1.0)
    ax.plot([rect_x + rect_w, rect_x + rect_w], [dim_y - tick_len / 2, dim_y + tick_len / 2], color="black", lw=1.0)
    ax.annotate(
        "", xy=(rect_x, dim_y), xytext=(rect_x + rect_w, dim_y),
        arrowprops=dict(arrowstyle="<|-|>", color="black", lw=1.0, mutation_scale=10)
    )
    w_label = f"Width: {width_mm} mm" if is_en else f"Chiều rộng: {width_mm} mm"
    ax.text(
        rect_x + rect_w / 2.0, dim_y + 0.02,
        w_label,
        ha="center", va="bottom",
        fontfamily=DIAGRAM_FONT, fontsize=10.5, color="black"
    )

    # 6. Right Dimension (Depth)
    dim_x = rect_x + rect_w + 0.03
    tick_w = 0.03
    ax.plot([dim_x - tick_w / 2, dim_x + tick_w / 2], [rect_y, rect_y], color="black", lw=1.0)
    ax.plot([dim_x - tick_w / 2, dim_x + tick_w / 2], [rect_y + rect_h, rect_y + rect_h], color="black", lw=1.0)
    ax.annotate(
        "", xy=(dim_x, rect_y), xytext=(dim_x, rect_y + rect_h),
        arrowprops=dict(arrowstyle="<|-|>", color="black", lw=1.0, mutation_scale=10)
    )
    h_label = f"Depth: {depth_mm} mm" if is_en else f"Chiều sâu:\n{depth_mm} mm"
    ax.text(
        dim_x + 0.04, rect_y + rect_h / 2.0,
        h_label,
        ha="left", va="center",
        fontfamily=DIAGRAM_FONT, fontsize=10.5, fontstyle="italic", color="black"
    )

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    plt.savefig(output_path, bbox_inches="tight", dpi=dpi, facecolor="white")
    plt.close(fig)
    return output_path


def calc_inflow_dim(q_m3h: float, width_mm: float, height_mm: float) -> Dict[str, Any]:
    """Calculates inflow air velocity from DIM flow rate and sash opening dimensions (Method 1)."""
    area_m2 = (width_mm * height_mm) / 1_000_000.0
    velocity = q_m3h / (area_m2 * 3600.0) if area_m2 > 0 else 0.0
    vel_rounded = round(velocity, 2)
    is_ok = vel_rounded >= 0.40
    return {
        "area_m2": round(area_m2, 4),
        "q_m3h": round(q_m3h, 2),
        "velocity": vel_rounded,
        "limit": "≥ 0.40",
        "assessment": "Đạt" if is_ok else "Không đạt",
        "assessment_en": "Pass" if is_ok else "Fail"
    }


def calc_inflow_kanomax(readings: Any, k_factor: Optional[float] = None) -> Dict[str, Any]:
    """Calculates inflow air velocity from Kanomax Climomaster 6501 readings (Method 2)."""
    flat_vals = []
    if isinstance(readings, list):
        for item in readings:
            if isinstance(item, list):
                flat_vals.extend([float(v) for v in item if v is not None])
            elif item is not None and isinstance(item, (int, float)):
                flat_vals.append(float(item))
    if not flat_vals:
        return {"avg": 0.0, "velocity": 0.0, "assessment": "N/A"}

    raw_avg = sum(flat_vals) / len(flat_vals)
    final_vel = raw_avg * float(k_factor) if (k_factor and float(k_factor) > 0) else raw_avg
    vel_rounded = round(final_vel, 2)
    is_ok = vel_rounded >= 0.40
    return {
        "count": len(flat_vals),
        "raw_avg": round(raw_avg, 3),
        "k_factor": k_factor,
        "velocity": vel_rounded,
        "velocity_precise": final_vel,
        "limit": "≥ 0.40",
        "assessment": "Đạt" if is_ok else "Không đạt",
        "assessment_en": "Pass" if is_ok else "Fail"
    }


DOWNFLOW_TOLERANCE_PCT = 20.0   # N2.TC.BSC.02 / N2.TC.LAF.01: mọi điểm trong Vtb ± 20 % (chốt 30/10/2026)


def calc_downflow(readings: List[List[float]], tolerance_pct: float = DOWNFLOW_TOLERANCE_PCT,
                  eval_uniformity: bool = False) -> Dict[str, Any]:
    """Tốc độ downflow trung bình (0,25 – 0,50 m/s) và độ đồng đều (mọi điểm trong Vtb ± tolerance_pct).

    Thông thường KHÔNG xét độ đồng đều (người dùng chốt 30/10/2026): độ đồng đều vẫn được tính sẵn
    (max, min, ngưỡng trên/dưới, is_uniform_ok) nhưng chỉ đưa vào đánh giá khi eval_uniformity=True
    (results_data["downflow"]["eval_uniformity"] = True)."""
    flat_vals = [float(v) for row in readings for v in row if v is not None]
    if not flat_vals:
        return {"avg": 0.0, "assessment": "N/A"}
    
    avg_val = sum(flat_vals) / len(flat_vals)
    max_val = max(flat_vals)
    min_val = min(flat_vals)
    
    upper_threshold = avg_val * (1.0 + tolerance_pct / 100.0)
    lower_threshold = avg_val * (1.0 - tolerance_pct / 100.0)
    
    is_speed_ok = 0.25 <= avg_val <= 0.50
    is_uniform_ok = (min_val >= lower_threshold) and (max_val <= upper_threshold)
    is_overall_ok = is_speed_ok and (is_uniform_ok or not eval_uniformity)

    return {
        "count": len(flat_vals),
        "avg": round(avg_val, 2),
        "avg_precise": avg_val,
        "max": round(max_val, 2),
        "min": round(min_val, 2),
        "tolerance_pct": tolerance_pct,
        "upper_threshold": round(upper_threshold, 3),
        "lower_threshold": round(lower_threshold, 3),
        "is_speed_ok": is_speed_ok,
        "is_uniform_ok": is_uniform_ok,
        "eval_uniformity": eval_uniformity,
        "assessment": "Đạt" if is_overall_ok else "Không đạt",
        "assessment_en": "Pass" if is_overall_ok else "Fail"
    }


def calc_lux(readings: List[List[float]], ambient_lux: Optional[float] = None) -> Dict[str, Any]:
    """Calculates average workspace illuminance."""
    flat_vals = [float(v) for row in readings for v in row if v is not None]
    if not flat_vals:
        return {"avg": 0.0, "assessment": "N/A"}
    
    avg_lux = sum(flat_vals) / len(flat_vals)
    is_ok = avg_lux >= 750.0
    return {
        "count": len(flat_vals),
        "avg": round(avg_lux, 1),
        "avg_rounded": int(round(avg_lux)),
        "ambient": ambient_lux,
        "limit": "≥ 750",
        "assessment": "Đạt" if is_ok else "Không đạt",
        "assessment_en": "Pass" if is_ok else "Fail"
    }


def calc_noise(ambient_dba: float, op_dba: float) -> Dict[str, Any]:
    """
    Độ ồn sau hiệu chỉnh ồn nền theo bảng của quy trình N2.TC.BSC.01/.02, LAF.01
    (Core/Tools/noise_corr.py — ΔN làm tròn số nguyên; ≤ 2 → không hiệu chỉnh, cảnh báo giảm ồn nền).
    """
    nc = background_noise_correction(ambient_dba, op_dba)
    if nc["warning"]:
        print(f"[WARNING] [Noise] {nc['warning']}")
    corrected_dba = nc["corrected_dba"]
    is_ok = corrected_dba <= 65.0

    return {
        "ambient": nc["ambient"],
        "operating": nc["operating"],
        "delta_n": nc["delta_n"],
        "delta_n_int": nc["delta_n_int"],
        "correction": nc["correction"],
        "corrected_dba": corrected_dba,
        "noise_status": nc["status"],
        "noise_warning": nc["warning"],
        "limit": "≤ 65,0",
        "assessment": "Đạt" if is_ok else "Không đạt",
        "assessment_en": "Pass" if is_ok else "Fail"
    }


def format_hepa_leak_val(leak_pct: Any, is_en: bool = False) -> str:
    """Formats HEPA filter leakage percentage according to QUATEST 3 standards.
    Rule: Với BSC quét rò rỉ lọc là 0,0000% thì thể hiện trên văn bản là '< 0,0001%' (VN) hoặc '< 0.0001%' (EN).
    Nếu >= 0.0001%: thể hiện 4 chữ số thập phân (vd: '0,0006 %', '0,2578 %').
    """
    dec = "." if is_en else ","
    if leak_pct is None:
        return f"< 0{dec}0001%"
    s_val = str(leak_pct).strip()
    if "<" in s_val:
        return "< 0.0001%" if is_en else "< 0,0001%"
    try:
        num = float(s_val.replace("%", "").replace(",", ".").strip())
        if num < 0.0001:
            return "< 0.0001%" if is_en else "< 0,0001%"
        return f"{num:.4f}".replace(".", dec) + " %"
    except Exception:
        if s_val in ("0", "0.0", "0,0", "0.00", "0,00", "0.000", "0,000", "0.0000", "0,0000", "0%", "0.0%", "0,0%", "0.0000%", "0,0000%"):
            return "< 0.0001%" if is_en else "< 0,0001%"
        return f"{s_val.replace('.', dec)} %"


def calc_hepa_leak(downflow_pct: float, exhaust_pct: float, aerosol_conc: float = 15.0) -> Dict[str, Any]:
    """Evaluates HEPA filter leakage (limit <= 0.01%)."""
    df_ok = downflow_pct <= 0.01
    ex_ok = exhaust_pct <= 0.01
    is_ok = df_ok and ex_ok
    return {
        "aerosol_conc": aerosol_conc,
        "downflow_pct": downflow_pct,
        "exhaust_pct": exhaust_pct,
        "downflow_formatted_vn": format_hepa_leak_val(downflow_pct, is_en=False),
        "downflow_formatted_en": format_hepa_leak_val(downflow_pct, is_en=True),
        "exhaust_formatted_vn": format_hepa_leak_val(exhaust_pct, is_en=False),
        "exhaust_formatted_en": format_hepa_leak_val(exhaust_pct, is_en=True),
        "downflow_assessment": "Đạt" if df_ok else "Không đạt",
        "exhaust_assessment": "Đạt" if ex_ok else "Không đạt",
        "assessment": "Đạt" if is_ok else "Không đạt",
        "assessment_en": "Pass" if is_ok else "Fail"
    }


def calc_uvc(readings: List[float]) -> Dict[str, Any]:
    """Calculates average UVC radiation intensity."""
    if not readings:
        return {"avg": 0.0, "assessment": "-"}
    avg_uvc = sum(readings) / len(readings)
    return {
        "count": len(readings),
        "avg": round(avg_uvc, 1),
        "avg_rounded": int(round(avg_uvc)),
        "assessment": "-",
        "assessment_en": "-"
    }



def _first_not_none(data: Dict[str, Any], *keys: str) -> Any:
    """Returns the first present, non-None value (keeps legitimate 0 values, unlike a chain of `or`)."""
    for k in keys:
        v = data.get(k)
        if v is not None:
            return v
    return None


def verify_bsc_calculations(dossier_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Performs independent recalculation from raw measurements and compares against
    scan/reference reported values. Returns a list of mismatch warnings if any discrepancy is found.
    Rule: Formulas in the original Excel file calculate independently; scan values are reference only.
    """
    mismatches = []

    # 1. Inflow Verification
    inflow_data = dossier_data.get("inflow", {})
    inflow_method = inflow_data.get("method")
    if inflow_method is None:
        inflow_method = 2 if inflow_data.get("readings") else 1

    calc_inflow_val = None
    if inflow_method == 2:
        m2_calc = calc_inflow_kanomax(inflow_data.get("readings", []), inflow_data.get("k_factor"))
        calc_inflow_val = m2_calc.get("velocity")
    else:
        q_readings = inflow_data.get("q_readings", [])
        q_val = sum(q_readings) / len(q_readings) if q_readings else inflow_data.get("q_m3h", 0)
        w_mm = inflow_data.get("width_mm", 0)
        h_mm = inflow_data.get("height_mm", 0)
        m1_calc = calc_inflow_dim(q_val, w_mm, h_mm)
        calc_inflow_val = m1_calc.get("velocity")

    scan_inflow = _first_not_none(inflow_data, "scan_velocity", "scan_val", "reference_velocity", "velocity")
    if calc_inflow_val is not None and scan_inflow is not None:
        try:
            scan_f = float(scan_inflow)
            diff = round(calc_inflow_val - scan_f, 2)
            if abs(diff) > 0.01:
                mismatches.append({
                    "parameter": "Tốc độ dòng khí qua cửa làm việc (Inflow)",
                    "parameter_en": "Inflow air velocity",
                    "calculated": calc_inflow_val,
                    "scan_reference": scan_f,
                    "diff": diff,
                    "unit": "m/s",
                    "message": f"Inflow: Giá trị tính toán độc lập từ công thức = {calc_inflow_val:.2f} m/s, nhưng số ghi trên scan = {scan_f:.2f} m/s (chênh lệch: {diff:+.2f} m/s)!"
                })
        except Exception as e:
            print(f"[Warning] Inflow mismatch check skipped (bad data): {e}")

    # 2. Downflow Verification
    df_data = dossier_data.get("downflow", {})
    df_readings = df_data.get("readings", [])
    if df_readings:
        df_calc = calc_downflow(df_readings, df_data.get("tolerance_pct", DOWNFLOW_TOLERANCE_PCT),
                                 bool(df_data.get("eval_uniformity", False)))
        calc_df_val = df_calc.get("avg")
        scan_df = _first_not_none(df_data, "scan_avg", "scan_val", "reference_val", "avg")
        if scan_df is not None:
            try:
                scan_f = float(scan_df)
                diff = round(calc_df_val - scan_f, 2)
                if abs(diff) > 0.01:
                    mismatches.append({
                        "parameter": "Tốc độ dòng khí cấp qua lọc HEPA (Downflow)",
                        "parameter_en": "Downflow air velocity",
                        "calculated": calc_df_val,
                        "scan_reference": scan_f,
                        "diff": diff,
                        "unit": "m/s",
                        "message": f"Downflow: Giá trị tính toán độc lập từ công thức = {calc_df_val:.2f} m/s, nhưng số ghi trên scan = {scan_f:.2f} m/s (chênh lệch: {diff:+.2f} m/s)!"
                    })
            except Exception as _err:
                print(f"[Warning] bsc_calc.py:420 {type(_err).__name__}: {_err}")

    # 3. Lux Verification
    lux_data = dossier_data.get("lux", {})
    lux_readings = lux_data.get("readings", [])
    if lux_readings:
        lux_calc = calc_lux(lux_readings)
        calc_lux_val = lux_calc.get("avg_rounded")
        scan_lux = _first_not_none(lux_data, "scan_avg", "scan_val", "reference_val", "avg_rounded", "avg")
        if scan_lux is not None:
            try:
                scan_i = int(round(float(scan_lux)))
                diff = calc_lux_val - scan_i
                if abs(diff) > 1:
                    mismatches.append({
                        "parameter": "Cường độ sáng (Lux)",
                        "parameter_en": "Illuminance (Lux)",
                        "calculated": calc_lux_val,
                        "scan_reference": scan_i,
                        "diff": diff,
                        "unit": "lux",
                        "message": f"Lux: Giá trị tính toán độc lập từ công thức = {calc_lux_val} lux, nhưng số ghi trên scan = {scan_i} lux (chênh lệch: {diff:+d} lux)!"
                    })
            except Exception as _err:
                print(f"[Warning] bsc_calc.py:444 {type(_err).__name__}: {_err}")

    # 4. Noise Verification
    noise_data = dossier_data.get("noise", {})
    if "ambient" in noise_data and "operating" in noise_data:
        noise_calc = calc_noise(float(noise_data["ambient"]), float(noise_data["operating"]))
        calc_corr = noise_calc.get("corrected_dba")
        scan_corr = _first_not_none(noise_data, "scan_corrected_dba", "scan_val", "corrected_dba")
        if scan_corr is not None:
            try:
                scan_f = float(scan_corr)
                diff = round(calc_corr - scan_f, 1)
                if abs(diff) > 0.1:
                    mismatches.append({
                        "parameter": "Độ ồn hiệu chỉnh (Noise level)",
                        "parameter_en": "Corrected noise level",
                        "calculated": calc_corr,
                        "scan_reference": scan_f,
                        "diff": diff,
                        "unit": "dBA",
                        "message": f"Noise: Giá trị hiệu chỉnh tính toán độc lập = {calc_corr:.1f} dBA, nhưng số ghi trên scan = {scan_f:.1f} dBA (chênh lệch: {diff:+.1f} dBA)!"
                    })
            except Exception as _err:
                print(f"[Warning] bsc_calc.py:467 {type(_err).__name__}: {_err}")

    # 5. UVC Verification
    uvc_data = dossier_data.get("uvc")
    if uvc_data and uvc_data.get("readings"):
        uvc_calc = calc_uvc(uvc_data["readings"])
        calc_uvc_val = uvc_calc.get("avg_rounded")
        scan_uvc = _first_not_none(uvc_data, "scan_avg", "scan_val", "reference_val", "avg_rounded")
        if scan_uvc is not None:
            try:
                scan_i = int(round(float(scan_uvc)))
                diff = calc_uvc_val - scan_i
                if abs(diff) > 1:
                    mismatches.append({
                        "parameter": "Cường độ bức xạ UVC",
                        "parameter_en": "UVC radiation intensity",
                        "calculated": calc_uvc_val,
                        "scan_reference": scan_i,
                        "diff": diff,
                        "unit": "µW/cm²",
                        "message": f"UVC: Giá trị tính toán độc lập từ công thức = {calc_uvc_val} µW/cm², nhưng số ghi trên scan = {scan_i} µW/cm² (chênh lệch: {diff:+d} µW/cm²)!"
                    })
            except Exception as _err:
                print(f"[Warning] bsc_calc.py:490 {type(_err).__name__}: {_err}")

    # 6. Assessment (Pass/Fail) Verification - AGENTS.md section 11
    #    Optional keys: "scan_assessment" / "reference_assessment" (what the field sheet concluded).
    for key, label_vn, label_en, calc_assessment in _assessment_checks(dossier_data):
        scan_assessment = (dossier_data.get(key, {}) or {}).get("scan_assessment") \
            or (dossier_data.get(key, {}) or {}).get("reference_assessment")
        if scan_assessment is None or calc_assessment in (None, "", "N/A", "-"):
            continue
        if _norm_assessment(scan_assessment) != _norm_assessment(calc_assessment):
            mismatches.append({
                "parameter": label_vn,
                "parameter_en": label_en,
                "calculated": calc_assessment,
                "scan_reference": scan_assessment,
                "diff": None,
                "unit": "",
                "severity": "CRITICAL",
                "kind": "assessment",
                "message": f"{label_vn}: kết luận tính toán độc lập = '{calc_assessment}', nhưng kết luận trên scan = '{scan_assessment}'!"
            })

    return mismatches


def _norm_assessment(v: Any) -> str:
    """Normalizes Pass/Fail wording (Đạt / Không đạt / Pass / Fail) for comparison."""
    t = str(v).strip().lower()
    if t in ("đạt", "dat", "pass", "passed", "complied", "phù hợp", "phu hop"):
        return "pass"
    if t in ("không đạt", "khong dat", "fail", "failed", "did not comply", "không phù hợp", "khong phu hop"):
        return "fail"
    return t


def _assessment_checks(dossier_data: Dict[str, Any]):
    """Yields (section_key, label_vn, label_en, independent_assessment) for sections that report Pass/Fail."""
    out = []
    df = dossier_data.get("downflow", {})
    if df.get("readings"):
        out.append(("downflow", "Đánh giá tốc độ dòng khí cấp", "Downflow assessment",
                    calc_downflow(df["readings"], df.get("tolerance_pct", DOWNFLOW_TOLERANCE_PCT),
                                  bool(df.get("eval_uniformity", False))).get("assessment")))
    lx = dossier_data.get("lux", {})
    if lx.get("readings"):
        out.append(("lux", "Đánh giá cường độ sáng", "Illuminance assessment",
                    calc_lux(lx["readings"]).get("assessment")))
    return out


def print_mismatch_warnings(mismatches: List[Dict[str, Any]]):
    """Prints highlighted warning box if any discrepancies were detected (AGENTS.md section 11 tags)."""
    if not mismatches:
        return
    print("\n" + "!" * 78)
    print(" [CẢNH BÁO / WARNING] PHÁT HIỆN SỐ LIỆU TÍNH TOÁN LỆCH VỚI GHI CHÉP HIỆN TRƯỜNG!")
    print(" Giá trị trên scan chỉ dùng để tham khảo. Đã giữ nguyên công thức tính trong Excel.")
    print("-" * 78)
    for idx, m in enumerate(mismatches, 1):
        tag = "[CRITICAL WARNING] [Assessment Mismatch]" if m.get("kind") == "assessment" else "[WARNING] [Data Mismatch]"
        print(f" {idx}. {tag} {m['message']}")
    print("!" * 78 + "\n")

def _gather_and_optimize_bsc_photos(dossier_data: Dict[str, Any], output_excel_path: str, max_photos: int = 4) -> List[str]:
    """
    Gathers field photos from dossier_data or dossier directories (photo/ or Photo/),
    corrects EXIF orientation, resizes them cleanly, and returns list of paths to optimized photos.
    """
    raw_photos = []
    # 1. From dossier_data["photos"]
    d_photos = dossier_data.get("photos")
    if d_photos and isinstance(d_photos, list):
        for p in d_photos:
            p_path = p.get("path") if isinstance(p, dict) else p
            if p_path and os.path.isfile(p_path):
                raw_photos.append(p_path)

    # 2. From input_dir
    input_dir = dossier_data.get("input_dir")
    candidate_dirs = []
    if input_dir and os.path.isdir(input_dir):
        candidate_dirs.append(input_dir)

    # 3. Fallback: relative to output_excel_path
    if output_excel_path:
        parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
        if os.path.basename(parent_dir).lower() in ("thuchien", "tbkq"):
            dossier_root = os.path.dirname(parent_dir)
        else:
            dossier_root = parent_dir
        if dossier_root not in candidate_dirs and os.path.isdir(dossier_root):
            candidate_dirs.append(dossier_root)

    img_exts = ('.jpg', '.jpeg', '.png')
    if not raw_photos:
        for cdir in candidate_dirs:
            for sub in ["photo", "Photo", "Photo Thuc Hien", "photos"]:
                p_sub = os.path.join(cdir, sub)
                if os.path.isdir(p_sub):
                    found = [
                        os.path.join(p_sub, f) for f in sorted(os.listdir(p_sub))
                        if f.lower().endswith(img_exts) and not f.startswith(("thumb_", "opt_"))
                    ]
                    if found:
                        raw_photos.extend(found)
                        break
            if raw_photos:
                break

    if not raw_photos or not HAS_PIL_IMAGE:
        return raw_photos[:max_photos]

    # Optimize and fix EXIF orientation
    optimized_photos = []
    for p_path in raw_photos[:max_photos]:
        try:
            with PILImage.open(p_path) as im:
                im = ImageOps.exif_transpose(im)
                w, h = im.size
                max_w = 1600
                if w > max_w:
                    new_h = int(h * (max_w / w))
                    im = im.resize((max_w, new_h), PILImage.Resampling.LANCZOS)
                c_dir = os.path.join(os.path.dirname(p_path), "_optimized")
                os.makedirs(c_dir, exist_ok=True)
                out_p = os.path.join(c_dir, f"opt_{os.path.splitext(os.path.basename(p_path))[0]}.jpg")
                im.convert("RGB").save(out_p, format="JPEG", quality=90)
                optimized_photos.append(out_p)
        except Exception as e:
            print(f"[bsc_calc] Photo optimization fallback for {p_path}: {e}")
            optimized_photos.append(p_path)

    return optimized_photos


def build_bsc_calculation_workbook(dossier_data: Dict[str, Any], output_excel_path: str,
                                   cals: Optional[Dict[str, Any]] = None, finalize: Optional[bool] = None):
    """
    Builds a standardized QUATEST 3 BSC calculation workbook (.xlsx) by populating
    the official QUATEST 3 BSC template (Core/Templates/TinhToanBSC_template.xlsx).
    Preserves 100% of the original formatting, sheets ('Bảng tính', 'TBKQ-VN', 'TBKQ - ENG'),
    formulas, fonts, and borders.
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    workspace_dir = os.path.dirname(os.path.dirname(current_dir))
    template_candidates = [
        os.path.join(current_dir, "..", "Templates", "TinhToanBSC_template.xlsx"),
        os.path.join(workspace_dir, "Core", "Templates", "TinhToanBSC_template.xlsx"),
        os.path.join(workspace_dir, "New From 2026", "BSC, Fumehood", "BSC REF", "Q326D059003 - 1.xlsx")
    ]
    template_path = None
    for cand in template_candidates:
        if os.path.exists(cand):
            template_path = cand
            break

    if not template_path:
        raise FileNotFoundError("Could not find standard QUATEST 3 BSC Excel template!")

    wb = openpyxl.load_workbook(template_path)
    fill_workbook_tags(wb, dossier_data)  # ô thông tin hồ sơ trong template là TAG (Core/Templates/TAGS.md)
    ws = wb["Bảng tính"] if "Bảng tính" in wb.sheetnames else wb.active

    # Extract dossier parameters
    meta = dossier_data.get("metadata", {})
    code = dossier_data.get("code") or meta.get("code", "")
    sub_no = dossier_data.get("sub_no") or meta.get("sub_no", "")
    reg_date = dossier_data.get("reg_date") or meta.get("reg_date", "")
    sample_date = dossier_data.get("sample_date") or meta.get("sample_date", "")
    start_time = dossier_data.get("start_time") or meta.get("start_time", "11:00")
    end_time = dossier_data.get("end_time") or meta.get("end_time", "16:00")

    cab = dossier_data.get("cabinet", {})
    cab_name = cab.get("cabinet_name") or dossier_data.get("cabinet_name", "Tủ cấy BIOCYT")
    brand = cab.get("brand") or dossier_data.get("brand", "")
    model = cab.get("model") or dossier_data.get("model", "")
    serial = cab.get("serial") or dossier_data.get("serial", "")
    dev_id = cab.get("device_id") or dossier_data.get("device_id", "VS0012")

    # 1. Header Information (Rows 1 - 5)
    ws["C1"] = code
    ws["G1"] = sub_no if sub_no else ""
    ws["I1"] = reg_date
    ws["C2"] = sample_date
    ws["F2"] = start_time
    ws["I2"] = end_time

    ws["B3"] = cab_name
    ws["G3"] = brand
    ws["B4"] = "II"
    ws["G4"] = model
    ws["B5"] = serial
    ws["H5"] = dev_id
    ws["K1"] = None

    # Update Page Setup Headers for all sheets
    cert_no_str = f"{code}/{sub_no}" if sub_no else code
    cert_date_str = sample_date or dossier_data.get("cert_date") or reg_date
    for sname in wb.sheetnames:
        s = wb[sname]
        s.oddHeader.left.text = f"Số: {cert_no_str}"
        s.oddHeader.right.text = f"Ngày {cert_date_str}" 

    # 2. Inflow Section (Rows 9 - 31)
    inflow_data = dossier_data.get("inflow", {})
    inflow_method = inflow_data.get("method")
    if inflow_method is None:
        if inflow_data.get("readings"):
            inflow_method = 2
        else:
            inflow_method = 1

    no_border = Border()
    thin_black = Side(style='thin', color='000000')
    from openpyxl.cell.cell import MergedCell
    font_12 = Font(name="Times New Roman", size=12)
    align_center = Alignment(horizontal="center", vertical="center")

    if inflow_method == 2:
        # Method 1 (Rows 10 - 15) - Leave as dash, NO numbers filled
        ws["A10"] = "Phương pháp 1:  DIM - TSI 8380"   # mặc định DIM = TSI 8380 (người dùng chốt 11/11/2026)
        ws["A12"] = "-"
        ws["A14"] = "-"
        ws["A15"] = "-"
        for col_letter in ["B", "C", "D", "E", "F"]:
            ws[f"{col_letter}12"] = "-"
        ws["G12"] = "-"
        ws["J12"] = "-"

        # Method 2 (Rows 17 - 26)
        ws["A17"] = "Phương pháp 2: Kanomax Climomaster 6501.       "
        ws["A18"] = "Vị trí đo"
        ws["A19"] = inflow_data.get("location", "Mép trên cửa kính")

        # Determine actual number of measurement columns
        m2_readings = inflow_data.get("readings", [])
        if m2_readings and isinstance(m2_readings[0], list):
            row1_vals = m2_readings[0]
            row2_vals = m2_readings[1] if len(m2_readings) > 1 else []
            max_c = max(len(row1_vals), len(row2_vals))
        elif m2_readings:
            max_c = len(m2_readings)
        else:
            max_c = 8
        end_inflow_idx = 2 + max_c
        end_inflow_col = openpyxl.utils.get_column_letter(end_inflow_idx)

        # Clear existing borders and values in Method 2 range up to col 13 (M)
        for r in range(21, 26):
            for c in range(3, 14):
                ws.cell(r, c).value = None
                ws.cell(r, c).border = no_border

        # Apply outer boundary box only according to actual number of measurement columns (C21 to end_inflow_col25)
        for r in range(21, 26):
            for c in range(3, end_inflow_idx + 1):
                top = thin_black if r == 21 else None
                bottom = thin_black if r == 25 else None
                left = thin_black if c == 3 else None
                right = thin_black if c == end_inflow_idx else None
                if top or bottom or left or right:
                    ws.cell(r, c).border = Border(top=top, bottom=bottom, left=left, right=right)

        if m2_readings and isinstance(m2_readings[0], list):
            # 2D grid: row 1 at row 22, row 2 at row 24
            row1_vals = m2_readings[0]
            row2_vals = m2_readings[1] if len(m2_readings) > 1 else []
            for i, v in enumerate(row1_vals):
                c_idx = 3 + i
                cell = ws.cell(22, c_idx)
                cell.value = float(v)
                cell.font = font_12
                cell.alignment = align_center
                cell.number_format = "0.00"
            for i, v in enumerate(row2_vals):
                c_idx = 3 + i
                cell = ws.cell(24, c_idx)
                cell.value = float(v)
                cell.font = font_12
                cell.alignment = align_center
                cell.number_format = "0.00"

            ws["K22"] = "TB (m/s)"
            ws["L22"] = f"=AVERAGE(C22:{end_inflow_col}22,C24:{end_inflow_col}24)"
        elif m2_readings:
            # 1D list
            for i, v in enumerate(m2_readings):
                cell = ws.cell(23, 3 + i)
                cell.value = float(v)
                cell.font = font_12
                cell.alignment = align_center
                cell.number_format = "0.00"
            ws["K22"] = "TB (m/s)"
            ws["L22"] = f"=AVERAGE(C23:{end_inflow_col}23)"
        else:
            ws["K22"] = "TB (m/s)"
            ws["L22"] = f"=AVERAGE(C23:{end_inflow_col}23)"

        k_factor = inflow_data.get("k_factor")
        if k_factor:
            ws["K23"] = "k"
            ws["L23"] = float(k_factor)
            ws["K24"] = "Vtb"
            ws["L24"] = "=L22*L23"
            ws["D31"] = "=L24"
        else:
            ws["K23"] = None
            ws["L23"] = None
            ws["K24"] = None
            ws["L24"] = None
            ws["D31"] = "=L22"

        ws["F31"] = 0.4
        ws["H31"] = "KQD"
        ws["J31"] = '=IF(D31="N/A","N/A",IF(AND(IF(F31="KQD",TRUE,D31>=F31),IF(H31="KQD",TRUE,D31<=H31)), "Đạt", "Không đạt"))'

    else:
        # Method 1 (DIM TSI)
        w_mm = inflow_data.get("width_mm", 1800)
        h_mm = inflow_data.get("height_mm", 180)
        q_readings = inflow_data.get("q_readings", [932, 936, 932, 936, 936])

        ws["A10"] = "Phương pháp 1:  DIM - TSI 8380"
        ws["A14"] = h_mm
        ws["A15"] = w_mm
        ws["A12"] = "=A14*A15/1000000"

        for i, v in enumerate(q_readings[:5]):
            col_letter = openpyxl.utils.get_column_letter(2 + i)  # B to F
            ws[f"{col_letter}12"] = v

        ws["G12"] = "=AVERAGE(B12:F12)"
        ws["J12"] = "=G12/A12/3600"
        ws["D31"] = "=J12"
        ws["F31"] = 0.4
        ws["H31"] = "KQD"
        ws["J31"] = '=IF(D31="N/A","N/A",IF(AND(IF(F31="KQD",TRUE,D31>=F31),IF(H31="KQD",TRUE,D31<=H31)), "Đạt", "Không đạt"))'

        # Method 2 (rows 17 - 26) not used: keep template structure & black frame C21:J25 without numbers (with diagonal slash)
        ws["A17"] = "Phương pháp 2: Kanomax Climomaster 6501.       "
        ws["A18"] = "Vị trí đo"
        ws["A19"] = "Mép trên cửa kính"

        # Clear numbers in rows 21 to 25 and maintain the outer black box border C21:J25 with diagonal slash
        for r in range(21, 26):
            for c in range(3, 14):
                cell = ws.cell(r, c)
                if not isinstance(cell, MergedCell):
                    cell.value = None
                cell.border = no_border
        strike_box(ws, "C21:J25")

        ws["K22"] = "TB (m/s)"
        ws["L22"] = "-"
        ws["K23"] = "k"
        ws["L23"] = "-"
        ws["K24"] = "Vtb"
        ws["L24"] = "-"

    # 3. Downflow Section (Rows 33 - 49)
    downflow_data = dossier_data.get("downflow", {})

    # Clear existing template values and borders in Downflow range up to col 13 (M)
    for r in range(37, 42):
        for c in range(3, 14):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell.value = None
            cell.border = no_border

    if downflow_data.get("skip"):
        strike_box(ws, "C37:I41")
        for ref in ("D46", "F46", "H46", "J46", "D48", "D49", "H48", "H49", "J48"):
            ws[ref] = "-"
    else:
        df_readings = downflow_data.get("readings", [])
        if not df_readings:
            df_readings = [
                [0.57, 0.53, 0.48, 0.45, 0.42, 0.42, 0.43, 0.47, 0.47],
                [0.56, 0.52, 0.49, 0.44, 0.42, 0.43, 0.40, 0.49, 0.48],
                [0.54, 0.52, 0.46, 0.44, 0.44, 0.44, 0.42, 0.48, 0.46]
            ]

        font_12 = Font(name="Times New Roman", size=12)
        align_center = Alignment(horizontal="center", vertical="center")

        df_rows = [37, 39, 41]
        max_cols = 0
        for r_idx, row_vals in zip(df_rows, df_readings):
            if len(row_vals) > max_cols:
                max_cols = len(row_vals)
            for c_idx, val in enumerate(row_vals):
                col_letter = openpyxl.utils.get_column_letter(3 + c_idx)  # Starts at C
                cell = ws[f"{col_letter}{r_idx}"]
                cell.value = val
                cell.font = font_12
                cell.alignment = align_center
                cell.number_format = "0.00"

        end_col_idx = 2 + max_cols  # e.g. 7 (G) for 5 cols, 11 (K) for 9 cols
        end_col_letter = openpyxl.utils.get_column_letter(end_col_idx)

        # Apply outer boundary box only (no internal borders)
        for r in range(37, 42):
            for c in range(3, end_col_idx + 1):
                top = thin_black if r == 37 else None
                bottom = thin_black if r == 41 else None
                left = thin_black if c == 3 else None
                right = thin_black if c == end_col_idx else None
                if top or bottom or left or right:
                    ws.cell(r, c).border = Border(top=top, bottom=bottom, left=left, right=right)

        ws["D46"] = f"=AVERAGE(C37:{end_col_letter}41)"
        ws["D48"] = f"=MAX(C37:{end_col_letter}41)"
        ws["D49"] = f"=MIN(C37:{end_col_letter}41)"
        ws["E43"] = downflow_data.get("tolerance_pct", DOWNFLOW_TOLERANCE_PCT)
        # Đánh giá bằng công thức Excel (AGENTS mục 11), biên 0,25 / 0,50 tính cả hai đầu
        ws["J46"] = ('=IF(D46="N/A","N/A",IF(AND(IF(F46="KQD",TRUE,D46>=F46),IF(H46="KQD",TRUE,D46<=H46)),'
                     '"Đạt","Không đạt"))')
        # Độ đồng đều: luôn tính sẵn H48/H49 (D46 ± E43 %), chỉ đánh giá khi eval_uniformity=True; mặc định "-" (không xét)
        if downflow_data.get("eval_uniformity"):
            ws["J48"] = '=IF(AND(D48<=H48,D49>=H49),"Đạt","Không đạt")'
        else:
            ws["J48"] = "-"

    # 4. Lux Section (Rows 50 - 64)
    lux_data = dossier_data.get("lux", {})
    lux_readings = lux_data.get("readings", [])
    if not lux_readings:
        lux_readings = [
            [956.7, 1122.7, 1160.9, 1084.6, 951.5],
            [956.1, 1104.7, 1102.6, 1091.0, 878.8],
            [937.4, 849.6, 1030.7, 990.1, 867.8]
        ]

    lux_row_indices = [53, 55, 57]
    lux_cols = ["C", "E", "G", "I", "K", "M"]

    # Clear all existing template lux cells
    for r_idx in lux_row_indices:
        for c_name in lux_cols:
            ws[f"{c_name}{r_idx}"].value = None

    for r_idx, row_vals in zip(lux_row_indices, lux_readings):
        for c_idx, val in enumerate(row_vals):
            if c_idx < len(lux_cols):
                col_name = lux_cols[c_idx]
                cell = ws[f"{col_name}{r_idx}"]
                cell.value = val
                cell.font = font_12
                cell.alignment = align_center
                cell.number_format = "0.0"

    end_lux_col = lux_cols[len(lux_readings[0]) - 1] if lux_readings else "K"
    end_lux_idx = openpyxl.utils.column_index_from_string(end_lux_col)

    # Update note formula to count active columns
    ws['A51'] = None    # phiếu chỉ in bảng số liệu — bỏ ghi chú bố trí điểm đo (AGENTS mục 11)
    ws['A34'] = None

    # Clear existing borders in Lux range up to col M
    for r in range(53, 58):
        for c in range(3, 14):
            ws.cell(r, c).border = no_border

    # Apply outer boundary box only (no internal borders)
    for r in range(53, 58):
        for c in range(3, end_lux_idx + 1):
            top = thin_black if r == 53 else None
            bottom = thin_black if r == 57 else None
            left = thin_black if c == 3 else None
            right = thin_black if c == end_lux_idx else None
            if top or bottom or left or right:
                ws.cell(r, c).border = Border(top=top, bottom=bottom, left=left, right=right)

    ws["D62"] = f"=AVERAGE(C53:{end_lux_col}57)"
    ws["J62"] = '=IF(D62="N/A","N/A",IF(D62>=F62,"Đạt","Không đạt"))'

    # 5. Noise Section (Rows 66 - 78)
    noise_data = dossier_data.get("noise", {})
    amb_db = float(noise_data.get("ambient", 44.7))
    op_db = float(noise_data.get("operating", 65.9))

    # Hiệu chỉnh ồn nền theo bảng quy trình (Core/Tools/noise_corr.py)
    noise_correction_value(amb_db, op_db)            # chỉ để in cảnh báo ΔN ≤ 2 dB nếu có

    ws["I68"] = amb_db
    ws["I69"] = op_db
    ws["I70"] = noise_correction_formula("I68", "I69")   # công thức tra bảng, không ghi số tĩnh (07/11/2026)
    ws["I72"] = "=I69-I70"
    ws["D76"] = "=I72"
    ws["F76"] = "KQD"
    ws["H76"] = 65
    ws["J76"] = '=IF(D76="N/A","N/A",IF(AND(IF(F76="KQD",TRUE,D76>=F76),IF(H76="KQD",TRUE,D76<=H76)), "Đạt", "Không đạt"))'

    # 6. HEPA Leak Section (Rows 80 - 95)
    hepa_data = dossier_data.get("hepa", {})
    df_leak = hepa_data.get("downflow_pct", 0.2578)
    ex_leak = hepa_data.get("exhaust_pct", 0.0021)
    aero_conc = hepa_data.get("aerosol_conc", 12)

    ws["J82"] = f"Nồng độ phun: {aero_conc} µg/L (tính theo hướng dẫn của NSX thiết bị tại manual của máy ATI Generator: 3375 * N (số đầu phun)/ tổng lưu lượng gió (CFM)"
    ws["D83"] = df_leak
    ws["H83"] = '=IF(D83<=0.01,"Đạt","Không đạt")'
    ws["D87"] = ex_leak
    ws["H87"] = '=IF(D87<=0.01,"Đạt","Không đạt")'

    # 7. UVC Section (Rows 112 - 123)
    uvc_data = dossier_data.get("uvc")
    # Clear existing borders and values in UVC range up to col 10 (J)
    for r in range(114, 119):
        for c in range(3, 11):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell.value = None
            cell.border = no_border

    if uvc_data and uvc_data.get("readings"):
        u_readings = uvc_data["readings"]
        for i, v in enumerate(u_readings):
            c_idx = 4 + i  # D116, E116, F116, G116...
            cell = ws.cell(116, c_idx)
            cell.value = float(v)
            cell.font = font_12
            cell.alignment = align_center
            cell.number_format = "0"

        # Actual result columns: readings start at col D (4), end at actual result column (e.g. col G for 4 readings)
        end_uvc_idx = 3 + len(u_readings)
        end_uvc_col = openpyxl.utils.get_column_letter(end_uvc_idx)

        # Apply outer boundary box only according to actual measurement points (D114 to end_uvc_col118)
        for r in range(114, 119):
            for c in range(4, end_uvc_idx + 1):
                top = thin_black if r == 114 else None
                bottom = thin_black if r == 118 else None
                left = thin_black if c == 4 else None
                right = thin_black if c == end_uvc_idx else None
                if top or bottom or left or right:
                    ws.cell(r, c).border = Border(top=top, bottom=bottom, left=left, right=right)

        ws["D122"] = f"=AVERAGE(D114:{end_uvc_col}118)"
        ws["J122"] = "-"   # UVC không có mức quy định -> không đánh giá (người dùng chốt 30/10/2026)
    else:
        # Không đo UVC: giữ khung lưới đo C114:I118 theo form gốc, gạch chéo (người dùng chốt 30/10/2026)
        strike_box(ws, "C114:I118")
        ws["D122"] = "-"
        ws["J122"] = "-"

    # 8. Particle Section (Rows 125 - 136)
    part_data = dossier_data.get("particle")
    has_particle = False
    p_readings = None
    if part_data:
        p_readings = part_data.get("readings")
        if p_readings is None and part_data.get("sizes"):
            p_readings = []
            for s_info in part_data["sizes"]:
                v_str = str(s_info[2]).split()[0]
                try:
                    p_readings.append(int(v_str))
                except ValueError:
                    try:
                        p_readings.append(float(v_str.replace(",", ".")))
                    except ValueError:
                        p_readings.append(v_str)
        if p_readings is not None:
            has_particle = True

    if has_particle and p_readings is not None:
        ws["A125"] = "8. Đo nồng độ hạt bụi trong tủ"
        cols = ["D", "F", "H", "J"]
        for col_name, v in zip(cols, p_readings):
            ws[f"{col_name}130"] = v
        ws["D136"] = '=IF(AND(D130<=D133, F130<=F133, H130<=H133, J130<=J133), "Đạt", "Không đạt")'
        if part_data.get("file_zero"):
            ws["C126"] = part_data["file_zero"]
        if part_data.get("file_ghi"):
            ws["C127"] = part_data["file_ghi"]
    else:
        ws["A125"] = "8. Đo nồng độ hạt bụi trong tủ"
        ws["C126"] = "-"
        ws["C127"] = "-"
        for col_name in ["D", "F", "H", "J"]:
            ws[f"{col_name}130"] = "-"
        ws["D136"] = "-"

    # 9. Insert Photos into Worksheets
    photos = _gather_and_optimize_bsc_photos(dossier_data, output_excel_path, max_photos=4)

    # 9.1. Sheet 'Bảng tính' (Row 144: Hình ảnh quá trình thực hiện)
    bt_anchors = ["A145", "G145", "A163", "G163"]
    for p_cell in bt_anchors:
        ws[p_cell] = None

    if photos and HAS_PIL_IMAGE:
        for idx, p_path in enumerate(photos[:4]):
            anchor = bt_anchors[idx]
            try:
                with PILImage.open(p_path) as pim:
                    w, h = pim.size
                    ratio = w / h if h > 0 else 1.0
                    target_w = 330
                    target_h = int(target_w / ratio)
                    if target_h > 247:
                        target_h = 247
                        target_w = int(target_h * ratio)
                img = OpenpyxlImage(p_path)
                img.width = target_w
                img.height = target_h
                ws.add_image(img, anchor)
            except Exception as e:
                print(f"[bsc_calc] Failed to add photo to Bảng tính {anchor}: {e}")

    # 9.2. Sheet 'TBKQ-VN' (Row 126: 10.4. Hình ảnh quá trình thực hiện)
    if "TBKQ-VN" in wb.sheetnames:
        ws_vn = wb["TBKQ-VN"]
        vn_anchors = ["A127", "J127", "A142", "J142"]
        for c in vn_anchors:
            ws_vn[c] = None

        if photos and HAS_PIL_IMAGE:
            for idx, p_path in enumerate(photos[:4]):
                anchor = vn_anchors[idx]
                try:
                    with PILImage.open(p_path) as pim:
                        w, h = pim.size
                        ratio = w / h if h > 0 else 1.0
                        target_w = 276
                        target_h = int(target_w / ratio)
                        if target_h > 195:
                            target_h = 195
                            target_w = int(target_h * ratio)
                    img = OpenpyxlImage(p_path)
                    img.width = target_w
                    img.height = target_h
                    ws_vn.add_image(img, anchor)
                except Exception as e:
                    print(f"[bsc_calc] Failed to add photo to TBKQ-VN {anchor}: {e}")

    # 9.3. Sheet 'TBKQ - ENG' (Row 129: 10.4. Photos of performance process)
    if "TBKQ - ENG" in wb.sheetnames:
        ws_en = wb["TBKQ - ENG"]
        en_anchors = ["A130", "J130", "A143", "J143"]
        for c in en_anchors:
            ws_en[c] = None

        if photos and HAS_PIL_IMAGE:
            for idx, p_path in enumerate(photos[:4]):
                anchor = en_anchors[idx]
                try:
                    with PILImage.open(p_path) as pim:
                        w, h = pim.size
                        ratio = w / h if h > 0 else 1.0
                        target_w = 276
                        target_h = int(target_w / ratio)
                        if target_h > 215:
                            target_h = 215
                            target_w = int(target_h * ratio)
                    img = OpenpyxlImage(p_path)
                    img.width = target_w
                    img.height = target_h
                    ws_en.add_image(img, anchor)
                except Exception as e:
                    print(f"[bsc_calc] Failed to add photo to TBKQ - ENG {anchor}: {e}")

    # 9. Update TBKQ-VN and TBKQ - ENG sheets
    # Clean Vietnamese parameter names in Bảng tính
    ws["A9"] = f'{ws["M9"].value or 1}. Tốc độ trung bình của dòng khí qua cửa làm việc'
    ws["A33"] = f'{ws["M33"].value or 2}. Tốc độ trung bình của dòng khí cấp qua lọc HEPA'
    ws["A83"] = "Lọc khí cấp"
    ws["A87"] = "Lọc khí thải"

    if "TBKQ-VN" in wb.sheetnames:
        ws_vn = wb["TBKQ-VN"]
        # Ensure purely Vietnamese parameter names (remove English subtitles)
        ws_vn["B68"] = "Tốc độ trung bình dòng khí qua cửa làm việc"
        ws_vn["B71"] = "Tốc độ trung bình dòng khí cấp qua lọc HEPA"
        ws_vn["B74"] = "Cường độ sáng trung bình bên trong tủ"
        ws_vn["B77"] = "Đo độ ồn"
        ws_vn["B82"] = "Mức rò rỉ của các lọc HEPA"
        ws_vn["B83"] = "Lọc khí cấp"
        ws_vn["B86"] = "Lọc khí thải"
        ws_vn["B92"] = "Kiểm tra hình thái dòng khí"
        ws_vn["B109"] = "Cường độ tia UV dải C"
        ws_vn["B112"] = "Đo nồng độ hạt trong tủ"

        if dossier_data.get("client_name_vn"):
            ws_vn["T11"] = dossier_data["client_name_vn"]
        if dossier_data.get("client_address_vn"):
            ws_vn["T13"] = dossier_data["client_address_vn"]
        ws_vn["T15"] = "Giám định tủ an toàn sinh học cấp 2."
        ws_vn["U19"] = ""
        if sample_date:
            ws_vn["T33"] = sample_date

        eqs_vn = dossier_data.get("equipments", [])
        if eqs_vn:
            for idx, eq_name in enumerate(eqs_vn[:7]):
                eq_text = eq_name if eq_name.startswith("-") else f"- {eq_name};"
                ws_vn[f"B{23+idx}"] = eq_text

    if "TBKQ - ENG" in wb.sheetnames:
        ws_en = wb["TBKQ - ENG"]
        if dossier_data.get("client_name_en"):
            ws_en["T9"] = dossier_data["client_name_en"]
        if dossier_data.get("client_address_en"):
            ws_en["T11"] = dossier_data["client_address_en"]
        ws_en["T13"] = "Inspection of Biosafety Cabinet Class II."
        ws_en["U17"] = ""

        eqs_en = dossier_data.get("equipments_en", [])
        if eqs_en:
            for idx, eq_name in enumerate(eqs_en[:7]):
                eq_text = eq_name if eq_name.startswith("-") else f"- {eq_name};"
                ws_en[f"B{22+idx}"] = eq_text

    # Verify calculations independently and warn on mismatches
    mismatches = verify_bsc_calculations(dossier_data)
    if mismatches:
        print_mismatch_warnings(mismatches)

    # Viền ngoài thường, không dùng viền dày (người dùng chốt 30/10/2026)
    thin_outside_borders(wb)

    # Save to destination
    os.makedirs(os.path.dirname(os.path.abspath(output_excel_path)), exist_ok=True)
    wb.save(output_excel_path)
    print(f"[bsc_calc] Saved template-based calculation workbook -> {output_excel_path}")

    # Phần "sau hiệu chỉnh" theo giấy hiệu chuẩn (người dùng chốt 07/11/2026, Core/Tools/bsc_cal_blocks.py):
    # có `cals` thì chèn ngay (finalize mặc định True); script còn hậu xử lý ô theo địa chỉ form gốc thì truyền
    # finalize=False rồi tự gọi finalize_bsc_workbook(path, cals) SAU khi hậu xử lý.
    if finalize is None:
        finalize = bool(cals)
    if finalize:
        try:
            from .bsc_cal_blocks import finalize_bsc_workbook
            finalize_bsc_workbook(output_excel_path, cals)
            return output_excel_path
        except Exception as ex:
            print(f"[bsc_calc] Note: không thêm được phần sau hiệu chỉnh: {ex}")

    # Re-evaluate and persist calculated values via Excel COM if available
    try:
        import win32com.client
        import gc
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            abs_p = os.path.abspath(output_excel_path)
            com_wb = excel.Workbooks.Open(abs_p)
            excel.CalculateFull()
            com_wb.Save()
            com_wb.Close(SaveChanges=True)
            print(f"[bsc_calc] Recalculated formulas with Excel COM -> {abs_p}")
        finally:
            com_wb = None
            excel.Quit()
            del excel
            gc.collect()
    except Exception as ex:
        print(f"[bsc_calc] Note: Excel COM calculation skipped: {ex}")

    return output_excel_path


class _BscRows:
    """Dò dòng theo nhãn cột A của sheet 'Bảng tính' (E53, 07/11/2026): phiếu đã thêm phần "sau hiệu chỉnh"
    (bsc_cal_blocks.finalize_bsc_workbook chèn dòng) làm lệch địa chỉ form gốc. Không thấy nhãn -> dòng form gốc."""

    def __init__(self, ws_formula):
        self.labels = {}
        for r in range(1, (ws_formula.max_row or 0) + 1):
            v = ws_formula.cell(r, 1).value
            if isinstance(v, str) and v.strip():
                self.labels[r] = v

    def find(self, text: str, default: int, exact: bool = False, after: int = 0) -> int:
        t = text.strip().lower()
        for r in sorted(self.labels):
            if r <= after:
                continue
            v = self.labels[r].strip().lower()
            if (v == t) if exact else (t in v):
                return r
        return default


def parse_bsc_excel(excel_path: str) -> Dict[str, Any]:
    """
    Parses a standardized BSC calculation workbook (.xlsx) and extracts all
    metadata and test results into a structured dictionary ready for Dossier generation.
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb["Bảng tính"] if "Bảng tính" in wb.sheetnames else wb.active
    _wbf = openpyxl.load_workbook(excel_path, data_only=False)
    ws_f = _wbf["Bảng tính"] if "Bảng tính" in _wbf.sheetnames else _wbf.active
    R = _BscRows(ws_f)
    # Dòng theo nhãn (form gốc ở mặc định) — phiếu có phần "sau hiệu chỉnh": kết quả đọc từ ô kết quả (đã trỏ sang b)),
    # số đọc lấy ở lưới a) (giữ vị trí tương đối với tiêu đề mục)
    r_door = 14
    for r in range(1, (ws_f.max_row or 0) + 1):
        if str(ws_f.cell(r, 2).value or "").strip().startswith("Độ mở cửa"):
            r_door = r
            break
    r_pp2 = R.find("Phương pháp 2", 17)
    r_inres = R.find("- Vận tốc gió trung bình", 31)
    r_df_title = R.find("Tốc độ trung bình của dòng khí cấp", 33)
    r_df_tol = R.find("Độ không đồng đều", 43, after=r_df_title)
    r_df_res = R.find("Vận tốc gió trung bình (m/s)", 46, exact=True, after=r_df_title)
    r_df_uni = R.find("Độ đồng đều của vận tốc gió", 48, after=r_df_title)
    r_lux_title = R.find("Cường độ sáng trung bình bên trong", 50)
    r_lux_res = R.find("Cường độ sáng trung bình (lux)", 62, after=r_lux_title)
    r_noise = R.find("Khi tủ không vận hành", 68)
    r_noise_res = R.find("Mức cường độ âm", 76, after=r_noise)
    r_hepa_in = R.find("Lọc khí cấp", 83, exact=True)
    r_hepa_ex = R.find("Lọc khí thải", 87, exact=True)
    r_uvc_title = R.find("Đo cường độ UVC", 112)
    r_uvc_raw = r_uvc_title + (5 if str(ws_f.cell(r_uvc_title + 1, 1).value or "").startswith("a)") else 4)
    r_uvc_res = R.find("Cường độ UVC", 122, after=r_uvc_title + 1)
    r_part = R.find("Giá trị đo, hạt/m3", 130)

    # Header & Cabinet Metadata
    code = ws["C1"].value or ""
    sub_no = ws["G1"].value or ""
    reg_date = ws["I1"].value or ""
    sample_date = ws["C2"].value or ""
    start_time = ws["F2"].value or "09:30"
    end_time = ws["I2"].value or "11:30"

    cab_name = ws["B3"].value or "Tủ an toàn sinh học cấp 2"
    brand = ws["G3"].value or ""
    model = ws["G4"].value or ""
    serial = ws["B5"].value or ""
    # Mã thiết bị nằm ở ô H5 (gộp H5:J5; F5:G5 là nhãn "Mã thiết bị:") — cùng ô mà bộ ghi dùng (sửa 30/10/2026, trước đọc nhầm G5)
    dev_id = ws["H5"].value or ws["G5"].value or "N/A"

    # 1. Inflow
    w_mm = ws[f"A{r_door + 1}"].value or 930
    h_mm = ws[f"A{r_door}"].value or 200
    b12_val = ws["B12"].value
    k22, k23, k24 = r_pp2 + 5, r_pp2 + 6, r_pp2 + 7          # lưới Kanomax a) (form gốc 22–24)
    is_m2 = (b12_val in ("-", None, "")) and (ws[f"C{k22}"].value is not None or ws[f"C{k23}"].value is not None
                                              or ws[f"L{k22}"].value is not None)

    inflow_vel = ws[f"D{r_inres}"].value or ws[f"L{k24}"].value or ws[f"L{k22}"].value or ws["J12"].value or 0.40
    if isinstance(inflow_vel, (int, float)):
        inflow_vel = round(float(inflow_vel), 2)

    # 2. Downflow
    df_readings = []
    for r in range(r_df_title + 4, r_df_title + 10):
        row_vals = [ws.cell(r, c).value for c in range(3, 15) if ws.cell(r, c).value is not None and isinstance(ws.cell(r, c).value, (int, float))]
        if row_vals:
            df_readings.append(row_vals)
    df_avg = ws[f"D{r_df_res}"].value or (sum(v for row in df_readings for v in row) / sum(len(row) for row in df_readings) if df_readings else 0.28)
    if isinstance(df_avg, (int, float)):
        df_avg = round(float(df_avg), 2)
    df_tol = ws[f"E{r_df_tol}"].value or DOWNFLOW_TOLERANCE_PCT
    # J48 = "-" khi không xét độ đồng đều; là công thức (đọc data_only ra "Đạt"/"Không đạt") khi có xét
    _j48 = ws_f[f"J{r_df_uni}"].value
    df_eval_uniformity = isinstance(_j48, str) and _j48.strip() not in ("", "-")

    # 3. Lux
    lux_readings = []
    for r in range(r_lux_title + 3, r_lux_title + 9):
        row_vals = [ws.cell(r, c).value for c in range(3, 15) if ws.cell(r, c).value is not None and isinstance(ws.cell(r, c).value, (int, float))]
        if row_vals:
            lux_readings.append(row_vals)
    flat_lux = [v for row in lux_readings for v in row]
    lux_avg = ws[f"D{r_lux_res}"].value or (sum(flat_lux) / len(flat_lux) if flat_lux else 750)
    lux_rounded = int(round(float(lux_avg))) if isinstance(lux_avg, (int, float)) else 750

    # 4. Noise
    amb_db = ws[f"I{r_noise}"].value or 47.7
    op_db = ws[f"I{r_noise + 1}"].value or 60.7
    corr_db = ws[f"D{r_noise_res}"].value or ws[f"I{r_noise + 4}"].value
    if corr_db is None:   # chưa tính lại bằng Excel (vd chạy trên Linux): ồn vận hành − hiệu chỉnh I70
        try:
            from .noise_corr import noise_correction_value
            corr_db = round(float(op_db) - noise_correction_value(float(amb_db), float(op_db)), 1)
        except (TypeError, ValueError):
            corr_db = op_db
    noise_assess_val = ws[f"J{r_noise_res}"].value
    try:
        corr_f = float(str(corr_db).replace(",", ".").strip())
        noise_pass = corr_f <= 65.0
    except Exception:
        noise_pass = True

    if noise_assess_val in ("Đạt", "Không đạt"):
        noise_ass_vn = noise_assess_val
        noise_ass_en = "Pass" if noise_ass_vn == "Đạt" else "Fail"
    else:
        noise_ass_vn = "Đạt" if noise_pass else "Không đạt"
        noise_ass_en = "Pass" if noise_pass else "Fail"

    # 5. HEPA Leak
    df_leak = ws[f"D{r_hepa_in}"].value if ws[f"D{r_hepa_in}"].value is not None else 0.0006
    ex_leak = ws[f"D{r_hepa_ex}"].value if ws[f"D{r_hepa_ex}"].value is not None else 0.0000

    # 6. UVC
    uvc_vals = [ws.cell(r_uvc_raw, c).value for c in range(3, 15) if ws.cell(r_uvc_raw, c).value is not None and isinstance(ws.cell(r_uvc_raw, c).value, (int, float))]
    uvc_avg = ws[f"D{r_uvc_res}"].value or (sum(uvc_vals) / len(uvc_vals) if uvc_vals else None)
    uvc_rounded = None
    if uvc_avg is not None and str(uvc_avg).strip() not in ('-', '', 'None'):
        try:
            uvc_rounded = int(round(float(uvc_avg)))
        except (ValueError, TypeError):
            uvc_rounded = None

    reg_date_str = reg_date.strftime("%d/%m/%Y") if hasattr(reg_date, "strftime") else str(reg_date)
    sample_date_str = sample_date.strftime("%d/%m/%Y") if hasattr(sample_date, "strftime") else str(sample_date)

    parsed = {
        "metadata": {
            "code": str(code),
            "sub_no": str(sub_no) if sub_no else "",
            "reg_date": reg_date_str,
            "sample_date": sample_date_str,
            "start_time": str(start_time),
            "end_time": str(end_time)
        },
        "cabinet": {
            "cabinet_name": str(cab_name),
            "brand": str(brand),
            "model": str(model),
            "serial": str(serial),
            "device_id": str(dev_id)
        },
        "inflow": (
            {
                "method": 2,
                "method_name": "Kanomax Climomaster 6501",
                "width_mm": w_mm if isinstance(w_mm, (int, float)) else 930,
                "height_mm": h_mm if isinstance(h_mm, (int, float)) else 200,
                "velocity": inflow_vel,
                "readings": [
                    [ws.cell(k22, c).value for c in range(3, 11) if isinstance(ws.cell(k22, c).value, (int, float))],
                    [ws.cell(k24, c).value for c in range(3, 11) if isinstance(ws.cell(k24, c).value, (int, float))]
                ] if ws[f"C{k22}"].value is not None else [
                    [ws.cell(k23, c).value for c in range(3, 11) if isinstance(ws.cell(k23, c).value, (int, float))]
                ]
            } if is_m2 else {
                "method": 1,
                "method_name": "DIM TSI 8380",
                "width_mm": w_mm,
                "height_mm": h_mm,
                "q_m3h": b12_val if isinstance(b12_val, (int, float)) else 267.84,
                "area_m2": ws["A12"].value if isinstance(ws["A12"].value, (int, float)) else ((float(w_mm) * float(h_mm)) / 1_000_000.0),
                "velocity": inflow_vel
            }
        ),
        "downflow": {
            "readings": df_readings,
            "avg": df_avg,
            "tolerance_pct": df_tol,
            "eval_uniformity": df_eval_uniformity
        },
        "lux": {
            "readings": lux_readings,
            "avg": lux_avg,
            "avg_rounded": lux_rounded
        },
        "noise": {
            "ambient": amb_db,
            "operating": op_db,
            "corrected_dba": corr_db,
            "assessment": noise_ass_vn,
            "assessment_en": noise_ass_en
        },
        "hepa": {
            "downflow_pct": df_leak,
            "exhaust_pct": ex_leak
        }
    }
    if uvc_avg is not None and str(uvc_avg).strip() not in ('-', '', 'None') and uvc_rounded is not None:
        parsed["uvc"] = {
            "readings": uvc_vals,
            "avg": uvc_avg,
            "avg_rounded": uvc_rounded
        }

    ws_info = wb["Thông tin"] if "Thông tin" in wb.sheetnames else None
    if ws_info:
        c_vn = str(ws_info["C9"].value or "").strip()
        c_en = str(ws_info["Q9"].value or "").strip()
        a_vn = str(ws_info["C11"].value or "").strip()
        a_en = str(ws_info["Q11"].value or "").strip()
        l_vn = str(ws_info["C12"].value or ws_info["C11"].value or "").strip()
        l_en = str(ws_info["Q12"].value or ws_info["Q11"].value or "").strip()
        if c_vn:
            parsed["client_name_vn"] = c_vn
        if c_en:
            parsed["client_name_en"] = c_en
        if a_vn:
            parsed["client_address_vn"] = a_vn
        if a_en:
            parsed["client_address_en"] = a_en
        if l_vn:
            parsed["location_vn"] = l_vn
        if l_en:
            parsed["location_en"] = l_en

    p_vals = [ws[f"{c}{r_part}"].value for c in ["D", "F", "H", "J"]]
    if any(v not in (None, "-", "") for v in p_vals):
        p_nums = [float(v) if isinstance(v, (int, float)) else 0.0 for v in p_vals]
        limits = [10200, 3520, 832, 29]
        ass_vn = ["Đạt" if p <= lim else "Không đạt" for p, lim in zip(p_nums, limits)]
        ass_en = ["Pass" if p <= lim else "Fail" for p, lim in zip(p_nums, limits)]
        parsed["particle"] = {
            "readings": [v if isinstance(v, (int, float)) else 0 for v in p_vals],
            "sizes": [
                ("Kích thước hạt d ≥ 0,3 µm", "Particle size d ≥ 0.3 µm", f"{p_vals[0] if p_vals[0] not in (None, '-') else 0} hạt/m³", f"{p_vals[0] if p_vals[0] not in (None, '-') else 0} particles/m³", "≤ 10 200\nhạt/m³ ⁽²⁾", "≤ 10,200\nparticles/m³ ⁽²⁾", ass_vn[0], ass_en[0]),
                ("Kích thước hạt d ≥ 0,5 µm", "Particle size d ≥ 0.5 µm", f"{p_vals[1] if p_vals[1] not in (None, '-') else 0} hạt/m³", f"{p_vals[1] if p_vals[1] not in (None, '-') else 0} particles/m³", "≤ 3 520\nhạt/m³ ⁽²⁾", "≤ 3,520\nparticles/m³ ⁽²⁾", ass_vn[1], ass_en[1]),
                ("Kích thước hạt d ≥ 1,0 µm", "Particle size d ≥ 1.0 µm", f"{p_vals[2] if p_vals[2] not in (None, '-') else 0} hạt/m³", f"{p_vals[2] if p_vals[2] not in (None, '-') else 0} particles/m³", "≤ 832\nhạt/m³ ⁽²⁾", "≤ 832\nparticles/m³ ⁽²⁾", ass_vn[2], ass_en[2]),
                ("Kích thước hạt d ≥ 5,0 µm", "Particle size d ≥ 5.0 µm", f"{p_vals[3] if p_vals[3] not in (None, '-') else 0} hạt/m³", f"{p_vals[3] if p_vals[3] not in (None, '-') else 0} particles/m³", "≤ 29\nhạt/m³ ⁽²⁾", "≤ 29\nparticles/m³ ⁽²⁾", ass_vn[3], ass_en[3]),
            ]
        }

    return parsed
