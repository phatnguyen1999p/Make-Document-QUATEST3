"""
ISO 8573-1:2010 Calculation Engine for Compressed Air (Khí nén) — QUATEST 3 (Room N2).
Implements:
1. Volume conversion factor N* to ISO 8573-1 reference conditions (20 °C, 1 bar abs)
2. Solid Particle Concentration (Lasair III 8-channel & AeroTrak 9350 6-channel) & Table 1 Classification
3. Pressure Dewpoint & CS Instruments 'CS Humidity Calculator' exact thermodynamic model & Table 2 Classification
4. Total Oil Content (Aerosol + Vapor) & Table 3 Classification
5. Microbiological Contaminants with Feller Correction
6. Direct reading, verification, and bug-fixing (Bẫy 15, Bẫy 41, K6 bug, H8 water vapor bug) for TinhToanKhiNen.xlsx
"""
import os
import re
import sys
import math
import json
import argparse
import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

# ==============================================================================
# 1. REFERENCE CONDITION FACTOR N* (ISO 8573-1:2010)
# ==============================================================================
def calc_n_star(t_measured_c: float, p_measured_bar: float = 1.0) -> float:
    """
    N* = (1 / Pdd) * ((273 + Tdd) / (273 + 20))
     matches Excel formula: =(1/C37)*((273+C36)/(273+20))
    """
    return (1.0 / p_measured_bar) * ((273.0 + t_measured_c) / 293.0)


def excel_round(val: float, ndigits: int = 0) -> float | int:
    """Standard half-away-from-zero rounding matching Excel ROUND()."""
    factor = 10.0 ** ndigits
    scaled = val * factor
    if scaled >= 0:
        res = math.floor(scaled + 0.5) / factor
    else:
        res = math.ceil(scaled - 0.5) / factor
    return int(res) if ndigits == 0 else res


# ==============================================================================
# 2. SOLID PARTICLES (ISO 8573-1:2010 Clause 5.1, Table 1)
# ==============================================================================
def classify_iso8573_particles(
    d_01_05: int | None,
    d_05_10: int,
    d_10_50: int,
    d_gt_50: int,
) -> dict:
    """
    Classify solid particle concentration (particles/m3 at reference conditions)
    according to ISO 8573-1:2010 Table 1.
    """
    # Individual band classes (when d_gt_50 == 0)
    if d_01_05 is None:
        c_band1 = None
    elif d_01_05 <= 20000:
        c_band1 = 1
    elif d_01_05 <= 400000:
        c_band1 = 2
    else:
        c_band1 = 3  # > 400,000 -> at least Class 3 (Table 1 does not specify 0.1-0.5 limit for Class 3+)

    if d_05_10 <= 400:
        c_band2 = 1
    elif d_05_10 <= 6000:
        c_band2 = 2
    elif d_05_10 <= 90000:
        c_band2 = 3
    else:
        c_band2 = 4  # > 90,000 -> at least Class 4

    if d_10_50 <= 10:
        c_band3 = 1
    elif d_10_50 <= 100:
        c_band3 = 2
    elif d_10_50 <= 1000:
        c_band3 = 3
    elif d_10_50 <= 10000:
        c_band3 = 4
    elif d_10_50 <= 100000:
        c_band3 = 5
    else:
        c_band3 = 6  # > 100,000 -> Class > 5

    if d_gt_50 > 0:
        overall_excel = "Không phân loại cho có hạt 5 µm"
        overall_tbkq_vi = "Không phân loại (3)"
        overall_tbkq_en = "Not classified (3)"
        overall_class_num = None
    else:
        bands = [x for x in (c_band1, c_band2, c_band3) if x is not None]
        max_c = max(bands)
        overall_class_num = max_c
        if max_c > 5:
            overall_excel = "Class >5"
            overall_tbkq_vi = "> Class 5"
            overall_tbkq_en = "> Class 5"
        else:
            overall_excel = f"Class {max_c}"
            overall_tbkq_vi = f"Class {max_c}"
            overall_tbkq_en = f"Class {max_c}"

    return {
        "d_0.1_0.5_m3": d_01_05,
        "d_0.5_1.0_m3": d_05_10,
        "d_1.0_5.0_m3": d_10_50,
        "d_gt_5.0_m3": d_gt_50,
        "d_gt_0.5_total_m3": d_05_10 + d_10_50 + d_gt_50,
        "band_classes": {
            "0.1_0.5": c_band1,
            "0.5_1.0": c_band2,
            "1.0_5.0": c_band3,
            "gt_5.0": "Có hạt 5 µm" if d_gt_50 > 0 else 0,
        },
        "class_num": overall_class_num,
        "class_excel": overall_excel,
        "class_tbkq_vi": overall_tbkq_vi,
        "class_tbkq_en": overall_tbkq_en,
    }


def calc_particles_from_lasair(
    channels_8: list[float],
    n_star: float,
    dilution_factor: float = 1.0,
) -> dict:
    """
    Calculate ISO 8573-1 particle bands from 8 Lasair III differential channels (particles/m3):
    channels_8 = [0.1, 0.15, 0.20, 0.25, 0.30, 0.50, 1.0, 5.0 um]
    """
    if len(channels_8) != 8:
        raise ValueError("Lasair III requires 8 channel values: [0.1, 0.15, 0.2, 0.25, 0.3, 0.5, 1.0, 5.0]")
    raw_01_05 = sum(channels_8[0:5]) * dilution_factor
    raw_05_10 = channels_8[5] * dilution_factor
    raw_10_50 = channels_8[6] * dilution_factor
    raw_gt_50 = channels_8[7] * dilution_factor

    std_01_05 = excel_round(raw_01_05 * n_star, 0)
    std_05_10 = excel_round(raw_05_10 * n_star, 0)
    std_10_50 = excel_round(raw_10_50 * n_star, 0)
    std_gt_50 = excel_round(raw_gt_50 * n_star, 0)

    res = classify_iso8573_particles(std_01_05, std_05_10, std_10_50, std_gt_50)
    res["raw_bands_m3"] = {
        "0.1_0.5": excel_round(raw_01_05, 0),
        "0.5_1.0": excel_round(raw_05_10, 0),
        "1.0_5.0": excel_round(raw_10_50, 0),
        "gt_5.0": excel_round(raw_gt_50, 0),
    }
    return res


def calc_particles_from_aerotrak(
    channels_6: list[float],
    n_star: float,
    dilution_factor: float = 1.0,
    is_cumulative: bool = False,
) -> dict:
    """
    Calculate ISO 8573-1 particle bands from 6 AeroTrak 9350 channels:
    channels_6 = [0.3, 0.5, 1.0, 2.0, 3.0, 5.0 um] (per m3).
    If is_cumulative=True (mode Σ), converts to differential Δ first.
    """
    if len(channels_6) != 6:
        raise ValueError("AeroTrak 9350 requires 6 channel values: [0.3, 0.5, 1.0, 2.0, 3.0, 5.0]")
    if is_cumulative:
        c03 = channels_6[0] - channels_6[1]
        c05 = channels_6[1] - channels_6[2]
        c10 = channels_6[2] - channels_6[3]
        c20 = channels_6[3] - channels_6[4]
        c30 = channels_6[4] - channels_6[5]
        c50 = channels_6[5]
    else:
        c03, c05, c10, c20, c30, c50 = channels_6

    raw_03_05 = c03 * dilution_factor
    raw_05_10 = c05 * dilution_factor
    raw_10_50 = (c10 + c20 + c30) * dilution_factor
    raw_gt_50 = c50 * dilution_factor

    std_03_05 = excel_round(raw_03_05 * n_star, 0)
    std_05_10 = excel_round(raw_05_10 * n_star, 0)
    std_10_50 = excel_round(raw_10_50 * n_star, 0)
    std_gt_50 = excel_round(raw_gt_50 * n_star, 0)

    res = classify_iso8573_particles(std_03_05, std_05_10, std_10_50, std_gt_50)
    res["raw_bands_m3"] = {
        "0.3_0.5": excel_round(raw_03_05, 0),
        "0.5_1.0": excel_round(raw_05_10, 0),
        "1.0_5.0": excel_round(raw_10_50, 0),
        "gt_5.0": excel_round(raw_gt_50, 0),
    }
    return res


# ==============================================================================
# 3. PRESSURE DEWPOINT & CS HUMIDITY CALCULATOR (ISO 8573-1:2010 Clause 5.2, Table 2)
# ==============================================================================
def _sonntag_vapor_pressure_hpa(temp_c: float) -> float:
    """
    Exact WMO 2008 / Sonntag 1990 saturation vapor pressure (hPa) used by CS Humidity Calculator:
    - Over water (T >= 0 °C): es(T) = 6.112 * exp(17.62 * T / (243.12 + T))
    - Over ice   (T <  0 °C): es(T) = 6.112 * exp(22.46 * T / (272.62 + T))
    """
    if temp_c >= 0.0:
        return 6.112 * math.exp((17.62 * temp_c) / (243.12 + temp_c))
    else:
        return 6.112 * math.exp((22.46 * temp_c) / (272.62 + temp_c))


def _invert_sonntag_temp_c(e_hpa: float) -> float:
    """Invert WMO 2008 / Sonntag 1990 vapor pressure (hPa) to Dewpoint (°Ctd)."""
    ln_r = math.log(e_hpa / 6.112)
    if e_hpa >= 6.112:
        return (243.12 * ln_r) / (17.62 - ln_r)
    else:
        return (272.62 * ln_r) / (22.46 - ln_r)


def classify_iso8573_dewpoint(td2_7bar_c: float) -> str:
    """Classify pressure dewpoint at 7 bar gauge according to ISO 8573-1:2010 Table 2."""
    if td2_7bar_c <= -70.0:
        return "Class 1"
    elif td2_7bar_c <= -40.0:
        return "Class 2"
    elif td2_7bar_c <= -20.0:
        return "Class 3"
    elif td2_7bar_c <= 3.0:
        return "Class 4"
    elif td2_7bar_c <= 7.0:
        return "Class 5"
    elif td2_7bar_c <= 10.0:
        return "Class 6"
    else:
        return "> Class 6"


def calc_dewpoint_cs(
    t1_start_c: float,
    t1_end_c: float,
    tkhi_start_c: float,
    tkhi_end_c: float,
    p_start_bar: float,
    p_end_bar: float,
    p_ref_bar: float = 7.0,
    p_atm_bar: float = 1.01325,
    round_inputs_1dec: bool = True,
) -> dict:
    """
    Exact replica of CS Instruments 'CS Humidity Calculator' + ISO 8573-1 Table 2 classification.
    Note: Technicians enter the 1-decimal displayed averages from the Excel sheet into CS Humidity Calculator
    (e.g. -41.9 °Ctd, 21.8 °C, 7.3 bar). Setting round_inputs_1dec=True matches CS software screenshot 100%.
    """
    t1_avg_exact = (t1_start_c + t1_end_c) / 2.0
    tkhi_avg_exact = (tkhi_start_c + tkhi_end_c) / 2.0
    p_avg_exact = (p_start_bar + p_end_bar) / 2.0

    t1_in = excel_round(t1_avg_exact, 1) if round_inputs_1dec else t1_avg_exact
    tkhi_in = excel_round(tkhi_avg_exact, 1) if round_inputs_1dec else tkhi_avg_exact
    p1_in = excel_round(p_avg_exact, 1) if round_inputs_1dec else p_avg_exact

    es_tkhi = _sonntag_vapor_pressure_hpa(tkhi_in)
    e1 = _sonntag_vapor_pressure_hpa(t1_in)
    rh1 = (e1 / es_tkhi) * 100.0

    # Pressure conversion to p_ref_bar (7 bar gauge)
    p_ratio = (p_ref_bar + p_atm_bar) / (p1_in + p_atm_bar)
    e2 = e1 * p_ratio
    rh2 = rh1 * p_ratio
    td2 = _invert_sonntag_temp_c(e2)

    # Absolute humidity (mg/m3) = 216.679 * e_hPa / (T + 273.15) * 1000
    abs_hum_1_mg_m3 = 216.679 * e1 / (tkhi_in + 273.15) * 1000.0
    abs_hum_2_mg_m3 = 216.679 * e2 / (tkhi_in + 273.15) * 1000.0

    # Volume mixing ratio (ppmV) & Humidity grade (g/kg)
    p1_total_hpa = (p1_in + p_atm_bar) * 1000.0
    ppm_v = (e1 / (p1_total_hpa - e1)) * 1e6
    hum_grade_g_kg = 622.0 * e1 / (p1_total_hpa - e1)

    td2_rounded_2 = excel_round(td2, 2)
    td2_rounded_1 = excel_round(td2, 1)
    iso_class = classify_iso8573_dewpoint(td2_rounded_2)

    return {
        "averages_exact": {
            "T1_c": round(t1_avg_exact, 4),
            "Tkhi_c": round(tkhi_avg_exact, 4),
            "P_bar": round(p_avg_exact, 4),
        },
        "cs_inputs": {
            "Dewpoint_Ctd": t1_in,
            "Temperature_C": tkhi_in,
            "OverPressure_bar": p1_in,
            "TargetOverPressure_bar": p_ref_bar,
        },
        "cs_measured_condition": {
            "RelativeHumidity_pct": round(rh1, 10),
            "Dewpoint_Ctd": round(t1_in, 2),
            "AbsoluteHumidity_mg_m3": excel_round(abs_hum_1_mg_m3, 2),
            "AbsoluteHumidity_tbkq_int": excel_round(abs_hum_1_mg_m3, 0),
            "HumidityGrade_g_kg": excel_round(hum_grade_g_kg, 3),
            "VaporRatio_ppm": excel_round(ppm_v, 2),
            "SaturationVaporPressure_hPa": excel_round(es_tkhi, 3),
            "PartialVaporPressure_hPa": excel_round(e1, 3),
        },
        "cs_converted_7bar": {
            "RelativeHumidity_pct": round(rh2, 10),
            "Dewpoint_T2_Ctd": td2_rounded_2,
            "Dewpoint_T2_tbkq_1dec": td2_rounded_1,
            "AbsoluteHumidity_mg_m3": excel_round(abs_hum_2_mg_m3, 2),
            "PartialVaporPressure_hPa": excel_round(e2, 3),
        },
        "class_iso8573": iso_class,
    }


# ==============================================================================
# 4. TOTAL OIL CONTENT (ISO 8573-1:2010 Clause 5.3, Table 3)
# ==============================================================================
def _parse_kqtn_ug(val) -> tuple[float, float | None]:
    """
    Parse KQTN value (e.g., 15.2, '< 10', '< 5', 'KPH').
    Returns (numeric_value_for_sum, lod_ug_if_below_lod).
    """
    if isinstance(val, (int, float)):
        return float(val), None
    s = str(val or "").strip().replace(",", ".")
    if not s or s.upper() == "KPH" or s.startswith("<"):
        m = re.search(r"(\d+(?:\.\d+)?)", s)
        lod = float(m.group(1)) if m else None
        return 0.0, lod
    try:
        return float(s), None
    except ValueError:
        return 0.0, None


def classify_iso8573_oil(oil_std_mg_m3: float | str) -> str:
    """Classify Total Oil content (mg/m3 at reference conditions) per ISO 8573-1:2010 Table 3."""
    if isinstance(oil_std_mg_m3, str):
        if "< 0,01" in oil_std_mg_m3 or "< 0.01" in oil_std_mg_m3:
            return "Class 1"
        val = float(oil_std_mg_m3.replace("<", "").replace(",", ".").strip())
    else:
        val = float(oil_std_mg_m3)

    if val <= 0.01:
        return "Class 1"
    elif val <= 0.1:
        return "Class 2"
    elif val <= 1.0:
        return "Class 3"
    elif val <= 5.0:
        return "Class 4"
    else:
        return "> Class 4"


def calc_total_oil(
    flow_aerosol_l_min: float,
    time_aerosol_min: float,
    kqtn_aerosol_ug,
    flow_vapor_l_min: float,
    time_vapor_min: float,
    kqtn_vapor_ug,
    n_star: float,
) -> dict:
    """
    Calculate Total Oil Content (Aerosol + Vapor) in mg/m3 at reference conditions
    and classify according to ISO 8573-1:2010 Table 3.
    """
    v_aerosol_l = flow_aerosol_l_min * time_aerosol_min
    v_vapor_l = flow_vapor_l_min * time_vapor_min

    ug_aer, lod_aer = _parse_kqtn_ug(kqtn_aerosol_ug)
    ug_vap, lod_vap = _parse_kqtn_ug(kqtn_vapor_ug)

    c_aer = (ug_aer / v_aerosol_l) if v_aerosol_l > 0 else 0.0
    c_vap = (ug_vap / v_vapor_l) if v_vapor_l > 0 else 0.0
    c_sum = c_aer + c_vap
    c_std = c_sum * n_star

    # Check effective LOD in mg/m3 if KPH/<LOD
    lod_aer_mg_m3 = (lod_aer / v_aerosol_l * n_star) if (lod_aer and v_aerosol_l > 0) else None
    lod_vap_mg_m3 = (lod_vap / v_vapor_l * n_star) if (lod_vap and v_vapor_l > 0) else None

    if c_sum <= 0.01:
        res_vi = "< 0,01"
        res_en = "< 0.01"
        iso_class = "Class 1"
    else:
        val_r = excel_round(c_std, 2)
        res_en = f"{val_r:.2f}"
        res_vi = res_en.replace(".", ",")
        iso_class = classify_iso8573_oil(val_r)

    return {
        "v_aerosol_L": v_aerosol_l,
        "v_vapor_L": v_vapor_l,
        "c_aerosol_mg_m3": round(c_aer, 6),
        "c_vapor_mg_m3": round(c_vap, 6),
        "c_total_std_mg_m3": round(c_std, 6),
        "result_vi": res_vi,
        "result_en": res_en,
        "class_iso8573": iso_class,
        "effective_lod_mg_m3": {
            "aerosol": round(lod_aer_mg_m3, 4) if lod_aer_mg_m3 else None,
            "vapor": round(lod_vap_mg_m3, 4) if lod_vap_mg_m3 else None,
        },
    }


# ==============================================================================
# 5. MICROBIOLOGICAL CONTAMINANTS (Feller Correction)
# ==============================================================================
def calc_microbiology_feller(
    counts_dict: dict[str, int | float | str],
    n_holes: int = 218,
    sample_vol_l: float = 500.0,
) -> dict:
    """
    Calculate Feller-corrected CFU (Pr) and microbial density (CFU/m3).
    Matches Excel formula in Tinh_Toan_Khi_Co_Ap_V6:
      Pr = N * (-LN(1 - r / N))
      CFU/m3 = ROUND(Pr / (V_L / 1000), 0)
    Handles '-', '< 1', 'KPH', '………' safely without #VALUE! (fixes Bẫy 15).
    """
    vol_m3 = sample_vol_l / 1000.0
    out = {}
    for param, raw_val in counts_dict.items():
        s = str(raw_val).strip() if raw_val is not None else "-"
        if s in ("-", ""):
            out[param] = {"raw": s, "Pr": "-", "cfu_m3_vi": "-", "cfu_m3_en": "-"}
        elif "…" in s or "..." in s:
            out[param] = {"raw": s, "Pr": "………", "cfu_m3_vi": "………", "cfu_m3_en": "………"}
        elif s.startswith("<") or s.upper() == "KPH" or s == "0":
            out[param] = {
                "raw": s,
                "Pr": 0.0,
                "cfu_m3_vi": "Không có chỉ thị",
                "cfu_m3_en": "No indication",
                "cfu_m3_limit_fmt": f"< {int(round(1.0 / vol_m3))}" if vol_m3 > 0 else "< 1",
            }
        else:
            r = float(s.replace(",", "."))
            if r <= 0:
                pr = 0.0
                cfu_m3 = 0
            else:
                pr = n_holes * (-math.log(1.0 - (r / n_holes)))
                cfu_m3 = excel_round(pr / vol_m3, 0)
            out[param] = {
                "raw": r,
                "Pr": round(pr, 4),
                "cfu_m3_vi": str(cfu_m3),
                "cfu_m3_en": str(cfu_m3),
            }
    return out


# ==============================================================================
# 6. AUDIT & FIX EXCEL WORKBOOK (TinhToanKhiNen.xlsx)
# ==============================================================================
def dew_rows(ws) -> dict:
    """Dòng bảng điểm sương theo nhãn cột B (form gốc: dew 87, temp 88, pres 89). Phiếu có hiệu chỉnh theo giấy
    hiệu chuẩn (khinen_calc.finalize_khinen_workbook) có thêm 'dew_hc', 'pres_hc' (dòng phần b) sau hiệu chỉnh)."""
    out = {}
    for r in range(80, min((ws.max_row or 0), 140) + 1):
        v = str(ws.cell(r, 2).value or "").strip().lower()
        if not v:
            continue
        hc = "sau hiệu chỉnh" in v
        if v.startswith("nhiệt độ điểm sương"):
            out.setdefault("dew_hc" if hc else "dew", r)
        elif v.startswith("nhiệt độ khí"):
            out.setdefault("temp", r)
        elif v.startswith("áp suất"):
            out.setdefault("pres_hc" if hc else "pres", r)
    out.setdefault("dew", 87)
    out.setdefault("temp", 88)
    out.setdefault("pres", 89)
    out["off"] = out["dew"] - 87          # số dòng đã chèn phía trên (dòng ảnh / giờ đo dời theo)
    return out


def audit_and_calc_from_excel(xlsx_path: str, fix_bugs: bool = False) -> dict:
    """
    Read input parameters from an existing TinhToanKhiNen.xlsx, run the independent
    ISO 8573-1 calculation engine, compare with cached Excel results, and detect/fix:
    - Bẫy 15: #VALUE! in Microbiology (rows 82-83, D11:D15) and unused AeroTrak (rows 67-68)
    - Bug K6: K6 formula referencing E67 (AeroTrak) instead of D6 (1.0-5.0 um particle result)
    - Bug H8: Hardcoded old water content string in H8 not matching CS Humidity Calculator
    """
    wb_v = openpyxl.load_workbook(xlsx_path, data_only=True)
    wb_f = openpyxl.load_workbook(xlsx_path, data_only=False)
    ws_v = wb_v["Tinh_Toan_Khi_Co_Ap_V6"]
    ws_f = wb_f["Tinh_Toan_Khi_Co_Ap_V6"]

    # 1. Read Dewpoint & Ambient inputs (form gốc rows 87-89; phiếu có phần "sau hiệu chỉnh" dò theo nhãn cột B)
    R = dew_rows(ws_v)
    t1_row = R.get("dew_hc", R["dew"])
    p_row = R.get("pres_hc", R["pres"])
    t1_start = float(ws_v[f"C{t1_row}"].value)
    t1_end = float(ws_v[f"D{t1_row}"].value)
    tkhi_start = float(ws_v[f"C{R['temp']}"].value)
    tkhi_end = float(ws_v[f"D{R['temp']}"].value)
    p_start = float(ws_v[f"C{p_row}"].value)
    p_end = float(ws_v[f"D{p_row}"].value)
    p_dd = float(ws_v["C37"].value or 1.0)

    dew_res = calc_dewpoint_cs(t1_start, t1_end, tkhi_start, tkhi_end, p_start, p_end)
    n_star = calc_n_star(dew_res["averages_exact"]["Tkhi_c"], p_dd)

    # 2. Read Oil inputs (rows 43-46)
    flow_aer = float(ws_v["C43"].value or 35)
    flow_vap = float(ws_v["D43"].value or 1)
    time_aer = float(ws_v["C44"].value or 240)
    time_vap = float(ws_v["D44"].value or 240)
    kq_aer = ws_v["C46"].value
    kq_vap = ws_v["D46"].value
    oil_res = calc_total_oil(flow_aer, time_aer, kq_aer, flow_vap, time_vap, kq_vap, n_star)

    # 3. Read Particle inputs (Lasair C54:J54 or AeroTrak C63:H63)
    dev_name = str(ws_v["C52"].value or "")
    if "Lasair" in dev_name and ws_v["C54"].value not in (None, "-"):
        ch8 = [float(ws_v.cell(54, col).value or 0) for col in range(3, 11)]
        dil = float(ws_v["C55"].value or 1.0)
        part_res = calc_particles_from_lasair(ch8, n_star, dil)
    else:
        ch6 = [float(ws_v.cell(63, col).value or 0) for col in range(3, 9)]
        dil = float(ws_v["C64"].value or 1.0)
        part_res = calc_particles_from_aerotrak(ch6, n_star, dil)

    # 4. Read Microbiology inputs (rows 77-81)
    n_holes = int(float(ws_v["C77"].value or 218))
    v_micro_l = float(ws_v["C78"].value or 500)
    micro_in = {
        str(ws_v.cell(80, c).value or f"Col{c}"): ws_v.cell(81, c).value
        for c in range(3, 8)
    }
    micro_res = calc_microbiology_feller(micro_in, n_holes, v_micro_l)

    # 5. Audit Excel formula bugs & discrepancies
    bugs_found = []
    k6_formula = str(ws_f["K6"].value or "")
    if "E67" in k6_formula:
        bugs_found.append(
            f"[Bug K6]: Cell K6 formula '{k6_formula}' references E67 (AeroTrak) instead of D6! "
            f"(Cached K6={ws_v['K6'].value}, true Class for 1.0-5.0um={part_res['band_classes']['1.0_5.0']})"
        )

    excel_g87 = ws_v[f"G{R['dew']}"].value
    if excel_g87 is not None and abs(float(excel_g87) - dew_res["cs_converted_7bar"]["Dewpoint_T2_Ctd"]) > 0.05:
        bugs_found.append(
            f"[Dewpoint G87 mismatch]: Excel G87={excel_g87} vs CS Calculator={dew_res['cs_converted_7bar']['Dewpoint_T2_Ctd']}"
        )

    h8_val = str(ws_v["H8"].value or "")
    expected_h8 = f"{dew_res['cs_measured_condition']['AbsoluteHumidity_tbkq_int']} mg/m3"
    if h8_val and str(dew_res["cs_measured_condition"]["AbsoluteHumidity_tbkq_int"]) not in h8_val:
        bugs_found.append(
            f"[Stale H8 Water Content]: Excel H8='{h8_val}' vs CS Calculator Absolute Humidity='{expected_h8}' "
            f"({dew_res['cs_measured_condition']['AbsoluteHumidity_mg_m3']} mg/m3)"
        )

    for coord in ("D11", "D12", "D13", "D14", "D15", "C67", "D67", "F67", "C82", "C83"):
        val_c = ws_v[coord].value
        if isinstance(val_c, str) and val_c.startswith("#"):
            bugs_found.append(f"[Bẫy 15 - Formula Error]: Cell {coord} has '{val_c}' (formula: {ws_f[coord].value})")

    if fix_bugs:
        # Fix K6 formula to reference D6 instead of E67
        if "E67" in str(ws_f["K6"].value or ""):
            ws_f["K6"].value = str(ws_f["K6"].value).replace("E67", "D6")
        # Fix AeroTrak rows 67-68 when C63:H63 is '-'
        ws_f["C67"].value = '=IF(ISNUMBER(C63),C63*C64,"-")'
        ws_f["D67"].value = '=IF(ISNUMBER(D63),D63*C64,"-")'
        ws_f["E67"].value = '=IF(ISNUMBER(E63),SUM(E63:G63)*C64,"-")'
        ws_f["F67"].value = '=IF(ISNUMBER(H63),H63*C64,"-")'
        for col_let in ("C", "D", "E", "F"):
            ws_f[f"{col_let}68"].value = f'=IF(ISNUMBER({col_let}67),ROUND({col_let}67*$C$38,0),"-")'
        # Fix Microbiology rows 82-83 when C81:G81 is '-' (Bẫy 15)
        for col_let in ("C", "D", "E", "F", "G"):
            ws_f[f"{col_let}82"].value = (
                f'=IF(OR({col_let}81="-",{col_let}81=""),"-",'
                f'IF(OR({col_let}81=0,{col_let}81="< 1",{col_let}81="KPH"),0,$C$77*(-LN(1-({col_let}81/$C$77)))))'
            )
            ws_f[f"{col_let}83"].value = (
                f'=IF({col_let}82="-","-",'
                f'IF(OR({col_let}81="< 1",{col_let}81="KPH"),"Không có chỉ thị",ROUND(IF({col_let}82=0,0,({col_let}82/($C$78/1000))),0)))'
            )
        wb_f.save(xlsx_path)
        print(f"Fixed formula bugs (K6, AeroTrak R67:68, Feller R82:83) in {os.path.basename(xlsx_path)}")

    wb_v.close()
    wb_f.close()

    report = {
        "file": xlsx_path,
        "N_star": round(n_star, 8),
        "particles": part_res,
        "dewpoint": dew_res,
        "total_oil": oil_res,
        "microbiology": micro_res,
        "bugs_found": bugs_found,
    }
    return report


def main():
    parser = argparse.ArgumentParser(description="ISO 8573-1:2010 Calculation & Verification Engine")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_aud = sub.add_parser("audit-excel", help="Audit & calculate ISO 8573-1 from TinhToanKhiNen.xlsx")
    p_aud.add_argument("xlsx", help="Path to TinhToanKhiNen.xlsx")
    p_aud.add_argument("--fix", action="store_true", help="Fix formula bugs (K6, Bẫy 15) in-place")

    p_dp = sub.add_parser("dewpoint", help="Run CS Humidity Calculator & ISO 8573-1 Table 2 classification")
    p_dp.add_argument("--t1", required=True, help="Start,End dewpoint T1 (°C) e.g. '-39.77,-44.09'")
    p_dp.add_argument("--tkhi", required=True, help="Start,End gas temp (°C) e.g. '22.02,21.57'")
    p_dp.add_argument("--p", required=True, help="Start,End gauge pressure (bar) e.g. '7.31,7.32'")

    args = parser.parse_args()
    if args.cmd == "audit-excel":
        rep = audit_and_calc_from_excel(args.xlsx, fix_bugs=args.fix)
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    elif args.cmd == "dewpoint":
        t1_s, t1_e = map(float, args.t1.split(","))
        tk_s, tk_e = map(float, args.tkhi.split(","))
        p_s, p_e = map(float, args.p.split(","))
        res = calc_dewpoint_cs(t1_s, t1_e, tk_s, tk_e, p_s, p_e)
        print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
