# -*- coding: utf-8 -*-
"""
Compressed Air Calculation Engine & Workbook Builder (QUATEST 3 Framework)
==========================================================================
Integrates:
1. Exact Sonntag 1990 / WMO 2008 dew point conversion to 7 bar gauge matching HumCalc.exe
2. HumCalc screenshot rendering via PIL with pixel precision matching the native software
3. ISO 8573-1:2010 Table 1, 2, 3 classification (Particles, Dewpoint, Oil)
4. Standard QUATEST 3 Compressed Air calculation workbook builder (TinhToanKhiNen_template.xlsx)
   - Injects photos (Location, Dew point start/end, Rendered HumCalc)
   - Populates parameters and formulas
   - Injects PageSetup headers (LeftHeader: Số: <code_cert> - KN<stt>, RightHeader: <Ngày>)
   - Fixes all formula bugs (K6, Bẫy 15, H8 stale string, AeroTrak rows 67-68)
   - Full calculation via Excel COM / python
5. Calculation Excel parser (parse_compressed_air_workbook) to extract data for TBKQ / BCKQ
6. Multi-point Dossier aggregation (create_dossier_from_compressed_air)
7. End-to-end pipeline runner
"""

import os
import re
import shutil
import zipfile
import openpyxl
from typing import Dict, List, Any, Optional
from PIL import Image, ImageDraw, ImageFont

from .iso8573_cal import calc_dewpoint_cs, audit_and_calc_from_excel
from ..Models.dossier import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation
from .tag_excel import fill_workbook_tags
from .excel_borders import thin_outside_borders
from .equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

# Default templates directory
MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
CORE_DIR = os.path.dirname(MODULE_DIR)
TEMPLATES_DIR = os.path.join(CORE_DIR, "Templates")
ASSETS_DIR = os.path.join(TEMPLATES_DIR, "assets")
DEFAULT_TEMPLATE_XLSX = os.path.join(TEMPLATES_DIR, "TinhToanKhiNen_template.xlsx")
DEFAULT_HUMCALC_BASE = os.path.join(ASSETS_DIR, "humcalc_base.png")


# ==============================================================================
# 1. HUMCALC SCREENSHOT RENDERER
# ==============================================================================
def render_humcalc_screenshot(
    base_image_path: str,
    out_image_path: str,
    values: Dict[str, Any]
) -> str:
    """
    Renders an exact, pixel-perfect CS Instruments Humidity Calculator screenshot
    by stamping values onto the base screenshot template with native font.
    """
    if not os.path.exists(base_image_path):
        base_image_path = DEFAULT_HUMCALC_BASE

    base = Image.open(base_image_path).convert("RGB")
    draw = ImageDraw.Draw(base)

    # Use native Segoe UI or Arial font
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    font_path = os.path.join(fonts_dir, "segoeui.ttf")
    if not os.path.exists(font_path):
        font_path = os.path.join(fonts_dir, "arial.ttf")
    font = ImageFont.truetype(font_path, 12) if os.path.exists(font_path) else ImageFont.load_default()

    boxes = {
        # Inputs (Left column top)
        "tbSetPar": ((148, 68, 235, 89), "left", 3),
        "tbSetTemp": ((148, 100, 235, 121), "left", 3),
        "tbSetOverPres": ((148, 131, 235, 152), "left", 3),
        # Target pressure input (Right column top)
        "tbSetOverPresCC": ((615, 173, 693, 192), "left", 3),

        # Measured Condition results (Left column bottom)
        "tbParVal": ((148, 266, 235, 287), "right", 4),
        "tbDewP": ((148, 295, 235, 316), "right", 4),
        "tbAbsHum": ((148, 324, 235, 345), "right", 4),
        "tbAbsHumMg": ((148, 353, 235, 374), "right", 4),
        "tbHumGrade": ((148, 382, 235, 403), "right", 4),
        "tbVapRatio": ((148, 411, 235, 432), "right", 4),
        "tbSatPres": ((148, 440, 235, 461), "right", 4),
        "tbParPres": ((148, 469, 235, 490), "right", 4),

        # Converted 7 bar results (Right column bottom)
        "tbParValCC": ((557, 266, 634, 287), "right", 4),
        "tbDewPCC": ((557, 295, 634, 316), "right", 4),
        "tbAbsHumCC": ((557, 324, 634, 345), "right", 4),
        "tbAbsHumMgCC": ((557, 353, 634, 374), "right", 4),
        "tbHumGradeCC": ((557, 382, 634, 403), "right", 4),
        "tbVapRatioCC": ((557, 411, 634, 432), "right", 4),
        "tbSatPresCC": ((557, 440, 634, 461), "right", 4),
        "tbParPresCC": ((557, 469, 634, 490), "right", 4),
    }

    for key, (rect, align, margin) in boxes.items():
        # Clear existing text box with white fill
        draw.rectangle(rect, fill=(255, 255, 255))
        val_str = str(values.get(key, ""))
        if not val_str:
            continue
        bbox = font.getbbox(val_str)
        text_w = bbox[2] - bbox[0]
        x1, y1, x2, y2 = rect
        y = y1 + 3
        if align == "left":
            x = x1 + margin
        else:
            x = x2 - margin - text_w
        draw.text((x, y), val_str, font=font, fill=(0, 0, 0))

    os.makedirs(os.path.dirname(os.path.abspath(out_image_path)), exist_ok=True)
    base.save(out_image_path)
    return out_image_path


def generate_humcalc_values(dew_res: Dict[str, Any]) -> Dict[str, str]:
    """
    Builds the dictionary of string values for render_humcalc_screenshot
    from the output of calc_dewpoint_cs().
    """
    cs_in = dew_res["cs_inputs"]
    cs_meas = dew_res["cs_measured_condition"]
    cs_conv = dew_res["cs_converted_7bar"]

    t1_in = cs_in["Dewpoint_Ctd"]
    tkhi_in = cs_in["Temperature_C"]
    p1_in = cs_in["OverPressure_bar"]
    pref_in = cs_in["TargetOverPressure_bar"]

    def fmt(v: float, dec: int) -> str:
        return f"{v:.{dec}f}".replace(".", ",")

    abs_g_m3_1 = cs_meas["AbsoluteHumidity_mg_m3"] / 1000.0
    abs_g_m3_2 = cs_conv["AbsoluteHumidity_mg_m3"] / 1000.0

    return {
        "tbSetPar": fmt(t1_in, 1),
        "tbSetTemp": fmt(tkhi_in, 1),
        "tbSetOverPres": fmt(p1_in, 1),
        "tbSetOverPresCC": str(int(round(pref_in))),

        "tbParVal": fmt(cs_meas["RelativeHumidity_pct"], 3),
        "tbDewP": fmt(cs_meas["Dewpoint_Ctd"], 2),
        "tbAbsHum": fmt(abs_g_m3_1, 2),
        "tbAbsHumMg": fmt(cs_meas["AbsoluteHumidity_mg_m3"], 2),
        "tbHumGrade": fmt(cs_meas["HumidityGrade_g_kg"], 3),
        "tbVapRatio": fmt(cs_meas["VaporRatio_ppm"], 2),
        "tbSatPres": fmt(cs_meas["SaturationVaporPressure_hPa"], 3),
        "tbParPres": fmt(cs_meas["PartialVaporPressure_hPa"], 3),

        "tbParValCC": fmt(cs_conv["RelativeHumidity_pct"], 3),
        "tbDewPCC": fmt(cs_conv["Dewpoint_T2_Ctd"], 2),
        "tbAbsHumCC": fmt(abs_g_m3_2, 2),
        "tbAbsHumMgCC": fmt(cs_conv["AbsoluteHumidity_mg_m3"], 2),
        "tbHumGradeCC": fmt(cs_meas["HumidityGrade_g_kg"], 3),
        "tbVapRatioCC": fmt(cs_meas["VaporRatio_ppm"], 2),
        "tbSatPresCC": fmt(cs_meas["SaturationVaporPressure_hPa"], 3),
        "tbParPresCC": fmt(cs_conv["PartialVaporPressure_hPa"], 3),
    }


# ==============================================================================
# 2. ZIP ARCHIVE MEDIA REPLACEMENT
# ==============================================================================
def place_box_images(ws, items: List[Any], temp_dir: str):
    """
    Template lưu ảnh kiểu "ảnh đặt trong ô" (richData) — openpyxl không đọc/giữ được nên ô hiện '#VALUE!'
    và update_zip_media không có gì để thay. Chèn ảnh nổi vừa khung vùng gộp (giữ tỉ lệ, không crop):
    items = [(vùng gộp "B17:D34", đường dẫn ảnh), ...].
    """
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.utils import range_boundaries, get_column_letter
    for rng, path in items:
        if not path or not os.path.isfile(path):
            continue
        c0, r0, c1, r1 = range_boundaries(rng)
        box_w = sum((ws.column_dimensions[get_column_letter(c)].width or 8.43) * 7 + 5 for c in range(c0, c1 + 1))
        box_h = sum((ws.row_dimensions[r].height or 15) * 4 / 3 for r in range(r0, r1 + 1))
        im = Image.open(path).convert("RGB")
        if im.width > 1200:
            im = im.resize((1200, round(im.height * 1200 / im.width)), Image.Resampling.LANCZOS)
        small = os.path.join(temp_dir, f"box_{c0}_{r0}.jpg")
        im.save(small, "JPEG", quality=88)
        scale = min((box_w - 6) / im.width, (box_h - 6) / im.height)
        pic = XLImage(small)
        pic.width, pic.height = round(im.width * scale), round(im.height * scale)
        ws.cell(r0, c0).value = None
        ws.add_image(pic, f"{get_column_letter(c0)}{r0}")


def update_zip_media(xlsx_path: str, media_replacements: Dict[str, str]):
    """
    Safely replaces media files (image1, image2, image3, image4) inside
    an existing xlsx zip archive without corrupting OpenXML relationships.
    """
    temp_path = xlsx_path + ".tmp"
    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(temp_path, "w") as zout:
        for item in zin.infolist():
            if item.filename in media_replacements and os.path.isfile(media_replacements[item.filename]):
                with open(media_replacements[item.filename], "rb") as f:
                    data = f.read()
                zout.writestr(item, data)
            else:
                zout.writestr(item, zin.read(item.filename))
    os.replace(temp_path, xlsx_path)


# ==============================================================================
# 3. WORKBOOK BUILDER (TinhToanKhiNen.xlsx)
# ==============================================================================
# ==============================================================================
# HIỆU CHỈNH THEO GIẤY HIỆU CHUẨN — CS DP510 (điểm sương °Ctd, áp suất bar)
# Theo công tắc chung cleanroom_calc.AP_HIEU_CHINH (tạm tắt 09/11/2026): tắt -> khinen_cals() trả {} -> phiếu như cũ.
# Bật: phiếu thêm "a) Kết quả đọc (chưa hiệu chỉnh)" trên bảng điểm sương và "b) Kết quả sau hiệu chỉnh …" dưới
# (công thức FORECAST trỏ sheet ẩn HieuChuan); điểm sương quy đổi 7 bar, hàm lượng nước, ảnh CS tính từ số sau
# hiệu chỉnh; TBKQ/BCKQ đọc phần b) (iso8573_cal.dew_rows). Máy đếm hạt (AeroTrak, Lasair) không có bảng sai số.
# ==============================================================================
KHINEN_CAL_KEYS = (("dewpoint", "dew", "Nhiệt độ điểm sương T1 sau hiệu chỉnh (ºC)"),
                   ("pressure", "pres", "Áp suất sau hiệu chỉnh (bar)"))


def khinen_cals(mt_code: str, on_date: str, device: str = "") -> Dict[str, Any]:
    """Bảng hiệu chuẩn máy đo điểm sương CS DP510 (mã MT, vd 'MT10-1') còn hiệu lực ngày đo.
    Trả {"dewpoint": CalTable, "pressure": CalTable} (khóa thiếu = không có bảng / công tắc tắt)."""
    from .cleanroom_calc import CalTable
    dev = device or f"CS Instruments DP510 ({mt_code})"
    out = {"dewpoint": CalTable.from_library(mt_code, "Nhiệt độ", on_date, device=dev, unit="°Ctd"),
           "pressure": CalTable.from_library(mt_code, "Chênh áp / áp suất", on_date, device=dev, unit="bar")}
    return {k: v for k, v in out.items() if v is not None and v.points}


def _hc_rows_below(ws) -> int:
    """Số dòng phần b) đã chèn dưới bảng điểm sương (0 nếu phiếu không hiệu chỉnh)."""
    from .iso8573_cal import dew_rows
    R = dew_rows(ws)
    n = sum(k in R for k in ("dew_hc", "pres_hc"))
    return n + 1 if n else 0


def finalize_khinen_workbook(path: str, cals: Dict[str, Any]) -> str:
    """Chèn phần a) / b) vào bảng điểm sương (Excel COM, để Excel tự dời tham chiếu, vùng gộp, ảnh, vùng in)."""
    import gc
    import win32com.client
    from .bsc_cal_blocks import _P, _retry, _title, _write_cal_sheet, _hc_formula
    cals = {k: v for k, v in (cals or {}).items() if v is not None and getattr(v, "points", None)}
    if not cals:
        return path
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        xl = _P(xl)
        wb = xl.Workbooks.Open(os.path.abspath(path))
        xl.ScreenUpdating = False
        xl.Calculation = -4135
        ws = wb.Worksheets("Tinh_Toan_Khi_Co_Ap_V6")
        cal_rows = _write_cal_sheet(wb, cals)
        # a) trên dòng tiêu đề bảng (86) -> dòng đọc 88–90
        _retry(ws.Rows("86:86").Insert)
        _title(ws, 86, "a) Kết quả đọc (chưa hiệu chỉnh)")
        r_dew, r_pres = 88, 90
        keys = [(k, src, lbl) for k, src, lbl in KHINEN_CAL_KEYS if k in cals]
        n = len(keys) + 1
        first = r_pres + 1
        _retry(ws.Rows(f"{first}:{first + n - 1}").Insert)
        try:                                                   # xlDecimalSeparator = 3 (mảng International, gốc 1)
            dec_sep = str(xl.International[2]) or "."
        except Exception:  # noqa: BLE001
            dec_sep = "."
        certs = sorted({c.cert_no for c in cals.values() if c.cert_no})
        _title(ws, first, "b) Kết quả sau hiệu chỉnh theo giấy hiệu chuẩn " + ", ".join(certs) +
               " (điểm sương quy đổi tính từ kết quả sau hiệu chỉnh)")
        for i, (k, src, lbl) in enumerate(keys, 1):
            r = first + i
            raw = r_dew if src == "dew" else r_pres
            _retry(ws.Rows(f"{raw}:{raw}").Copy)
            ws.Rows(f"{r}:{r}").PasteSpecial(-4122)            # xlPasteFormats
            ws.Range(f"F{r}:J{r}").UnMerge()
            ws.Range(f"F{r}:J{r}").ClearContents()
            ws.Cells(r, 2).Value = lbl
            for col in ("C", "D"):
                v = ws.Range(f"{col}{raw}").Value
                if isinstance(v, (int, float)):
                    ws.Range(f"{col}{r}").Formula = _hc_formula(f"{col}{raw}", float(v), cals[k], cal_rows[k])
                    # như số đọc (2 số lẻ); NumberFormatLocal theo dấu thập phân của máy (VN: ',' — gán "0.00" ra "0\.00")
                    ws.Range(f"{col}{r}").NumberFormatLocal = f"0{dec_sep}00"
            ws.Range(f"E{r}").Formula = f"=AVERAGE(C{r}:D{r})"
        xl.CutCopyMode = False
        xl.Calculation = -4105
        ws.Activate()
        _retry(xl.CalculateFull)
        wb.Save()
        print(f"[khinen_calc] {os.path.basename(path)}: thêm phần sau hiệu chỉnh ({', '.join(k for k, *_ in keys)})")
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        wb = None
        xl.Quit()
        del xl
        gc.collect()
    return path


def build_compressed_air_workbook(
    point_data: Dict[str, Any],
    output_path: str,
    template_path: Optional[str] = None
) -> str:
    """
    Builds a complete, standardized QUATEST 3 Compressed Air calculation workbook (.xlsx)
    for a given sampling point (e.g. KN1, KN2), matching 100% of formatting, headers,
    media photos, HumCalc rendered screenshots, and ISO 8573-1 calculations.
    """
    template_file = template_path or DEFAULT_TEMPLATE_XLSX
    if not os.path.isfile(template_file):
        raise FileNotFoundError(f"Compressed air Excel template not found at {template_file}")

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    shutil.copy2(template_file, output_path)

    # 1. Extract inputs
    code = point_data.get("code", "")
    point_label = point_data.get("point_label", "KN1")
    cert_date = point_data.get("cert_date", "")
    loc_name = point_data.get("location_name", "")
    loc_symbol = point_data.get("location_symbol", point_label)

    dew_in = point_data.get("dewpoint", {})
    t1_s = float(dew_in.get("dewp_start", -40.0))
    t1_e = float(dew_in.get("dewp_end", -40.0))
    tk_s = float(dew_in.get("temp_start", 20.0))
    tk_e = float(dew_in.get("temp_end", 20.0))
    p_s = float(dew_in.get("pres_start", 7.0))
    p_e = float(dew_in.get("pres_end", 7.0))
    time_start = dew_in.get("time_start", "")
    time_end = dew_in.get("time_end", "")

    # Hiệu chỉnh theo giấy hiệu chuẩn DP510 (point_data["cals"] = khinen_cals(...); rỗng khi công tắc tắt)
    from .cleanroom_calc import apply_cal
    cals = {k: v for k, v in (point_data.get("cals") or {}).items() if v is not None and getattr(v, "points", None)}
    c_dew, c_pres = cals.get("dewpoint"), cals.get("pressure")
    t1c_s, t1c_e = apply_cal(t1_s, c_dew), apply_cal(t1_e, c_dew)
    pc_s, pc_e = apply_cal(p_s, c_pres), apply_cal(p_e, c_pres)

    # Calculate exact CS dewpoint (từ số sau hiệu chỉnh nếu có)
    dew_res = calc_dewpoint_cs(t1c_s, t1c_e, tk_s, tk_e, pc_s, pc_e)
    td2_conv = dew_res["cs_converted_7bar"]["Dewpoint_T2_Ctd"]
    abs_hum_water = dew_res["cs_measured_condition"]["AbsoluteHumidity_mg_m3"]

    # 2. Render HumCalc screenshot
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(output_path)), "_temp_media")
    os.makedirs(temp_dir, exist_ok=True)
    rendered_humcalc_png = os.path.join(temp_dir, f"{point_label}_humcalc_rendered.png")
    humcalc_vals = generate_humcalc_values(dew_res)
    render_humcalc_screenshot(DEFAULT_HUMCALC_BASE, rendered_humcalc_png, humcalc_vals)

    # 3. Prepare photos & media replacements
    replacements = {"xl/media/image3.png": rendered_humcalc_png}

    photo_start = dew_in.get("photo_start")
    if photo_start and os.path.isfile(photo_start):
        replacements["xl/media/image2.jpeg"] = photo_start

    photo_end = dew_in.get("photo_end")
    if photo_end and os.path.isfile(photo_end):
        replacements["xl/media/image4.jpeg"] = photo_end

    loc_photo = point_data.get("location_photo")
    if loc_photo and os.path.isfile(loc_photo):
        # Check if template has image1.jpeg or image1.png
        with zipfile.ZipFile(output_path, "r") as z:
            names = z.namelist()
            has_jpeg = "xl/media/image1.jpeg" in names
        if has_jpeg and not loc_photo.lower().endswith((".jpg", ".jpeg")):
            conv_jpeg = os.path.join(temp_dir, f"{point_label}_loc.jpeg")
            im = Image.open(loc_photo).convert("RGB")
            im.save(conv_jpeg, "JPEG", quality=95)
            replacements["xl/media/image1.jpeg"] = conv_jpeg
        elif has_jpeg:
            replacements["xl/media/image1.jpeg"] = loc_photo
        else:
            replacements["xl/media/image1.png"] = loc_photo

    update_zip_media(output_path, replacements)

    # 4. Populate Excel cells via openpyxl
    wb = openpyxl.load_workbook(output_path)
    fill_workbook_tags(wb, point_data)  # ô thông tin hồ sơ trong template là TAG (Core/Templates/TAGS.md)
    ws = wb["Tinh_Toan_Khi_Co_Ap_V6"] if "Tinh_Toan_Khi_Co_Ap_V6" in wb.sheetnames else wb.active

    # Location info
    if loc_name:
        ws["H2"].value = loc_name
    if loc_symbol:
        ws["H3"].value = loc_symbol

    # Dewpoint table (rows 87-89)
    ws["C87"].value = t1_s
    ws["D87"].value = t1_e
    ws["C88"].value = tk_s
    ws["D88"].value = tk_e
    ws["C89"].value = p_s
    ws["D89"].value = p_e

    # Times
    if time_start:
        ws["C92"].value = str(time_start)
    if time_end:
        ws["C103"].value = str(time_end)

    # Converted Dew point & Water content
    ws["G87"].value = td2_conv
    ws["H8"].value = abs_hum_water

    # Dust particles inputs
    dust_in = point_data.get("dust", {})
    raw_channels = dust_in.get("raw_channels", [])
    device_name = dust_in.get("device", "Lasair III 110")
    # ISO 8573-1 xếp cấp theo DẢI cỡ hạt -> ô C54/C63 nhận số VI PHÂN (Lasair cột Δ, AeroTrak 9350 cột Diff)
    # (người dùng chốt 11/11/2026, AGENTS mục 8C). `mode` = "diff" | "cumul" (Σ / Cumul -> tự đổi sang Δ).
    if raw_channels:
        from .particle_mode import norm_mode, to_diff, DIFF, CUMUL
        mode = norm_mode(dust_in.get("mode"))
        if mode == CUMUL:
            print(f"[INFO] [Particle Mode] {device_name}: số cộng dồn Σ -> đổi sang Δ cho ISO 8573-1")
            raw_channels = [round(x) for x in to_diff(raw_channels)]
        elif mode != DIFF:
            print(f"[WARNING] [Particle Mode] {device_name}: chưa khai dust['mode'] — khí nén phải dùng cột "
                  f"Δ (Lasair) / Diff (AeroTrak 9350); kiểm tra lại phiếu in: {raw_channels}")
    ws["C52"].value = device_name

    if "Lasair" in device_name and len(raw_channels) >= 8:
        for idx, val in enumerate(raw_channels[:8]):
            ws.cell(54, 3 + idx).value = val
        ws["C55"].value = dust_in.get("dilution_factor", 1.0)
    elif len(raw_channels) >= 6:
        for idx, val in enumerate(raw_channels[:6]):
            ws.cell(63, 3 + idx).value = val
        ws["C64"].value = dust_in.get("dilution_factor", 1.0)

    # Total Oil inputs (rows 43-46)
    oil_in = point_data.get("oil", {})
    if "flow_aerosol" in oil_in:
        ws["C43"].value = float(oil_in["flow_aerosol"])
    if "flow_vapor" in oil_in:
        ws["D43"].value = float(oil_in["flow_vapor"])
    if "time_aerosol" in oil_in:
        ws["C44"].value = float(oil_in["time_aerosol"])
    if "time_vapor" in oil_in:
        ws["D44"].value = float(oil_in["time_vapor"])
    if "kqtn_aerosol" in oil_in:
        ws["C46"].value = oil_in["kqtn_aerosol"]
    if "kqtn_vapor" in oil_in:
        ws["D46"].value = oil_in["kqtn_vapor"]
    # Số phiếu KQTN, vd "MTTN26019695.01" -> I43 "MTTN" + J43 "26019695.01"; "HDTN26013537" -> I44/J44
    for key, row in (("kqtn_no_vapor", 43), ("kqtn_no_aerosol", 44)):
        m = re.match(r"([A-Z]+)\s*(.+)", str(oil_in.get(key) or ""))
        if m:
            ws[f"I{row}"].value, ws[f"J{row}"].value = m.group(1), m.group(2)

    # Customer internal limits
    if "customer_limit_05" in dust_in:
        ws["I6"].value = dust_in["customer_limit_05"]
    if "water_limit" in point_data:
        ws["I8"].value = point_data["water_limit"]
    if "customer_limit" in oil_in:
        ws["I9"].value = oil_in["customer_limit"]

    # Fix formula bugs (K6, AeroTrak rows 67-68, Feller rows 82-83)
    if "E67" in str(ws["K6"].value or ""):
        ws["K6"].value = str(ws["K6"].value).replace("E67", "D6")

    ws["C67"].value = '=IF(ISNUMBER(C63),C63*C64,"-")'
    ws["D67"].value = '=IF(ISNUMBER(D63),D63*C64,"-")'
    ws["E67"].value = '=IF(ISNUMBER(E63),SUM(E63:G63)*C64,"-")'
    ws["F67"].value = '=IF(ISNUMBER(H63),H63*C64,"-")'
    for col_let in ("C", "D", "E", "F"):
        ws[f"{col_let}68"].value = f'=IF(ISNUMBER({col_let}67),ROUND({col_let}67*$C$38,0),"-")'

    for col_let in ("C", "D", "E", "F", "G"):
        ws[f"{col_let}82"].value = (
            f'=IF(OR({col_let}81="-",{col_let}81=""),"-",'
            f'IF(OR({col_let}81=0,{col_let}81="< 1",{col_let}81="KPH"),0,$C$77*(-LN(1-({col_let}81/$C$77)))))'
        )
        ws[f"{col_let}83"].value = (
            f'=IF({col_let}82="-","-",'
            f'IF(OR({col_let}81="< 1",{col_let}81="KPH"),"Không có chỉ thị",ROUND(IF({col_let}82=0,0,({col_let}82/($C$78/1000))),0)))'
        )

    # Set PageSetup headers via openpyxl
    lh_text = f'&"Times New Roman,Regular"&13Số: {code} - {point_label}' if code else f'&"Times New Roman,Regular"&13Số: - {point_label}'
    rh_text = f'&"Times New Roman,Regular"&13 {cert_date}' if cert_date else ""
    for sheet in wb.worksheets:
        sheet.oddHeader.left.text = lh_text
        sheet.oddHeader.right.text = rh_text

    place_box_images(ws, [("B17:D34", loc_photo), ("B91:B100", photo_start), ("B102:B111", photo_end),
                          ("D92:J111", rendered_humcalc_png)], temp_dir)

    thin_outside_borders(wb)  # viền ngoài thường, không dùng viền dày (người dùng chốt 30/10/2026)
    wb.save(output_path)
    wb.close()

    # 5. Evaluate full workbook & set PageSetup via Excel COM if available
    try:
        import win32com.client
        import gc
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            abs_out = os.path.abspath(output_path)
            x_wb = excel.Workbooks.Open(abs_out)
            for x_ws in x_wb.Sheets:
                ps = x_ws.PageSetup
                ps.LeftHeader = lh_text
                ps.RightHeader = rh_text
            x_wb.Application.CalculateFull()
            x_wb.Save()
            x_wb.Close()
        finally:
            ps = None
            x_ws = None
            x_wb = None
            excel.Quit()
            del excel
            gc.collect()
    except Exception as e:
        print(f"[build_compressed_air_workbook] Excel COM warning: {e}")

    if cals:
        finalize_khinen_workbook(output_path, cals)

    # Clean up temp media directory
    shutil.rmtree(temp_dir, ignore_errors=True)
    print(f"[build_compressed_air_workbook] Successfully created {output_path}")
    return output_path


# ==============================================================================
# 4. WORKBOOK PARSER (TinhToanKhiNen.xlsx -> Structured Data)
# ==============================================================================
def parse_compressed_air_workbook(xlsx_path: str) -> Dict[str, Any]:
    """
    Parses an existing TinhToanKhiNen.xlsx file, runs independent ISO 8573-1 calculations,
    and returns a clean structured dictionary of results for reporting.
    """
    audit_rep = audit_and_calc_from_excel(xlsx_path, fix_bugs=False)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Tinh_Toan_Khi_Co_Ap_V6"] if "Tinh_Toan_Khi_Co_Ap_V6" in wb.sheetnames else wb.active

    loc_name = str(ws["H2"].value or "").strip()
    loc_symbol = str(ws["H3"].value or "").strip()

    # Normalize location symbol if generic "KN" or empty
    if not loc_symbol or loc_symbol.upper() == "KN":
        m_fn = re.search(r"[-_\s]?(KN\d+)[-_\s]?", os.path.basename(xlsx_path), re.IGNORECASE)
        if m_fn:
            loc_symbol = m_fn.group(1).upper()

    from .iso8573_cal import dew_rows
    _off = dew_rows(ws)["off"] + _hc_rows_below(ws)   # phiếu có phần sau hiệu chỉnh: ô giờ đo dời xuống
    t_start = str(ws[f"C{92 + _off}"].value or "").strip()
    t_end = str(ws[f"C{103 + _off}"].value or "").strip()

    # Extract time interval e.g. "15:02 – 15:21"
    m1 = re.search(r"(\d+h\d+|\d+:\d+)", t_start)
    m2 = re.search(r"(\d+h\d+|\d+:\d+)", t_end)
    t1_clean = m1.group(1).replace("h", ":") if m1 else ""
    t2_clean = m2.group(1).replace("h", ":") if m2 else ""
    time_str = f"{t1_clean} – {t2_clean}" if (t1_clean and t2_clean) else (t1_clean or t2_clean)

    # Dew point results
    dp = audit_rep["dewpoint"]
    tkhi_avg = dp["averages_exact"]["Tkhi_c"]
    p_avg = dp["averages_exact"]["P_bar"]
    t1_meas = dp["cs_measured_condition"]["Dewpoint_Ctd"]
    ahum_meas = dp["cs_measured_condition"]["AbsoluteHumidity_tbkq_int"]
    td2_7bar = dp["cs_converted_7bar"]["Dewpoint_T2_tbkq_1dec"]
    ahum_7bar = dp["cs_converted_7bar"]["AbsoluteHumidity_mg_m3"]
    dp_class = dp["class_iso8573"]

    # Particles results
    part = audit_rep["particles"]
    d01_05 = part["d_0.1_0.5_m3"]
    d05_10 = part["d_0.5_1.0_m3"]
    d10_50 = part["d_1.0_5.0_m3"]
    d_gt_50 = part["d_gt_5.0_m3"]
    d_gt_05_total = part["d_gt_0.5_total_m3"]
    part_class_vi = part["class_tbkq_vi"]
    part_class_en = part["class_tbkq_en"]

    # Oil results
    oil = audit_rep["total_oil"]
    oil_val_vi = oil["result_vi"]
    oil_val_en = oil["result_en"]
    oil_class = oil["class_iso8573"]

    # Customer internal limits
    dust_req = str(ws["I6"].value or "").strip()
    water_req = str(ws["I8"].value or "").strip()
    oil_req = str(ws["I9"].value or "").strip()

    # Extract location image if available
    loc_image_path = None
    try:
        with zipfile.ZipFile(xlsx_path, "r") as z:
            for item in z.namelist():
                if "xl/media/image1" in item:
                    ext = os.path.splitext(item)[1]
                    import tempfile
                    temp_img_dir = os.path.join(tempfile.gettempdir(), "quatest3_ca_media")
                    os.makedirs(temp_img_dir, exist_ok=True)
                    out_img = os.path.join(temp_img_dir, f"{loc_symbol}_extracted{ext}")
                    with open(out_img, "wb") as f_out:
                        f_out.write(z.read(item))
                    loc_image_path = out_img
                    break
    except Exception as _err:
        print(f"[Warning] khinen_calc.py:475 {type(_err).__name__}: {_err}")

    wb.close()

    return {
        "file": xlsx_path,
        "location_name": loc_name,
        "location_symbol": loc_symbol,
        "time_str": time_str,
        "location_photo": loc_image_path,
        "dewpoint": {
            "Tkhi": tkhi_avg,
            "P": p_avg,
            "T1": t1_meas,
            "Ahum": ahum_meas,
            "Td_7bar": td2_7bar,
            "Ahum_7bar": ahum_7bar,
            "class": dp_class,
            "requirement": water_req,
        },
        "particles": {
            "d_01_05": d01_05,
            "d_05_10": d05_10,
            "d_10_50": d10_50,
            "d_gt_50": d_gt_50,
            "d_gt_05_total": d_gt_05_total,
            "class_vi": part_class_vi,
            "class_en": part_class_en,
            "requirement": dust_req,
        },
        "total_oil": {
            "val_vi": oil_val_vi,
            "val_en": oil_val_en,
            "class": oil_class,
            "requirement": oil_req,
        },
        "microbiology": audit_rep.get("microbiology", {}),
    }


# ==============================================================================
# 5. DOSSIER AGGREGATION & REPORT PIPELINE
# ==============================================================================
def create_dossier_from_compressed_air_workbooks(
    xlsx_paths: List[str],
    code: str,
    cert_date: str = "20/10/2026",
    sample_date: str = "06/10/2026",
    sample_time_vn: str = "Ngày 06/10/2026 từ 08:00 đến 17:00",
    sample_time_en: str = "From 06/10/2026",
    sub_no: int = 2,
    client: Optional[ClientInfo] = None,
    personnel: Optional[PersonnelInfo] = None,
    equipments: Optional[List[Equipment]] = None,
    scope_vn: str = "Quan trắc chất lượng hệ thống khí nén theo tiêu chuẩn ISO 8573-1:2010.",
    scope_en: str = "Monitoring compressed air quality according to ISO 8573-1:2010.",
    standards_vn: str = "ISO 8573-1:2010; ISO 8573-2:2018; ISO 8573-7:2014; QUATEST 3 3030:2025",
    standards_en: str = "ISO 8573-1:2010; ISO 8573-2:2018; ISO 8573-7:2014; QUATEST 3 3030:2025",
    basis_vn: Optional[str] = None,
    basis_en: Optional[str] = None,
) -> Dossier:
    """
    Aggregates one or multiple TinhToanKhiNen.xlsx workbooks (e.g. KN1, KN2, KN3)
    into a standardized Dossier object ready for DossierRunner (TBKQ VN/EN, BCKQ, PDF).
    """
    parsed_points = [parse_compressed_air_workbook(p) for p in xlsx_paths]

    # Default Client
    if client is None:
        client = ClientInfo(
            name_vn="Công ty TNHH Nextern Việt Nam",
            name_en="Nextern Vietnam Limited Liability Company",
            address_vn="Đường số 7, Khu công nghiệp VSIP II-A, thị xã Tân Uyên, tỉnh Bình Dương",
            address_en="No. 7 street, VSIP II-A Industrial Park, Tan Uyen Town, Binh Duong Province"
        )

    # Default Personnel
    if personnel is None:
        personnel = PersonnelInfo(
            executor="Phát",
            sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
            sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
            sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
            sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
        )

    # Default Equipment
    if equipments is None:
        equipments = [
            eq("lasair_iii_110"),
            eq("cs_ds400")
        ]

    # Reconcile symbols across points:
    # If any point has KN2, KN3... but another point has 'KN', normalize 'KN' to 'KN1'
    has_numbered_kn = any(re.search(r"KN[1-9]\d*", (p.get("location_symbol") or "").upper()) for p in parsed_points)
    for idx, p in enumerate(parsed_points):
        sym = (p.get("location_symbol") or "").strip().upper()
        if not sym or sym == "KN":
            m_fn = re.search(r"KN(\d+)", os.path.basename(p.get("file", "")), re.IGNORECASE)
            if m_fn:
                sym = f"KN{m_fn.group(1)}"
            elif has_numbered_kn or len(parsed_points) > 1:
                sym = f"KN{idx + 1}"
            else:
                sym = "KN"
        p["location_symbol"] = sym

    # Build Monitoring Locations & Photos
    locations = []
    photos = []
    oil_rows = []
    dp_rows = []
    part_rows = []

    for idx, p in enumerate(parsed_points):
        sym = p["location_symbol"] or f"KN{idx + 1}"
        name = p["location_name"] or f"Đường ống khí nén vị trí {sym}"
        locations.append(
            MonitoringLocation(
                name_vn=name,
                name_en=f"Pneumatic pipeline at {sym}",
                points_summary_vn=[sym],
                points_summary_en=[sym],
            )
        )

        if p.get("location_photo") and os.path.isfile(p["location_photo"]):
            photos.append({
                "path": p["location_photo"],
                "caption_vn": f"Vị trí lấy mẫu {sym}",
                "caption_en": f"Sampling location {sym}"
            })

        # Oil row
        oil_rows.append({
            "loc": sym,
            "val": p["total_oil"]["val_vi"],
            "class": p["total_oil"]["class"]
        })

        # Dewpoint row: loc with time, values: [Tkhi, P, T1, T2_7bar, Ahum_7bar, Class]
        t_label = f"{sym}\n({p['time_str']})" if p["time_str"] else sym
        t1_val = p['dewpoint'].get('T1', p['dewpoint'].get('Ahum', 0.0))
        dp_vals = [
            f"{p['dewpoint']['Tkhi']:.1f}".replace(".", ","),
            f"{p['dewpoint']['P']:.1f}".replace(".", ","),
            f"{t1_val:.1f}".replace(".", ",") if isinstance(t1_val, (int, float)) else str(t1_val),
            f"{p['dewpoint']['Td_7bar']:.1f}".replace(".", ","),
            f"{p['dewpoint']['Ahum_7bar']:.2f}".replace(".", ","),
            p['dewpoint']['class']
        ]
        dp_rows.append({"loc": t_label, "values": dp_vals})

        # Particle row: values: [0.1-0.5, 0.5-1.0, 1.0-5.0, Class]
        def fmt_count(val):
            return f"{val:,}".replace(",", " ") if isinstance(val, (int, float)) else str(val)

        p_vals = [
            fmt_count(p["particles"]["d_01_05"]),
            fmt_count(p["particles"]["d_05_10"]),
            fmt_count(p["particles"]["d_10_50"]),
            p["particles"]["class_vi"]
        ]
        part_rows.append({
            "loc": sym,
            "values": p_vals,
            "d_gt_50": p["particles"]["d_gt_50"],
            "class_vi": p["particles"]["class_vi"],
            "class_en": p["particles"]["class_en"]
        })

    results_data = {
        "oil": {"rows": oil_rows},
        "dewpoint": {"rows": dp_rows},
        "particles": {"rows": part_rows},
        "parsed_points": parsed_points
    }

    if basis_vn is None:
        basis_vn = f"ISO 8573-1:2010 và Yêu cầu kỹ thuật do {client.name_vn} cung cấp."
    if basis_en is None:
        basis_en = f"ISO 8573-1:2010 and Technical requirements provided by {client.name_en}."

    dossier = Dossier(
        code=code,
        cert_date=cert_date,
        sample_date=sample_date,
        sample_time_vn=sample_time_vn,
        sample_time_en=sample_time_en,
        sub_no=sub_no,
        is_multi_cert=True,
        dossier_type="compressed_air",
        scope_vn=scope_vn,
        scope_en=scope_en,
        location_desc_vn=f"{client.name_vn} - {client.address_vn}",
        location_desc_en=f"{client.name_en} - {client.address_en}",
        standards_vn=standards_vn,
        standards_en=standards_en,
        basis_vn=basis_vn,
        basis_en=basis_en,
        client=client,
        personnel=personnel,
        equipments=equipments,
        locations=locations,
        photos=photos,
        results_data=results_data
    )
    return dossier
