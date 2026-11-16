# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D134804 - Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản (02/11/2026).

2 chứng thư:
  /1  Tủ an toàn sinh học cấp I  ESCO AC1-6A1, S/N 2003-4060 (QA.E19)
      -> BSCHandler, bỏ downflow + lọc khí cấp (tủ cấp I không có dòng khí cấp), inflow mức cấp I 0,7 - 1,0 m/s
  /2  Tủ an toàn sinh học cấp II ESCO AC2-4E8, S/N 2019-140833 (QA.E54) -> BSCHandler đầy đủ (không đo bụi)

Quyết định của người dùng (02/11/2026):
  - Tủ 1: inflow lấy theo Kanomax đo tại mép trên cửa (TB 0,82 m/s), mức EN 12469 cho tủ cấp I 0,7 - 1,0 m/s
    (phiếu có thêm DIM 692,2 m³/h, mở 220 x 1850 mm -> 0,47 m/s, không dùng).
  - Tủ 1, cường độ sáng: ô đầu (viết đè) đọc 746,8 lux.
  - Tủ 2 inflow: k = 0,457 theo nhãn ESCO trên tủ ("Multiply the SM average with correction factor of 0.457"),
    1,075 x 0,457 = 0,49 m/s (phiếu ghi tay "0,41" không dùng).
Suy luận của AI (cần người dùng xác nhận):
  - Tủ 2 downflow hàng 3, điểm 5 đọc 0,31 (có thể 0,34; TB vẫn 0,36 m/s).

Số liệu đọc từ phiếu ghi chép (Input/134804/20261102135959707.pdf) và phiếu in nhiệt ATI 2i
(Input/134804/20261102140034478.pdf).

Chạy:  python build_134804.py            (DOCX + XLSX, không PDF)
       python build_134804.py --pdf      (thêm xuất PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: (network_dir để trống; chép lên Z bằng robocopy theo AGENTS mục 0).
"""

import os
import sys
import argparse

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)

import openpyxl
from openpyxl.styles import Alignment, Border

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.bsc import BSCHandler
from Core.Tools.bsc_calc import build_bsc_calculation_workbook
from Core.Tools.excel_borders import strike_box
from Core.Tools.noise_corr import noise_correction_value
from Core.Tools.qa_check import check_single_docx
from Core.Builders.tag_context import build_tem_tu_sheet
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D134804"
REG_DATE = "23/10/2026"
SAMPLE_DATE = "02/11/2026"
import datetime as _dt
CERT_DATE = _dt.date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày hiện tại, tự cập nhật mỗi lần dựng (người dùng chốt 06/11/2026, AGENTS mục 2)
START, END = "09:00", "11:30"
INPUT_DIR = os.path.join(WORKSPACE, "Input", "134804")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photo", "Use")

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản",
    name_en="VISSAN Joint Stock Company",
    address_vn="420 Nơ Trang Long, phường Bình Lợi Trung, Thành phố Hồ Chí Minh",
    address_en="420 No Trang Long Street, Binh Loi Trung Ward, Ho Chi Minh City",
)
LOCATION_VN = "Phòng Quản lý Chất lượng - 420 Nơ Trang Long, phường Bình Lợi Trung, Thành phố Hồ Chí Minh"
LOCATION_EN = "Quality Management Department - 420 No Trang Long Street, Binh Loi Trung Ward, Ho Chi Minh City"

EQ = {
    "kanomax": eq("kanomax_6501"),
    "ati": eq("ati_6d_2i"),
    "smoke": eq("smoke"),
    "testo": eq("testo_440"),
    "rion": eq("rion_nl53"),
    "uvc": eq("extech_sdl470"),
}
EQUIPMENTS = [EQ["kanomax"], EQ["ati"], EQ["smoke"], EQ["testo"], EQ["rion"], EQ["uvc"]]


def _photos(sub: str):
    d = os.path.join(PHOTO_DIR, sub)
    if not os.path.isdir(d):
        return []
    return [{"path": os.path.join(d, f), "caption_vn": "", "caption_en": ""}
            for f in sorted(os.listdir(d)) if f.lower().endswith((".jpg", ".jpeg", ".png"))]


def _avg(grid):
    vals = [v for row in grid for v in row]
    return sum(vals) / len(vals)


def _base_kwargs():
    # Số phụ /1, /2 do assign_cert_numbers() tự đánh theo thứ tự báo cáo trong main()
    return dict(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=True,
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        client=CLIENT, personnel=PersonnelInfo(executor="Phát"),
        input_dir=INPUT_DIR, tbkq_dir=os.path.join(INPUT_DIR, "TBKQ"), output_dir="", network_dir="",
        dossier_type="bsc",
        standards_vn="EN 12469:2000", standards_en="EN 12469:2000",
        basis_vn="EN 12469:2000", basis_en="EN 12469:2000",
        equipments=EQUIPMENTS,
        assessment={"basis_vn": "EN 12469:2000"},
    )


def _noise(amb: float, op: float) -> dict:
    corr = noise_correction_value(amb, op)
    return {"ambient": amb, "operating": op, "corrected_dba": round(op - corr, 1)}


def _uvc(readings):
    avg = sum(readings) / len(readings)
    return {"readings": readings, "avg": round(avg, 1), "avg_rounded": int(round(avg))}


# ---------------------------------------------------------------------------
# /1 - Tủ an toàn sinh học cấp I ESCO AC1-6A1
# ---------------------------------------------------------------------------
T1_INFLOW = [0.83, 0.90, 0.82, 0.81, 0.74]          # Kanomax, mép trên cửa kính, khoảng mở 220 mm, rộng 1850 mm
T1_INFLOW_RANGE = (0.70, 1.00)                      # EN 12469:2000, tủ cấp I
T1_LUX = [[746.8, 850.1, 659.7, 816.1, 714.3],
          [840.5, 840.8, 857.8, 826.8, 737.2],
          [610.6, 750.2, 855.4, 720.2, 638.8]]
T1_UVC = [132, 191, 205, 132, 69]
# Hình thái dòng khí tủ cấp I — theo mức ghi trên phiếu ghi chép (kết quả "OK")
T1_AF_ROWS = [
    {"param_vn": "Khu vực không gian làm việc của tủ",
     "param_en": "Inside cabinet workspace",
     "result_vn": "- Dòng khói đi vào khe hút khí bên trong tủ;\n- Khói không thoát ra ngoài tủ.",
     "result_en": "- Smoke was drawn into the internal exhaust slots;\n- No smoke escaped.",
     "limit_vn": "- Dòng khói đi vào khe hút khí bên trong tủ;\n- Khói không thoát ra ngoài tủ. ⁽¹⁾",
     "limit_en": "- Smoke is drawn into the internal exhaust slots;\n- No smoke escapes. ⁽¹⁾",
     "height": 1000},
    {"param_vn": "Khu vực dọc theo chu vi phía bên ngoài của vùng mở cửa tủ",
     "param_en": "Area along the outer perimeter of the work opening",
     "result_vn": "- Dòng khói hướng vào trong tủ, không thoát ra ngoài.",
     "result_en": "- Smoke was drawn into the cabinet, no smoke escaped.",
     "limit_vn": "- Dòng khói hướng vào trong tủ, không thoát ra ngoài. ⁽¹⁾",
     "limit_en": "- Smoke is drawn into the cabinet, no smoke escapes. ⁽¹⁾",
     "height": 760},
]


def tu1_dossier() -> Dossier:
    v_in = sum(T1_INFLOW) / len(T1_INFLOW)
    lux_avg = _avg(T1_LUX)
    rd = {
        "start_time": START, "end_time": END,
        "file_label": "BSC", "bckq_label": "bsc",
        "cabinet": {
            "cabinet_name": "Tủ an toàn sinh học cấp 1",
            "cabinet_name_en": "Biosafety cabinet class I",
            "cap_tu": "I",
            "brand": "ESCO", "model": "AC1-6A1", "serial": "2003-4060", "device_id": "QA.E19",
            "seal_item_name": "“Tủ an toàn sinh học cấp 1”",
        },
        "inflow": {"method": 2, "method_name": "Kanomax Climomaster 6501", "readings": T1_INFLOW,
                   "velocity": round(v_in, 2), "limit_range": T1_INFLOW_RANGE,
                   "width_mm": 1850, "height_mm": 220},
        "downflow": {"skip": True},
        "lux": {"readings": T1_LUX, "avg": lux_avg, "avg_rounded": int(round(lux_avg))},
        "noise": dict(_noise(53.9, 63.8), file_ambient="0182", file_operating="0180"),
        "hepa": {"downflow_pct": None, "skip_downflow": True, "exhaust_pct": 0.0018, "aerosol_conc": 28,
                 "exhaust_gasket_pct": 0.0018, "exhaust_media_pct": 0.0005},
        "airflow_pattern": {"rows": [dict(r, assess_vn="Đạt", assess_en="Pass") for r in T1_AF_ROWS]},
        "uvc": _uvc(T1_UVC),
        "tasks_vn": ["Đo tốc độ dòng khí qua cửa làm việc;", "Đo cường độ sáng bên trong tủ;",
                     "Đo độ ồn của tủ khi vận hành;", "Kiểm tra độ rò rỉ của lọc HEPA khí thải;",
                     "Kiểm tra hình thái dòng khí của tủ;", "Đo cường độ tia UV dải C trong khoang làm việc."],
    }
    return Dossier(
        **_base_kwargs(),
        procedure_vn="QUATEST 3 3030:2025, N2.TC.BSC.01",
        procedure_en="QUATEST 3 3030:2025, N2.TC.BSC.01",
        scope_vn="Giám định tủ an toàn sinh học cấp 1",
        scope_en="Inspection of Biosafety cabinet class I",
        locations=[MonitoringLocation(name_vn="Tủ an toàn sinh học cấp 1 ESCO AC1-6A1",
                                      name_en="Biosafety cabinet class I ESCO AC1-6A1")],
        photos=_photos("1"),
        results_data=rd,
    )


# ---------------------------------------------------------------------------
# /2 - Tủ an toàn sinh học cấp II ESCO AC2-4E8
# ---------------------------------------------------------------------------
T2_INFLOW = [1.11, 1.12, 1.09, 1.07, 1.08, 1.03, 1.02, 1.08]   # Kanomax, mép trên cửa kính, 8 điểm
T2_K = 0.457                                                 # nhãn ESCO trên tủ (người dùng chốt 02/11/2026)
T2_DOWNFLOW = [[0.36, 0.39, 0.40, 0.39, 0.39, 0.40, 0.36],
               [0.35, 0.35, 0.34, 0.35, 0.35, 0.33, 0.34],
               [0.36, 0.36, 0.35, 0.36, 0.31, 0.34, 0.35]]
T2_LUX = [[1024.6, 1342.8, 1363.3, 1112.9],
          [991.0, 1341.9, 1360.2, 1096.0],
          [816.6, 988.1, 1135.6, 935.5]]
T2_UVC = [97, 160, 147, 95]


def tu2_dossier() -> Dossier:
    v_in = sum(T2_INFLOW) / len(T2_INFLOW) * T2_K
    df_avg = _avg(T2_DOWNFLOW)
    lux_avg = _avg(T2_LUX)
    rd = {
        "start_time": START, "end_time": END,
        "file_label": "BSC", "bckq_label": "bsc",
        "cabinet": {
            "cabinet_name": "Tủ an toàn sinh học cấp 2",
            "cabinet_name_en": "Biosafety cabinet class II",
            "cap_tu": "II",
            "brand": "ESCO", "model": "AC2-4E8", "serial": "2019-140833", "device_id": "QA.E54",
            "seal_item_name": "“Tủ an toàn sinh học cấp 2”",
        },
        "inflow": {"method": 2, "method_name": "Kanomax Climomaster 6501", "readings": T2_INFLOW,
                   "k_factor": T2_K, "velocity": round(v_in, 2), "width_mm": 1220, "height_mm": 135},
        "downflow": {"readings": T2_DOWNFLOW, "avg": round(df_avg, 2), "tolerance_pct": 20,
                     "width_mm": 1220, "depth_mm": 560},
        "lux": {"readings": T2_LUX, "avg": lux_avg, "avg_rounded": int(round(lux_avg))},
        "noise": dict(_noise(53.9, 60.5), file_ambient="0182", file_operating="0181"),
        "hepa": {"downflow_pct": 0.0032, "exhaust_pct": 0.0082, "aerosol_conc": 27,
                 "downflow_gasket_pct": 0.0018, "downflow_media_pct": 0.0032,
                 "exhaust_gasket_pct": 0.0082, "exhaust_media_pct": 0.0060},
        "airflow_pattern": {},
        "uvc": _uvc(T2_UVC),
        "tasks_vn": ["Đo tốc độ dòng khí qua cửa làm việc;", "Đo tốc độ dòng khí cấp qua lọc HEPA;",
                     "Đo cường độ sáng bên trong tủ;", "Đo độ ồn của tủ khi vận hành;",
                     "Kiểm tra độ rò rỉ của các lọc HEPA;", "Kiểm tra hình thái dòng khí của tủ;",
                     "Đo cường độ tia UV dải C trong khoang làm việc."],
    }
    return Dossier(
        **_base_kwargs(),
        procedure_vn="QUATEST 3 3030:2025, N2.TC.BSC.02",
        procedure_en="QUATEST 3 3030:2025, N2.TC.BSC.02",
        scope_vn="Giám định tủ an toàn sinh học cấp 2",
        scope_en="Inspection of Biosafety cabinet class II",
        locations=[MonitoringLocation(name_vn="Tủ an toàn sinh học cấp 2 ESCO AC2-4E8",
                                      name_en="Biosafety cabinet class II ESCO AC2-4E8")],
        photos=_photos("2"),
        results_data=rd,
    )


# ---------------------------------------------------------------------------
# Phiếu tính toán Excel
# ---------------------------------------------------------------------------
def _calc_common(d: Dossier) -> dict:
    c = dict(d.results_data)
    c.update({
        "code": d.code, "sub_no": d.sub_no, "reg_date": d.reg_date, "sample_date": d.sample_date,
        "cert_date": d.cert_date, "start_time": START, "end_time": END, "photos": d.photos, "input_dir": d.input_dir,
        "client_name_vn": d.client.name_vn, "client_name_en": d.client.name_en,
        "client_address_vn": d.client.address_vn, "client_address_en": d.client.address_en,
        "scope_vn": d.scope_vn, "scope_en": d.scope_en,
        "location_vn": d.location_desc_vn, "location_en": d.location_desc_en,
        "leader_name_vn": "Lê Đăng Huyền", "gdv_name_vn": "Nguyễn Xuân Nam",
        "leader_name_en": "Le Dang Huyen", "gdv_name_en": "Nguyen Xuan Nam",
        "equipments": [e.name_vn for e in d.equipments], "equipments_en": [e.name_en for e in d.equipments],
    })
    cab = d.results_data.get("cabinet", {})
    for k in ("cabinet_name", "brand", "model", "serial", "device_id"):
        c[k] = cab.get(k, "")
    return c


def build_excel(d: Dossier) -> str:
    out = os.path.join(d.thuchien_dir, f"{d.code} - {d.sub_no} - TinhToanBSC.xlsx")
    os.makedirs(d.thuchien_dir, exist_ok=True)
    rd = d.results_data
    cab = rd["cabinet"]
    class1 = cab["cap_tu"] == "I"
    build_bsc_calculation_workbook(_calc_common(d), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Bảng tính"]
    vn = wb["TBKQ-VN"] if "TBKQ-VN" in wb.sheetnames else None
    en = wb["TBKQ - ENG"] if "TBKQ - ENG" in wb.sheetnames else None
    ws["B4"] = cab["cap_tu"]
    ws["A19"] = f"Mép trên cửa kính - khoảng mở {rd['inflow']['height_mm']} mm, rộng {rd['inflow']['width_mm']} mm"
    ws["A79"] = None   # ghi chú "(*) Không đánh giá do ồn nền..." không áp dụng
    if vn is not None:
        vn["T15"] = d.scope_vn + "."
        vn["F30"] = d.procedure_vn
    if en is not None:
        en["T13"] = d.scope_en + "."

    if class1:
        # 1. Inflow tủ cấp I: Kanomax đo trực tiếp (k = 1), mức EN 12469 cấp I 0,7 - 1,0 m/s
        ws["F29"] = "Mức quy định EN 12469 (tủ cấp I)"
        ws["F31"], ws["H31"] = T1_INFLOW_RANGE
        if vn is not None:
            vn["L68"] = "(0,70-1,00) m/s ⁽¹⁾"
        # 2. Downflow: tủ cấp I không có dòng khí cấp -> giữ khung, gạch chéo (AGENTS mục 8B)
        for m in list(ws.merged_cells.ranges):
            if any(r in str(m) for r in ("37", "38", "39", "40", "41")):
                ws.unmerge_cells(str(m))
        for r in range(37, 42):
            for c in range(3, 14):
                ws.cell(r, c).value = None
                ws.cell(r, c).border = Border()
        strike_box(ws, "C37:I41")
        for ref in ("D46", "F46", "H46", "J46", "D48", "D49", "H48", "H49", "J48"):
            ws[ref] = "-"
        ws["A34"] = "Ghi chú: Tủ cấp I không có dòng khí cấp qua lọc HEPA - không thực hiện chỉ tiêu này."
        # 5. HEPA: chỉ có lọc khí thải
        ws["D83"], ws["F83"], ws["H83"] = "-", "-", "-"
        # Đánh lại số mục: 1 inflow, 2 lux, 3 ồn, 4 HEPA, 5 hình thái dòng khí, 6 UVC
        ws["M50"], ws["M66"], ws["M80"], ws["M97"] = 2, 3, 4, 5
        ws["A112"] = "6. Đo cường độ UVC"
        # 6. Hình thái dòng khí tủ cấp I
        for row, af in zip((101, 106), T1_AF_ROWS):
            ws[f"A{row}"] = af["param_vn"]
            ws[f"D{row}"] = af["result_vn"]
            ws[f"H{row}"] = af["limit_vn"]
            ws[f"L{row}"] = "Đạt"
            for col in "ADH":
                ws[f"{col}{row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if vn is not None:
            for r in (71, 72, 73, 83, 84, 85):
                vn.row_dimensions[r].hidden = True
            vn["B93"] = T1_AF_ROWS[0]["param_vn"]
            vn["B101"] = T1_AF_ROWS[1]["param_vn"]
            vn["A109"] = 6
        if en is not None:
            for r in (70, 71, 72, 82, 83, 84):
                en.row_dimensions[r].hidden = True
            for row, af in zip((92, 102), T1_AF_ROWS):
                en[f"B{row}"] = af["param_en"]
                en[f"H{row}"] = af["result_en"]
                en[f"L{row}"] = af["limit_en"]
                en[f"P{row}"] = "Pass"
            en["A112"] = 6
    else:
        if vn is not None:
            vn["A109"] = 7
    wb.save(out)
    _recalc(out)
    print(f"[134804] Excel /{d.sub_no} -> {out}")
    return out


def _recalc(path: str):
    """Tính lại công thức bằng Excel COM (openpyxl lưu lại làm mất giá trị đã tính); không bám phiên Excel đang mở."""
    try:
        import gc
        import win32com.client
    except ImportError:
        print("[134804] Không có pywin32 - bỏ qua tính lại Excel")
        return
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        excel.CalculateFull()
        wb.Close(SaveChanges=True)
        wb = None
    finally:
        excel.Quit()
        del excel
        gc.collect()


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()

    jobs = [tu1_dossier(), tu2_dossier()]
    assign_cert_numbers(jobs)   # 2 báo cáo -> /1, /2
    summary, tems = [], []
    for d in jobs:
        d.ensure_output_directories()
        print(f"\n===== {d.cert_no_vn} =====")
        handler = BSCHandler(d)
        vn, en = handler.build_tbkq(is_en=False), handler.build_tbkq(is_en=True)
        tems.append(handler.build_tem_tu())
        xlsx = build_excel(d)
        for p in (vn, en):
            summary.append((os.path.basename(p), check_single_docx(p, code=d.code).get("status")))
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            ex = PDFExporter(d)
            ex.export_files([vn, en])
            ex.export_excel_files([xlsx])
    # BCKQ gộp 2 tủ (AGENTS mục 7)
    bckq = BCKQBuilder.build_merged(jobs)
    summary.insert(0, (os.path.basename(bckq), check_single_docx(bckq, code=CODE).get("status")))
    # Gom tem 2 tủ vào 1 file (AGENTS mục 8C), xóa tem lẻ vừa sinh
    sheet = build_tem_tu_sheet(tems, os.path.join(INPUT_DIR, "TBKQ", "134804 - tem tu.docx"))
    if sheet and os.path.isfile(sheet):
        for t in tems:
            if t and os.path.isfile(t) and os.path.abspath(t) != os.path.abspath(sheet):
                os.remove(t)
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(jobs[0]).export_files([p for p in (bckq, sheet) if p])
    print("\n===== QA =====")
    for name, st in summary:
        print(f"  {st}\t{name}")


if __name__ == "__main__":
    main()
