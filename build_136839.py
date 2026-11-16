# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D136839 - Chi nhánh Công ty TNHH Nước giải khát Swire Coca-Cola Việt Nam tại tỉnh Tây Ninh (06/11/2026).

1 chứng thư (mã trần): Tủ an toàn sinh học cấp 2 BioAir S.p.A., model SafeMate ECO+ 1.5, S/N Y1819,
mã quản lý S_QA_MLA_TAG_030 -> BSCHandler đầy đủ (không đo bụi). Chứng thư Việt - Anh (PYC mục 15).

Nguồn số liệu (chỉ đọc):
  - Input/136839/20261106163042996.pdf  : BBLM (trang 1-2) + phiếu ghi chép BM3030.01 + Ghi chép các thông số tủ (trang 3-6)
  - Input/136839/20261106163058822.pdf  : phiếu in nhiệt ATI 2i (4 lần quét, upstream 22 µg/L)
  - Input/136839/20261105160102681.pdf  : phiếu kiểm soát công việc, PYC 27/10/2026, báo giá QONV26015971, PO
    (tr. 1 phiếu kiểm soát công việc, 2 trắng, 3–6 PYC + điều kiện chung (2 bản), 7–10 báo giá, 11–16 PO + trang trắng)
  -> Edoc: BBLM = 163042996 tr. 1–2; Ghi Chep = 163042996 tr. 3–6 + phiếu ATI; Khac = 160102681 (bỏ trang trắng)
  - Ảnh nhãn tủ (Photo/USE/20261106_094331.jpg): BioAir S.p.A., SafeMate ECO+ 1.5, s.n. Y1819, năm 2024
  - Giấy hiệu chuẩn (thư viện HieuChuan/): Kanomax MT01-1 KT3-02665AHL5 (Tốc độ gió), Testo 440 MT38-1
    DEDL26001061.02.2 (Độ rọi), Extech MT25 V11.CN5.1093.25 (UVC). TSI 8375 (MT31) chỉ có bảng chênh áp -> inflow DIM 1 phần.
  - MST 0300792451-009 (tax_helper): tên quốc tế "BRANCH OF SWIRE COCA-COLA BEVERAGES VIET NAM LIMITED LIABILITY
    COMPANY IN TAY NINH PROVINCE"; tên VN đăng ký "CHI NHÁNH CÔNG TY TNHH NƯỚC GIẢI KHÁT COCA-COLA VIỆT NAM TẠI TỈNH
    TÂY NINH" (không có "Swire" — lệch PYC/BBLM; người dùng chốt 07/11/2026 giữ tên theo PYC).

Quyết định của người dùng:
  - 09/11/2026: "bỏ hiệu chuẩn cả trong 136839" — tạm không hiệu chỉnh theo giấy hiệu chuẩn (cờ HIEU_CHINH = False):
    phiếu tính toán 1 phần, TBKQ/BCKQ dùng số đọc (downflow, độ sáng, UVC). Hiệu chỉnh ồn nền giữ nguyên.
  - 06/11/2026: Ồn nền = 59,8 dBA (> 55) -> ΔN = 63,7 - 59,8 = 3,9 -> 4 -> trừ 2 -> độ ồn vận hành hiệu chỉnh 61,7 dBA.
  - 06/11/2026: Địa điểm ghi đủ theo BBLM (không ghi gọn).
  - 07/11/2026: dựng lại theo quy tắc mới — phiếu tính toán 2 phần (đọc / sau hiệu chỉnh); máy đo: Kanomax MT01-1,
    Testo 440 MT38-1 (người dùng chọn, dù MT38-1 không đạt kiểm định 17/04/2026). Số trên TBKQ/BCKQ = sau hiệu chỉnh.
  - 07/11/2026: tên VN giữ theo PYC ("Chi nhánh Công ty TNHH Nước giải khát Swire Coca-Cola Việt Nam tại tỉnh
    Tây Ninh"), không đổi theo tên đăng ký MST (không có "Swire").
Suy luận của AI (cần người dùng xác nhận):
  - HEPA: 2 lần quét 10:19 là lọc cấp (max 0,0010 % / 0,0006 %), 2 lần 10:20 là lọc thải (0,0015 % / 0,0006 %);
    phiếu ghi chép ghi lọc cấp 0,0010 %, lọc thải 0,0015 % (khớp max).
  - Tên EN khách hàng theo tên quốc tế MST (AGENTS mục 15).

Chạy:  python build_136839.py            (DOCX + XLSX, không PDF)
       python build_136839.py --pdf      (thêm xuất PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: (network_dir để trống; chép lên Z bằng robocopy theo AGENTS mục 0).
"""

import os
import sys
import argparse

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_co_Z_"))

import openpyxl

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.bsc import BSCHandler
from Core.Tools.bsc_calc import build_bsc_calculation_workbook
from Core.Tools.bsc_cal_blocks import finalize_bsc_workbook, corrected_values
from Core.Tools.cleanroom_calc import CalTable
from Core.Tools.scan_splitter import write_edoc
from Core.Tools.noise_corr import noise_correction_value
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D136839"
REG_DATE = "27/10/2026"
SAMPLE_DATE = "06/11/2026"
import datetime as _dt
CERT_DATE = _dt.date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày hiện tại, tự cập nhật mỗi lần dựng (người dùng chốt 06/11/2026, AGENTS mục 2)
START, END = "09:00", "11:00"
INPUT_DIR = os.path.join(WORKSPACE, "Input", "136839")
PHOTO_DIR = os.path.join(INPUT_DIR, "136839 0- Photo", "USE")

CLIENT = ClientInfo(
    name_vn="Chi nhánh Công ty TNHH Nước giải khát Swire Coca-Cola Việt Nam tại tỉnh Tây Ninh",
    name_en="Branch of Swire Coca-Cola Beverages Viet Nam Limited Liability Company in Tay Ninh Province",  # tên quốc tế MST
    address_vn="Lô C12, đường Dọc 2, Khu công nghiệp Phú An Thạnh, xã Bến Lức, tỉnh Tây Ninh",
    address_en="Lot C12, Doc 2 Street, Phu An Thanh Industrial Park, Ben Luc Commune, Tay Ninh Province",
)
# Địa điểm ghi đủ theo BBLM (người dùng nhắc 06/11/2026 ở hồ sơ 136247: không rút gọn tên công ty)
LOCATION_VN = ("Phòng thí nghiệm tại Chi nhánh Công ty TNHH Nước giải khát Swire Coca-Cola Việt Nam tại tỉnh Tây Ninh"
               " - Lô C12, đường Dọc 2, Khu công nghiệp Phú An Thạnh, xã Bến Lức, tỉnh Tây Ninh")
LOCATION_EN = ("Laboratory at Branch of Swire Coca-Cola Beverages Viet Nam Limited Liability Company in Tay Ninh Province"
               " - Lot C12, Doc 2 Street, Phu An Thanh Industrial Park, Ben Luc Commune, Tay Ninh Province")

# Inflow phương pháp 1 (DIM TSI 8375); downflow Kanomax
EQUIPMENTS = [eq("tsi_8375"), eq("kanomax_6501"), eq("ati_6d_2i"), eq("smoke"),
              eq("testo_440"), eq("rion_nl53"), eq("extech_sdl470")]

# ---------------------------------------------------------------------------
# Số liệu phiếu ghi chép (06/11/2026)
# ---------------------------------------------------------------------------
INFLOW_Q = [433, 434, 433, 437, 436]          # m³/h, DIM TSI 8375; TB 434,6
INFLOW_H_MM, INFLOW_W_MM = 200, 1500          # độ mở 0,2 m, rộng 1,5 m -> 0,300 m²
DOWNFLOW = [[0.32, 0.33, 0.32, 0.32, 0.31, 0.32, 0.33, 0.33, 0.35],
            [0.31, 0.31, 0.31, 0.32, 0.31, 0.31, 0.33, 0.33, 0.35],
            [0.36, 0.34, 0.34, 0.34, 0.34, 0.33, 0.33, 0.37, 0.37]]   # tổng hàng 2,93 / 2,88 / 3,12 khớp phiếu
LUX = [[1157.6, 1378.6, 1490.3, 1384.0, 1075.8],
       [1109.5, 1380.5, 1476.2, 1364.2, 1079.7],
       [958.8, 1223.0, 1241.7, 1138.0, 917.8]]
NOISE_AMB, NOISE_OP = 59.8, 63.7              # file 0187 (nền, người dùng xác nhận 59,8 — 06/11/2026), 0186 (vận hành)
UVC = [211, 311, 380, 307, 212]

# Giấy hiệu chuẩn còn hiệu lực ngày đo (AGENTS mục 8B: phiếu tủ 2 phần). TSI 8375 chưa có bảng lưu lượng -> None.
# Hồ sơ này TẠM BỎ QUA HIỆU CHỈNH theo giấy hiệu chuẩn (người dùng 09/11/2026: "bỏ hiệu chuẩn cả trong 136839")
# -> phiếu tính toán 1 phần, số trên TBKQ/BCKQ = số đọc. Hiệu chỉnh ồn nền (quy trình đo) vẫn giữ.
# Bật lại: AP_HIEU_CHINH = True trong Core/Tools/cleanroom_calc.py -> hiệu chỉnh lại theo giấy.
from Core.Tools.cleanroom_calc import AP_HIEU_CHINH as HIEU_CHINH   # công tắc chung (09/11/2026: tạm tắt)
CALS = {
    "downflow": CalTable.from_library("MT01-1", "Tốc độ gió", SAMPLE_DATE, device="Kanomax Climomaster 6501 (MT01-1)", unit="m/s"),
    "lux": CalTable.from_library("MT38-1", "Độ rọi", SAMPLE_DATE, device="Testo 440 (MT38-1)", unit="lux"),
    "uvc": CalTable.from_library("MT25", "UVC", SAMPLE_DATE, device="Extech SDL470 (MT25)", unit="µW/cm²"),
    "inflow_dim": None,
} if HIEU_CHINH else {}


def _edoc():
    write_edoc(INPUT_DIR, "136839", {
        "BBLM": [("20261106163042996.pdf", [1, 2])],
        "Ghi Chep": [("20261106163042996.pdf", [3, 4, 5, 6]), ("20261106163058822.pdf", None)],
        "Khac": [("20261105160102681.pdf", None)],
    })


def _photos():
    if not os.path.isdir(PHOTO_DIR):
        return []
    return [{"path": os.path.join(PHOTO_DIR, f), "caption_vn": "", "caption_en": ""}
            for f in sorted(os.listdir(PHOTO_DIR)) if f.lower().endswith((".jpg", ".jpeg", ".png"))]


def _avg(grid):
    vals = [v for row in grid for v in row]
    return sum(vals) / len(vals)


def dossier() -> Dossier:
    q_avg = sum(INFLOW_Q) / len(INFLOW_Q)
    v_in = q_avg / (INFLOW_H_MM * INFLOW_W_MM / 1e6) / 3600
    # Số trên TBKQ/BCKQ = sau hiệu chỉnh, khớp phần b) của phiếu tính toán
    df_avg = _avg(corrected_values(DOWNFLOW, CALS.get("downflow")))
    lux_avg = _avg(corrected_values(LUX, CALS.get("lux")))
    corr = noise_correction_value(NOISE_AMB, NOISE_OP)
    uvc_c = corrected_values(UVC, CALS.get("uvc"))
    uvc_avg = sum(uvc_c) / len(uvc_c)
    rd = {
        "start_time": START, "end_time": END,
        "file_label": "BSC", "bckq_label": "bsc",
        "cabinet": {
            "cabinet_name": "Tủ an toàn sinh học cấp 2",
            "cabinet_name_en": "Biosafety cabinet class II",
            "cap_tu": "II",
            "brand": "BioAir S.p.A.", "model": "SafeMate ECO+ 1.5", "serial": "Y1819",
            "device_id": "S_QA_MLA_TAG_030",
            "seal_item_name": "“Tủ an toàn sinh học cấp 2”",
        },
        "inflow": {"method": 1, "method_name": "DIM TSI 8375", "q_readings": INFLOW_Q,
                   "width_mm": INFLOW_W_MM, "height_mm": INFLOW_H_MM,
                   "velocity": round(v_in, 2), "scan_velocity": 0.402},
        "downflow": {"readings": DOWNFLOW, "avg": round(df_avg, 2), "tolerance_pct": 20,
                     "width_mm": 1530},
        "lux": {"readings": LUX, "avg": lux_avg, "avg_rounded": int(round(lux_avg))},
        "noise": {"ambient": NOISE_AMB, "operating": NOISE_OP, "corrected_dba": round(NOISE_OP - corr, 1),
                  "file_ambient": "0187", "file_operating": "0186"},
        "hepa": {"downflow_pct": 0.0010, "exhaust_pct": 0.0015, "aerosol_conc": 22,
                 "downflow_gasket_pct": 0.0010, "downflow_media_pct": 0.0006,
                 "exhaust_gasket_pct": 0.0015, "exhaust_media_pct": 0.0006},
        "airflow_pattern": {},
        "uvc": {"readings": UVC, "avg": round(uvc_avg, 1), "avg_rounded": int(round(uvc_avg))},
        "tasks_vn": ["Đo tốc độ dòng khí qua cửa làm việc;", "Đo tốc độ dòng khí cấp qua lọc HEPA;",
                     "Đo cường độ sáng bên trong tủ;", "Đo độ ồn của tủ khi vận hành;",
                     "Kiểm tra độ rò rỉ của các lọc HEPA;", "Kiểm tra hình thái dòng khí của tủ;",
                     "Đo cường độ tia UV dải C trong khoang làm việc."],
    }
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False,
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        client=CLIENT, personnel=PersonnelInfo(executor="Phát"),
        input_dir=INPUT_DIR, tbkq_dir=os.path.join(INPUT_DIR, "TBKQ"), output_dir="", network_dir="",
        dossier_type="bsc",
        standards_vn="EN 12469:2000", standards_en="EN 12469:2000",
        basis_vn="EN 12469:2000", basis_en="EN 12469:2000",
        equipments=EQUIPMENTS,
        assessment={"basis_vn": "EN 12469:2000"},
        procedure_vn="QUATEST 3 3030:2025, N2.TC.BSC.02",
        procedure_en="QUATEST 3 3030:2025, N2.TC.BSC.02",
        scope_vn="Giám định tủ an toàn sinh học cấp 2",
        scope_en="Inspection of Biosafety cabinet class II",
        locations=[MonitoringLocation(name_vn="Tủ an toàn sinh học cấp 2 BioAir SafeMate ECO+ 1.5",
                                      name_en="Biosafety cabinet class II BioAir SafeMate ECO+ 1.5")],
        photos=_photos(),
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
    out = os.path.join(d.thuchien_dir, f"{d.code} - TinhToanBSC.xlsx")
    os.makedirs(d.thuchien_dir, exist_ok=True)
    build_bsc_calculation_workbook(_calc_common(d), out, cals=CALS, finalize=False)
    wb = openpyxl.load_workbook(out)
    ws = wb["Bảng tính"]
    ws["B4"] = "II"
    ws["A10"] = "Phương pháp 1:  DIM - TSI 8375"      # theo phiếu ghi chép (bộ dựng ghi cứng 8380)
    ws["A79"] = None   # ghi chú "(*) Không đánh giá do ồn nền..." không áp dụng
    vn = wb["TBKQ-VN"] if "TBKQ-VN" in wb.sheetnames else None
    en = wb["TBKQ - ENG"] if "TBKQ - ENG" in wb.sheetnames else None
    if vn is not None:
        vn["T15"] = d.scope_vn + "."
        vn["F30"] = d.procedure_vn
        vn["A109"] = 7
    if en is not None:
        en["T13"] = d.scope_en + "."
    wb.save(out)
    finalize_bsc_workbook(out, CALS)   # phần b) sau hiệu chỉnh + công thức ồn nền, tính lại bằng Excel COM
    print(f"[136839] Excel -> {out}")
    return out


def _compact_cover(path: str):
    """Tên KH + địa điểm dài (2 - 3 dòng) làm bìa tràn: thu khoảng cách đoạn trong Table 0 3pt -> 1,5pt
    (như khi Căn cứ từ 2 tiêu chuẩn, AGENTS mục 14), giữ khoảng ký 3 cm."""
    import docx
    from Core.Builders.tag_context import shrink_cover_break
    doc = docx.Document(path)
    # Địa điểm ghi đủ theo BBLM (3 - 4 dòng) -> bìa EN tràn: khoảng cách đoạn trong ô Table 0 về 0pt (như Căn cứ >= 4 mục)
    shrink_cover_break(doc, compact=2)
    doc.save(path)


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()

    d = dossier()
    assign_cert_numbers([d])   # 1 báo cáo -> mã trần
    d.ensure_output_directories()
    _edoc()
    print(f"\n===== {d.cert_no_vn} =====")
    handler = BSCHandler(d)
    vn, en = handler.build_tbkq(is_en=False), handler.build_tbkq(is_en=True)
    for p in (vn, en):
        _compact_cover(p)
    tem = handler.build_tem_tu()
    xlsx = build_excel(d)
    bckq = BCKQBuilder(d).build()
    summary = [(os.path.basename(p), check_single_docx(p, code=d.code).get("status")) for p in (bckq, vn, en)]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(d)
        ex.export_files([p for p in (vn, en, bckq, tem) if p])
        ex.export_excel_files([xlsx])
    print("\n===== QA =====")
    for name, st in summary:
        print(f"  {st}\t{name}")
    print(f"  tem: {tem}")


if __name__ == "__main__":
    main()
