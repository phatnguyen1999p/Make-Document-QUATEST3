# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D136247 - Công ty Cổ phần Phân bón Bình Điền (06/11/2026).

1 chứng thư (mã trần): Tủ an toàn sinh học cấp 2 ESCO SC2-4E1 (Streamline), S/N 2016-105879, mã quản lý N/A
-> BSCHandler đầy đủ (không đo bụi). PYC mục 15: chứng thư **Tiếng Việt** -> chỉ dựng TBKQ VN (không bản EN).

Nguồn số liệu (chỉ đọc):
  - Input/136247/20261106164304322.pdf : BBLM (tr. 1-2) + phiếu ghi chép BM3030.01 + Ghi chép các thông số tủ (tr. 3-6)
  - Input/136247/20261106164318353.pdf : phiếu in nhiệt ATI 2i (lọc cấp 0,0042 % / 0,0038 %; lọc thải 7,4149 % / 0,0006 %)
  - Input/136247/20261105160050224.pdf : phiếu kiểm soát công việc, (tr. 2 hóa đơn YIWU SHIZHAO - không thuộc hồ sơ),
    PYC 26/10/2026, báo giá QONV26015972
  - Ảnh nhãn tủ (Photo/20261106_134852.jpg): ESCO SC2-4E1, Serial 2016-105879, sản xuất 2016;
    nhãn "Cabinet Airflow Setpoints" (20261106_134857.jpg): inflow 0,45 ± 0,025 m/s, k = 0,435, downflow 0,27 - 0,36 m/s
  - Giấy hiệu chuẩn (thư viện HieuChuan/): Kanomax MT01-1 KT3-02665AHL5 (Tốc độ gió), Testo 440 MT38-1
    DEDL26001061.02.2 (Độ rọi), Extech MT25 V11.CN5.1093.25 (UVC).
  - MST 0302975517 (tax_helper): tên quốc tế "BÌNH ĐIỀN FERTILIZER JOINT STOCK COMPANY"; địa chỉ ĐKKD
    "C12/21 đường Lê Khả Phiêu, Xã Tân Nhựt, TP Hồ Chí Minh" (khác địa chỉ PYC/BBLM — đang giữ theo PYC, chờ người dùng).
  -> Edoc: BBLM = 164304322 tr. 1–2; Ghi Chep = 164304322 tr. 3–6 + phiếu ATI; Khac = 160050224 trừ tr. 2 (hóa đơn lạ),
     bỏ trang trắng.

KHÔNG ĐẠT theo số liệu (09/11/2026, mức EN 12469:2000; không hiệu chỉnh): rò rỉ lọc HEPA khí thải
7,4149 % (> 0,01 %; Alarm Exceeded: Yes). Inflow 0,44 m/s (≥ 0,40), downflow 0,33 m/s (0,25 – 0,50), độ ồn 62,5 dBA -> Đạt.
(Theo mức NSX ESCO — MUC_NSX = True — inflow 0,442 trong 0,425 – 0,475 Đạt, downflow 0,33 trong 0,27 – 0,36 Đạt.)

Quyết định của người dùng:
  - 09/11/2026: "Tạm thời hồ sơ này bỏ qua hiệu chỉnh" — không hiệu chỉnh theo giấy hiệu chuẩn (cờ HIEU_CHINH = False):
    gió vào 0,442 m/s (TB 1,0175 × k 0,435), gió xuống 0,33 m/s, độ sáng 1116 lux, UVC 132 µW/cm²; phiếu 1 phần.
  - 09/11/2026: "Tạm đánh giá theo yêu cầu EN 12469" -> inflow ≥ 0,40 m/s, downflow 0,25 – 0,50 m/s; bỏ ghi chú ⁽²⁾
    NSX và "Tài liệu kỹ thuật của nhà sản xuất ESCO" ở Căn cứ (cờ MUC_NSX = False; đặt True để về mức NSX).
  - 09/11/2026: độ ồn khi vận hành = 62,5 dBA (thay 69,2 đọc trước đó) -> đạt (≤ 65 dBA).
  - 07/11/2026: dựng lại theo quy tắc mới — phiếu tính toán 2 phần; máy đo Kanomax MT01-1, Testo 440 MT38-1
    (MT38-1 không đạt kiểm định 17/04/2026, người dùng vẫn chọn). Số trên TBKQ/BCKQ = sau hiệu chỉnh.
  - 07/11/2026: tủ có chỉ tiêu không đạt vẫn cấp tem.
  - 06/11/2026:
  - (09/11/2026 tạm bỏ — MUC_NSX = False) Mức inflow / downflow theo mức công bố của NSX (nhãn ESCO "Cabinet Airflow Setpoints"):
    inflow 0,45 ± 0,025 = 0,425 - 0,475 m/s, downflow 0,27 - 0,36 m/s; ghi chú ⁽²⁾, Căn cứ thêm tài liệu NSX ESCO.
  - Địa điểm giám định ghi đủ theo BBLM: "Phòng Vi sinh - Công ty Cổ phần Phân bón Bình Điền - <địa chỉ>".
Suy luận của AI (cần người dùng xác nhận):
  - Cường độ sáng: hàng 2 ô 2 đọc 1324,3; hàng 3 ô 4 (viết đè) đọc 947,8.

Chạy:  python build_136247.py            (DOCX + XLSX, không PDF)
       python build_136247.py --pdf      (thêm xuất PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
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

CODE = "Q326D136247"
REG_DATE = "26/10/2026"
SAMPLE_DATE = "06/11/2026"
import datetime as _dt
CERT_DATE = _dt.date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày hiện tại, tự cập nhật mỗi lần dựng (người dùng chốt 06/11/2026, AGENTS mục 2)
START, END = "13:30", "15:00"
INPUT_DIR = os.path.join(WORKSPACE, "Input", "136247")
PHOTO_DIR = os.path.join(INPUT_DIR, "136247 Photo", "USE")

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần Phân bón Bình Điền",
    name_en="Binh Dien Fertilizer Joint Stock Company",   # tên quốc tế MST (bỏ dấu tiếng Việt cho bản EN)
    address_vn="Khu công nghiệp Long Định - Long Cang, xã Long Cang, tỉnh Tây Ninh",
    address_en="Long Dinh - Long Cang Industrial Park, Long Cang Commune, Tay Ninh Province",
)
LOCATION_VN = ("Phòng Vi sinh - Công ty Cổ phần Phân bón Bình Điền - Khu công nghiệp Long Định - Long Cang, "
               "xã Long Cang, tỉnh Tây Ninh")
LOCATION_EN = ("Microbiology Laboratory - Binh Dien Fertilizer Joint Stock Company - Long Dinh - Long Cang Industrial Park, "
               "Long Cang Commune, Tay Ninh Province")

# Mức đánh giá inflow / downflow: True = mức NSX (ESCO), False = mức EN 12469:2000 (≥ 0,40; 0,25 – 0,50 m/s)
# Người dùng 09/11/2026: "Tạm đánh giá theo yêu cầu EN 12469" -> False
MUC_NSX = False
# Mức công bố của nhà sản xuất — nhãn ESCO "Cabinet Airflow Setpoints" SC2-4E1 (ảnh 20261106_134857.jpg)
INFLOW_NSX = (0.425, 0.475)    # 0,45 ± 0,025 m/s (DIM và phương pháp thứ cấp)
DF_NSX = (0.27, 0.36)
NSX_BASIS_VN = "ESCO, nhãn thông số dòng khí (Cabinet Airflow Setpoints) trên tủ model SC2-4E1"
NSX_BASIS_EN = "ESCO, Cabinet Airflow Setpoints label on model SC2-4E1"

# Inflow phương pháp 2 (Kanomax, k theo nhãn ESCO); không dùng DIM
EQUIPMENTS = [eq("kanomax_6501"), eq("ati_6d_2i"), eq("smoke"),
              eq("testo_440"), eq("rion_nl53"), eq("extech_sdl470")]

# ---------------------------------------------------------------------------
# Số liệu phiếu ghi chép (06/11/2026)
# ---------------------------------------------------------------------------
INFLOW = [1.02, 0.96, 1.09, 1.06, 1.05, 0.95, 1.01, 1.00]   # Kanomax, mép trên cửa kính, 8 điểm; TB 1,02
INFLOW_K = 0.435                                           # nhãn ESCO "correction factor of 0.435"
INFLOW_W_MM = 1220
DOWNFLOW = [[0.32, 0.33, 0.33, 0.32, 0.32, 0.36, 0.37],
            [0.35, 0.34, 0.33, 0.33, 0.32, 0.33, 0.36],
            [0.34, 0.33, 0.32, 0.31, 0.32, 0.32, 0.34]]   # tổng hàng 2,35 / 2,36 / 2,28 khớp phiếu; 1220 x 570 mm
LUX = [[961.5, 1217.2, 1235.0, 1231.8],
       [1018.6, 1324.3, 1420.8, 1325.8],
       [798.6, 921.7, 988.7, 947.8]]
NOISE_AMB, NOISE_OP = 49.6, 62.5              # ồn nền file 0189; ồn vận hành 62,5 theo người dùng 09/11/2026 (trước 69,2)
UVC = [100, 160, 157, 111]

# Giấy hiệu chuẩn còn hiệu lực ngày đo (AGENTS mục 8B: phiếu tủ 2 phần)
_KANOMAX = CalTable.from_library("MT01-1", "Tốc độ gió", SAMPLE_DATE, device="Kanomax Climomaster 6501 (MT01-1)", unit="m/s")
# Hồ sơ này TẠM BỎ QUA HIỆU CHỈNH theo giấy hiệu chuẩn (người dùng 09/11/2026: "gió vào là 0,442 gió xuống là 0,33",
# "Tạm thời hồ sơ này bỏ qua hiệu chỉnh") -> phiếu tính toán 1 phần, số trên TBKQ/BCKQ = số đọc.
# Bật lại: AP_HIEU_CHINH = True trong Core/Tools/cleanroom_calc.py -> hiệu chỉnh lại (inflow 0,51, downflow 0,41, lux 1196, UVC 135).
from Core.Tools.cleanroom_calc import AP_HIEU_CHINH as HIEU_CHINH   # công tắc chung (09/11/2026: tạm tắt)
CALS = {
    "downflow": _KANOMAX,
    "inflow_kanomax": _KANOMAX,
    "lux": CalTable.from_library("MT38-1", "Độ rọi", SAMPLE_DATE, device="Testo 440 (MT38-1)", unit="lux"),
    "uvc": CalTable.from_library("MT25", "UVC", SAMPLE_DATE, device="Extech SDL470 (MT25)", unit="µW/cm²"),
} if HIEU_CHINH else {}


def _edoc():
    write_edoc(INPUT_DIR, "136247", {
        "BBLM": [("20261106164304322.pdf", [1, 2])],
        "Ghi Chep": [("20261106164304322.pdf", [3, 4, 5, 6]), ("20261106164318353.pdf", None)],
        "Khac": [("20261105160050224.pdf", [1] + list(range(3, 11)))],   # tr. 2 hóa đơn YIWU không thuộc hồ sơ
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
    # Số trên TBKQ/BCKQ = sau hiệu chỉnh, khớp phần b) của phiếu tính toán
    inf_c = corrected_values(INFLOW, CALS.get("inflow_kanomax"))
    v_in = sum(inf_c) / len(inf_c) * INFLOW_K
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
            "brand": "ESCO", "model": "SC2-4E1", "serial": "2016-105879", "device_id": "N/A",
            "seal_item_name": "“Tủ an toàn sinh học cấp 2”",
        },
        "inflow": {"method": 2, "method_name": "Kanomax Climomaster 6501", "readings": INFLOW,
                   "k_factor": INFLOW_K, "velocity": round(v_in, 2),
                   "width_mm": INFLOW_W_MM},
        "downflow": {"readings": DOWNFLOW, "avg": round(df_avg, 2), "tolerance_pct": 20,
                     "width_mm": 1220, "depth_mm": 570},
        "lux": {"readings": LUX, "avg": lux_avg, "avg_rounded": int(round(lux_avg))},
        "noise": {"ambient": NOISE_AMB, "operating": NOISE_OP, "corrected_dba": round(NOISE_OP - corr, 1),
                  "file_ambient": "0189"},
        "hepa": {"downflow_pct": 0.0042, "exhaust_pct": 7.4149, "aerosol_conc": 29,
                 "downflow_gasket_pct": 0.0042, "downflow_media_pct": 0.0038,
                 "exhaust_gasket_pct": 7.4149, "exhaust_media_pct": 0.0006},
        "airflow_pattern": {},
        "uvc": {"readings": UVC, "avg": round(uvc_avg, 1), "avg_rounded": int(round(uvc_avg))},
        **({"nsx": {"basis_vn": NSX_BASIS_VN, "basis_en": NSX_BASIS_EN}} if MUC_NSX else {}),
        "tasks_vn": ["Đo tốc độ dòng khí qua cửa làm việc;", "Đo tốc độ dòng khí cấp qua lọc HEPA;",
                     "Đo cường độ sáng bên trong tủ;", "Đo độ ồn của tủ khi vận hành;",
                     "Kiểm tra độ rò rỉ của các lọc HEPA;", "Kiểm tra hình thái dòng khí của tủ;",
                     "Đo cường độ tia UV dải C trong khoang làm việc."],
    }
    if MUC_NSX:
        rd["inflow"].update(limit_range=INFLOW_NSX, limit_source="nsx")
        rd["downflow"].update(limit_range=DF_NSX, limit_source="nsx")
    std_vn = "EN 12469:2000" + ("; Tài liệu kỹ thuật của nhà sản xuất ESCO" if MUC_NSX else "")
    std_en = "EN 12469:2000" + ("; Technical documents of the manufacturer ESCO" if MUC_NSX else "")
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False,
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        client=CLIENT, personnel=PersonnelInfo(executor="Phát"),
        input_dir=INPUT_DIR, tbkq_dir=os.path.join(INPUT_DIR, "TBKQ"), output_dir="", network_dir="",
        dossier_type="bsc",
        standards_vn=std_vn, standards_en=std_en, basis_vn=std_vn, basis_en=std_en,
        equipments=EQUIPMENTS,
        assessment={"basis_vn": "EN 12469:2000"},
        procedure_vn="QUATEST 3 3030:2025, N2.TC.BSC.02",
        procedure_en="QUATEST 3 3030:2025, N2.TC.BSC.02",
        scope_vn="Giám định tủ an toàn sinh học cấp 2",
        scope_en="Inspection of Biosafety cabinet class II",
        locations=[MonitoringLocation(name_vn="Tủ an toàn sinh học cấp 2 ESCO SC2-4E1",
                                      name_en="Biosafety cabinet class II ESCO SC2-4E1")],
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
    ws["A19"] = f"Mép trên cửa kính - rộng {INFLOW_W_MM} mm"
    ws["A79"] = None   # ghi chú "(*) Không đánh giá do ồn nền..." không áp dụng
    # Mức NSX (ESCO) cho inflow, downflow — công thức J31, J46 giữ nguyên, chỉ đổi ô mức
    if MUC_NSX:
        ws["F29"] = "Mức quy định NSX (ESCO)"
        ws["F31"], ws["H31"] = INFLOW_NSX
        ws["F44"] = "Mức quy định NSX (ESCO)"
        ws["F46"], ws["H46"] = DF_NSX
    vn = wb["TBKQ-VN"] if "TBKQ-VN" in wb.sheetnames else None
    en = wb["TBKQ - ENG"] if "TBKQ - ENG" in wb.sheetnames else None
    if vn is not None:
        vn["T15"] = d.scope_vn + "."
        vn["F30"] = d.procedure_vn
        vn["A109"] = 7
        if MUC_NSX:
            vn["L68"] = "(0,425-0,475) m/s ⁽²⁾"
            vn["L71"] = "(0,27-0,36) m/s ⁽²⁾"
        for row in (vn.iter_rows() if MUC_NSX else ()):
            for c in row:
                if isinstance(c.value, str) and "Mức quy định tại EN 12469:2000" in c.value and "⁽²⁾" not in c.value:
                    c.value = c.value.rstrip() + f"\n⁽²⁾ Mức công bố của nhà sản xuất: {NSX_BASIS_VN}."
    if en is not None:
        en["T13"] = d.scope_en + "."
    wb.save(out)
    finalize_bsc_workbook(out, CALS)   # phần b) sau hiệu chỉnh + công thức ồn nền, tính lại bằng Excel COM
    print(f"[136247] Excel -> {out}")
    return out


def _compact_cover(path: str):
    """Giữ bìa TBKQ 1 trang: thu khoảng cách đoạn trong Table 0 3pt -> 1,5pt (AGENTS mục 14), giữ khoảng ký 3 cm."""
    import docx
    from Core.Builders.tag_context import shrink_cover_break
    doc = docx.Document(path)
    shrink_cover_break(doc, compact=True)
    # Giá trị Không đạt -> in đậm (AGENTS mục 3; Core chưa làm - E41): ô Kết quả đo của hàng có Đánh giá "Không đạt"
    for t in doc.tables:
        for r in t.rows:
            cells = r.cells
            if len(cells) >= 5 and cells[-1].text.strip() in ("Không đạt", "Fail"):
                for p in cells[2].paragraphs:
                    for run in p.runs:
                        run.bold = True
    _shrink_photos(doc, 4.2)
    doc.save(path)


def _shrink_photos(doc, width_cm: float):
    """7 ảnh (3 hàng) + dòng ghi chú NSX làm hàng ảnh cuối rớt sang trang 3 trống: thu ảnh grid 3 cột
    4,8 -> width_cm (vẫn trong 4 - 5 cm, AGENTS mục 3), giữ tỉ lệ, không crop."""
    from docx.oxml.ns import qn
    new_cx = int(width_cm * 360000)
    for t in doc.tables:
        inl = t._tbl.findall(".//" + qn("wp:inline"))
        if len(inl) < 5:
            continue
        for el in inl:
            ext = el.find(qn("wp:extent"))
            cx, cy = int(ext.get("cx")), int(ext.get("cy"))
            if cx <= new_cx:
                continue
            ncy = int(cy * new_cx / cx)
            ext.set("cx", str(new_cx)); ext.set("cy", str(ncy))
            for a_ext in el.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}ext"):
                if a_ext.get("cx"):
                    a_ext.set("cx", str(new_cx)); a_ext.set("cy", str(ncy))


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
    vn = handler.build_tbkq(is_en=False)   # PYC: chứng thư Tiếng Việt, không dựng bản EN
    _compact_cover(vn)
    tem = handler.build_tem_tu()
    xlsx = build_excel(d)
    bckq = BCKQBuilder(d).build()
    summary = [(os.path.basename(p), check_single_docx(p, code=d.code).get("status")) for p in (bckq, vn)]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(d)
        ex.export_files([p for p in (vn, bckq, tem) if p])
        ex.export_excel_files([xlsx])
    print("\n===== QA =====")
    for name, st in summary:
        print(f"  {st}\t{name}")
    print(f"  tem: {tem}")


if __name__ == "__main__":
    main()
