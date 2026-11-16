# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D132488 - Công ty Cổ phần Thành Thành Công - Biên Hòa (27/10/2026).

4 chứng thư:
  /1  Tủ cấy ESCO AVC-4D8 (TBKN46-01)  -> BSCHandler (bỏ inflow, lọc khí thải; downflow & lux theo mức NSX ESCO)
  /2  Tủ hút HUYHOANG (TBKN29-2)        -> FumeHoodHandler, form KHÔNG đánh giá (không đo cường độ sáng)
  /3  Tủ hút PROLAB (TBKN29-4)          -> FumeHoodHandler, form KHÔNG đánh giá
  /4  Tủ hút PROLAB (TBKN29-5)          -> FumeHoodHandler, form KHÔNG đánh giá

Quyết định của người dùng (27/10/2026):
  - Tủ cấy: hình thái dòng khí có làm, Đạt.
    Người dùng cập nhật 02/11/2026 (theo mẫu TBKQ tủ cấy Q326D045837/2): 2 vị trí (bên trong không gian tủ;
    chu vi ngoài vùng mở cửa tủ), mức tham chiếu CAG-003:2022 (ghi chú ⁽⁴⁾), thay câu cũ 27/10
    "dòng khói đi xuống và thoát ra ngoài, không quẩn trong tủ" / mức EN 12469.
  - Tủ hút: form không đánh giá, mức khuyến nghị SEFA 1-2010 (inflow 0,35 - 0,51 m/s; lux ≥ 861; ồn ≤ 85 dBA).
    Không làm hình thái dòng khí.
  - Tủ 4: khoảng mở cửa 450 mm (phiếu bỏ trống).
Người dùng chốt 30/10/2026: tủ cấy đánh giá downflow / cường độ sáng theo catalogue ESCO Airstream AVC-4D_
  (0,45 m/s tại điểm cài đặt ban đầu -> chấp nhận ± 20 % = 0,36 - 0,54 m/s; đèn 1 110 lux -> ≥ 1 110 lux).
  Độ ồn giữ ≤ 65 dBA (EN 12469 / LAF.01; catalogue 56,5 dBA đo trong buồng tiêu âm, không dùng làm mức).

Số liệu đọc từ phiếu ghi chép hiện trường (Input/132488/<1..4>/132488 - <n>.pdf) và phiếu in nhiệt
(Input/132488/1/20261027174658684.pdf: HEPA ATI 2i, AeroTrak 9350 sample 469/470).

Chạy:  python build_132488.py            (DOCX + XLSX, không PDF)
       python build_132488.py --pdf      (thêm xuất PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: (network_dir để trống).
"""

import os
import sys
import argparse

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)

import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.bsc import BSCHandler
from Core.Builders.handlers.fumehood import FumeHoodHandler
from Core.Tools.bsc_calc import build_bsc_calculation_workbook
from Core.Tools.fumehood_calc import build_fumehood_calculation_workbook
from Core.Tools.noise_corr import noise_correction_value
from Core.Tools.qa_check import check_single_docx
from Core.Builders.tag_context import build_tem_tu_sheet
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D132488"
REG_DATE = "19/10/2026"      # BBLM ghi "19/10/2025" (nghi viết nhầm năm); phiếu ghi chép ghi 19/10/2026
SAMPLE_DATE = "27/10/2026"
import datetime as _dt
CERT_DATE = _dt.date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày hiện tại, tự cập nhật mỗi lần dựng (người dùng chốt 06/11/2026, AGENTS mục 2)
START, END = "10:00", "13:30"
INPUT_DIR = os.path.join(WORKSPACE, "Input", "132488")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photos")

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần Thành Thành Công - Biên Hòa",
    name_en="Thanh Thanh Cong - Bien Hoa Joint Stock Company",
    address_vn="Ấp Tân Lợi, xã Tân Phú, tỉnh Tây Ninh",
    address_en="Tan Loi Hamlet, Tan Phu Commune, Tay Ninh Province",
)
LOCATION_VN = "Công ty Cổ phần Thành Thành Công - Biên Hòa - Ấp Tân Lợi, xã Tân Phú, tỉnh Tây Ninh"
LOCATION_EN = "Thanh Thanh Cong - Bien Hoa Joint Stock Company - Tan Loi Hamlet, Tan Phu Commune, Tay Ninh Province"

EQ = {
    "kanomax": eq("kanomax_6501"),
    "testo": eq("testo_440"),
    "rion": eq("rion_nl53"),
    "ati": eq("ati_6d_2i"),
    "aerotrak": eq("aerotrak_9350"),
    "smoke": eq("smoke"),
}


def _photos(sub: str, folder: str):
    d = os.path.join(PHOTO_DIR, sub, folder)
    if not os.path.isdir(d):
        return []
    return [{"path": os.path.join(d, f), "caption_vn": "", "caption_en": ""}
            for f in sorted(os.listdir(d)) if f.lower().endswith((".jpg", ".jpeg", ".png"))]


def _personnel():
    return PersonnelInfo(executor="Phát")


def _base_kwargs():
    # Số phụ /1, /2, /3, /4 do assign_cert_numbers() tự đánh theo thứ tự báo cáo trong main()
    return dict(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=True,
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        client=CLIENT, personnel=_personnel(),
        input_dir=INPUT_DIR, tbkq_dir=os.path.join(INPUT_DIR, "TBKQ"), output_dir="", network_dir="",
    )


def _noise_corr(ambient: float, operating: float) -> float:
    """Hiệu chỉnh độ ồn theo bảng của quy trình (chung: Core/Tools/noise_corr.py, chốt 30/10/2026).
    ΔN làm tròn số nguyên; ΔN ≤ 2 → không hiệu chỉnh và in cảnh báo giảm ồn nền."""
    return noise_correction_value(ambient, operating)


def _hood_corr(h: dict) -> float:
    """Hiệu chỉnh ồn của một tủ hút; `noise_corr_override` (nếu có) giữ số đã chốt trên phiếu."""
    ov = h.get("noise_corr_override")
    return float(ov) if ov is not None else _noise_corr(h["amb"], h["op"])


def _avg(grid):
    vals = [v for row in grid for v in row]
    return sum(vals) / len(vals)


# ---------------------------------------------------------------------------
# /1 - Tủ cấy ESCO AVC-4D8
# ---------------------------------------------------------------------------
TU_CAY_DOWNFLOW = [
    [0.61, 0.60, 0.64, 0.56, 0.53, 0.79, 0.83],
    [0.87, 0.74, 0.80, 0.76, 0.68, 0.69, 0.68],
    [0.80, 0.73, 0.71, 0.62, 0.69, 0.67, 0.59],
]
TU_CAY_LUX = [
    [633.0, 700.7, 668.3, 731.7],
    [704.0, 781.8, 727.0, 766.2],
    [603.5, 676.2, 742.6, 680.0],
]
TU_CAY_PARTICLE = [1, 0, 0, 0]   # hạt/m³ tích lũy d ≥ 0,3 / 0,5 / 1,0 / 5,0 µm (sample 470, 1000 L)

# Mức công bố của nhà sản xuất ESCO, catalogue Airstream AVC-4D_ (1,2 m) — người dùng chốt 30/10/2026
DF_NSX = (0.36, 0.54)   # 0,45 m/s ± 20 %
LUX_NSX = 1110          # cường độ đèn khi ánh sáng nền bằng 0
# Ghi chú ⁽³⁾ rút gọn, bỏ phần trong ngoặc (0,45 m/s ± 20 %; 1 110 lux) — người dùng sửa trên Z và chốt 01/11/2026.
# Downflow / độ sáng tủ cấy vẫn đánh giá theo mức NSX (Không đạt) ở cả TBKQ và BCKQ — người dùng chốt 01/11/2026.
NSX_BASIS_VN = "ESCO, catalogue Airstream model AVC-4D_"
NSX_BASIS_EN = "ESCO, Airstream model AVC-4D_ catalogue"

# Hình thái dòng khí tủ cấy — người dùng cập nhật 02/11/2026 theo mẫu TBKQ tủ cấy Q326D045837/2.
# Mức tham chiếu CAG-003:2022; "{ref}" = ký hiệu ghi chú do BSCHandler.airflow_mark() đánh số (⁽⁴⁾).
AF_BASIS = "CAG-003:2022"
AF_ROWS = [
    {"param_vn": "Bên trong không gian tủ\n- Dọc đường trung tâm, phía trên khoảng mở làm việc"
                 "\n- Phía sau cửa mở làm việc\n- Dọc theo mép cửa mở làm việc",
     "param_en": "Inside cabinet workspace\n- Along center line, above work opening"
                 "\n- Behind work sash\n- Along edge of work sash",
     "result_vn": "- Dòng khói hướng xuống, không có điểm tù;\n- Dòng khói thoát ra khỏi mặt thao tác.",
     "result_en": "- Smooth downward flow, no dead spots;\n- Smoke exited off the work surface.",
     "limit_vn": "- Dòng khói hướng xuống, không có điểm tù;\n- Dòng khói thoát ra khỏi mặt thao tác. {ref}",
     "limit_en": "- Smooth downward flow, no dead spots;\n- Smoke exits off the work surface. {ref}",
     "height": 1750},   # VN 6 dòng chữ (mục 9: chiều cao theo bản VN, áp cả EN)
    {"param_vn": "- Khu vực dọc theo chu vi phía bên ngoài của vùng mở cửa tủ",
     "param_en": "- Area along the outer perimeter of the work opening",
     "result_vn": "- Dòng khói hướng ra phía trước tủ.",
     "result_en": "- Smoke flowed outward to the front of the cabinet.",
     "limit_vn": "- Dòng khói hướng ra phía trước tủ. {ref}",
     "limit_en": "- Smoke flows outward to the front of the cabinet. {ref}",
     "height": 820},    # VN 3 dòng
]


def tu_cay_dossier() -> Dossier:
    amb, op = 52.1, 62.2
    corr = _noise_corr(amb, op)
    df_avg = _avg(TU_CAY_DOWNFLOW)
    lux_avg = _avg(TU_CAY_LUX)
    pass_vn, pass_en = "Đạt", "Pass"

    def size_row(lbl_vn, lbl_en, v, lim_vn, lim_en, lim_num):
        ok = v <= lim_num
        return (lbl_vn, lbl_en, f"{v} hạt/m³", f"{v} particles/m³", lim_vn, lim_en,
                pass_vn if ok else "Không đạt", pass_en if ok else "Fail")

    rd = {
        "start_time": START, "end_time": END,
        "file_label": "Tu cay",
        "bckq_label": "tu cay",
        "cabinet": {
            "cabinet_name": "Tủ cấy",
            "cabinet_name_en": "Laminar flow cabinet",
            "brand": "ESCO", "model": "AVC-4D8", "serial": "2014-95852", "device_id": "TBKN46-01",
            "seal_item_name": "“Tủ cấy”",
        },
        "nsx": {"basis_vn": NSX_BASIS_VN, "basis_en": NSX_BASIS_EN},
        "inflow": {"skip": True},
        "downflow": {"readings": TU_CAY_DOWNFLOW, "avg": round(df_avg, 2), "tolerance_pct": 20,
                     "limit_range": DF_NSX, "limit_source": "nsx",
                     "width_mm": 1200, "depth_mm": 600},
        "lux": {"readings": TU_CAY_LUX, "avg": lux_avg, "avg_rounded": int(round(lux_avg)),
                "limit_min": LUX_NSX, "limit_source": "nsx"},
        "noise": {"ambient": amb, "operating": op, "corrected_dba": round(op - corr, 1), "file_no": ""},
        "hepa": {"downflow_pct": 0.0024, "exhaust_pct": None, "skip_exhaust": True, "aerosol_conc": 15,
                 "downflow_gasket_pct": 0.0024, "downflow_media_pct": 0.0012},
        "airflow_pattern": {"basis_vn": AF_BASIS, "basis_en": AF_BASIS,
                            "rows": [dict(r, assess_vn=pass_vn, assess_en=pass_en) for r in AF_ROWS]},
        "particle": {
            "clean_class": "ISO Class 5", "sample_volume_l": 1000.0, "sample_time_min": 20,
            "readings": TU_CAY_PARTICLE, "file_zero": "469", "file_ghi": "470",
            "sizes": [
                size_row("Kích thước hạt d ≥ 0,3 µm", "Particle size d ≥ 0.3 µm", TU_CAY_PARTICLE[0], "≤ 10 200\nhạt/m³ ⁽²⁾", "≤ 10,200\nparticles/m³ ⁽²⁾", 10200),
                size_row("Kích thước hạt d ≥ 0,5 µm", "Particle size d ≥ 0.5 µm", TU_CAY_PARTICLE[1], "≤ 3 520\nhạt/m³ ⁽²⁾", "≤ 3,520\nparticles/m³ ⁽²⁾", 3520),
                size_row("Kích thước hạt d ≥ 1,0 µm", "Particle size d ≥ 1.0 µm", TU_CAY_PARTICLE[2], "≤ 832\nhạt/m³ ⁽²⁾", "≤ 832\nparticles/m³ ⁽²⁾", 832),
                size_row("Kích thước hạt d ≥ 5,0 µm", "Particle size d ≥ 5.0 µm", TU_CAY_PARTICLE[3], "≤ 29\nhạt/m³ ⁽²⁾", "≤ 29\nparticles/m³ ⁽²⁾", 29),
            ],
        },
        "tasks_vn": ["Đo tốc độ dòng khí cấp qua lọc HEPA;", "Đo cường độ sáng bên trong tủ;",
                     "Đo độ ồn của tủ khi vận hành;", "Kiểm tra độ rò rỉ của lọc HEPA khí cấp;",
                     "Kiểm tra hình thái dòng khí của tủ;", "Đo nồng độ hạt bụi trong tủ."],
    }
    return Dossier(
        **_base_kwargs(),
        dossier_type="bsc",
        scope_vn="Giám định tủ cấy",
        scope_en='Measuring/ validating "Laminar flow cabinet"',
        procedure_vn="QUATEST 3 3030:2025, N2.TC.LAF.01",
        procedure_en="QUATEST 3 3030:2025, N2.TC.LAF.01",
        # Căn cứ thêm CAG-003:2022 (mức hình thái dòng khí ⁽⁴⁾) — người dùng chốt 02/11/2026
        standards_vn=f"EN 12469:2000; ISO 14644-1:2015; {AF_BASIS}; Tài liệu kỹ thuật của nhà sản xuất ESCO",
        standards_en=f"EN 12469:2000; ISO 14644-1:2015; {AF_BASIS}; Manufacturer's technical data (ESCO)",
        basis_vn=f"EN 12469:2000; ISO 14644-1:2015; {AF_BASIS}; Tài liệu kỹ thuật của nhà sản xuất ESCO",
        basis_en=f"EN 12469:2000; ISO 14644-1:2015; {AF_BASIS}; Manufacturer's technical data (ESCO)",
        equipments=[EQ["kanomax"], EQ["testo"], EQ["rion"], EQ["ati"], EQ["aerotrak"], EQ["smoke"]],
        locations=[MonitoringLocation(name_vn="Tủ cấy ESCO AVC-4D8", name_en="Laminar flow cabinet ESCO AVC-4D8")],
        photos=_photos("1", "Used"),
        results_data=rd,
        # Đánh giá BCKQ (phù hợp / không phù hợp / không đánh giá) tự lấy từ bảng kết quả TBKQ
        assessment={"basis_vn": f"EN 12469:2000; ISO 14644-1:2015; {AF_BASIS}"},
    )


# ---------------------------------------------------------------------------
# /2 /3 /4 - Tủ hút (form không đánh giá, mức khuyến nghị SEFA 1-2010)
# ---------------------------------------------------------------------------
HOODS = [
    dict(sub=2, brand="HUYHOANG", model="N/A", serial="N/A", dev="TBKN29-2", width=1300, opening=450,
         inflow=[[0.42, 0.44, 0.45, 0.45, 0.44, 0.45, 0.48], [0.48, 0.45, 0.44, 0.48, 0.46, 0.47, 0.51]],
         lux=None, amb=57.5, op=72.8, noise_files=("0171", "0172")),
    dict(sub=3, brand="PROLAB", model="N/A", serial="N/A", dev="TBKN29-4", width=1700, opening=450,
         inflow=[[0.41, 0.46, 0.50, 0.48, 0.51, 0.50, 0.44], [0.45, 0.48, 0.46, 0.45, 0.44, 0.47, 0.46]],
         lux=[[673.8, 753.8, 855.3, 779.1, 690.1], [759.2, 870.9, 1028.4, 973.2, 720.0], [853.4, 1034.9, 1124.6, 1089.9, 895.4]],
         amb=60.9, op=63.5, noise_files=("0174", "0173")),
         # Số trên phiếu là số đo chưa hiệu chỉnh (người dùng xác nhận 30/10/2026): ΔN 2,6 ≈ 3 → trừ 3 → 60,5 dBA.
    dict(sub=4, brand="PROLAB", model="N/A", serial="N/A", dev="TBKN29-5", width=1500, opening=450,
         inflow=[[0.45, 0.40, 0.39, 0.40, 0.39, 0.40, 0.41], [0.40, 0.42, 0.44, 0.46, 0.47, 0.42, 0.41]],
         lux=[[671.3, 756.5, 786.4, 750.7, 715.8], [846.4, 931.5, 955.9, 891.2, 823.5], [915.0, 1011.0, 1087.7, 1035.5, 941.4]],
         amb=57.9, op=65.0, noise_files=("0176", "0175")),
]
INFLOW_REC_VN, INFLOW_REC_EN = "0,35 - 0,51 m/s (*)", "0.35 - 0.51 m/s (*)"


def hood_dossier(h: dict) -> Dossier:
    corr = _hood_corr(h)
    v_avg = _avg(h["inflow"])
    rd = {
        "start_time": START, "end_time": END,
        "bckq_label": "fumehood",
        "cabinet": {"cabinet_name": "Tủ hút", "cabinet_type": "Tủ hút", "brand": h["brand"], "model": h["model"],
                    "serial": h["serial"], "device_id": h["dev"],
                    "chi_tieu_vn": "\n".join(
                        ["1. Đo tốc độ dòng khí qua cửa làm việc;"]
                        + (["2. Đo cường độ sáng bên trong tủ;", "3. Đo độ ồn của tủ khi vận hành."] if h["lux"]
                           else ["2. Đo độ ồn của tủ khi vận hành."]))},
        "openings": [{
            "width_mm": h["width"], "height_mm": h["opening"], "opening_pct": None,
            "avg": round(v_avg, 2), "limit_vn": INFLOW_REC_VN, "limit_en": INFLOW_REC_EN,
            "readings": h["inflow"],
        }],
        "lux": ({"readings": h["lux"], "avg": _avg(h["lux"]), "avg_rounded": int(round(_avg(h["lux"])))}
                if h["lux"] else {"skip": True}),
        "noise": {"ambient": h["amb"], "operating": h["op"], "corrected_dba": round(h["op"] - corr, 1),
                  "file_ambient": h["noise_files"][0], "file_operating": h["noise_files"][1]},
        "airflow_pattern": {"skip": True},
        "request_content_vn": "Giám định tủ hút",
        "request_content_en": 'Measuring/ validating "Fume hood"',
    }
    return Dossier(
        **_base_kwargs(),
        dossier_type="fumehood",
        has_assessment=False,
        scope_vn="Giám định tủ hút",
        scope_en='Measuring/ validating "Fume hood"',
        procedure_vn="QUATEST 3 3030:2025, N2.TC.T.01",
        procedure_en="QUATEST 3 3030:2025, N2.TC.T.01",
        standards_vn="SEFA 1-2010",
        standards_en="SEFA 1-2010",
        equipments=[EQ["kanomax"]] + ([EQ["testo"]] if h["lux"] else []) + [EQ["rion"]],
        locations=[MonitoringLocation(name_vn=f"Tủ hút {h['brand']} ({h['dev']})", name_en=f"Fume hood {h['brand']} ({h['dev']})")],
        photos=_photos(str(h["sub"]), "Use"),
        results_data=rd,
        assessment={},
    )


# ---------------------------------------------------------------------------
# Phiếu tính toán Excel
# ---------------------------------------------------------------------------
THIN = Side(style="thin", color="000000")


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


def build_excel_tu_cay(d: Dossier) -> str:
    out = os.path.join(d.thuchien_dir, f"{d.code} - {d.sub_no} - TinhToanTuCay.xlsx")
    os.makedirs(d.thuchien_dir, exist_ok=True)
    calc = _calc_common(d)
    # Inflow không thực hiện: truyền giá trị giả cho bộ dựng (phương pháp 1), sau đó gạch "-" toàn bộ ô inflow
    calc["inflow"] = {"method": 1, "width_mm": 1000, "height_mm": 1000, "q_readings": [3600]}
    build_bsc_calculation_workbook(calc, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Bảng tính"]
    ws["B4"] = "Tủ cấy (LAF)"
    ws["H5"] = d.results_data["cabinet"]["device_id"]
    # 1. Inflow: không thực hiện -> gạch "-" (giữ khung theo form gốc, AGENTS mục 8B)
    for ref in ("A12", "B12", "C12", "D12", "E12", "F12", "G12", "J12", "A14", "A15", "D31", "F31", "H31", "J31"):
        ws[ref] = "-"
    ws["A10"] = "Phương pháp 1:  DIM - TSI 8375"
    ws["A28"] = "Đánh giá (không thực hiện chỉ tiêu này)"
    # 2. Downflow: mức NSX ESCO 0,45 m/s ± 20 % (giữ công thức trung bình/max/min; độ đồng đều không xét)
    ws["F44"] = "Mức quy định NSX (ESCO)"
    ws["F46"], ws["H46"] = DF_NSX
    ws["J46"] = '=IF(AND(D46>=F46,D46<=H46),"Đạt","Không đạt")'
    ws["J48"] = "-"
    # 3. Cường độ sáng: mức NSX ESCO ≥ 1 110 lux
    ws["F60"] = "Mức quy định NSX (ESCO)"
    ws["F62"], ws["H62"] = LUX_NSX, "-"
    ws["J62"] = '=IF(D62>=F62,"Đạt","Không đạt")'
    # Sheet TBKQ-VN / TBKQ - ENG: mức theo NSX (ký hiệu ⁽³⁾ như TBKQ)
    for name in ("TBKQ-VN", "TBKQ - ENG"):
        if name in wb.sheetnames:
            for row in wb[name].iter_rows():
                for cell in row:
                    if cell.value in ("(0,25-0,50) m/s ⁽¹⁾", "(0.25-0.50) m/s ⁽¹⁾"):
                        cell.value = "(0,36-0,54) m/s ⁽³⁾" if name == "TBKQ-VN" else "(0.36-0.54) m/s ⁽³⁾"
                    elif cell.value == "≥ 750 lux ⁽¹⁾":
                        cell.value = "≥ 1 110 lux ⁽³⁾" if name == "TBKQ-VN" else "≥ 1,110 lux ⁽³⁾"
    # 5. HEPA: chỉ có lọc khí cấp (gasket 0,0024 %, media 0,0012 % -> lấy max)
    ws["D83"] = 0.0024
    ws["H83"] = '=IF(D83<=0.01,"Đạt","Không đạt")'
    ws["D87"], ws["F87"], ws["H87"] = "-", "-", "-"
    # 6. Hình thái dòng khí: 2 vị trí, mức tham chiếu CAG-003:2022 (ký hiệu như TBKQ)
    af_mark = BSCHandler(d).airflow_mark()
    ws["H99"] = f"Mức quy định tham chiếu tại {AF_BASIS}"
    for row, af in zip((101, 106), AF_ROWS):
        ws[f"A{row}"] = af["param_vn"]
        ws[f"D{row}"] = af["result_vn"]
        ws[f"H{row}"] = af["limit_vn"].replace("{ref}", af_mark)
        ws[f"L{row}"] = "Đạt"
        for col in "ADH":
            ws[f"{col}{row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    if "TBKQ-VN" in wb.sheetnames:
        wb["TBKQ-VN"]["B101"] = AF_ROWS[1]["param_vn"]
    if "TBKQ - ENG" in wb.sheetnames:
        en = wb["TBKQ - ENG"]
        for row, af in zip((92, 102), AF_ROWS):
            en[f"B{row}"] = af["param_en"]
            en[f"H{row}"] = af["result_en"]
            en[f"L{row}"] = af["limit_en"].replace("{ref}", af_mark)
            en[f"P{row}"] = "Pass"
    ws["A79"] = None   # ghi chú "(*) Không đánh giá do ồn nền..." không áp dụng
    # 8. Bụi: ISO 5, cả 4 kích thước đạt
    ws["D136"] = '=IF(AND(D130<=D133,F130<=F133,H130<=H133,J130<=J133),"Đạt","Không đạt")'
    wb.save(out)
    print(f"[132488] Excel tủ cấy -> {out}")
    return out


def build_excel_hood(d: Dossier, h: dict) -> str:
    out = os.path.join(d.thuchien_dir, f"{d.code} - {d.sub_no} - TinhToanFumehood.xlsx")
    os.makedirs(d.thuchien_dir, exist_ok=True)
    build_fumehood_calculation_workbook(_calc_common(d), out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Bảng tính"]
    ws["B3"] = '"Tủ hút"'
    ws["B4"] = "Tủ hút"
    ws["K1"] = None
    ws["G12"], ws["J12"] = "-", "-"     # phương pháp 1 (DIM) không dùng
    ws.oddHeader.left.text = f"Số: {d.cert_no_vn}"
    ws.oddHeader.right.text = f"Ngày {d.cert_date}"
    f12 = Font(name="Times New Roman", size=12)
    center = Alignment(horizontal="center", vertical="center")

    # 1. Inflow: lưới 2 x 7 điểm (phiếu ghi chép) -> bỏ gộp ô D:E, G:H để ghi đủ 7 cột C..I
    for rng in ("D20:E20", "G20:H20", "D21:E21", "G21:H21", "D23:E23", "G23:H23", "D24:E24"):
        try:
            ws.unmerge_cells(rng)
        except (KeyError, ValueError):
            pass
    for r in range(20, 25):
        for c in range(3, 10):
            ws.cell(r, c).value = None
            ws.cell(r, c).border = Border(left=THIN if c == 3 else None, right=THIN if c == 9 else None,
                                          top=THIN if r == 20 else None, bottom=THIN if r == 24 else None)
    for row_idx, vals in zip((21, 23), h["inflow"]):
        for i, v in enumerate(vals):
            cell = ws.cell(row_idx, 3 + i)
            cell.value = float(v)
            cell.font, cell.alignment, cell.number_format = f12, center, "0.00"
    ws["A17"] = f"Khoảng mở cửa: {h['opening']} mm"
    ws["J20"] = f"Khoảng mở cửa : {h['opening']}mm"
    ws["C25"] = f"Kích thước: {h['width']} mm"
    ws["F28"] = "Mức khuyến nghị SEFA 1-2010"
    ws["D30"] = "=AVERAGE(C20:I24)"
    ws["F30"], ws["H30"] = 0.35, 0.51
    ws["J30"] = "-"   # form không đánh giá
    # Bỏ khoảng mở 50 % / 25 % (không thực hiện): ẩn dòng 32-62
    for r in range(32, 63):
        ws.row_dimensions[r].hidden = True

    # 2. Cường độ sáng
    if h["lux"]:
        cols = [3, 5, 7, 9, 11]
        for r_i, row_idx in enumerate((66, 68, 70)):
            for c_i, v in enumerate(h["lux"][r_i]):
                cell = ws.cell(row_idx, cols[c_i])
                cell.value = float(v)
                cell.font, cell.alignment, cell.number_format = f12, center, "0.0"
        for r in range(66, 71):
            for c in range(3, 12):
                ws.cell(r, c).border = Border(left=THIN if c == 3 else None, right=THIN if c == 11 else None,
                                              top=THIN if r == 66 else None, bottom=THIN if r == 70 else None)
        ws["J75"] = "-"
        ws["M49"] = 3
    else:
        for r in range(63, 78):
            ws.row_dimensions[r].hidden = True
        ws["M49"] = 2   # Đo độ ồn thành mục 2

    # 3. Độ ồn (giữ công thức I84/I85)
    ws["A82"] = "Khi tủ vận hành, vị trí đo cách mặt trước của tủ 1m"
    ws["I81"], ws["I82"], ws["I83"] = h["amb"], h["op"], _hood_corr(h)
    ws["J89"] = "-"
    ws["A92"] = None
    # 4. Hình thái dòng khí: không thực hiện -> ẩn
    for r in range(93, 109):
        ws.row_dimensions[r].hidden = True
    wb.save(out)
    print(f"[132488] Excel tủ hút /{d.sub_no} -> {out}")
    return out


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--only", type=str, default="", help="Chỉ chạy các số phụ, vd 1,3")
    ap.add_argument("--bckq-rieng", action="store_true",
                    help="BCKQ từng tủ (mặc định gộp 4 tủ vào 1 file ThucHien/BCKQ - tu - Q326D132488.docx)")
    args = ap.parse_args()
    only ={int(x) for x in args.only.split(",") if x.strip()} if args.only else None

    jobs = [(tu_cay_dossier(), None)] + [(hood_dossier(h), h) for h in HOODS]
    assign_cert_numbers([d for d, _ in jobs])   # hồ sơ nhiều báo cáo -> tự thêm /1, /2, /3, /4
    summary = []
    tems = []
    for d, h in jobs:
        if only and d.sub_no not in only:
            continue
        d.ensure_output_directories()
        print(f"\n===== {d.cert_no_vn} ({d.dossier_type}) =====")
        # Gộp BCKQ các tủ vào 1 file (người dùng cho phép 30/10/2026); --only / --bckq-rieng thì làm từng tủ
        bckq = BCKQBuilder(d).build() if (only or args.bckq_rieng) else None
        handler = BSCHandler(d) if h is None else FumeHoodHandler(d)
        if h is None:
            vn, en = handler.build_tbkq(is_en=False), handler.build_tbkq(is_en=True)
        else:
            vn, en = handler.build_tbkq(is_en=False, has_assessment=False), handler.build_tbkq(is_en=True, has_assessment=False)
        tem = handler.build_tem_tu()
        tems.append(tem)
        xlsx = build_excel_tu_cay(d) if h is None else build_excel_hood(d, h)
        docs = [p for p in (bckq, vn, en) if p]
        for p in docs:
            st = check_single_docx(p, code=d.code).get("status")
            summary.append((os.path.basename(p), st))
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            ex = PDFExporter(d)
            # tem lẻ chỉ xuất PDF khi chạy --only (khi chạy đủ 4 tủ thì xuất PDF file gom bên dưới)
            ex.export_files(docs + ([tem] if (tem and only) else []))
            ex.export_excel_files([xlsx])
    if not only and not args.bckq_rieng:
        bckq = BCKQBuilder.build_merged([d for d, _ in jobs])
        summary.insert(0, (os.path.basename(bckq), check_single_docx(bckq, code=jobs[0][0].code).get("status")))
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            PDFExporter(jobs[0][0]).export_files([bckq])
    # Gom tem 4 tủ vào 1 file, 3 tủ / trang (người dùng chốt 30/10/2026) -> TBKQ/132488 - tem tu.docx
    if not only and len(tems) > 1:
        sheet = build_tem_tu_sheet(tems, os.path.join(INPUT_DIR, "TBKQ", "132488 - tem tu.docx"))
        # Gom xong thì bỏ 4 file tem lẻ vừa sinh (người dùng yêu cầu dọn file 30/10/2026)
        if sheet and os.path.isfile(sheet):
            for t in tems:
                if t and os.path.isfile(t) and os.path.abspath(t) != os.path.abspath(sheet):
                    os.remove(t)
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            PDFExporter(jobs[0][0]).export_files([sheet])
    print("\n===== QA =====")
    for name, st in summary:
        print(f"  {st}\t{name}")


if __name__ == "__main__":
    main()
