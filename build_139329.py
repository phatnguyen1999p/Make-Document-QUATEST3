# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D139329 - Công ty Cổ phần Sanofi Việt Nam (10/11/2026) — BSC + tủ cấy (LAF).

2 chứng thư (assign_cert_numbers tự đánh /1, /2):
  /1  Tủ an toàn sinh học cấp 2 LaboGene ApS Mars 900, S/N L.06140002, mã 0264QC00 -> BSCHandler đầy đủ
      (inflow DIM TSI 8380, downflow, lux, ồn, HEPA cấp + thải, hình thái dòng khí, bụi; không đo UVC)
  /2  Tủ cấy sạch LaboGene ApS Fortuna 900 (Vertical laminair air flow), S/N L.06140003, mã 0265QC00
      -> BSCHandler bỏ inflow, lọc khí thải; hình thái dòng khí 2 vị trí theo CAG-003:2022 (AGENTS 8C); bụi; không UVC
BCKQ gộp 1 file (BCKQ - tu - <mã>), tem 2 tủ gom 1 file, Excel tính toán từng tủ.

Nguồn (chép từ Z "Q326D139329 - Sanofi Việt Nam - BSC - LAF", chỉ đọc):
  - Scan/20261109155059847.pdf (8 tr.): 1 phiếu kiểm soát công việc (phân công "Phát + Quân" 02/11), 2 PYC 02/11/2026
    (mục 11, 15 để trống), 3 điều kiện chung, 4 trắng, 5 PYC (bản 2), 6 điều kiện chung, 7 PO E900192188
    (dòng 4 "Giám định tủ an toàn sinh học cấp 2", dòng 5 "Giám định tủ cấy vi sinh"; dòng 1–3 nước uống — hồ sơ
    "Q326D139329 - Sanofi Việt Nam - Nước uống" riêng trên Z), 8 trắng.
  - Scan/20261110174852216.pdf (10 tr.): 1–2 BBLM (09:30–11:30 ngày 10/11/2026, người thực hiện Nguyễn Trần Phát),
    3 phiếu ghi chép BM3030.01 (để trống, gạch), 4–6 Ghi chép thông số tủ /1, 7–9 tủ /2, 10 trắng.
  - Scan/20261110174955390.pdf (4 tr.): 1 phiếu in Lasair III S/N 137662 (zero 10:08, "1" 11:01, "2" 10:44 — 0 hạt/m³
    mọi kênh), 2 ATI 2i tủ /2 (lọc cấp 0,0003 % / 0,0000 %), 3 ATI 2i tủ /1 (thải gasket 0,0032 % / media 0,0008 %;
    cấp gasket 0,0005 % / media 0,0000 %), 4 trùng trang 1.
  - Kq Do On/1_1 octave.xlsx (Rion NL-53): 0194 63,8 dBA (10:54), 0195 70,2 dBA 15,7 s (bỏ dở), 0196 62,0 dBA (11:13),
    0197 46,4 dBA (11:20, ồn nền).
  - Ảnh nhãn tủ: Photos/1/20261110_111812.jpg (Mars 900, Biohazard safety cabinet class II, L.06140002, 2014, working
    area 510 × 900 mm), Photos/2/20261110_111833.jpg (Fortuna 900, Vertical laminair air flow, L.06140003, 2014,
    600 × 900 mm). Ảnh chèn: Photos/<n>/Use (người dùng chọn).
  - Old test result/Q326D045837 - 1, 2 (chứng thư cũ cùng 2 tủ, 05/2026) — chỉ tham khảo tên/mã tủ.
  - MST 0312233458 (tax_helper): tên quốc tế "SANOFI VIETNAM SHAREHOLDING COMPANY" (PO cũng ghi tên này).
  -> Edoc: BBLM = 174852216 tr. 1–2; Ghi Chep = 174852216 tr. 3–9 + 174955390 tr. 1–3; Khac = 155059847 (bỏ trắng)

Quyết định của người dùng (10/11/2026):
  - Độ ồn vận hành tủ /1 = 63,8 dBA theo máy (file Rion 0194, ảnh màn hình), phiếu ghi 63,4.
  - Tên EN khách hàng "Sanofi Vietnam Joint Stock Company" như chứng thư cũ 045837 (MST ghi "Shareholding").
  - Tủ cấy /2 đánh giá theo IEST-RP-CC002.4 như 045837/2: downflow 0,40–0,50 m/s, lux ≥ 540, ồn ≤ 70 dBA
    (ghi chú ⁽¹⁾ "Mức quy định tham chiếu tại IEST-RP-CC002.4"); hình thái dòng khí CAG-003:2022; bụi ISO 14644-1.
  - Độ rọi đọc 1020,4 (tủ /1); 890,1 và 1002,7 (tủ /2).
  - Giữ: ngôn ngữ Việt - Anh, bụi d ≥ 5 µm mức "-", máy đếm hạt Lasair III 110.
  - 11/11/2026: tên thiết bị /1 = "Tủ an toàn sinh học" theo tên khách hàng đặt (EN "Biosafety cabinet"); nội dung
    yêu cầu giữ "Giám định tủ an toàn sinh học cấp 2" (PO, BBLM).
  - 11/11/2026: inflow phương pháp lưu lượng (DIM) đo bằng TSI 8380 (người dùng điều chỉnh, phiếu ghi 8375);
    mặc định BSC dùng TSI 8380 khi đo inflow theo lưu lượng.

Suy luận của AI (còn lại):
  - Tủ /2: tên "Tủ cấy sạch" (BBLM sửa tay), nội dung yêu cầu "Giám định tủ cấy vi sinh" (PO, BBLM) — như 045837/2.
  - Cửa làm việc tủ /1: 0,20 m × 0,90 m = 0,18 m² (phiếu); nồng độ phun 37 µg/L (6 jet), tủ /2 14 µg/L (2 jet).
  - Tên đầy đủ IEST-RP-CC002.4 ở Căn cứ: "Unidirectional-Flow, Clean-Air Devices" (Core/Tools/standards.py).

Chạy:  python build_139329.py            (DOCX + XLSX, không PDF)
       python build_139329.py --pdf      (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: (network_dir để trống; chép lên Z bằng ai_files.py theo AGENTS mục 0).
"""

import os
import sys
import argparse
import datetime as _dt

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_co_Z_"))

import openpyxl
from openpyxl.styles import Alignment

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.bsc import BSCHandler
from Core.Builders.tag_context import build_tem_tu_sheet, shrink_cover_break
from Core.Tools.bsc_calc import build_bsc_calculation_workbook
from Core.Tools.bsc_cal_blocks import finalize_bsc_workbook
from Core.Tools.scan_splitter import write_edoc
from Core.Tools.noise_corr import noise_correction_value
from Core.Tools.qa_check import check_single_docx
from Core.Tools.excel_borders import strike_box
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D139329"
REG_DATE = "02/11/2026"
SAMPLE_DATE = "10/11/2026"
CERT_DATE = _dt.date.today().strftime("%d/%m/%Y")   # không có KQTN -> ngày hiện tại (AGENTS mục 2)
START, END = "09:30", "11:30"
INPUT_DIR = os.path.join(WORKSPACE, "Input", "139329")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photos")

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần Sanofi Việt Nam",
    name_en="Sanofi Vietnam Joint Stock Company",   # người dùng chốt 10/11/2026 (như chứng thư cũ 045837)
    address_vn="Lô I-8-2, đường D8, Khu công nghệ cao, phường Tăng Nhơn Phú, Thành phố Hồ Chí Minh",
    address_en="Lot I-8-2, D8 Street, Saigon Hi-Tech Park, Tang Nhon Phu Ward, Ho Chi Minh City",
)
LOCATION_VN = ("Nhà máy của Công ty Cổ phần Sanofi Việt Nam - Lô I-8-2, đường D8, Khu công nghệ cao,"
               " phường Tăng Nhơn Phú, Thành phố Hồ Chí Minh")   # BBLM mục 5
LOCATION_EN = ("Factory of Sanofi Vietnam Joint Stock Company - Lot I-8-2, D8 Street, Saigon Hi-Tech Park,"
               " Tang Nhon Phu Ward, Ho Chi Minh City")

PARTICLE_FILE_ZERO = "139329 zero (10:08)"

# ---------------------------------------------------------------------------
# Số liệu phiếu ghi chép (10/11/2026)
# ---------------------------------------------------------------------------
BSC = dict(
    inflow_q=[282, 284, 282, 284, 283], h_mm=200, w_mm=900,          # DIM TSI 8380, 0,18 m²; phiếu 0,44 m/s
    downflow=[[0.36, 0.32, 0.33, 0.30, 0.30],
              [0.32, 0.30, 0.30, 0.30, 0.31],
              [0.31, 0.31, 0.35, 0.35, 0.35]],                         # phiếu TB 0,32
    lux=[[1020.4, 1146.0, 989.9], [956.2, 1084.0, 873.8], [845.2, 975.7, 813.4]],
    noise_amb=46.4, noise_op=63.8, noise_files=("0197", "0194"),       # phiếu ghi 63,4 — file/ảnh máy 63,8
    hepa=dict(downflow_pct=0.0005, exhaust_pct=0.0032, aerosol_conc=37,
              downflow_gasket_pct=0.0005, downflow_media_pct=0.0000,
              exhaust_gasket_pct=0.0032, exhaust_media_pct=0.0008),
    particle=[0, 0, 0, 0], particle_file="1 (11:01)",
)
LAF = dict(
    downflow=[[0.46, 0.49, 0.48, 0.50, 0.43],
              [0.42, 0.44, 0.46, 0.44, 0.40],
              [0.41, 0.39, 0.43, 0.42, 0.40]],                         # tổng hàng 2,36 / 2,16 / 2,05 khớp phiếu, TB 0,44
    lux=[[785.1, 890.1, 761.3], [845.5, 1002.7, 826.1], [870.7, 1038.5, 866.8]],   # 890,1 người dùng đọc 10/11
    noise_amb=46.4, noise_op=62.0, noise_files=("0197", "0196"),
    hepa=dict(downflow_pct=0.0003, exhaust_pct=None, skip_exhaust=True, aerosol_conc=14,
              downflow_gasket_pct=0.0003, downflow_media_pct=0.0000),
    particle=[0, 0, 0, 0], particle_file="2 (10:44)",
)

# Hình thái dòng khí tủ cấy — AGENTS 8C (mẫu 045837/2, 132488/1): 2 vị trí, mức tham chiếu CAG-003:2022
AF_BASIS = "CAG-003:2022"
# Tủ cấy /2: mức tham chiếu IEST-RP-CC002.4 như chứng thư cũ 045837/2 (người dùng chốt 10/11/2026)
LAF_BASIS = "IEST-RP-CC002.4"
LAF_DF_RANGE = (0.40, 0.50)
LAF_LUX_MIN = 540
LAF_NOISE_MAX = 70
AF_ROWS = [
    {"param_vn": "Bên trong không gian tủ\n- Dọc đường trung tâm, phía trên khoảng mở làm việc"
                 "\n- Phía sau cửa mở làm việc\n- Dọc theo mép cửa mở làm việc",
     "param_en": "Inside cabinet workspace\n- Along center line, above work opening"
                 "\n- Behind work sash\n- Along edge of work sash",
     "result_vn": "- Dòng khói hướng xuống, không có điểm tù;\n- Dòng khói thoát ra khỏi mặt thao tác.",
     "result_en": "- Smooth downward flow, no dead spots;\n- Smoke exited off the work surface.",
     "limit_vn": "- Dòng khói hướng xuống, không có điểm tù;\n- Dòng khói thoát ra khỏi mặt thao tác. {ref}",
     "limit_en": "- Smooth downward flow, no dead spots;\n- Smoke exits off the work surface. {ref}",
     "height": 1750},
    {"param_vn": "- Khu vực dọc theo chu vi phía bên ngoài của vùng mở cửa tủ",
     "param_en": "- Area along the outer perimeter of the work opening",
     "result_vn": "- Dòng khói hướng ra phía trước tủ.",
     "result_en": "- Smoke flowed outward to the front of the cabinet.",
     "limit_vn": "- Dòng khói hướng ra phía trước tủ. {ref}",
     "limit_en": "- Smoke flows outward to the front of the cabinet. {ref}",
     "height": 820},
]


def _avg(grid):
    vals = [v for row in grid for v in row]
    return sum(vals) / len(vals)


def _photos(sub: str):
    d = os.path.join(PHOTO_DIR, sub, "Use")
    if not os.path.isdir(d):
        return []
    return [{"path": os.path.join(d, f), "caption_vn": "", "caption_en": ""}
            for f in sorted(os.listdir(d)) if f.lower().endswith((".jpg", ".jpeg", ".png"))]


def _particle(vals, file_ghi):
    """Nồng độ hạt cộng dồn (Lasair III, 5 phút × 28,3 L/phút), giới hạn ISO 5 (ISO 14644-1:2015 Bảng 1);
    d ≥ 5 µm: ISO 5 không có giới hạn -> mức "-", không đánh giá (AGENTS mục 15)."""
    def row(lbl_vn, lbl_en, v, lim_vn, lim_en, lim_num):
        if lim_num is None:
            ass_vn = ass_en = "-"
        else:
            ok = v <= lim_num
            ass_vn, ass_en = ("Đạt", "Pass") if ok else ("Không đạt", "Fail")
        return (lbl_vn, lbl_en, f"{v} hạt/m³", f"{v} particles/m³", lim_vn, lim_en, ass_vn, ass_en)
    return {
        "clean_class": "ISO Class 5", "sample_volume_l": 141.6, "sample_time_min": 5,
        "readings": vals, "file_zero": PARTICLE_FILE_ZERO, "file_ghi": file_ghi,
        "sizes": [
            row("Kích thước hạt d ≥ 0,3 µm", "Particle size d ≥ 0.3 µm", vals[0], "≤ 10 200\nhạt/m³ ⁽²⁾", "≤ 10,200\nparticles/m³ ⁽²⁾", 10200),
            row("Kích thước hạt d ≥ 0,5 µm", "Particle size d ≥ 0.5 µm", vals[1], "≤ 3 520\nhạt/m³ ⁽²⁾", "≤ 3,520\nparticles/m³ ⁽²⁾", 3520),
            row("Kích thước hạt d ≥ 1,0 µm", "Particle size d ≥ 1.0 µm", vals[2], "≤ 832\nhạt/m³ ⁽²⁾", "≤ 832\nparticles/m³ ⁽²⁾", 832),
            row("Kích thước hạt d ≥ 5,0 µm", "Particle size d ≥ 5.0 µm", vals[3], "-", "-", None),
        ],
    }


def _base_kwargs():
    return dict(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=True,
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        client=CLIENT, personnel=PersonnelInfo(executor="Phát"),
        input_dir=INPUT_DIR, tbkq_dir=os.path.join(INPUT_DIR, "TBKQ"), output_dir="", network_dir="",
        dossier_type="bsc",
    )


def bsc_dossier() -> Dossier:
    s = BSC
    v_in = sum(s["inflow_q"]) / len(s["inflow_q"]) / (s["h_mm"] * s["w_mm"] / 1e6) / 3600
    corr = noise_correction_value(s["noise_amb"], s["noise_op"])
    lux_avg = _avg(s["lux"])
    rd = {
        "start_time": START, "end_time": END, "file_label": "BSC", "bckq_label": "bsc",
        "cabinet": {
            # tên thiết bị theo khách hàng đặt (nhãn trên tủ "Tủ an toàn sinh học") — người dùng chốt 11/11/2026
            "cabinet_name": "Tủ an toàn sinh học", "cabinet_name_en": "Biosafety cabinet",
            "cap_tu": "II", "bsc_class": 2,
            "brand": "LaboGene ApS", "model": "Mars 900", "serial": "L.06140002", "device_id": "0264QC00",
            "seal_item_name": "“Tủ an toàn sinh học”",
        },
        "inflow": {"method": 1, "method_name": "DIM TSI 8380", "q_readings": s["inflow_q"],
                   "width_mm": s["w_mm"], "height_mm": s["h_mm"], "velocity": round(v_in, 2), "scan_velocity": 0.44},
        "downflow": {"readings": s["downflow"], "avg": round(_avg(s["downflow"]), 2), "tolerance_pct": 20,
                     "width_mm": 900},
        "lux": {"readings": s["lux"], "avg": lux_avg, "avg_rounded": int(round(lux_avg))},
        "noise": {"ambient": s["noise_amb"], "operating": s["noise_op"],
                  "corrected_dba": round(s["noise_op"] - corr, 1),
                  "file_ambient": s["noise_files"][0], "file_operating": s["noise_files"][1]},
        "hepa": dict(s["hepa"]),
        "airflow_pattern": {},
        "particle": _particle(s["particle"], s["particle_file"]),
    }
    return Dossier(
        **_base_kwargs(),
        scope_vn="Giám định tủ an toàn sinh học cấp 2", scope_en="Inspection of Biosafety cabinet class II",
        procedure_vn="QUATEST 3 3030:2025, N2.TC.BSC.02", procedure_en="QUATEST 3 3030:2025, N2.TC.BSC.02",
        standards_vn="EN 12469:2000; ISO 14644-1:2015", standards_en="EN 12469:2000; ISO 14644-1:2015",
        basis_vn="EN 12469:2000; ISO 14644-1:2015", basis_en="EN 12469:2000; ISO 14644-1:2015",
        equipments=[eq("tsi_8380"), eq("kanomax_6501"), eq("testo_440"), eq("rion_nl53"),
                    eq("ati_6d_2i"), eq("smoke"), eq("lasair_iii_110")],
        locations=[MonitoringLocation(name_vn="Tủ an toàn sinh học LaboGene Mars 900",
                                      name_en="Biosafety cabinet LaboGene Mars 900")],
        photos=_photos("1"),
        results_data=rd,
        assessment={"basis_vn": "EN 12469:2000; ISO 14644-1:2015"},
    )


def laf_dossier() -> Dossier:
    s = LAF
    corr = noise_correction_value(s["noise_amb"], s["noise_op"])
    lux_avg = _avg(s["lux"])
    pass_vn, pass_en = "Đạt", "Pass"
    rd = {
        "start_time": START, "end_time": END, "file_label": "Tu cay", "bckq_label": "tu cay",
        "cabinet": {
            "cabinet_name": "Tủ cấy sạch", "cabinet_name_en": "Laminar flow cabinet",
            "brand": "LaboGene ApS", "model": "Fortuna 900", "serial": "L.06140003", "device_id": "0265QC00",
            "seal_item_name": "“Tủ cấy sạch”",
        },
        "limit_basis": LAF_BASIS,
        "seal_standard": f"{LAF_BASIS};\nISO 14644-1:2015",
        "inflow": {"skip": True},
        "downflow": {"readings": s["downflow"], "avg": round(_avg(s["downflow"]), 2), "tolerance_pct": 20,
                     "width_mm": 900, "depth_mm": 600, "limit_range": LAF_DF_RANGE},
        "lux": {"readings": s["lux"], "avg": lux_avg, "avg_rounded": int(round(lux_avg)), "limit_min": LAF_LUX_MIN},
        "noise": {"ambient": s["noise_amb"], "operating": s["noise_op"],
                  "corrected_dba": round(s["noise_op"] - corr, 1), "limit_max": LAF_NOISE_MAX,
                  "file_ambient": s["noise_files"][0], "file_operating": s["noise_files"][1]},
        "hepa": dict(s["hepa"]),
        "airflow_pattern": {"basis_vn": AF_BASIS, "basis_en": AF_BASIS,
                            "rows": [dict(r, assess_vn=pass_vn, assess_en=pass_en) for r in AF_ROWS]},
        "particle": _particle(s["particle"], s["particle_file"]),
    }
    std = f"{LAF_BASIS}; {AF_BASIS}; ISO 14644-1:2015"
    return Dossier(
        **_base_kwargs(),
        scope_vn="Giám định tủ cấy vi sinh", scope_en="Inspection of Laminar flow cabinet",
        procedure_vn="QUATEST 3 3030:2025, N2.TC.LAF.01", procedure_en="QUATEST 3 3030:2025, N2.TC.LAF.01",
        standards_vn=std, standards_en=std, basis_vn=std, basis_en=std,
        equipments=[eq("kanomax_6501"), eq("testo_440"), eq("rion_nl53"), eq("ati_6d_2i"),
                    eq("smoke"), eq("lasair_iii_110")],
        locations=[MonitoringLocation(name_vn="Tủ cấy sạch LaboGene Fortuna 900",
                                      name_en="Laminar flow cabinet LaboGene Fortuna 900")],
        photos=_photos("2"),
        results_data=rd,
        assessment={"basis_vn": std},
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


def _particle_cells(ws, d: Dossier):
    """Bụi: d ≥ 5 µm ISO 5 không có giới hạn -> J133 '-', đánh giá theo 3 kênh còn lại; ghi file máy đếm."""
    ws["J133"] = "-"
    ws["D136"] = '=IF(AND(D130<=D133,F130<=F133,H130<=H133),"Đạt","Không đạt")'


def build_excel_bsc(d: Dossier) -> str:
    out = os.path.join(d.thuchien_dir, f"{d.code} - {d.sub_no} - TinhToanBSC.xlsx")
    os.makedirs(d.thuchien_dir, exist_ok=True)
    build_bsc_calculation_workbook(_calc_common(d), out, cals={}, finalize=False)
    wb = openpyxl.load_workbook(out)
    ws = wb["Bảng tính"]
    ws["B4"] = "II"
    ws["A79"] = None
    _particle_cells(ws, d)
    wb.save(out)
    finalize_bsc_workbook(out, {})
    print(f"[139329] Excel /{d.sub_no} -> {out}")
    return out


def build_excel_laf(d: Dossier) -> str:
    out = os.path.join(d.thuchien_dir, f"{d.code} - {d.sub_no} - TinhToanTuCay.xlsx")
    os.makedirs(d.thuchien_dir, exist_ok=True)
    calc = _calc_common(d)
    # Inflow không thực hiện: bộ dựng cần số giả (phương pháp 1) rồi gạch "-" (giữ khung form, AGENTS 8B)
    calc["inflow"] = {"method": 1, "width_mm": 1000, "height_mm": 1000, "q_readings": [3600]}
    build_bsc_calculation_workbook(calc, out, cals={}, finalize=False)
    wb = openpyxl.load_workbook(out)
    ws = wb["Bảng tính"]
    ws["B4"] = "Tủ cấy (LAF)"
    for ref in ("A12", "B12", "C12", "D12", "E12", "F12", "G12", "J12", "A14", "A15", "D31", "F31", "H31", "J31"):
        ws[ref] = "-"
    ws["A10"] = "Phương pháp 1:  DIM - TSI 8380"
    ws["A28"] = "Đánh giá (không thực hiện chỉ tiêu này)"
    # Mức IEST-RP-CC002.4 (người dùng chốt 10/11/2026): downflow 0,40–0,50; lux ≥ 540; ồn ≤ 70 — đánh giá vẫn là công thức
    for ref in ("F44", "F60", "F74"):
        ws[ref] = f"Mức quy định {LAF_BASIS}"
        # thu chữ cho vừa ô (nới hàng làm ngắt trang cắt ngang bảng downflow)
        ws[ref].alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    ws["F46"], ws["H46"] = LAF_DF_RANGE
    ws["F62"] = LAF_LUX_MIN
    ws["H76"] = LAF_NOISE_MAX
    # HEPA: chỉ có lọc khí cấp
    ws["D87"], ws["F87"], ws["H87"] = "-", "-", "-"
    # Hình thái dòng khí: 2 vị trí, mức tham chiếu CAG-003:2022 (ký hiệu như TBKQ)
    af_mark = BSCHandler(d).airflow_mark()
    ws["H99"] = f"Mức quy định tham chiếu tại {AF_BASIS}"
    for row, af in zip((101, 106), AF_ROWS):
        ws[f"A{row}"] = af["param_vn"]
        ws[f"D{row}"] = af["result_vn"]
        ws[f"H{row}"] = af["limit_vn"].replace("{ref}", af_mark)
        ws[f"L{row}"] = "Đạt"
        for col in "ADH":
            ws[f"{col}{row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for r in range(101, 106):          # ô A101:C105 6 dòng chữ -> nới hàng cho khỏi cắt chữ
        ws.row_dimensions[r].height = 27
    ws["A79"] = None
    _particle_cells(ws, d)
    wb.save(out)
    finalize_bsc_workbook(out, {})
    print(f"[139329] Excel /{d.sub_no} -> {out}")
    return out


def _edoc():
    write_edoc(INPUT_DIR, "139329", {
        "BBLM": [(os.path.join("Scan", "20261110174852216.pdf"), [1, 2])],
        "Ghi Chep": [(os.path.join("Scan", "20261110174852216.pdf"), [3, 4, 5, 6, 7, 8, 9]),
                     (os.path.join("Scan", "20261110174955390.pdf"), [1, 2, 3])],
        "Khac": [(os.path.join("Scan", "20261109155059847.pdf"), None)],
    })


def _compact_cover(path: str, compact):
    import docx
    doc = docx.Document(path)
    shrink_cover_break(doc, compact=compact)
    doc.save(path)


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()

    jobs = [(bsc_dossier(), build_excel_bsc, True), (laf_dossier(), build_excel_laf, 2)]
    assign_cert_numbers([d for d, _, _ in jobs])
    _edoc()
    summary, tems, pdf_docs, xlsxs = [], [], [], []
    for d, xl_fn, compact in jobs:
        d.ensure_output_directories()
        print(f"\n===== {d.cert_no_vn} =====")
        h = BSCHandler(d)
        vn, en = h.build_tbkq(is_en=False), h.build_tbkq(is_en=True)
        for p in (vn, en):
            _compact_cover(p, compact)
        tems.append(h.build_tem_tu())
        xlsxs.append(xl_fn(d))
        for p in (vn, en):
            summary.append((os.path.basename(p), check_single_docx(p, code=d.code).get("status")))
        pdf_docs += [vn, en]
    ds = [d for d, _, _ in jobs]
    bckq = BCKQBuilder.build_merged(ds)
    summary.insert(0, (os.path.basename(bckq), check_single_docx(bckq, code=CODE).get("status")))
    sheet = build_tem_tu_sheet(tems, os.path.join(INPUT_DIR, "TBKQ", "139329 - tem tu.docx"))
    if sheet and os.path.isfile(sheet):
        for t in tems:
            if t and os.path.isfile(t) and os.path.abspath(t) != os.path.abspath(sheet):
                os.remove(t)
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(ds[0])
        ex.export_files(pdf_docs + [bckq, sheet])
        ex.export_excel_files(xlsxs)
    print("\n===== QA =====")
    for name, st in summary:
        print(f"  {st}\t{name}")


if __name__ == "__main__":
    main()
