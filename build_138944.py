# -*- coding: utf-8 -*-
r"""
Q326D138944 — Công ty TNHH Bosch Việt Nam — nước thải sau HTXL tại hố ga đấu nối KCN (NT1), 1 chứng thư (mã trần) +
BCKQ. CHƯA CÓ KQTN -> bản khung "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D138944 - Bosch - NT` về Input\138944, chỉ đọc):
  - `BBLLM.pdf` (2 trang, scan): BBLM ngày 11/11/2026, PYC Q326D138944 ngày 30/10/2026; địa chỉ = địa điểm "Khu công nghiệp
    Long Thành, xã An Phước, Thành phố Đồng Nai"; "Lấy mẫu và phân tích chất lượng nước thải"; 9h20 – 10h00; Hanna HI-98190
    (pH), HACH DR300 (clo), Hanna HI98192 (EC/TDS); QUATEST3 3030.2025; N2.TC.N.02; N2.TC.N.06; chỉ định: anh Lương Hiếu
    Trúc (người dùng xác nhận 18/11/2026) – FCM; NT1 "Nước thải sau HTXL tại hố ga đấu nối KCN"; đo nhanh 9h44: pH 7,622/7,610/7,617,
    clo dư 0,09×3 mg/L, TDS 426/430/429 mg/L, nhiệt độ 30,5/30,4/30,5 °C; GĐV Nguyễn Minh Quân.
  - `138944 - Biên bản hiện trường.xlsx` sheet "Phụ lục I": 30 chỉ tiêu (Cr³⁺, α, β, độ màu, Cr⁶⁺, F⁻, Cl⁻, BVTV P, PCB,
    COD, T-N, T-P, amoni, S²⁻, CN⁻, Coliform, dầu mỡ ĐTV, dầu mỡ khoáng, phenol, As, Pb, Hg, Cd, Cu, Zn, Ni, Mn, Fe,
    TSS, BOD5). Ảnh `Photo\image_01 … 12.jpg`.

Quyết định của người dùng (18/11/2026): người chỉ định "Lương Hiếu Trúc"; quy chuẩn chưa có -> Căn cứ "-", tô vàng.

Suy luận của AI (đã được người dùng xem 18/11/2026):
  1. (bỏ — người dùng xác nhận tên Lương Hiếu Trúc)
  2. Không có quy chuẩn / mức -> Căn cứ "-", không đánh giá. Clo dư 0,09 -> "< 0,1" (AGENTS 8C).
  3. pH, nhiệt độ, TDS = trung bình 3 lần (số từng lần ghi ở BCKQ). Ảnh vị trí hố ga: `image_10.jpg` (image_10 – 12 cùng cảnh).

Chạy:  python build_138944.py [--pdf]
"""
import argparse
import datetime
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER as P
from nt_params import params

IN = os.path.join(WORKSPACE, "Input", "138944")
NAMES = ["Crôm III (Cr³⁺)", "Tổng hoạt độ phóng xạ α", "Tổng hoạt độ phóng xạ β", "Độ màu", "Crôm VI (Cr⁶⁺)",
         "Florua (F⁻)", "Clorua (Cl⁻)", "Hóa chất BVTV phospho hữu cơ", "Polyclo biphenyl (PCB)", "COD", "Tổng nitơ",
         "Tổng phốt pho", "Amoni (NH₄⁺ tính theo N)", "Sulfua (S²⁻)", "Xyanua (CN⁻)", "Tổng Coliform",
         "Dầu mỡ động thực vật", "Tổng dầu mỡ khoáng", "Tổng phenol", "Asen (As)", "Chì (Pb)", "Thủy ngân (Hg)",
         "Cadimi (Cd)", "Đồng (Cu)", "Kẽm (Zn)", "Niken (Ni)", "Mangan (Mn)", "Sắt (Fe)",
         "Tổng chất rắn lơ lửng (TSS)", "BOD₅"]
F = "Đo tại hiện trường"
FIELD = [("pH", "-", F), ("Nhiệt độ", "°C", F), ("TDS", "mg/L", F), ("Clo dư", "mg/L", F)]
SITE = "Khu công nghiệp Long Thành, xã An Phước, Thành phố Đồng Nai"
WITNESS = "anh Lương Hiếu Trúc – FCM"

CONFIG = dict(
    code6="138944", loai="nuoc thai", kind="nt",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2)
    client_vn="Công ty TNHH Bosch Việt Nam", address_vn=SITE,
    reg_date="30/10/2026", dates=["11/11/2026"], time_bckq="Ngày 11/11/2026 từ 09:20 đến 10:00",
    site_vn=f"Công ty TNHH Bosch Việt Nam - {SITE}",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.06",
    equip=["hanna_hi98190", "hach_dr300", "hanna_hi98192"], executor="Quân",
    witness_vn=WITNESS, req_date=P, hl_can_cu=True,
    field_devices={"pH": "Hanna HI 98190", "Nhiệt độ": "Hanna HI 98190", "TDS": "Hanna HI 98192", "Clo dư": "Hach DR300"},
    bckq_notes=["- Số đo hiện trường ghi là trung bình 3 lần đo trên BBLM: pH 7,622/7,610/7,617; nhiệt độ 30,5/30,4/30,5 °C; "
                "TDS 426/430/429 mg/L; clo dư 0,09/0,09/0,09 mg/L (dưới 0,1 mg/L -> \"< 0,1\")."],
    samples=[dict(kind="nt", ky_hieu="NT1", kqtn_key="NT1", time="09:44",
                  vi_tri="Nước thải sau hệ thống xử lý nước thải tại hố ga đấu nối khu công nghiệp",
                  field={"pH": "7,62", "Nhiệt độ": "30,5", "TDS": "428", "Clo dư": "< 0,1"},
                  field_tbkq=["pH", "Nhiệt độ", "TDS", "Clo dư"],
                  equip=["hanna_hi98190", "hach_dr300", "hanna_hi98192"], params=FIELD + params(NAMES),
                  chi_tieu="pH, nhiệt độ, TDS, clo dư (đo tại hiện trường) và 30 thông số theo yêu cầu khách hàng",
                  photo=os.path.join(IN, "Photo", "image_10.jpg"))],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
