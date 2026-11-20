# -*- coding: utf-8 -*-
r"""
Q326D137038 — Công ty Cổ phần Đầu tư Nước Tân Hiệp — nước sinh hoạt 1 mẫu (QCVN 01-1:2024/BYT), 1 chứng thư (mã trần).
CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.1.Tân Hiệp\Q326D137038 - Nhà máy nước Tân Hiệp- Nước sinh hoạt - 16.11.2026` về Input\137038):
  - `0. BBLM - 137038.pdf` (2 trang): ngày 16/11/2026, PYC Q326D137038 ngày 27/10/2026, 09:45 – 10:15; "Lấy mẫu, giám định
    chất lượng nước sinh hoạt theo QCVN 01-1:2024/BYT"; địa điểm "Phòng thí nghiệm của Nhà máy nước Tân Hiệp 2"; Mettler
    Seven2Go S2 + Hach DR900; QTGĐ 3030.2025, N2.TC.N.02; chỉ định: chị Mai Thị Hậu; 137038-NSH "Vòi nước dịch vụ - Phòng
    thí nghiệm" 10:05, pH 7,81 / 30,0 °C, clo dư tự do 0,96 mg/L, mùi: không có mùi lạ. Người thực hiện: Hồ Nguyễn Huy.
  - `0. Phiếu YCTN -Q326D137038.N2.1.pdf` (MTTN26022165): dự kiến trả 26/11/2026; mẫu QC (amoni).
  - Ảnh `Q326D137038 - photo\0. Photo\Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026): quy trình NSH để trống (tô vàng); chức vụ "trưởng ban hóa nghiệm" đúng.

Suy luận của AI (cần người dùng xác nhận):
  1. Danh sách 10 thông số nhóm A, phương pháp, mức theo khung nước sinh hoạt 134486/1 (QCVN 01-1:2024/BYT).
  2. Quy trình NSH để "………" tô vàng như 134486 (BBLM ghi N2.TC.N.02).

Chạy:  python build_137038.py [--pdf]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, nsh_base

IN = os.path.join(WORKSPACE, "Input", "137038")

CONFIG = dict(
    code6="137038", loai="nuoc sinh hoat", kind="nsh",
    client_vn="Công ty Cổ phần Đầu tư Nước Tân Hiệp",
    address_vn="Số 226, đường Dương Công Khi, xã Hóc Môn, Thành phố Hồ Chí Minh",
    reg_date="27/10/2026", dates=["16/11/2026"], time_bckq="Ngày 16/11/2026 từ 09:45 đến 10:15",
    site_vn="Phòng thí nghiệm của Nhà máy nước Tân Hiệp 2",
    scope_vn="Lấy mẫu và giám định chất lượng nước sinh hoạt",
    procedure="QUATEST 3 3030:2025, ………",       # BBLM N2.TC.N.02 — để trống như 134486
    equip=["mettler_s2", "hach_dr900"], executor="Huy",
    witness_vn="chị Mai Thị Hậu – trưởng ban hóa nghiệm",
    req_date="26/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2",
                   "Clo dư tự do": "Hach DR900", "Mùi": "-"},
    samples=[
        dict(nsh_base(), ky_hieu="137038-NSH", kqtn_key="NSH", time="10:05",
             vi_tri="Vòi nước dịch vụ - Phòng thí nghiệm",
             field={"pH": "7,81", "Nhiệt độ": "30,0", "Clo dư tự do": "0,96", "Mùi": "Không có mùi lạ"},
             photo=os.path.join(IN, "Q326D137038 - photo", "0. Photo", "Vị trí lấy mẫu.jpg")),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
