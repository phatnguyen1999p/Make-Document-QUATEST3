# -*- coding: utf-8 -*-
r"""
Q326D137037 — Công ty Cổ phần Đầu tư Nước Tân Hiệp — nước sinh hoạt 1 mẫu (QCVN 01-1:2024/BYT), 1 chứng thư (mã trần).
ĐÃ CÓ KQTN MTTN26021357 ngày 17/11/2026.

Nguồn (chép từ Z `0.1.Tân Hiệp\Q326D137037 - Nhà máy nước Tân Hiệp- Nước sinh hoạt - 09.11.2026` về Input\137037, chỉ đọc):
  - `137037 - BBLM.pdf` (2 trang, scan): ngày 09/11/2026, PYC Q326D137037 ngày 27/10/2026, 09:15 – 10:00; "Lấy mẫu, giám định
    chất lượng nước sinh hoạt theo QCVN 01-1:2024/BYT"; địa điểm "Phòng thí nghiệm của Nhà máy nước Tân Hiệp 2"; Mettler
    Seven2Go S2 + Hach DR900; QTGĐ 3030.2025, N2.TC.N.02; chỉ định: chị Mai Thị Hậu – trưởng ban hóa nghiệm; 137037-NSH
    "Vòi nước dịch vụ - Phòng thí nghiệm" 09:35, pH 7,71 / 29,9 °C, clo dư tự do 0,93 mg/L, mùi: không có mùi lạ;
    kiểm tra pH 4,03 / 7,00 / 10,00 đạt. Người thực hiện: Hồ Nguyễn Huy.
  - `PYC-Q326D137037.N2.1.pdf` (MTTN26021357): nhận 06/11/2026, dự kiến trả 18/11/2026.
  - `KQ-Q326D137037-MTTN26021357.pdf` (17/11/2026, nhận mẫu 09/11/2026): permanganat, amoni, As, độ màu, độ đục KPH;
    E. coli, Coliforms < 1.
  - Khung làm tay của người dùng `3. TBKQ - Q326D137037.docx`, `2. BCKQTĐ - Q326D137037.docx` (bảng người chứng kiến còn
    số liệu hồ sơ cũ 104561 — không dùng).
  - Ảnh `photo\Vị trí lấy mẫu.jpg`.

Quyết định của người dùng: 18/11/2026 "làm … có KQTN rồi"; quy trình "………" tô vàng; địa điểm theo BBLM (đều "oke"). (137038 cùng khách, 16/11/2026: quy trình NSH để trống, tô vàng.)

Suy luận của AI (cần người dùng xác nhận):
  1. Quy trình để "………" tô vàng như 137038 (BBLM ghi N2.TC.N.02, khung làm tay ghi N2.TC.N.01).
  2. Địa điểm theo BBLM "Phòng thí nghiệm của Nhà máy nước Tân Hiệp 2" (khung: "Nhà máy nước Tân Hiệp 2, số 226, …").
  3. Phương pháp, kết quả theo KQTN (permanganat "< 1" theo LOD 1 của phiếu; khung ghi "< 1,5 ⁽²⁾").

Chạy:  python build_137037.py [--pdf]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, nsh_base

IN = os.path.join(WORKSPACE, "Input", "137037")

CONFIG = dict(
    code6="137037", loai="nuoc sinh hoat", kind="nsh",
    client_vn="Công ty Cổ phần Đầu tư Nước Tân Hiệp",
    address_vn="Số 226, đường Dương Công Khi, xã Hóc Môn, Thành phố Hồ Chí Minh",
    reg_date="27/10/2026", dates=["09/11/2026"], time_bckq="Ngày 09/11/2026 từ 09:15 đến 10:00",
    site_vn="Phòng thí nghiệm của Nhà máy nước Tân Hiệp 2",
    scope_vn="Lấy mẫu và giám định chất lượng nước sinh hoạt",
    procedure="QUATEST 3 3030:2025, ………",       # BBLM N2.TC.N.02 — để trống như 134486
    equip=["mettler_s2", "hach_dr900"], executor="Huy",
    witness_vn="chị Mai Thị Hậu – trưởng ban hóa nghiệm",
    req_date="18/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2",
                   "Clo dư tự do": "Hach DR900", "Mùi": "-"},
    samples=[
        dict(nsh_base(), ky_hieu="137037-NSH", kqtn_key="NSH", time="09:35",
             vi_tri="Vòi nước dịch vụ - Phòng thí nghiệm",
             field={"pH": "7,71", "Nhiệt độ": "29,9", "Clo dư tự do": "0,93", "Mùi": "Không có mùi lạ"},
             photo=os.path.join(IN, "photo", "Vị trí lấy mẫu.jpg")),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
