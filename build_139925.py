# -*- coding: utf-8 -*-
r"""
Q326D139925 — Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO — KCN Nhơn Trạch 5 — nước thải tháng 11 (9 mẫu, 9 chứng thư:
N1 – N7, N9, N10). CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D139925 - IDICO Nhơn Trạch 5 - Nước Thải tháng 11` về Input\139925, chỉ đọc):
  - `0.  BBLM - NHÓM HUY.pdf` (2 trang): ngày 05/11/2026, PYC Q326D139925 ngày 03/11/2026, 08:50 – 11:00; "nước thải theo
    QCVN 40:2011/BTNMT"; địa điểm "Khu công nghiệp Nhơn Trạch 5, phường Nhơn Trạch, TP. Đồng Nai"; Hach DR900 + Mettler;
    N2.TC.N.06; chỉ định: anh Bùi Thanh Hải – quản đốc nhà máy; N1 10:20 pH 7,11 / 32,8 °C; N2 10:40 pH 7,05 / 33,3;
    N4 10:15 pH 6,45 / 30,9; N5 09:20, N6 09:40, N7 09:50 (chỉ lấy mẫu).
  - `0. BBLM - NHÓM VIỆT.pdf` (2 trang): 08:50 – 11:00, Hach DR300 + Mettler; N3 09:20 pH 7,67 / 34,4; N9 pH 8,70 / 37,0 (không
    ghi giờ); N10 10:00 (chỉ lấy mẫu); N8 (Triển Vũ), N11 (Oto Vina): không có nước thải chảy ra, không lấy mẫu.
  - `PYC-Q326D139925.N2.1.pdf` (MTTN21044): chỉ tiêu từng mẫu (lấy từ đây), dự kiến trả 19/11/2026; mẫu QC.
  - Ảnh `0. vi tri\Nn - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026): số chứng thư theo số mẫu (N9 -> /9), tô vàng chỗ nhảy cóc; bỏ máy clo (không có số đo).

Suy luận của AI (cần người dùng xác nhận):
  1. Chỉ tiêu theo PYC thử nghiệm; tên / phương pháp sẽ thay theo KQTN khi có.
  2. BBLM ghi máy clo dư (DR900, DR300) nhưng không có số clo dư -> chỉ ghi Mettler Seven2Go S2.
  3. Không có cột / hệ số QCVN 40:2011 -> không đánh giá (như tháng 10, 135228).
  4. Giờ lấy mẫu N9 không ghi -> "………" tô vàng.

Chạy:  python build_139925.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER, pyc_index, pyc_sample

IN = os.path.join(WORKSPACE, "Input", "139925")
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
PYC, _ = pyc_index(os.path.join(IN, "PYC-Q326D139925.N2.1.pdf"))
NM = "Nhà máy xử lý nước thải Khu công nghiệp Nhơn Trạch 5"
TT = ("Nhiệt độ", "pH")


def s(key, time, vi_tri, ph=None, t=None, tbkq=("pH",)):
    f = {"pH": ph, "Nhiệt độ": t} if ph else {}
    return pyc_sample("139925", key, PYC, time, vi_tri, f, tbkq, photo=PH(key))


CONFIG = dict(
    code6="139925", loai="nuoc thai", kind="nt",
    cert_by_key=True,                    # số chứng thư theo số mẫu (Nn -> /n), tô vàng khi nhảy cóc
    client_vn="Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    address_vn="Số 48, khu phố 1, phường Long Hưng, Thành phố Đồng Nai",
    reg_date="03/11/2026", dates=["05/11/2026"], time_bckq="Ngày 05/11/2026 từ 08:50 đến 11:00",
    site_vn="Khu công nghiệp Nhơn Trạch 5, phường Nhơn Trạch, Thành phố Đồng Nai",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    equip=["mettler_s2"], executor="Huy + Việt",
    witness_vn="anh Bùi Thanh Hải – quản đốc nhà máy",
    req_date="19/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2"},
    bckq_notes=["- Mẫu N8 (Nước thải Công ty TNHH Vật liệu KHKT Triển Vũ Việt Nam) và N11 (Nước thải Công ty TNHH Oto Vina): "
                "tại thời điểm lấy mẫu không có nước thải chảy ra tại hố ga, không lấy mẫu (theo BBLM)."],
    samples=[
        s("N1", "10:20", f"Nước thải đầu vào {NM}", "7,11", "32,8", TT),
        s("N2", "10:40", f"Nước thải đầu ra {NM}", "7,05", "33,3", TT),
        s("N3", "09:20", "Nước thải Công ty TNHH Da thuộc TaiYu", "7,67", "34,4"),
        s("N4", "10:15", "Nước thải Công ty Cổ phần Dệt Renze", "6,45", "30,9"),
        s("N5", "09:20", "Nước thải Công ty TNHH Hyosung Đồng Nai (D2)"),
        s("N6", "09:40", "Nước thải Công ty TNHH Hyosung Đồng Nai (N4)"),
        s("N7", "09:50", "Nước thải Công ty TNHH Hyosung Vina Industrial Machinery"),
        s("N9", PLACEHOLDER, "Nước thải Công ty TNHH Hyundai Welding Vina", "8,70", "37,0"),
        s("N10", "10:00", "Nước thải Công ty Cổ phần KCN Đồng Nai"),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
