# -*- coding: utf-8 -*-
r"""
Q326D139923 — Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO — KCN Mỹ Xuân A — nước thải tháng 11 (6 mẫu, 6 chứng thư).
CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D139923 - IDICO Mỹ Xuân A - Nước thải - tháng 11` về Input\139923, chỉ đọc):
  - `0. BBLM - 139923.pdf` (2 trang): ngày 12/11/2026, PYC Q326D139923 ngày 03/11/2026, 09:15 – 11:30; "Lấy mẫu, giám định
    chất lượng nước thải theo QCVN 40:2025/BTNMT"; địa điểm (viết tay) "Khu công nghiệp Mỹ Xuân A, phường Phú Mỹ, TP. Hồ Chí
    Minh"; Hach DR900 + Mettler Seven2Go S2; N2.TC.N.06; chỉ định: anh Lê Tấn Anh Vũ – chuyên viên; N1 09:30 pH 7,10 /
    29,9 °C, clo dư 0,13, clo dư tự do 0,08 mg/L; N2 10:05 pH 6,68 / 31,0; N3 09:50 pH 8,21 / 33,6; N4 10:30 pH 4,59 / 33,5;
    N5 10:50 pH 7,88 / 30,5; N6 (không ghi giờ) pH 7,44 / 29,0; ghi chú: clo dư tự do của N1 theo yêu cầu khách hàng.
    Người thực hiện: Hồ Nguyễn Huy.
  - `0. Phiếu YCTN - Q326D139923.N2.1.pdf` (MTTN26021652): dự kiến trả 30/11/2026; mẫu QC.
  - Khung làm tay của người dùng `3. TBKQ - 1…6 - Q326D139923.docx`, `2. BCKQ - Q326D139923.docx` (tên vị trí, thông số,
    phương pháp — lấy từ đây; không có mức).
  - Ảnh `0. vi tri\N1…N6 - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026): quy trình NSH để trống, clo dư < 0,1 -> "< 0,1" (tô vàng), giờ trống để "………".

Suy luận của AI (cần người dùng xác nhận):
  1. BBLM "theo QCVN 40:2025/BTNMT" nhưng không có cột / hệ số -> không đánh giá (như tháng 10, 135225).
  2. Clo dư tự do N1 đọc 0,08 mg/L (< 0,1) -> ghi "< 0,1" (AGENTS mục 8C, điểm 5).
  3. Giờ lấy mẫu N6 BBLM không ghi -> "………" tô vàng.

Chạy:  python build_139923.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER, clo

IN = os.path.join(WORKSPACE, "Input", "139923")
DR = lambda n: os.path.join(IN, f"3. TBKQ - {n} - Q326D139923.docx")      # noqa: E731
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
KCN = "Khu công nghiệp Mỹ Xuân A"
HG = "Hố ga đấu nối vào hệ thống thu gom KCN – Công ty"


def s(n, time, vi_tri, ph, t, equip=("mettler_s2",), **extra):
    f = {"pH": ph, "Nhiệt độ": t}
    f.update(extra)
    return dict(ky_hieu=f"139923-N{n}", kqtn_key=f"N{n}", time=time, vi_tri=vi_tri, draft=DR(n), field=f,
                equip=list(equip), photo=PH(f"N{n}"))


CONFIG = dict(
    code6="139923", loai="nuoc thai", kind="nt",
    client_vn="Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    address_vn="Số 48, khu phố 1, phường Long Hưng, Thành phố Đồng Nai",
    reg_date="03/11/2026", dates=["12/11/2026"], time_bckq="Ngày 12/11/2026 từ 09:15 đến 11:30",
    site_vn=f"{KCN}, phường Phú Mỹ, Thành phố Hồ Chí Minh",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    equip=["mettler_s2", "hach_dr900"], executor="Huy",
    witness_vn="anh Lê Tấn Anh Vũ – chuyên viên",
    req_date="30/11/2026",
    highlight=["< 0,1"],
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR900",
                   "Clo dư tự do": "Hach DR900"},
    bckq_notes=["- Clo dư tự do của mẫu N1 đo theo yêu cầu của khách hàng (theo BBLM)."],
    samples=[
        s(1, "09:30", f"Mương quan trắc Nhà máy XLNT tập trung {KCN}", "7,10", "29,9", ("mettler_s2", "hach_dr900"),
          **{"Clo dư": clo("0,13"), "Clo dư tự do": clo("0,08")}),
        s(2, "10:05", f"Đầu vào Nhà máy XLNT tập trung {KCN}", "6,68", "31,0"),
        s(3, "09:50", "Mương quan trắc Nhà máy bia Heineken", "8,21", "33,6"),
        s(4, "10:30", f"{HG} TNHH Thức ăn gia súc Ngô Nguyễn", "4,59", "33,5"),
        s(5, "10:50", f"{HG} TNHH Fritta Việt Nam", "7,88", "30,5"),
        s(6, PLACEHOLDER, f"{HG} TNHH Lixil Việt Nam tại Bà Rịa Vũng Tàu", "7,44", "29,0"),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
