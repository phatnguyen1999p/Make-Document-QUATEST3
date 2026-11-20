# -*- coding: utf-8 -*-
r"""
Q326D135225 — Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO — nước thải KCN Mỹ Xuân A (6 mẫu, 6 chứng thư).

Nguồn (chép từ Z `0.VIMCERTS078\Q326D135225 - IDICO Mỹ Xuân A - Nước thải` về Input\135225, chỉ đọc):
  - `0. BBLM.pdf` (5 BBLM 1 trang, BM3030.02, cùng PYC Q326D135225 ngày 23/10/2026, ngày 26/10/2026, quy trình N2.TC.N.06,
    người chỉ định: anh Lê Tấn Anh Vũ – chuyên viên, người thực hiện: Hồ Nguyễn Huy):
      tr.1 N1, N2 — Nhà máy XLNT tập trung KCN Mỹ Xuân A, đường 2B; 08:15 – 09:00; Hach DR300 + Mettler Seven2Go S2;
           N1 pH 7,05 – 30,3 °C – clo dư 0,6 mg/L; N2 pH 6,25 – 30,9 °C.
      tr.2 N3 — Mương quan trắc nhà máy bia Heineken; địa điểm như tr.1; 10:10 – 10:30; pH 8,13 – 34,6 °C.
      tr.3 N4 — Hố ga đấu nối vào hệ thống thu gom KCN, Công ty TNHH CJ RC – đường số 5; 09:05 – 09:15; pH 7,33 – 31,2 °C.
      tr.4 N5 — Hố ga đấu nối …, Công ty TNHH MTV Thiết bị vệ sinh Viglacera – đường số 5; 09:15 – 09:30; pH 7,22 – 29,7 °C.
      tr.5 N6 — Hố ga đấu nối …, Công ty TNHH Thủy tinh Malaya Việt Nam – đường số 7C; 09:40 – 10:00; pH 8,67 – 30,0 °C.
  - `4. edoc\135225 - Danh sách vị trí + chỉ tiêu.pdf`: kế hoạch quan trắc tháng 10/2026 của KH (căn cứ GPMT số
    156/GPMT-BTNMT ngày 02/5/2024): N1 27, N2 25, N3 14, N4 – N6 13 chỉ tiêu; nhiệt độ chỉ ở N1, clo dư chỉ ở N1.
  - `PYC-Q326D135225.N2.1.pdf` (PYC thử nghiệm MTTN26020079): ngày dự kiến trả kết quả 12/11/2026.
  - KQTN `Q326D135225 - TBKQ\124975 edoc\KQ-Q326D135225-MTTN26020079.01 – .06` (N1 – N6), ngày 13/11/2026.
  - Ảnh `0. vi tri\N1 … N6 - Vị trí lấy mẫu.jpg`; khung làm tay của người dùng trên Z — chỉ đối chiếu.

Quyết định của người dùng (16/11/2026):
  - Bảng mức GPMT 156 cho N1, N3: người dùng gửi sau (chưa đánh giá).
  - Không ghi trễ hạn KQTN (để trống, tô vàng).

Suy luận của AI (cần người dùng xác nhận):
  1. Không có mức đánh giá (khung làm tay: Căn cứ "-"; BCKQ làm tay có bảng Đánh giá N1, N3 theo phụ lục GPMT 156/GPMT-BTNMT
     nhưng chưa có giá trị mức) -> TBKQ không đánh giá. Nếu cần đánh giá N1, N3: cần bảng mức của phụ lục GPMT.
  2. KQTN ghi lấy mẫu / nhận mẫu 26/10/2026 — khớp BBLM.
  3. Địa điểm giám định từng chứng thư theo BBLM của mẫu đó (N4 – N6 tại doanh nghiệp đấu nối).
  4. BBLM ghi "theo QCVN 40:2025/BTNMT" (không phải 40:2011) — không dùng vì không có mức.

Chạy:  python build_135225.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE

IN = os.path.join(WORKSPACE, "Input", "135225")
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
TAIL = "Khu công nghiệp Mỹ Xuân A, phường Phú Mỹ, Thành phố Hồ Chí Minh"
NM = f"Nhà máy xử lý nước thải tập trung Khu công nghiệp Mỹ Xuân A, đường 2B, {TAIL}"

CONFIG = dict(
    code6="135225", loai="nuoc thai", kind="nt",
    client_vn="Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    address_vn="Số 48, khu phố 1, phường Long Hưng, Thành phố Đồng Nai",
    reg_date="23/10/2026", dates=["26/10/2026"], time_bckq="Ngày 26/10/2026 từ 08:15 đến 10:30",
    site_vn=TAIL,
    site_bckq=(f"Nhà máy xử lý nước thải tập trung Khu công nghiệp Mỹ Xuân A, đường 2B; Công ty TNHH CJ RC, Công ty TNHH MTV "
               f"Thiết bị vệ sinh Viglacera – đường số 5; Công ty TNHH Thủy tinh Malaya Việt Nam – đường số 7C, {TAIL}"),
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    equip=["mettler_s2", "hach_dr300"], executor="Huy",
    witness_vn="anh Lê Tấn Anh Vũ – chuyên viên, đại diện Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    req_date="12/11/2026",
    tre_han="………",                     # người dùng chốt 16/11/2026: không ghi trễ hạn (để trống, tô vàng)
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR300"},
    samples=[
        dict(ky_hieu="135225-N1", kqtn_key="N1", time="08:15 – 09:00", site_vn=NM, equip=["mettler_s2", "hach_dr300"],
             vi_tri="Mương quan trắc Nhà máy xử lý nước thải tập trung Khu công nghiệp Mỹ Xuân A",
             field={"pH": "7,05", "Nhiệt độ": "30,3", "Clo dư": "0,6"}, field_tbkq=["Nhiệt độ", "pH", "Clo dư"],
             photo=PH("N1")),
        dict(ky_hieu="135225-N2", kqtn_key="N2", time="08:15 – 09:00", site_vn=NM, equip=["mettler_s2"],
             vi_tri="Đầu vào Nhà máy xử lý nước thải tập trung Khu công nghiệp Mỹ Xuân A",
             field={"pH": "6,25", "Nhiệt độ": "30,9"}, field_tbkq=["pH"], photo=PH("N2")),
        dict(ky_hieu="135225-N3", kqtn_key="N3", time="10:10 – 10:30", site_vn=NM, equip=["mettler_s2"],
             vi_tri="Mương quan trắc Nhà máy bia Heineken",
             field={"pH": "8,13", "Nhiệt độ": "34,6"}, field_tbkq=["pH"], photo=PH("N3")),
        dict(ky_hieu="135225-N4", kqtn_key="N4", time="09:05 – 09:15", equip=["mettler_s2"],
             site_vn=f"Công ty TNHH CJ RC – đường số 5, {TAIL}",
             vi_tri="Hố ga đấu nối vào hệ thống thu gom Khu công nghiệp – Công ty TNHH CJ RC",
             field={"pH": "7,33", "Nhiệt độ": "31,2"}, field_tbkq=["pH"], photo=PH("N4")),
        dict(ky_hieu="135225-N5", kqtn_key="N5", time="09:15 – 09:30", equip=["mettler_s2"],
             site_vn=f"Công ty TNHH MTV Thiết bị vệ sinh Viglacera – đường số 5, {TAIL}",
             vi_tri="Hố ga đấu nối vào hệ thống thu gom Khu công nghiệp – Công ty TNHH MTV Thiết bị vệ sinh Viglacera",
             field={"pH": "7,22", "Nhiệt độ": "29,7"}, field_tbkq=["pH"], photo=PH("N5")),
        dict(ky_hieu="135225-N6", kqtn_key="N6", time="09:40 – 10:00", equip=["mettler_s2"],
             site_vn=f"Công ty TNHH Thủy tinh Malaya Việt Nam – đường số 7C, {TAIL}",
             vi_tri="Hố ga đấu nối vào hệ thống thu gom Khu công nghiệp – Công ty TNHH Thủy tinh Malaya Việt Nam",
             field={"pH": "8,67", "Nhiệt độ": "30,0"}, field_tbkq=["pH"], photo=PH("N6")),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
