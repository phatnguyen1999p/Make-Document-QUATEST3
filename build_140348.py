# -*- coding: utf-8 -*-
r"""
Q326D140348 — Công ty Cổ phần Phát triển Doanh nghiệp nhỏ và vừa Nhật Bản — nước thải 2 mẫu (đầu vào / đầu ra trạm XLNT)
theo Quy định đấu nối KCN Nhơn Trạch III, 2 chứng thư. CHƯA CÓ KQTN -> bản khung, "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D140348 - PT DN Vừa và nhỏ Nhật Bản - NT` về Input\140348, chỉ đọc):
  - `0. BBLM - 140348.pdf` (2 trang): ngày 12/11/2026, PYC Q326D140348 ngày 03/11/2026, 13:30 – 14:05; "Lấy mẫu, giám định
    chất lượng nước thải theo Quy định đấu nối KCN Nhơn Trạch III"; địa điểm "KNX Nhật Bản (JSC), KCN Nhơn Trạch III giai
    đoạn 2, phường Nhơn Trạch, thành phố Đồng Nai"; Hanna HI 98192 (TDS) + Mettler Seven2Go S2; N2.TC.N.06; chỉ định: anh
    Đào Trung Trí – nhân viên; NT1 đầu vào 13:35 pH 8,15 / 32,0 °C; NT2 đầu ra 13:50 pH 5,06 / 31,2 °C; ghi chú: NT1 nước
    có màu nâu nhạt, nhiều cặn bẩn, có mùi hôi; NT2 nước đục, có cặn lơ lửng. Người thực hiện: Hồ Nguyễn Huy.
  - `0. Phiếu  YCTN -Q326D140348.N2.1.pdf` (MTTN26021651): 24 chỉ tiêu / mẫu, dự kiến trả 30/11/2026; mẫu QC.
  - Khung làm tay của người dùng `3. TBKQ - 1/2 - Q326D140348.docx`, `2. BCKQ - Q326D140348.docx` (thông số, phương pháp,
    mức theo Quy định giới hạn tiếp nhận nước thải KCN Nhơn Trạch III – giai đoạn 2, Phụ lục 1; `0. Mức so sánh.jpg`).
  - Ảnh `Q326D140348 - Photo\0. photo\NT1/NT2 - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026): giữ máy TDS Hanna HI 98192 trong danh sách thiết bị.

Suy luận của AI (cần người dùng xác nhận):
  1. (không còn)

Chạy:  python build_140348.py [--pdf] [--only 1,2]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, params_from_draft, limits_from_draft

IN = os.path.join(WORKSPACE, "Input", "140348")
DR = lambda n: os.path.join(IN, f"3. TBKQ - {n} - Q326D140348.docx")                         # noqa: E731
PH = lambda k: os.path.join(IN, "Q326D140348 - Photo", "0. photo", f"{k} - Vị trí lấy mẫu.jpg")  # noqa: E731
QD = ("Quy định giới hạn tiếp nhận nước thải của hệ thống xử lý nước thải tập trung Khu công nghiệp Nhơn Trạch III – "
      "giai đoạn 2 (Phụ lục 1)")
SITE = "KNX Nhật Bản (JSC), Khu công nghiệp Nhơn Trạch III giai đoạn 2, phường Nhơn Trạch, Thành phố Đồng Nai"


def nt(n, time, vi_tri, ph, t):
    return dict(ky_hieu=f"140348-NT{n}", kqtn_key=f"NT{n}", time=time, vi_tri=vi_tri,
                params=params_from_draft(DR(n)), limits=limits_from_draft(DR(n)), standard_vn=QD,
                limit_head="Giá trị tối đa ^(1)", limit_note=f"Giá trị tối đa theo {QD}.",
                intro=f"So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên với mức quy định tương ứng tại "
                      f"{QD}:",
                field={"pH": ph, "Nhiệt độ": t}, photo=PH(f"NT{n}"))


CONFIG = dict(
    code6="140348", loai="nuoc thai", kind="nt",
    client_vn="Công ty Cổ phần Phát triển Doanh nghiệp nhỏ và vừa Nhật Bản",
    address_vn=SITE,
    reg_date="03/11/2026", dates=["12/11/2026"], time_bckq="Ngày 12/11/2026 từ 13:30 đến 14:05",
    site_vn=SITE,
    scope_vn="Lấy mẫu và giám định chất lượng nước thải theo Quy định đấu nối Khu công nghiệp Nhơn Trạch III",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    equip=["mettler_s2", "hanna_hi98192"], executor="Huy",
    witness_vn="anh Đào Trung Trí – nhân viên",
    req_date="30/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2"},
    bckq_notes=["- Hiện trạng (theo BBLM): NT1 nước có màu nâu nhạt, nhiều cặn bẩn, có mùi hôi; NT2 nước đục, có cặn "
                "lơ lửng."],
    samples=[
        nt(1, "13:35", "Đầu vào trạm xử lý nước thải", "8,15", "32,0"),
        nt(2, "13:50", "Đầu ra trạm xử lý nước thải", "5,06", "31,2"),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
