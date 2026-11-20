# -*- coding: utf-8 -*-
r"""
Q326D138693 — Công ty Cổ phần Việt Úc - Bạc Liêu (MST 1900345344) — nước thải sau xử lý, 1 chứng thư (mã trần), tiếng
Việt. BẢN KHUNG: CHƯA CÓ KQTN. Khuôn chung + nguồn chung: xem docstring `build_138689.py`.

Nguồn riêng:
  - BBLM 18/11/2026, 08:30 – 09:10; "theo QCVN 40:2011/BTNMT, (Cột B, Kq = 0,9; Kf = 1)"; tên đơn vị in "CÔNG TY VIỆT
    ÚC – BẠC LIÊU", viết tay thêm "CỔ PHẦN"; địa điểm viết tay "Khu xử lý nước thải của Công ty cổ phần Việt Úc – Bạc
    Liêu", Khóm Nhà Mát, phường Hiệp Thành, tỉnh Cà Mau; vị trí "Nước thải sau xử lý", 138693 - BL, 8h45: pH 8,57 /
    31,2 °C, clo dư 0,20 mg/L; ghi chú "Nước chảy mạnh, nước hơi đục, ít cặn; trời nắng, gió nhẹ".
  - YCTN `0. Phiếu YCTN-Q326D138693.N2.1.pdf` (MTTN26020937: .01 BL 7 thông số; .02 "138691 - 138693 QC" amoni).

Chạy:  python build_138693.py [--pdf]
"""
import argparse
import os

from nuoc_common import run, WORKSPACE
from build_138689 import make, Q40, Q40B

CONFIG = make(
    "138693", "Công ty Cổ phần Việt Úc - Bạc Liêu", "Khóm Nhà Mát, phường Hiệp Thành, tỉnh Cà Mau",
    "Khu xử lý nước thải của Công ty Cổ phần Việt Úc - Bạc Liêu, Khóm Nhà Mát, phường Hiệp Thành, tỉnh Cà Mau",
    "Nước thải sau xử lý", "138693 - BL", "BL", "18/11/2026", "08:30", "09:10", "08:45", "8,57", "0,20",
    *Q40, {k: v for k, v in Q40B.items() if k != "dmk"}, 0.9, 1.0, ["col", "nh4", "tn", "tss", "cod", "bod", "tp"],
    os.path.join(WORKSPACE, "Input", "138693", "Q326D138693 - Photo", "0. Photo", "Vị trí lấy mẫu.jpg"),
    notes=["- Nhiệt độ nước thải đo tại hiện trường: 31,2 °C.",
           "- Hiện trạng (BBLM): nước chảy mạnh, nước hơi đục, ít cặn; trời nắng, gió nhẹ."])

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
