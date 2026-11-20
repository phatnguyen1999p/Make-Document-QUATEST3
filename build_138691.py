# -*- coding: utf-8 -*-
r"""
Q326D138691 — Công ty Cổ phần Nhà máy Thủy sản Việt Úc (Cà Mau, MST 1900698808) — nước thải tại mương quan trắc,
1 chứng thư (mã trần), tiếng Việt. BẢN KHUNG: CHƯA CÓ KQTN. Khuôn chung + nguồn chung: xem docstring `build_138689.py`.

Nguồn riêng:
  - BBLM 18/11/2026, 09:20 – 10:00; "theo QCVN 11:2015/BTNMT, (Cột B, Kq = 0,9; Kf = 1)"; địa điểm viết tay "Khu xử lý
    nước thải của công ty cổ phần Nhà máy Thủy sản Việt Úc", Ấp Giồng Nhãn, phường Hiệp Thành, tỉnh Cà Mau; vị trí "Nước
    thải tại mương quan trắc", 138691 - CM, 09h30: pH 7,85 / 30,4 °C, clo dư 0,07 mg/L; ghi chú "Nước chảy chậm, nước
    trong không màu; trời nắng".
  - YCTN MTTN26020867 (1 mẫu CM, 8 thông số — có "dầu mỡ động thực vật"); PYC KH ghi "Tổng dầu mỡ khoáng".

Suy luận của AI (cần người dùng xác nhận):
  1. Quy chuẩn ghi đủ `QCVN 11-MT:2015/BTNMT` (BBLM viết "QCVN 11:2015/BTNMT"); giá trị C cột B: BOD₅ 50, COD 150, TSS 100,
     amoni 20, tổng N 60, tổng P 20, dầu mỡ động thực vật 20, clo dư 2, coliform 5 000, pH 5,5 – 9.
  2. Dầu mỡ ghi "dầu mỡ khoáng" theo PYC KH (người dùng 19/11/2026: "ghi theo yêu cầu khách trước … tô vàng"); YCTN ghi
     động thực vật. QCVN 11-MT:2015 không có mức dầu mỡ khoáng -> mức "-" (không đánh giá).
  3. Clo dư 0,07 -> "< 0,1" (Hach DR900 dưới 0,1 mg/L).

Chạy:  python build_138691.py [--pdf]
"""
import argparse
import os

from nuoc_common import run, WORKSPACE
from build_138689 import make

# QCVN 11-MT:2015 cột B không có dầu mỡ khoáng (chỉ "tổng dầu, mỡ động thực vật" 20 mg/L) -> dầu mỡ khoáng mức "-"
Q11B = {"pH": "5,5 – 9", "clo": 2, "col": 5000, "nh4": 20, "dmk": "-", "tn": 60, "tss": 100, "cod": 150, "bod": 50,
        "tp": 20}
CONFIG = make(
    "138691", "Công ty Cổ phần Nhà máy Thủy sản Việt Úc", "Ấp Giồng Nhãn, phường Hiệp Thành, tỉnh Cà Mau",
    "Khu xử lý nước thải của Công ty Cổ phần Nhà máy Thủy sản Việt Úc, Ấp Giồng Nhãn, phường Hiệp Thành, tỉnh Cà Mau",
    "Nước thải tại mương quan trắc", "138691 - CM", "CM", "18/11/2026", "09:20", "10:00", "09:30", "7,85", "< 0,1",
    "QCVN 11-MT:2015/BTNMT", "Quy chuẩn kỹ thuật quốc gia về nước thải chế biến thủy sản",
    Q11B, 0.9, 1.0, ["col", "nh4", "dmk", "tn", "tss", "cod", "bod", "tp"],
    os.path.join(WORKSPACE, "Input", "138691", "Q326D138691 - Photo", "0. Photo", "Vị trí lấy mẫu.jpg"),
    notes=["- Nhiệt độ nước thải đo tại hiện trường: 30,4 °C.",
           "- Clo dư đọc trên máy 0,07 mg/L -> ghi \"< 0,1\".",
           "- Hiện trạng (BBLM): nước chảy chậm, nước trong không màu; trời nắng.",
           "- Dầu mỡ: PYC khách hàng ghi \"tổng dầu mỡ khoáng\", phiếu YCTN ghi \"dầu mỡ động thực vật\" — ghi theo PYC khách hàng."],
    highlight=["QCVN 11-MT:2015/BTNMT"],               # tên quy chuẩn AI ghi đủ (BBLM "QCVN 11:2015") — tô vàng
    hl_names=["Dầu mỡ khoáng"])  # theo PYC KH (YCTN: động thực vật) — tô vàng

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
