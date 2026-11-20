# -*- coding: utf-8 -*-
r"""
Q326D139847 — Văn phòng đại diện của British-American Tobacco Marketing (Singapore) Private Limited tại Thành phố
Hồ Chí Minh — 1 mẫu nước uống tại máy RO (QCVN 06-1:2010/BYT), 1 chứng thư (mã trần) + BCKQ.
CHƯA CÓ KQTN -> bản khung, kết quả "………" tô vàng.

Nguồn (chép từ Z `Q326D139847 - Văn phòng đại diện của British-American Tobacco …` về Input\139847, chỉ đọc):
  - `BBLM.pdf` (2 trang, scan): BBLM ngày 05/11/2026, PYC Q326D139847 ngày 09/11/2026 (sau ngày BBLM — lệch); địa chỉ
    = địa điểm "Phòng 18.02, Tầng 18, Tòa nhà The Hallmark, số 15, Đường Trần Bạch Đằng, Phường An Khánh, Thành phố Hồ
    Chí Minh, Việt Nam"; "Lấy mẫu và phân tích chất lượng nước uống"; 9h30 – 10h00; thiết bị gạch bỏ; QUATEST3 3030.2025;
    N2.TC.N.02; N2.TC.N.07; chỉ định: anh Bùi Nguyễn Nhựt Anh – CHT; 01 "Nước uống tại máy RO", ký hiệu / niêm phong để
    trống; chỉ tiêu như 139854 (Sb … bào tử vi khuẩn kị khí khử sulfit, Clor, α, β); GĐV Nguyễn Minh Quân.
  - `139847 - BBLM.docx` (bản in sẵn), ảnh `Photo\` (4 ảnh).

Quyết định của người dùng (18/11/2026):
  - Ký hiệu mẫu: "không cần để kí hiệu, còn nếu form có bắt buộc thì cứ ghi NU1 rồi tô vàng" -> bảng 9.1 có cột ký hiệu
    -> "NU1" tô vàng.
  - Ngày: "đề ngày 05/11 rồi tô vàng" -> ngày giám định 05/11/2026 tô vàng (ngày PYC 09/11/2026 vẫn tô vàng).
  - Căn cứ QCVN 06-1:2010/BYT như 139854: đúng.

Suy luận của AI (đã được người dùng xem 18/11/2026):
  1. Ký hiệu mẫu "NU1" (BBLM để trống).
  2. Ngày giám định 05/11/2026 theo BBLM; ngày PYC 09/11/2026 theo BBLM (lệch thứ tự — hỏi).
  3. Căn cứ, mức, danh sách chỉ tiêu như 139854 (QCVN 06-1:2010/BYT + α, β).

Chạy:  python build_139847.py [--pdf]
"""
import argparse
import datetime
import glob
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER as P
import build_139854 as H

IN = os.path.join(WORKSPACE, "Input", "139847")
PHOTO = sorted(glob.glob(os.path.join(IN, "Photo", "*.jpg")))
ADDR = ("Phòng 18.02, tầng 18, tòa nhà The Hallmark, số 15, đường Trần Bạch Đằng, phường An Khánh, Thành phố Hồ Chí "
        "Minh")
CLIENT = ("Văn phòng đại diện của British-American Tobacco Marketing (Singapore) Private Limited tại Thành phố Hồ Chí "
          "Minh")
sample = H.mau(1, "Nước uống tại máy RO")
sample.update(cert_no=None, photo=PHOTO[0] if PHOTO else None)

CONFIG = dict(
    code6="139847", loai="nuoc uong", kind="nu", word="chỉ tiêu",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2)
    client_vn=CLIENT, address_vn=ADDR,
    reg_date="09/11/2026", dates=["05/11/2026"], time_bckq="Ngày 05/11/2026 từ 09:30 đến 10:00",
    site_vn=f"{CLIENT} - {ADDR}",
    scope_vn="Lấy mẫu và giám định chất lượng nước uống",
    procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.07",
    equip=[], executor="Quân",
    witness_vn="anh Bùi Nguyễn Nhựt Anh – CHT",
    req_date=P, highlight=["NU1", "09/11/2026", "05/11/2026"],
    samples=[sample],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
