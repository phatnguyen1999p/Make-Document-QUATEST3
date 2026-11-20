# -*- coding: utf-8 -*-
r"""
Q326D139854 — Công ty Cổ phần Hadico — nước uống tại Công ty TNHH Intel Products Việt Nam: 2 mẫu NU1, NU2
(QCVN 06-1:2010/BYT), 2 chứng thư /1, /2 + BCKQ. CHƯA CÓ KQTN -> bản khung, kết quả "………" tô vàng.

Nguồn (chép từ Z `Q326D139854 - Hadico - Intel - NU` về Input\139854, chỉ đọc):
  - `BBLM.pdf` (2 trang, scan): BBLM ngày 09/11/2026, PYC Q326D139854 ngày 03/11/2026; "CÔNG TY CỔ PHẦN HADICO", "14 Đường
    46 phường Tân Tạo, TP Hồ Chí Minh"; "Lấy mẫu và phân tích chất lượng nước uống"; địa điểm "CÔNG TY TNHH INTEL PRODUCTS
    VIỆT NAM - Địa điểm: Lô I2, đường D1, Khu Công nghệ cao TP.HCM, Phường Tăng Nhơn Phú, TP.HCM"; 10h00 – 10h30; thiết
    bị: gạch bỏ; QUATEST3 3030.2025; N2.TC.N.02; N2.TC.N.07; chỉ định: anh Nguyễn Ngọc Long – giám sát; NU1 "Vòi rửa rau
    số 3", NU2 "Vòi rửa rau số 4" (viết tay, đọc được như vậy); không niêm phong; chỉ tiêu: Sb, As, Ba, B, bromat, Cd,
    clorat, clorit, tổng Cr, Cu, CN⁻, F⁻, Pb, Mn, Hg, Mo, Ni, nitrat, nitrit, Se, tổng α, tổng β, E. coli, tổng
    Coliform, Streptococci feacal, P. aeruginosa, bào tử vi khuẩn kị khí khử sulfit, Clor; GĐV Nguyễn Minh Quân.
  - `139854 - Biên bản hiện trường.xlsx` (bản in sẵn của BBLM, ghi chép), ảnh `Photo\NU1.jpg`, `NU2.jpg`.

Quyết định của người dùng (18/11/2026): căn cứ QCVN 06-1:2010/BYT + mức α, β: đúng; tên vị trí "Vòi rửa rau số 3 / số 4":
đúng.

Suy luận của AI (đã được người dùng xem 18/11/2026):
  1. Căn cứ QCVN 06-1:2010/BYT phụ lục II, III (BBLM không ghi quy chuẩn; danh sách chỉ tiêu trùng QCVN 06-1); mức như
     131229, thêm tổng hoạt độ phóng xạ α ≤ 0,5 Bq/L, β ≤ 1 Bq/L (phương pháp chưa biết -> "………").
  2. Tên vị trí "Vòi rửa rau số 3 / số 4" đọc từ chữ viết tay.
  3. Không có PYC thử nghiệm -> ngày yêu cầu cấp KQTN "………"; chứng thư tiếng Việt.

Chạy:  python build_139854.py [--pdf] [--only 1,2]
"""
import argparse
import datetime
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER as P
from build_131229 import LIMITS as L06, QC06
from build_144954 import NU_PARAMS

IN = os.path.join(WORKSPACE, "Input", "139854")
PARAMS = [p for p in NU_PARAMS if p[2] != "Đo tại hiện trường" and "TDS" not in p[0]]
PARAMS += [("Tổng hoạt độ phóng xạ α", "Bq/L", P), ("Tổng hoạt độ phóng xạ β", "Bq/L", P)]
LIMITS = dict(L06, **{"Tổng hoạt độ phóng xạ α": "0,5", "Tổng hoạt độ phóng xạ β": "1"})
CHI_TIEU = f"21 chỉ tiêu hóa lý, tổng hoạt độ phóng xạ α, β và 05 chỉ tiêu vi sinh tại {QC06}, phụ lục II và phụ lục III"


def mau(n, vi_tri):
    return dict(kind="nu", cert_no=n, ky_hieu=f"NU{n}", kqtn_key=f"NU{n}", vi_tri=vi_tri, time="-", field={},
                field_tbkq=[], equip=[], params=PARAMS, limits=LIMITS, chi_tieu=CHI_TIEU,
                photo=os.path.join(IN, "Photo", f"NU{n}.jpg"),
                standard_vn=f"{QC06} - Quy chuẩn kỹ thuật quốc gia đối với nước khoáng thiên nhiên và nước uống đóng chai",
                limit_head="Giới hạn tối đa ^(1)",
                limit_note=f"Mức giới hạn tối đa cho phép quy định tại phụ lục II và phụ lục III, {QC06}.",
                eval_words=("Chỉ tiêu", "không đạt", "đạt"),
                intro=f"So sánh kết quả thử nghiệm các chỉ tiêu của mẫu nước uống nêu trên so với mức quy định tương "
                      f"ứng tại {QC06}, phụ lục II và phụ lục III:")


CONFIG = dict(
    code6="139854", loai="nuoc uong", kind="nu", word="chỉ tiêu",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2)
    client_vn="Công ty Cổ phần Hadico",
    address_vn="14 đường 46, phường Tân Tạo, Thành phố Hồ Chí Minh",
    reg_date="03/11/2026", dates=["09/11/2026"], time_bckq="Ngày 09/11/2026 từ 10:00 đến 10:30",
    site_vn="Công ty TNHH Intel Products Việt Nam - Lô I2, đường D1, Khu Công nghệ cao Thành phố Hồ Chí Minh, phường "
            "Tăng Nhơn Phú, Thành phố Hồ Chí Minh",
    scope_vn="Lấy mẫu và giám định chất lượng nước uống",
    procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.07",
    equip=[], executor="Quân",
    witness_vn="anh Nguyễn Ngọc Long – giám sát",
    req_date=P,
    samples=[mau(1, "Vòi rửa rau số 3"), mau(2, "Vòi rửa rau số 4")],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
