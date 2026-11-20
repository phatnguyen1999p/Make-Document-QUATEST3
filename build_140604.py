# -*- coding: utf-8 -*-
r"""
Q326D140604 — Công ty Cổ phần Đầu tư và Kinh doanh Nước sạch Sài Gòn — nước mặt (nước thô trạm bơm Hóa An), 1 mẫu,
1 chứng thư (mã trần). CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D140604 - Nước sạch Sài Gòn - NM tháng 11` về Input\140604, chỉ đọc):
  - `0. BBLM 14604.pdf` (2 trang): ngày 12/11/2026, PYC Q326D140604 ngày 04/11/2026, 11:45 – 12:10; "Lấy mẫu, giám định chất
    lượng nước mặt theo QCVN 08:2023/BTNMT"; địa điểm "Trạm bơm nước thô Hóa An, phường Biên Hòa, thành phố Đồng Nai";
    Hanna HI 9142 (DO) + Mettler Seven2Go S2; QTGĐ 3030.2025, N2.TC.N.04; chỉ định: anh Nguyễn Thái Minh – nhân viên;
    140604-NM "Nước thô trạm bơm Hóa An": pH 6,99, t° 30,4 °C, DO 4,7 mg/L; ghi chú: nước trong hầm thu, nước đục có cặn
    lơ lửng; trời nắng, nước chảy chậm. Người thực hiện: Hồ Nguyễn Huy.
  - `0. Phiếu YCTN-Q326D140604.N2.1.pdf` (MTTN26022164): 8 chỉ tiêu (E. coli, TOC, T-N, TSS, COD, BOD₅, T-P, vi sinh);
    dự kiến trả 27/11/2026; mẫu QC (T-P).
  - Chứng thư tháng trước `Q326D131874 - TBKQ\3. TBKQ - Q326D131874.docx` (mức QCVN 08:2023/BTNMT theo bảng giới hạn nguồn
    nước thô của khách hàng `Bang chi tieu gioi han ve nguon nuoc tho.pdf`).
  - Ảnh `Q326D140604 - Photo\0. Photo\Vị trí lấy mẫu.jpg`.

Quyết định của người dùng: (chưa có)

Suy luận của AI (cần người dùng xác nhận):
  1. Chỉ tiêu theo PYC thử nghiệm (8 chỉ tiêu + pH, DO đo hiện trường) — chứng thư 131874 có 49 chỉ tiêu.
  2. Mức đánh giá lấy theo chứng thư 131874 (tô vàng cột mức).
  3. Máy đo DO Hanna HI 9142 thêm vào danh mục thiết bị (`hanna_hi9142`) với tên "Thiết bị đo oxy hòa tan Hanna HI 9142,
     USA".
  4. Vị trí lấy mẫu ghi theo BBLM "Nước thô trạm bơm Hóa An" (131874 ghi "Bể chứa nước thô").

Chạy:  python build_140604.py [--pdf]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, FIELD

IN = os.path.join(WORKSPACE, "Input", "140604")
STD = "QCVN 08:2023/BTNMT - Quy chuẩn kỹ thuật quốc gia về chất lượng nước mặt"
NOTE = "Giới hạn theo bảng chỉ tiêu giới hạn về nguồn nước thô do khách hàng cung cấp (QCVN 08:2023/BTNMT)."
PARAMS = [
    ("pH", "-", FIELD), ("Oxy hòa tan (DO)", "mg/L", FIELD),
    ("Nhu cầu oxy sinh hóa sau 5 ngày (BOD₅)", "mg/L", "TCVN 6001-1:2021"),
    ("Nhu cầu oxy hóa học (COD)", "mg/L", "SMEWW 5220D:2023"),
    ("Hàm lượng carbon hữu cơ toàn phần (TOCs)", "mg/L", "TCVN 6634:2000"),
    ("Hàm lượng tổng chất rắn lơ lửng (TSS)", "mg/L", "SMEWW 2540D:2023"),
    ("Tổng phospho (P)", "mg/L", "US EPA 200.7"),
    ("Hàm lượng nitơ tổng", "mg/L", "TCVN 6624-1:2000"),
    ("Coliforms", "MPN/100mL", "SMEWW 9221B:2023"),
    ("Escherichia coli", "MPN/100mL", "SMEWW 9221B&F:2023"),
]
LIMITS = {"pH": "6,5 – 8,5", "Oxy hòa tan (DO)": "≥ 6,0", "Nhu cầu oxy sinh hóa sau 5 ngày (BOD₅)": "≤ 4",
          "Nhu cầu oxy hóa học (COD)": "≤ 10", "Hàm lượng carbon hữu cơ toàn phần (TOCs)": "≤ 4",
          "Hàm lượng tổng chất rắn lơ lửng (TSS)": "≤ 25", "Tổng phospho (P)": "≤ 0,1",
          "Hàm lượng nitơ tổng": "≤ 0,6", "Coliforms": "≤ 1 000", "Escherichia coli": "20"}

CONFIG = dict(
    code6="140604", loai="nuoc mat", kind="nt",
    client_vn="Công ty Cổ phần Đầu tư và Kinh doanh Nước sạch Sài Gòn",
    address_vn="Số 2A1 Lê Văn Chí, phường Linh Xuân, Thành phố Hồ Chí Minh",
    reg_date="04/11/2026", dates=["12/11/2026"], time_bckq="Ngày 12/11/2026 từ 11:45 đến 12:10",
    site_vn="Trạm bơm nước thô Hóa An, phường Biên Hòa, Thành phố Đồng Nai",
    scope_vn="Lấy mẫu và giám định chất lượng nước mặt",
    procedure="QUATEST 3 3030:2025, N2.TC.N.04",
    equip=["mettler_s2", "hanna_hi9142"], executor="Huy",
    witness_vn="anh Nguyễn Thái Minh – nhân viên",
    req_date="27/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "DO": "Hanna HI 9142"},
    bckq_notes=["- Hiện trạng: nước trong hầm thu, nước đục có cặn lơ lửng; trời nắng, nước chảy chậm (theo BBLM)."],
    samples=[
        dict(ky_hieu="140604-NM", kqtn_key="NM", time="11:45 – 12:10", vi_tri="Nước thô trạm bơm Hóa An",
             params=PARAMS, limits=LIMITS, hl_limits=True, standard_vn=STD,
             field={"pH": "6,99", "Nhiệt độ": "30,4", "DO": "4,7"}, field_tbkq=["pH", "DO"],
             limit_head="Giới hạn ^(1)", limit_note=NOTE,
             intro="So sánh kết quả thử nghiệm các thông số của mẫu nước mặt nêu trên với mức giới hạn tương ứng:",
             photo=os.path.join(IN, "Q326D140604 - Photo", "0. Photo", "Vị trí lấy mẫu.jpg")),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
