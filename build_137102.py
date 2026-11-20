# -*- coding: utf-8 -*-
r"""
Q326D137102 — Công ty TNHH Ô tô Isuzu Việt Nam — nước thải 1 mẫu (QCVN 40:2011/BTNMT cột B, Kq = Kf = 1),
1 chứng thư (mã trần). ĐÃ CÓ KQTN MTTN26020965 (17/11/2026).

Nguồn (chép từ Z `0.VIMCERTS078\Q326D137102 - ISUZU - Nước thải` về Input\137102, chỉ đọc):
  - `0. BBLM - 137102.pdf` (2 trang): ngày 04/11/2026, PYC Q326D137102 ngày 28/10/2026, 09:30 – 10:00; "Lấy mẫu, giám định
    chất lượng nước thải theo yêu cầu khách hàng"; địa điểm "Hố ga giám sát nước thải sau xử lý của cửa hàng ô tô Isuzu tại
    số 695 Quang Trung, phường Thông Tây Hội, Tp. HCM"; Hach DR900 + Mettler Seven2Go S2; N2.TC.N.06; chỉ định: chị Lê
    Hạnh Nguyên – nhân viên môi trường; 137102-NT: pH 7,50, t° 31,4 °C, clo dư 0,16 mg/L; ghi chú: nước trong hố ga có
    nắp đậy, nước trong không màu, có cặn lơ lửng. Người thực hiện: Hồ Nguyễn Huy.
  - `0. PYC-Q326D137102.N2.1.pdf` (MTTN26020965): 16 chỉ tiêu, dự kiến trả 18/11/2026; mẫu QC.
  - Khung làm tay của người dùng `3. TBKQ - Q326D137102.docx`, `2. BCKQ - Q326D137102.docx` (thông số, phương pháp, mức,
    địa điểm, vị trí — lấy từ đây).
  - Ảnh `Q326D138689 - Photo\Vị trí lấy mẫu.jpg` (tên thư mục ảnh ghi 138689).

Quyết định của người dùng (18/11/2026): "có KQTN rồi" -> dựng lại theo KQTN MTTN26020965 (.01 NT, .02 QC, 17/11/2026).

Suy luận của AI (cần người dùng xác nhận):
  1. Mức đánh giá QCVN 40:2011/BTNMT cột B, Kq = Kf = 1 theo khung làm tay (BBLM ghi "theo yêu cầu khách hàng").

Chạy:  python build_137102.py [--pdf]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, params_from_draft, limits_from_draft

IN = os.path.join(WORKSPACE, "Input", "137102")
DR = os.path.join(IN, "3. TBKQ - Q326D137102.docx")
STD = "QCVN 40:2011/BTNMT - Quy chuẩn kỹ thuật quốc gia về nước thải công nghiệp"
NOTE = "Giá trị tối đa cho phép theo giá trị C - cột B, hệ số Kq = 1, Kf = 1 tại QCVN 40:2011/BTNMT."

CONFIG = dict(
    code6="137102", loai="nuoc thai", kind="nt",
    client_vn="Công ty TNHH Ô tô Isuzu Việt Nam",
    address_vn="Số 695 Quang Trung, phường Thông Tây Hội, Thành phố Hồ Chí Minh",
    reg_date="28/10/2026", dates=["04/11/2026"], time_bckq="Ngày 04/11/2026 từ 09:30 đến 10:00",
    site_vn="Cửa hàng ô tô Isuzu, số 695 Quang Trung, phường Thông Tây Hội, Thành phố Hồ Chí Minh",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    equip=["mettler_s2", "hach_dr900"], executor="Huy",
    witness_vn="chị Lê Hạnh Nguyên – nhân viên môi trường",
    req_date="18/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR900"},
    bckq_notes=["- Hiện trạng: nước trong hố ga có nắp đậy, nước trong không màu, có cặn lơ lửng (theo BBLM)."],
    samples=[
        dict(ky_hieu="137102-NT", kqtn_key="NT", time="09:30 – 10:00",
             vi_tri="Hố ga giám sát nước thải sau xử lý, trước khi thải vào hệ thống thoát nước chung của khu vực",
             params=params_from_draft(DR), standard_vn=STD,
             limits=dict(limits_from_draft(DR), **{   # QCVN 40:2011 cột B, Kq = Kf = 1 — tên theo KQTN (khung khác tên)
                 "Hàm lượng cadmi (Cd)": "0,1", "Hàm lượng crom III (Cr³⁺)": "1", "Hàm lượng crom VI (Cr⁶⁺)": "0,1",
                 "Tổng phospho (P)": "6"}),
             limit_head="Giá trị tối đa ^(1)", limit_note=NOTE,
             intro="So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên với mức quy định tương ứng tại "
                   "QCVN 40:2011/BTNMT (cột B, Kq = 1, Kf = 1):",
             field={"pH": "7,50", "Nhiệt độ": "31,4", "Clo dư": "0,16"},
             photo=os.path.join(IN, "Q326D138689 - Photo", "Vị trí lấy mẫu.jpg")),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
