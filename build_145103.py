# -*- coding: utf-8 -*-
r"""
Q326D145103 — Công ty Cổ phần Kỹ nghệ Đô Thành — 2 mẫu nước khách mang đến Trung tâm (mẫu gửi, không lấy mẫu hiện
trường): X. PET (/1), X. Khuôn (/2), so sánh QCVN 01-1:2024/BYT. CHƯA CÓ KQTN -> bản khung, kết quả "………" tô vàng.

Nguồn (chép từ Z `Q326D145103 - DO THANH KY NGHE - NC` về Input\145103, chỉ đọc):
  - Ảnh `145103 - PHOTO\…18e747ec….jpg`: PYC giám định Q326D145103 ngày 12/11/2026 — "Công ty Cổ phần Kỹ nghệ Đô Thành",
    "59-65 Huỳnh Mẫn Đạt, phường An Đông, thành phố Hồ Chí Minh", đại diện Lê Thị Khánh Vân – nhân viên, MST viết tay
    (đọc được 0303510875, chưa chắc — tra MST không ra); mục 11: "Mẫu nước uống (Nước sạch sử dụng cho mục đích sinh
    hoạt)", 2 bình, "Giám định mẫu nước cấp - Thông số phân tích/thử nghiệm: 10 thông số nhóm A, Pseudomonas aeruginosa,
    Staphylococcus aureus và đánh giá kết quả phân tích/thử nghiệm mẫu nước theo QCVN 01-1:2024/BYT"; mục 13 "TT3";
    mục 15 chỉ Tiếng Việt.
  - Ảnh `…90fc5de0….jpg`: phiếu kiểm soát công việc DG Q326D145103 (Kỹ nghệ Đô Thành), phân công Nghĩa 12/11/2026.
  - `PYC-Q326D145103.N2.1.pdf` (MTTN26021960, nhận 12/11/2026, dự kiến trả 20/11/2026): .01 "Q326D145103 - NSH (X. PET)",
    .02 "Q326D145103 - NSH (X. KHUON)" — Nhóm A hóa lý (màu, mùi, pH, đục, As, clo dư tự do, permanganate, amoni) + vi
    sinh (Coliforms, E. coli, P. aeruginosa, S. aureus).
  - Ảnh mẫu `MAU XUONG PET .jpg`, `MAU XUONG KHUON .jpg` (bình nhựa trong, nhãn "X. PET" / "X. Khuôn" + "Q326D145103").
  - Khung làm tay trên Z `2. BCKQDG.docx`, `3. TBKQ - XUONG PET/KHUON .docx` (chép từ hồ sơ Q326D128096 — khuôn mẫu gửi;
    danh sách thông số, phương pháp, mức lấy từ đây).

Quyết định của người dùng (17/11/2026):
  - Tên mẫu theo khai báo: giữ như AI ghi, tô vàng. Mô tả tình trạng mẫu: ok. Nội dung yêu cầu "Giám định mẫu nước cấp".
  - Mức Coliforms: giữ "< 1" theo khung, tô vàng.

Suy luận của AI (đã được người dùng xem 17/11/2026):
  1. Tên mẫu theo khai báo "Mẫu nước sạch sử dụng cho mục đích sinh hoạt – X. PET / X. Khuôn" (PYC ghi "Mẫu nước uống
     (Nước sạch sử dụng cho mục đích sinh hoạt)", PYC thử nghiệm ghi "NSH (X. PET)").
  2. Mô tả tình trạng mẫu theo ảnh (bình nhựa ~5 lít, không thấy niêm phong); ngày nhận mẫu 12/11/2026; người gửi mẫu
     Lê Thị Khánh Vân.
  3. Mức QCVN theo khung làm tay (Coliforms < 1 CFU/100 mL như khung).
  4. Nội dung yêu cầu "Giám định mẫu nước cấp" theo PYC; quy trình chỉ `QUATEST 3 3030:2025` (mẫu gửi); thiết bị "-".

Chạy:  python build_145103.py [--pdf]
"""
import argparse
import datetime
import os
from nuoc_common import run, WORKSPACE, params_from_draft

IN = os.path.join(WORKSPACE, "Input", "145103")
DRAFT = os.path.join(IN, "3.  TBKQ - XUONG PET .docx")
_P = params_from_draft(DRAFT, with_limit=True)
PARAMS = [(n, u, m) for n, u, m, _l in _P]
LIMITS = {n: l.replace("Trong khoảng (", "Trong khoảng\n(") for n, _u, _m, l in _P}
MO_TA = ("Chất lỏng trong, không cặn, được chứa trong bình nhựa 5 lít. Trên thân bình có dán nhãn ghi tên mẫu "
         "\"{}\" và mã hồ sơ Q326D145103.")
CHI_TIEU = "10 thông số nhóm A, Pseudomonas aeruginosa và Staphylococcus aureus tại QCVN 01-1:2024/BYT"


def mau(n, nhan, kqtn, photo):
    return dict(kind="nsh", gui_mau=True, cert_no=n, ky_hieu=nhan, kqtn_key=kqtn,
                ten_mau=f"Mẫu nước sạch sử dụng cho mục đích sinh hoạt – {nhan}", mo_ta=MO_TA.format(nhan),
                chi_tieu=CHI_TIEU, params=PARAMS, limits=LIMITS, hl_limits=["Coliforms tổng số"], field={}, equip=[],
                photo=os.path.join(IN, "145103 - PHOTO", photo), time="-", vi_tri=nhan,
                standard_vn="QCVN 01-1:2024/BYT - Quy chuẩn kỹ thuật quốc gia về chất lượng nước sạch sử dụng cho mục "
                            "đích sinh hoạt",
                limit_head="Ngưỡng giới hạn ^(1)",
                limit_note="Ngưỡng giới hạn tối đa nêu tại QCVN 01-1:2024/BYT.",
                intro="So sánh kết quả thử nghiệm các thông số của mẫu nước nêu trên với mức quy định tương ứng tại "
                      "QCVN 01-1:2024/BYT:")


CONFIG = dict(
    code6="145103", loai="nuoc sinh hoat", kind="nsh", gui_mau=True,
    client_vn="Công ty Cổ phần Kỹ nghệ Đô Thành",
    address_vn="59-65 Huỳnh Mẫn Đạt, phường An Đông, Thành phố Hồ Chí Minh",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2, 16/11/2026)
    reg_date="12/11/2026", dates=["12/11/2026"], time_bckq="Ngày 12/11/2026",
    site_vn="Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3, 49 Pasteur, phường Sài Gòn, Thành phố Hồ Chí Minh",
    site_bckq="49 Pasteur, phường Sài Gòn, Thành phố Hồ Chí Minh (mẫu do khách hàng mang trực tiếp đến Trung tâm Kỹ "
              "thuật 3)",
    scope_vn="Giám định mẫu nước cấp",
    procedure="QUATEST 3 3030:2025",
    equip=[], executor="Nghĩa", witness_vn="-",
    ngay_nhan_mau="12/11/2026", nguoi_gui="Lê Thị Khánh Vân",
    req_date="20/11/2026",
    highlight=["Mẫu nước sạch sử dụng cho mục đích sinh hoạt – X. PET",
               "Mẫu nước sạch sử dụng cho mục đích sinh hoạt – X. Khuôn"],
    samples=[mau(1, "X. PET", "NSH (X. PET)", "MAU XUONG PET  .jpg"),
             mau(2, "X. Khuôn", "NSH (X. KHUON)", "MAU XUONG KHUON .jpg")],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
