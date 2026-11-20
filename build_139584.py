# -*- coding: utf-8 -*-
r"""
Q326D139584 — Công ty Cổ phần Đầu tư Xây dựng & Thương mại Phú Điền — nước thải tháng 11 (4 mẫu, 4 chứng thư):
N1 nước thải đầu ra NM Bình Hưng (GPMT 3336), N2 / N3 nước thải đầu vào / đầu ra NM Tham Lương – Bến Cát (N3 theo GP 157),
N4 bùn thải NM Bình Hưng (QCVN 07:2025/BNNMT). CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D139584 - Phú Điền - Nước thải Tháng 11` về Input\139584, chỉ đọc):
  - `139584 - BBLM - BÌNH HƯNG.pdf` (2 trang) — NỘI DUNG LÀ BBLM NM Tham Lương – Bến Cát (tên file đặt ngược): 04/11/2026,
    08:50 – 09:36, PYC Q326D139584 ngày 02/11/2026, "theo hợp đồng số 12/KT3-DL26"; Mettler Seven2Go S2; N2.TC.N.06;
    chỉ định: anh Trần Ngọc Tiến – quản lý nhà máy; N2 (đầu vào) 09:05 pH 7,46 – 29,9 °C; N3 (đầu ra) 09:23 pH 7,68 – 29,3 °C.
  - `139584 - BBLM - THAM LƯƠNG.pdf` (2 trang) — NỘI DUNG LÀ BBLM NM Bình Hưng: 11:00 – 11:30; Hach DR300 + Mettler;
    N2.TC.N.06, N2.TC.B.01; chỉ định: anh Võ Trung Trực – kỹ sư vận hành; N1 (đầu ra) 11:09 pH 7,50 – 31,2 °C, clo dư 0,3;
    N4 bùn thải. Người thực hiện cả hai BBLM: Nguyễn Hoàng Việt.
  - `PYC-Q326D139584.N2.1.pdf` (PYC thử nghiệm MTTN26020826): ngày dự kiến trả kết quả 18/11/2026.
  - Hồ sơ tháng trước cùng khách `Q326D122373 - TBKQ\3. TBKQ - 122373 - 1…4.docx` (danh sách thông số, phương pháp,
    mức GPMT 3336 cho N1, GP 157 bảng 1 cho N3, ngưỡng bùn thải cho N4) — AGENTS mục 1 "hồ sơ lặp lại": chỉ lấy khung, mức.
  - Ảnh `Q326D139584 - Photo\N1…N4\Nn - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026):
  - Tên 2 file BBLM trên Z đặt ngược: "khoan hãy động tới, để tôi check lại".
  - Danh sách thông số / mức N1, N3, ngưỡng N4 lấy theo chứng thư tháng 10 (122373) — tô vàng cột mức để soát.
  - TDS N2, N3 để ……… (đồng ý).

Suy luận của AI (cần người dùng xác nhận):
  1. Danh sách thông số, phương pháp và mức đánh giá N1, N3, ngưỡng N4 lấy theo chứng thư tháng 10 (122373).
  2. TDS N2, N3: tháng 10 đo hiện trường bằng Hanna HI 98192, BBLM tháng 11 không ghi máy TDS và không có số -> để "………".
  3. Ngưỡng bùn thải Cu, Zn, Cr VI ghi "Không quy định" như chứng thư 122373/4.
  4. Thời điểm lấy mẫu N4 không ghi trên BBLM -> ghi khung 11:00 – 11:30.

Chạy:  python build_139584.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, params_from_draft, limits_from_draft

IN = os.path.join(WORKSPACE, "Input", "139584")
PH = lambda k: os.path.join(IN, "Q326D139584 - Photo", k, f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
OLD = lambda n: os.path.join(IN, "Q326D122373 - TBKQ", f"3. TBKQ - 122373 - {n}.docx")   # noqa: E731
HCM = "Thành phố Hồ Chí Minh"
BH = f"Nhà máy xử lý nước thải Bình Hưng, ấp 37, xã Bình Hưng, {HCM}"
TL = f"Nhà máy xử lý nước thải lưu vực Tham Lương – Bến Cát, phường An Phú Đông, {HCM}"
GPMT = "phụ lục kèm theo Giấy phép môi trường (Cấp điều chỉnh lần 1) số 3336/GPMT-SNNMT-CCBVMT ngày 15/05/2026"
GP157 = "Bảng 1 - Giấy phép xả nước thải vào nguồn nước số 157/GP-BTNMT ngày 22/01/2018"
CT_TL = ("pH, nhiệt độ, BOD₅, COD, TSS, tổng chất rắn hòa tan, asen, thủy ngân, chì, cadimi, crom VI, crom III, đồng, kẽm, "
         "niken, mangan, sắt, tổng xyanua, tổng phenol, tổng dầu mỡ khoáng, sunfua, amoni, tổng nitơ, tổng phốt pho, "
         "Coliform, tổng các chất hoạt động bề mặt, dầu mỡ động thực vật (27 thông số)")
LEACH = "US EPA 1311 & US EPA 200.7"

CONFIG = dict(
    code6="139584", loai="nuoc thai", kind="nt",
    client_vn="Công ty Cổ phần Đầu tư Xây dựng & Thương mại Phú Điền",
    address_vn=f"Số 31 Tân Canh, phường Tân Sơn Hòa, {HCM}",
    reg_date="02/11/2026", dates=["04/11/2026"],
    time_bckq="Ngày 04/11/2026 từ 08:50 đến 09:36 (Nhà máy Tham Lương – Bến Cát) và từ 11:00 đến 11:30 "
              "(Nhà máy Bình Hưng)",
    site_vn=BH, site_bckq=f"{TL}; {BH}",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    scope_bckq="Lấy mẫu và giám định chất lượng nước thải, bùn thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06", procedure_bckq="QUATEST 3 3030:2025, N2.TC.N.06, N2.TC.B.01",
    equip=["mettler_s2", "hach_dr300"], executor="Việt",
    witness_vn="anh Trần Ngọc Tiến – quản lý Nhà máy xử lý nước thải lưu vực Tham Lương – Bến Cát và anh Võ Trung Trực – "
               "kỹ sư vận hành Nhà máy xử lý nước thải Bình Hưng",
    req_date="18/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR300"},
    samples=[
        dict(ky_hieu="139584-N1", kqtn_key="N1", time="11:09", site_vn=BH, equip=["mettler_s2", "hach_dr300"],
             vi_tri="Điểm nước thải đầu ra tại nhà máy Bình Hưng",
             chi_tieu=f"33 chất ô nhiễm (trừ lưu lượng) tại {GPMT}",
             params=params_from_draft(OLD(1)), limits=limits_from_draft(OLD(1)), hl_limits=True,
             standard_vn=GPMT[0].upper() + GPMT[1:],
             limit_head="Giới hạn cho phép ^(1)",
             limit_note=f"Giá trị giới hạn tối đa cho phép của các chất ô nhiễm quy định tại {GPMT}.",
             intro=f"So sánh kết quả thử nghiệm các chất ô nhiễm của mẫu nước thải nêu trên với mức quy định tương ứng "
                   f"tại {GPMT}:", eval_words=("Chất ô nhiễm", "không đạt", "đạt"),
             field={"pH": "7,50", "Nhiệt độ": "31,2", "Clo dư": "0,3"}, photo=PH("N1")),
        dict(ky_hieu="139584-N2", kqtn_key="N2", time="09:05", site_vn=TL, equip=["mettler_s2"],
             vi_tri="Điểm nước thải đầu vào tại nhà máy xử lý nước thải lưu vực Tham Lương - Bến Cát", chi_tieu=CT_TL,
             params=params_from_draft(OLD(2)), field={"pH": "7,46", "Nhiệt độ": "29,9"}, photo=PH("N2")),
        dict(ky_hieu="139584-N3", kqtn_key="N3", time="09:23", site_vn=TL, equip=["mettler_s2"],
             vi_tri="Điểm nước thải đầu ra tại nhà máy xử lý nước thải lưu vực Tham Lương - Bến Cát", chi_tieu=CT_TL,
             params=params_from_draft(OLD(3)), limits=limits_from_draft(OLD(3)), hl_limits=True, standard_vn=GP157,
             limit_head="Giới hạn cho phép ^(1)", limit_note=f"Giới hạn tối đa cho phép tại {GP157}.",
             intro=f"So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên với mức quy định tương ứng tại "
                   f"{GP157}:",
             field={"pH": "7,68", "Nhiệt độ": "29,3"}, photo=PH("N3")),
        dict(ky_hieu="139584-N4", kqtn_key="N4", time="11:00 – 11:30", site_vn=BH, equip=[],
             scope_vn="Lấy mẫu và giám định chất lượng bùn thải", procedure="QUATEST 3 3030:2025, N2.TC.B.01",
             standard_vn="QCVN 07:2025/BNNMT - Quy chuẩn kỹ thuật quốc gia về ngưỡng chất thải nguy hại",
             vi_tri="Điểm lấy mẫu bùn thải tại nhà máy Bình Hưng",
             chi_tieu="Thành phần nguy hại theo nồng độ ngâm chiết: asen, cadmi, đồng, chì, kẽm, niken, thủy ngân, crom VI",
             params=[("Asen (As)", "mg/L", LEACH), ("Niken (Ni)", "mg/L", LEACH), ("Cadmi (Cd)", "mg/L", LEACH),
                     ("Chì (Pb)", "mg/L", LEACH), ("Đồng (Cu)", "mg/L", LEACH), ("Kẽm (Zn)", "mg/L", LEACH),
                     ("Thủy ngân (Hg)", "mg/L", "US EPA 1311 & SMEWW 3112B:2023"),
                     ("Crom VI (Cr⁶⁺)", "mg/L", "US EPA 1311 & SMEWW 3500-Cr B:2023")],
             limits={"Asen (As)": "≥ 5,0", "Niken (Ni)": "≥ 70", "Cadmi (Cd)": "≥ 1,0", "Chì (Pb)": "≥ 5,0",
                     "Đồng (Cu)": "Không quy định", "Kẽm (Zn)": "Không quy định", "Thủy ngân (Hg)": "≥ 0,2",
                     "Crom VI (Cr⁶⁺)": "Không quy định"}, hl_limits=True,
             cols=["name", "method", "result", "limit"],
             headers=["STT", "Thành phần nguy hại", "Phương pháp", "Kết quả, mg/L", "Ngưỡng CTNH ^(1), mg/L"],
             widths=[8, 32, 30, 14, 16],
             limit_note="Ngưỡng chất thải nguy hại tính theo nồng độ ngâm chiết tại QCVN 07:2025/BNNMT, bảng 2.",
             intro="So sánh kết quả thử nghiệm nồng độ ngâm chiết của các thành phần nguy hại với các ngưỡng nguy hại "
                   "tương ứng nêu tại QCVN 07:2025/BNNMT:",
             eval_words=("Chỉ tiêu", "cao hơn mức quy định", "thấp hơn mức quy định"), photo=PH("N4")),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
