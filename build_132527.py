# -*- coding: utf-8 -*-
r"""
Q326D132527 — Công ty TNHH KPack — nước uống tháng 11/2026: 2 mẫu nước máy lọc KP1, KP2 (QCVN 06-1:2010/BYT phụ lục
II, III), 2 chứng thư /1, /2 Việt - Anh, BCKQ gộp. ĐÃ CÓ KQTN (MTTN26019547.01, .02 ngày 16/11/2026).

Nguồn (chép từ Z `Q326D132527 - K Pack - Nước uống tháng 11.2026` về Input\132527, chỉ đọc):
  - `0. BBLM.pdf` (1 trang, scan): BBLM ngày viết tay đọc là 03/11/2026 (khung làm tay ghi 05/11/2026), PYC Q326D132527
    ngày 19/10/2026; "Công ty TNHH Kpack", "Lô B1-5, đường D2, KCN Tây Bắc Củ Chi, xã Tân An Hội, Tp. HCM"; "Lấy mẫu,
    giám định chất lượng nước uống"; địa điểm = địa chỉ; 10h45 – 11h15; thiết bị "-"; QTGĐ 3030.2025, N2.TC.N.07;
    KP1 "Máy lọc nước Karofi B3", KP2 "Máy lọc nước B1" (không niêm phong); "21 chỉ tiêu hóa lý và 05 chỉ tiêu vi sinh
    tại QCVN 06-1:2010/BYT, PL II + III"; đại diện KH ký đọc được "Lê Tuấn Khanh" (chưa chắc); TTKT3 ký "Hoàng Minh".
  - `PYC-Q326D132527.N2.1.pdf` (MTTN26019547, nhận 20/10/2026, dự kiến trả 14/11/2026).
  - `Q326D132527 - TBKQ\Edoc\KQ-Q326D132527-MTTN26019547.01, .02.pdf`: KQTN 16/11/2026, nhận mẫu 20/10/2026.
  - `Payment Q326D132527.pdf`: "CÔNG TY TNHH KPACK", MST 3700635434 (tên quốc tế "KPACK COMPANY LIMITED").
  - Khung làm tay `3. TBKQ - 1/2 - vn/en - Q326D132527.docx` (09/11/2026): vị trí VN/EN, Việt - Anh.
  - Ảnh `0. vi tri\KP1.jpg`, `KP2.jpg`.

Quyết định của người dùng (17/11/2026):
  - Ngày giám định: "ghi ngày 3 đi nhưng highlight lên" -> 03/11/2026, tô vàng (TBKQ VN/EN, BCKQ).
  - Người chỉ định: "Lê Tuấn Khanh, chắc không cần ghi chữ anh".
  - Địa điểm: giữ như BBLM.
  - Trễ hạn: "không ghi trễ hạn, highlight lên … cứ để thời gian đúng như bình thường" -> BCKQ "[X] không", ngày gửi mẫu
    20/10/2026 giữ, cả hai tô vàng.

Suy luận của AI (đã được người dùng xem 17/11/2026):
  1. Ngày giám định 03/11/2026 theo BBLM (khung làm tay 05/11/2026).
  2. Người chỉ định (BCKQ) "anh Lê Tuấn Khanh" đọc từ chữ ký BBLM — chưa chắc, tô vàng.
  3. Địa điểm giám định ghi như BBLM (không có tên đơn vị -> QA [Location Short]).
  4. Ngày gửi mẫu 20/10/2026 theo KQTN (trước ngày lấy mẫu) -> "KQTN có trễ hạn: có" — như 131229, tô vàng.
  5. Người thực hiện Minh (chữ ký BBLM) -> GĐV ký Nguyễn Xuân Nam. Mức QCVN, tên chỉ tiêu EN như 131229.
  6. Địa chỉ EN dùng "industrial park", "commune" (khung ghi "industrial zone", "ward").

Chạy:  python build_132527.py [--pdf]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE
from build_131229 import LIMITS, NAMES_EN, QC06, CHI_TIEU

IN = os.path.join(WORKSPACE, "Input", "132527")
SITE = "Lô B1-5, đường D2, Khu công nghiệp Tây Bắc Củ Chi, xã Tân An Hội, Thành phố Hồ Chí Minh"
SITE_EN = "Lot B1-5, road D2, Tay Bac Cu Chi industrial park, Tan An Hoi commune, Ho Chi Minh City"


def mau(n, vi_tri, vi_tri_en):
    return dict(kind="nu", cert_no=n, ky_hieu=f"KP{n}", kqtn_key=f"KP{n}", vi_tri=vi_tri, vi_tri_en=vi_tri_en,
                time="-", field={}, field_tbkq=[], equip=[], chi_tieu=CHI_TIEU, limits=LIMITS,
                order=list(LIMITS),                 # cùng thứ tự chỉ tiêu cho KP1, KP2 (phiếu KP2 xếp vi sinh lên đầu)
                photo=os.path.join(IN, "0. vi tri", f"KP{n}.jpg"),
                standard_vn=f"{QC06} - Quy chuẩn kỹ thuật quốc gia đối với nước khoáng thiên nhiên và nước uống đóng chai",
                standard_en=f"{QC06} - National technical regulation for natural mineral water and bottled drinking water",
                limit_head="Giới hạn tối đa ^(1)", limit_head_en="Maximum limit ^(1)",
                limit_note=f"Mức giới hạn tối đa cho phép quy định tại phụ lục II và phụ lục III, {QC06}.",
                limit_note_en=f"Maximum allowable limit specified in annex II and annex III, {QC06}.",
                chi_tieu_en=f"21 physicochemical and 05 microbiological parameters in {QC06}, annex II and annex III",
                eval_words=("Chỉ tiêu", "không đạt", "đạt"),
                intro=f"So sánh kết quả thử nghiệm các chỉ tiêu của mẫu nước uống nêu trên so với mức quy định tương "
                      f"ứng tại {QC06}, phụ lục II và phụ lục III:",
                intro_en=f"Comparing the test result of parameters of drinking water sample mentioned above with the "
                         f"corresponding specification in {QC06}, annex II and annex III:")


CONFIG = dict(
    code6="132527", loai="nuoc uong", kind="nu", word="chỉ tiêu",
    en=True, names_en=NAMES_EN,
    cert_date=None,                                 # có KQTN -> ngày KQTN muộn nhất (AGENTS mục 2)
    client_vn="Công ty TNHH KPack", client_en="KPack Company Limited",
    address_vn=SITE, address_en=SITE_EN,
    reg_date="19/10/2026", dates=["03/11/2026"], time_bckq="Ngày 03/11/2026 từ 10:45 đến 11:15",
    site_vn=SITE, site_en=SITE_EN,
    scope_vn="Lấy mẫu và giám định chất lượng nước uống", scope_en="Sampling and quality inspection of drinking water",
    procedure="QUATEST 3 3030:2025, N2.TC.N.07",
    equip=[], executor="Minh",
    witness_vn="Lê Tuấn Khanh",
    req_date="14/11/2026",
    highlight=["03/11/2026"], highlight_en=["03/11/2026"],     # ngày giám định theo BBLM, chờ soát
    tre_han="[  ] có		[X] không",                # người dùng 17/11/2026: không ghi trễ hạn, tô vàng
    highlight_bckq=["Ngày gửi mẫu", "KQTN có trễ hạn"],
    samples=[
        mau(1, "Máy lọc nước Karofi B3", "Karofi water purifier B3"),
        mau(2, "Máy lọc nước B1", "Water purifier B1"),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
