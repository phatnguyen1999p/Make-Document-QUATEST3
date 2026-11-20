# -*- coding: utf-8 -*-
r"""
Q326D131229 — Công ty TNHH L.S Pack — nước uống tháng 11/2026: 6 mẫu nước máy lọc LSP1 – LSP6 (QCVN 06-1:2010/BYT
phụ lục II, III), 6 chứng thư /1 – /6, BCKQ gộp. ĐÃ CÓ KQTN (MTTN26019574.01 – .06, ngày 16/11/2026).

Nguồn (chép từ Z `Q326D0131229 - L.S Pack - Nước uống tháng 11.2026` về Input\131229, chỉ đọc):
  - `0. BBLM.pdf` (2 trang, scan): BBLM ngày 03/11/2026, PYC Q326D131229 ngày 20/10/2026; "Lấy mẫu, giám định chất lượng
    nước uống"; địa điểm Nhà máy 1 (Lô A2-2, KCN Tây Bắc Củ Chi…) và Nhà máy 2 (Lô C2-3-3, đường D4, KCN Tây Bắc Củ
    Chi…); 9h00 – 10h30; thiết bị "-"; QTGĐ 3030.2025, N2.TC.N.07; vị trí viết tay: LSP1 Máy lọc nước Karofi A3, LSP2
    Máy lọc nước Suntech A1, LSP3 Máy lọc nước Karofi A2, LSP4 Máy lọc nước Karofi 1 - Nhà máy 2, LSP5 Máy lọc nước
    Karofi 2 - Nhà máy 2, LSP6 Máy lọc nước Suntech - Nhà máy 2 (không niêm phong, không ghi giờ từng mẫu); "Chỉ tiêu
    thực hiện: 21 chỉ tiêu hóa lý và 5 chỉ tiêu vi sinh QCVN 06-1:2010/BYT, phụ lục II + III"; đại diện KH ký: Trần
    Thị Bông; đại diện TTKT3 ký: "Hoàng Minh".
  - `PYC-Q326D131229.N2.1.pdf` (MTTN26019574, nhận 20/10/2026, dự kiến trả 14/11/2026): LSP1 – LSP6, 26 chỉ tiêu.
  - `Q326D131229 - TBKQ\77708 - edoc\KQ-Q326D131229-MTTN26019574.01 … .06.pdf`: KQTN ngày 16/11/2026, nhận mẫu
    20/10/2026; kết quả vi sinh "<1⁽ᶜ⁾" (⁽ᶜ⁾: "< 1 CFU" xem là không phát hiện).
  - `Payment Q326D131229.pdf`: báo giá — MST 0302419037, "Số lượng: 8".
  - Khung làm tay trên Z `3. TBKQ - 1…6 - vn/en - Q326D131229.docx`, `2. BCKQ - Q326D131229.docx` (09/11/2026): vị trí,
    địa điểm từng nhà máy, mức QCVN lấy từ đây; tên chỉ tiêu, phương pháp, kết quả theo KQTN.
  - Ảnh `0. vi tri\LSP1 … LSP6.jpg`.

Quyết định của người dùng (17/11/2026):
  - Có thêm bản TBKQ tiếng Anh (chứng thư Việt - Anh).
  - Người thực hiện Minh: đúng (GĐV ký Nguyễn Xuân Nam).
  - Ngày gửi mẫu 20/10/2026 (theo KQTN) và "KQTN có trễ hạn: có": giữ như vậy, tô vàng.
  - Làm 6 chứng thư (báo giá ghi số lượng 8).

Suy luận của AI (cần người dùng xác nhận):
  1. Tên EN theo MST 0302419037 (tax_helper: "L.S PACK COMPANY LIMITED"), địa chỉ / địa điểm / vị trí EN theo khung
     làm tay; tên chỉ tiêu EN dịch từ tên KQTN (NAMES_EN).
  2. Địa điểm giám định theo khung làm tay: LSP1 – 3 Nhà máy L.S Pack 1, LSP4 – 6 Nhà máy L.S Pack 2.
  3. Tên vị trí "Suntech" (BBLM viết tay đọc được "Sunteck"/"Suntech") theo khung làm tay.
  4. Mức QCVN 06-1:2010/BYT theo khung làm tay (vi sinh "Không phát hiện").
  5. Ngày yêu cầu cấp KQTN 14/11/2026 (PYC thử nghiệm); KQTN ngày nhận mẫu 20/10/2026 sớm hơn ngày lấy mẫu 03/11/2026.

Chạy:  python build_131229.py [--pdf] [--only 1,3]

TỪ 18/11/2026: TBKQ VN/EN điền vào FORM GỐC của người dùng bằng `fill_131229.py` (người dùng: "giữ lại form cũ … chỉ điền thôi không cần sửa"). KHÔNG chạy lại script này — nó dựng lại TBKQ theo KQTN (sai thứ tự / tên chỉ tiêu) và BCKQ (người dùng đã sửa BCKQ trên Z 18/11/2026).
"""
import argparse
import os
from nuoc_common import run, WORKSPACE

IN = os.path.join(WORKSPACE, "Input", "131229")
QC06 = "QCVN 06-1:2010/BYT"
KPH_VS = "Không phát hiện"
LIMITS = {
    "Hàm lượng asen (As)": "0,01", "Hàm lượng bari (Ba)": "0,7", "Hàm lượng bor (B)": "0,5",
    "Hàm lượng cadmi (Cd)": "0,003", "Hàm lượng chì (Pb)": "0,01", "Hàm lượng crom (Cr)": "0,05",
    "Hàm lượng cyanid (CN⁻)": "0,07", "Hàm lượng fluorid (F‾)": "1,5", "Hàm lượng mangan (Mn)": "0,4",
    "Hàm lượng molybden (Mo)": "0,07", "Hàm lượng niken (Ni)": "0,07", "Hàm lượng nitrat (NO₃⁻)": "50",
    "Hàm lượng nitrit (NO₂⁻)": "3", "Hàm lượng selen (Se)": "0,01", "Hàm lượng stibi (Sb)": "0,02",
    "Hàm lượng thủy ngân (Hg)": "0,006", "Hàm lượng tổng clo (Cl₂)": "5", "Hàm lượng đồng (Cu)": "2",
    "Hàm lượng bromat (BrO₃⁻)": "0,01", "Hàm lượng clorate (ClO₃⁻)": "0,7", "Hàm lượng clorit (ClO₂⁻)": "0,7",
    "Coliforms": KPH_VS, "Escherichia coli": KPH_VS, "Pseudomonas aeruginosa": KPH_VS,
    "Enterococci (Streptococci faecal)": KPH_VS, "Bào tử vi khuẩn kỵ khí khử sulfit (Clostridia)": KPH_VS,
}
KCN = "khu công nghiệp Tây Bắc Củ Chi, ấp Bàu Tre 2, xã Tân An Hội, Thành phố Hồ Chí Minh"
NM1 = f"Nhà máy L.S Pack 1, lô A2-2, {KCN}"
NM2 = f"Nhà máy L.S Pack 2, lô C2-3-3, đường D4, {KCN}"
CHI_TIEU = f"21 chỉ tiêu hóa lý và 05 chỉ tiêu vi sinh tại {QC06}, phụ lục II và phụ lục III"
KCN_EN = "Tay Bac Cu Chi industrial park, Bau Tre 2 hamlet, Tan An Hoi commune, Ho Chi Minh City"
NM1_EN = f"L.S Pack factory 1, lot A2-2, {KCN_EN}"
NM2_EN = f"L.S Pack factory 2, lot C2-3-3, road D4, {KCN_EN}"
NAMES_EN = {
    "Hàm lượng asen (As)": "Arsenic (As)", "Hàm lượng bari (Ba)": "Barium (Ba)", "Hàm lượng bor (B)": "Boron (B)",
    "Hàm lượng cadmi (Cd)": "Cadmium (Cd)", "Hàm lượng chì (Pb)": "Lead (Pb)", "Hàm lượng crom (Cr)": "Chromium (Cr)",
    "Hàm lượng cyanid (CN⁻)": "Cyanide (CN⁻)", "Hàm lượng fluorid (F‾)": "Fluoride (F‾)",
    "Hàm lượng mangan (Mn)": "Manganese (Mn)", "Hàm lượng molybden (Mo)": "Molybdenum (Mo)",
    "Hàm lượng niken (Ni)": "Nickel (Ni)", "Hàm lượng nitrat (NO₃⁻)": "Nitrate (NO₃⁻)",
    "Hàm lượng nitrit (NO₂⁻)": "Nitrite (NO₂⁻)", "Hàm lượng selen (Se)": "Selenium (Se)",
    "Hàm lượng stibi (Sb)": "Antimony (Sb)", "Hàm lượng thủy ngân (Hg)": "Mercury (Hg)",
    "Hàm lượng tổng clo (Cl₂)": "Total chlorine (Cl₂)", "Hàm lượng đồng (Cu)": "Copper (Cu)",
    "Hàm lượng bromat (BrO₃⁻)": "Bromate (BrO₃⁻)", "Hàm lượng clorate (ClO₃⁻)": "Chlorate (ClO₃⁻)",
    "Hàm lượng clorit (ClO₂⁻)": "Chlorite (ClO₂⁻)",
    "Bào tử vi khuẩn kỵ khí khử sulfit (Clostridia)": "Sulfite-reducing anaerobic spores (Clostridia)",
}


def mau(n, vi_tri, site, vi_tri_en, site_en):
    return dict(kind="nu", cert_no=n, ky_hieu=f"LSP{n}", kqtn_key=f"LSP {n}", vi_tri=vi_tri, site_vn=site,
                vi_tri_en=vi_tri_en, site_en=site_en,
                chi_tieu_en=f"21 physicochemical and 05 microbiological parameters in {QC06}, annex II and annex III",
                # rút gọn để EN cùng số dòng VN, VN = EN số trang (AGENTS mục 3, 9)
                standard_en=f"{QC06} - National technical regulation for natural mineral water and bottled drinking water",
                limit_head_en="Maximum limit ^(1)",
                limit_note_en=f"Maximum allowable limit specified in annex II and annex III, {QC06}.",
                intro_en=f"Comparing the test result of parameters of drinking water sample mentioned above with the "
                         f"corresponding specification in {QC06}, annex II and annex III:",
                time="-", field={}, field_tbkq=[], equip=[], chi_tieu=CHI_TIEU, limits=LIMITS,
                photo=os.path.join(IN, "0. vi tri", f"LSP{n}.jpg"),
                standard_vn=f"{QC06} - Quy chuẩn kỹ thuật quốc gia đối với nước khoáng thiên nhiên và nước uống đóng chai",
                limit_head="Giới hạn tối đa ^(1)",
                limit_note=f"Mức giới hạn tối đa cho phép quy định tại phụ lục II và phụ lục III, {QC06}.",
                eval_words=("Chỉ tiêu", "không đạt", "đạt"),
                intro=f"So sánh kết quả thử nghiệm các chỉ tiêu của mẫu nước uống nêu trên so với mức quy định tương "
                      f"ứng tại {QC06}, phụ lục II và phụ lục III:")


CONFIG = dict(
    code6="131229", loai="nuoc uong", kind="nu", word="chỉ tiêu",   # nước uống: cột "Chỉ tiêu" (khung làm tay)
    cert_date=None,                                 # có KQTN -> ngày KQTN muộn nhất (AGENTS mục 2)
    en=True, names_en=NAMES_EN,                     # chứng thư Việt - Anh (người dùng chốt 17/11/2026)
    client_vn="Công ty TNHH L.S Pack", client_en="L.S Pack Company Limited",
    address_en="Lot A2-2, Tay Bac Cu Chi industrial park, Bau Tre 2 hamlet, Tan An Hoi commune, Ho Chi Minh City",
    scope_en="Sampling and quality inspection of drinking water",
    address_vn="Lô A2-2, Khu công nghiệp Tây Bắc Củ Chi, ấp Bàu Tre 2, xã Tân An Hội, Thành phố Hồ Chí Minh",
    reg_date="20/10/2026", dates=["03/11/2026"], time_bckq="Ngày 03/11/2026 từ 09:00 đến 10:30",
    site_vn=NM1,
    site_bckq=f"{NM1}; {NM2}",
    scope_vn="Lấy mẫu và giám định chất lượng nước uống",
    procedure="QUATEST 3 3030:2025, N2.TC.N.07",
    equip=[], executor="Minh",
    witness_vn="chị Trần Thị Bông",
    req_date="14/11/2026",
    highlight_bckq=["Ngày gửi mẫu", "KQTN có trễ hạn"],   # giữ theo KQTN, tô vàng (người dùng 17/11/2026)
    samples=[
        mau(1, "Máy lọc nước Karofi A3", NM1, "Karofi A3 water purifier", NM1_EN),
        mau(2, "Máy lọc nước Suntech A1", NM1, "Suntech A1 water purifier", NM1_EN),
        mau(3, "Máy lọc nước Karofi A2", NM1, "Karofi A2 water purifier", NM1_EN),
        mau(4, "Máy lọc nước Karofi 1 – Nhà máy 2", NM2, "Karofi 1 water purifier – Factory 2", NM2_EN),
        mau(5, "Máy lọc nước Karofi 2 – Nhà máy 2", NM2, "Karofi 2 water purifier – Factory 2", NM2_EN),
        mau(6, "Máy lọc nước Suntech – Nhà máy 2", NM2, "Suntech water purifier – Factory 2", NM2_EN),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
