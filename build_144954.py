# -*- coding: utf-8 -*-
r"""
Q326D144954 — Công ty TNHH Điện máy Orro Home — nước uống (máy lọc Philips) N1 – N3 theo QCVN 06-1:2010/BYT và nước cấp
(bộ lọc Philips) N4 – N5 theo QCVN 01-1:2024/BYT; 5 chứng thư /1 – /5, BCKQ gộp.
CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `Q326D144954 - ORRO HOME` về Input\144954, chỉ đọc):
  - `0. BBLM.pdf` (2 trang, scan): BBLM ngày 16/11/2026, PYC Q326D144954 ngày 12/11/2026; "CÔNG TY TNHH ĐIỆN MÁY ORRO
    HOME", "Số 27, đường số 4C, khu nhà ở Rạch Bà Tánh, xã Bình Hưng, Thành phố Hồ Chí Minh"; "Lấy mẫu, giám định chất lượng
    nước uống, nước cấp"; địa điểm viết tay "Văn phòng ORRO HOME" + in sẵn "Số 21-23 Phạm Hùng (Nối dài), khu dân cư Đại
    Phúc Green Villas, xã Bình Hưng, Tp.Hồ Chí Minh"; 14h15 – 15h30; thiết bị in sẵn: Mettler Seven2Go S2, Hanna HI 98192,
    Hach DR300; QTGĐ 3030.2025, N2.TC.N.07, N2.TC.N.01; chỉ định: Lê Thị Hồng Anh – nhân viên; 05 vị trí:
      N1 Máy lọc và cung cấp nước nóng Philips AUT9124/90 — nước uống, pH 8,73, TDS 9 mg/L (viết "8?" gạch, sửa 9);
      N2 Máy lọc và cung cấp nước nóng, lạnh Philips ADD6996WH/90 — nước uống, pH 7,37, TDS 6 mg/L;
      N3 Máy lọc và cung cấp nước nóng Philips AUT7005/9 — nước uống, pH 6,67, TDS 3 mg/L;
      N4 Bộ lọc nước tại vòi – Philips AWP3704/97 — nước cấp, pH 7,25, 32,3 °C, clo dư 0,04 mg/L, không có mùi lạ;
      N5 Bộ lọc nước đầu nguồn – Philips AWP1830 — nước cấp, pH 7,27, 31,0 °C, clo dư 0,67 mg/L, không có mùi lạ;
    kiểm tra chuẩn pH 4,01 / 7,00 / 9,99 — đạt; GĐV ký "Hoàng Minh".
  - Ảnh màn hình máy đo `Q326D144954 - photo\đo pH/TDS/Clo - Nn.jpg`: khớp số BBLM (TDS N1 0,009 g/L); máy đo clo trong
    ảnh là Hach DR900 (BBLM in sẵn DR300).
  - `Q326D144954 - thuc hien\PYC-Q326D144954.N2.1.pdf` (MTTN26022136, nhận 14/11/2026, dự kiến trả 30/11/2026): N1 – N3 —
    18 kim loại / anion + TDS (SMEWW 2540C:2023) + clorit, clorat, bromat + vi sinh QCVN 6-1 (5 chỉ tiêu).
  - `PYC-Q326D144954.N2.2.pdf` (MTTN26022134, nhận 14/11/2026, dự kiến trả 24/11/2026): N4, N5 — permanganate, amoni, As,
    độ màu, độ đục + vi sinh QCVN 01-1:2024 (Coliforms, E. coli, P. aeruginosa, S. aureus).
  - Khung làm tay `3. TBKQ - 1.docx` (chép từ 124254 cùng khách, nước uống) và `3. TBKQ - 4.docx` (chép từ 096118, nước
    sinh hoạt): mức QCVN lấy từ đây; tên chỉ tiêu, phương pháp theo PYC / KQTN.
  - Ảnh vị trí `0. vi tri\N1 … N5.jpg`.

Quyết định của người dùng (17/11/2026):
  - Máy đo clo: không đổi (Hach DR300 theo BBLM). Địa chỉ khách: ghi theo BBLM. Chỉ dựng bản tiếng Việt.
  - Giữ "nước cấp" cho N4 – N5 và "pH, TDS; 22 chỉ tiêu hóa lý và 05 chỉ tiêu vi sinh" cho N1 – N3.

Suy luận của AI (đã được người dùng xem 17/11/2026):
  1. Người thực hiện "Minh" theo chữ ký BBLM -> GĐV ký Nguyễn Xuân Nam.
  2. Địa chỉ khách theo BBLM (số 27, đường số 4C); khung 124254 ghi "Số 44, đường số 6A".
  3. Máy đo clo theo BBLM (Hach DR300), ảnh màn hình là DR900.
  4. Clo dư N4 đọc 0,04 -> ghi "< 0,1" (AGENTS 8C nước NSH điểm 5) -> N4 không đạt khoảng 0,2 – 1,0 khi đánh giá.
  5. pH, TDS của nước uống: QCVN 06-1 không quy định -> chỉ tiêu không đánh giá.
  6. Chỉ dựng TBKQ tiếng Việt (không có PYC giám định để xem mục 15; 124254 chỉ có bản VN).
  7. Nội dung yêu cầu: N1 – N3 "nước uống", N4 – N5 "nước cấp" (theo BBLM).

Chạy:  python build_144954.py [--pdf] [--only 1,4]
"""
import argparse
import datetime
import os
from nuoc_common import run, WORKSPACE

IN = os.path.join(WORKSPACE, "Input", "144954")
P = "………"
F = "Đo tại hiện trường"
QC06, QC01 = "QCVN 06-1:2010/BYT", "QCVN 01-1:2024/BYT"
SITE = ("Văn phòng Orro Home, số 21-23 Phạm Hùng (nối dài), khu dân cư Đại Phúc Green Villas, xã Bình Hưng, "
        "Thành phố Hồ Chí Minh")
E200, A4110B, A4110D = "US EPA 200.8", "SMEWW 4110B:2023", "SMEWW 4110D:2023"
ISO9308 = "ISO 9308-1:2014/Amd 1:2016"
NU_PARAMS = [
    ("pH", "-", F), ("TDS", "mg/L", F),
    ("Hàm lượng asen (As)", "mg/L", E200), ("Hàm lượng bari (Ba)", "mg/L", E200), ("Hàm lượng bor (B)", "mg/L", E200),
    ("Hàm lượng cadmi (Cd)", "mg/L", E200), ("Hàm lượng chì (Pb)", "mg/L", E200), ("Hàm lượng crom (Cr)", "mg/L", E200),
    ("Hàm lượng cyanid (CN⁻)", "mg/L", "TCVN 6181:1996"), ("Hàm lượng fluorid (F‾)", "mg/L", A4110B),
    ("Hàm lượng mangan (Mn)", "mg/L", E200), ("Hàm lượng molybden (Mo)", "mg/L", E200),
    ("Hàm lượng niken (Ni)", "mg/L", E200), ("Hàm lượng nitrat (NO₃⁻)", "mg/L", A4110B),
    ("Hàm lượng nitrit (NO₂⁻)", "mg/L", A4110B), ("Hàm lượng selen (Se)", "mg/L", E200),
    ("Hàm lượng stibi (Sb)", "mg/L", E200), ("Hàm lượng thủy ngân (Hg)", "mg/L", E200),
    ("Hàm lượng tổng clo (Cl₂)", "mg/L", "TCVN 6225-2:2021"), ("Hàm lượng đồng (Cu)", "mg/L", E200),
    ("Tổng chất rắn hòa tan (TDS)", "mg/L", "SMEWW 2540C:2023"),
    ("Hàm lượng bromat (BrO₃⁻)", "mg/L", A4110D), ("Hàm lượng clorate (ClO₃⁻)", "mg/L", A4110D),
    ("Hàm lượng clorit (ClO₂⁻)", "mg/L", A4110D),
    ("Coliforms", "CFU/ 250 mL", ISO9308), ("Escherichia coli", "CFU/ 250 mL", ISO9308),
    ("Pseudomonas aeruginosa", "CFU/ 250 mL", "ISO 16266:2006"),
    ("Enterococci (Streptococci faecal)", "CFU/ 250 mL", "ISO 7899-2:2000 (QCVN 06-1:2010/BYT)"),
    ("Bào tử vi khuẩn kỵ khí khử sulfit (Clostridia)", "CFU/ 50 mL", "ISO 6461-2:1986"),
]
KPH_VS = "Không phát hiện"
NU_LIMITS = {                                       # QCVN 06-1:2010/BYT phụ lục II, III (khung 124254)
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
NSH_PARAMS = [
    ("pH", "-", F), ("Clo dư tự do", "mg/L", F),
    ("Độ màu", "Pt.Co", "SMEWW 2120C:2023"), ("Độ đục", "NTU", "SMEWW 2130B:2023"),
    ("Hàm lượng asen (As)", "mg/L", E200), ("Chỉ số permanganate", "mg/L", "TCVN 6186:1996"),
    ("Hàm lượng amoni (NH₃ và NH₄⁺ tính theo N)", "mg/L", "US EPA 350.2 (So màu)"),
    ("Mùi", "-", "Cảm quan tại hiện trường"),
    ("Coliforms", "CFU/ 100 mL", ISO9308), ("Escherichia coli", "CFU/ 100 mL", ISO9308),
    ("Pseudomonas aeruginosa", "CFU/ 100 mL", "ISO 16266:2006"),
    ("Staphylococcus aureus", "CFU/ 100 mL", "SMEWW 9213B:2023"),
]
NSH_LIMITS = {                                      # QCVN 01-1:2024/BYT (khung 096118 / 145103)
    "pH": "Trong khoảng\n(6,0 – 8,5)", "Clo dư tự do": "Trong khoảng\n(0,2 – 1,0)", "Độ màu": "15", "Độ đục": "2",
    "Hàm lượng asen (As)": "0,01", "Chỉ số permanganate": "2", "Hàm lượng amoni (NH₃ và NH₄⁺ tính theo N)": "1",
    "Mùi": "Không có mùi lạ", "Coliforms": "< 1", "Escherichia coli": "< 1", "Pseudomonas aeruginosa": "< 1",
    "Staphylococcus aureus": "< 1",
}


def nu(n, vi_tri, ph, tds):
    return dict(kind="nu", loai="nuoc uong", cert_no=n, ky_hieu=f"144954-N{n}", kqtn_key=f"N{n}", vi_tri=vi_tri,
                time="-", field={"pH": ph, "TDS": tds}, field_tbkq=["pH", "TDS"], equip=["mettler_s2", "hanna_hi98192"],
                params=NU_PARAMS, limits=NU_LIMITS, photo=os.path.join(IN, "0. vi tri", f"N{n}.jpg"),
                scope_vn="Lấy mẫu và giám định chất lượng nước uống", procedure="QUATEST 3 3030:2025, N2.TC.N.07",
                chi_tieu=f"- pH, TDS;\n- 22 chỉ tiêu hóa lý và 05 chỉ tiêu vi sinh nêu tại {QC06}, phụ lục II và phụ "
                         f"lục III.",
                standard_vn=f"{QC06} - Quy chuẩn kỹ thuật quốc gia đối với nước khoáng thiên nhiên và nước uống đóng chai",
                limit_head="Giới hạn tối đa ^(1)",
                limit_note=f"Mức giới hạn tối đa cho phép quy định tại phụ lục II và phụ lục III, {QC06}.",
                eval_words=("Chỉ tiêu", "không đạt", "đạt"),
                intro=f"So sánh kết quả thử nghiệm các chỉ tiêu của mẫu nước uống nêu trên so với mức quy định tương "
                      f"ứng tại {QC06}, phụ lục II và phụ lục III:")


def nsh(n, vi_tri, ph, t, clo):
    return dict(kind="nsh", loai="nuoc cap", word="thông số", cert_no=n, ky_hieu=f"144954-N{n}", kqtn_key=f"N{n}", vi_tri=vi_tri,
                time="-", field={"pH": ph, "Nhiệt độ": t, "Clo dư tự do": clo, "Mùi": "Không có mùi lạ"},
                field_tbkq=["pH", "Clo dư tự do", "Mùi"], equip=["mettler_s2", "hach_dr300"],
                params=NSH_PARAMS, limits=NSH_LIMITS, photo=os.path.join(IN, "0. vi tri", f"N{n}.jpg"),
                scope_vn="Lấy mẫu và giám định chất lượng nước cấp", procedure="QUATEST 3 3030:2025, N2.TC.N.01",
                chi_tieu=f"10 thông số nhóm A, Pseudomonas aeruginosa, Staphylococcus aureus nêu tại {QC01}",
                standard_vn=f"{QC01} - Quy chuẩn kỹ thuật quốc gia về chất lượng nước sạch sử dụng cho mục đích sinh "
                            f"hoạt",
                limit_head="Ngưỡng giới hạn cho phép ^(1)",
                limit_note=f"Ngưỡng giới hạn tối đa cho phép quy định tại {QC01}.",
                intro=f"So sánh kết quả thử nghiệm các thông số của mẫu nước nêu trên so với mức quy định tương ứng tại "
                      f"{QC01}:")


CONFIG = dict(
    code6="144954", loai="nuoc", kind="nu", word="chỉ tiêu",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2, 16/11/2026)
    client_vn="Công ty TNHH Điện máy Orro Home",
    address_vn="Số 27, đường số 4C, khu nhà ở Rạch Bà Tánh, xã Bình Hưng, Thành phố Hồ Chí Minh",
    reg_date="12/11/2026", dates=["16/11/2026"], time_bckq="Ngày 16/11/2026 từ 14:15 đến 15:30",
    site_vn=SITE,
    scope_vn="Lấy mẫu và giám định chất lượng nước uống",
    scope_bckq="Lấy mẫu và giám định chất lượng nước uống, nước cấp",
    procedure="QUATEST 3 3030:2025, N2.TC.N.07",
    procedure_bckq="QUATEST 3 3030:2025, N2.TC.N.07, N2.TC.N.01",
    equip=["mettler_s2", "hanna_hi98192", "hach_dr300"], executor="Minh",
    witness_vn="chị Lê Thị Hồng Anh – nhân viên",
    req_date="24/11/2026 (N4, N5); 30/11/2026 (N1 – N3)",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "TDS": "Hanna HI 98192", "Nhiệt độ": "Mettler Toledo Seven2Go S2",
                   "Clo dư tự do": "Hach DR300", "Mùi": "-"},
    bckq_notes=["- Clo dư tự do mẫu N4 đọc trên máy 0,04 mg/L (dưới 0,1 mg/L) -> ghi \"< 0,1\"; TDS mẫu N1 máy hiển thị "
                "0,009 g/L = 9 mg/L."],
    samples=[
        nu(1, "Máy lọc và cung cấp nước nóng Philips AUT9124/90", "8,73", "9"),
        nu(2, "Máy lọc và cung cấp nước nóng, lạnh Philips ADD6996WH/90", "7,37", "6"),
        nu(3, "Máy lọc và cung cấp nước nóng Philips AUT7005/9", "6,67", "3"),
        nsh(4, "Bộ lọc nước tại vòi – Philips AWP3704/97", "7,25", "32,3", "< 0,1"),
        nsh(5, "Bộ lọc nước đầu nguồn – Philips AWP1830", "7,27", "31,0", "0,67"),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
