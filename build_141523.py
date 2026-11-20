# -*- coding: utf-8 -*-
r"""
Q326D141523 — Công ty TNHH Giấy Lee & Man Việt Nam (Cần Thơ) — 41 mẫu: nước uống NU1 – NU34 (mỗi mẫu một chứng thư
/1 – /34), nước sinh hoạt NSH1A, NSH2A, NSH1B, NSH2B (mỗi mẫu một chứng thư /35 – /38), nước thải NT1 – NT3 (/39 – /41). CHƯA CÓ KQTN -> bản khung; chỉ tiêu thử nghiệm để "………" tô vàng (người dùng chốt).

Nguồn (chép từ Z về Input\141523, chỉ đọc):
  - `0. BBLM.pdf` (4 trang): BBLM Q326D141523/2 ngày 12/11/2026 (PYC Q326D141523/2 ngày 05/11/2026) — "Lấy mẫu và phân tích
    chất lượng nước sinh hoạt và nước uống"; N2.TC.N.02, N2.TC.N.07; Mettler Seven2Go S2, Hanna HI93703 (độ đục), Hanna
    HI98192 (TDS), Hach DR300; chỉ định: chị Huỳnh Thị Thoa – Tổ trưởng hóa nghiệm; NSH1A "Nước sinh hoạt 1 – Nhóm A",
    NSH2A, NSH1B, NSH2B (đo 3 lần); "Nước uống tại vòi của máy lọc nước" NU1 -> NU34 (chỉ lấy mẫu); không ghi giờ.
    BBLM Q326D141523/1 ngày 12/11/2026 (PYC /1 ngày 05/11/2026) — nước thải, 09:30 – 16:30, N2.TC.N.02, N2.TC.N.06;
    NT1 "Nước thải trước hồ sinh học" 10:15 – 10:30, NT2 "Nước thải sau hồ sinh học" 10:40 – 10:50, NT3 "Nước thải hồ thu
    gom" 11:00 – 11:15 (pH, nhiệt độ, clo dư 3 lần; NT3 clo dư "Không xác định"). Người thực hiện: Nguyễn Minh Quý.
  - `0. Danh sách vị trí nước uống.pdf`: NU1 – NU34 "Máy lọc nước - mã số: VNLM202510-01 … -34", mã chứng thư
    Q326D141523/1 … /34.
  - Excel biên bản hiện trường `141523 - thuc hien\141523-1/-2 …xlsx`; ảnh `0. vi tri\NSH1, NSH2, NT1 – NT3, NU1 – NU34.jpg`.
  - `2. BCKQ.docx`, `3. TBKQ - 1/35/37.docx` trên Z là khung của khách LEGO (Q326D100120) — chỉ tham khảo cách chia.

Quyết định của người dùng (16/11/2026):
  - Khung LEGO trên Z sai khách -> làm lại cho Lee & Man.
  - Chia riêng 4 mẫu nước sinh hoạt, chia riêng 34 mẫu nước uống, 3 mẫu nước thải gộp chung 1 chứng thư.
  - 17/11/2026: mỗi mẫu một chứng thư -> nước thải NT1, NT2, NT3 = /39, /40, /41; quy chuẩn "khoan hãy so" (giữ ………)
    — đã thay bằng quyết định 18/11/2026 dưới.
  - Chỉ tiêu nước thải, nước sinh hoạt, nước uống: để trống (tô vàng).

Quyết định của người dùng (18/11/2026): "so kết quả nước thải theo QCVN 40:2011/BTNMT và QCVN 12:2015/BTNMT, cột A,
(Kq=1,2 và Kf=0,9) so cả 2 quy chuẩn. Nước sinh hoạt thì theo QCVN 01-1:2024/BYT. nước uống thì theo QCVN 6-1:2010/BYT".
  - Nước thải: 2 cột mức, Cmax = C × Kq × Kf = C × 1,08 (QCVN 40 mục 2.1.2: nhiệt độ, màu, pH, coliform, α, β lấy
    Cmax = C; QCVN 12 mục 2.1.2: nhiệt độ, pH lấy Cmax = C; QCVN 12 COD / độ màu theo cơ sở mới — áp cho mọi cơ sở từ
    01/01/2020). Giá trị C: QCVN 40:2011 bảng 1 cột A (bản tiếng Anh trên Z, bản tiếng Việt là ảnh scan); QCVN 12:2015
    bảng 1 cột A (Z `TIÊU CHUẨN VIỆT NAM. Tong hop cac quy chuan moi truong`).
  - Nước sinh hoạt: QCVN 01-1:2024/BYT (mức như khung NSH 134486/145103; TDS ≤ 1000 mg/L — AI theo QCVN, cần soát).
  - Nước uống: QCVN 06-1:2010/BYT phụ lục II, III (mức như 131229).

Suy luận của AI (cần người dùng xác nhận):
  1. Số chứng thư: NU1 – NU34 = /1 – /34 (theo danh sách KH), NSH1A, NSH2A, NSH1B, NSH2B = /35 – /38, nước thải = /39
     (tô vàng /35 – /39).
  2. Số đo hiện trường = trung bình 3 lần đo (pH 2 số lẻ, độ đục 2 số lẻ, TDS số nguyên); clo dư dưới 0,1 -> "< 0,1".
  3. Ngày yêu cầu cấp KQTN tạm = 12/11 + 8 = 20/11/2026 (chưa có PYC thử nghiệm).
  4. Ảnh NSH1.jpg dùng cho NSH1A và NSH1B, NSH2.jpg cho NSH2A và NSH2B.
  5. Máy đo độ đục Hanna HI 93703 thêm vào danh mục (`hanna_hi93703`).

Chạy:  python build_141523.py [--pdf] [--only 1,35,39]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER, FIELD

IN = os.path.join(WORKSPACE, "Input", "141523")
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k}.jpg")   # noqa: E731
P = PLACEHOLDER
BLANK = [(P, P, P)]                                          # chỉ tiêu thử nghiệm để trống (người dùng chốt)
SITE = "Cụm Công nghiệp tập trung Phú Hữu A - giai đoạn 1, xã Châu Thành, Thành phố Cần Thơ"
BYT = dict(procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.07", chi_tieu=P)
from build_131229 import LIMITS as NU_LIMITS                 # QCVN 06-1:2010/BYT phụ lục II, III
QC06, QC01 = "QCVN 06-1:2010/BYT", "QCVN 01-1:2024/BYT"
NU_STD = dict(standard_vn=f"{QC06} - Quy chuẩn kỹ thuật quốc gia đối với nước khoáng thiên nhiên và nước uống đóng chai",
              limits=NU_LIMITS, limit_head="Giới hạn tối đa ^(1)", word="chỉ tiêu",
              limit_note=f"Mức giới hạn tối đa cho phép quy định tại phụ lục II và phụ lục III, {QC06}.",
              eval_words=("Chỉ tiêu", "không đạt", "đạt"),
              intro=f"So sánh kết quả thử nghiệm các chỉ tiêu của mẫu nước uống nêu trên so với mức quy định tương ứng "
                    f"tại {QC06}, phụ lục II và phụ lục III:")
NSH_LIMITS = {"pH": "Trong khoảng\n(6,0 – 8,5)", "Clo dư": "Trong khoảng\n(0,2 – 1,0)", "Độ đục": "2",
              "Mùi": "Không có mùi lạ", "TDS": "1000", "Độ màu": "15", "Asen (As)": "0,01", "Chỉ số permanganate": "2",
              "Amoni (NH₃ và NH₄⁺ tính theo N)": "1", "Coliforms": "< 1", "Escherichia coli": "< 1"}
NSH_STD = dict(standard_vn=f"{QC01} - Quy chuẩn kỹ thuật quốc gia về chất lượng nước sạch sử dụng cho mục đích sinh "
                           f"hoạt", limits=NSH_LIMITS, limit_head="Ngưỡng giới hạn cho phép ^(1)",
               limit_note=f"Ngưỡng giới hạn tối đa cho phép quy định tại {QC01}.",
               intro=f"So sánh kết quả thử nghiệm các thông số của mẫu nước nêu trên so với mức quy định tương ứng tại "
                     f"{QC01}:")


def _vn(x):
    return f"{x:.6f}".rstrip("0").rstrip(".").replace(".", ",")


KQ, KF = 1.2, 0.9


def cmax(c, exempt=False):
    """Cmax = C × Kq × Kf (cột A, Kq = 1,2; Kf = 0,9); thông số miễn hệ số -> Cmax = C."""
    return _vn(c) if exempt else _vn(round(c * KQ * KF, 6))


Q40 = {  # QCVN 40:2011/BTNMT bảng 1 cột A — (C, miễn hệ số?)
    "Nhiệt độ": (40, True), "Độ màu": (50, True), "BOD₅": (30, False), "COD": (75, False), "TSS": (50, False),
    "Asen (As)": (0.05, False), "Thủy ngân (Hg)": (0.005, False), "Chì (Pb)": (0.1, False), "Cadmi (Cd)": (0.05, False),
    "Crom VI (Cr⁶⁺)": (0.05, False), "Crom III (Cr³⁺)": (0.2, False), "Đồng (Cu)": (2, False), "Kẽm (Zn)": (3, False),
    "Niken (Ni)": (0.2, False), "Mangan (Mn)": (0.5, False), "Sắt (Fe)": (1, False), "Xyanua (CN⁻)": (0.07, False),
    "Tổng phenol": (0.1, False), "Dầu mỡ khoáng": (5, False), "Sunfua (S²⁻)": (0.2, False), "Florua (F⁻)": (5, False),
    "Amoni (NH₄⁺ tính theo N)": (5, False), "Tổng nitơ": (20, False), "Tổng phospho": (4, False),
    "Clorua (Cl⁻)": (500, False), "Clo dư": (1, False), "Hóa chất BVTV clo hữu cơ": (0.05, False),
    "Hóa chất BVTV phospho hữu cơ": (0.3, False), "Polychlorinated biphenyl (PCB)": (0.003, False),
    "Coliform": (3000, True), "Tổng hoạt độ phóng xạ α": (0.1, True), "Tổng hoạt độ phóng xạ β": (1.0, True)}
Q12 = {  # QCVN 12:2015/BTNMT bảng 1 cột A (COD, độ màu: cơ sở mới)
    "Nhiệt độ": (40, True), "BOD₅": (30, False), "COD": (75, False), "TSS": (50, False), "Độ màu": (50, False),
    "Halogen hữu cơ dễ bị hấp thụ (AOX)": (7.5, False), "Dioxin": (15, False)}
NT_LIMITS = dict({k: cmax(c, e) for k, (c, e) in Q40.items()}, pH="6 – 9")
NT_LIMITS2 = dict({k: cmax(c, e) for k, (c, e) in Q12.items()}, pH="6 – 9")
NT_STD = dict(standard_vn="- QCVN 40:2011/BTNMT - Quy chuẩn kỹ thuật quốc gia về nước thải công nghiệp;\n"
                          "- QCVN 12:2015/BTNMT - Quy chuẩn kỹ thuật quốc gia về nước thải công nghiệp giấy và bột giấy.",
              limits=NT_LIMITS, limits2=NT_LIMITS2,
              limit_head="QCVN 40:2011/BTNMT cột A ^(1)", limit_head2="QCVN 12:2015/BTNMT cột A ^(2)",
              limit_note="Giá trị tối đa cho phép Cmax = C × Kq × Kf theo QCVN 40:2011/BTNMT, cột A (Kq = 1,2; Kf = 0,9); "
                         "nhiệt độ, màu, pH, coliform, tổng hoạt độ phóng xạ α, β áp dụng Cmax = C.",
              limit_note2="Giá trị tối đa cho phép Cmax = C × Kq × Kf theo QCVN 12:2015/BTNMT, cột A (Kq = 1,2; Kf = 0,9); "
                          "nhiệt độ, pH áp dụng Cmax = C.",
              intro="So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên so với giá trị tối đa cho phép "
                    "tương ứng tại QCVN 40:2011/BTNMT, cột A (Kq = 1,2; Kf = 0,9):",
              intro2="So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên so với giá trị tối đa cho phép "
                     "tương ứng tại QCVN 12:2015/BTNMT, cột A (Kq = 1,2; Kf = 0,9):")


def nu(n):
    return dict(BYT, **NU_STD, kind="nu", ky_hieu=f"NU{n}", kqtn_key=f"NU{n}", cert_no=n, time=P,
                vi_tri=f"Máy lọc nước - mã số: VNLM202510-{n:02d}", scope_vn="Lấy mẫu và giám định chất lượng nước uống",
                params=BLANK, field={}, equip=[], photo=PH(f"NU{n}"))


def nsh(key, n, vi_tri, field, tb, equip, photo):
    params = [(k, {"pH": "-", "Mùi": "-", "Độ đục": "NTU", "Clo dư": "mg/L", "TDS": "mg/L"}[k],
               "Cảm quan" if k == "Mùi" else FIELD) for k in tb] + BLANK
    return dict(BYT, **NSH_STD, kind="nsh", ky_hieu=key, kqtn_key=key, cert_no=n, time=P, vi_tri=vi_tri,
                scope_vn="Lấy mẫu và giám định chất lượng nước sinh hoạt", params=params, field=field,
                field_tbkq=tb, equip=equip, photo=PH(photo), hl_cert=True)


A = ["mettler_s2", "hanna_hi93703", "hach_dr300"]
NT_MEMBERS = [
    dict(ky_hieu="NT1", vi_tri="Nước thải trước hồ sinh học", time="10:15 – 10:30", photo=PH("NT1"),
         field={"pH": "7,80", "Nhiệt độ": "33,4", "Clo dư": "0,2"}),
    dict(ky_hieu="NT2", vi_tri="Nước thải sau hồ sinh học", time="10:40 – 10:50", photo=PH("NT2"),
         field={"pH": "7,84", "Nhiệt độ": "32,3", "Clo dư": "0,1"}),
    dict(ky_hieu="NT3", vi_tri="Nước thải hồ thu gom", time="11:00 – 11:15", photo=PH("NT3"),
         field={"pH": "7,30", "Nhiệt độ": "32,3", "Clo dư": "Không xác định"}),
]

def nt(m, n):
    """Mỗi mẫu nước thải một chứng thư (người dùng chốt 17/11/2026: "tách chứng thư mỗi cái trước"; căn cứ chưa so)."""
    return dict(kind="nt", ky_hieu=m["ky_hieu"], kqtn_key=m["ky_hieu"], cert_no=n, time=m["time"], vi_tri=m["vi_tri"],
                photo=m["photo"], field=m["field"], chi_tieu=P, **NT_STD,
                scope_vn="Lấy mẫu và giám định chất lượng nước thải",
                procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.06",
                params=[("Nhiệt độ", "°C", FIELD), ("pH", "-", FIELD), ("Clo dư", "mg/L", FIELD)] + BLANK,
                equip=["mettler_s2", "hach_dr300"], hl_cert=True)


CONFIG = dict(
    code6="141523", loai="nuoc", kind="nu",
    cert_date=__import__("datetime").date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2)
    client_vn="Công ty TNHH Giấy Lee & Man Việt Nam", address_vn=SITE,
    reg_date="05/11/2026", dates=["12/11/2026"], time_bckq="Ngày 12/11/2026 từ 09:30 đến 16:30",
    site_vn=SITE,
    scope_vn="Lấy mẫu và giám định chất lượng nước uống",
    scope_bckq="Lấy mẫu và giám định chất lượng nước uống, nước sinh hoạt, nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.07",
    procedure_bckq="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.06, N2.TC.N.07",
    equip=["mettler_s2", "hanna_hi93703", "hanna_hi98192", "hach_dr300"], executor="Quý",
    witness_vn="chị Huỳnh Thị Thoa – Tổ trưởng hóa nghiệm",
    req_date="20/11/2026",
    highlight=["< 0,1", "Không xác định"],
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2",
                   "Clo dư": "Hach DR300", "TDS": "Hanna HI 98192", "Mùi": "-"},
    bckq_notes=["- Độ đục (đo hiện trường, Hanna HI 93703): NSH1A 0,35 NTU; NSH2A 0,12 NTU (theo BBLM).",
                "- Số đo hiện trường ghi là trung bình 3 lần đo trên BBLM.",
                "- Chỉ tiêu thử nghiệm của các mẫu: để trống, chờ xác nhận."],
    samples=[nu(n) for n in range(1, 35)] + [
        nsh("NSH1A", 35, "Nước sinh hoạt 1 – Nhóm A",
            {"pH": "7,47", "Mùi": "Không có mùi lạ", "Độ đục": "0,35", "Clo dư": "< 0,1"},
            ["pH", "Mùi", "Độ đục", "Clo dư"], A, "NSH1"),
        nsh("NSH2A", 36, "Nước sinh hoạt 2 – Nhóm A",
            {"pH": "7,44", "Mùi": "Không có mùi lạ", "Độ đục": "0,12", "Clo dư": "< 0,1"},
            ["pH", "Mùi", "Độ đục", "Clo dư"], A, "NSH2"),
        nsh("NSH1B", 37, "Nước sinh hoạt 1 – Nhóm B", {"TDS": "84"}, ["TDS"], ["hanna_hi98192"], "NSH1"),
        nsh("NSH2B", 38, "Nước sinh hoạt 2 – Nhóm B", {"TDS": "80"}, ["TDS"], ["hanna_hi98192"], "NSH2"),
    ] + [nt(m, 39 + i) for i, m in enumerate(NT_MEMBERS)],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
