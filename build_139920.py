# -*- coding: utf-8 -*-
r"""
Q326D139920 — Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO — KCN Hựu Thạnh — tháng 11: nước sinh hoạt N1 – N3, N14
(QCVN 01-1:2024/BYT) và nước thải N4 – N13 (14 mẫu, 14 chứng thư /n = Nn). CHƯA CÓ KQTN -> bản khung, "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D139920 - IDICO Hữu Thạnh - Nước thải, Nước sinh hoạt - thang 11` về Input\139920):
  - `0. BBLM Nhóm Huy.pdf` (2 trang): ngày 10/11/2026, PYC Q326D139920 ngày 03/11/2026, 09:30 – 14:00; "nước sinh hoạt theo
    QCVN 01-1:2024/BYT và nước thải theo QCVN 40:2011/BTNMT"; địa điểm "Khu công nghiệp Hựu Thạnh, xã Đức Hòa, tỉnh Tây
    Ninh"; Hach DR900 + Mettler; N2.TC.N.02, N2.TC.N.06; chỉ định: anh Võ Văn Chí Công – quản đốc; N1 09:40, N2 10:05,
    N3 10:20, N14 09:50 (nước sinh hoạt: pH, t°, clo dư, mùi); N4 11:20, N5 11:05 (+ clo dư tự do 0,38), N6 13:50 (nước
    thải: pH, t°, clo dư); ghi chú: clo dư tự do N5 theo yêu cầu KH; mẫu N6 không thực hiện chỉ tiêu Coliform.
  - `0. BBLM - NHÓM VIỆT.pdf` (3 trang): 09:30 – 12:00, Mettler; N7 – N13 (pH, t°, giờ).
  - `PYC-Q326D139920.N2.1.pdf` (MTTN26021087): chỉ tiêu N1 – N13 (lấy từ đây), dự kiến trả 20/11/2026; QC1, QC2. PYC
    KHÔNG có mẫu N14.
  - Ảnh `0. vi tri\Nn - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026): quy trình NSH để trống (tô vàng); clo dư < 0,1 -> "< 0,1" (tô vàng);
  không có mẫu N14 (bỏ chứng thư /14); N6 bỏ Coliform — đồng ý.

Suy luận của AI (cần người dùng xác nhận):
  1. Nước sinh hoạt: 10 thông số nhóm A, phương pháp, mức theo khung 134486/1; quy trình để "………" tô vàng như 134486.
  2. N14 (nước sạch tại bể chứa PCCC) có trên BBLM nhưng không có trên PYC thử nghiệm -> vẫn dựng chứng thư /14 (khung NSH).
  3. N6 bỏ Coliforms (BBLM: không thực hiện) dù PYC có.
  4. Clo dư đọc dưới 0,1 mg/L (N4 0,08; N6 0,09; N14 0,00) -> "< 0,1".
  5. Nước thải: chỉ tiêu theo PYC, không đánh giá (không có cột / hệ số QCVN 40:2011).

Chạy:  python build_139920.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER, clo, nsh_base, pyc_index, pyc_sample

IN = os.path.join(WORKSPACE, "Input", "139920")
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
PYC, _ = pyc_index(os.path.join(IN, "PYC-Q326D139920.N2.1.pdf"))
PYC["N6"] = [p for p in PYC["N6"] if not p[0].lower().startswith("coliform")]
KCN = "KCN Hựu Thạnh"
NSHP = "QUATEST 3 3030:2025, " + PLACEHOLDER


def nsh(n, time, vi_tri, ph, t, c):
    return dict(nsh_base(), ky_hieu=f"139920-N{n}", kqtn_key=f"N{n}", time=time, vi_tri=vi_tri, procedure=NSHP,
                equip=["mettler_s2", "hach_dr900"], photo=PH(f"N{n}"),
                field={"pH": ph, "Nhiệt độ": t, "Clo dư tự do": clo(c), "Clo dư": clo(c), "Mùi": "Không có mùi lạ"})


def nt(n, time, vi_tri, ph, t, tbkq=("pH",), equip=("mettler_s2",), **extra):
    f = {"pH": ph, "Nhiệt độ": t}
    f.update(extra)
    return pyc_sample("139920", f"N{n}", PYC, time, vi_tri, f, tbkq, equip, photo=PH(f"N{n}"), kind="nt")


DR = ("mettler_s2", "hach_dr900")
CONFIG = dict(
    code6="139920", loai="nuoc", kind="nt",
    client_vn="Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    address_vn="Số 48, khu phố 1, phường Long Hưng, Thành phố Đồng Nai",
    reg_date="03/11/2026", dates=["10/11/2026"], time_bckq="Ngày 10/11/2026 từ 09:30 đến 14:00",
    site_vn="Khu công nghiệp Hựu Thạnh, xã Đức Hòa, tỉnh Tây Ninh",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    scope_bckq="Lấy mẫu và giám định chất lượng nước sinh hoạt, nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    procedure_bckq="QUATEST 3 3030:2025, " + PLACEHOLDER + ", N2.TC.N.06",
    equip=["mettler_s2", "hach_dr900"], executor="Huy + Việt",
    witness_vn="anh Võ Văn Chí Công – quản đốc nhà máy",
    req_date="20/11/2026",
    highlight=["< 0,1"],
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR900",
                   "Clo dư tự do": "Hach DR900", "Mùi": "-"},
    bckq_notes=["- Clo dư tự do của mẫu N5 đo theo yêu cầu của khách hàng; mẫu N6 không thực hiện chỉ tiêu Coliform "
                "(theo BBLM)."],
    samples=[
        nsh(1, "09:40", f"Mẫu nước sinh hoạt: Đường N2, {KCN}", "6,89", "31,8", "0,47"),
        nsh(2, "10:05", f"Mẫu nước sinh hoạt: Đường D5, {KCN}", "6,83", "31,2", "0,45"),
        nsh(3, "10:20", f"Mẫu nước sinh hoạt: Đường D6, {KCN} (vị trí cấp nước Công ty TNHH Nước giải khát Suntory "
                        "PepsiCo Việt Nam - CN LA)", "6,93", "31,9", "0,79"),
        nt(4, "11:20", f"Mẫu nước thải đầu vào Nhà máy XLNT TT {KCN}", "7,48", "32,8", ("Nhiệt độ", "pH", "Clo dư"), DR,
           **{"Clo dư": clo("0,08")}),
        nt(5, "11:05", f"Mẫu nước thải đầu ra Nhà máy XLNT TT {KCN}", "7,98", "31,7", ("Nhiệt độ", "pH", "Clo dư"), DR,
           **{"Clo dư": clo("0,46"), "Clo dư tự do": clo("0,38")}),
        nt(6, "13:50", "Mẫu nước thải Công ty TNHH Nước giải khát Suntory PepsiCo Việt Nam - CN LA", "7,49", "32,2",
           ("pH", "Clo dư"), DR, **{"Clo dư": clo("0,09")}),
        nt(7, "09:55", "Mẫu nước thải Công ty TNHH Greencans Industries (Việt Nam)", "8,86", "32,6"),
        nt(8, "10:10", "Mẫu nước thải Công ty Cổ phần MTK Hữu Thạnh", "9,19", "31,1"),
        nt(9, "10:35", "Mẫu nước thải Công ty TNHH HD Long An", "8,42", "30,7"),
        nt(10, "10:50", "Mẫu nước thải Công ty TNHH TDCH Plastic & Hardware", "7,32", "31,7"),
        nt(11, "11:10", "Mẫu nước thải Công ty TNHH Vật liệu mới Quan Feng", "6,82", "30,9"),
        nt(12, "11:25", "Mẫu nước thải Công ty Cổ phần Ống thép Cường Phát Sài Gòn", "5,93", "31,9"),
        nt(13, "11:45", "Mẫu nước thải Công ty TNHH New Jin Feng Việt Nam", "6,11", "30,8"),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
