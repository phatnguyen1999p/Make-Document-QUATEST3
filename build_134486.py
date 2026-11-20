# -*- coding: utf-8 -*-
r"""
Q326D134486 — Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO — KCN Hữu Thạnh: nước sinh hoạt N1 – N3 (QCVN 01-1:2024/BYT)
và nước thải N4 – N11 (11 mẫu, 11 chứng thư). CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D134486 - IDICO Hữu Thạnh - Nước thải, Nước sinh hoạt` về Input\134486, chỉ đọc):
  - `0. BBLM - 1.pdf` (2 trang): ngày 29/10/2026, 09:15 – 14:15, PYC Q326D134486 ngày 23/10/2026; "Lấy mẫu, giám định chất
    lượng nước sinh hoạt theo QCVN 01-1:2024/BYT và nước thải theo QCVN 40:2011/BTNMT"; địa điểm "Khu công nghiệp Hữu Thạnh,
    xã Đức Hòa, tỉnh Tây Ninh"; Hach DR900 + Mettler Seven2Go S2; N2.TC.N.02, N2.TC.N.06; chỉ định: anh Võ Văn Chí Công –
    quản đốc; N1 – N6 (giờ lấy mẫu, đo nhanh pH, t°, clo dư, mùi); ghi chú: clo dư tự do của mẫu N5 thực hiện theo yêu cầu KH.
  - `0. BBLM - 2.pdf` (4 trang): 09:45 – 14:00, Mettler Seven2Go S2; N7 – N11 (pH, t°, giờ), N12, N13 "không thực hiện được".
  - `4. edoc\134486 - Danh sách vị trí + chỉ tiêu.pdf`; khung TBKQ làm tay `3. TBKQ - 1…11 - Q326D134486.docx` (danh sách
    thông số, phương pháp, mức NSH lấy từ đây — chỉ đọc).
  - Ảnh `0. vi tri\N1 … N11 - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026):
  - Quy trình nước sinh hoạt: để trống, tô vàng (BBLM ghi N2.TC.N.02).
  - Người thực hiện: Huy.

Suy luận của AI (cần người dùng xác nhận):
  1. Quy trình NSH theo BBLM `N2.TC.N.02` (AGENTS mặc định NSH `N2.TC.N.01`).
  2. Danh sách thông số / phương pháp lấy theo khung làm tay; khi có KQTN sẽ thay bằng tên / phương pháp trên phiếu.
  3. Clo dư tự do N5 (0,83 mg/L, KH yêu cầu thêm) chỉ ghi ở BCKQ; TBKQ /5 theo khung (clo dư 0,96).
  4. Người thực hiện "Huy + Việt" theo khung làm tay (2 BBLM chỉ có chữ ký Hồ Nguyễn Huy).
  5. Ngày yêu cầu cấp KQTN tạm = ngày lấy mẫu + 8 = 06/11/2026 (chưa có PYC thử nghiệm).

Quyết định của người dùng (20/11/2026): "quy trình thực hiện nước sinh hoạt là N2.TC.N.01 … sửa và xóa tô vàng"
  -> NSH ghi QUATEST 3 3030:2025, N2.TC.N.01 (thay ………/N2.TC.N.02 tô vàng), BCKQ N2.TC.N.01, N2.TC.N.06.

Chạy:  python build_134486.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE

IN = os.path.join(WORKSPACE, "Input", "134486")
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
DR = lambda n: os.path.join(IN, f"3. TBKQ - {n} - Q326D134486.docx")       # noqa: E731
KCN = "Khu công nghiệp Hữu Thạnh"
NSH_LIM = {"pH": "Trong khoảng\n(6,0 – 8,5)", "Clo dư tự do": "Trong khoảng\n(0,2 – 1,0)", "Màu sắc": "15", "Độ đục": "2",
           "Asen (As)": "0,01", "Chỉ số permanganat": "2", "Amoni (NH₃ và NH₄⁺ tính theo N)": "1",
           "Mùi": "Không có mùi lạ", "Coliform tổng số": "< 1", "Escherichia coli": "< 1"}
NSH = dict(kind="nsh", scope_vn="Lấy mẫu và giám định chất lượng nước sinh hoạt",
           procedure="QUATEST 3 3030:2025, N2.TC.N.01", equip=["mettler_s2", "hach_dr900"],
           standard_vn="QCVN 01-1:2024/BYT - Quy chuẩn kỹ thuật quốc gia về chất lượng nước sạch sử dụng cho mục đích "
                       "sinh hoạt",
           chi_tieu="10 thông số nhóm A nêu tại QCVN 01-1:2024/BYT", limits=NSH_LIM,
           limit_head="Ngưỡng giới hạn cho phép ^(1)",
           limit_note="Ngưỡng giới hạn tối đa cho phép quy định tại QCVN 01-1:2024/BYT.",
           intro="So sánh kết quả thử nghiệm các thông số của mẫu nước nêu trên so với mức quy định tương ứng tại "
                 "QCVN 01-1:2024/BYT:")


def nsh(n, time, vi_tri, ph, t, clo):
    return dict(NSH, ky_hieu=f"134486-N{n}", kqtn_key=f"N{n}", time=time, vi_tri=vi_tri, draft=DR(n), photo=PH(f"N{n}"),
                field={"pH": ph, "Nhiệt độ": t, "Clo dư": clo, "Clo dư tự do": clo, "Mùi": "Không có mùi lạ"})


def nt(n, time, vi_tri, ph, t, clo=None, equip=("mettler_s2",)):
    f = {"pH": ph, "Nhiệt độ": t}
    if clo:
        f["Clo dư"] = clo
    return dict(kind="nt", ky_hieu=f"134486-N{n}", kqtn_key=f"N{n}", time=time, vi_tri=vi_tri, draft=DR(n),
                photo=PH(f"N{n}"), field=f, equip=list(equip))


CONFIG = dict(
    code6="134486", loai="nuoc", kind="nt",
    client_vn="Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    address_vn="Số 48, khu phố 1, phường Long Hưng, Thành phố Đồng Nai",
    reg_date="23/10/2026", dates=["29/10/2026"], time_bckq="Ngày 29/10/2026 từ 09:15 đến 14:15",
    site_vn=f"{KCN}, xã Đức Hòa, tỉnh Tây Ninh",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    scope_bckq="Lấy mẫu và giám định chất lượng nước sinh hoạt, nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    procedure_bckq="QUATEST 3 3030:2025, N2.TC.N.01, N2.TC.N.06",
    equip=["mettler_s2", "hach_dr900"], executor="Huy",
    witness_vn="anh Võ Văn Chí Công – quản đốc, đại diện Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    req_date="06/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR900",
                   "Clo dư tự do": "Hach DR900", "Mùi": "-"},
    bckq_notes=["- Mẫu N12, N13: không thực hiện được (theo BBLM).",
                "- Clo dư tự do của mẫu N5 (0,83 mg/L) đo theo yêu cầu của khách hàng (theo BBLM)."],
    samples=[
        nsh(1, "10:25", f"Mẫu nước sinh hoạt: Đường N2, {KCN}", "6,88", "32,5", "0,56"),
        nsh(2, "10:45", f"Mẫu nước sinh hoạt: Đường D5, {KCN}", "6,68", "31,4", "0,37"),
        nsh(3, "10:35", f"Mẫu nước sinh hoạt: Đường D6, {KCN} (vị trí cấp nước Công ty TNHH Nước giải khát Suntory "
                        "PepsiCo Việt Nam - Chi nhánh Long An)", "6,87", "32,1", "0,67"),
        nt(4, "10:55", f"Nước thải đầu vào Nhà máy xử lý nước thải tập trung {KCN}", "6,87", "31,5", "0,00",
           ("mettler_s2", "hach_dr900")),
        nt(5, "09:55", f"Nước thải đầu ra Nhà máy xử lý nước thải tập trung {KCN}", "8,33", "31,5", "0,96",
           ("mettler_s2", "hach_dr900")),
        nt(6, "14:00", "Nước thải Công ty TNHH Nước giải khát Suntory PepsiCo Việt Nam - Chi nhánh Long An", "7,98", "33,6"),
        nt(7, "09:50", "Nước thải Công ty TNHH Khoa học Kỹ thuật Kingtec Việt Nam", "7,39", "31,7"),
        nt(8, "10:25", "Nước thải Công ty Cổ phần Bán lẻ Kỹ thuật số FPT", "6,83", "30,9"),
        nt(9, "10:40", "Nước thải Công ty TNHH New Jin Feng Việt Nam", "5,73", "31,7"),
        nt(10, "11:10", "Nước thải Công ty TNHH Suntone Industry Việt Nam", "6,55", "32,6"),
        nt(11, "13:30", "Nước thải Công ty TNHH TDCH Plastic & Hardware", "7,00", "31,7"),
    ],
)
CONFIG["samples"][4]["field"]["Clo dư tự do"] = "0,83"

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
