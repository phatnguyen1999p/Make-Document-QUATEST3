# -*- coding: utf-8 -*-
r"""
Q326D138952 — Công ty TNHH MTV Phát triển Đô thị và Khu công nghiệp IDICO — nước dưới đất tháng 11 (QCVN 09:2023/BTNMT),
32 giếng: N1 – N18 (02/11/2026) và M1 – M14 (03/11/2026), 32 chứng thư (/1 – /18 = N1 – N18, /19 – /32 = M1 – M14).
CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D138952 - IDICO Tuy Hạ - Nước dưới đất tháng 11` về Input\138952, chỉ đọc):
  - `0. BBLM\0. BBLM nhóm 1 ngày 02.11.pdf` (Nguyễn Hoàng Việt): PYC Q326D138952 ngày 30/10/2026, 08:30 – 11:20; "Lấy mẫu,
    giám định chất lượng nước dưới đất theo QCVN 09:2023/BTNMT"; địa điểm (viết tay) "Khu công nghiệp Nhơn Trạch 1,
    phường Nhơn Trạch, thành phố Đồng Nai"; Mettler Seven2Go S2; N2.TC.N.05; chỉ định: chị Bùi Thị Thủy – nhân viên;
    N1 – N9 (giếng H1, DK, 453A, H3A, H2A, QT1, B2, 452, LK5A; giờ, pH, t°).
  - `0. BBLM\0. 138952 - BBLM nhóm 2 ngày 02.11.pdf` (Hồ Nguyễn Huy): 08:30 – 11:00, KCN Nhơn Trạch 1; N10 – N18 (giếng
    LK1A, C8, C5, C6, C4, LK4, C3, H4A, LK2A).
  - `0. BBLM\0. 138952 - BBLM nhóm 1 ngày 03.11.pdf` (Việt): 09:00 – 10:30, địa điểm "Khu công nghiệp Nhơn Trạch 5"; M1 – M7
    (giếng KT9, KT8, KT7, KT6, KT5, QS4, KT4).
  - `0. BBLM\0. BBLM nhóm 2 ngày 03.11.pdf` (Huy): 09:00 – 10:30, KCN Nhơn Trạch 5; M8 KT14, M9 KT15, M10 KT16, M11 KT17,
    M12 KT2, "M14" KT18 (10:00), "M14" KT3 (10:10) — ký hiệu M14 ghi 2 lần; ghi chú: trời nắng, nước các giếng trong,
    không màu, miệng giếng có nắp đậy.
  - `0. BBLM\0. PYC-Q326D138952.N2.1.pdf` (MTTN26020754, N1 – N18, dự kiến trả 14/11/2026) và `.N2.2.pdf` (MTTN26020850,
    M1 – M14, dự kiến trả 16/11/2026); giấy nhận mẫu KH `BM-QLM 03 … TUY HA THANG 11.docx` ("41 chỉ tiêu theo QCVN
    09:2023/BTNMT"; tên giếng "45-2", "45-3A").
  - Khung chứng thư nước dưới đất của người dùng `3. TBKQ - 1.docx` (hồ sơ Q326D100302 — 41 thông số, phương pháp, mức).
  - Ảnh `Q326D138952 - Photo\…`.

Quyết định của người dùng (16/11/2026):
  - 41 chỉ tiêu (giữ); BBLM "ghi đúng hết", khoan sửa — chỉ tô vàng (KT18 = M13 giữ, tô vàng).
  - Tên giếng ghi theo khách: 45-3A, 45-2. Giữ nguyên địa điểm như BBLM, không sửa chữ "Tuy Hạ".
  - Ảnh chưa đặt tên thì để trống, không thêm; ghi chú "ảnh chưa đặt tên nên chưa thêm".

Suy luận của AI (cần người dùng xác nhận):
  1. 41 thông số, phương pháp, mức QCVN 09:2023/BTNMT theo khung Q326D100302 (tô vàng cột mức); TDS theo PYC (phòng thí
     nghiệm, SMEWW 2540C:2023) thay vì đo hiện trường.
  2. Giếng KT18 = M13 (BBLM ghi "M14" hai lần; ảnh có M13) — tô vàng.
  4. Địa điểm theo BBLM: ngày 02/11 KCN Nhơn Trạch 1, ngày 03/11 KCN Nhơn Trạch 5 (thư mục Z ghi "Tuy Hạ").

Chạy:  python build_138952.py [--pdf] [--only 1,19]
"""
import argparse
import glob
import os
from nuoc_common import run, WORKSPACE, params_from_draft, limits_from_draft

IN = os.path.join(WORKSPACE, "Input", "138952")
PHD = os.path.join(IN, "Q326D138952 - Photo")
DR = os.path.join(IN, "3. TBKQ - 1.docx")
STD = "QCVN 09:2023/BTNMT - Quy chuẩn kỹ thuật quốc gia về chất lượng nước dưới đất"
NT1 = "Khu công nghiệp Nhơn Trạch 1, phường Nhơn Trạch, Thành phố Đồng Nai"
NT5 = "Khu công nghiệp Nhơn Trạch 5, phường Nhơn Trạch, Thành phố Đồng Nai"
PARAMS = [(n, u, "SMEWW 2540C:2023" if "TDS" in n else m) for n, u, m in params_from_draft(DR)]
LIMITS = limits_from_draft(DR)


def photo(key, group, well):
    # chỉ ảnh đã đặt tên "… - Vị trí lấy mẫu"; ảnh chưa đặt tên (nhóm Việt) không thêm (người dùng chốt 16/11/2026)
    f = glob.glob(os.path.join(PHD, group, f"{key} - Vị trí lấy m*.jpg"))
    return f[0] if f else None


def g(key, well, time, ph, t, date, group, site):
    return dict(ky_hieu=f"138952-{key}", kqtn_key=key, time=time, vi_tri=f"Giếng {well}", date=date, site_vn=site,
                params=PARAMS, limits=LIMITS, hl_limits=True, standard_vn=STD,
                limit_head="Giá trị giới hạn ^(1)", limit_note="Giá trị giới hạn tối đa nêu tại QCVN 09:2023/BTNMT.",
                intro="So sánh kết quả thử nghiệm các thông số của mẫu nước dưới đất nêu trên với mức quy định tương ứng "
                      "nêu tại QCVN 09:2023/BTNMT:",
                chi_tieu="41 thông số tại bảng 1, QCVN 09:2023/BTNMT",
                field={"pH": ph, "Nhiệt độ": t}, photo=photo(key, group, well),
                photo_note="ảnh chưa đặt tên nên chưa thêm")


D1, D2 = "02/11/2026", "03/11/2026"
V1, H1, V2, H2 = "Nhóm Việt - ngày 02.11", "Nhóm Huy ngày 02.11", "Nhóm Việt - ngày 03.11", "Nhóm Huy ngày 03.11"
N = [("N1", "H1", "08:50", "4,69", "30,2"), ("N2", "DK", "09:05", "4,66", "30,3"), ("N3", "45-3A", "09:20", "5,13", "30,1"),
     ("N4", "H3A", "09:35", "4,57", "30,0"), ("N5", "H2A", "09:50", "4,51", "30,0"), ("N6", "QT1", "10:05", "5,58", "29,7"),
     ("N7", "B2", "10:20", "4,33", "30,1"), ("N8", "45-2", "10:30", "4,31", "30,2"), ("N9", "LK5A", "10:40", "4,70", "29,9")]
NH = [("N10", "LK1A", "10:25", "4,28", "30,4"), ("N11", "C8", "08:55", "4,37", "29,9"), ("N12", "C5", "09:15", "4,71", "29,8"),
      ("N13", "C6", "09:25", "4,26", "30,3"), ("N14", "C4", "09:30", "4,27", "30,0"), ("N15", "LK4", "09:40", "4,44", "30,0"),
      ("N16", "C3", "09:50", "4,70", "30,1"), ("N17", "H4A", "10:10", "4,49", "30,3"), ("N18", "LK2A", "10:20", "4,13", "30,4")]
M = [("M1", "KT9", "09:05", "4,19", "29,6"), ("M2", "KT8", "09:15", "3,43", "29,7"), ("M3", "KT7", "09:23", "4,42", "29,9"),
     ("M4", "KT6", "09:30", "4,52", "29,7"), ("M5", "KT5", "09:40", "4,12", "30,1"), ("M6", "QS4", "09:45", "6,80", "30,3"),
     ("M7", "KT4", "10:15", "4,55", "30,6")]
MH = [("M8", "KT14", "09:10", "3,76", "30,4"), ("M9", "KT15", "09:20", "5,06", "29,9"), ("M10", "KT16", "09:30", "5,35", "30,1"),
      ("M11", "KT17", "09:40", "5,24", "30,0"), ("M12", "KT2", "09:50", "5,12", "30,6"), ("M13", "KT18", "10:00", "4,57", "30,6"),
      ("M14", "KT3", "10:10", "4,28", "29,8")]

CONFIG = dict(
    code6="138952", loai="nuoc duoi dat", kind="nt",
    client_vn="Công ty TNHH MTV Phát triển Đô thị và Khu công nghiệp IDICO",
    address_vn="Đường Tôn Đức Thắng, Khu công nghiệp Nhơn Trạch 1, phường Nhơn Trạch, Thành phố Đồng Nai",
    reg_date="30/10/2026", dates=[D1, D2],
    time_bckq="Ngày 02/11/2026 từ 08:30 đến 11:20; ngày 03/11/2026 từ 09:00 đến 10:30",
    site_vn=NT1, site_bckq=f"{NT1}; {NT5}",
    scope_vn="Lấy mẫu và giám định chất lượng nước dưới đất",
    procedure="QUATEST 3 3030:2025, N2.TC.N.05",
    equip=["mettler_s2"], executor="Huy + Việt",
    witness_vn="chị Bùi Thị Thủy – nhân viên",
    req_date="16/11/2026",
    highlight=["Giếng KT18", "ảnh chưa đặt tên nên chưa thêm"],
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2"},
    bckq_notes=["- Hiện trạng ngày 03/11/2026: trời nắng, nước của các giếng trong, không màu, miệng giếng có nắp đậy "
                "(theo BBLM)."],
    samples=[g(*x, D1, V1, NT1) for x in N] + [g(*x, D1, H1, NT1) for x in NH]
            + [g(*x, D2, V2, NT5) for x in M] + [g(*x, D2, H2, NT5) for x in MH],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
