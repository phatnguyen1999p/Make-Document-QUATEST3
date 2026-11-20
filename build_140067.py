# -*- coding: utf-8 -*-
r"""
Q326D140067 — Công ty TNHH MTV Phát triển Đô thị và Khu công nghiệp IDICO — KCN Nhơn Trạch 1 — nước thải tháng 11
(10 mẫu, 10 chứng thư: N1 – N6, N8 – N11). CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D140067 - IDICO Nhơn trạch 1 - Nước thải  tháng 11` về Input\140067, chỉ đọc):
  - `0. BBLM - 1.pdf` (2 trang, Hồ Nguyễn Huy): ngày 06/11/2026, PYC Q326D140067 ngày 03/11/2026, 09:00 – 11:30; "nước thải
    theo QCVN 40:2011/BTNMT"; địa điểm "KCN Nhơn Trạch 1, phường Nhơn Trạch, thành phố Đồng Nai"; Hach DR900 + Mettler;
    N2.TC.N.06; chỉ định: chị Dương Thị Thanh Hoa – quản đốc; N1 pH 7,29 / 32,8 °C, clo dư 0,01 (chữ viết tay, chưa chắc);
    N2 pH 7,52 / 33,8; N3 09:15 pH 6,38 / 29,8; N4 10:05 pH 8,44 / 37,8; N5 09:45 pH 8,48 / 32,9; N6 09:30 pH 7,43 / 33,0;
    ghi chú: mẫu N2 chỉ tiêu clo dư và clo dư tự do không thực hiện theo yêu cầu khách hàng.
  - `0. BBLM - 2.pdf` (2 trang, Nguyễn Hoàng Việt): 09:00 – 11:15, Mettler; N8 10:04 pH 8,59 / 33,4; N9 10:37 pH 7,65 / 31,1;
    N10 11:02 pH 7,73 / 32,3; N11 09:40 pH 6,28 / 29,0; N7 (Posco VST), N12 (Tung Kuang): hố ga đấu nối không có nước
    chảy ra, không lấy mẫu.
  - `PYC-Q326D140067.N2.1.pdf` (MTTN26021030): chỉ tiêu từng mẫu (tên mẫu N1 – N11 trên PYC ghi "135226 - Nn"), dự kiến
    trả 20/11/2026; QC1, QC2.
  - Ảnh `0. vi tri\Nn - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026): số chứng thư theo số mẫu (N9 -> /9), tô vàng chỗ nhảy cóc; clo dư N1 = 0,01 (-> "< 0,1", tô vàng); tô vàng chỗ phiếu yêu cầu thử nghiệm ghi tên mẫu "135226 - Nn".

Suy luận của AI (cần người dùng xác nhận):
  1. Chỉ tiêu theo PYC thử nghiệm; tên / phương pháp sẽ thay theo KQTN khi có.
  2. Clo dư N1 đọc "0,01" (< 0,1) -> "< 0,1"; giờ lấy mẫu N1, N2 không ghi -> "………".
  3. Không có cột / hệ số QCVN 40:2011 -> không đánh giá.
  5. PYC ghi tên mẫu "135226 - Nn" (mã hồ sơ tháng trước) — KQTN có thể ghi tên mẫu như vậy; khóa tra KQTN theo đuôi Nn.

Chạy:  python build_140067.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER, clo, pyc_index, pyc_sample

IN = os.path.join(WORKSPACE, "Input", "140067")
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
PYC, _ = pyc_index(os.path.join(IN, "PYC-Q326D140067.N2.1.pdf"))
TT = ("Nhiệt độ", "pH", "Clo dư")


def s(key, time, vi_tri, ph, t, tbkq=("pH",), equip=("mettler_s2",), photo=None, **extra):
    f = {"pH": ph, "Nhiệt độ": t}
    f.update(extra)
    return pyc_sample("140067", key, PYC, time, vi_tri, f, tbkq, equip, photo=photo or PH(key))


CONFIG = dict(
    code6="140067", loai="nuoc thai", kind="nt",
    cert_by_key=True,                    # số chứng thư theo số mẫu (Nn -> /n), tô vàng khi nhảy cóc
    client_vn="Công ty TNHH MTV Phát triển Đô thị và Khu công nghiệp IDICO",
    address_vn="Đường Tôn Đức Thắng, Khu công nghiệp Nhơn Trạch 1, phường Nhơn Trạch, Thành phố Đồng Nai",
    reg_date="03/11/2026", dates=["06/11/2026"], time_bckq="Ngày 06/11/2026 từ 09:00 đến 11:30",
    site_vn="Khu công nghiệp Nhơn Trạch 1, phường Nhơn Trạch, Thành phố Đồng Nai",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    equip=["mettler_s2", "hach_dr900"], executor="Huy + Việt",
    witness_vn="chị Dương Thị Thanh Hoa – quản đốc",
    req_date="20/11/2026",
    highlight=["< 0,1", "135226 - N1 … N11"],
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR900"},
    bckq_notes=["- Mẫu N7 (Công ty TNHH Posco VST) và N12 (Công ty CP Công nghiệp Tung Kuang): tại thời điểm lấy mẫu hố ga "
                "đấu nối không có nước chảy ra, không lấy mẫu (theo BBLM).",
                "- Mẫu N2: chỉ tiêu clo dư và clo dư tự do không thực hiện theo yêu cầu khách hàng (theo BBLM).",
                "- Phiếu yêu cầu thử nghiệm MTTN26021030 ghi tên mẫu \"135226 - N1 … N11\" (mã hồ sơ tháng 10)."],
    samples=[
        s("N1", PLACEHOLDER, "Nước thải đầu vào Nhà máy xử lý nước thải Khu công nghiệp Nhơn Trạch 1", "7,29", "32,8",
          TT, ("mettler_s2", "hach_dr900"), **{"Clo dư": clo("0,01")}),
        s("N2", PLACEHOLDER, "Nước thải đầu ra tại mương đo lưu lượng sau hệ thống xử lý nước thải", "7,52", "33,8",
          ("Nhiệt độ", "pH")),
        s("N3", "09:15", "Nước thải Công ty TNHH Công nghiệp T.A Việt Nam", "6,38", "29,8"),
        s("N4", "10:05", "Nước thải Công ty TNHH IL Kwang Vina", "8,44", "37,8"),
        s("N5", "09:45", "Nước thải Công ty TNHH SX Thun và NL May Premier VN", "8,48", "32,9"),
        s("N6", "09:30", "Nước thải Công ty TNHH White Glove", "7,43", "33,0"),
        s("N8", "10:04", "Nước thải Công ty TNHH Thực phẩm NFC", "8,59", "33,4"),
        s("N9", "10:37", "Nước thải Công ty CP TBVS Caesar VN", "7,65", "31,1"),
        s("N10", "11:02", "Nước thải Công ty TNHH Tea Yang", "7,73", "32,3"),
        s("N11", "09:40", "Nước thải Công ty TNHH MTV Austdoor Nhơn Trạch", "6,28", "29,0",
          photo=os.path.join(IN, "0. vi tri", "N11 -  Vị trí lấy mãu.jpg")),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
