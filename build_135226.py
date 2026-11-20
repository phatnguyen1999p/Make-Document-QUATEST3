# -*- coding: utf-8 -*-
r"""
Q326D135226 — Công ty TNHH MTV Phát triển Đô thị và Khu công nghiệp IDICO — nước thải KCN Nhơn Trạch 1 (11 mẫu N1 – N11,
11 chứng thư /1 – /11). ĐÃ CÓ KQTN MTTN26020080.01 – .11 (+ .12 mẫu QC), ngày 17/11/2026.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D135226 - IDICO Nhơn trạch 1 - Nước thải` về Input\135226, chỉ đọc):
  - `0. BBLM.pdf` (3 trang + 1 trang trắng): ngày 27/10/2026, PYC Q326D135226 ngày 23/10/2026; "CÔNG TY TNHH MTV PHÁT TRIỂN
    ĐÔ THỊ VÀ KCN IDICO", "Đường Tôn Đức Thắng, KCN Nhơn Trạch 1, phường Nhơn Trạch, Tp. Đồng Nai"; "Lấy mẫu, giám định chất
    lượng nước thải theo QCVN 40:2011/BTNMT"; địa điểm viết tay "Khu xử lý nước thải KCN Nhơn Trạch 1, phường Nhơn Trạch,
    TP. Đồng Nai"; 08h45 – 15h30; Hach DR900 (clo dư), Mettler Seven2Go S2; QTGĐ 3030.2025, N2.TC.N.06; chỉ định: chị
    Dương Thị Thanh Hoa – quản đốc; 13 dòng, N12 (Austdoor Nhơn Trạch), N13 (Tung Kuang) gạch bỏ; đo nhanh:
      N1 pH 6,53, t° 32,4, clo dư 0,00; N2 7,41, 33,0, clo dư 0,42; N3 7,83, 31,1; N4 7,98, 35,4; N5 7,33, 33,9;
      N6 7,46, 31,6; N7 7,62, 35,3; N8 7,17, 33,0; N9 6,93, 32,4; N10 7,49, 31,3; N11 7,43, 30,5;
    kiểm tra pH 4,00 / 7,01 / 9,93 đạt; GĐV Hồ Nguyễn Huy.
  - `PYC-Q326D135226.N2.1.pdf` (MTTN26020080, nhận 25/10/2026, dự kiến trả 13/11/2026).
  - KQTN `Q326D135226 - TBKQ\Edoc\KQ-Q326D135226-MTTN26020080.01 – .12.pdf` (17/11/2026, nhận mẫu 27/10/2026).
  - Khung làm tay của người dùng `3. TBKQ - 1 … 11 - Q326D135226.docx` (16/11/2026): mỗi mẫu một chứng thư, Căn cứ "-",
    tên vị trí, hàng đo hiện trường từng mẫu (N1, N2: nhiệt độ, pH, clo dư; N3 – N11: pH) — lấy theo đây.
  - Ảnh `0. vi tri\N1 … N11 - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (18/11/2026): "làm … có KQTN rồi"; Ni N9 "Error" tô vàng để đó; Căn cứ "-" tô vàng (đổi
  thì người dùng báo); địa điểm theo BBLM; giữ nguyên "NL" (N5), tô vàng; N7 giữ như khung.

Suy luận của AI (cần người dùng xác nhận):
  1. Căn cứ "-", không đánh giá (khung làm tay; BBLM ghi "theo QCVN 40:2011/BTNMT" nhưng không có cột / Kq, Kf).
  2. Địa điểm theo BBLM: "Khu xử lý nước thải Khu công nghiệp Nhơn Trạch 1, …" (khung ghi "Khu công nghiệp Nhơn Trạch 1").
  3. Tên vị trí theo khung làm tay, mở rộng viết tắt: "SX" -> Sản xuất, "NL" -> Nguyên liệu (chưa chắc, tô vàng),
     "VN" -> Việt Nam, "CP" -> Cổ phần, "TBVS" -> Thiết bị vệ sinh, "CN" -> Công nghiệp. N7 khung không có "Hố ga nước
     thải -" — giữ như khung.
  4. Clo dư N1 đọc 0,00 -> "< 0,1" (như khung).

Chạy:  python build_135226.py [--pdf] [--only 1,3]

TỪ 18/11/2026: TBKQ điền vào FORM GỐC của người dùng bằng `fill_135226.py` (như 131229). Chạy lại script này sẽ dựng lại TBKQ theo KQTN (sai thứ tự / tên chỉ tiêu) — chỉ dùng cho BCKQ nếu cần.
"""
import argparse
import os
from nuoc_common import run, WORKSPACE

IN = os.path.join(WORKSPACE, "Input", "135226")
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
DR = lambda n: os.path.join(IN, f"3. TBKQ - {n} - Q326D135226.docx")       # noqa: E731
HG = "Hố ga nước thải - "
SAMPLES = [  # (n, vị trí, field)
    (1, "Điểm nước thải đầu vào tại Nhà máy xử lý nước thải Khu công nghiệp Nhơn Trạch 1",
     {"Nhiệt độ": "32,4", "pH": "6,53", "Clo dư": "< 0,1"}),
    (2, "Điểm nước thải đầu ra tại mương đo lưu lượng sau hệ thống xử lý nước thải",
     {"Nhiệt độ": "33,0", "pH": "7,41", "Clo dư": "0,42"}),
    (3, HG + "Công ty TNHH May mặc USV Malaysia", {"pH": "7,83", "Nhiệt độ": "31,1"}),
    (4, HG + "Công ty TNHH 1L Kwang Vina", {"pH": "7,98", "Nhiệt độ": "35,4"}),
    (5, HG + "Công ty TNHH Sản xuất Thun và NL May Premier Việt Nam", {"pH": "7,33", "Nhiệt độ": "33,9"}),
    (6, HG + "Công ty TNHH White glove", {"pH": "7,46", "Nhiệt độ": "31,6"}),
    (7, "Công ty TNHH Posco VST", {"pH": "7,62", "Nhiệt độ": "35,3"}),
    (8, HG + "Công ty TNHH Thực phẩm NFC", {"pH": "7,17", "Nhiệt độ": "33,0"}),
    (9, HG + "Công ty Cổ phần Thiết bị vệ sinh Caesar Việt Nam", {"pH": "6,93", "Nhiệt độ": "32,4"}),
    (10, HG + "Công ty TNHH Hirota Precision", {"pH": "7,49", "Nhiệt độ": "31,3"}),
    (11, HG + "Công ty TNHH Công nghiệp Kang Yuan Việt Nam", {"pH": "7,43", "Nhiệt độ": "30,5"}),
]

CONFIG = dict(
    code6="135226", loai="nuoc thai", kind="nt",
    cert_date=None,                                 # có KQTN -> ngày KQTN
    client_vn="Công ty TNHH MTV Phát triển Đô thị và Khu công nghiệp IDICO",
    address_vn="Đường Tôn Đức Thắng, Khu công nghiệp Nhơn Trạch 1, phường Nhơn Trạch, Thành phố Đồng Nai",
    reg_date="23/10/2026", dates=["27/10/2026"], time_bckq="Ngày 27/10/2026 từ 08:45 đến 15:30",
    site_vn="Khu xử lý nước thải Khu công nghiệp Nhơn Trạch 1, phường Nhơn Trạch, Thành phố Đồng Nai",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    equip=["mettler_s2", "hach_dr900"], executor="Huy",
    witness_vn="chị Dương Thị Thanh Hoa – quản đốc",
    req_date="13/11/2026",
    highlight=["NL May Premier", "Error"], hl_can_cu=True,   # Error: KQTN N9 ghi Ni "Error" — chờ người dùng
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR900"},
    bckq_notes=["- Mẫu N12 (Công ty TNHH MTV Austdoor Nhơn Trạch), N13 (Công ty Cổ phần Công nghiệp Tung Kuang): gạch bỏ "
                "trên BBLM, không lấy mẫu.",
                "- Clo dư mẫu N1 đọc trên máy 0,00 mg/L -> ghi \"< 0,1\"."],
    samples=[dict(kind="nt", cert_no=n, ky_hieu=f"N{n}", kqtn_key=f"N{n}", time="-", vi_tri=vt, field=f,
                  draft=DR(n), equip=["mettler_s2", "hach_dr900"] if n <= 2 else ["mettler_s2"], photo=PH(f"N{n}"),
                  photo_cm={9: 6.0, 11: 5.0}.get(n, 7.0),   # /9, /11 vừa 1 trang phụ lục (/7 13 dòng: vẫn 2 trang)
                  col_widths={9: [7, 32, 11, 35, 15],
                              11: [6, 39, 10, 31, 14]}.get(n))   # cột Phương pháp rộng: /7, /9, /11 vừa 1 trang phụ lục
             for n, vt, f in SAMPLES],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
