# -*- coding: utf-8 -*-
r"""
Q326D138167 — Công ty TNHH Jones Lang Lasalle — nước thải tại Công ty TNHH Intel Products Việt Nam: NT1 (đầu vào HTXL),
NT2 (đầu ra tại hố ga đấu nối KCN), 2 chứng thư /1, /2 + BCKQ. CHƯA CÓ KQTN -> bản khung "………" tô vàng.
Phần ĐO ỒN NGOÀI TRỜI (4 lần đo, BBLM dòng 03) CHƯA dựng — chờ người dùng chốt cách trình bày.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D138167 - JLL - NT` về Input\138167, chỉ đọc):
  - `0\BBLM.pdf` (2 trang, scan): BBLM ngày 09/11/2026, PYC Q326D138167 ngày 30/10/2026; "Công ty TNHH Jones Lang Lasalle",
    "37 Tôn Đức Thắng, Phường Bến Nghé, Quận 1, TP. Hồ Chí Minh"; "Lấy mẫu và phân tích chất lượng nước thải"; địa điểm
    "CÔNG TY TNHH INTEL PRODUCTS VIỆT NAM - Địa điểm: Lô I2, đường D1, Khu Công nghệ cao TP.HCM, Phường Tăng Nhơn Phú,
    TP.HCM"; 10h30 – 15h00; Hanna HI-98190 (pH), Hanna HI98192 (EC/TDS); QUATEST3 3030.2025; N2.TC.N.02; N2.TC.N.06;
    chỉ định: Trần Thị Liễu Quyên – nhân viên EHS; NT1 "Nước thải đầu vào HTXL", NT2 "Nước thải đầu ra tại hố ga đấu nối
    KCN" (chữ cuối đọc "KCN"/"KCNC" — chưa chắc); 03 "Đo ồn ngoài trời", tiếng ồn, 4 (lần); đo nhanh 3 lần:
      NT1 10:55 pH 7,046/7,047/7,048, t° 31,4×3, TDS 0,414/0,415/0,414 g/L;
      NT2 14:45 pH 7,477/7,475/7,475, t° 33,8×3, TDS 0,329/0,330/0,329 g/L; GĐV Nguyễn Minh Quân.
  - `0\138167 - Biên bản hiện trường.xlsx` sheet "Phụ lục I": 25 chỉ tiêu (TSS, BOD5, COD, F⁻, amoni, T-N, T-P, CN⁻, Hg,
    Cu, Fe, Zn, Pb, As, Ni, Cr tổng, CHĐBM anion, PCB, tổng phenol, Coliform, độ màu, Cd, Mn, Cr⁶⁺, dầu mỡ ĐTV).
  - `0\KQ On\1_1 octave.xlsx` + Rion NL (Auto_0190 – 0193): LAeq 30 phút 56,7 / 54,2 / 56,4 / 61,4 dBA lúc 11:16 /
    11:58 / 12:51 / 13:31 ngày 09/11/2026 (chưa dùng).
  - Ảnh `0\Photo\`.

Quyết định của người dùng (18/11/2026):
  1. Đo ồn: "không, làm chung 1 bản" -> 01 chứng thư (mã trần) gồm NT1, NT2 (mỗi mẫu một cột) + bảng đo ồn ngoài trời.
  2. Địa chỉ: "xem lại hồ sơ Jones Lang Lasalle gần nhất xem TBKQ là để địa chỉ nào thì để địa chỉ đó" -> theo TBKQ
     126159 (tháng 10/2026): "Số 37, đường Tôn Đức Thắng, phường Sài Gòn, Thành phố Hồ Chí Minh". Tên khách giữ theo
     BBLM (tô vàng; 126159 ghi "Công ty TNHH Jones Lang LaSalle (Việt Nam)").
  3. Quy chuẩn nước thải chưa có: Căn cứ "-", tô vàng.

Suy luận của AI (cần người dùng xác nhận):
  - Bảng đo ồn: vị trí "Ngoài trời" (BBLM chỉ ghi "Đo ồn ngoài trời"), khung giờ = giờ bắt đầu + 30 phút, máy Rion
    NL-53 (file .rnh), không mức / không đánh giá, câu VIMCERTS tô vàng; nội dung yêu cầu thêm "đo tiếng ồn ngoài trời"
    (tô vàng).
  1. Tên / địa chỉ khách theo BBLM (tô vàng); các hồ sơ JLL khác ghi "Công ty TNHH Jones Lang LaSalle (Việt Nam)",
     "Số 37, đường Tôn Đức Thắng, phường Sài Gòn".
  2. Không có quy chuẩn / mức -> Căn cứ "-", không đánh giá; phương pháp theo KQTN khi về.
  3. pH, nhiệt độ, TDS = trung bình 3 lần (TDS g/L -> mg/L).

Chạy:  python build_138167.py [--pdf]
"""
import argparse
import datetime
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER as P, _set_hl
from nt_params import params
from build_nuoc_mau import note_table, cell_text, _keep_rows, sub_head
from docx.oxml.ns import qn

IN = os.path.join(WORKSPACE, "Input", "138167")
NAMES = ["Tổng chất rắn lơ lửng (TSS)", "BOD₅", "COD", "Florua (F⁻)", "Amoni (NH₄⁺ tính theo N)", "Tổng nitơ",
         "Tổng phốt pho", "Xyanua (CN⁻)", "Thủy ngân (Hg)", "Đồng (Cu)", "Sắt (Fe)", "Kẽm (Zn)", "Chì (Pb)",
         "Asen (As)", "Niken (Ni)", "Tổng crôm (Cr)", "Chất hoạt động bề mặt anion", "Polyclo biphenyl (PCB)",
         "Tổng phenol", "Tổng Coliform", "Độ màu", "Cadimi (Cd)", "Mangan (Mn)", "Crôm VI (Cr⁶⁺)",
         "Dầu mỡ động thực vật"]
FIELD = [("pH", "-", "Đo tại hiện trường"), ("Nhiệt độ", "°C", "Đo tại hiện trường"),
         ("TDS", "mg/L", "Đo tại hiện trường")]
CLIENT = "Công ty TNHH Jones Lang Lasalle"
ADDR = "Số 37, đường Tôn Đức Thắng, phường Sài Gòn, Thành phố Hồ Chí Minh"   # theo TBKQ JLL 126159
SCOPE = "Lấy mẫu và giám định chất lượng nước thải; đo tiếng ồn ngoài trời"
CHI_TIEU = "pH, nhiệt độ, TDS (đo tại hiện trường) và 25 thông số theo yêu cầu khách hàng"
NOISE = [("11:16 – 11:46", "56,7"), ("11:58 – 12:28", "54,2"), ("12:51 – 13:21", "56,4"), ("13:31 – 14:01", "61,4")]


def nt(n, vi_tri, time, ph, t, tds):
    return dict(ky_hieu=f"NT{n}", kqtn_key=f"NT{n}", vi_tri=vi_tri, time=time, chi_tieu=CHI_TIEU,
                field={"pH": ph, "Nhiệt độ": t, "TDS": tds}, field_tbkq=["pH", "Nhiệt độ", "TDS"],
                photo=os.path.join(IN, "0", "Photo", {1: "NT1.jpg", 2: "NT2.jpeg"}[n]))


def noise_section(b, doc, note_proto, is_en):
    """Bảng đo tiếng ồn ngoài trời (4 lần đo 30 phút, Rion) — người dùng 18/11/2026 "làm chung 1 bản"."""
    sub_head(b, "- Kết quả đo tiếng ồn ngoài trời:")
    t = b.table(1 + len(NOISE), 4, [10, 34, 30, 26])
    for j, h in enumerate(["STT", "Vị trí đo", "Thời gian đo", "Mức âm tương đương LAeq, dBA"]):
        cell_text(t.rows[0].cells[j], h)
    for i, (tm, v) in enumerate(NOISE, start=1):
        cell_text(t.cell(i, 0), str(i))
        cell_text(t.cell(i, 2), tm)
        cell_text(t.cell(i, 3), v)
    t.cell(1, 1).merge(t.cell(len(NOISE), 1))
    cell_text(t.cell(1, 1), "Ngoài trời")
    for r_el in t.cell(1, 1)._tc.iter(qn("w:r")):
        _set_hl(r_el)
    _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:-1])
    nt_el = note_table(b, note_proto, [("●", "Các thông số được chứng nhận VIMCERTS 078."),
                                       ("●", "Thời gian đo mỗi lần 30 phút ngày 09/11/2026; không có mức quy định, "
                                             "không đánh giá.")])
    for r_el in nt_el.findall(qn("w:tr"))[1].iter(qn("w:r")):
        _set_hl(r_el)


CONFIG = dict(
    code6="138167", loai="nuoc thai", kind="nt",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2)
    client_vn=CLIENT, address_vn=ADDR,
    reg_date="30/10/2026", dates=["09/11/2026"], time_bckq="Ngày 09/11/2026 từ 10:30 đến 15:00",
    site_vn="Công ty TNHH Intel Products Việt Nam - Lô I2, đường D1, Khu Công nghệ cao Thành phố Hồ Chí Minh, phường "
            "Tăng Nhơn Phú, Thành phố Hồ Chí Minh",
    scope_vn=SCOPE,
    procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.06",
    equip=["hanna_hi98190", "hanna_hi98192", "rion_nl53"], executor="Quân",
    witness_vn="chị Trần Thị Liễu Quyên – nhân viên EHS",
    req_date=P, highlight=[CLIENT, "; đo tiếng ồn ngoài trời"], hl_can_cu=True,
    field_devices={"pH": "Hanna HI 98190", "Nhiệt độ": "Hanna HI 98190", "TDS": "Hanna HI 98192"},
    bckq_notes=["- Số đo hiện trường ghi là trung bình 3 lần đo trên BBLM: NT1 pH 7,046/7,047/7,048, nhiệt độ 31,4 °C, "
                "TDS 0,414/0,415/0,414 g/L; NT2 pH 7,477/7,475/7,475, nhiệt độ 33,8 °C, TDS 0,329/0,330/0,329 g/L.",
                "- Đo ồn ngoài trời (BBLM dòng 03, 4 lần đo 30 phút, máy Rion): LAeq 56,7 / 54,2 / 56,4 / 61,4 dBA lúc "
                "11:16 / 11:58 / 12:51 / 13:31 (máy Rion NL-53, file Auto_0190 – 0193, 1_1 octave.xlsx)."],
    samples=[dict(kind="nt", ky_hieu="NT1, NT2", kqtn_key="NT", chi_tieu=CHI_TIEU, field_tbkq=["pH", "Nhiệt độ", "TDS"],
                  params=FIELD + params(NAMES), equip=["hanna_hi98190", "hanna_hi98192", "rion_nl53"],
                  vi_tri="-", time="-", after_results=noise_section,
                  members=[nt(1, "Nước thải đầu vào hệ thống xử lý nước thải", "10:55", "7,05", "31,4", "414"),
                           nt(2, "Nước thải đầu ra tại hố ga đấu nối khu công nghệ cao", "14:45", "7,48", "33,8",
                              "329")])],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
