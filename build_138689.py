# -*- coding: utf-8 -*-
r"""
Việt Úc — nước thải sau xử lý, 1 mẫu / hồ sơ, 1 chứng thư (mã trần), chứng thư TIẾNG VIỆT (PYC KH mục 15).
Dùng chung cho 3 hồ sơ: 138689 (Sóc Trăng), 138691 (Cà Mau, `build_138691.py`), 138693 (Bạc Liêu, `build_138693.py`).
BẢN KHUNG: CHƯA CÓ KQTN -> phương pháp + kết quả phòng thí nghiệm "………" tô vàng (AGENTS mục 13); ngày ban hành = ngày dựng.

Nguồn chung (chép từ Z `0.VIMCERTS078\Q326D<6 số> - Việt Úc - …` về Input\<6 số>, chỉ đọc):
  - `0. BBLM - <6 số>.pdf` (2 trang, scan, đã ký): PYC ngày 30/10/2026; thiết bị Hach DR900 + Mettler Seven2Go S2;
    QTGĐ 3030.2025, N2.TC.N.06; chỉ định "Anh Trần Thanh Liêm – Chức vụ: chuyên viên kỹ thuật"; 1 vị trí; trang 2 kiểm
    tra pH 4 / 7 / 10 đạt; người thực hiện Hồ Nguyễn Huy.
  - `0. Phiếu yêu cầu KH - <6 số>.pdf` (scan): chứng thư Tiếng Việt; người liên hệ Trần Thanh Liêm.
  - Phiếu YCTN `PYC-Q326D<6 số>.N2.1.pdf` / `0. Phiếu YCTN-…` (thông số thử nghiệm, dự kiến trả 30/11/2026).
  - `3. TBKQ.docx`, `2. BCKQ.docx` = bản chép TBKQ / BCKQ hồ sơ 120476 (khách khác) — chỉ tham khảo cách trình bày,
    KHÔNG phải form làm sẵn của hồ sơ này (không điền vào). `Q326D… - Thực hiện\*.docx` là mẫu BBLM / ghi chép soạn sẵn
    (138689 - BBLM.docx còn ghi hồ sơ 147190 bùn thải) — không dùng.

138689 — Công ty TNHH Thủy sản Việt Úc - Sóc Trăng (MST 2200732572):
  - BBLM 17/11/2026, 13:50 – 14:45; "theo QCVN 40:2011/BTNMT, (Cột B, Kq = 1,2, Kf = 1,1)"; địa điểm viết tay "Khu xử lý
    nước thải của công ty TNHH Thủy sản Việt Úc – Sóc Trăng", Ấp Huỳnh Kỳ, xã Vĩnh Hải, thành phố Cần Thơ; vị trí "Nước
    thải sau xử lý", 138689 - Str, 14h20: pH 8,74 / 33,0 °C, clo dư 0,18 mg/L.
  - YCTN MTTN26020758 (.01 STr 8 thông số, .02 QC amoni).

Quyết định của người dùng (19/11/2026): "làm hồ sơ 138689, 138691, 138693"; trả lời 5 điểm hỏi: "1. oke, tô vàng nhé
  2. oke, tô vàng cho tôi 3. ghi theo yêu cầu khách trước, có gì cần thay đổi thì tôi báo, tô vàng nhé 4. oke 5. oke"
  -> mức Cmax giữ, tô vàng cột mức (cả 3 hồ sơ); 138691 tên QCVN 11-MT:2015/BTNMT + bảng C tô vàng, dầu mỡ ghi "dầu mỡ
  khoáng" theo PYC KH (tô vàng); nhiệt độ không đưa vào TBKQ; tên 138693 theo PYC.

Suy luận của AI (cần người dùng xác nhận):
  1. Cmax = C × Kq × Kf (pH, coliform không nhân hệ số); giá trị C cột B theo quy chuẩn ghi trên BBLM.
  2. Bảng kết quả gồm pH, clo dư (đo hiện trường, PYC KH có yêu cầu) + thông số YCTN; nhiệt độ có đo trên BBLM nhưng PYC
     KH không yêu cầu -> không đưa vào bảng (ghi ở BCKQ).
  3. Không dựng theo `3. TBKQ.docx` (bản của hồ sơ 120476).

Chạy:  python build_138689.py [--pdf]
"""
import argparse
import datetime
import os

from nuoc_common import run, WORKSPACE, PLACEHOLDER as P

M = "Đo tại hiện trường"
NAMES = {  # khóa ngắn -> (tên thông số, đơn vị) theo form 120476 (người dùng 19/11/2026: "sửa lại theo form 120476")
    "pH": ("pH", "-"), "clo": ("Clo dư", "mg/L"),
    "col": ("Tổng coliforms", "MPN/ 100 mL"), "nh4": ("Amoni (tính theo Nitơ)", "mg/L"),
    "dmk": ("Dầu mỡ khoáng", "mg/L"), "dtv": ("Dầu mỡ động thực vật", "mg/L"),
    "tn": ("Tổng nitơ", "mg/L"), "tss": ("Chất rắn lơ lửng (TSS)", "mg/L"),
    "cod": ("Nhu cầu oxy hóa học (COD) quy về O₂", "mg/L"), "bod": ("Nhu cầu oxy sinh hoá (BOD₅) quy về O₂", "mg/L"),
    "tp": ("Tổng phốt pho", "mg/L")}
CT = {"pH": "pH", "clo": "clo dư", "tss": "TSS", "bod": "BOD₅", "cod": "COD", "nh4": "amoni", "tn": "tổng nitơ",
      "tp": "tổng phốt pho", "dmk": "dầu mỡ khoáng", "dtv": "dầu mỡ động thực vật", "col": "tổng coliforms"}
ORDER = ["tss", "bod", "cod", "nh4", "tn", "tp", "dmk", "dtv", "col"]   # thứ tự thông số như form 120476
EXEMPT = {"pH", "col"}                               # không nhân Kq, Kf


def vn(x):
    s = f"{x:.4f}".rstrip("0").rstrip(".")
    if "." not in s and len(s) > 3:                  # 5000 -> "5 000" (như các khung nước thải)
        s = f"{int(s):,}".replace(",", " ")
    return s.replace(".", ",")


def make(code6, client, address, site, vi_tri, ky, kq_key, date, t0, t1, tsample, ph, clo, std_code, std_title,
         col_c, kq, kf, keys, photo, notes=(), highlight=(), hl_limits=True, hl_names=()):
    """keys: thông số YCTN theo thứ tự phiếu; col_c: {khóa: C cột B} (pH là chuỗi khoảng)."""
    f = kq * kf
    limits = {}
    for k, c in col_c.items():
        name = NAMES[k][0]
        limits[name] = c if isinstance(c, str) else (vn(c) if k in EXEMPT else vn(round(c * f, 4)))
    keys = sorted(keys, key=ORDER.index)
    params = [(NAMES["pH"][0], "-", M), (NAMES["clo"][0], "mg/L", M)] + [(NAMES[k][0], NAMES[k][1], P) for k in keys]
    kqs, kfs = vn(kq), vn(kf)
    return dict(
        code6=code6, loai="nuoc thai", kind="nt",
        cert_date=datetime.date.today().strftime("%d/%m/%Y"),     # chưa KQTN -> ngày dựng (AGENTS mục 2)
        client_vn=client, address_vn=address,
        reg_date="30/10/2026", dates=[date], time_bckq=f"Ngày {date} từ {t0} đến {t1}",
        site_vn=site, scope_vn="Lấy mẫu và giám định chất lượng nước thải",
        procedure="QUATEST 3 3030:2025, N2.TC.N.06",
        equip=["mettler_s2", "hach_dr900"], executor="Huy",
        witness_vn="anh Trần Thanh Liêm – chuyên viên kỹ thuật",
        req_date="30/11/2026",
        field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR900"},
        chi_dinh=True,                                  # "được chỉ định bởi" như form 120476
        bckq_notes=list(notes),
        highlight=list(highlight),
        samples=[dict(
            ky_hieu=ky, kqtn_key=kq_key, time=tsample, vi_tri=vi_tri, params=params,
            standard_vn=f"{std_code} - {std_title}", limits=limits,
            limit_head="Giá trị tối đa ^(1)",
            limit_note=f"Giá trị tối đa tham chiếu được tính toán theo giá trị C - cột B, hệ số Kq = {kqs} và "
                       f"Kf = {kfs} tại {std_code}.",            # câu ghi chú theo form 120476
            col_widths=[7, 27, 12, 24, 14, 16],                # gần form 120476; cột đơn vị đủ "MPN/ 100 mL" 2 dòng
            intro=f"So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên với giá trị tối đa cho phép tương "
                  f"ứng tại {std_code}, cột B (Kq = {kqs}; Kf = {kfs}):",
            chi_tieu=", ".join(CT[k] for k in ["pH", "clo"] + keys),   # kiểu liệt kê như form 120476
            field={"pH": ph, "Clo dư": clo}, photo=photo,
            hl_limits=hl_limits,                       # mức tính Cmax chờ người dùng soát (người dùng 19/11/2026: tô vàng)
            highlight=list(hl_names))],
    )


# QCVN 40:2011/BTNMT bảng 1 cột B (giá trị C)
Q40B = {"pH": "5,5 – 9", "clo": 2, "col": 5000, "nh4": 10, "dmk": 10, "tn": 40, "tss": 100, "cod": 150, "bod": 50,
        "tp": 6}
Q40 = ("QCVN 40:2011/BTNMT", "Quy chuẩn kỹ thuật quốc gia về nước thải công nghiệp")

CONFIG = make(
    "138689", "Công ty TNHH Thủy sản Việt Úc - Sóc Trăng", "Ấp Huỳnh Kỳ, xã Vĩnh Hải, Thành phố Cần Thơ",
    "Khu xử lý nước thải của Công ty TNHH Thủy sản Việt Úc - Sóc Trăng, Ấp Huỳnh Kỳ, xã Vĩnh Hải, Thành phố Cần Thơ",
    "Nước thải sau xử lý", "138689 - Str", "STr", "17/11/2026", "13:50", "14:45", "14:20", "8,74", "0,18",
    *Q40, Q40B, 1.2, 1.1, ["col", "nh4", "dmk", "tn", "tss", "cod", "bod", "tp"],
    os.path.join(WORKSPACE, "Input", "138689", "Q326D138689 - Photo", "0.Photo", "Vị trí lấy mẫu.jpg"),
    notes=["- Nhiệt độ nước thải đo tại hiện trường: 33,0 °C."])

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
