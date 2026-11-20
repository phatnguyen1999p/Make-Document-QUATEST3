# -*- coding: utf-8 -*-
r"""
Q326D142430 — Công ty TNHH Giấy Lee & Man Việt Nam — nước thải hằng tuần (2 mẫu NT1, NT2; 2 chứng thư /1, /2).
CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D142430 - LEEMAN - Nước thải Hằng Tuần` về Input\142430, chỉ đọc):
  - `0\0. BBLM - 142430.pdf` (2 trang, scan): BBLM ngày 16/11/2026, PYC Q326D142430 ngày 06/11/2026; "Lấy mẫu và phân tích
    chất lượng nước thải"; địa điểm "Cụm Công nghiệp tập trung Phú Hữu A - giai đoạn 1, Xã Châu Thành, Thành Phố Cần Thơ,
    Việt Nam"; 10h05 – 10h50; Mettler Seven2Go S2, Hach DR300; QUATEST3 3030.2025; N2.TC.N.02; N2.TC.N.06; chỉ định: Anh Võ
    Duy Phương – Kỹ sư vận hành; NT1 "Trước hồ sinh thái", NT2 "Điểm xả thải" (không niêm phong); đo nhanh 2 lần:
    NT1 10h20 – 10h25 pH 7,741 / 7,748, t° 35,0 / 35,1, clo dư 0,6 / 0,6; NT2 10h40 – 10h45 pH 7,795 / 7,797, t° 32,9 /
    32,9, clo dư 0,6 / 0,6; GĐV ký: Nguyễn Hoàng Việt.
  - `0\0. Phiếu kiểm soát công việc.pdf` (6 trang): PKSCV (phân công "Huy + Quân", "Việt + Minh", "+ Quân"), PYC giám
    định Q326D142430 ngày 06/11/2026 (2 bản) — "Giám định mẫu nước thải hằng tuần", 2 mẫu, "Cụ thể theo nội dung Điều 2 và
    phụ lục 1 hợp đồng số …/QN.26/N2-LEEMAN (VLM-2610-007)", MST 6300034360, chứng thư Tiếng Việt; trang trắng; điều kiện
    chung.
  - `0\PYC-Q326D142430.N2.1.pdf` (MTTN26022102, nhận 13/11/2026, dự kiến trả 27/11/2026): NT1, NT2 mỗi mẫu 25 chỉ tiêu;
    mẫu QC (BOD₅).
  - `0\142430-1 - VIMCERTS - Biên bản hiện trường.xlsx` (biên bản in sẵn), ảnh `0\0.Photo\0.NT1`, `0.NT2`.

Quyết định của người dùng (17/11/2026):
  - Quy chuẩn / mức: "Cứ để đó, chưa cần thêm gì vào" -> Căn cứ "-", không đánh giá.
  - Số đo hiện trường = trung bình 2 lần: "oke highlight chỗ đó" -> ô số đo hiện trường tô vàng (TBKQ, BCKQ).
  - Người thực hiện Việt: giữ nguyên. Địa điểm giám định: giữ như BBLM (không thêm tên đơn vị; [Location Short] chấp nhận).

Suy luận của AI (đã được người dùng xem 17/11/2026):
  1. Chưa có hợp đồng / phụ lục 1 (VLM-2610-007) -> không biết quy chuẩn, mức -> Căn cứ "-", không đánh giá.
  2. Số đo hiện trường = trung bình 2 lần đo trên BBLM (pH 2 số lẻ, nhiệt độ 1 số lẻ).
  3. Người thực hiện theo chữ ký BBLM: Việt (PKSCV ghi nhiều người).
  4. Ngày yêu cầu cấp KQTN = 27/11/2026 (ngày dự kiến trả trên PYC thử nghiệm).

Chạy:  python build_142430.py [--pdf] [--only 1,2]
"""
import argparse
import datetime
import os
from nuoc_common import run, WORKSPACE, pyc_index, pyc_sample

IN = os.path.join(WORKSPACE, "Input", "142430")
PH = lambda k: os.path.join(IN, "0", "0.Photo", f"0.{k}", {"NT1": "NT1-Vị trí lấy mẫu .jpg",   # noqa: E731
                                                          "NT2": "NT2-Vị trí lấy mẫu.jpg"}[k])
PYC, _ = pyc_index(os.path.join(IN, "0", "PYC-Q326D142430.N2.1.pdf"))
TT = ("Nhiệt độ", "pH", "Clo dư")


def s(key, time, vi_tri, ph, t, clo):
    return pyc_sample("142430", key, PYC, time, vi_tri, {"pH": ph, "Nhiệt độ": t, "Clo dư": clo}, TT,
                      ("mettler_s2", "hach_dr300"), photo=PH(key), ky_hieu=key,   # ký hiệu theo BBLM
                      hl_field=True)   # số đo hiện trường (TB 2 lần) tô vàng — người dùng 17/11/2026


CONFIG = dict(
    code6="142430", loai="nuoc thai", kind="nt",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2, 16/11/2026)
    client_vn="Công ty TNHH Giấy Lee & Man Việt Nam",
    address_vn="Cụm Công nghiệp tập trung Phú Hữu A - giai đoạn 1, xã Châu Thành, Thành phố Cần Thơ",
    reg_date="06/11/2026", dates=["16/11/2026"], time_bckq="Ngày 16/11/2026 từ 10:05 đến 10:50",
    site_vn="Cụm Công nghiệp tập trung Phú Hữu A - giai đoạn 1, xã Châu Thành, Thành phố Cần Thơ",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.06",
    equip=["mettler_s2", "hach_dr300"], executor="Việt",
    witness_vn="anh Võ Duy Phương – Kỹ sư vận hành",
    req_date="27/11/2026",
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR300"},
    bckq_notes=["- Số đo hiện trường ghi là trung bình 2 lần đo trên BBLM (NT1: pH 7,741 / 7,748, nhiệt độ 35,0 / 35,1 °C, "
                "clo dư 0,6 / 0,6 mg/L; NT2: pH 7,795 / 7,797, nhiệt độ 32,9 / 32,9 °C, clo dư 0,6 / 0,6 mg/L)."],
    samples=[
        s("NT1", "10:20 – 10:25", "Trước hồ sinh thái", "7,74", "35,1", "0,6"),
        s("NT2", "10:40 – 10:45", "Điểm xả thải", "7,80", "32,9", "0,6"),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
