# -*- coding: utf-8 -*-
r"""
Q326D143325 — Công ty Cổ phần Môi trường Phú Xuân — nước dưới đất tháng 11 (QCVN 09:2023/BTNMT), 18 giếng, 18 chứng thư.
BẢN KHUNG: CHƯA CÓ BBLM ĐÃ KÝ, CHƯA CÓ KQTN -> ngày / giờ lấy mẫu, pH, nhiệt độ, người chứng kiến, người thực hiện, ảnh,
phương pháp và kết quả phòng thí nghiệm để "………" tô vàng (AGENTS mục 13).

Nguồn (chép từ Z `0.VIMCERTS078\Q326D143325 - Phú Xuân - nước ngầm tháng 11\Q326D143325 - Vĩnh Tân đợt 1` về Input\143325,
chỉ đọc):
  - `0. Phiếu YCTN-Q326D143325.N2.1.pdf` (MTTN26022147, nhận 14/11/2026, dự kiến trả 07/12/2026): 21 mẫu = .01 – .15
    VT5-1 … VT7-8, .16 – .18 QC, .19 VT6-5, .20 PQL-1, .21 PQL-2; mỗi mẫu 11 chỉ tiêu: chỉ số permanganat, Coliforms,
    amoni, As, Cl⁻, Mn, nitrat, nitrit, Fe, TDS, độ cứng.
  - BBLM soạn sẵn (Word, chưa ngày giờ, chưa số đo, chưa ký): `143325 - BBLM Vĩnh Tân 5 / 6 / 7.docx`,
    `143325 - BBLM Phạm Quốc Liêm.docx` — PYC Q326D143325 ngày 10/11/2026; địa chỉ "Số 113/37/44, đường 30/4, khu phố Vinh
    Sơn, phường Phú Lợi, Thành phố Hồ Chí Minh"; "Lấy mẫu, giám định chất lượng nước dưới đất theo QCVN 09:2023/BTNMT";
    Mettler Seven2Go S2; QTGĐ 3030.2025, N2.TC.N.05; địa điểm và tên giếng từng trại.
  - Khung chứng thư người dùng hồ sơ trước cùng khách `3. TBKQ - 1 - 1 - Q326D128082.docx` (đợt 10/2026): cách trình bày,
    đơn vị, mức QCVN 09:2023/BTNMT (28 thông số; hồ sơ này chỉ 11 thông số theo YCTN).
  - `0. Phiếu yêu cầu KH + Báo giá.pdf` (scan).
  - Thư mục Z "Vĩnh Tân đợt 2" có BBLM Vĩnh Tân 3 (VT3-1 … 3) không có trong YCTN và BBLM Vĩnh Tân 6 chỉ 4 giếng -> không
    dùng.

Quyết định của người dùng (18/11/2026): "Dựng khung luôn".

Suy luận của AI (cần người dùng xác nhận):
  1. 18 chứng thư theo thứ tự YCTN (bỏ mẫu QC): /1 – /15 = VT5-1 … VT7-8, /16 = VT6-5, /17 = PQL-1, /18 = PQL-2.
  2. 11 thông số theo YCTN + pH, nhiệt độ đo hiện trường (BBLM); mức QCVN 09:2023/BTNMT theo khung 128082.
  3. Tên vị trí, địa điểm theo BBLM soạn sẵn (mở rộng viết tắt CN, CP); VT6-3, VT6-4 cùng tên "Lê Thị Bích Thủy" chưa
     ghi số giếng (tô vàng).
  4. Nội dung yêu cầu "nước dưới đất" theo BBLM (khung 128082 ghi "nước ngầm").

Chạy:  python build_143325.py [--pdf] [--only 1,17]
"""
import argparse
import datetime
from nuoc_common import run, PLACEHOLDER as P

STD = "QCVN 09:2023/BTNMT - Quy chuẩn kỹ thuật quốc gia về chất lượng nước dưới đất"
HCM = "xã Minh Thạnh, Thành phố Hồ Chí Minh"
VT5 = "Trại Vĩnh Tân 5 Chi nhánh Công ty Cổ phần Thực phẩm Vĩnh Tân"
SITE = {"VT5": f"{VT5} tại ấp Tân Minh, {HCM}", "VT6": f"Trại Vĩnh Tân 6 tại ấp Tân Minh, {HCM}",
        "VT7": f"Trại Vĩnh Tân 7 tại ấp Đồng Sơn, {HCM}",
        "PQL": "Hộ kinh doanh Phạm Quốc Liêm, xã Long Hòa, Thành phố Hồ Chí Minh"}
F = "Đo tại hiện trường"
PARAMS = [("pH", "-", F), ("Nhiệt độ", "°C", F),
          ("Chỉ số permanganat", "mg/L", P), ("Tổng Coliform", "MPN/ 100 mL", P),
          ("Amoni (NH₄⁺ tính theo N)", "mg/L", P), ("Asen (As)", "mg/L", P), ("Clorua (Cl⁻)", "mg/L", P),
          ("Mangan (Mn)", "mg/L", P), ("Nitrat (NO₃⁻ tính theo N)", "mg/L", P), ("Nitrit (NO₂⁻ tính theo N)", "mg/L", P),
          ("Sắt (Fe)", "mg/L", P), ("Tổng chất rắn hòa tan (TDS)", "mg/L", P),
          ("Độ cứng (tính theo CaCO₃)", "mg/L", P)]
LIMITS = {"pH": "5,8 - 8,5", "Chỉ số permanganat": "4", "Tổng Coliform": "3", "Amoni (NH₄⁺ tính theo N)": "1",
          "Asen (As)": "0,05", "Clorua (Cl⁻)": "250", "Mangan (Mn)": "0,5", "Nitrat (NO₃⁻ tính theo N)": "15",
          "Nitrit (NO₂⁻ tính theo N)": "1", "Sắt (Fe)": "5", "Tổng chất rắn hòa tan (TDS)": "1500",
          "Độ cứng (tính theo CaCO₃)": "500"}
CHI_TIEU = ("pH, nhiệt độ (đo tại hiện trường), chỉ số permanganat, Coliform, amoni, asen, clorua, mangan, nitrat, nitrit, "
            "sắt, tổng chất rắn hòa tan, độ cứng")

WELLS = [  # (ký hiệu, vị trí theo BBLM) — thứ tự chứng thư theo YCTN
    ("VT5-1", f"Giếng 1 – {VT5}"), ("VT5-2", f"Giếng 2 – {VT5}"), ("VT5-3", f"Giếng 3 – {VT5}"),
    ("VT6-1", "Trại Vĩnh Tân 6 Hộ kinh doanh Võ Thị Ba – Giếng 1"),
    ("VT6-2", "Trại Vĩnh Tân 6 Hộ kinh doanh Võ Thị Ba – Giếng 2"),
    ("VT6-3", "Trại Vĩnh Tân 6 Hộ kinh doanh Lê Thị Bích Thủy"),
    ("VT6-4", "Trại Vĩnh Tân 6 Hộ kinh doanh Lê Thị Bích Thủy"),
    ("VT7-1", "Trại Vĩnh Tân 7 Hộ kinh doanh Võ Thị Ba"),
    ("VT7-2", "Trại Vĩnh Tân 7 Hộ kinh doanh Lê Thị Bích Thủy – Giếng 1"),
    ("VT7-3", "Trại Vĩnh Tân 7 Hộ kinh doanh Lê Thị Bích Thủy – Giếng 2"),
    ("VT7-4", "Trại Vĩnh Tân 7 Hộ kinh doanh Lê Thị Bích Thủy – Giếng 3"),
    ("VT7-5", "Trại Vĩnh Tân 7 Hộ kinh doanh Lê Thị Bích Thủy – Giếng 4"),
    ("VT7-6", "Trại Vĩnh Tân 7 Hộ kinh doanh Võ Thị Thanh Thảo – Giếng 1"),
    ("VT7-7", "Trại Vĩnh Tân 7 Hộ kinh doanh Võ Thị Thanh Thảo – Giếng 2"),
    ("VT7-8", "Trại Vĩnh Tân 7 Hộ kinh doanh Võ Thị Thanh Thảo – Giếng 3"),
    ("VT6-5", "Trại Vĩnh Tân 6 Hộ kinh doanh Võ Thị Thanh Thảo"),
    ("PQL-1", "Hộ kinh doanh Phạm Quốc Liêm – Giếng 1"), ("PQL-2", "Hộ kinh doanh Phạm Quốc Liêm – Giếng 2"),
]


def g(key, vi_tri):
    return dict(kind="nt", ky_hieu=f"143325-{key}", kqtn_key=key, time=P, vi_tri=vi_tri, site_vn=SITE[key[:3]],
                params=PARAMS, limits=LIMITS, standard_vn=STD, chi_tieu=CHI_TIEU,
                limit_head="Giá trị giới hạn ^(1)", limit_note="Giá trị giới hạn tối đa nêu tại QCVN 09:2023/BTNMT.",
                intro="So sánh kết quả thử nghiệm các thông số của mẫu nước dưới đất nêu trên với mức quy định tương ứng "
                      "nêu tại QCVN 09:2023/BTNMT:",
                field={"pH": P, "Nhiệt độ": P}, photo_note=P)


CONFIG = dict(
    code6="143325", loai="nuoc duoi dat", kind="nt",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2)
    client_vn="Công ty Cổ phần Môi trường Phú Xuân",
    address_vn="Số 113/37/44, đường 30/4, khu phố Vinh Sơn, phường Phú Lợi, Thành phố Hồ Chí Minh",
    reg_date="10/11/2026", dates=[P], time_bckq=P,
    site_vn=SITE["VT5"], site_bckq="; ".join(SITE.values()),
    scope_vn="Lấy mẫu và giám định chất lượng nước dưới đất",
    procedure="QUATEST 3 3030:2025, N2.TC.N.05",
    equip=["mettler_s2"], executor=P, witness_vn=P,
    req_date="07/12/2026",
    highlight=["Hộ kinh doanh Lê Thị Bích Thủy"],
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2"},
    bckq_notes=["- Bản khung: chưa có BBLM đã ký (ngày, giờ, số đo pH / nhiệt độ, người chứng kiến, người thực hiện) và "
                "chưa có KQTN MTTN26022147 (dự kiến trả 07/12/2026)."],
    samples=[g(k, v) for k, v in WELLS],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
