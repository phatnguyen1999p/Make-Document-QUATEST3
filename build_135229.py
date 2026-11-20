# -*- coding: utf-8 -*-
r"""
Q326D135229 — Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO — KCN Phú Mỹ 2: nước sinh hoạt N1 – N5 (QCVN 01-1:2024/BYT),
nước thải N6 – N16 (16 mẫu, 16 chứng thư). CHƯA CÓ KQTN -> bản khung, kết quả phòng thí nghiệm "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D135229 - IDICO Phú Mỹ 2 - Nước thải` về Input\135229, chỉ đọc):
  - `0. BBLM - 1.pdf` (4 trang, Nguyễn Hoàng Việt): ngày 28/10/2026, PYC Q326D135229 ngày 23/10/2026, 09:30 – 12:30; "nước sinh
    hoạt theo QCVN 01-1:2024/BYT và nước thải theo QCVN 40:2011/BTNMT"; địa điểm (viết tay) "Khu công nghiệp Phú Mỹ 2,
    phường Tân Phước, TP. HCM"; Hach DR900 + Mettler Seven2Go S2; N2.TC.N.02, N2.TC.N.06; chỉ định: anh Trần Huỳnh Chiêu –
    chuyên viên; N8 – N11 (trạm bơm, chỉ lấy mẫu) 10:04, 10:08, 09:53, 10:16; N12 pH 8,26 / 32,8 °C, clo dư 0,03 (không ghi
    giờ); N13 10:35 pH 7,39 / 30,2; N14 11:00 pH 7,60 / 32,3; N15 11:25 pH 7,88 / 32,1; N16 11:40 pH 7,97 / 31,9.
  - `0. BBLM - 2.pdf` (3 trang, Hồ Nguyễn Huy): N1 – N5 nước sinh hoạt (pH, t°, clo dư, mùi), N6 đầu vào pH 8,08 / 30,9,
    N7 đầu ra pH 8,31 / 30,9, clo dư 0,31, clo dư tự do 0,14 (theo yêu cầu KH); BBLM này không ghi giờ lấy mẫu từng mẫu.
  - `0. Phiếu YCTN - Q326D135229.N2.1.pdf` (MTTN26020081): dự kiến trả 13/11/2026; mẫu QC.
  - Khung làm tay của người dùng `3. TBKQ - 1…16 - Q326D135229.docx`, `2. BCKQ - Q326D135229.docx` (tên vị trí, thông số,
    phương pháp, mức NSH và mức Giấy phép môi trường 91/GPMT-BTNMT cho N12 – N16 — lấy từ đây).
  - Ảnh `0. vi tri\N1…N16 - Vị trí lấy mẫu.jpg`.

Quyết định của người dùng (16/11/2026): quy trình NSH để trống, clo dư < 0,1 -> "< 0,1" (tô vàng), giờ trống để "………".

Suy luận của AI (cần người dùng xác nhận):
  1. Giờ lấy mẫu N1 – N7 và N12: BBLM không ghi -> "………" tô vàng.
  2. Clo dư N5 (0,00) và N12 (0,03) đọc dưới 0,1 mg/L -> "< 0,1" (AGENTS mục 8C).
  3. Quy trình NSH để "………" tô vàng như 134486 (BBLM, khung làm tay ghi N2.TC.N.02).

Quyết định của người dùng (20/11/2026): "quy trình thực hiện nước sinh hoạt là N2.TC.N.01 … sửa và xóa tô vàng"
  -> NSH ghi QUATEST 3 3030:2025, N2.TC.N.01 (thay ………/N2.TC.N.02 tô vàng), BCKQ N2.TC.N.01, N2.TC.N.06.

Người dùng (20/11/2026): "Phú Mỹ 2 có kết quả rồi, điền vào luôn đi, điền kết quả theo form" -> KQTN MTTN26020081.01–.17
  (20/11/2026); TBKQ điền vào form `3. TBKQ - n` bằng `python fill_form_nhieu.py 135229 --pdf` (script này chỉ còn dùng
  dựng BCKQ: chạy rồi chép lại TBKQ của fill_form_nhieu).

Chạy:  python build_135229.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER, NSH_STD, clo, params_from_draft, limits_from_draft

IN = os.path.join(WORKSPACE, "Input", "135229")
DR = lambda n: os.path.join(IN, f"3. TBKQ - {n} - Q326D135229.docx")      # noqa: E731
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731
KCN = "KCN Phú Mỹ II"
GP = "phụ lục kèm theo Giấy phép môi trường số 91/GPMT-BTNMT ngày 18/04/2025"
NSH = dict(kind="nsh", scope_vn="Lấy mẫu và giám định chất lượng nước sinh hoạt", procedure="QUATEST 3 3030:2025, N2.TC.N.01",
           standard_vn=NSH_STD, chi_tieu="10 thông số nhóm A nêu tại QCVN 01-1:2024/BYT",
           limit_head="Ngưỡng giới hạn cho phép ^(1)",
           limit_note="Ngưỡng giới hạn tối đa cho phép quy định tại QCVN 01-1:2024/BYT.",
           intro="So sánh kết quả thử nghiệm các thông số của mẫu nước nêu trên so với mức quy định tương ứng tại "
                 "QCVN 01-1:2024/BYT:", equip=["mettler_s2", "hach_dr900"])
GPMT = dict(standard_vn=GP[0].upper() + GP[1:], limit_head="Giới hạn tối đa ^(1)",
            limit_note=f"Giới hạn tối đa cho phép tại {GP}.",
            intro=f"So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên với mức quy định tương ứng tại {GP}:")


def nsh(n, vi_tri, ph, t, c):
    return dict(NSH, ky_hieu=f"135229-N{n}", kqtn_key=f"N{n}", time=PLACEHOLDER, vi_tri=vi_tri, draft=DR(n),
                params=params_from_draft(DR(n)), limits=limits_from_draft(DR(n)), photo=PH(f"N{n}"),
                field={"pH": ph, "Nhiệt độ": t, "Clo dư tự do": clo(c), "Clo dư": clo(c), "Mùi": "Không có mùi lạ"})


def nt(n, time, vi_tri, ph=None, t=None, equip=("mettler_s2",), gp=False, **extra):
    f = {"pH": ph, "Nhiệt độ": t} if ph else {}
    f.update(extra)
    d = dict(kind="nt", ky_hieu=f"135229-N{n}", kqtn_key=f"N{n}", time=time, vi_tri=vi_tri,
             params=params_from_draft(DR(n)), field=f, equip=list(equip) if ph else [], photo=PH(f"N{n}"))
    if gp:
        d.update(GPMT, limits=limits_from_draft(DR(n)))
    return d


CONFIG = dict(
    code6="135229", loai="nuoc", kind="nt",
    client_vn="Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    address_vn="Số 48, khu phố 1, phường Long Hưng, Thành phố Đồng Nai",
    reg_date="23/10/2026", dates=["28/10/2026"], time_bckq="Ngày 28/10/2026 từ 09:30 đến 12:30",
    site_vn="Khu công nghiệp Phú Mỹ 2, phường Tân Phước, Thành phố Hồ Chí Minh",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    scope_bckq="Lấy mẫu và giám định chất lượng nước sinh hoạt, nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    procedure_bckq="QUATEST 3 3030:2025, N2.TC.N.01, N2.TC.N.06",
    equip=["mettler_s2", "hach_dr900"], executor="Huy + Việt",
    witness_vn="anh Trần Huỳnh Chiêu – chuyên viên",
    req_date=" ",                                   # người dùng 20/11/2026: "để trống ngày yêu cầu nhận kq"
    tre_han="[  ] có		[  ] không",               # không có ngày yêu cầu -> không đánh ô trễ hạn (như 131873)
    highlight=["< 0,1"],
    field_devices={"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ": "Mettler Toledo Seven2Go S2", "Clo dư": "Hach DR900",
                   "Clo dư tự do": "Hach DR900", "Mùi": "-"},
    bckq_notes=["- Clo dư tự do của mẫu N7 đo theo yêu cầu của khách hàng (theo BBLM)."],
    samples=[
        nsh(1, "Mẫu chung 2 khu công nghiệp Phú Mỹ II và Phú Mỹ II mở rộng", "7,27", "30,6", "0,51"),
        nsh(2, "Giữa KCN PMII: Ngã 4 đường D3 x N6 (Gần Công ty TNHH Trần Tiến Thịnh)", "7,01", "30,4", "0,45"),
        nsh(3, "Cuối KCN PMII: Ngã 3 đường N6 x N2 (Gần Công ty TNHH Kho Vận Cảng Bàn Thạch)", "7,07", "30,2", "0,40"),
        nsh(4, "Giữa KCN PMII MR: Ngã 3 đường D3 x D9 (Gần Công ty TNHH Kính Nổi Siêu Trắng Phú Mỹ II)",
            "6,97", "30,6", "0,43"),
        nsh(5, "Cuối KCN PMII MR: Cuối đường D9 (Gần Công ty Cổ phần Sản xuất Container Hòa Phát)", "6,94", "30,7", "0,00"),
        nt(6, PLACEHOLDER, f"Nước thải đầu vào Nhà máy XLNT tập trung {KCN}", "8,08", "30,9"),
        nt(7, PLACEHOLDER, f"Nước thải đầu ra Nhà máy XLNT tập trung {KCN}", "8,31", "30,9", ("mettler_s2", "hach_dr900"),
           **{"Clo dư": clo("0,31"), "Clo dư tự do": clo("0,14")}),
        nt(8, "10:04", "Nước thải Trạm bơm chuyển bậc 1"),
        nt(9, "10:08", "Nước thải Trạm bơm chuyển bậc 2"),
        nt(10, "09:53", "Nước thải Trạm bơm chuyển bậc 3"),
        nt(11, "10:16", "Nước thải Trạm bơm trung chuyển PMII MR"),
        nt(12, PLACEHOLDER, "Nước thải Công ty TNHH Hyosung Đồng Nai - BIO Bà Rịa - Vũng Tàu", "8,26", "32,8",
           ("mettler_s2", "hach_dr900"), gp=True, **{"Clo dư": clo("0,03")}),
        nt(13, "10:35", "Nước thải Công ty TNHH Thương mại Dịch vụ Hong Ming Vũng Tàu", "7,39", "30,2", gp=True),
        nt(14, "11:00", "Nước thải Công ty TNHH Ống thép hàn xoắn Nippon Steel Việt Nam", "7,60", "32,3", gp=True),
        nt(15, "11:25", "Nước thải Công ty Cổ phần PET LIKE", "7,88", "32,1", gp=True),
        nt(16, "11:40", "Nước thải Công ty Cổ phần Thép Povina", "7,97", "31,9", gp=True),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
