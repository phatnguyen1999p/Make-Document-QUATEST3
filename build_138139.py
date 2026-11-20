# -*- coding: utf-8 -*-
r"""
Q326D138139 — Công ty TNHH Thương mại và Dịch vụ Nhất Tín (đề nghị đổi tên: Công ty TNHH Công nghệ và Dịch vụ Nhất Tín)
— nước thải sau HTXL tại bể quan trắc, Nhà máy Thuận Phương Long An; 6 mẫu NT1 – NT6 (10:20 – 11:10), 01 chứng thư
gộp (mỗi mẫu một cột kết quả, như hồ sơ trạm quan trắc 132335) + BCKQ. ĐÃ CÓ KQTN (MTTN26020975.01 – .07, 13/11/2026).

Nguồn (chép từ Z `0.VIMCERTS078\Q326D138139 - Công ty TNHH Thương mại và Dịch vụ Nhất Tín` về Input\138139, chỉ đọc):
  - `Thuc hiện\BBLM.pdf` (2 trang, scan): BBLM ngày 04/11/2026, PYC Q326D138139 ngày 29/10/2026; "535/32 Quốc Lộ 13, phường
    Hiệp Bình, TPHCM"; "Lấy mẫu và phân tích chất lượng nước thải"; địa điểm "Công ty TNHH may thêu Thuận Phương - Nhà máy
    Thuận Phương Long An - Ấp 3, xã Long Cang, tỉnh Tây Ninh, Việt Nam"; bắt đầu 10h15, kết thúc viết tay đọc "10h20"
    (trước giờ lấy NT6 — chưa chắc); máy đo pH Hanna HI-98190; QUATEST3 3030.2025; N2.TC.N.02; N2.TC.N.06; chỉ định:
    anh Nguyễn Anh Tuấn – nhân viên; NT1 – NT6 "Nước thải sau HTXL tại bể quan trắc", đo pH/t° 3 lần:
      NT1 10:20 pH 8,014/8,014/8,013, t° 32,7/32,7/32,8;  NT2 10:30 8,034/8,033/8,034, 32,8×3;
      NT3 10:40 8,063/8,063/8,062, 32,7×3;               NT4 10:50 8,078/8,077/8,077, 33,0×3;
      NT5 11:00 7,979/7,980/7,980, 33,0×3;               NT6 11:10 7,995/7,994/7,995, 33,2×3;
    (cột vị trí NT1 ghi giờ đọc giống "10h10"); ghi chú: "Đơn vị yêu cầu đề nghị thay đổi tên đơn vị yêu cầu thành:
    'Công ty TNHH Công nghệ và Dịch vụ Nhất Tín'"; GĐV Nguyễn Minh Quân.
  - `Thuc hiện\138139 - du lieu bien ban.json`: QCVN 40:2025/BTNMT, thông số pH, t°, COD, NH4, TSS, VIMCERTS.
  - KQTN `Q326D138139 - TBKQ\Edoc\KQ-…MTTN26020975.01 – .06` (tên mẫu chỉ ghi "Nước thải"), `.07` Mẫu QC (TSS).
  - Ảnh `Photo\NT.jpg`.

Quyết định của người dùng (18/11/2026):
  1. Tên khách: "Công ty TNHH Thương mại và Dịch vụ Nhất Tín" (tên in sẵn, không theo đề nghị đổi tên).
  2. Phiếu .01 – .06 ↔ NT1 – NT6 theo thứ tự: đúng.
  3. Giờ kết thúc: "để 11:10 và tô vàng".
  4. Ngày gửi mẫu (KQTN ghi nhận mẫu 03/11): "ghi 04/11 và tô vàng".
  5. Căn cứ / ngày yêu cầu cấp KQTN: "để nguyên và tô vàng" (Căn cứ "-" tô vàng, ngày yêu cầu "………").
Quyết định của người dùng (20/11/2026, lần 2): BCKQ "bỏ tô vàng hết, sửa ngày yêu cầu thành 20/11, đánh giá là không
  trễ, sửa hình sang chính giữa" -> bỏ highlight_bckq (giờ 11:10, ngày gửi mẫu 04/11 giữ số), req_date 20/11/2026,
  trễ hạn "không"; ảnh 1 tấm căn giữa (nuoc_common.photo_grid).
Quyết định của người dùng (20/11/2026): "chỉnh sửa lại tên công ty giúp tôi là: Công ty TNHH Công nghệ và Dịch vụ Nhất
  Tín" -> thay quyết định 1 ngày 18/11: tên khách (TBKQ + BCKQ) = "Công ty TNHH Công nghệ và Dịch vụ Nhất Tín".

Suy luận của AI (đã được người dùng xem 18/11/2026):
  1. Tên khách trên TBKQ theo đề nghị ghi trên BBLM: "Công ty TNHH Công nghệ và Dịch vụ Nhất Tín" (tô vàng).
  2. Phiếu .01 – .06 ↔ NT1 – NT6 theo thứ tự (phiếu không ghi ký hiệu mẫu).
  3. pH, nhiệt độ = trung bình 3 lần đo trên BBLM (số từng lần ghi ở BCKQ).
  4. Không có mức (QCVN 40:2025 chưa có cột / hệ số) -> Căn cứ "-", không đánh giá; giờ kết thúc để "………".
  5. Ngày yêu cầu cấp KQTN chưa có (không có PYC thử nghiệm) -> "………".

Chạy:  python build_138139.py [--pdf]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE, PLACEHOLDER as P

IN = os.path.join(WORKSPACE, "Input", "138139")
NEW_NAME = "Công ty TNHH Công nghệ và Dịch vụ Nhất Tín"
FIELD = [  # ký hiệu, giờ, pH TB, t° TB
    ("NT1", "10:20", "8,01", "32,7"), ("NT2", "10:30", "8,03", "32,8"), ("NT3", "10:40", "8,06", "32,7"),
    ("NT4", "10:50", "8,08", "33,0"), ("NT5", "11:00", "7,98", "33,0"), ("NT6", "11:10", "7,99", "33,2"),
]
MEMBERS = [dict(ky_hieu=k, kqtn_key=k, time=t, vi_tri="Nước thải sau hệ thống xử lý tại bể quan trắc", field={"pH": ph, "Nhiệt độ": tt}, field_tbkq=["pH", "Nhiệt độ"],
                photo=os.path.join(IN, "Photo", "NT.jpg") if k == "NT1" else None) for k, t, ph, tt in FIELD]

CONFIG = dict(
    code6="138139", loai="nuoc thai", kind="nt",
    cert_date=None,                                 # có KQTN -> ngày KQTN (AGENTS mục 2)
    client_vn=NEW_NAME,                             # người dùng 20/11/2026: đổi tên theo đề nghị trên BBLM
    address_vn="535/32 Quốc Lộ 13, phường Hiệp Bình, Thành phố Hồ Chí Minh",
    reg_date="29/10/2026", dates=["04/11/2026"], time_bckq="Ngày 04/11/2026 từ 10:15 đến 11:10",
    site_vn="Công ty TNHH may thêu Thuận Phương - Nhà máy Thuận Phương Long An - Ấp 3, xã Long Cang, tỉnh Tây Ninh",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.02, N2.TC.N.06",
    equip=["hanna_hi98190"], executor="Quân",
    witness_vn="anh Nguyễn Anh Tuấn – nhân viên",
    kph_lod=True,                                   # KPH -> giới hạn phát hiện, không "<" (AGENTS mục 13, 19/11/2026)
    req_date="20/11/2026",                          # người dùng 20/11/2026
    kqtn_map={f"MTTN26020975.0{i}": f"NT{i}" for i in range(1, 7)},
    field_devices={"pH": "Hanna HI 98190", "Nhiệt độ": "Hanna HI 98190"},
    send_date="04/11/2026",                        # hl_can_cu bỏ: người dùng 20/11/2026 "Bỏ tô vàng"
    tre_han=__import__("Core.Tools.kqtn_late", fromlist=["KHONG"]).KHONG,                  # người dùng 20/11/2026: "đánh giá là không trễ"
    bckq_notes=["- Số đo hiện trường ghi là trung bình 3 lần đo trên BBLM: NT1 pH 8,014/8,014/8,013, nhiệt độ 32,7/32,7/"
                "32,8 °C; NT2 8,034/8,033/8,034, 32,8 °C; NT3 8,063/8,063/8,062, 32,7 °C; NT4 8,078/8,077/8,077, 33,0 °C; "
                "NT5 7,979/7,980/7,980, 33,0 °C; NT6 7,995/7,994/7,995, 33,2 °C.",
                "- Ghi chú BBLM: đơn vị yêu cầu đề nghị thay đổi tên đơn vị yêu cầu thành \"Công ty TNHH Công nghệ và "
                "Dịch vụ Nhất Tín\"."],
    samples=[dict(kind="nt", ky_hieu="NT1 – NT6", kqtn_key="NT", members=MEMBERS, member_times=True, photo_cm=6.0, field_tbkq=["pH", "Nhiệt độ"],
                  params=[("pH", "-", "Đo tại hiện trường"), ("Nhiệt độ", "°C", "Đo tại hiện trường")],
                  vi_tri="Nước thải sau hệ thống xử lý tại bể quan trắc", time="10:20 – 11:10",
                  chi_tieu="pH, nhiệt độ, COD, amoni, TSS", equip=["hanna_hi98190"])],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf)
