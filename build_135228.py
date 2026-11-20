# -*- coding: utf-8 -*-
r"""
Q326D135228 — Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO — nước thải KCN Nhơn Trạch 5 (7 mẫu, 7 chứng thư).

Nguồn (chép từ Z `0.VIMCERTS078\\Q326D135228 - IDICO Nhơn trạch 5 - Nước thải` về Input\\135228, chỉ đọc):
  - `0. BBLM.pdf` (2 trang, BM3030.02): ngày 27/10/2026, 09:00 – 13:30, PYC Q326D135228 ngày 23/10/2026; địa điểm viết tay
    "IDICO – TSC đường N5, KCN Nhơn Trạch 5, phường Nhơn Trạch, thành phố Đồng Nai"; thiết bị in sẵn Hach DR300 (clo dư) +
    Hanna HI98190 (pH); quy trình N2.TC.N.06; người chỉ định: anh Bùi Thanh Hải – quản đốc nhà máy XLNT Nhơn Trạch 5;
    9 vị trí N1 – N9, đo nhanh N1 – N4 (pH, t°), giờ lấy mẫu từng mẫu; ghi chú: N1, N2 bỏ chỉ tiêu clo dư; N7, N9 không
    có nước. Người thực hiện: Nguyễn Hoàng Việt.
  - `4. edoc\\135228 - Danh sách vị trí + chỉ tiêu.pdf`: danh sách chỉ tiêu theo vị trí của KH.
  - KQTN `Q326D135228 - TBKQ\\Edoc\\KQ-Q326D135228-MTTN26020190.01–.06, .08` (N1 – N6, N8), `.10` (mẫu QC amoni),
    ngày 13/11/2026; phiếu ghi "Thời gian lấy mẫu 26/10/2026", ngày nhận mẫu 26/10/2026.
  - Ảnh `0. vi tri\\N1 … N8 - Vị trí lấy mẫu.jpg`.
  - Khung làm tay của người dùng trên Z (`2. BCKQ`, `3. TBKQ - 1…7`) — chỉ đối chiếu, không ghi đè.

Quyết định của người dùng (16/11/2026):
  - Thiết bị giữ theo BBLM (Hanna HI 98190), tô vàng; ngày lấy mẫu 27/10/2026 (tô vàng); địa điểm giữ "IDICO – TSC
    đường N5…" (tô vàng); nhiệt độ N4 = 30,4; chưa cần đánh giá; không ghi trễ hạn KQTN (để trống, tô vàng).

Suy luận của AI (cần người dùng xác nhận):
  1. Thiết bị theo BBLM: Hanna HI98190 (pH) cho N1 – N4; khung làm tay ghi "Máy đo pH Seven2Go S2, Mettler Toledo".
     N5, N6, N8 không đo hiện trường -> Thiết bị `-`.
  2. Ngày lấy mẫu: BBLM 27/10/2026, KQTN ghi lấy mẫu / nhận mẫu 26/10/2026 -> TBKQ theo BBLM; BCKQ ngày gửi mẫu theo KQTN.
  3. Địa điểm chép theo BBLM "IDICO – TSC đường N5, …" (chưa rõ "TSC" viết tắt của gì — giữ nguyên chữ BBLM).
  4. Tên vị trí theo BBLM (không thêm "Hố ga nước thải -" như khung làm tay).
  5. Nhiệt độ N4 đọc 30,4 °C (chữ viết tay, chưa chắc 30,4 / 30,5).
  6. BBLM ghi "theo QCVN 40:2011/BTNMT" nhưng không có mức / hệ số Kq, Kf -> không đánh giá, Căn cứ "-".
  7. Ngày yêu cầu cấp KQTN: không có PYC thử nghiệm -> tạm ngày lấy mẫu + 8 = 04/11/2026 (AGENTS mục 13).

Quyết định của người dùng (20/11/2026): "bỏ tô vàng tất cả, phần địa điểm … giữ từ KCN Nhơn Trạch 5 trở đi thôi, xóa đoạn
  IDICO – TSC đường N5 đi, thiết bị sử dụng làm giống BBLM" -> địa điểm "Khu công nghiệp Nhơn Trạch 5, phường Nhơn Trạch,
  Thành phố Đồng Nai"; thiết bị Hach DR300 + Hanna HI 98190 (cả 7 chứng thư, như BBLM); không tô vàng; ngày ban hành
  TBKQ 20/11/2026 (người dùng sửa trên Z).

Chạy:  python build_135228.py [--pdf] [--only 1,3]
"""
import argparse
import os
from nuoc_common import run, WORKSPACE

IN = os.path.join(WORKSPACE, "Input", "135228")
PH = lambda k: os.path.join(IN, "0. vi tri", f"{k} - Vị trí lấy mẫu.jpg")   # noqa: E731

CONFIG = dict(
    code6="135228", loai="nuoc thai", kind="nt",
    client_vn="Công ty TNHH MTV Dịch vụ Khu công nghiệp IDICO",
    address_vn="Số 48, khu phố 1, phường Long Hưng, Thành phố Đồng Nai",
    reg_date="23/10/2026", dates=["27/10/2026"], time_bckq="Ngày 27/10/2026 từ 09:00 đến 13:30",
    site_vn="Khu công nghiệp Nhơn Trạch 5, phường Nhơn Trạch, Thành phố Đồng Nai",
    scope_vn="Lấy mẫu và giám định chất lượng nước thải",
    procedure="QUATEST 3 3030:2025, N2.TC.N.06",
    equip=["hach_dr300", "hanna_hi98190"], executor="Việt",   # như BBLM (người dùng 20/11)
    witness_vn="anh Bùi Thanh Hải – quản đốc Nhà máy xử lý nước thải Khu công nghiệp Nhơn Trạch 5",
    req_date="04/11/2026",
    # Người dùng chốt 16/11/2026: giữ theo BBLM nhưng tô vàng để soát (thiết bị Hanna, ngày 27/10, địa điểm "IDICO – TSC");
    # không ghi trễ hạn KQTN (để trống, tô vàng)
    tre_han="………", no_highlight=True,                # người dùng 20/11/2026: bỏ tô vàng tất cả
    field_devices={"pH": "Hanna HI98190", "Nhiệt độ": "Hanna HI98190"},
    bckq_notes=["- Mẫu N7 (Nước thải Công ty TNHH Công nghiệp HW) và N9 (Công ty TNHH Jang In Furniture Việt Nam): "
                "tại thời điểm lấy mẫu không có nước thải, không lấy mẫu (theo BBLM).",
                "- Mẫu N1, N2: theo yêu cầu của anh Bùi Thanh Hải, bỏ chỉ tiêu clo dư (theo BBLM)."],
    samples=[
        dict(ky_hieu="135228-N1", kqtn_key="N1", time="13:12",
             vi_tri="Nước thải đầu vào Nhà máy xử lý nước thải Khu công nghiệp Nhơn Trạch 5",
             field={"pH": "7,6", "Nhiệt độ": "32,6"}, field_tbkq=["Nhiệt độ", "pH"], photo=PH("N1")),
        dict(ky_hieu="135228-N2", kqtn_key="N2", time="12:17",
             vi_tri="Nước thải đầu ra Nhà máy xử lý nước thải Khu công nghiệp Nhơn Trạch 5",
             field={"pH": "7,0", "Nhiệt độ": "32,7"}, field_tbkq=["Nhiệt độ", "pH"], photo=PH("N2")),
        dict(ky_hieu="135228-N3", kqtn_key="N3", time="10:30",
             vi_tri="Nước thải Công ty TNHH Da thuộc TaiYu",
             field={"pH": "8,0", "Nhiệt độ": "34,1"}, field_tbkq=["pH"], photo=PH("N3")),
        dict(ky_hieu="135228-N4", kqtn_key="N4", time="11:11",
             vi_tri="Nước thải Công ty Cổ phần Dệt Renze",
             field={"pH": "6,6", "Nhiệt độ": "30,4"}, field_tbkq=["pH"], photo=PH("N4")),
        dict(ky_hieu="135228-N5", kqtn_key="N5", time="09:05",
             vi_tri="Nước thải Công ty TNHH Hyosung Đồng Nai", photo=PH("N5")),
        dict(ky_hieu="135228-N6", kqtn_key="N6", time="09:30",
             vi_tri="Nước thải Công ty TNHH Hyosung Vina Industrial Machinery", photo=PH("N6")),
        dict(ky_hieu="135228-N8", kqtn_key="N8", time="11:40",
             vi_tri="Công ty Cổ phần Trung tâm gia công Posco Việt Nam – Chi nhánh Đồng Nai - VHPC", photo=PH("N8")),
    ],
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    run(CONFIG, pdf=a.pdf, only={int(x) for x in a.only.split(",") if x.strip()} or None)
