# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D139353 - Công ty Cổ phần Dược phẩm Phenikaa (12/11/2026) — khí nén, 1 chứng thư, chỉ lấy mẫu tổng dầu
tại KN2. Bộ dựng chung: build_khinen_tong_dau.py.

Nguồn (Input/139353 — chép từ Z 16/11/2026, chỉ đọc):
  20261116172250646.pdf — tr. 1–2 BBLM BM3030.02: PYC Q326D139353 ngày 02/11/2026; địa chỉ "Lô đất CN2-10, Khu công
      nghiệp Công nghệ cao 2, Khu công nghệ cao Hòa Lạc, xã Hòa Lạc, Thành phố Hà Nội"; "Giám định chất lượng khí nén";
      địa điểm (viết tay) "Nhà máy Nghiên cứu và Sản xuất Dược phẩm Phenikaa - <địa chỉ trên>"; 10h30 – 15h30 ngày
      12/11/2026; bộ lấy mẫu tổng dầu (Pall, ORBO); N2.TC.KN.01; báo giá QONV26016982; vị trí 1 "Khu vực kỹ thuật nhà
      sấy 2 - Điểm xa nhất hệ thống khí nén - (Có lọc than hoạt tính)" lấy mẫu tổng dầu, 01, KN2; tr. 2 "Mã phòng:
      P1140", "Mã điểm: CA-P1140-01-I"; người thực hiện Nguyễn Tiến Phát.
      tr. 3 Ghi chép: hệ thống chạy, nhà xưởng không sản xuất (trạng thái nghỉ), "Máy Beko"; 7 bar / 0 bar; 35 / 1
      L/phút; 10h30 – 14h30; nhiệt độ khí 24,71 °C; Charcoal / Cassette KN2 (sửa từ KN1). tr. 4 trắng.
  Scan/Q326D139353 - Phenikaa - scan.pdf — PKSCV, PYC (MST 0103831711; báo giá QONV26016728 ngày 26/10/2026; "Khí nén
      sạch – 01 điểm – Đo tổng hàm lượng dầu khí nén – Iso 8573"; chứng thư Việt-Anh; ngày dự kiến trả KQ trống), điều
      kiện chung, báo giá.
  Kn2/BM00003–09.bmp — màn hình DP510 12/11/2026 (nhiệt độ khí 24,71 °C, khớp phiếu ghi chép).
  KN2.png — ảnh vị trí người dùng đánh dấu KN2 (Untitled Diagram.drawio: file người dùng đánh dấu ảnh KN1–KN3).

Quyết định của người dùng (16/11/2026, trả lời hỏi gộp):
  - Mô tả hệ thống khí nén: "Bỏ mục mô tả" (TBKQ bỏ đoạn; BCKQ ghi "Không có thông tin.").
  - Ảnh DP510: "Giữ chỉ tổng dầu" (chỉ lấy nhiệt độ khí cho N*, không mục điểm sương, thiết bị chỉ bộ lấy mẫu dầu).
  - Tên vị trí: "Giữ có mã" (mã phòng / mã điểm như BBLM tr. 2).
  - Chép Z: "Chờ KQTN".

Suy luận của AI (cần người dùng xác nhận):
  - Chỉ tiêu tổng dầu; thiết bị chỉ bộ lấy mẫu tổng dầu; DP510 chỉ để lấy nhiệt độ khí cho N*.
  - Tên vị trí ghép mã phòng / mã điểm (BBLM tr. 2) — cùng mã điểm với KN1 hồ sơ 139363 (khác ở lọc than hoạt tính).
  - Tên EN theo MST ("PHENIKAA PHARMACEUTICALS JOINT STOCK COMPANY").
  - Báo giá: BBLM QONV26016982, PYC QONV26016728.
  - Chưa có KQTN -> ………; ngày yêu cầu cấp KQTN = 20/11/2026.

Chạy:  python build_139353.py [--pdf]
Chép Z: python Core/Tools/ai_files.py chep 139353 --ghi
"""
import os
import argparse
from build_khinen_tong_dau import Case, run, ClientInfo, WORKSPACE

INP = os.path.join(WORKSPACE, "Input", "139353")
ADDR_VN = "Lô đất CN2-10, Khu công nghiệp Công nghệ cao 2, Khu công nghệ cao Hòa Lạc, xã Hòa Lạc, Thành phố Hà Nội"
ADDR_EN = "Land lot CN2-10, Hi-Tech Industrial Park 2, Hoa Lac Hi-Tech Park, Hoa Lac commune, Ha Noi City"

CASE = Case(
    code6="139353",
    client=ClientInfo(
        name_vn="Công ty Cổ phần Dược phẩm Phenikaa",
        name_en="Phenikaa Pharmaceuticals Joint Stock Company",
        address_vn=ADDR_VN, address_en=ADDR_EN,
    ),
    location_vn=f"Nhà máy Nghiên cứu và Sản xuất Dược phẩm Phenikaa - {ADDR_VN}",
    location_en=f"Phenikaa Pharmaceutical Research and Manufacturing Factory - {ADDR_EN}",
    sym="KN2",
    name_vn="Khu vực kỹ thuật nhà sấy 2 - Điểm xa nhất hệ thống khí nén (có lọc than hoạt tính) - "
            "Mã phòng: P1140, mã điểm: CA-P1140-01-I",
    name_en="Technical area of drying building 2 - Farthest point of the compressed air system (with activated "
            "carbon filter) - Room code: P1140, point code: CA-P1140-01-I",
    reg_date="02/11/2026", sample_date="12/11/2026", start="10:30", end="15:30",
    oil_from="10:30", oil_to="14:30", gas_temp=24.71,
    photo=os.path.join(INP, "KN2.png"),
    scan_new=os.path.join(INP, "20261116172250646.pdf"),
    scan_khac=os.path.join(INP, "Scan", "Q326D139353 - Phenikaa - scan.pdf"),
    kqtn=None,
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    run(CASE, pdf=ap.parse_args().pdf)
