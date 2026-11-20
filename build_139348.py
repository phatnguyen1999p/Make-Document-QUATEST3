# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D139348 - Công ty Cổ phần Công nghiệp PSI Việt Nam (12/11/2026) — khí nén, 1 chứng thư, chỉ lấy mẫu tổng
dầu tại KN3 (nhà máy Phenikaa Hòa Lạc). Bộ dựng chung: build_khinen_tong_dau.py.

Nguồn (Input/139348 — chép từ Z 16/11/2026, chỉ đọc):
  20261116172223146.pdf — tr. 1–2 BBLM BM3030.02: PYC Q326D139348 ngày 02/11/2026; địa chỉ "Ô 1A, lô 3, Cụm công nghiệp
      Lai Xá, xã Hoài Đức, Thành phố Hà Nội"; "Giám định chất lượng khí nén"; địa điểm (viết tay) "Nhà máy Nghiên cứu và
      Sản xuất Dược phẩm Phenikaa - Lô đất CN2-10, …, xã Hòa Lạc, Thành phố Hà Nội"; 10h30 – 15h30 ngày 12/11/2026; bộ
      lấy mẫu tổng dầu (Pall, ORBO); N2.TC.KN.01; báo giá QONV26016988; vị trí 1 "Phòng máy sấy hấp thụ - Sau bộ lọc
      cuối" lấy mẫu tổng dầu, 01, KN3; tr. 2 "Mã phòng: P1148", "Mã điểm: CA-P1216-02-I"; người thực hiện Nguyễn Tiến Phát.
      tr. 3 Ghi chép: hệ thống chạy, nhà xưởng không sản xuất (trạng thái nghỉ), "Máy Beko"; 7 bar / 0 bar; 35 / 1
      L/phút; 11h00 – 15h00; nhiệt độ khí 24,16 °C; Charcoal / Cassette KN3. tr. 4 trắng.
  Scan/Q326D139348 - PSI Việt Nam - scan.pdf — PKSCV, PYC (MST 0106088879; báo giá QONV26016988 ngày 30/10/2026; "Khí nén
      sạch – 01 điểm – Đo tổng hàm lượng dầu khí nén – Iso 8573"; chứng thư Việt-Anh; ngày dự kiến trả KQ trống), điều
      kiện chung, báo giá.
  KN3/BM00010–20.bmp — màn hình DP510 12/11/2026 (nhiệt độ khí 24,04 – 24,32 °C; phiếu ghi 24,16 °C).
  KN3.png — ảnh vị trí người dùng đánh dấu KN3.

Quyết định của người dùng (16/11/2026, trả lời hỏi gộp):
  - Mô tả hệ thống khí nén: "Bỏ mục mô tả" (TBKQ bỏ đoạn; BCKQ ghi "Không có thông tin.").
  - Ảnh DP510: "Giữ chỉ tổng dầu" (chỉ lấy nhiệt độ khí cho N*, không mục điểm sương, thiết bị chỉ bộ lấy mẫu dầu).
  - Tên vị trí: "Giữ có mã" (mã phòng / mã điểm như BBLM tr. 2).
  - Chép Z: "Chờ KQTN".

Suy luận của AI (cần người dùng xác nhận):
  - Chỉ tiêu tổng dầu; thiết bị chỉ bộ lấy mẫu tổng dầu; DP510 chỉ để lấy nhiệt độ khí cho N*.
  - Tên vị trí ghép mã phòng / mã điểm (BBLM tr. 2): mã phòng P1148 nhưng mã điểm ghi "CA-P1216-02-I" (khác số phòng)
    — giữ nguyên như chữ viết tay.
  - Tên EN theo MST ("PSI VIET NAM INDUSTRY JOINT STOCK COMPANY").
  - Chưa có KQTN -> ………; ngày yêu cầu cấp KQTN = 20/11/2026.

Chạy:  python build_139348.py [--pdf]
Chép Z: python Core/Tools/ai_files.py chep 139348 --ghi
"""
import os
import argparse
from build_khinen_tong_dau import Case, run, ClientInfo, WORKSPACE

INP = os.path.join(WORKSPACE, "Input", "139348")
FACTORY_VN = ("Nhà máy Nghiên cứu và Sản xuất Dược phẩm Phenikaa - Lô đất CN2-10, Khu công nghiệp Công nghệ cao 2, "
              "Khu công nghệ cao Hòa Lạc, xã Hòa Lạc, Thành phố Hà Nội")
FACTORY_EN = ("Phenikaa Pharmaceutical Research and Manufacturing Factory - Land lot CN2-10, Hi-Tech Industrial Park 2, "
              "Hoa Lac Hi-Tech Park, Hoa Lac commune, Ha Noi City")

CASE = Case(
    code6="139348",
    client=ClientInfo(
        name_vn="Công ty Cổ phần Công nghiệp PSI Việt Nam",
        name_en="PSI Viet Nam Industry Joint Stock Company",
        address_vn="Ô 1A, lô 3, Cụm công nghiệp Lai Xá, xã Hoài Đức, Thành phố Hà Nội",
        address_en="Plot 1A, lot 3, Lai Xa Industrial Cluster, Hoai Duc commune, Ha Noi City",
    ),
    location_vn=FACTORY_VN, location_en=FACTORY_EN,
    sym="KN3",
    name_vn="Phòng máy sấy hấp thụ - Sau bộ lọc cuối - Mã phòng: P1148, mã điểm: CA-P1216-02-I",
    name_en="Adsorption dryer room - After the final filter - Room code: P1148, point code: CA-P1216-02-I",
    reg_date="02/11/2026", sample_date="12/11/2026", start="10:30", end="15:30",
    oil_from="11:00", oil_to="15:00", gas_temp=24.16,
    photo=os.path.join(INP, "KN3.png"),
    scan_new=os.path.join(INP, "20261116172223146.pdf"),
    scan_khac=os.path.join(INP, "Scan", "Q326D139348 - PSI Việt Nam - scan.pdf"),
    kqtn=None,
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    run(CASE, pdf=ap.parse_args().pdf)
