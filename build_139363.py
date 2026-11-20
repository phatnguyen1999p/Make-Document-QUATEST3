# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D139363 - Công ty TNHH Kỹ thuật Cơ điện lạnh Quí Long (12/11/2026) — khí nén, 1 chứng thư, chỉ lấy mẫu
tổng dầu tại KN1 (nhà máy Phenikaa Hòa Lạc). Bộ dựng chung: build_khinen_tong_dau.py.

Nguồn (Input/139363 — chép từ Z 16/11/2026, chỉ đọc):
  20261116172235607.pdf — tr. 1–2 BBLM BM3030.02: PYC Q326D139363 ngày 02/11/2026; địa chỉ "15 Lam Sơn, phường Đức
      Nhuận, Thành phố Hồ Chí Minh"; "Giám định chất lượng khí nén"; địa điểm (viết tay) "Nhà máy Nghiên cứu và Sản xuất
      Dược phẩm Phenikaa - Lô đất CN2-10, …, xã Hòa Lạc, Thành phố Hà Nội"; 10h30 – 15h30 ngày 12/11/2026; thiết bị bộ
      lấy mẫu tổng dầu (Pall, ORBO); N2.TC.KN.01; báo giá QONV26016980; vị trí 1 "Khu vực kỹ thuật nhà sấy 2 - Điểm xa
      nhất hệ thống khí nén - (Không lọc than hoạt tính)" lấy mẫu tổng dầu, 01, KN1; tr. 2 "Mã phòng: P1140", "Mã điểm:
      CA-P1140-01-I"; người thực hiện Nguyễn Tiến Phát.
      tr. 3 Ghi chép BM3030.01: hệ thống khí nén chạy, nhà xưởng không sản xuất (trạng thái nghỉ), ghi chú "Máy Beko";
      7 bar / lấy mẫu 0 bar; 35 / 1 L/phút; 10h30 – 14h30; nhiệt độ khí 24,87 °C; Charcoal / Cassette KN1. tr. 4 trắng.
  Scan/Q326D139363 - Quí Long - scan.pdf — PKSCV, PYC (MST 0304906371; báo giá QONV26016728 ngày 26/10/2026;
      "Khí nén sạch – 01 điểm – Đo tổng hàm lượng dầu khí nén – Iso 8573"; chứng thư Việt-Anh; "Ngày dự kiến trả kết
      quả" trống), điều kiện chung, báo giá.
  KN1/BM00000–02.bmp — màn hình DP510 12/11/2026 (nhiệt độ khí 24,80 – 24,87 °C, khớp phiếu ghi chép).
  KN1.png — ảnh vị trí người dùng đã đánh dấu KN1.

Quyết định của người dùng (16/11/2026, trả lời hỏi gộp):
  - Mô tả hệ thống khí nén: "Bỏ mục mô tả" (TBKQ bỏ đoạn; BCKQ ghi "Không có thông tin.").
  - Ảnh DP510: "Giữ chỉ tổng dầu" (chỉ lấy nhiệt độ khí cho N*, không mục điểm sương, thiết bị chỉ bộ lấy mẫu dầu).
  - Tên vị trí: "Giữ có mã" (mã phòng / mã điểm như BBLM tr. 2).
  - Chép Z: "Chờ KQTN".

Suy luận của AI (cần người dùng xác nhận):
  - Chỉ dựng chỉ tiêu tổng dầu (BBLM, PYC); ảnh DP510 chỉ dùng lấy nhiệt độ khí cho N*, không lập mục điểm sương,
    thiết bị chỉ ghi bộ lấy mẫu tổng dầu (BBLM mục 7).
  - Tên vị trí ghép mã phòng / mã điểm ở BBLM tr. 2; "nhà sấy 2" đọc từ chữ viết tay.
  - Tên EN theo MST (tên quốc tế "QUI LONG REFRIGERATION ELECTRICAL ENGINEERING TECHNOLOGY CO., LTD").
  - Báo giá: BBLM QONV26016980, PYC QONV26016728 (trùng số báo giá trên PYC Phenikaa 139353).
  - Chưa có KQTN -> kết quả tổng dầu, ngày ban hành để ……… ; ngày yêu cầu cấp KQTN = 12/11 + 8 = 20/11/2026.

Chạy:  python build_139363.py          (DOCX + XLSX + Edoc)
       python build_139363.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu)
Chép Z: python Core/Tools/ai_files.py chep 139363 --ghi
"""
import os
import argparse
from build_khinen_tong_dau import Case, run, ClientInfo, WORKSPACE

INP = os.path.join(WORKSPACE, "Input", "139363")
FACTORY_VN = ("Nhà máy Nghiên cứu và Sản xuất Dược phẩm Phenikaa - Lô đất CN2-10, Khu công nghiệp Công nghệ cao 2, "
              "Khu công nghệ cao Hòa Lạc, xã Hòa Lạc, Thành phố Hà Nội")
FACTORY_EN = ("Phenikaa Pharmaceutical Research and Manufacturing Factory - Land lot CN2-10, Hi-Tech Industrial Park 2, "
              "Hoa Lac Hi-Tech Park, Hoa Lac commune, Ha Noi City")

CASE = Case(
    code6="139363",
    client=ClientInfo(
        name_vn="Công ty TNHH Kỹ thuật Cơ điện lạnh Quí Long",
        name_en="Qui Long Refrigeration Electrical Engineering Technology Co., Ltd",
        address_vn="15 Lam Sơn, phường Đức Nhuận, Thành phố Hồ Chí Minh",
        address_en="15 Lam Son, Duc Nhuan ward, Ho Chi Minh City",
    ),
    location_vn=FACTORY_VN, location_en=FACTORY_EN,
    sym="KN1",
    name_vn="Khu vực kỹ thuật nhà sấy 2 - Điểm xa nhất hệ thống khí nén (không lọc than hoạt tính) - "
            "Mã phòng: P1140, mã điểm: CA-P1140-01-I",
    name_en="Technical area of drying building 2 - Farthest point of the compressed air system (without activated "
            "carbon filter) - Room code: P1140, point code: CA-P1140-01-I",
    reg_date="02/11/2026", sample_date="12/11/2026", start="10:30", end="15:30",
    oil_from="10:30", oil_to="14:30", gas_temp=24.87,
    photo=os.path.join(INP, "KN1.png"),
    scan_new=os.path.join(INP, "20261116172235607.pdf"),
    scan_khac=os.path.join(INP, "Scan", "Q326D139363 - Quí Long - scan.pdf"),
    kqtn=None,
)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    run(CASE, pdf=ap.parse_args().pdf)
