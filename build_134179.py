# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D134179 - Công ty TNHH Duy Tân Long An (03/11/2026) — 2 chứng thư (số theo phiếu ghi chép):
  /1 Khí nén: 01 vị trí KN1 (ống dẫn khí nén sau lọc tại khu vực máy nén khí) — tổng dầu, điểm sương, nồng độ hạt
     (AeroTrak 9350 + HPD III), vi sinh khí nén (SAS Pinocchio Super II: tổng số VSV hiếu khí, E. coli, Salmonella).
  /2 Vi sinh: không khí K1–K6 (Buck B30120, 100 L/phút x 1 phút, tổng số VSV hiếu khí) + bề mặt BM1–BM14
     (swab, 10 mL BPW: tổng số VSV hiếu khí, Coliforms), blank BM15.

Nguồn (Input/134179/_Z — chép nguyên thư mục hồ sơ trên Z 13/11/2026, chỉ đọc):
  0. BBLM.pdf — BM3030.02 2 trang: PYC Q326D134179 ngày 22/10/2026; thực hiện 03/11/2026 08:45–13:15; quy trình
      N2.TC.KN.01, N2.TC.09; nhà máy hoạt động sản xuất; K1–K6, BM1 đến BM14, KN; "Hệ khí nén … gồm 01 máy nén khí
      Belliss Marcom có dùng dầu bôi trơn. Theo yêu cầu Công ty, không thể hiện sơ đồ hệ thống khí nén tại TBKQ";
      người thực hiện Phạm Quốc Nghĩa.
  0. Ghi chep khi nen .pdf — BM3030.01 số Q326D134179/1, 3 trang: tổng dầu (hệ thống ~6,3 bar, lấy mẫu ~3 bar, hạt
      35 L/phút, hơi 1,0 L/phút, 9h00–13h00, khí ~33 °C); particle (HPD ~4 bar, 50 L/phút, 20 phút, zero 1250, đo
      1251); vi sinh (100 L/phút, 05 phút/đĩa, mẫu KN1, blank KN2: TS, E. coli, Sal); điểm sương (file máy CS 510);
      mô tả hệ thống: máy nén Belliss Marcom → bình chứa → máy sấy Gardner Denver → 03 lọc Belliss Marcom nối tiếp
      → vị trí lấy mẫu (máy nén có dùng dầu).
  0. Ghi chep vi sinh .pdf — BM3030.01 số Q326D134179/2, 3 trang: APBuck 100 L/phút, 1 phút, K1–K6; swab 1–14,
      15 = blank ("Diện tích quét" để trống).
  1.  TinhToanKhiNen.xlsx — phiếu người dùng nhập 13/11/2026 (AeroTrak Δ, điểm sương −1,82/−1,73 °C, 32,8/32,7 °C,
      6,29/6,25 bar, quy đổi 7 bar −0,7 °C; dầu hạt "< 10", dầu hơi "< 5" µg/mẫu; vi sinh "< 1").
  1.  Tinh Toan VSKK.xlsx — phiếu người dùng: số Q326D134179/KK ngày 10/11/2026, hiệu chỉnh bảng Conversion Buck
      (380 lỗ): K1–K6 = 27, 15, 10, 23, 11, 33 CFU/đĩa -> 280, 150, 100, 240, 110, 340 CFU/m³.
  THUC HIEN -/D134179 - KQ DO BUI/D134179 - DIF .pdf — AeroTrak 9350-02 S/N 93502239001, Count Mode Δ, # / m³:
      1250 zero (5 phút), 1251 đo 08:58:07, 20 phút, 1000 L: 198297 / 35128 / 8991 / 5184 / 50 / 3.
  THUC HIEN -/D134179 - DIEM SUONG/D261103/BM00007 - DO 1 (11:41:17), BM00020 - DO 2 (12:06:43) — màn hình DP510.
  THUC HIEN -/D134179 - PHOTO/D134179 - photo - BAO CAO — 7 ảnh người dùng chọn (1 khí nén, 6 vi sinh).
  Q326D134179 - KQTN: HDTN26013748.01.1 (10/11/2026, dầu hạt KPH), .01.2 (định tính dầu chuẩn);
      VSTN26027754.01–.06 (09/11/2026, K1–K6), .08 KN1, .09 KN2 (BLANK); VSTN26027755.01–.15 (11/11/2026,
      BM1–BM15; BM15 = blank).
  Mẫu trình bày: TBKQ / BCKQ hồ sơ Q326D103230 người dùng để sẵn trên Z ("2. BCKQTĐ…", "3. TBKQ…"), TBKQ 901
      (11/2025, cùng khách).

Quyết định của người dùng (13/11/2026):
  - "chạy xong khoan chép lên Z" — KHÔNG chép lên Z cho tới khi người dùng bảo.
  - Tin nhắn giao việc (người dùng chuyển 13/11/2026): "điểm sương + VS khí nén + bụi, a làm excel tính rồi, còn chờ
    KQ dầu hơi thôi, KQVS không khí + bề mặt có hết rồi" -> điểm sương, bụi, vi sinh khí nén lấy theo phiếu Excel;
    CHƯA có KQTN dầu hơi: tổng dầu và ngày ban hành /1 để "………" tô vàng (OIL_PENDING); /2 chốt theo KQTN.
  - "ok, chép lên Z, các nội dung do Claude làm chép vào 1 thư mục là "Claude - work"" -> file AI đặt trong
    <thư mục hồ sơ trên Z>\\Claude - work\\{ThucHien, TBKQ, pdf}: python Core/Tools/ai_files.py chep 134179 --z "<…>\\Claude - work" --ghi

Quyết định của người dùng (14/11/2026): "Tham khảo các nội dung trong hồ sơ này [TBKQ - 1 Q326D124796, K Pack — vi
  sinh không khí + swab], và thực hiện dựng lại hồ sơ Honchuan, Duy Tan" -> bộ /2 vi sinh dựng theo khuôn K Pack:
  9.1 = 3 dòng dẫn + bảng "Vị trí | Ký hiệu vị trí | Nội dung thực hiện" + ảnh gom theo nhóm có dòng chú thích nhóm;
  9.2 = bảng kết quả chỉ ghi ký hiệu vị trí, đơn vị ở tiêu đề; Ghi chú theo thứ tự "< 10" -> VILAS -> VIAS.
  Giữ theo AGENTS (khác bản K Pack): tên máy chuẩn, "QUATEST 3 3030:2025, N2.TC.09", kết quả ghi như KQTN.
  Suy luận của AI khi áp khuôn (chờ người dùng xác nhận): chú thích nhóm ảnh và thứ tự ghi chú lấy theo K Pack (lệch
  AGENTS mục 3); mẫu bề mặt để một nhóm "trên bề mặt" (K Pack tách thiết bị / tay công nhân vì khác chỉ tiêu).
  Chỉ dựng lại bộ /2: python build_134179.py --chi-vi-sinh --pdf (bộ /1 khí nén và 2 phiếu Excel giữ nguyên).

Suy luận của AI (cần người dùng xác nhận):
  - Chỉ dựng TBKQ tiếng Việt (không có PYC trong thư mục; TBKQ 901 năm trước và file mẫu 103230 đều chỉ tiếng Việt).
  - Ký hiệu khí nén ghi KN1 (BBLM ghi "KN" — AGENTS mục 6: vị trí đầu ghi KN1).
  - Dầu hạt KPH -> "< 10" µg/mẫu như phiếu người dùng (KQTN không in LOD). Khi có phiếu MTTN: điền KQTN_DAU_HOI,
    OIL_VAPOR rồi chạy lại (ngày ban hành /1 = ngày KQTN muộn nhất của bộ khí nén).
  - Ngày ban hành /2 = 11/11/2026 (phiếu KQTN muộn nhất của bộ vi sinh: VSTN26027755).
  - HDTN26013748 ghi "Ngày nhận mẫu 25/10/2026" (trước ngày lấy mẫu 03/11/2026) — ghi nguyên theo phiếu, chờ quyết.
  - Ngày yêu cầu cấp KQTN = ngày lấy mẫu + 8 = 11/11/2026 (không có PYC — AGENTS mục 13).
  - Diện tích swab 100 cm² (phiếu ghi chép để trống; theo TBKQ 901) -> đơn vị CFU/ 100 cm².
  - Tên vị trí swab 1–5 đọc là "4000-35" (chữ viết tay có thể là "400-35"); TBKQ không ghi tên nhân viên
    (ghi chép: "Ng Khải Tâm", "Đinh Khánh Trung" — đọc chưa chắc), BCKQ ghi kèm.
  - "Belliss Marcom" ghi theo BBLM / ghi chép (tên hãng thường viết "Belliss & Morcom").
  - Nhiệt độ khí để tính N* = 32,8 °C; điểm sương quy đổi 7 bar = −0,7 °C theo phiếu người dùng nhập.
  - Địa điểm giám định: BBLM ghi "tại địa chỉ nêu tại mục 3" -> "Công ty TNHH Duy Tân Long An - <địa chỉ>".
  - Tên tiếng Anh khách hàng do AI dịch (chưa có MST để tra) — chỉ dùng trong dữ liệu, không in trên TBKQ VN.

Chạy:  python build_134179.py          (DOCX + XLSX + Edoc)
       python build_134179.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: trong code (network_dir để trống) — chép lên Z theo AGENTS mục 0 KHI NGƯỜI DÙNG BẢO.
"""

import os
import re
import sys
import copy
import argparse

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_ton_tai_Z"))

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm
from docx.text.paragraph import Paragraph

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.handlers.compressed_air import CompressedAirHandler
from Core.Builders.tag_context import PLACEHOLDER, style_note_tables, appendix_total_field
from Core.Styling.elements import add_document_ending_line
from Core.Styling.styles import set_cell_shading
from Core.Styling.tables import (format_table_structure, set_table_autofit_window, insert_image_to_cell,
                                 add_tbl_header, add_cant_split, PHOTO_WIDTH_3COL)
from Core.Tools.khinen_calc import build_compressed_air_workbook, create_dossier_from_compressed_air_workbooks
from Core.Tools.cleanroom_calc import CleanroomCalcBook, LEFT
from Core.Tools.excel_borders import thin_outside_borders, strike_box
from Core.Tools.kqtn_late import tre_han_text
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D134179"
CODE6 = "134179"
REG_DATE = "22/10/2026"            # BBLM mục 1 (PYC)
SAMPLE_DATE = "03/11/2026"         # BBLM, ghi chép
START, END = "08:45", "13:15"      # BBLM mục 6
CERT_DATE = "11/11/2026"           # phiếu KQTN muộn nhất hiện có: VSTN26027755 ngày 11/11/2026 (AGENTS mục 2)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SRC = os.path.join(INPUT_DIR, "_Z")                       # bản chép thư mục hồ sơ trên Z (chỉ đọc)
SRC_TH = os.path.join(SRC, "THUC HIEN -")
PHOTO_DIR = os.path.join(SRC_TH, "D134179 - PHOTO", "D134179 - photo - BAO CAO")   # ảnh người dùng chọn
DEW_DIR = os.path.join(SRC_TH, "D134179 - DIEM SUONG", "D261103")
SCAN_BBLM = os.path.join(SRC, "0. BBLM.pdf")
SCAN_GC_KN = os.path.join(SRC, "0. Ghi chep khi nen .pdf")
SCAN_GC_VS = os.path.join(SRC, "0. Ghi chep vi sinh .pdf")
SCAN_DUST = os.path.join(SRC_TH, "D134179 - KQ DO BUI", "D134179 - DIF .pdf")
USER_VSKK = os.path.join(SRC, "1.  Tinh Toan VSKK.xlsx")
NGAY_YC_KQTN = "11/11/2026"        # TẠM: ngày lấy mẫu + 8 ngày (không có PYC — AGENTS mục 13)

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Duy Tân Long An",
    name_en="Duy Tan Long An Company Limited",          # AI dịch — chưa có MST để tra tên quốc tế
    address_vn="A18, Tân Đô Hải Sơn, Khu công nghiệp Tân Đô, xã Đức Hòa, tỉnh Tây Ninh",      # BBLM mục 3
    address_en="A18, Tan Do Hai Son, Tan Do Industrial Park, Duc Hoa commune, Tay Ninh province",
)
LOCATION_VN = f"{CLIENT.name_vn} - {CLIENT.address_vn}"   # BBLM mục 5: "tại địa chỉ nêu tại mục 3"
LOCATION_EN = f"{CLIENT.name_en} - {CLIENT.address_en}"
EXECUTOR = "Nghĩa"                 # BBLM: Phạm Quốc Nghĩa -> GĐV ký Phạm Quốc Nghĩa; người báo cáo Nghĩa + Minh

# ---------------------------------------------------------------------------
# /1 Khí nén
# ---------------------------------------------------------------------------
KN_SYM = "KN1"
KN_NAME_VN = "Ống dẫn khí nén sau lọc tại khu vực máy nén khí"          # BBLM tr. 2
KN_NAME_EN = "Compressed air pipeline after filters at the air compressor area"
KN_PHOTO = os.path.join(PHOTO_DIR, "VI TRI KN  .jpg")
KN_GAS_TEMP = 32.8                 # °C — phiếu người dùng nhập (N*); màn hình DP510 32,84 / 32,77
KN_T2_7BAR = -0.7                  # °C — điểm sương quy đổi 7 bar theo phiếu người dùng nhập (CS Humidity Calculator)
KN_DUST = [198297, 35128, 8991, 5184, 50, 3]   # AeroTrak file 1251, Δ hạt/m³: 0,3 0,5 1 2 3 5 µm (1000 L, 20 phút)
KN_DEW = {"dewp_start": -1.82, "dewp_end": -1.73, "temp_start": 32.8, "temp_end": 32.7,
          "pres_start": 6.29, "pres_end": 6.25, "time_start": "Lúc 11h41", "time_end": "Lúc 12h06",
          "photo_start": os.path.join(DEW_DIR, "BM00007 - DO 1 .jpg"),
          "photo_end": os.path.join(DEW_DIR, "BM00020 - DO 2 .jpg")}
KQTN_DAU_HAT = {"so": "HDTN26013748", "phieu": "HDTN26013748.01.1", "ngay": "10/11/2026", "ngay_nhan_mau": "25/10/2026"}
KQTN_DAU_HOI = {"so": "", "phieu": "", "ngay": "", "ngay_nhan_mau": ""}      # CHƯA CÓ phiếu MTTN trong thư mục KQTN
KQTN_VS_KN = {"so": "VSTN26027754", "phieu": "VSTN26027754.08", "blank": "VSTN26027754.09",
              "ngay": "09/11/2026", "ngay_nhan_mau": "03/11/2026"}
OIL_AEROSOL, OIL_VAPOR = "< 10", "< 5"      # µg/mẫu — phiếu người dùng nhập (dầu hạt KPH; dầu hơi chờ phiếu MTTN)
OIL_PENDING = not KQTN_DAU_HOI["so"]        # chưa có KQTN dầu hơi -> tổng dầu, ngày ban hành /1 để '………' tô vàng
CERT_DATE_KN = PLACEHOLDER if OIL_PENDING else CERT_DATE
# Vi sinh khí nén (VSTN26027754.08, mẫu KN1, 3 đĩa, 100 L/phút x 5 phút = 500 L): (chỉ tiêu, KQTN, kết quả TBKQ)
KN_MICRO = [("Tổng số vi sinh vật hiếu khí", "< 1 CFU/đĩa", "Không có chỉ thị"),
            ("Escherichia coli", "< 1 CFU/đĩa", "Không có chỉ thị"),
            ("Salmonella spp.", "KPH/mẫu", "Không có chỉ thị")]
SYSTEM_CHAIN_VN = ["Máy nén khí Belliss Marcom", "Bình chứa", "Máy sấy khí Gardner Denver",
                   "03 bộ lọc Belliss Marcom nối tiếp", "Vị trí lấy mẫu"]          # ghi chép khí nén tr. 3
SYSTEM_CHAIN_EN = ["Belliss Marcom air compressor", "Air receiver", "Gardner Denver air dryer",
                   "03 Belliss Marcom filters in series", "Sampling location"]

# ---------------------------------------------------------------------------
# /2 Vi sinh không khí + bề mặt
# ---------------------------------------------------------------------------
SO_VSKK, NGAY_VSKK = f"{CODE}/KK", "10/11/2026"      # số, ngày phiếu tính toán theo phiếu người dùng
VS_FLOW_LPM, VS_MINUTES, BUCK_HOLES = 100, 1, 380    # ghi chép; bảng Conversion Buck B30120 (phiếu người dùng)
KQTN_VSKK = {"so": "VSTN26027754", "ngay": "09/11/2026", "ngay_nhan_mau": "03/11/2026"}
KQTN_VSBM = {"so": "VSTN26027755", "ngay": "11/11/2026", "ngay_nhan_mau": "03/11/2026"}
AIR = [  # (ký hiệu, vị trí — BBLM + ghi chép, CFU/đĩa theo VSTN26027754.01–.06)
    ("K1", "Khu vực máy Siapi 4", 27),
    ("K2", "Khu vực máy 2000-10", 15),
    ("K3", "Khu vực máy 4000-12", 10),
    ("K4", "Khu vực máy 450-22", 23),
    ("K5", "Khu vực máy 450-17", 11),
    ("K6", "Khu vực máy ASB1-2", 33),
]
SWAB = [  # (số mẫu ghi chép, vị trí TBKQ, tên người trên ghi chép (chỉ BCKQ), TSVSVHK, Coliforms) — VSTN26027755.01–.14
    (1, "Máng hứng sản phẩm 4000-35", "", "< 10", "< 10"),
    (2, "Tay gắp sản phẩm 4000-35", "", "< 10", "< 10"),
    (3, "Gá phôi sản phẩm 4000-35", "", "< 10", "< 10"),
    (4, "Bao bì thành phẩm 4000-35", "", "< 10", "< 10"),
    (5, "Băng tải máy thổi 4000-35", "", "2,0 × 10¹", "< 10"),
    (6, "Găng tay nhân viên xưởng PET", "Ng Khải Tâm", "5,8 × 10²", "< 10"),
    (7, "Áo nhân viên xưởng PET", "Ng Khải Tâm", "< 10", "< 10"),
    (8, "Tay nhân viên xưởng PET", "Ng Khải Tâm", "2,7 × 10²", "< 10"),
    (9, "Áo nhân viên kỹ thuật PET (cổ áo)", "Đinh Khánh Trung", "1,9 × 10³", "< 10"),
    (10, "Tay nhân viên kỹ thuật PET", "Đinh Khánh Trung", "< 10", "< 10"),
    (11, "Thùng đựng bán thành phẩm 450-9", "", "< 10", "< 10"),
    (12, "Băng tải hứng phôi 450-9", "", "< 10", "< 10"),
    (13, "Máng hứng bán thành phẩm 450-9", "", "< 10", "< 10"),
    (14, "Tay robot 450-16", "", "< 10", "< 10"),
]
SWAB_BLANK = 15                    # BM15: < 10 CFU/mẫu cả 2 chỉ tiêu (VSTN26027755.15)
UNIT_BM = "CFU/ 100 cm²"
VS_PHOTO_GROUPS = [   # ảnh gom theo nhóm, mỗi nhóm một dòng chú thích — theo TBKQ mẫu Q326D124796 (K Pack)
    ("Lấy mẫu vi sinh không khí", ["khong khi  .jpg", "khong khi .jpg"]),
    ("Lấy mẫu vi sinh trên bề mặt", ["be mat - vi tri  .jpg", "vi tri  -  .jpg", "co ao -  .jpg", "tay NV  .jpg"]),
]
VS_CONTENT_AIR = "Lấy mẫu và xác định tổng số vi sinh vật hiếu khí trong không khí"
VS_CONTENT_BM = "Lấy mẫu và xác định tổng số vi sinh vật hiếu khí và Coliforms trên bề mặt"
NOTE_VILAS = "Phương pháp thử nghiệm được công nhận ISO 17025 - VILAS 004."
NOTE_VIAS = "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004."
BM_ROW_TWIPS = 440                 # chiều cao hàng bảng kết quả bề mặt (chỉnh theo PDF để bảng + Ghi chú trọn trang)
BCKQ_PHOTO_W2, BCKQ_PHOTO_W3 = Cm(7.0), Cm(4.6)   # ảnh BCKQ: lưới 2 cột / 3 cột
EQUIP_SWAB_VN = "Tăm bông tiệt trùng và dung dịch đệm pepton (BPW)"
EQUIP_SWAB_EN = "Sterile swab and buffered peptone water (BPW)"


def buck_corrected(r: int) -> int:
    """Số khuẩn lạc sau hiệu chỉnh theo bảng Conversion Buck B30120 (N lỗ) — như phiếu người dùng:
    Pr = ROUND(N × Σ 1/(N - i + 1), 0), tổng cộng dồn làm tròn 6 số lẻ."""
    s = 0.0
    for i in range(1, r + 1):
        s = round(s + 1 / (BUCK_HOLES - i + 1), 6)
    return round(s * BUCK_HOLES)


def air_value(cfu: int) -> int:
    return round(buck_corrected(cfu) / (VS_FLOW_LPM * VS_MINUTES / 1000.0))


def _check_user_vskk():
    """Đối chiếu số tự tính với phiếu người dùng (cột P = CFU/m³) — lệch thì dừng (AGENTS mục 11)."""
    ws = openpyxl.load_workbook(USER_VSKK, data_only=True)["Sheet1"]
    for i, (sym, _, cfu) in enumerate(AIR):
        r = 8 + i
        if str(ws[f"A{r}"].value).strip() != sym or ws[f"B{r}"].value != cfu or ws[f"P{r}"].value != air_value(cfu):
            raise SystemExit(f"[FATAL] [Data Mismatch] VSKK {sym}: phiếu người dùng {ws[f'B{r}'].value} CFU/đĩa -> "
                             f"{ws[f'P{r}'].value} CFU/m³, code {cfu} -> {air_value(cfu)}")


# ---------------------------------------------------------------------------
# Tiện ích DOCX (theo build_137035 / build_137058)
# ---------------------------------------------------------------------------
def _optimized(path: str, max_w: int = 1200) -> str:
    """Ảnh nén trong Input/134179/Photo/_optimized (AGENTS mục 1, 3): xoay EXIF, ảnh dọc xoay 90°, không crop."""
    from PIL import Image, ImageOps
    out_dir = os.path.join(INPUT_DIR, "Photo", "_optimized")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0].replace(" ", "_") + ".jpg")
    im = ImageOps.exif_transpose(Image.open(path))
    if im.height > im.width:
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=88)
    return out


def _ptext(el):
    import unicodedata          # template BCKQ dùng dấu tổ hợp (NFD) -> so chữ theo NFC
    return unicodedata.normalize("NFC", "".join(t.text or "" for t in el.iter(qn("w:t"))))


def _hl(run):
    rpr = run._r.get_or_add_rPr()
    if rpr.find(qn("w:highlight")) is None:
        h = OxmlElement("w:highlight")
        h.set(qn("w:val"), "yellow")
        rpr.append(h)


def add_runs(p, text, size, bold=False, italic=False):
    """Ghi chữ vào đoạn; chỉ phần '………' được tô vàng (AGENTS mục 13)."""
    for i, piece in enumerate(text.split(PLACEHOLDER)):
        if i:
            r = p.add_run(PLACEHOLDER)
            r.font.size, r.bold, r.italic = Pt(size), bold, italic
            r.font.name = "Times New Roman"
            _hl(r)
        if piece:
            r = p.add_run(piece)
            r.font.size, r.bold, r.italic = Pt(size), bold, italic
            r.font.name = "Times New Roman"


def _highlight_placeholders(doc):
    """Tô vàng mọi '………' còn sót (thân văn bản, bảng) để người dùng thấy ô chờ dữ liệu."""
    for r in list(doc.element.body.iter(qn("w:r"))):
        if not any(PLACEHOLDER in (t.text or "") for t in r.findall(qn("w:t"))):
            continue
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        if rpr.find(qn("w:highlight")) is None:
            h = OxmlElement("w:highlight")
            h.set(qn("w:val"), "yellow")
            rpr.append(h)


def _highlight_header_placeholders(doc):
    for sec in doc.sections:
        for part in (sec.header, sec.first_page_header, sec.footer):
            if part.is_linked_to_previous:
                continue
            for r in part._element.iter(qn("w:r")):
                if any(PLACEHOLDER in (t.text or "") for t in r.findall(qn("w:t"))):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr")
                        r.insert(0, rpr)
                    if rpr.find(qn("w:highlight")) is None:
                        h = OxmlElement("w:highlight")
                        h.set(qn("w:val"), "yellow")
                        rpr.append(h)


def _keep(p_el):
    pPr = p_el.get_or_add_pPr()
    if pPr.find(qn("w:keepNext")) is None:
        pPr.insert(0, OxmlElement("w:keepNext"))


def _keep_rows(rows):
    for tr in rows:
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)


def _keep_table_head(tbl_el, n=4):
    rows = tbl_el.findall(qn("w:tr"))
    _keep_rows(rows[:min(n, len(rows) - 1)])


def _keep_table_tail(tbl_el, n=2):
    _keep_rows(tbl_el.findall(qn("w:tr"))[-n:])


def _is_note_table(el) -> bool:
    return el is not None and el.tag == qn("w:tbl") and _ptext(el).strip().startswith(("Ghi chú", "Note"))


def _fix_cover_lists(doc):
    """Thiết bị / Căn cứ (Table 0): bỏ numPr (tránh '- - '), thụt 0 để mọi dòng '- ' thẳng hàng (AGENTS mục 14)."""
    for row in doc.tables[0].rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if not p.text.lstrip().startswith("- "):
                    continue
                pPr = p._p.get_or_add_pPr()
                for el in pPr.findall(qn("w:numPr")) + pPr.findall(qn("w:ind")):
                    pPr.remove(el)
                ind = OxmlElement("w:ind")
                ind.set(qn("w:left"), "0")
                ind.set(qn("w:hanging"), "0")
                pPr.append(ind)


def _keep_headings_with_tables(doc):
    """Tiêu đề mục đứng ngay trước bảng kết quả -> keepNext; tiêu đề cột + 3 dòng đầu đi liền (AGENTS mục 3)."""
    for el in doc.element.body.iterchildren(qn("w:tbl")):
        prev = el.getprevious()
        if prev is not None and prev.tag == qn("w:p") and _ptext(prev).strip():
            _keep(prev)
            rows = el.findall(qn("w:tr"))
            _keep_rows(rows[:min(5, len(rows) - 1)])


def _superscript_note_marks(doc):
    """'(n)' đầu dòng trong bảng Ghi chú -> chỉ số trên (AGENTS mục 3)."""
    from docx.text.run import Run
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        if not _is_note_table(tbl):
            continue
        for r_el in list(tbl.iter(qn("w:r"))):
            ts = r_el.findall(qn("w:t"))
            m = re.match(r"^\s*(\(\d\))(.*)$", ts[0].text or "") if len(ts) == 1 else None
            if not m:
                continue
            run = Run(r_el, None)
            if run.font.superscript:
                continue
            if m.group(2):
                rest = copy.deepcopy(r_el)
                r_el.addnext(rest)
                Run(rest, None).text = m.group(2)
            run.text = m.group(1)
            run.font.superscript = True


def _replace_in_paragraphs(root, mapping: dict) -> int:
    """Thay chuỗi trong từng đoạn (kể cả khi chuỗi nằm vắt qua nhiều run)."""
    n = 0
    for p in root.iter(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        txt = "".join(t.text or "" for t in ts)
        if not any(a in txt for a in mapping):
            continue
        for a, b in mapping.items():
            while a in txt:
                start = txt.index(a)
                pos, first = 0, None
                for t in ts:
                    L_ = len(t.text or "")
                    if pos + L_ > start:
                        first = t
                        break
                    pos += L_
                new = txt.replace(a, b, 1)
                before = sum(len(t.text or "") for t in ts[:ts.index(first)])
                first.text = new[before:]
                first.set(qn("xml:space"), "preserve")
                for t in ts[ts.index(first) + 1:]:
                    t.text = ""
                txt = new
                n += 1
    return n


def _set_text(p_el, text: str, superscript_tail: str = ""):
    """Đặt lại chữ của đoạn, giữ định dạng run đầu; `superscript_tail` = phần cuối in chỉ số trên (vd '(1)')."""
    runs = p_el.findall(qn("w:r"))
    proto = copy.deepcopy(runs[0])
    for r in runs:
        p_el.remove(r)
    for ch in [c for c in p_el if c.tag != qn("w:pPr")]:
        p_el.remove(ch)

    def run(txt, sup=False):
        r = copy.deepcopy(proto)
        for c in [c for c in r if c.tag != qn("w:rPr")]:
            r.remove(c)
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        for va in rpr.findall(qn("w:vertAlign")):
            rpr.remove(va)
        if sup:
            va = OxmlElement("w:vertAlign")
            va.set(qn("w:val"), "superscript")
            rpr.append(va)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = txt
        r.append(t)
        p_el.append(r)
    run(text)
    if superscript_tail:
        run(superscript_tail, True)


def _set_cell(tc, text: str, superscript_tail: str = ""):
    """Ô bảng (chép từ bảng mẫu): giữ đoạn đầu có run, bỏ đoạn thừa, ghi chữ mới (mỗi '\\n' một đoạn)."""
    ps = tc.findall(qn("w:p"))
    base = next((p for p in ps if p.findall(qn("w:r"))), ps[0])
    for p in ps:
        if p is not base:
            tc.remove(p)
    lines = text.split("\n")
    cur = base
    for i, line in enumerate(lines):
        if i:
            nxt = copy.deepcopy(base)
            cur.addnext(nxt)
            cur = nxt
        _set_text(cur, line, superscript_tail if i == len(lines) - 1 else "")


def _vmerge(tc, restart: bool):
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:vMerge")):
        tcPr.remove(old)
    vm = OxmlElement("w:vMerge")
    if restart:
        vm.set(qn("w:val"), "restart")
    tcPr.append(vm)


def note_rows(proto, lines):
    """Bảng 'Ghi chú' chép từ bảng mẫu: dòng tiêu đề + mỗi ghi chú một dòng (marker | nội dung)."""
    el = copy.deepcopy(proto)
    rows = el.findall(qn("w:tr"))
    item = rows[-1]
    for r in rows[1:]:
        el.remove(r)
    for marker, text in lines:
        tr = copy.deepcopy(item)
        el.append(tr)
        tcs = tr.findall(qn("w:tc"))
        _set_cell(tcs[0], marker)
        _set_cell(tcs[1], text)
    return el


class Body:
    """Ghi tuần tự các khối vào trước một mốc (mặc định: sectPr cuối thân văn bản)."""

    def __init__(self, doc, text_proto, size):
        self.doc, self.proto, self.size = doc, text_proto, size
        self.sect = doc.element.body.find(qn("w:sectPr"))

    def emit(self, el):
        self.sect.addprevious(el)
        return el

    def para(self, text="", bold=False, italic=False, indent_cm=None, keep_next=False, space_before=None):
        el = copy.deepcopy(self.proto)
        for ch in [c for c in el if c.tag != qn("w:pPr")]:
            el.remove(ch)
        self.emit(el)
        p = Paragraph(el, self.doc._body)
        add_runs(p, text, self.size, bold, italic)
        pf = p.paragraph_format
        if indent_cm is not None:
            pf.left_indent = Cm(indent_cm)
            pf.first_line_indent = Cm(0)
        if keep_next:
            pf.keep_with_next = True
        if space_before is not None:
            pf.space_before = Pt(space_before)
        return p

    def table(self, rows, cols, widths_pct, header_rows=1, size=13, row_height=440):
        t = self.doc.add_table(rows=rows, cols=cols)
        self.emit(t._tbl)
        format_table_structure(t, row_height=row_height)
        for mar in list(t._tbl.iter(qn("w:tcMar"))):     # E51: tcMar 40/40 làm Word cộng thêm ~4,5 pt/dòng
            mar.getparent().remove(mar)
        total = 9355
        for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths_pct):
            gc.set(qn("w:w"), str(round(total * w / 100)))
        for row in t.rows:
            for c, w in zip(row._tr.findall(qn("w:tc")), widths_pct):
                tcPr = c.get_or_add_tcPr()
                old = tcPr.find(qn("w:tcW"))
                if old is not None:
                    tcPr.remove(old)
                tcW = OxmlElement("w:tcW")
                tcW.set(qn("w:w"), str(round(50 * w)))
                tcW.set(qn("w:type"), "pct")
                tcPr.insert(0, tcW)
        for r in t.rows[:header_rows]:
            add_tbl_header(r)
            for c in r.cells:
                set_cell_shading(c, "D9D9D9")
        t._size = size
        t._gap = 1 if row_height >= 440 else 0
        return t


def cell_text(table, cell, text, bold=False, italic=False, align="center"):
    for p in list(cell.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = cell.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    for i, line in enumerate(text.split("\n")):
        p = p0 if i == 0 else cell.add_paragraph()
        p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
        pf = p.paragraph_format
        pf.space_before, pf.space_after, pf.line_spacing = Pt(table._gap), Pt(table._gap), 1.0
        add_runs(p, line, table._size, bold, italic)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def photo_grid(body: Body, files, width=PHOTO_WIDTH_3COL):
    """Ảnh hiện trường: lưới 3 cột 4–5 cm (chuẩn 4,8 cm), không chú thích (AGENTS mục 3)."""
    photos = [_optimized(os.path.join(PHOTO_DIR, f)) for f in files]
    ncol = 3
    nrow = (len(photos) + ncol - 1) // ncol
    t = body.doc.add_table(rows=nrow, cols=ncol)
    body.emit(t._tbl)
    set_table_autofit_window(t)
    for i, path in enumerate(photos):
        cell = t.rows[i // ncol].cells[i % ncol]
        insert_image_to_cell(cell, path, width=width)
        for p in cell.paragraphs:
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(3)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for row in t.rows:
        add_cant_split(row)
    return t


def photo_groups(body: Body, groups, width2=Cm(7.8), width3=Cm(5.0)):
    """Ảnh gom theo nhóm như TBKQ mẫu Q326D124796 (K Pack): mỗi nhóm một lưới ảnh + một dòng chú thích nhóm
    (13pt nghiêng, giữa). Nhóm 3 ảnh hoặc từ 5 ảnh: lưới 3 cột (4–5 cm); còn lại lưới 2 cột (7–8 cm) — AGENTS mục 3."""
    for caption, files in groups:
        photos = [_optimized(os.path.join(PHOTO_DIR, f)) for f in files]
        ncol = 3 if (len(photos) == 3 or len(photos) >= 5) else min(2, len(photos))
        width = width3 if ncol == 3 else width2
        t = body.doc.add_table(rows=(len(photos) + ncol - 1) // ncol, cols=ncol)
        body.emit(t._tbl)
        set_table_autofit_window(t)
        for i, path in enumerate(photos):
            cell = t.rows[i // ncol].cells[i % ncol]
            insert_image_to_cell(cell, path, width=width)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(3)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        for row in t.rows:
            add_cant_split(row)
        _keep_rows([t.rows[-1]._tr])            # hàng ảnh cuối đi liền dòng chú thích nhóm
        cap = body.para(caption, italic=True)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = cap.paragraph_format
        pf.left_indent = pf.first_line_indent = Cm(0)
        pf.space_before, pf.space_after = Pt(0), Pt(6)


def unshade(cell):
    """Bỏ nền xám của ô tiêu đề con (tên chỉ tiêu) — như bảng kết quả của TBKQ mẫu Q326D124796."""
    tcPr = cell._tc.get_or_add_tcPr()
    for shd in tcPr.findall(qn("w:shd")):
        tcPr.remove(shd)


def unit_header(table, cell, text, unit):
    """Ô tiêu đề 'Số khuẩn lạc có …, <đơn vị>' — đơn vị in nghiêng như TBKQ mẫu."""
    cell_text(table, cell, text)
    add_runs(cell.paragraphs[0], unit, table._size, italic=True)


def location_table(b: Body, with_names=False, row_height=440):
    """Bảng 'Vị trí | Ký hiệu vị trí | Nội dung thực hiện' (TBKQ mẫu Q326D124796); BCKQ ghi thêm tên nhân viên."""
    n_air, n_bm = len(AIR), len(SWAB)
    t = b.table(1 + n_air + n_bm, 3, [43, 19, 38], header_rows=1, size=13, row_height=row_height)
    for j, h in enumerate(["Vị trí", "Ký hiệu vị trí", "Nội dung thực hiện"]):
        cell_text(t, t.rows[0].cells[j], h)
    i = 1
    for sym, pos, _ in AIR:
        cell_text(t, t.rows[i].cells[0], pos, align="left")
        cell_text(t, t.rows[i].cells[1], sym)
        i += 1
    for no, pos, who, _a, _c in SWAB:
        cell_text(t, t.rows[i].cells[0], f"{pos} - {who}" if (with_names and who) else pos, align="left")
        cell_text(t, t.rows[i].cells[1], f"BM{no}")
        i += 1
    cell_text(t, t.rows[1].cells[2].merge(t.rows[n_air].cells[2]), VS_CONTENT_AIR)
    cell_text(t, t.rows[n_air + 1].cells[2].merge(t.rows[n_air + n_bm].cells[2]), VS_CONTENT_BM)
    _keep_table_head(t._tbl, 4)
    return t


# ---------------------------------------------------------------------------
# /1 Khí nén — phiếu tính toán, dossier, TBKQ, BCKQ
# ---------------------------------------------------------------------------
def build_kn_workbook() -> str:
    """Phiếu tính toán khí nén KN1 (template TinhToanKhiNen, giữ công thức) -> ThucHien."""
    hoi = KQTN_DAU_HOI["phieu"]
    point_data = {
        "code": CODE, "point_label": KN_SYM, "cert_date": CERT_DATE_KN,
        "location_name": KN_NAME_VN, "location_symbol": KN_SYM, "location_photo": KN_PHOTO,
        "dewpoint": dict(KN_DEW),
        # Hiệu chỉnh DP510: công tắc chung AP_HIEU_CHINH đang tắt (AGENTS 8B); chưa rõ máy MT10-n -> không khai "cals".
        "dust": {"device": "AeroTrak 9350", "raw_channels": KN_DUST, "mode": "diff", "dilution_factor": 1.0},
        "oil": {"flow_aerosol": 35, "flow_vapor": 1, "time_aerosol": 240, "time_vapor": 240,   # ghi chép: 9h00–13h00
                "kqtn_aerosol": OIL_AEROSOL, "kqtn_vapor": OIL_VAPOR,
                "kqtn_no_aerosol": KQTN_DAU_HAT["phieu"], **({"kqtn_no_vapor": hoi} if hoi else {})},
    }
    path = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - {KN_SYM} - TinhToanKhiNen.xlsx")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    build_compressed_air_workbook(point_data, path)
    return path


def kn_particles():
    """[0,3–0,5; 0,5–1; 1–5; > 5 µm] hạt/m³ quy về 1 bar, 20 °C với N* = (273 + 32,8)/293 như phiếu người dùng nhập."""
    n = (273 + KN_GAS_TEMP) / (273 + 20)
    a, b, c, e, f, g = KN_DUST
    return [round(a * n), round(b * n), round((c + e + f) * n), round(g * n)]


def finalize_kn_workbook(path: str):
    """Hoàn thiện phiếu khí nén SAU khi đã đọc số liệu vào dossier: kết quả bụi trỏ sang khối AeroTrak (dòng 68),
    khối Lasair không dùng -> gạch chéo (AGENTS 8B); N* theo nhiệt độ khí 32,8 °C như phiếu người dùng nhập; xóa số
    phiếu MTTN mẫu còn trong template; ghi số liệu vi sinh; rồi tính lại bằng Excel COM (AGENTS mục 11)."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Tinh_Toan_Khi_Co_Ap_V6"]
    ws["C36"].value = KN_GAS_TEMP
    ws["C4"].value = "- Kích thước 0,3 - 0,5 µm"
    for cell, ref in (("D4", "C68"), ("D5", "D68"), ("D6", "E68"), ("D7", "F68")):
        ws[cell].value = f"={ref}"
    ws["C52"].value = "-"
    strike_box(ws, "C54:J59")                 # khối Lasair III 110 (không dùng)
    ws["C85"].value = "DP 510"                # ghi chép: "file máy CS 510" — không ghi số máy
    hoi = re.match(r"([A-Z]+)(.+)", KQTN_DAU_HOI["phieu"] or "")
    ws["I43"].value, ws["J43"].value = (hoi.group(1), hoi.group(2)) if hoi else ("MTTN", None)
    if OIL_PENDING:                           # chờ KQTN dầu hơi: ô nhập để '………', ô kết quả tổng dầu tô vàng (tạm)
        from openpyxl.styles import PatternFill
        ws["D46"].value = PLACEHOLDER
        for ref in ("D46", "C48", "C49", "D9", "F9", "H9", "J43"):
            ws[ref].fill = PatternFill("solid", fgColor="FFFF00")
    # Vi sinh khí nén (VSTN26027754.08): tổng số, E. coli < 1 CFU/đĩa; Salmonella spp. KPH
    ws["C78"].value = 500
    ws["C81"].value = "< 1"
    ws["G80"].value, ws["G81"].value = "E.coli", "< 1"
    ws["H80"].value, ws["H81"].value, ws["H83"].value = "Salmonella spp.", "KPH", "Không có chỉ thị"
    ws["I77"].value = KQTN_VS_KN["phieu"]
    from openpyxl.styles import Alignment
    for ref in ("H2", "J44"):
        a = ws[ref].alignment
        ws[ref].alignment = Alignment(horizontal=a.horizontal, vertical=a.vertical, shrink_to_fit=True)
    if "DiemSuong" in wb.sheetnames:          # sheet phụ (bảng hiệu chuẩn DP510) -> ẩn, PDF không in trang trắng
        wb["DiemSuong"].sheet_state = "hidden"
    thin_outside_borders(wb)
    wb.save(path)
    wb.close()
    try:
        import win32com.client
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            x = xl.Workbooks.Open(os.path.abspath(path))
            xl.CalculateFull()
            got = [x.Worksheets("Tinh_Toan_Khi_Co_Ap_V6").Range(a).Value for a in ("D4", "D5", "D6", "D7")]
            x.Save()
            x.Close()
        finally:
            xl.Quit()
        if [round(v) for v in got] != kn_particles():
            raise SystemExit(f"[FATAL] [Data Mismatch] bụi khí nén: phiếu {got} khác TBKQ {kn_particles()}")
    except ImportError as e:
        print(f"[build_134179] Excel COM warning: {e}")


def compressed_air_dossier(xlsx: str) -> Dossier:
    d = create_dossier_from_compressed_air_workbooks(
        [xlsx], code=CODE, cert_date=CERT_DATE_KN, sample_date=SAMPLE_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=1, client=CLIENT, personnel=PersonnelInfo(executor=EXECUTOR),
        equipments=[eq("aerotrak_9350_hpd3"), eq("cs_dp510"), eq("oil_kit"), eq("sas_pinocchio")],
        scope_vn="Lấy mẫu, phân tích chất lượng khí nén", scope_en="Sampling and analysis of compressed air quality",
        standards_vn="ISO 8573-1:2010", standards_en="ISO 8573-1:2010",
    )
    d.reg_date = REG_DATE
    d.procedure_vn = d.procedure_en = "QUATEST 3 3030:2025, N2.TC.KN.01"      # ghi chép /1 mục 6
    d.location_desc_vn, d.location_desc_en = LOCATION_VN, LOCATION_EN
    loc = d.locations[0]
    loc.name_vn, loc.name_en = KN_NAME_VN, KN_NAME_EN
    d.photos = [{"path": _optimized(KN_PHOTO), "caption_vn": f"Vị trí lấy mẫu {KN_SYM}",
                 "caption_en": f"Sampling location {KN_SYM}"}]
    rd = d.results_data
    if OIL_PENDING:                           # chưa có KQTN dầu hơi -> không ghi kết quả tổng dầu (AGENTS mục 13)
        for r in rd["oil"]["rows"]:
            r["val"], r["class"] = PLACEHOLDER, PLACEHOLDER
    for r in rd["dewpoint"]["rows"]:          # hàm lượng nước ghi số nguyên (như 133266); T2 theo phiếu người dùng
        r["values"][4] = str(round(float(r["values"][4].replace(",", "."))))
        if r["values"][3] != f"{KN_T2_7BAR:.1f}".replace(".", ","):
            print(f"[WARNING] [Data Mismatch] điểm sương 7 bar: code {r['values'][3]}, phiếu người dùng {KN_T2_7BAR}")
            r["values"][3] = f"{KN_T2_7BAR:.1f}".replace(".", ",")
    vals = kn_particles()                     # N* theo 32,8 °C (phiếu người dùng); code Core lấy trung bình 32,75 °C
    row = rd["particles"]["rows"][0]
    core = row["values"][:3]
    row["values"][:3] = [f"{v:,}".replace(",", " ") for v in vals[:3]]
    row["d_gt_50"] = vals[3]
    print(f"[build_134179] bụi khí nén (N* theo {KN_GAS_TEMP} °C): {row['values'][:3]} + > 5 µm {vals[3]} (Core, nhiệt độ trung bình: {core})")
    rd["system_status_vn"], rd["factory_status_vn"] = "Đang hoạt động", "Đang hoạt động sản xuất"   # ghi chép tr. 1
    rd["system_chain_vn"], rd["system_chain_en"] = SYSTEM_CHAIN_VN, SYSTEM_CHAIN_EN
    hoi = KQTN_DAU_HOI["phieu"] or f"MTTN{PLACEHOLDER}"
    rd["kqtn_by_location"] = {KN_SYM: f"- Dầu hơi: xem tại phiếu {hoi}\n- Dầu hạt: xem tại phiếu {KQTN_DAU_HAT['phieu']}"}

    def kq(k):
        if not k["so"]:
            return {"so": f"MTTN{PLACEHOLDER}", "ngay": PLACEHOLDER, "ngay_gui": PLACEHOLDER, "ngay_nhan": PLACEHOLDER,
                    "ngay_yc": NGAY_YC_KQTN, "tre_han": "[  ] có\t\t[  ] không"}
        return {"so": k["so"], "ngay": k["ngay"], "ngay_gui": k["ngay_nhan_mau"], "ngay_nhan": k["ngay"],
                "ngay_yc": NGAY_YC_KQTN, "tre_han": tre_han_text(k["ngay"], k["ngay_nhan_mau"], NGAY_YC_KQTN, PLACEHOLDER)}
    rd["kqtn"] = [kq(KQTN_DAU_HAT), kq(KQTN_DAU_HOI), kq(KQTN_VS_KN)]
    rd["bckq_label"] = "khi nen"
    d.input_dir, d.output_dir, d.network_dir = INPUT_DIR, "", ""
    d.thuchien_dir = os.path.join(INPUT_DIR, "ThucHien")
    d.tbkq_dir = os.path.join(INPUT_DIR, "TBKQ")
    d.edoc_dir = os.path.join(d.tbkq_dir, "Edoc")
    d.pdf_dir = os.path.join(INPUT_DIR, "pdf")
    return d


def _prev_text_p(el):
    prev = el.getprevious()
    while prev is not None and not (prev.tag == qn("w:p") and _ptext(prev).strip()):
        prev = prev.getprevious()
    return prev


def _kn_tbkq_patch(doc):
    """TBKQ khí nén sau khi handler dựng: (1) AeroTrak đo từ 0,3 µm -> đổi '0,1' thành '0,3' ở tiêu đề cột và dòng chỉ
    tiêu; (2) bỏ sơ đồ hệ thống (BBLM: theo yêu cầu Công ty không thể hiện tại TBKQ); (3) thêm chỉ tiêu + mục
    'Nhiễm bẩn vi sinh vật' (template khí nén chưa có — theo mẫu TBKQ 103230 người dùng để trên Z)."""
    body = doc.element.body
    # (1)
    for p in body.iter(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        txt = "".join(t.text or "" for t in ts)
        if "d ≤ 0,5" in txt and "0,1" in txt and len(txt) < 60:
            for t in ts:
                if t.text and "0,1" in t.text:
                    t.text = t.text.replace("0,1", "0,3")
                    break
    # (3a) dòng chỉ tiêu trong bảng vị trí
    for p in list(body.iter(qn("w:p"))):
        if _ptext(p).strip() == "Nồng độ hạt bụi có kích thước (0,1-5,0) µm.":
            _set_text(p, "Nồng độ hạt bụi có kích thước (0,3-5,0) µm;")
            extra = copy.deepcopy(p)
            p.addnext(extra)
            _set_text(extra, "Độ nhiễm bẩn vi sinh vật.")
    # (2)
    for p in list(body.iterchildren(qn("w:p"))):
        if _ptext(p).strip().startswith("Mô tả hệ thống khí nén"):
            chain = p.getnext()
            body.remove(p)
            if chain is not None and chain.tag == qn("w:p") and "Belliss" in _ptext(chain):
                body.remove(chain)
    # (3b) mục vi sinh: chép tiêu đề + bảng + đoạn ngăn + bảng Ghi chú của mục tổng dầu rồi ghi lại nội dung
    tbls = list(body.iterchildren(qn("w:tbl")))
    oil = next(t for t in tbls if "Hàm lượng tổng dầu, mg/m" in _ptext(t))
    oil_head = _prev_text_p(oil)
    between, nxt = [], oil.getnext()
    while nxt is not None and not _is_note_table(nxt):
        between.append(nxt)
        nxt = nxt.getnext()
    oil_note = nxt
    last_note = [t for t in tbls if _is_note_table(t)][-1]
    after_note = []                     # các đoạn (trống) giữa bảng Ghi chú dầu và tiêu đề mục kế -> chép cùng kiểu
    nxt = oil_note.getnext()
    while nxt is not None and nxt.tag == qn("w:p") and not _ptext(nxt).strip():
        after_note.append(nxt)
        nxt = nxt.getnext()

    anchor = last_note
    def put(el):
        nonlocal anchor
        anchor.addnext(el)
        anchor = el
        return el
    for g in after_note:
        put(copy.deepcopy(g))
    head = put(copy.deepcopy(oil_head))
    _set_text(head, "Nhiễm bẩn vi sinh vật")
    _keep(head)
    tbl = put(copy.deepcopy(oil))
    rows = tbl.findall(qn("w:tr"))
    data_proto = rows[1]
    for r in rows[2:]:
        tbl.remove(r)
    tcs = rows[0].findall(qn("w:tc"))
    _set_cell(tcs[0], "Vị trí thực hiện")
    _set_cell(tcs[1], "Chỉ tiêu")
    _set_cell(tcs[2], "Số khuẩn lạc trong khí nén, CFU/m³ ", "(1)")
    for i, (name, _kqtn, val) in enumerate(KN_MICRO):
        tr = data_proto if i == 0 else copy.deepcopy(data_proto)
        if i:
            tbl.append(tr)
        tcs = tr.findall(qn("w:tc"))
        _set_cell(tcs[0], KN_SYM if i == 0 else "")
        _vmerge(tcs[0], restart=(i == 0))
        _set_cell(tcs[1], name)
        _set_cell(tcs[2], val)
        for h in tr.iter(qn("w:trHeight")):       # dòng 0,7 cm để mục vi sinh nằm trọn trang phụ lục 2 (đã soát PDF)
            h.set(qn("w:val"), "400")
    _keep_rows(tbl.findall(qn("w:tr")))
    grid = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))      # cột 18 / 34 / 48 % để tiêu đề cột nằm 1 dòng
    total = sum(int(g.get(qn("w:w"))) for g in grid)
    widths = [round(total * w) for w in (0.18, 0.34, 0.48)]
    for g, w in zip(grid, widths):
        g.set(qn("w:w"), str(w))
    for tr in tbl.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), widths):
            tcW = tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tcW is not None:
                tcW.set(qn("w:w"), str(w))
                tcW.set(qn("w:type"), "dxa")
    for g in between:
        put(copy.deepcopy(g))
    note = put(note_rows(oil_note, [
        ("(1) …", "Kết quả quy về điều kiện áp suất tuyệt đối 1 bar, nhiệt độ 20 °C theo ISO 8573-7:2003."),
        ("●", "Kết quả được biểu thị “Không có chỉ thị” khi không phát hiện vi sinh vật mục tiêu trong khí nén "
              "được lấy mẫu."),
        ("●", "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004."),
    ]))
    _keep_rows(note.findall(qn("w:tr"))[:-1])


def _resize_pictures(root, width_cm: float):
    """Đổi độ rộng ảnh (giữ tỉ lệ, không crop — AGENTS mục 3)."""
    new_cx = int(Cm(width_cm))
    for inline in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline"):
        ext = inline.find("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
        new_cy = round(cy * new_cx / cx)
        ext.set("cx", str(new_cx)); ext.set("cy", str(new_cy))
        for a_ext in inline.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}ext"):
            if a_ext.get("cx"):
                a_ext.set("cx", str(new_cx)); a_ext.set("cy", str(new_cy))


def _keep_tails_together(doc):
    """Bảng kết quả (2 dòng cuối) + đoạn ngăn + bảng Ghi chú cùng trang; dòng cuối bảng Ghi chú không keepNext
    khi không có câu Đánh giá theo sau (AGENTS mục 3, E32)."""
    for tbl in list(doc.element.body.iterchildren(qn("w:tbl"))):
        note, gaps = tbl.getnext(), []
        while note is not None and note.tag == qn("w:p") and not _ptext(note).strip():
            gaps.append(note)
            note = note.getnext()
        if _is_note_table(tbl) or not _is_note_table(note):
            continue
        for g in gaps:
            _keep(g)
        _keep_rows(tbl.findall(qn("w:tr"))[-2:])
        _keep_rows(note.findall(qn("w:tr"))[:-1])


KN_PHOTO_WIDTH_CM = 8.0            # 1 ảnh: thu nhỏ để mục 9.2.1 (bảng + Ghi chú) nằm trọn trang phụ lục 1 — đã soát PDF


def build_tbkq_kn(d: Dossier) -> str:
    src = CompressedAirHandler(d).build_tbkq(is_en=False)
    doc = Document(src)
    _kn_tbkq_patch(doc)
    for t in doc.element.body.iterchildren(qn("w:tbl")):
        if "Vị trí lấy mẫu KN1" in _ptext(t) and t.find(f".//{qn('w:drawing')}") is not None:
            _resize_pictures(t, KN_PHOTO_WIDTH_CM)
    _fix_cover_lists(doc)
    _keep_headings_with_tables(doc)
    _keep_tails_together(doc)
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    _highlight_header_placeholders(doc)
    style_note_tables(doc)          # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - viet - 1 - khi nen - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)              # tên handler 'TBKQ - vn - 1 - khi nen - …' (file do code tạo, E41)
    print(f"[build_134179] TBKQ khí nén -> {dst}")
    return dst


def build_bckq_kn(d: Dossier) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - khi nen - {CODE}.docx")
    doc = Document(out)
    body = doc.element.body
    so_phieu = f"{CODE} - {KN_SYM}"
    _replace_in_paragraphs(body, {
        f"phiếu Tính toán kết quả nồng độ hạt bụi số {CODE} ngày": f"phiếu Kết quả tính toán số {so_phieu} ngày",
        f"phiếu Tính toán kết quả nhiệt độ điểm sương số {CODE} ngày": f"phiếu Kết quả tính toán số {so_phieu} ngày",
        "Nồng độ hạt bụi có kích thước (0,1-5,0) µm.": "Nồng độ hạt bụi có kích thước (0,3-5,0) µm; Độ nhiễm bẩn vi sinh vật.",
    })
    # Mục "Vi sinh trong khí nén" sau mục điểm sương (trước "Gửi mẫu thử nghiệm:")
    kids = list(body.iterchildren(qn("w:p")))
    dp_head = next(p for p in kids if _ptext(p).strip() == "Nhiệt độ điểm sương")
    dp_ref = dp_head.getnext()
    head = copy.deepcopy(dp_head)
    dp_ref.addnext(head)
    _set_text(head, "Vi sinh trong khí nén")
    _keep(head)
    kq = "; ".join(f"{n}: {k}" for n, k, _ in KN_MICRO)
    lines = [
        f"Xem tại phiếu KQTN số {KQTN_VS_KN['phieu']} ngày {KQTN_VS_KN['ngay']} (mẫu {KN_SYM}, 03 đĩa, 100 lít/phút × "
        f"05 phút/đĩa = 500 lít: {kq}).",
        "- KQ được biểu thị < 1 CFU/đĩa (KPH đối với Salmonella spp.) → không phát hiện vi sinh vật mục tiêu trong "
        "thể tích khí nén được lấy mẫu ~ Không có chỉ thị.",
        f"- KQTN của mẫu Blank (KN2, phiếu {KQTN_VS_KN['blank']}): tổng số vi sinh vật hiếu khí < 1 CFU/đĩa, "
        "Escherichia coli < 1 CFU/đĩa, Salmonella spp. KPH → môi trường lấy mẫu đảm bảo kết quả.",
    ]
    anchor = head
    for i, line in enumerate(lines):
        p = copy.deepcopy(dp_ref)
        anchor.addnext(p)
        _set_text(p, line)
        if i < len(lines) - 1:
            _keep(p)
        anchor = p
    _keep_headings_with_tables(doc)
    for t in doc.tables[1:]:
        if len(t.rows) > 1 and t.rows[0].cells[0].text.strip().startswith("Vị trí"):
            add_tbl_header(t.rows[0])
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "- Hình ảnh")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    print(f"[build_134179] BCKQ khí nén -> {out}")
    return out


# ---------------------------------------------------------------------------
# /2 Vi sinh không khí + bề mặt
# ---------------------------------------------------------------------------
def vs_dossier() -> Dossier:
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=2, is_multi_cert=True, dossier_type="cleanroom",       # dùng khung bìa phòng sạch, phụ lục dựng lại
        scope_vn=("- Lấy mẫu và xác định mức nhiễm bẩn vi sinh vật trong không khí;\n"
                  "- Lấy mẫu và xác định mức nhiễm bẩn vi sinh vật trên bề mặt."),
        scope_en=("- Sampling and determining microbial contamination in workplace air;\n"
                  "- Sampling and determining microbial contamination on surface."),
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.09", procedure_en="QUATEST 3 3030:2025, N2.TC.09",   # ghi chép /2
        standards_vn="-", standards_en="-",                           # không có mức quy định
        client=CLIENT, personnel=PersonnelInfo(executor=EXECUTOR),
        equipments=[eq("buck_b30120"), Equipment(name_vn=EQUIP_SWAB_VN, name_en=EQUIP_SWAB_EN)],
        locations=[], photos=[], results_data={"bckq_label": "vi sinh"},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


def build_tbkq_vs(d: Dossier) -> str:
    src = CleanroomHandler(d).build_tbkq(is_en=False)
    doc = Document(src)
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    sect_i = next(i for i, el in enumerate(kids)
                  if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None)
    after = kids[sect_i + 1:]
    heading = next(el for el in after if el.tag == qn("w:p") and _ptext(el).strip() == "Nội dung thực hiện")
    text_proto = copy.deepcopy(heading.getnext())
    note_proto = copy.deepcopy(next(el for el in after if _is_note_table(el)))
    head_proto = copy.deepcopy(heading)
    for el in after:
        if el.tag != qn("w:sectPr"):
            body_el.remove(el)

    b = Body(doc, text_proto, 13)

    def head(text):
        el = copy.deepcopy(head_proto)
        for r in el.findall(qn("w:r"))[1:]:
            el.remove(r)
        el.find(qn("w:r")).find(qn("w:t")).text = text
        b.emit(el)
        _keep(el)
        return el

    # Bố cục phụ lục theo TBKQ mẫu Q326D124796 (K Pack) — người dùng yêu cầu 14/11/2026:
    # 9.1 = 3 dòng dẫn + bảng 'Vị trí | Ký hiệu vị trí | Nội dung thực hiện' + ảnh gom theo nhóm;
    # 9.2 = bảng kết quả chỉ ghi ký hiệu vị trí; Ghi chú: '< 10' -> VILAS -> VIAS.
    head("Nội dung thực hiện")
    b.para(f"- Khu vực lấy mẫu theo chỉ định của {CLIENT.name_vn}.")
    b.para("- Trạng thái nhà xưởng: hoạt động sản xuất.")
    b.para("- Nội dung thực hiện và ký hiệu vị trí:", keep_next=True)
    n_air, n_bm = len(AIR), len(SWAB)
    location_table(b)
    b.para("- Hình ảnh quá trình lấy mẫu:", keep_next=True, space_before=6)
    photo_groups(b, VS_PHOTO_GROUPS)

    # 9.2 Kết quả giám định
    head("Kết quả giám định")
    b.para("9.2.1 Độ nhiễm bẩn vi sinh trong không khí", bold=True, keep_next=True)
    t = b.table(2 + n_air, 2, [26, 74], header_rows=2, size=13)
    cell_text(t, t.rows[0].cells[0].merge(t.rows[1].cells[0]), "Vị trí lấy mẫu")
    unit_header(t, t.rows[0].cells[1], "Số khuẩn lạc có trong không khí, ", "CFU/m³")
    cell_text(t, t.rows[1].cells[1], "Tổng số vi sinh vật hiếu khí", italic=True)
    unshade(t.rows[1].cells[1])
    for k, (sym, _pos, cfu) in enumerate(AIR, start=2):
        cell_text(t, t.rows[k].cells[0], sym)
        cell_text(t, t.rows[k].cells[1], str(air_value(cfu)))
    _keep_rows(t._tbl.findall(qn("w:tr")))
    nt = b.emit(note_rows(note_proto, [
        ("●", "Kết quả được biểu thị “< 10 CFU/m³” khi không phát hiện vi sinh vật mục tiêu trên thể tích không khí "
              "được lấy mẫu."),
        ("●", NOTE_VILAS),
        ("●", NOTE_VIAS),
    ]))
    _keep_rows(nt.findall(qn("w:tr"))[:-1])

    b.para("9.2.2 Độ nhiễm bẩn vi sinh trên bề mặt", bold=True, keep_next=True, space_before=6)
    t = b.table(2 + n_bm, 3, [26, 39, 35], header_rows=2, size=13, row_height=BM_ROW_TWIPS)
    r0 = t.rows[0].cells
    cell_text(t, r0[0].merge(t.rows[1].cells[0]), "Vị trí lấy mẫu")
    unit_header(t, r0[1].merge(r0[2]), "Số khuẩn lạc có trên bề mặt, ", UNIT_BM)
    for j, h in ((1, "Tổng số vi sinh vật hiếu khí"), (2, "Coliforms")):
        cell_text(t, t.rows[1].cells[j], h, italic=True)
        unshade(t.rows[1].cells[j])
    for k, (no, _pos, _who, apc, coli) in enumerate(SWAB, start=2):
        cells = t.rows[k].cells
        cell_text(t, cells[0], f"BM{no}")
        cell_text(t, cells[1], apc)
        cell_text(t, cells[2], coli)
    if n_bm <= 10:                     # bảng ngắn (≤ 10 dòng số liệu): giữ trọn một trang (AGENTS mục 15)
        _keep_rows(t._tbl.findall(qn("w:tr")))
    else:
        _keep_table_head(t._tbl, 5)
        _keep_table_tail(t._tbl, 3)    # bảng ngắt trang: 3 dòng cuối + Ghi chú đi cùng nhau (đã soát PDF: 11 / 3 dòng)
    nt = b.emit(note_rows(note_proto, [
        ("●", f"Kết quả được biểu thị “< 10 {UNIT_BM}” khi không phát hiện vi sinh vật mục tiêu trên diện tích bề mặt "
              "được lấy mẫu."),
        ("●", NOTE_VILAS),
        ("●", NOTE_VIAS),
    ]))
    _keep_rows(nt.findall(qn("w:tr"))[:-1])

    _fix_cover_lists(doc)
    add_document_ending_line(doc)
    appendix_total_field(doc)
    style_note_tables(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - viet - 2 - vi sinh - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)              # tên handler phòng sạch 'TBKQ - vn - 2 - phong sach - …' (file do code tạo)
    print(f"[build_134179] TBKQ vi sinh -> {dst}")
    return dst


def build_bckq_vs(d: Dossier) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 2 - vi sinh - {CODE}.docx")
    doc = Document(out)
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    h1 = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Nội dung thực hiện"))
    h2 = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Kết quả đo đạc"))
    gdv = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Báo cáo của GĐV"))
    text_proto = copy.deepcopy(h1.getnext())
    i1, i2, ig = kids.index(h1), kids.index(h2), kids.index(gdv)
    for el in kids[i1 + 1:i2] + kids[i2 + 1:ig]:
        body_el.remove(el)

    b = Body(doc, text_proto, 13)
    b.sect = h2
    n_air, n_bm = len(AIR), len(SWAB)
    b.para(f"- Các vị trí lấy mẫu và chỉ tiêu phân tích được chỉ định bởi đại diện {CLIENT.name_vn}, cụ thể như sau:",
           keep_next=True)
    location_table(b, with_names=True, row_height=400)             # bảng 3 cột như TBKQ (mẫu Q326D124796), thêm tên nhân viên
    b.para(f"- Vi sinh không khí: thiết bị Buck Bio-Culture B30120, lưu lượng {VS_FLOW_LPM} lít/phút, thời gian "
           f"{VS_MINUTES:02d} phút/mẫu ({VS_FLOW_LPM * VS_MINUTES} lít). Vi sinh bề mặt: tăm bông tiệt trùng, 10 mL "
           f"BPW, diện tích quét 100 cm²; 01 mẫu blank (BM{SWAB_BLANK}).", space_before=6)
    b.para("- Tại thời điểm thực hiện, các khu vực lấy mẫu đang hoạt động sản xuất.")
    b.para("- Hình ảnh các vị trí lấy mẫu:", keep_next=True)
    photo_groups(b, VS_PHOTO_GROUPS, width2=BCKQ_PHOTO_W2, width3=BCKQ_PHOTO_W3)

    b.sect = gdv
    b.para("4.6.1 Kết quả lấy mẫu vi sinh không khí", bold=True, keep_next=True)
    t = b.table(1 + n_air, 4, [34, 16, 22, 28], header_rows=1, size=13)
    for j, h in enumerate(["Vị trí lấy mẫu", "Ký hiệu mẫu", "KQTN (CFU/ đĩa)", "KQTT (CFU/ m3)"]):
        cell_text(t, t.rows[0].cells[j], h)
    for k, (sym, pos, _cfu) in enumerate(AIR, start=1):
        cell_text(t, t.rows[k].cells[0], pos, align="left")
        cell_text(t, t.rows[k].cells[1], sym)
    cell_text(t, t.rows[1].cells[2].merge(t.rows[n_air].cells[2]),
              f"Xem tại phiếu KQTN số {KQTN_VSKK['so']}.01 đến .06 ngày {KQTN_VSKK['ngay']}")
    cell_text(t, t.rows[1].cells[3].merge(t.rows[n_air].cells[3]),
              f"Xem tại phiếu Tính toán kết quả vi sinh không khí số {SO_VSKK} ngày {NGAY_VSKK}")
    _keep_rows(t._tbl.findall(qn("w:tr"))[:-1])
    b.para("- KQTT (CFU/m³) = KQTN sau hiệu chỉnh theo bảng Conversion của thiết bị Buck Bio-Culture B30120 chia thể "
           f"tích không khí được lấy mẫu ({VS_FLOW_LPM * VS_MINUTES} lít).", space_before=6)
    b.para("- Không lấy mẫu Blank vi sinh không khí (phiếu ghi chép không ghi mẫu blank).")
    b.para("4.6.2 Kết quả lấy mẫu vi sinh trên bề mặt", bold=True, keep_next=True, space_before=6)
    t = b.table(3, 3, [34, 26, 40], header_rows=1, size=13)
    for j, h in enumerate(["Nội dung", "Ký hiệu mẫu", "KQTN (CFU/ mẫu)"]):
        cell_text(t, t.rows[0].cells[j], h)
    cell_text(t, t.rows[1].cells[0], "Vi sinh bề mặt")
    cell_text(t, t.rows[1].cells[1], f"BM1 → BM{n_bm}")
    cell_text(t, t.rows[2].cells[0], "Mẫu Blank")
    cell_text(t, t.rows[2].cells[1], f"BM{SWAB_BLANK}")
    cell_text(t, t.rows[1].cells[2].merge(t.rows[2].cells[2]),
              f"Xem tại phiếu KQTN số {KQTN_VSBM['so']}.01 đến .{SWAB_BLANK:02d} ngày {KQTN_VSBM['ngay']}")
    _keep_rows(t._tbl.findall(qn("w:tr")))
    b.para("- KQTN được biểu thị < 10 CFU/ mẫu → không phát hiện khuẩn lạc trên diện tích bề mặt swab (100 cm²) ~ "
           "< 10 CFU/ 100 cm².", space_before=6)
    b.para(f"- KQTN của mẫu Blank (BM{SWAB_BLANK}): < 10 CFU/ mẫu (tổng số vi sinh vật hiếu khí, Coliforms) → môi "
           "trường lấy mẫu đảm bảo kết quả.")
    b.para("Gửi mẫu thử nghiệm:", bold=True, keep_next=True, space_before=6)
    for stt, k in zip("ab", (KQTN_VSKK, KQTN_VSBM)):
        b.para(f"{stt}. Phiếu KQTN số {k['so']} ngày {k['ngay']}", keep_next=True)
        b.para(f"Ngày gửi mẫu: {k['ngay_nhan_mau']}", keep_next=True)
        b.para(f"Ngày yêu cầu cấp KQTN: {NGAY_YC_KQTN}", keep_next=True)
        b.para(f"Ngày nhận KQTN: {k['ngay']}", keep_next=True)
        tre = tre_han_text(k["ngay"], k["ngay_nhan_mau"], NGAY_YC_KQTN, PLACEHOLDER)
        b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre}")

    _keep(h2)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "- Hình ảnh")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    print(f"[build_134179] BCKQ vi sinh -> {out}")
    return out


def build_excel_vs() -> str:
    """Phiếu tính toán vi sinh không khí (form phiếu tính toán chung, chỉ bảng số liệu — AGENTS mục 11);
    bảng Conversion Buck B30120 ở sheet ẩn, công thức động như phiếu người dùng."""
    sheet = "BuckConversion"
    book = CleanroomCalcBook()
    ws = book.extra_sheet("Vi sinh không khí", so=SO_VSKK, date=NGAY_VSKK, widths=(6, 30, 10, 11, 10, 10, 15, 17, 13),
                          title="TÍNH TOÁN KẾT QUẢ VI SINH KHÔNG KHÍ")
    hdr = ("STT", "Vị trí lấy mẫu", "Ký hiệu mẫu", "Lưu lượng, L/phút", "Thời gian, phút", "Thể tích, L",
           "Tổng số VSV hiếu khí theo KQTN, CFU/đĩa", "Sau hiệu chỉnh theo bảng Conversion Buck B30120, CFU/đĩa",
           "Kết quả, CFU/m³")
    for c, lbl in enumerate(hdr, 1):
        book.put(ws, 1, c, lbl, bold=True)
    ws.row_dimensions[1].height = 62
    tbl = f"{sheet}!$A$2:$B${BUCK_HOLES + 1}"
    for i, (sym, pos, cfu) in enumerate(AIR):
        r = 2 + i
        book.put(ws, r, 1, i + 1)
        book.put(ws, r, 2, pos, align=LEFT)
        book.put(ws, r, 3, sym)
        book.put(ws, r, 4, VS_FLOW_LPM)
        book.put(ws, r, 5, VS_MINUTES)
        book.put(ws, r, 6, f"=D{r}*E{r}")
        book.put(ws, r, 7, cfu)
        book.put(ws, r, 8, f'=IF(G{r}="","",IF(OR(G{r}="< 1",G{r}=0),"< 1",VLOOKUP(G{r},{tbl},2,FALSE)))')
        book.put(ws, r, 9, f'=IF(H{r}="","",IF(H{r}="< 1","< "&ROUND(1000/F{r},0),ROUND(H{r}*1000/F{r},0)))', fmt="0")
    wb_ = book.wb.create_sheet(sheet)
    wb_["A1"], wb_["B1"], wb_["C1"], wb_["D1"], wb_["E1"] = "r (CFU/đĩa)", "Pr (sau hiệu chỉnh)", "Σ 1/(N-i+1)", "N (số lỗ)", BUCK_HOLES
    for k in range(1, BUCK_HOLES + 1):
        r = k + 1
        wb_[f"A{r}"] = k
        wb_[f"C{r}"] = "=ROUND(1/$E$1,6)" if k == 1 else f"=ROUND(C{r - 1}+1/($E$1-A{r}+1),6)"
        wb_[f"B{r}"] = f"=ROUND(C{r}*$E$1,0)"
    wb_.sheet_state = "hidden"
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - 2 - TinhToanViSinh.xlsx")
    book.save(out)
    print(f"[build_134179] Excel vi sinh -> {out}")
    return out


# ---------------------------------------------------------------------------
def split_edoc() -> list:
    """BBLM = '0. BBLM.pdf'; Ghi chép = ghi chép khí nén (3 tr.) + ghi chép vi sinh (3 tr.) + bảng AeroTrak (tr. 2
    trắng — finish_edoc bỏ). Thư mục 'E DOC' trên Z là của người dùng, không đụng tới (AGENTS 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    outs = []
    for name, srcs in ((f"{CODE6} - BBLM.pdf", [SCAN_BBLM]),
                       (f"{CODE6} - Ghi Chep.pdf", [SCAN_GC_KN, SCAN_GC_VS, SCAN_DUST])):
        dst = fitz.open()
        for s in srcs:
            dst.insert_pdf(fitz.open(s))
        finish_edoc(dst, INPUT_DIR, name)
        p = os.path.join(edoc, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_134179] Edoc -> {p}")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--chi-vi-sinh", action="store_true",
                    help="Chỉ dựng lại TBKQ + BCKQ bộ /2 vi sinh (giữ nguyên bộ /1 khí nén và 2 phiếu Excel)")
    args = ap.parse_args()

    _check_user_vskk()
    if args.chi_vi_sinh:
        vs = vs_dossier()
        vs.ensure_output_directories()
        docs = [build_bckq_vs(vs), build_tbkq_vs(vs)]
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            PDFExporter(vs).export_files(docs)
        print("\n===== QA =====")
        for p in docs:
            print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")
        return
    kn_xlsx = build_kn_workbook()
    kn = compressed_air_dossier(kn_xlsx)
    finalize_kn_workbook(kn_xlsx)                 # sau khi đã đọc số liệu vào dossier
    vs = vs_dossier()
    assign_cert_numbers([kn, vs])
    for d in (kn, vs):
        d.ensure_output_directories()
    split_edoc()

    kn_vn = build_tbkq_kn(kn)
    kn_bckq = build_bckq_kn(kn)
    vs_vn = build_tbkq_vs(vs)
    vs_bckq = build_bckq_vs(vs)
    vs_xlsx = build_excel_vs()

    docs = [kn_bckq, kn_vn, vs_bckq, vs_vn]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(kn).export_files(docs[:2])
        PDFExporter(vs).export_files(docs[2:])
        PDFExporter(kn).export_excel_files([kn_xlsx, vs_xlsx])
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
