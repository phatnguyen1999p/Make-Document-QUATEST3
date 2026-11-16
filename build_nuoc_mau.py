# -*- coding: utf-8 -*-
"""
Hồ sơ ẢO nước / bùn thải — để người dùng xem form trước khi làm hồ sơ thật (người dùng yêu cầu 16/11/2026:
"Tự dựng hồ sơ ảo, và chạy local, không chép lên Z, để t xem"). Mọi số liệu, khách hàng, ảnh đều GIẢ LẬP.

| Mã           | Loại                       | Căn cứ                                   | Câu công nhận            |
|--------------|----------------------------|------------------------------------------|--------------------------|
| Q326D999911  | Nước uống (máy lọc)        | QCVN 06-1:2010/BYT phụ lục II, III       | VILAS 004 + VIAS 004     |
| Q326D999912  | Nước sinh hoạt (vòi)       | QCVN 01-1:2024/BYT nhóm A                | VILAS 004 + VIAS 004     |
| Q326D999913  | Nước thải (VN + EN)        | QCVN 40:2011/BTNMT cột A, Kq 0,9, Kf 1,0 | VIMCERTS 078 (dòng đầu)  |
| Q326D999914  | Bùn thải (+ phiếu Excel)   | QCVN 07:2025/BNNMT                       | VIMCERTS 078 (dòng đầu)  |

Quy tắc áp (AGENTS mục 2, 3, 8C, 14 — người dùng chốt 16/11/2026): `QUATEST 3 3030:2025` có dấu cách; tên máy theo
danh mục `eq(...)`; tên chỉ tiêu nước uống / sinh hoạt theo phiếu KQTN; luôn có `Đánh giá:`; clo dư hiện trường
< 0,1 ghi `< 0,1`; `Từ ngày`; header nước BYT `- Ngày: `; nước thải / bùn: `mg/L`, "thông số", VIMCERTS dòng đầu.
Kết quả "Không đạt" in đậm (AGENTS mục 3); đánh giá tính từ số liệu, không ghi cứng.

Suy luận của AI (cần người dùng xem): 9.2 bùn thải bỏ ký hiệu ⁽¹⁾ "Thành phần nguy hại được chứng nhận VIMCERTS 078"
của mẫu cũ, thay bằng dòng `● Các thông số được chứng nhận VIMCERTS 078.` đầu Ghi chú, ngưỡng đánh số ⁽¹⁾ ⁽²⁾ ⁽³⁾
theo bảng 1, 2, 3 QCVN 07; ảnh 1 cột rộng 8 cm (mẫu cũ 1 ảnh căn giữa); nước thải VN + EN.

Chạy:  python build_nuoc_mau.py            (DOCX + XLSX, QA)
       python build_nuoc_mau.py --pdf      (thêm PDF qua Word / Excel COM + giấy tiêu đề UpChungThu)
       python build_nuoc_mau.py --only 999913
Kết quả: Input/9999xx/ThucHien, TBKQ, pdf. KHÔNG chép lên Z (mã thử Q326D9999xx — AGENTS mục 0).
"""

import os
import re
import sys
import copy
import argparse
import datetime

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ["QUATEST3_NETWORK_ROOT"] = r"C:\__khong_dong_bo_Z__"     # không bao giờ đồng bộ Z (hồ sơ ảo)

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm
from docx.text.paragraph import Paragraph

from Core.Models import Dossier, ClientInfo, PersonnelInfo
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, merge_same_vertical, style_note_tables, appendix_total_field
from Core.Styling.elements import add_document_ending_line
from Core.Styling.styles import set_cell_shading
from Core.Styling.tables import (format_table_structure, set_table_autofit_window, insert_image_to_cell,
                                 add_tbl_header, add_cant_split)
from Core.Tools.equipment_names import eq
from Core.Tools.qa_check import check_single_docx

REG_DATE = "03/11/2026"
SAMPLE_DATE = "05/11/2026"
KQTN_DATE = "12/11/2026"        # có KQTN -> ngày ban hành = ngày KQTN (AGENTS mục 2)
SEND_DATE = "05/11/2026"
REQ_DATE = "13/11/2026"
TODAY = datetime.date.today().strftime("%d/%m/%Y")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Nước Mẫu Thử Việt Nam",
    name_en="Vietnam Sample Water Testing Company Limited",
    address_vn="Lô A1, đường số 2, Khu công nghiệp Mẫu Thử, phường Tây Thạnh, Thành phố Hồ Chí Minh",
    address_en="Lot A1, road No. 2, Mau Thu Industrial Park, Tay Thanh ward, Ho Chi Minh City",
)
SITE_VN = "Nhà máy của Công ty TNHH Nước Mẫu Thử Việt Nam, lô A1, đường số 2, Khu công nghiệp Mẫu Thử, phường Tây Thạnh, Thành phố Hồ Chí Minh"
SITE_EN = "Factory of Vietnam Sample Water Testing Company Limited, lot A1, road No. 2, Mau Thu Industrial Park, Tay Thanh ward, Ho Chi Minh City"
WITNESS_VN = "Đại diện Công ty TNHH Nước Mẫu Thử Việt Nam – anh Nguyễn Văn An"
WITNESS_EN = "Representative of Vietnam Sample Water Testing Company Limited – Mr. Nguyen Van An"

FIELD = "Đo tại hiện trường"
FIELD_EN = "On-site measurement"
SENSE = "Cảm quan tại hiện trường"

# Mỗi chỉ tiêu: (tên VN, tên EN, tên gọn VN (câu Đánh giá), tên gọn EN, đơn vị, phương pháp VN, phương pháp EN, kết quả, mức)
# Kết quả / mức ghi dấu phẩy thập phân (VN); EN tự đổi. "^(2)" = ký hiệu ghi chú chỉ số trên.
SPECS = {
    "999911": dict(
        kind="nu", loai="nuoc uong", vi_tri=("Máy lọc nước nhà ăn", None), sample="999911 - NU",
        scope_vn="Lấy mẫu và giám định chất lượng nước uống", scope_en=None,
        procedure="QUATEST 3 3030:2025, N2.TC.N.07",
        standard_vn="QCVN 06-1:2010/BYT - Quy chuẩn kỹ thuật quốc gia đối với nước khoáng thiên nhiên và nước uống đóng chai",
        standard_en=None, equip=["hach_dr900"], seal=("08:45", "KT000011"), time=("08:30", "09:15"),
        chi_tieu_vn="- 11 chỉ tiêu hóa lý: asen, clo dư, độ đục, độ màu, mùi, amoni, chì, đồng, mangan, nitrat, nitrit\n"
                    "- 05 chỉ tiêu vi sinh tại QCVN 06-1:2010/BYT, phụ lục III",
        word="chỉ tiêu", limit_head="Giới hạn tối đa ^(1)",
        rows=[
            ("Hàm lượng asen (As)", None, "asen", None, "mg/L", "US EPA 200.8", None, "< 0,0005", "0,01"),
            ("Hàm lượng clo dư tự do (Cl₂)", None, "clo dư", None, "mg/L", FIELD, None, "< 0,1", "5"),
            ("Độ đục", None, "độ đục", None, "NTU", "SMEWW 2130B:2023", None, "< 0,5", "-"),
            ("Độ màu", None, "độ màu", None, "Pt.Co", "TCVN 6185:2015", None, "< 5", "-"),
            ("Mùi", None, "mùi", None, "-", SENSE, None, "Không có mùi lạ", "-"),
            ("Hàm lượng amoni (NH₃ và NH₄⁺ tính theo N)", None, "amoni", None, "mg/L", "US EPA 350.2 (So màu)", None, "< 0,1", "-"),
            ("Hàm lượng chì (Pb)", None, "chì", None, "mg/L", "US EPA 200.8", None, "< 0,003", "0,01"),
            ("Hàm lượng đồng (Cu)", None, "đồng", None, "mg/L", "US EPA 200.8", None, "< 0,01", "2"),
            ("Hàm lượng mangan (Mn)", None, "mangan", None, "mg/L", "US EPA 200.8", None, "< 0,01", "0,4"),
            ("Hàm lượng nitrat (NO₃⁻)", None, "nitrat", None, "mg/L", "SMEWW 4110B:2023", None, "< 0,4 ^(2)", "50"),
            ("Hàm lượng nitrit (NO₂⁻)", None, "nitrit", None, "mg/L", "SMEWW 4110B:2023", None, "< 0,02", "3"),
            ("Coliforms", None, "Coliforms", None, "CFU/ 250 mL", "ISO 9308-1:2014/Amd 1:2016", None, "3", "Không phát hiện"),
            ("Escherichia coli", None, "Escherichia coli", None, "CFU/ 250 mL", "ISO 9308-1:2014/Amd 1:2016", None, "< 1", "Không phát hiện"),
            ("Enterococci (Streptococci faecal)", None, "Enterococci", None, "CFU/ 250 mL", "ISO 7899-2:2000", None, "< 1", "Không phát hiện"),
            ("Pseudomonas aeruginosa", None, "Pseudomonas aeruginosa", None, "CFU/ 250 mL", "ISO 16266:2006", None, "< 1", "Không phát hiện"),
            ("Bào tử vi khuẩn kỵ khí khử sulfit (Clostridia)", None, "bào tử vi khuẩn kỵ khí khử sulfit", None, "CFU/ 50 mL", "ISO 6461-2:1986", None, "< 1", "Không phát hiện"),
        ],
        notes_vn=[("(1)", "Mức giới hạn tối đa cho phép quy định tại phụ lục II và phụ lục III, QCVN 06-1:2010/BYT."),
                  ("(2)", "Giới hạn định lượng của phương pháp thử."),
                  ("●", "Phương pháp thử nghiệm được công nhận ISO 17025 - VILAS 004."),
                  ("●", "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.")],
        intro_vn="So sánh kết quả thử nghiệm các chỉ tiêu của mẫu nước uống nêu trên so với mức quy định tương ứng tại "
                 "QCVN 06-1:2010/BYT, phụ lục II và phụ lục III:",
        kqtn="MTTN26099911", field_table=[("pH", None), ("Clo dư [mg/L]", "< 0,1"), ("Mùi", "Không có mùi lạ")],
    ),
    "999912": dict(
        kind="nsh", loai="nuoc sinh hoat", vi_tri=("Vòi nước nhà ăn", None), sample="999912 - NSH",
        scope_vn="Lấy mẫu và giám định chất lượng nước sinh hoạt", scope_en=None,
        procedure="QUATEST 3 3030:2025, N2.TC.N.01",
        standard_vn="QCVN 01-1:2024/BYT - Quy chuẩn kỹ thuật quốc gia về chất lượng nước sạch sử dụng cho mục đích sinh hoạt",
        standard_en=None, equip=["mettler_s2", "hach_dr900"], seal=None, time=("09:30", "10:00"),
        chi_tieu_vn="10 thông số nhóm A tại QCVN 01-1:2024/BYT",
        word="thông số", limit_head="Ngưỡng giới hạn cho phép ^(1)",
        rows=[
            ("pH", None, "pH", None, "-", FIELD, None, "7,45", "Trong khoảng (6,0 – 8,5)"),
            ("Clo dư tự do", None, "clo dư tự do", None, "mg/L", FIELD, None, "< 0,1", "Trong khoảng (0,2 – 1,0)"),
            ("Mùi", None, "mùi", None, "-", SENSE, None, "Không có mùi lạ", "Không có mùi lạ"),
            ("Độ màu", None, "độ màu", None, "Pt.Co", "TCVN 6185:2015", None, "< 5", "15"),
            ("Độ đục", None, "độ đục", None, "NTU", "SMEWW 2130B:2023", None, "< 0,5", "2"),
            ("Hàm lượng asen (As)", None, "asen", None, "mg/L", "US EPA 200.8", None, "< 0,0005", "0,01"),
            ("Chỉ số permanganate", None, "chỉ số permanganate", None, "mg/L", "TCVN 6186:1996", None, "< 1", "2"),
            ("Hàm lượng amoni (NH₃ và NH₄⁺ tính theo N)", None, "amoni", None, "mg/L", "US EPA 350.2 (So màu)", None, "< 0,1", "1"),
            ("Coliforms", None, "Coliforms", None, "CFU/ 100 mL", "ISO 9308-1:2014/Amd 1:2016", None, "< 1", "< 1"),
            ("Escherichia coli", None, "Escherichia coli", None, "CFU/ 100 mL", "ISO 9308-1:2014/Amd 1:2016", None, "< 1", "< 1"),
        ],
        notes_vn=[("(1)", "Ngưỡng giới hạn tối đa cho phép quy định tại QCVN 01-1:2024/BYT."),
                  ("●", "Phương pháp thử nghiệm được công nhận ISO 17025 - VILAS 004."),
                  ("●", "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.")],
        intro_vn="So sánh kết quả thử nghiệm các thông số của mẫu nước nêu trên so với mức quy định tương ứng tại "
                 "QCVN 01-1:2024/BYT:",
        kqtn="MTTN26099912",
        field_table=[("pH", "7,45"), ("Nhiệt độ [°C]", "29,8"), ("Clo dư [mg/L]", "< 0,1"), ("Mùi", "Không có mùi lạ")],
    ),
    "999913": dict(
        kind="nt", loai="nuoc thai", vi_tri=("Hố ga nước thải sau hệ thống xử lý", "Wastewater manhole after the treatment system"),
        sample="999913-NT",
        scope_vn="Lấy mẫu và giám định chất lượng nước thải", scope_en="Sampling and quality inspection of wastewater",
        procedure="QUATEST 3 3030:2025, N2.TC.N.06",
        standard_vn="QCVN 40:2011/BTNMT - Quy chuẩn kỹ thuật quốc gia về nước thải công nghiệp",
        standard_en="QCVN 40:2011/BTNMT - National Technical Regulation on Industrial Wastewater",
        equip=["mettler_s2", "hach_dr300"], seal=("10:20", "KT000013"), time=("10:10", "10:40"),
        chi_tieu_vn="Nhiệt độ, pH, BOD₅, COD, TSS, amoni, tổng nitơ, tổng phốt pho, clo dư, sắt, dầu mỡ khoáng, "
                    "Coliform (12 thông số)",
        chi_tieu_en="Temperature, pH, BOD₅, COD, TSS, ammonia, total nitrogen, total phosphorous, residual chlorine, iron, "
                    "mineral oil and fat, Coliform (12 parameters)",
        word="thông số", limit_head="Giới hạn tối đa ^(1)", limit_head_en="Limit ^(1)",
        rows=[
            ("Nhiệt độ", "Temperature", "nhiệt độ", "temperature", "°C", FIELD, FIELD_EN, "31,2", "40"),
            ("pH", "pH", "pH", "pH", "-", FIELD, FIELD_EN, "7,2", "6 – 9"),
            ("Nhu cầu oxy sinh hóa (BOD₅) quy về O₂", "Biological oxygen demand (BOD₅) as O₂", "BOD₅", "BOD₅", "mg/L",
             "TCVN 6001-1:2021", None, "12,4", "27"),
            ("Nhu cầu oxy hóa học (COD) quy về O₂", "Chemical oxygen demand (COD) as O₂", "COD", "COD", "mg/L",
             "SMEWW 5220D:2023", None, "38,5", "67,5"),
            ("Tổng chất rắn lơ lửng (TSS)", "Total suspended solids (TSS)", "TSS", "TSS", "mg/L", "SMEWW 2540D:2023", None, "21,0", "45"),
            ("Amoni (tính theo Nitơ)", "Ammonia (as N)", "amoni", "ammonia", "mg/L", "TCVN 5988:1995", None, "1,2", "4,5"),
            ("Tổng nitơ", "Total nitrogen", "tổng nitơ", "total nitrogen", "mg/L", "TCVN 6638:2000", None, "10,6", "18"),
            ("Tổng phốt pho", "Total phosphorous", "tổng phốt pho", "total phosphorous", "mg/L", "US EPA 200.7", None, "1,05", "3,6"),
            ("Clo dư", "Residual chlorine", "clo dư", "residual chlorine", "mg/L", FIELD, FIELD_EN, "1,4", "0,9"),
            ("Sắt", "Iron", "sắt", "iron", "mg/L", "SMEWW 3120B:2023", None, "0,32", "0,9"),
            ("Dầu mỡ khoáng", "Mineral oil and fat", "dầu mỡ khoáng", "mineral oil and fat", "mg/L", "SMEWW 5520B&F:2023", None,
             "< 1,4", "4,5"),
            ("Tổng Coliform", "Total Coliform", "Coliform", "Coliform", "MPN/ 100 mL", "SMEWW 9221B:2023", None, "2 400", "3 000"),
        ],
        notes_vn=[("●", "Các thông số được chứng nhận VIMCERTS 078."),
                  ("(1)", "Giới hạn tối đa cho phép tính toán theo QCVN 40:2011/BTNMT; cột A; hệ số Kq = 0,9; Kf = 1,0 "
                          "(thông tin do đại diện Công ty TNHH Nước Mẫu Thử Việt Nam cung cấp).")],
        notes_en=[("●", "Parameters are accredited to VIMCERTS 078."),
                  ("(1)", "Maximum permissible limit calculated according to QCVN 40:2011/BTNMT; column A; Kq = 0.9; "
                          "Kf = 1.0 (information provided by the representative of Vietnam Sample Water Testing Company "
                          "Limited).")],
        intro_vn="So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên với mức quy định tương ứng tại "
                 "QCVN 40:2011/BTNMT; cột A; hệ số Kq = 0,9; Kf = 1,0:",
        intro_en="Comparing the test result of parameters of wastewater sample mentioned above with the corresponding "
                 "limits in QCVN 40:2011/BTNMT; column A; Kq = 0.9; Kf = 1.0:",
        kqtn="MTTN26099913",
        field_table=[("pH", "7,2"), ("Nhiệt độ [°C]", "31,2"), ("Clo dư [mg/L]", "1,4")],
    ),
}

# Bùn thải: 3 bảng theo QCVN 07:2025/BNNMT (bảng 1 tính chất, bảng 2 ngâm chiết, bảng 3 hàm lượng tuyệt đối)
BT = dict(
    kind="bt", loai="bun thai", vi_tri=("Bao chứa bùn khô của hệ thống xử lý nước thải", None), sample="999914-BT",
    scope_vn="Lấy mẫu và giám định chất lượng bùn thải", procedure="QUATEST 3 3030:2025, N2.TC.B.01",
    standard_vn="QCVN 07:2025/BNNMT - Quy chuẩn kỹ thuật quốc gia về ngưỡng chất thải nguy hại", equip=[], seal=None,
    time=("13:30", "14:00"), kqtn="MTTN26099914",
    chi_tieu_vn="Tính chất nguy hại: tính kiềm/ tính axit (pH);\nThành phần nguy hại theo nồng độ ngâm chiết: antimon, "
                "asen, bari, cadimi, chì, thủy ngân, benzen, phenol;\nThành phần nguy hại theo hàm lượng tuyệt đối: "
                "tổng xyanua, PCB.",
    ph=("6,9", "pH ≥ 12,5 hoặc\npH ≤ 2,0"),
    leach=[("Antimon (Sb)", "US EPA 1311 & US EPA 200.7", "< 0,05", "≥ 15"),
           ("Asen (As)", "US EPA 1311 & US EPA 200.7", "< 0,5", "≥ 5,0"),
           ("Bari (Ba)", "US EPA 1311 & US EPA 200.7", "< 0,5", "≥ 100"),
           ("Cadimi (Cd)", "US EPA 1311 & US EPA 200.7", "< 0,01", "≥ 1,0"),
           ("Chì (Pb)", "US EPA 1311 & US EPA 200.7", "< 0,5", "≥ 5,0"),
           ("Thủy ngân (Hg)", "US EPA 1311 & SMEWW 3112B:2023", "< 0,02", "≥ 0,2"),
           ("Benzen", "US EPA 1311 & US EPA 5021A", "< 0,02", "≥ 0,5"),
           ("Phenol", "US EPA 1311 & US EPA 3510C & US EPA 3650B & US EPA 8041A", "< 0,5", "≥ 1 000")],
    # hàm lượng tuyệt đối: (tên, phương pháp, kết quả mẫu ướt mg/kg) — quy đổi về mẫu rắn khô ở phiếu Excel
    absolute=[("Tổng xyanua (CN⁻)", "US EPA 9013A & US EPA 9010C & US EPA 9014", "1,2", "≥ 590"),
              ("PCB\n+ PCB 28\n+ PCB 52\n+ PCB 101\n+ PCB 138\n+ PCB 153\n+ PCB 180",
               "US EPA 3550C & US EPA 3630C & US EPA 8270E",
               "< 0,003\n< 0,003\n< 0,003\n< 0,003\n< 0,003\n< 0,003", "≥ 5,0")],
    moisture="86,6",
)


# ---------------------------------------------------------------------------
# Tiện ích chữ / đánh giá
# ---------------------------------------------------------------------------
def en_num(s):
    return re.sub(r"(\d),(\d)", r"\1.\2", s or "")


def _num(s):
    s = (s or "").replace("\u00a0", " ").strip()
    m = re.match(r"^([<>≤≥])?\s*([\d ]+(?:,\d+)?)", s)
    if not m:
        return None, None
    return m.group(1), float(m.group(2).replace(" ", "").replace(",", "."))


def evaluate(result, limit):
    """True đạt / False không đạt / None không đánh giá (mức '-')."""
    lim = (limit or "").strip()
    res = re.sub(r"\^\(\d\)", "", result or "").strip()
    if lim in ("", "-"):
        return None
    if lim.lower() in ("không phát hiện", "< 1"):
        op, v = _num(res)
        return op == "<" or v == 0
    if lim.lower().startswith("không có mùi"):
        return res.lower().startswith("không có mùi")
    rng = re.search(r"([\d,]+)\s*[–-]\s*([\d,]+)", lim)
    op, v = _num(res)
    if v is None:
        return None
    if rng:
        lo, hi = (float(x.replace(",", ".")) for x in rng.groups())
        if op == "<":                       # "< 0,1" với khoảng (0,2 – 1,0) -> chắc chắn dưới khoảng
            return False if v <= lo else None
        if op == ">":
            return False if v >= hi else None
        return lo <= v <= hi
    _, lv = _num(lim.lstrip("≤ "))
    if lv is None:
        return None
    if op == "<":
        return True if v <= lv else None
    if op == ">":
        return False if v >= lv else None
    return v <= lv


def _set_rfonts(r):
    r.font.name = "Times New Roman"
    rpr = r._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), "Times New Roman")


def add_runs(p, text, size, bold=False, italic=False):
    """Ghi chữ; '^(n)' -> '(n)' chỉ số trên (AGENTS mục 3)."""
    for i, piece in enumerate(re.split(r"\^(\(\d\))", text)):
        if not piece:
            continue
        r = p.add_run(piece)
        r.font.size, r.bold, r.italic = Pt(size), bold, italic
        _set_rfonts(r)
        if i % 2 == 1:
            r.font.superscript = True


def add_mixed(p, parts, size):
    """parts = [(text, bold, italic), …]."""
    for text, b, it in parts:
        add_runs(p, text, size, b, it)


def _ptext(el):
    import unicodedata
    return unicodedata.normalize("NFC", "".join(t.text or "" for t in el.iter(qn("w:t"))))


def _keep(p_el, lines=False):
    pPr = p_el.get_or_add_pPr()
    tag = "w:keepLines" if lines else "w:keepNext"
    if pPr.find(qn(tag)) is None:
        pPr.insert(0, OxmlElement(tag))


def _keep_rows(tbl_el, rows):
    for tr in rows:
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)


class Body:
    """Ghi tuần tự các khối vào trước ``anchor`` (mặc định sectPr cuối)."""

    def __init__(self, doc, text_proto, size=13):
        self.doc, self.proto, self.size = doc, text_proto, size
        self.anchor = doc.element.body.find(qn("w:sectPr"))

    def emit(self, el):
        self.anchor.addprevious(el)
        return el

    def para(self, text="", bold=False, italic=False, keep_next=False, space_before=None, parts=None,
             indent_cm=None, tab_cm=None, align=None):
        el = copy.deepcopy(self.proto)
        for ch in [c for c in el if c.tag != qn("w:pPr")]:
            el.remove(ch)
        self.emit(el)
        p = Paragraph(el, self.doc._body)
        if parts:
            add_mixed(p, parts, self.size)
        else:
            add_runs(p, text, self.size, bold, italic)
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        if indent_cm is not None:
            pf.left_indent = Cm(indent_cm)
        if tab_cm:
            pf.tab_stops.add_tab_stop(Cm(tab_cm))
        if keep_next:
            pf.keep_with_next = True
        if space_before is not None:
            pf.space_before = Pt(space_before)
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return p

    def table(self, rows, cols, widths_pct, header_rows=1, row_height=440):
        t = self.doc.add_table(rows=rows, cols=cols)
        self.emit(t._tbl)
        format_table_structure(t, row_height=row_height)
        for mar in list(t._tbl.iter(qn("w:tcMar"))):
            mar.getparent().remove(mar)
        grid = t._tbl.tblGrid
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths_pct):
            gc.set(qn("w:w"), str(round(9355 * w / 100)))
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
            add_cant_split(row)
        for r in t.rows[:header_rows]:
            add_tbl_header(r)
            for c in r.cells:
                set_cell_shading(c, "D9D9D9")
        return t

    def photo(self, path, width_cm=7.0):          # 7 cm: đủ chỗ để mục 9.2 bắt đầu cùng trang (16/11/2026)
        t = self.doc.add_table(rows=1, cols=1)
        self.emit(t._tbl)
        set_table_autofit_window(t)
        insert_image_to_cell(t.rows[0].cells[0], path, width=Cm(width_cm))
        add_cant_split(t.rows[0])
        return t


def cell_text(cell, text, bold=False, italic=False, align="center", size=13):
    for p in list(cell.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = cell.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    for i, line in enumerate(str(text).split("\n")):
        p = p0 if i == 0 else cell.add_paragraph()
        p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
        pf = p.paragraph_format
        pf.space_before, pf.space_after, pf.line_spacing = Pt(1), Pt(1), 1.0
        pf.left_indent = pf.first_line_indent = Cm(0)
        add_runs(p, line, size, bold, italic)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def merge_text_runs(t, col, start):
    """Gộp dọc ô mức liền nhau cùng chữ khi là mức chữ ('-', 'Không phát hiện', '< 1'…); mức là số (0,9 / 0,9 của
    2 thông số khác nhau) giữ riêng từng ô."""
    rows = t.rows
    r = start
    while r < len(rows):
        txt = rows[r].cells[col].text.strip()
        e = r
        while e + 1 < len(rows) and rows[e + 1].cells[col].text.strip() == txt:
            e += 1
        if e > r and txt and not re.fullmatch(r"[≤≥<>]?\s*[\d ]+([,.]\d+)?", txt):
            for rr in range(r + 1, e + 1):
                for tnode in rows[rr].cells[col]._tc.iter(qn("w:t")):
                    tnode.text = ""
            m = rows[r].cells[col].merge(rows[e].cells[col])
            for p in m.paragraphs[1:]:
                if not p.text.strip():
                    p._p.getparent().remove(p._p)
            m.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        r = e + 1


def note_table(body, proto, lines):
    """Bảng 'Ghi chú/Note' theo template: '(n)' ở cột ký hiệu in chỉ số trên; style 11pt nghiêng làm sau."""
    el = copy.deepcopy(proto)
    rows = el.findall(qn("w:tr"))
    item = rows[-1]
    for r in rows[1:]:
        el.remove(r)
    for marker, text in lines:
        tr = copy.deepcopy(item)
        el.append(tr)
        for k, (tc, val) in enumerate(zip(tr.findall(qn("w:tc")), (marker, text))):
            ps = tc.findall(qn("w:p"))
            for extra in ps[1:]:
                tc.remove(extra)
            p_el = ps[0]
            runs = p_el.findall(qn("w:r"))
            proto_r = copy.deepcopy(runs[0]) if runs else OxmlElement("w:r")
            for ch in [c for c in p_el if c.tag != qn("w:pPr")]:
                p_el.remove(ch)
            r = copy.deepcopy(proto_r)
            for c in [c for c in r if c.tag != qn("w:rPr")]:
                r.remove(c)
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            for va in rpr.findall(qn("w:vertAlign")):
                rpr.remove(va)
            if k == 0 and re.fullmatch(r"\(\d\)", val):
                va = OxmlElement("w:vertAlign")
                va.set(qn("w:val"), "superscript")
                rpr.append(va)
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = val
            r.append(t)
            p_el.append(r)
    # Đoạn ngăn 1 pt giữa bảng kết quả và bảng Ghi chú: không có thì Word gộp 2 bảng (dòng tiêu đề lặp sang
    # trang sau trên bảng Ghi chú); thu về line 20 exact, before/after 0 (AGENTS mục 15 [Note Gap])
    gap = copy.deepcopy(body.proto)
    for ch in [c for c in gap if c.tag != qn("w:pPr")]:
        gap.remove(ch)
    pPr = gap.get_or_add_pPr() if hasattr(gap, "get_or_add_pPr") else gap.find(qn("w:pPr"))
    for sp_el in pPr.findall(qn("w:spacing")):
        pPr.remove(sp_el)
    sp_el = OxmlElement("w:spacing")
    for a, v in (("w:before", "0"), ("w:after", "0"), ("w:line", "20"), ("w:lineRule", "exact")):
        sp_el.set(qn(a), v)
    pPr.append(sp_el)
    _keep(gap)
    body.emit(gap)
    body.emit(el)
    _keep_rows(el, el.findall(qn("w:tr"))[:-1])
    return el


def fake_photo(path, label):
    """Ảnh giả lập (hồ sơ ảo) — khung ngang 4:3 có chữ."""
    if os.path.exists(path):
        return path
    from PIL import Image, ImageDraw, ImageFont
    os.makedirs(os.path.dirname(path), exist_ok=True)
    im = Image.new("RGB", (1200, 900), (205, 220, 232))
    dr = ImageDraw.Draw(im)
    try:
        f1 = ImageFont.truetype(r"C:\Windows\Fonts\times.ttf", 64)
        f2 = ImageFont.truetype(r"C:\Windows\Fonts\times.ttf", 44)
    except OSError:
        f1 = f2 = ImageFont.load_default()
    dr.rectangle((40, 40, 1160, 860), outline=(60, 90, 120), width=8)
    dr.text((600, 380), "ẢNH GIẢ LẬP", fill=(160, 30, 30), font=f1, anchor="mm")
    dr.text((600, 480), label, fill=(30, 40, 60), font=f2, anchor="mm")
    im.save(path, "JPEG", quality=85)
    return path


# ---------------------------------------------------------------------------
# Dossier + bìa
# ---------------------------------------------------------------------------
def make_dossier(code6, sp):
    st, en = sp["time"]
    return Dossier(
        code=f"Q326D{code6}", cert_date=KQTN_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {st} đến {en}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False, dossier_type="cleanroom",       # khung bìa phòng sạch, phụ lục dựng lại
        scope_vn=sp["scope_vn"], scope_en=sp.get("scope_en") or sp["scope_vn"],
        location_desc_vn=SITE_VN, location_desc_en=SITE_EN,
        procedure_vn=sp["procedure"], procedure_en=sp["procedure"],
        standards_vn=sp["standard_vn"], standards_en=sp.get("standard_en") or sp["standard_vn"],
        client=CLIENT, personnel=PersonnelInfo(executor="Huy"),
        equipments=[eq(k) for k in sp["equip"]], locations=[], photos=[], results_data={},
        input_dir=os.path.join(WORKSPACE, "Input", code6), output_dir="", network_dir="",
    )


def fix_cover(doc, is_en, header_colon):
    """Thiết bị trống -> '-'; header nước BYT 'Ngày:' (AGENTS mục 2 ngoại lệ)."""
    t0 = doc.tables[0]
    rows = list(t0.rows)
    for k, row in enumerate(rows[:-1]):
        label = row.cells[0].text.strip().lower()
        nxt = rows[k + 1]
        if ("thiết bị" in label or "equipment" in label) and PLACEHOLDER in nxt.cells[0].text:
            # thiết bị trống: template để danh sách ở hàng dưới -> ghi '-' ngay hàng nhãn, bỏ hàng '………'
            val = row.cells[-1]                                         # ô giá trị = ô cuối (VN / EN lưới cột khác nhau)
            for p in list(val.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            p0 = val.paragraphs[0]
            for r in list(p0.runs):
                r._r.getparent().remove(r._r)
            add_runs(p0, "-", 13)
            nxt._tr.getparent().remove(nxt._tr)
            break
    rows = list(t0.rows)
    for k, row in enumerate(rows[:-1]):
        label = row.cells[0].text.strip().lower()
        nxt = rows[k + 1]
        lines = [x for x in nxt.cells[0].text.split("\n") if x.strip()]
        if ("căn cứ" in label or "applicable" in label) and len(lines) == 1 and lines[0].startswith("- "):
            # 1 căn cứ: ghi ngay sau ':' của hàng nhãn, không gạch đầu dòng (AGENTS mục 14: '- …' chỉ khi từ 2 mục)
            val = row.cells[-1]                                         # ô giá trị = ô cuối (VN / EN lưới cột khác nhau)
            for p in list(val.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            p0 = val.paragraphs[0]
            for r in list(p0.runs):
                r._r.getparent().remove(r._r)
            add_runs(p0, lines[0][2:].strip().rstrip(";"), 13)
            p0.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            nxt._tr.getparent().remove(nxt._tr)
            break
    if header_colon and not is_en:
        for sec in doc.sections:
            for part in (sec.header, sec.first_page_header):
                if part.is_linked_to_previous:
                    continue
                for p_el in part._element.iter(qn("w:p")):
                    ts = list(p_el.iter(qn("w:t")))
                    full = "".join(t.text or "" for t in ts)
                    if "Số:" not in full or "Ngày:" in full:
                        continue
                    for t in ts:
                        if t.text and re.search(r"Ngày\s*$|Ngày\s+\d", t.text):
                            t.text = re.sub(r"Ngày(\s*)", r"Ngày: ", t.text, count=1).replace("Ngày:  ", "Ngày: ")
                            break


def appendix_protos(doc):
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    sect_i = next(i for i, el in enumerate(kids)
                  if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None)
    after = kids[sect_i + 1:]
    heading = next(el for el in after if el.tag == qn("w:p")
                   and _ptext(el).strip() in ("Performance content", "Nội dung thực hiện"))
    text_proto = copy.deepcopy(heading.getnext())
    note_proto = copy.deepcopy(next(el for el in after if el.tag == qn("w:tbl")
                                    and _ptext(el).strip().startswith(("Note", "Ghi chú"))))
    head_proto = copy.deepcopy(heading)
    for el in after:
        if el.tag != qn("w:sectPr"):
            body_el.remove(el)
    return text_proto, note_proto, head_proto


def make_head(body, head_proto):
    def head(text):
        el = copy.deepcopy(head_proto)
        for r in el.findall(qn("w:r"))[1:]:
            el.remove(r)
        el.find(qn("w:r")).find(qn("w:t")).text = text
        body.emit(el)
        _keep(el)
        return el
    return head


def sub_head(body, text):
    return body.para(text, bold=True, keep_next=True, space_before=6)


def assessment(body, intro, lines, is_en=False):
    """'Đánh giá:' + câu so sánh + các dòng '+ … : …' (cụm từ đánh giá đậm nghiêng — AGENTS mục 3)."""
    p = body.para("Assessment:" if is_en else "Đánh giá:", bold=True, keep_next=True, space_before=6)
    body.para(intro, keep_next=True)
    for k, (pre, key, val) in enumerate(lines):
        q = body.para(parts=[("+ " + (f"{pre} " if pre else ""), False, False), (key, True, True), ("\t: " + val, False, False)],
                      indent_cm=0.75, tab_cm=7.0, keep_next=k < len(lines) - 1)
        _keep(q._p, lines=True)
    return p


def eval_lines(rows, word, is_en):
    """Chia chỉ tiêu theo kết quả đánh giá -> 3 dòng (không đánh giá chỉ khi có)."""
    no_eval, fail, ok = [], [], []
    for r in rows:
        name = (r[3] if is_en else r[2]) or r[2]
        v = evaluate(r[7], r[8])
        (no_eval if v is None else ok if v else fail).append(name)
    lines = []
    if is_en:
        if no_eval:
            lines.append(("", "Non-evaluated parameter", ", ".join(no_eval)))
        lines.append(("", "Non-compliant parameter", ", ".join(fail) or "-"))
        lines.append(("", "Compliant parameter", "all parameters" if not (fail or no_eval) else "the remaining parameters"))
        return lines
    if no_eval:
        lines.append((word.capitalize(), "không đánh giá", ", ".join(no_eval)))
    lines.append((word.capitalize(), "không đạt", ", ".join(fail) or "-"))
    lines.append((word.capitalize(), "đạt", f"tất cả các {word}" if not (fail or no_eval) else f"các {word} còn lại"))
    return lines


# ---------------------------------------------------------------------------
# TBKQ nước (uống / sinh hoạt / thải)
# ---------------------------------------------------------------------------
def photos_for(code6, sp):
    d = os.path.join(WORKSPACE, "Input", code6, "Photo")
    loc = fake_photo(os.path.join(d, "Vi tri lay mau.jpg"), f"Vị trí lấy mẫu {sp['sample']}")
    seal = fake_photo(os.path.join(d, "Mau niem phong.jpg"), f"Mẫu niêm phong {sp['seal'][1]}") if sp.get("seal") else None
    return loc, seal


def build_tbkq_water(d, code6, sp, is_en):
    src = CleanroomHandler(d).build_tbkq(is_en=is_en)
    doc = Document(src)
    fix_cover(doc, is_en, header_colon=sp["kind"] in ("nu", "nsh"))
    text_proto, note_proto, head_proto = appendix_protos(doc)
    b = Body(doc, text_proto)
    head = make_head(b, head_proto)
    L = (lambda vn, en: en if is_en else vn)
    num = (lambda s: en_num(s) if is_en else s)
    loc_photo, seal_photo = photos_for(code6, sp)

    head(L("Nội dung thực hiện", "Performance content"))
    who = CLIENT.name_en if is_en else CLIENT.name_vn
    b.para(L(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm theo yêu cầu của {who}, cụ thể như sau:",
             f"- Sampling location and parameters were indicated by {who}, detailed as follows:"), keep_next=True)
    t = b.table(2, 3, [28, 22, 50])
    for j, h in enumerate(L(["Vị trí lấy mẫu", "Ký hiệu mẫu", "Chỉ tiêu thử nghiệm"],
                            ["Sampling location", "Sampling code", "Parameter"])):
        cell_text(t.rows[0].cells[j], h)
    cell_text(t.rows[1].cells[0], L(sp["vi_tri"][0], sp["vi_tri"][1] or sp["vi_tri"][0]))
    cell_text(t.rows[1].cells[1], sp["sample"])
    ct = sp.get("chi_tieu_en") if is_en else sp["chi_tieu_vn"]
    cell_text(t.rows[1].cells[2], ct, align="left" if "\n" in ct or len(ct) > 60 else "center")
    _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:1])
    if sp.get("seal"):
        b.para(L("- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong:",
                 "- Information on the sampling process and photo of sealed sample:"), keep_next=True, space_before=6)
        t = b.table(2, 4, [40, 18, 20, 22])
        for j, h in enumerate(L(["Người chứng kiến", "Ký hiệu mẫu", "Thời điểm lấy mẫu", "Số niêm phong"],
                                ["Witness", "Sampling code", "Sampling time", "Sealing number"])):
            cell_text(t.rows[0].cells[j], h)
        for j, v in enumerate([L(WITNESS_VN, WITNESS_EN), sp["sample"], sp["seal"][0], sp["seal"][1]]):
            cell_text(t.rows[1].cells[j], v)
        _keep_rows(t._tbl, t._tbl.findall(qn("w:tr")))
        b.photo(seal_photo)
    else:
        b.para(L("- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong: -",
                 "- Information on the sampling process and photo of sealed sample: -"), space_before=6)
    b.para(L("- Hình ảnh vị trí lấy mẫu:", "- Photo of sampling location:"), keep_next=True, space_before=6)
    b.photo(loc_photo)

    head(L("Kết quả giám định", "Inspection result"))
    rows = sp["rows"]
    word_col = {"chỉ tiêu": "Chỉ tiêu", "thông số": "Thông số"}[sp["word"]]
    hdr = L(["STT", word_col, "Đơn vị", "Phương pháp", "Kết quả", sp["limit_head"]],
            ["No.", "Parameter", "Unit", "Method", "Result", sp.get("limit_head_en", "Limit ^(1)")])
    widths = [7, 23, 13, 20, 15, 22] if sp["kind"] == "nsh" else [7, 31, 12, 22, 13, 15]   # NSH: mức "Trong khoảng (…)" dài
    t = b.table(1 + len(rows), 6, widths)
    for j, h in enumerate(hdr):
        cell_text(t.rows[0].cells[j], h)
    for i, r in enumerate(rows, start=1):
        name = (r[1] or r[0]) if is_en else r[0]
        meth = (r[6] or r[5]) if is_en else r[5]
        ok = evaluate(r[7], r[8])
        cells = t.rows[i].cells
        cell_text(cells[0], str(i))
        cell_text(cells[1], name, align="left")
        cell_text(cells[2], r[4])
        cell_text(cells[3], meth)
        cell_text(cells[4], num(r[7]), bold=(ok is False))          # Không đạt -> in đậm (AGENTS mục 3)
        lim = r[8]
        if is_en:
            lim = {"Không phát hiện": "Not detected", "Không có mùi lạ": "No strange odor"}.get(lim, en_num(lim))
            lim = lim.replace("Trong khoảng", "In range")
        cell_text(cells[5], lim)
    merge_same_vertical(t, 3, 1)                                     # gộp dọc phương pháp trùng nhau
    merge_text_runs(t, 5, 1)                                         # gộp dọc mức chữ ('-', 'Không phát hiện'), không gộp số
    for row in t.rows[1:]:
        for c in row.cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(13)
    trs = t._tbl.findall(qn("w:tr"))
    # Tiêu đề 9.2 + dòng tiêu đề bảng + 4 dòng đầu đi liền; phần còn lại được nối sang trang sau (lặp tiêu đề).
    # Không giữ trọn cả bảng ngắn: 999912 từng để trống nửa trang phụ lục 1 (người dùng báo 16/11/2026).
    _keep_rows(t._tbl, trs[:5])
    _keep_rows(t._tbl, trs[-2:])                                     # 2 dòng cuối đi liền Ghi chú + Đánh giá (AGENTS 3)
    notes = sp.get("notes_en") if is_en else sp["notes_vn"]
    if not is_en and sp["kind"] == "nu" and not any("^(2)" in r[7] for r in rows):
        notes = [n for n in notes if n[0] != "(2)"]
    nt = note_table(b, note_proto, notes)
    _keep_rows(nt, nt.findall(qn("w:tr")))                            # Ghi chú -> Đánh giá cùng trang
    assessment(b, sp.get("intro_en") if is_en else sp["intro_vn"], eval_lines(rows, sp["word"], is_en), is_en)
    return finish_tbkq(doc, src, d, sp, is_en)


def finish_tbkq(doc, src, d, sp, is_en):
    style_note_tables(doc)
    add_document_ending_line(doc)
    appendix_total_field(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - {sp['loai']} - {d.code}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    print(f"[build_nuoc_mau] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


# ---------------------------------------------------------------------------
# TBKQ bùn thải
# ---------------------------------------------------------------------------
def dry_value(wet, moisture):
    if wet.strip().startswith("<"):
        return wet
    v = float(wet.replace(",", ".")) * 100 / (100 - float(moisture.replace(",", ".")))
    return f"{v:.2f}".replace(".", ",")


def build_tbkq_sludge(d, code6, sp):
    src = CleanroomHandler(d).build_tbkq(is_en=False)
    doc = Document(src)
    fix_cover(doc, False, header_colon=False)
    text_proto, note_proto, head_proto = appendix_protos(doc)
    b = Body(doc, text_proto)
    head = make_head(b, head_proto)
    loc_photo, _ = photos_for(code6, sp)

    head("Nội dung thực hiện")
    b.para(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm theo yêu cầu của {CLIENT.name_vn}, cụ thể như sau:", keep_next=True)
    t = b.table(2, 3, [28, 17, 55])
    for j, h in enumerate(["Vị trí lấy mẫu", "Ký hiệu mẫu", "Chỉ tiêu thử nghiệm"]):
        cell_text(t.rows[0].cells[j], h)
    cell_text(t.rows[1].cells[0], sp["vi_tri"][0])
    cell_text(t.rows[1].cells[1], sp["sample"])
    cell_text(t.rows[1].cells[2], sp["chi_tieu_vn"], align="left")
    b.para("- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong: -", space_before=6)
    b.para("- Hình ảnh vị trí lấy mẫu:", keep_next=True, space_before=6)
    b.photo(loc_photo)

    head("Kết quả giám định")
    vim = ("●", "Các thông số được chứng nhận VIMCERTS 078.")
    # 9.2.1
    sub_head(b, "9.2.1. Tính chất nguy hại")
    t = b.table(2, 4, [30, 20, 14, 36])
    for j, h in enumerate(["Tính chất nguy hại", "Phương pháp", "Kết quả", "Ngưỡng CTNH ^(1)"]):
        cell_text(t.rows[0].cells[j], h)
    for j, v in enumerate(["Tính kiềm/ tính axit (pH)", "US EPA 9045D", sp["ph"][0], sp["ph"][1]]):
        cell_text(t.rows[1].cells[j], v)
    _keep_rows(t._tbl, t._tbl.findall(qn("w:tr")))
    nt = note_table(b, note_proto, [vim, ("(1)", "Ngưỡng chất thải nguy hại tại QCVN 07:2025/BNNMT, bảng 1.")])
    _keep_rows(nt, nt.findall(qn("w:tr"))[:-1])

    def comp_table(title, rows, unit, mark, note):
        sub_head(b, title)
        t = b.table(1 + len(rows), 5, [8, 30, 30, 15, 17])
        for j, h in enumerate(["STT", "Thành phần nguy hại", "Phương pháp", f"Kết quả, {unit}",
                               f"Ngưỡng CTNH ^({mark}), {unit}"]):
            cell_text(t.rows[0].cells[j], h)
        for i, (name, meth, res, thr) in enumerate(rows, start=1):
            c = t.rows[i].cells
            cell_text(c[0], str(i))
            cell_text(c[1], name, align="left")
            cell_text(c[2], meth.replace(" & ", " &\n"))
            cell_text(c[3], res)
            cell_text(c[4], thr)
        merge_same_vertical(t, 2, 1)
        trs = t._tbl.findall(qn("w:tr"))
        _keep_rows(t._tbl, trs[:4])
        _keep_rows(t._tbl, trs[-2:])
        nt = note_table(b, note_proto, [vim, (f"({mark})", note)])
        _keep_rows(nt, nt.findall(qn("w:tr"))[:-1])

    comp_table("9.2.2. Nồng độ ngâm chiết của các thành phần nguy hại", sp["leach"], "mg/L", 2,
               "Ngưỡng chất thải nguy hại tính theo nồng độ ngâm chiết tại QCVN 07:2025/BNNMT, bảng 2.")
    absolute = [(n, m, dry_value(r, sp["moisture"]) if "\n" not in r else r, thr) for n, m, r, thr in sp["absolute"]]
    comp_table("9.2.3. Hàm lượng tuyệt đối của các thành phần nguy hại", absolute, "ppm", 3,
               "Ngưỡng chất thải nguy hại tính theo hàm lượng tuyệt đối tại QCVN 07:2025/BNNMT, bảng 3.")

    def hazardous(rows):
        hi = []
        for name, _m, res, thr in rows:
            op, v = _num(res.split("\n")[0])
            _, lv = _num(thr.lstrip("≥> "))
            if v is not None and lv is not None and op != "<" and v >= lv:
                hi.append(name.split("\n")[0])
        return hi

    b.para("Đánh giá:", bold=True, keep_next=True, space_before=6)
    ph_ok = 2.0 < _num(sp["ph"][0])[1] < 12.5          # ngưỡng CTNH: pH ≥ 12,5 hoặc pH ≤ 2,0
    b.para("Độ pH " + ("không vượt" if ph_ok else "vượt") + " ngưỡng nguy hại tương ứng nêu tại QCVN 07:2025/BNNMT.",
           keep_next=True)
    for what, rows in (("nồng độ ngâm chiết", sp["leach"]), ("hàm lượng tuyệt đối", absolute)):
        hi = hazardous(rows)
        b.para(f"So sánh kết quả thử nghiệm {what} của các thành phần nguy hại với các ngưỡng nguy hại tương ứng nêu "
               "tại QCVN 07:2025/BNNMT:", keep_next=True)
        for k, (key, val) in enumerate((("cao hơn mức quy định", ", ".join(hi) or "-"),
                                        ("thấp hơn mức quy định", "Tất cả các chỉ tiêu" if not hi else "các chỉ tiêu còn lại"))):
            q = b.para(parts=[("+ Chỉ tiêu ", False, False), (key, True, True), ("\t: " + val, False, False)],
                       indent_cm=0.75, tab_cm=7.5, keep_next=(what == "nồng độ ngâm chiết" or k == 0))
            _keep(q._p, lines=True)
    return finish_tbkq(doc, src, d, sp, False)


# ---------------------------------------------------------------------------
# BCKQ (BM3030.03) — dùng khung BCKQ phòng sạch, thay mục 4.5 / 4.6 / Đánh giá / Gửi mẫu
# ---------------------------------------------------------------------------
def build_bckq(d, code6, sp):
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - {sp['loai']} - {d.code}.docx")
    doc = Document(out)
    for row in doc.tables[0].rows:                                    # thiết bị trống -> '-'
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs:
                    if PLACEHOLDER in r.text and "thiết bị" in row.cells[0].text.lower():
                        r.text = r.text.replace(PLACEHOLDER, "-")
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    h1 = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Nội dung thực hiện"))
    h2 = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Kết quả đo đạc"))
    gdv = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Báo cáo của GĐV"))
    text_proto = copy.deepcopy(h1.getnext())
    i1, i2, ig = kids.index(h1), kids.index(h2), kids.index(gdv)
    for el in kids[i1 + 1:i2] + kids[i2 + 1:ig]:
        body_el.remove(el)
    b = Body(doc, text_proto)
    loc_photo, seal_photo = photos_for(code6, sp)

    b.anchor = h2
    b.para(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm được chỉ định bởi đại diện {CLIENT.name_vn}, cụ thể như sau:",
           keep_next=True)
    t = b.table(2, 3, [28, 22, 50])
    for j, h in enumerate(["Vị trí thực hiện", "Kí hiệu mẫu", "Chỉ tiêu thử nghiệm"]):
        cell_text(t.rows[0].cells[j], h)
    cell_text(t.rows[1].cells[0], sp["vi_tri"][0])
    cell_text(t.rows[1].cells[1], sp["sample"])
    cell_text(t.rows[1].cells[2], sp["chi_tieu_vn"], align="left")
    if sp.get("seal"):
        b.para("- Thông tin về quá trình lấy mẫu niêm phong:", keep_next=True, space_before=6)
        t = b.table(2, 4, [40, 18, 20, 22])
        for j, h in enumerate(["Người chứng kiến", "Ký hiệu mẫu", "Thời điểm lấy mẫu", "Số niêm phong"]):
            cell_text(t.rows[0].cells[j], h)
        for j, v in enumerate([WITNESS_VN, sp["sample"], sp["seal"][0], sp["seal"][1]]):
            cell_text(t.rows[1].cells[j], v)
        b.photo(seal_photo)
    else:
        b.para("- Thông tin về quá trình lấy mẫu niêm phong: -", space_before=6)
    b.para("- Hình ảnh vị trí lấy mẫu:", keep_next=True, space_before=6)
    b.photo(loc_photo)

    b.anchor = gdv
    _keep(h2)
    if sp["kind"] == "bt":
        b.para(f"- KQTN hàm lượng tuyệt đối và nồng độ ngâm chiết của mẫu chất thải: xem tại phiếu KQTN số {sp['kqtn']} "
               f"ngày {KQTN_DATE};")
        b.para(f"- Tính toán hàm lượng tuyệt đối của chỉ tiêu phân tích theo mẫu rắn khô: xem tại phiếu Tính toán hàm "
               f"lượng tuyệt đối theo mẫu rắn khô số {d.code} ngày {KQTN_DATE}.")
    else:
        b.para(f"- KQTN mẫu {sp['sample']}\t: xem tại phiếu KQTN số {sp['kqtn']}", tab_cm=5.5)
        b.para("- Kết quả đo tại hiện trường:", keep_next=True)
        ft = [(h, v) for h, v in sp["field_table"] if v is not None]
        t = b.table(3, 1 + len(ft), [20] + [80 / len(ft)] * len(ft))
        cell_text(t.rows[0].cells[0], "Kí hiệu mẫu")
        cell_text(t.rows[1].cells[0], sp["sample"])
        cell_text(t.rows[2].cells[0], "Thiết bị sử dụng")
        names = {"pH": "Mettler Toledo Seven2Go S2", "Nhiệt độ [°C]": "Mettler Toledo Seven2Go S2",
                 "Clo dư [mg/L]": "Hach DR300" if "hach_dr300" in sp["equip"] else "Hach DR900", "Mùi": "-"}
        for j, (h, v) in enumerate(ft, start=1):
            cell_text(t.rows[0].cells[j], h)
            cell_text(t.rows[1].cells[j], v)
            cell_text(t.rows[2].cells[j], names[h], italic=names[h] != "-")
        _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:-1])
    if sp["kind"] == "bt":
        b.para("Đánh giá:", bold=True, keep_next=True, space_before=6)
        b.para("Độ pH không vượt ngưỡng nguy hại tương ứng nêu tại QCVN 07:2025/BNNMT.", keep_next=True)
        for what in ("nồng độ ngâm chiết", "hàm lượng tuyệt đối"):
            b.para(f"So sánh kết quả thử nghiệm các thành phần nguy hại theo {what} của chất thải với các ngưỡng nguy "
                   "hại tương ứng nêu tại QCVN 07:2025/BNNMT:", keep_next=True)
            b.para(parts=[("+ Chỉ tiêu ", False, False), ("cao hơn mức quy định", True, True), ("\t: -", False, False)],
                   indent_cm=0.75, tab_cm=7.5, keep_next=True)
            b.para(parts=[("+ Chỉ tiêu ", False, False), ("thấp hơn mức quy định", True, True),
                          ("\t: Tất cả các chỉ tiêu", False, False)], indent_cm=0.75, tab_cm=7.5)
    else:
        assessment(b, sp["intro_vn"], eval_lines(sp["rows"], sp["word"], False))
    from Core.Tools.kqtn_late import tre_han_text
    b.para("Gửi mẫu thử nghiệm:", bold=True, keep_next=True, space_before=6)
    for line in (f"Phiếu KQTN số {sp['kqtn']} ngày {KQTN_DATE}", f"Ngày gửi mẫu\t: {SEND_DATE}",
                 f"Ngày yêu cầu cấp KQTN\t: {REQ_DATE}", f"Ngày nhận KQTN\t: {KQTN_DATE}"):
        b.para(line, keep_next=True, tab_cm=5.5)
    b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre_han_text(KQTN_DATE, SEND_DATE, REQ_DATE, PLACEHOLDER)}")
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm")):
            p.paragraph_format.keep_with_next = True
    doc.save(out)
    print(f"[build_nuoc_mau] BCKQ -> {out}")
    return out


# ---------------------------------------------------------------------------
# Phiếu Excel bùn thải: quy đổi hàm lượng tuyệt đối về mẫu rắn khô (công thức động — AGENTS mục 11)
# ---------------------------------------------------------------------------
def build_sludge_excel(d, sp):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from Core.Tools.excel_borders import thin_outside_borders
    wb = Workbook()
    ws = wb.active
    ws.title = "Tinh toan"
    f = Font(name="Times New Roman", size=12)
    fb = Font(name="Times New Roman", size=12, bold=True)
    ws["A1"], ws["F1"] = f"Số: {d.code}", f"Ngày {KQTN_DATE}"
    ws["A3"] = "TÍNH TOÁN QUY ĐỔI HÀM LƯỢNG TUYỆT ĐỐI THEO MẪU RẮN KHÔ"
    ws.merge_cells("A3:F3")
    hdr = ["Kí hiệu mẫu", "Chỉ tiêu", "Kết quả, mg/kg", "Độ ẩm, %", "Hàm lượng chất rắn khô, %",
           "Kết quả quy đổi về khối lượng khô, mg/kg"]
    for j, h in enumerate(hdr, start=1):
        ws.cell(5, j, h)
    items = [("Tổng xyanua", "1,2")] + [(f"PCB-{n}", "< 0,003") for n in (28, 52, 101, 138, 153, 180)]
    r0 = 6
    for k, (name, val) in enumerate(items):
        r = r0 + k
        ws.cell(r, 2, name)
        if val.startswith("<"):
            ws.cell(r, 3, val)
        else:
            ws.cell(r, 3, float(val.replace(",", ".")))
            ws.cell(r, 3).number_format = "0.0"
        ws.cell(r, 4, f"=$D${r0}" if k else float(sp["moisture"].replace(",", ".")))
        ws.cell(r, 5, f"=100-D{r}")
        ws.cell(r, 6, f'=IF(ISNUMBER(C{r}),ROUND(C{r}*100/E{r},2),C{r})')
        ws.cell(r, 4).number_format = ws.cell(r, 5).number_format = "0.0"
    rN = r0 + len(items) - 1
    ws.cell(r0, 1, sp["sample"])
    ws.merge_cells(start_row=r0, start_column=1, end_row=rN, end_column=1)
    thin = Side(style="thin")
    for row in ws.iter_rows(min_row=5, max_row=rN, max_col=6):
        for c in row:
            c.font = fb if c.row == 5 else f
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ("A1", "F1"):
        ws[c].font = f
    ws["A3"].font = Font(name="Times New Roman", size=13, bold=True)
    ws["A3"].alignment = Alignment(horizontal="center")
    for col, w in zip("ABCDEF", (14, 18, 14, 11, 16, 22)):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[5].height = 48
    thin_outside_borders(wb)
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True                 # 1 trang
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    out = os.path.join(d.input_dir, "ThucHien", f"{d.code} - TinhToanChatThai.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    print(f"[build_nuoc_mau] Excel -> {out}")
    return out


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + giấy tiêu đề UpChungThu")
    ap.add_argument("--only", default="", help="Chỉ dựng các mã, vd 999911,999913")
    args = ap.parse_args()
    only = {s.strip() for s in args.only.split(",") if s.strip()}
    jobs = list(SPECS.items()) + [("999914", BT)]
    report = []
    for code6, sp in jobs:
        if only and code6 not in only:
            continue
        print(f"\n===== Q326D{code6} — {sp['loai']} (HỒ SƠ ẢO) =====")
        d = make_dossier(code6, sp)
        d.ensure_output_directories()
        docs, xls = [build_bckq(d, code6, sp)], []
        if sp["kind"] == "bt":
            docs.append(build_tbkq_sludge(d, code6, sp))
            xls.append(build_sludge_excel(d, sp))
        else:
            docs.append(build_tbkq_water(d, code6, sp, is_en=False))
            if sp.get("scope_en"):
                docs.append(build_tbkq_water(d, code6, sp, is_en=True))
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            ex = PDFExporter(d)
            ex.export_files(docs)
            if xls:
                ex.export_excel_files(xls)
        for p in docs:
            res = check_single_docx(p, code=d.code)
            report.append((res.get("status"), os.path.relpath(p, WORKSPACE), res.get("issues", [])))
    print("\n===== QA =====")
    for st, p, issues in report:
        print(f"  {st}\t{p}")
        for i in issues:
            print(f"      - {i}")


if __name__ == "__main__":
    main()
