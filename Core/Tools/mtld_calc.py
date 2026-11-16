# -*- coding: utf-8 -*-
"""
Core/Tools/mtld_calc.py
Bộ công cụ chuẩn hóa bảng tra cứu hóa chất (QCVN 03:2019/BYT, QCVN 01:2025/BYT,
QCVN 31:2026/BYT, QĐ 3733/2002/QĐ-BYT) và sinh các biểu mẫu tính toán Excel chuẩn
cho Quan trắc môi trường lao động (QTMTLĐ).

Tuân thủ AGENTS.md mục 11:
- Giữ nguyên 100% công thức Excel động (AVERAGE, IF, ROUND, COUNT, SUM, VLOOKUP, LET).
- Bổ sung đầy đủ các chất bị thiếu trong Sheet1 của bảng tính Hơi khí độc:
  + QCVN 01:2025/BYT (70 chất, Thông tư 09/2025/TT-BYT): bổ sung Butyl acrylat (55 mg/m³),
    Aceton cyanohydrin / Axeton xyanohydrin (4 mg/m³), Caprolactam dạng hơi (20 mg/m³),
    Đá talc bụi toàn phần (2 mg/m³), chuẩn hóa Cađimi về QCVN 03:2019/BYT.
  + QCVN 03:2019/BYT (50 chất, Thông tư 10/2019/TT-BYT): chuẩn hóa toàn bộ giá trị số
    (sửa lỗi '1.000' -> 1000, '9.000' -> 9000) và bổ sung đầy đủ STEL.
  + QCVN 31:2026/BYT (80 chất, Thông tư 38/2026/TT-BYT, hiệu lực 09/04/2027): bổ sung cột
    tham chiếu TWA (QCVN 31:2026).
  + Bổ sung hơn 180 tên đồng nghĩa / tên theo QCVN / tên tiếng Anh / công thức hóa học /
    tên trên phiếu KQTN (NOx, SO2, CO, CO2, NH3, HCl, H2SO4, HNO3, HF, Hexan, n-Hexan,
    Heptan, n-Heptan, Ethanol, Methanol, Acetone, Toluen, Xylen, Dầu khoáng, Sắt, Chì,
    Đồng, Kẽm, Mangan, Thiếc, Isopropanol, Etyl axetat, Butyl axetat...) để VLOOKUP không
    bao giờ bị #N/A.
"""

from __future__ import annotations

import copy
import os
import shutil
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT_DIR = Path(__file__).resolve().parents[2]
MTLD_SAMPLE_DIR = ROOT_DIR / "MTLD"
TEMPLATE_MTLD_DIR = ROOT_DIR / "Core" / "Templates" / "MTLD"


def _find_sample(folder_prefix: str, filename_substr: str) -> Path:
    """Tìm file mẫu trong thư mục MTLD bất kể chuẩn NFC hay NFD của tên file/thư mục."""
    norm_sub = unicodedata.normalize("NFC", filename_substr).lower()
    for d in MTLD_SAMPLE_DIR.iterdir():
        if d.is_dir() and d.name.startswith(folder_prefix):
            for p in d.rglob("*"):
                if p.is_file() and not p.name.startswith("~$"):
                    if norm_sub in unicodedata.normalize("NFC", p.name).lower():
                        return p
    raise FileNotFoundError(f"Không tìm thấy mẫu: {folder_prefix} / {filename_substr}")


# ==============================================================================
# 1. CƠ SỞ DỮ LIỆU ĐẦY ĐỦ QCVN 03:2019/BYT, QCVN 01:2025/BYT, QCVN 31:2026/BYT, QĐ 3733
# ==============================================================================

# STEL theo QĐ 3733/2002/QĐ-BYT cho các chất thông dụng & các chất chỉ có STEL (TWA = '-')
QD3733_STEL_MAP: Dict[str, float] = {
    "Acrolein": 0.8,
    "Acrylonitril": 2.5,
    "Alyl axetat": 2.0,
    "Amoniac": 25.0,
    "Amyl axetat": 525.0,
    "Anhydrit phtalic": 5.0,
    "Anilin": 8.0,
    "Asen và các hợp chất": 0.06,
    "Asin": 0.15,
    "Axeton": 1000.0,
    "Axeton xyanohydrin": 0.9,
    "Axetonitril": 100.0,
    "Axit axetic": 35.0,
    "Axit clohiđric": 7.5,
    "Axít fomic": 19.0,
    "Axít nitric": 10.0,
    "Axít phosphoric": 3.0,
    "Axit piric": 0.3,
    "Axít sunfuric": 2.0,
    "Bạc": 0.03,
    "Bạc (dạng hợp chất)": 0.03,
    "Bari oxit": 1.5,
    "Benzen": 15.0,
    "Benzonitril": 1.0,
    "Benzotriclorua": 0.2,
    "Benzyl clorua": 2.0,
    "Beryli và các hợp chất": 0.003,
    "Brom": 2.0,
    "1,3-Butadien": 40.0,
    "Butylaxetat": 700.0,
    "Butanol": 300.0,
    "Cađimi và các hợp chất": 0.05,
    "Cacbon đioxit": 1800.0,
    "Cacbon đisunfua": 25.0,
    "Cacbon monoxit": 40.0,
    "Cacbon tetraclorua": 20.0,
    "Canxi oxit": 5.0,
    "Caprolactam (bụi)": 3.0,
    "Caprolactam (khói)": 40.0,
    "Captan": 15.0,
    "Carbaryl": 5.0,
    "Chì tetraetyl": 0.01,
    "Chì và các hợp chất": 0.1,
    "Clo": 3.0,
    "Clo đioxit": 0.8,
    "Clobenzen": 200.0,
    "Clorofom": 20.0,
    "Coban và hợp chất": 0.1,
    "Crom VI oxit": 0.1,
    "Crom (III) (dạng hợp chất)": 1.0,
    "Crom (VI) (dạng hòa tan trong nước)": 0.02,
    "Cumen": 160.0,
    "Dầu khoáng (sương mù)": 10.0,
    "Dầu thông": 600.0,
    "Đibutyl phtalat": 5.0,
    "Đicloetan": 12.0,
    "1,1- Đicloetylen": 20.0,
    "Điclometan": 100.0,
    "Điclovos": 3.0,
    "Đimetyl amin": 5.0,
    "Đimetyfomamit": 20.0,
    "1,1- Đimetylhyđrazin": 0.5,
    "Đimetyl sunfat": 0.1,
    "Đinitrobanzen": 2.0,
    "Đinitrotoluen (DNT)": 2.0,
    "1,4 Đioxan": 20.0,
    "Đồng (bụi)": 1.0,
    "Đồng (hơi, khói)": 0.5,
    "Enđosunfan": 0.3,
    "Etanolamin": 15.0,
    "Ete cloetyl": 2.0,
    "Ete etyl": 1500.0,
    "Ete resorcinol monometyl": 0.1,
    "Etylen glycol (hơi)": 125.0,
    "Etylen oxit": 2.0,
    "Flo": 0.4,
    "Florua (các hợp chất F)": 2.5,
    "Fomanđehyt": 1.0,
    "Fufural": 20.0,
    "Fufuryl alcol": 40.0,
    "Heptan (tất cả các đồng phân)": 1250.0,
    "n-hexan": 180.0,
    "Hyđrazin (và...sunfate)": 0.1,
    "Hyđro florua": 0.5,
    "Hyđro sunfua": 15.0,
    "Hyđro xyanua": 0.6,
    "Hyđroxyt kiềm": 1.0,
    "Hyđroquinon": 1.0,
    "Khói hàn": 10.0,
    "Kẽm clorua": 2.0,
    "Kẽm oxit (bụi, khói)": 10.0,
    "Long não": 4.0,
    "Magie oxit": 10.0,
    "Malathion": 10.0,
    "Mangan và các hợp chất": 0.6,
    "Metyl acrylat": 40.0,
    "Metyl axêtat": 250.0,
    "Metyl etyl xeton": 300.0,
    "Metyl meta crylat": 150.0,
    "Naphtalen": 75.0,
    "Niken và các dạng hợp chất (hòa tan)": 0.25,
    "Nitơ đioxit": 10.0,
    "Nitơ mono oxit": 20.0,
    "Nitro benzen": 6.0,
    "Nitro toluen": 22.0,
    "Octan (tất cả các đồng phân)": 1400.0,
    "Ozon": 0.2,
    "Paraquat": 0.3,
    "Parathion": 0.1,
    "Phenol": 8.0,
    "Phosgen": 0.4,
    "Phosphin": 0.2,
    "n-Propyl axetat": 400.0,
    "Pyriđin": 10.0,
    "Rượu etylic": 3000.0,
    "Rượu metylic": 100.0,
    "Rượu propylic": 600.0,
    "Sáp parafin (khói)": 3.0,
    "Sắt (III) oxit (bụi, khói)": 10.0,
    "Selen và các hợp chất": 0.3,
    "Selen đioxit": 0.3,
    "Styren": 200.0,
    "Sunfua đioxit": 10.0,
    "Tetracloetylen": 150.0,
    "Thiếc (hữu cơ)": 0.2,
    "Thiếc (vô cơ)": 2.0,
    "Thiếc oxit": 4.0,
    "Thủy ngân và hợp chất thủy ngân vô cơ": 0.04,
    "Toluen": 300.0,
    "Toluen điisoxyanat": 0.1,
    "Tricloetylen": 40.0,
    "2,4,6-Trinitrotoluen (TNT)": 0.5,
    "Vinyl clorua": 5.0,
    "Xăng": 500.0,
    "Xyclohexan": 1000.0,
    "Xylen": 300.0,
}

# 50 chất theo QCVN 03:2019/BYT (Thông tư 10/2019/TT-BYT): (TWA, STEL)
QCVN_03_2019_MAP: Dict[str, Tuple[Optional[float], Optional[float]]] = {
    "Amoniac": (17.0, 25.0),
    "Anilin": (4.0, 8.0),
    "Asen và các hợp chất": (0.01, 0.04),
    "Asin": (0.05, 0.15),
    "Axeton": (200.0, 1000.0),
    "Axit axetic": (25.0, 35.0),
    "Axit clohiđric": (5.0, 7.5),
    "Axít sunfuric": (1.0, 2.0),
    "Benzen": (5.0, 15.0),
    "Butanol": (150.0, 300.0),
    "Cađimi và các hợp chất": (0.01, None),
    "Cacbon đioxit": (9000.0, 18000.0),
    "Cacbon đisunfua": (15.0, 25.0),
    "Cacbon monoxit": (20.0, 40.0),
    "Cacbon tetraclorua": (10.0, 20.0),
    "Clo": (1.5, 3.0),
    "Clorofom": (10.0, 20.0),
    "Coban và hợp chất": (0.05, 0.1),
    "Crom VI oxit": (0.05, 0.1),
    "Crom (III) (dạng hợp chất)": (0.5, 1.0),
    "Crom (VI) (dạng hòa tan trong nước)": (0.01, 0.02),
    "Điclometan": (50.0, 100.0),
    "Đồng (bụi)": (0.5, 1.0),
    "Đồng (hơi, khói)": (0.1, 0.5),
    "Flo": (0.2, 0.4),
    "Florua (các hợp chất F)": (1.0, 2.5),
    "Fomanđehyt": (0.5, 1.0),
    "n-hexan": (90.0, 180.0),
    "Hyđro sunfua": (10.0, 15.0),
    "Hyđro xyanua": (0.3, 0.6),
    "Kẽm oxit (bụi, khói)": (5.0, 10.0),
    "Mangan và các hợp chất": (0.3, 0.6),
    "Metyl axêtat": (100.0, 250.0),
    "Nhôm và hợp chất": (2.0, None),
    "Nicotin": (0.5, 1.5),
    "Nitơ đioxit": (5.0, 10.0),
    "Nitơ mono oxit": (10.0, 20.0),
    "Nitro benzen": (3.0, 6.0),
    "Nitro toluen": (11.0, 22.0),
    "Phenol": (4.0, 8.0),
    "Rượu etylic": (1000.0, 3000.0),
    "Rượu metylic": (50.0, 100.0),
    "Selen và các hợp chất": (0.1, 0.3),
    "Selen đioxit": (0.1, 0.3),
    "Sunfua đioxit": (5.0, 10.0),
    "Toluen": (100.0, 300.0),
    "2,4,6-Trinitrotoluen (TNT)": (0.1, 0.5),
    "Vinyl clorua": (1.0, 5.0),
    "Xăng": (300.0, 500.0),
    "Xylen": (100.0, 300.0),
}

# 70 chất theo QCVN 01:2025/BYT (Thông tư 09/2025/TT-BYT, hiệu lực 01/10/2025): (TWA, STEL)
QCVN_01_2025_MAP: Dict[str, Tuple[Optional[float], Optional[float]]] = {
    "Acrolein": (0.25, 0.8),
    "Acrylamit": (0.03, 0.2),
    "Acrylonitril": (4.3, 21.5),
    "Amyl axetat": (266.0, 525.0),
    "Anhydrit phtalic": (2.0, 5.0),
    "Antimon": (0.5, None),
    "ANTU": (0.3, None),
    "Atphan (bitum, nhựa đường)": (5.0, None),
    "Axeton xyanohydrin": (4.0, 5.0),
    "Axetonitril": (50.0, 100.0),
    "Axít fomic": (9.0, 19.0),
    "Axit metacrylic": (70.0, None),
    "Axít nitric": (5.0, 10.0),
    "Axít phosphoric": (1.0, 3.0),
    "Axit piric": (0.1, 0.3),
    "Axit tricloaxetic": (7.0, None),
    "Azinpho metyl": (0.2, None),
    "Bạc": (0.01, 0.03),
    "Bạc (dạng hợp chất)": (0.01, 0.03),
    "Benomyl": (10.0, 15.0),
    "Benzidin": (0.008, None),
    "Benzoyl peroxit": (5.0, None),
    "Benzyl clorua": (5.0, None),
    "Beryli và các hợp chất": (0.002, 0.01),
    "Brom": (0.7, 2.0),
    "1,3-Butadien": (20.0, 40.0),
    "Butylaxetat": (710.0, 950.0),
    "Butyl acrylat": (55.0, None),
    "Cacbonfuran": (0.1, None),
    "Canxi cacbonat": (10.0, None),
    "Canxi hyđroxit": (5.0, None),
    "Canxi oxit": (2.0, 5.0),
    "Canxi silicat": (10.0, None),
    "Canxi sunfat đihyđrat": (6.0, None),
    "Canxi xyanamit": (0.5, None),
    "Caprolactam (bụi)": (1.0, 3.0),
    "Caprolactam (khói)": (20.0, 40.0),
    "Captan": (5.0, 15.0),
    "Carbaryl": (5.0, 10.0),
    "Catechol": (20.0, None),
    "Chì tetraetyl": (0.075, 0.3),
    "Chì và các hợp chất": (0.05, 0.15),
    "Clo axetalđehyt": (3.0, None),
    "Clo đioxit": (0.3, 0.8),
    "Cloaxetophenon": (0.3, None),
    "Clobenzen": (100.0, 200.0),
    "Clopren": (36.0, None),
    "Cresol": (22.0, None),
    "Crotonalđehyt": (5.0, None),
    "Cumen": (245.0, None),
    "Dầu khoáng (sương mù)": (5.0, 10.0),
    "Dầu mỏ (naphta)": (350.0, None),
    "Dầu thông": (560.0, None),
    "Dầu thực vật (dạng sương)": (10.0, None),
    "Dung môi stoddard": (525.0, None),
    "Đá talc, hoạt thạch (bụi hô hấp)": (1.0, None),
    "Đá talc, hoạt thạch (chứa 1% quartz)": (2.0, None),
    "Đemeton": (0.1, 0.3),
    "Điazinon": (0.1, 0.3),
    "Điboran": (0.1, None),
    "Đibutyl phtalat": (5.0, None),
    "Đicloetan": (400.0, 810.0),
    "1,1- Đicloetylen": (8.0, 20.0),
    "Điclovos": (1.0, 3.0),
    "Điclotophos": (0.25, None),
    "Đimetyl amin": (18.0, 27.6),
    "Đimetyfomamit": (30.0, 60.0),
    "1,1- Đimetylhyđrazin": (0.2, 0.5),
    "Đimetyl sunfat": (0.5, None),
    "Đinitrobanzen": (1.0, 3.0),
    "Đinitrotoluen (DNT)": (1.5, 5.0),
}

# 80 chất theo QCVN 31:2026/BYT (Thông tư 38/2026/TT-BYT, hiệu lực 09/04/2027): TWA (mg/m³)
QCVN_31_2026_TWA_MAP: Dict[str, float] = {
    "Xenluloza (bụi tổng số)": 10.0,
    "Xenluloza (bụi hô hấp)": 5.0,
    "Xyanua": 1.0,
    "Xyanogen": 21.0,
    "Xyclohexan": 1050.0,
    "Xyclohexanol": 200.0,
    "1,4 Đioxan": 90.0,
    "Enđosunfan": 0.1,
    "ΕΡΝ": 0.5,
    "Etanolamin": 8.0,
    "Etyl-amin": 18.0,
    "Ete etyl": 1200.0,
    "Etyl mercaptan": 1.0,
    "Etylen đibromua": 1.0,
    "Etylen glycol (hơi)": 60.0,
    "Etylen glycol (sơn khí, hạt, mù sương)": 10.0,
    "Etylen glycol initrat (dinitrat)": 0.3,
    "Etylen oxit": 1.8,
    "Fufural": 8.0,
    "Fufuryl alcol": 40.0,
    "Rượu fufuryl": 40.0,
    "Hắc ín than đá (hơi)": 0.2,
    "Halothan": 400.0,
    "Heptaclo": 0.5,
    "Heptan (tất cả các đồng phân)": 1200.0,
    "Hyđrazin (và...sunfate)": 0.05,
    "Hyđro florua": 2.5,
    "Hyđroxyt kiềm": 2.0,
    "Hyđroquinon": 2.0,
    "Iot": 1.0,
    "Isopropyl glyxiđyl ete": 240.0,
    "Kali cyanua": 5.0,
    "Kẽm clorua": 1.0,
    "Kẽm sterat (bụi tổng số)": 10.0,
    "Kẽm sterat (bụi hô hấp)": 5.0,
    "Khói hàn": 5.0,
    "Long não": 2.0,
    "Sunfua tetraflorua": 0.4,
    "Magie oxit": 10.0,
    "Malathion": 10.0,
    "Metyl acrylat": 35.0,
    "Metyl mercaptan": 1.0,
    "Metyl meta crylat": 205.0,
    "Muội than": 3.5,
    "Naphtalen": 50.0,
    "Natri bisunfit": 5.0,
    "Natri cyanua": 5.0,
    "Natri nitrua": 0.3,
    "Niken và các dạng hợp chất (hòa tan)": 0.1,
    "Nikel cacbonyl": 0.12,
    "Nitro etan": 310.0,
    "Nitro metan": 50.0,
    "1-Nitropropan": 90.0,
    "2-Nitropropan": 36.0,
    "Nitroglyxerin": 0.5,
    "Ozon": 0.2,
    "Paraquat": 0.1,
    "Parathion": 0.1,
    "Phenyl hyđrazin": 0.4,
    "Phorat": 0.05,
    "Phosgen": 0.4,
    "Phosphin": 0.4,
    "Phospho (trắng, vàng)": 0.1,
    "Phospho triclorua": 1.5,
    "Pyriđin": 15.0,
    "Resorinol": 45.0,
    "Rượu alilic": 5.0,
    "Rượu propylic": 350.0,
    "Rượu propagyl": 2.0,
    "Sáp parafin (khói)": 2.0,
    "Sắt (III) oxit (bụi, khói)": 5.0,
    "Styren": 210.0,
    "Tetracloetylen": 170.0,
    "Tetranitrometan": 8.0,
    "Thiếc (hữu cơ)": 0.1,
    "Thiếc (vô cơ)": 2.0,
    "Thiếc oxit": 2.0,
    "Thiram": 5.0,
    "Thủy ngân và hợp chất thủy ngân vô cơ": 0.02,
    "Titan đioxit (bụi hô hấp)": 5.0,
    "Titan đioxit (bụi tổng số)": 10.0,
    "Toluen điisoxyanat": 0.04,
}


# Danh mục các chất bổ sung mới hoàn toàn (chưa có trong 356 dòng QĐ 3733)
# và các tên chuẩn theo QCVN / tên trên phiếu KQTN / công thức hóa học
# Cấu trúc: (Tên tra cứu, Tên gốc trong 356 dòng hoặc None nếu chất mới,
#            override_tuple=(twa_3733, twa_03_2019, twa_01_2025, stel_3733, stel_03_2019, stel_01_2025, twa_31_2026))
NEW_AND_ALIAS_CHEMICALS: List[Tuple[str, Optional[str], Optional[Tuple[Any, ...]]]] = [
    # --- Chất mới trong QCVN 01:2025/BYT & QCVN 03:2019/BYT ---
    ("Butyl acrylat", None, ("-", "-", 55.0, "-", "-", "-", "-")),
    ("Butyl acrylate", None, ("-", "-", 55.0, "-", "-", "-", "-")),
    ("n-Butyl acrylat", None, ("-", "-", 55.0, "-", "-", "-", "-")),
    ("Aceton cyanohydrin", "Axeton xyanohydrin", None),
    ("Acetone cyanohydrin", "Axeton xyanohydrin", None),
    ("Caprolactam (dạng hơi)", "Caprolactam (khói)", None),
    ("Caprolactam (dạng bụi)", "Caprolactam (bụi)", None),
    ("Đá talc (bụi toàn phần)", "Đá talc, hoạt thạch (chứa 1% quartz)", None),
    ("Đá talc (bụi hô hấp)", "Đá talc, hoạt thạch (bụi hô hấp)", None),
    ("Bạc và các hợp chất hòa tan", "Bạc", None),
    ("Cađimi oxit (khói)", None, (0.01, 0.005, "-", 0.05, "-", "-", "-")),
    ("Cadmi oxide (khói)", None, (0.01, 0.005, "-", 0.05, "-", "-", "-")),
    # --- Tên chuẩn trong QCVN 03:2019/BYT ---
    ("Aceton", "Axeton", None),
    ("Acetone", "Axeton", None),
    ("Hàm lượng Aceton", "Axeton", None),
    ("Acid acetic", "Axit axetic", None),
    ("Axit acetic", "Axit axetic", None),
    ("CH3COOH", "Axit axetic", None),
    ("Acid hydrochloric", "Axit clohiđric", None),
    ("Axit clohydric", "Axit clohiđric", None),
    ("HCl", "Axit clohiđric", None),
    ("Hơi HCl", "Axit clohiđric", None),
    ("Acid sulfuric", "Axít sunfuric", None),
    ("Axit sulfuric", "Axít sunfuric", None),
    ("Axit sunfuric", "Axít sunfuric", None),
    ("H2SO4", "Axít sunfuric", None),
    ("H₂SO₄", "Axít sunfuric", None),
    ("Hơi H2SO4", "Axít sunfuric", None),
    ("Amonia", "Amoniac", None),
    ("Ammonia", "Amoniac", None),
    ("NH3", "Amoniac", None),
    ("NH₃", "Amoniac", None),
    ("Hàm lượng Amoniac (NH3)", "Amoniac", None),
    ("Hàm lượng Amoniac", "Amoniac", None),
    ("Arsenic và hợp chất", "Asen và các hợp chất", None),
    ("Asen", "Asen và các hợp chất", None),
    ("As", "Asen và các hợp chất", None),
    ("Arsin", "Asin", None),
    ("AsH3", "Asin", None),
    ("Benzene", "Benzen", None),
    ("Hàm lượng Benzen", "Benzen", None),
    ("n-Butanol", "Butanol", None),
    ("Isobutanol", "Butanol", None),
    ("Rượu butylic", "Butanol", None),
    ("Cadmi và hợp chất", "Cađimi và các hợp chất", None),
    ("Cadimi", "Cađimi và các hợp chất", None),
    ("Cd", "Cađimi và các hợp chất", None),
    ("Hàm lượng Cadimi (Cd)", "Cađimi và các hợp chất", None),
    ("Carbon dioxide", "Cacbon đioxit", None),
    ("CO2", "Cacbon đioxit", None),
    ("CO₂", "Cacbon đioxit", None),
    ("Carbon disulfide", "Cacbon đisunfua", None),
    ("CS2", "Cacbon đisunfua", None),
    ("Carbon monoxide", "Cacbon monoxit", None),
    ("CO", "Cacbon monoxit", None),
    ("Carbon tetrachloride", "Cacbon tetraclorua", None),
    ("CCl4", "Cacbon tetraclorua", None),
    ("Chlor", "Clo", None),
    ("Cl2", "Clo", None),
    ("Cl₂", "Clo", None),
    ("Chloroform", "Clorofom", None),
    ("CHCl3", "Clorofom", None),
    ("Chromi (III) (dạng hợp chất)", "Crom (III) (dạng hợp chất)", None),
    ("Chromi (VI) (dạng hòa tan trong nước)", "Crom (VI) (dạng hòa tan trong nước)", None),
    ("Chromi (VI) oxide", "Crom VI oxit", None),
    ("Crom", "Crom (III) (dạng hợp chất)", None),
    ("Cr", "Crom (III) (dạng hợp chất)", None),
    ("Cr(VI)", "Crom (VI) (dạng hòa tan trong nước)", None),
    ("Cobalt và hợp chất", "Coban và hợp chất", None),
    ("Coban", "Coban và hợp chất", None),
    ("Co", "Coban và hợp chất", None),
    ("Dichloromethan", "Điclometan", None),
    ("Dichloromethane", "Điclometan", None),
    ("Dichloromethane*", "Điclometan", None),
    ("Methylene chloride", "Điclometan", None),
    ("CH2Cl2", "Điclometan", None),
    ("Đồng và hợp chất (dạng bụi)", "Đồng (bụi)", None),
    ("Đồng và hợp chất (dạng hơi, khói)", "Đồng (hơi, khói)", None),
    ("Hơi đồng trong không khí (Cu)", "Đồng (bụi)", None),
    ("Hơi đồng", "Đồng (hơi, khói)", None),
    ("Đồng", "Đồng (bụi)", None),
    ("Cu", "Đồng (bụi)", None),
    ("Ethanol", "Rượu etylic", None),
    ("Cồn", "Rượu etylic", None),
    ("Ethyl alcohol", "Rượu etylic", None),
    ("Hơi ethanol", "Rượu etylic", None),
    ("Fluor", "Flo", None),
    ("F2", "Flo", None),
    ("Fluoride", "Florua (các hợp chất F)", None),
    ("F-", "Florua (các hợp chất F)", None),
    ("Formaldehyde", "Fomanđehyt", None),
    ("Formaldehyt", "Fomanđehyt", None),
    ("Formaldehye", "Fomanđehyt", None),
    ("Hàm lượng Formaldehyde", "Fomanđehyt", None),
    ("Hàm lượng Formaldehye", "Fomanđehyt", None),
    ("HCHO", "Fomanđehyt", None),
    ("n-Hexan", "n-hexan", None),
    ("n-Hexan*", "n-hexan", None),
    ("n-Hexane", "n-hexan", None),
    ("Hexan", "n-hexan", None),
    ("Hexane", "n-hexan", None),
    ("Hàm lượng Hexan", "n-hexan", None),
    ("Hydro cyanide", "Hyđro xyanua", None),
    ("HCN", "Hyđro xyanua", None),
    ("Hydro sulfide", "Hyđro sunfua", None),
    ("H2S", "Hyđro sunfua", None),
    ("H₂S", "Hyđro sunfua", None),
    ("Kẽm oxide (khói)", "Kẽm oxit (bụi, khói)", None),
    ("Kẽm oxit", "Kẽm oxit (bụi, khói)", None),
    ("Hơi kẽm", "Kẽm oxit (bụi, khói)", None),
    ("Kẽm", "Kẽm oxit (bụi, khói)", None),
    ("Zn", "Kẽm oxit (bụi, khói)", None),
    ("ZnO", "Kẽm oxit (bụi, khói)", None),
    ("Mangan", "Mangan và các hợp chất", None),
    ("Hơi Mangan", "Mangan và các hợp chất", None),
    ("Mn", "Mangan và các hợp chất", None),
    ("Methanol", "Rượu metylic", None),
    ("Methyl alcohol", "Rượu metylic", None),
    ("CH3OH", "Rượu metylic", None),
    ("Methyl acetat", "Metyl axêtat", None),
    ("Methyl acetate", "Metyl axêtat", None),
    ("Metyl axetat", "Metyl axêtat", None),
    ("Nhôm và các hợp chất", "Nhôm và hợp chất", None),
    ("Nhôm", "Nhôm và hợp chất", None),
    ("Al", "Nhôm và hợp chất", None),
    ("Nitơ dioxide", "Nitơ đioxit", None),
    ("NO2", "Nitơ đioxit", None),
    ("NO₂", "Nitơ đioxit", None),
    ("NOx", "Nitơ đioxit", None),
    ("NOₓ", "Nitơ đioxit", None),
    ("Nitơ monoxide", "Nitơ mono oxit", None),
    ("NO", "Nitơ mono oxit", None),
    ("Nitrobenzene", "Nitro benzen", None),
    ("Nitrotoluen", "Nitro toluen", None),
    ("Selen", "Selen và các hợp chất", None),
    ("Se", "Selen và các hợp chất", None),
    ("Selen dioxide", "Selen đioxit", None),
    ("SeO2", "Selen đioxit", None),
    ("Sulfur dioxide", "Sunfua đioxit", None),
    ("SO2", "Sunfua đioxit", None),
    ("SO₂", "Sunfua đioxit", None),
    ("Toluene", "Toluen", None),
    ("Hàm lượng Toluen", "Toluen", None),
    ("TNT", "2,4,6-Trinitrotoluen (TNT)", None),
    ("Vinyl chloride", "Vinyl clorua", None),
    ("VCM", "Vinyl clorua", None),
    ("Xăng (Gasoline)", "Xăng", None),
    ("Gasoline", "Xăng", None),
    ("Xylene", "Xylen", None),
    ("Hàm lượng Xylen", "Xylen", None),
    # --- Tên chuẩn trong QCVN 01:2025/BYT ---
    ("Acrylamid", "Acrylamit", None),
    ("Acrylamide", "Acrylamit", None),
    ("Acrylonitrile", "Acrylonitril", None),
    ("Amyl acetat", "Amyl axetat", None),
    ("Amyl acetate", "Amyl axetat", None),
    ("Anhydrid phthalic", "Anhydrit phtalic", None),
    ("Sb", "Antimon", None),
    ("Asphalt (dạng khói)", "Atphan (bitum, nhựa đường)", None),
    ("Acetonitril", "Axetonitril", None),
    ("Acetonitrile", "Axetonitril", None),
    ("Acid formic", "Axít fomic", None),
    ("Axit formic", "Axít fomic", None),
    ("HCOOH", "Axít fomic", None),
    ("Acid methacrylic", "Axit metacrylic", None),
    ("Hydro nitrat (Acid nitric)", "Axít nitric", None),
    ("Acid nitric", "Axít nitric", None),
    ("Axit nitric", "Axít nitric", None),
    ("HNO3", "Axít nitric", None),
    ("HNO₃", "Axít nitric", None),
    ("Trihydro phosphat (Acid phosphoric)", "Axít phosphoric", None),
    ("Acid phosphoric", "Axít phosphoric", None),
    ("Axit phosphoric", "Axít phosphoric", None),
    ("H3PO4", "Axít phosphoric", None),
    ("H₃PO₄", "Axít phosphoric", None),
    ("Acid picric", "Axit piric", None),
    ("Acid trichloroacetic", "Axit tricloaxetic", None),
    ("Azinphos methyl", "Azinpho metyl", None),
    ("Ag", "Bạc", None),
    ("Benomyl (bụi toàn phần)", "Benomyl", None),
    ("Benzoyl peroxide", "Benzoyl peroxit", None),
    ("Benzyl chloride", "Benzyl clorua", None),
    ("Be", "Beryli và các hợp chất", None),
    ("Br2", "Brom", None),
    ("n-Butyl acetat", "Butylaxetat", None),
    ("Butyl acetat", "Butylaxetat", None),
    ("Butyl acetate", "Butylaxetat", None),
    ("Butyl axetat", "Butylaxetat", None),
    ("n-Butyl acetate", "Butylaxetat", None),
    ("Carbofuran", "Cacbonfuran", None),
    ("Calci carbonat (bụi toàn phần)", "Canxi cacbonat", None),
    ("CaCO3", "Canxi cacbonat", None),
    ("Calci hydroxide (bụi toàn phần)", "Canxi hyđroxit", None),
    ("Ca(OH)2", "Canxi hyđroxit", None),
    ("Calci oxide", "Canxi oxit", None),
    ("CaO", "Canxi oxit", None),
    ("Calci silicat (bụi toàn phần)", "Canxi silicat", None),
    ("Calci sulfat dihydrat (bụi toàn phần)", "Canxi sunfat đihyđrat", None),
    ("Calci cyanamid", "Canxi xyanamit", None),
    ("Chì tetraethyl", "Chì tetraetyl", None),
    ("Chì và các hợp chất vô cơ", "Chì và các hợp chất", None),
    ("Chì", "Chì và các hợp chất", None),
    ("Hơi chì", "Chì và các hợp chất", None),
    ("Hàm lượng chì (Pb)", "Chì và các hợp chất", None),
    ("Pb", "Chì và các hợp chất", None),
    ("Chloroacetaldehyd", "Clo axetalđehyt", None),
    ("Chlor dioxide", "Clo đioxit", None),
    ("ClO2", "Clo đioxit", None),
    ("Chloroacetophenol", "Cloaxetophenon", None),
    ("Chlorobenzen", "Clobenzen", None),
    ("Chlorobezen", "Clobenzen", None),
    ("Chlorobenzene", "Clobenzen", None),
    ("Chloropren", "Clopren", None),
    ("Crotonaldehyd", "Crotonalđehyt", None),
    ("Cumene", "Cumen", None),
    ("Dầu khoáng (dạng sương)", "Dầu khoáng (sương mù)", None),
    ("Dầu khoáng", "Dầu khoáng (sương mù)", None),
    ("Hơi dầu khoáng", "Dầu khoáng (sương mù)", None),
    ("Dầu mỏ (napthas)", "Dầu mỏ (naphta)", None),
    ("Dầu thực vật (dạng sương – bụi toàn phần)", "Dầu thực vật (dạng sương)", None),
    ("Dung môi Stoddard", "Dung môi stoddard", None),
    ("Demeton", "Đemeton", None),
    ("Diazinon", "Điazinon", None),
    ("Diboran", "Điboran", None),
    ("Dibutyl phthalat", "Đibutyl phtalat", None),
    ("DBP", "Đibutyl phtalat", None),
    ("1,1-Dichloroethan", "Đicloetan", None),
    ("1,2-Dichloroethan", "Đicloetan", None),
    ("Dichloroethane", "Đicloetan", None),
    ("1,1-Dichloroethylen", "1,1- Đicloetylen", None),
    ("Dichlorvos", "Điclovos", None),
    ("Dicrotophos", "Điclotophos", None),
    ("Dimethylamin", "Đimetyl amin", None),
    ("Dimethylformamid", "Đimetyfomamit", None),
    ("DMF", "Đimetyfomamit", None),
    ("1,1-Dimethylhydrazin", "1,1- Đimetylhyđrazin", None),
    ("Dimethyl sulfat", "Đimetyl sunfat", None),
    ("Dinitrobenzen", "Đinitrobanzen", None),
    ("Dinitrotoluen (DNT)", "Đinitrotoluen (DNT)", None),
    # --- Các chất phổ biến khác trong QĐ 3733 & QCVN 31:2026/BYT ---
    ("Heptan", "Heptan (tất cả các đồng phân)", None),
    ("n-Heptan", "Heptan (tất cả các đồng phân)", None),
    ("n-heptan", "Heptan (tất cả các đồng phân)", None),
    ("Heptane", "Heptan (tất cả các đồng phân)", None),
    ("n-Heptane", "Heptan (tất cả các đồng phân)", None),
    ("Cyclohexan", "Xyclohexan", None),
    ("Cyclohexane", "Xyclohexan", None),
    ("Cyclohexanol", "Xyclohexanol", None),
    ("Cyclohexanon", None, (50.0, "-", "-", 100.0, "-", "-", "-")),
    ("Cyclohexanone", None, (50.0, "-", "-", 100.0, "-", "-", "-")),
    ("Xyclohexanon", None, (50.0, "-", "-", 100.0, "-", "-", "-")),
    ("Etyl axetat", None, (400.0, "-", "-", 800.0, "-", "-", "-")),
    ("Etyl acetate", None, (400.0, "-", "-", 800.0, "-", "-", "-")),
    ("Ethyl axetat", None, (400.0, "-", "-", 800.0, "-", "-", "-")),
    ("Ethyl acetate", None, (400.0, "-", "-", 800.0, "-", "-", "-")),
    ("Isopropanol", "Rượu propylic", None),
    ("Isopropyl alcohol", "Rượu propylic", None),
    ("IPA", "Rượu propylic", None),
    ("Rượu isopropylic", "Rượu propylic", None),
    ("Rượu isopropyl", "Rượu propylic", None),
    ("Propanol", "Rượu propylic", None),
    ("2-Propanol", "Rượu propylic", None),
    ("MEK", "Metyl etyl xeton", None),
    ("Methyl ethyl ketone", "Metyl etyl xeton", None),
    ("2-Butanon", "Metyl etyl xeton", None),
    ("2-Butanone", "Metyl etyl xeton", None),
    ("MIBK", None, (205.0, "-", "-", 300.0, "-", "-", "-")),
    ("Methyl isobutyl ketone", None, (205.0, "-", "-", 300.0, "-", "-", "-")),
    ("THF", None, (200.0, "-", "-", 590.0, "-", "-", "-")),
    ("Tetrahydrofuran", None, (200.0, "-", "-", 590.0, "-", "-", "-")),
    ("Styrene", "Styren", None),
    ("Hàm lượng Styren", "Styren", None),
    ("Dietyl ether", "Ete etyl", None),
    ("Diethyl ether", "Ete etyl", None),
    ("Ethyl ether", "Ete etyl", None),
    ("Iso-Octan", "Octan (tất cả các đồng phân)", None),
    ("Iso-Octane", "Octan (tất cả các đồng phân)", None),
    ("Octan", "Octan (tất cả các đồng phân)", None),
    ("n-Pentan", None, (1000.0, "-", "-", 1800.0, "-", "-", "-")),
    ("n-Pentane", None, (1000.0, "-", "-", 1800.0, "-", "-", "-")),
    ("Pentan", None, (1000.0, "-", "-", 1800.0, "-", "-", "-")),
    ("HF", "Hyđro florua", None),
    ("Hydro fluoride", "Hyđro florua", None),
    ("Axit flohydric", "Hyđro florua", None),
    ("NaOH", "Hyđroxyt kiềm", None),
    ("KOH", "Hyđroxyt kiềm", None),
    ("Hydroxide kiềm", "Hyđroxyt kiềm", None),
    ("Sắt", "Sắt (III) oxit (bụi, khói)", None),
    ("Hơi sắt", "Sắt (III) oxit (bụi, khói)", None),
    ("Hơi sắt trong không khí (Fe)", "Sắt (III) oxit (bụi, khói)", None),
    ("Sắt III oxide (bụi, khói)", "Sắt (III) oxit (bụi, khói)", None),
    ("Fe", "Sắt (III) oxit (bụi, khói)", None),
    ("Fe2O3", "Sắt (III) oxit (bụi, khói)", None),
    ("Bari Oxit", "Bari oxit", None),
    ("BaO", "Bari oxit", None),
    ("Thiếc", "Thiếc oxit", None),
    ("Hơi thiếc", "Thiếc oxit", None),
    ("Thiếc dioxide", "Thiếc oxit", None),
    ("Sn", "Thiếc oxit", None),
    ("SnO2", "Thiếc oxit", None),
    ("Niken", "Niken và các dạng hợp chất (hòa tan)", None),
    ("Nickel", "Niken và các dạng hợp chất (hòa tan)", None),
    ("Nickel và các hợp chất (hòa tan)", "Niken và các dạng hợp chất (hòa tan)", None),
    ("Ni", "Niken và các dạng hợp chất (hòa tan)", None),
    ("Thủy ngân", "Thủy ngân và hợp chất thủy ngân vô cơ", None),
    ("Hg", "Thủy ngân và hợp chất thủy ngân vô cơ", None),
    ("O3", "Ozon", None),
    ("Ozone", "Ozon", None),
    ("Tetrachloroethylen", "Tetracloetylen", None),
    ("Perchloroethylene", "Tetracloetylen", None),
    ("PCE", "Tetracloetylen", None),
    ("Trichloroethylen", "Tricloetylen", None),
    ("TCE", "Tricloetylen", None),
    ("1,4-Dioxan", "1,4 Đioxan", None),
    ("Dioxan", "1,4 Đioxan", None),
    ("Ethylen oxide", "Etylen oxit", None),
    ("EO", "Etylen oxit", None),
    ("Methyl methacrylat", "Metyl meta crylat", None),
    ("MMA", "Metyl meta crylat", None),
    ("Methyl acrylat", "Metyl acrylat", None),
    ("TDI", "Toluen điisoxyanat", None),
    ("Toluen-2,4-diisocyanat", "Toluen điisoxyanat", None),
    ("Cyanide", "Xyanua", None),
    ("CN-", "Xyanua", None),
]


def _clean_num(val: Any) -> Any:
    """Chuyển chuỗi số có dấu chấm phân cách hàng nghìn ('1.000', '9.000') thành float/int chuẩn."""
    if val is None:
        return "-"
    if isinstance(val, (int, float)):
        return val
    s = str(val).strip()
    if s in ("", "-", "–", "−"):
        return "-"
    # Các giá trị kiểu '1.000', '9.000', '1.150', '1.800' trong Sheet1 gốc là 1000, 9000, 1150, 1800
    if s in ("1.000", "9.000", "1.150", "1.800", "1.500", "1.250", "3.000"):
        return int(s.replace(".", ""))
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return s


def build_complete_chemical_db(base_xlsx_path: Optional[Path] = None) -> List[Dict[str, Any]]:
    """
    Xây dựng danh sách tra cứu hóa chất hoàn chỉnh cho Sheet1 của TinhToan_HoiKhiDoc.xlsx.
    Mỗi phần tử gồm:
      stt, name, twa_3733, twa_03_2019, twa_01_2025,
      stel_3733, stel_03_2019, stel_01_2025, twa_31_2026
    """
    if base_xlsx_path is None:
        base_xlsx_path = _find_sample("Q326D033771", "Tinh Toan Hoi Khi Doc.xlsx")

    wb = openpyxl.load_workbook(base_xlsx_path, data_only=True)
    ws = wb["Sheet1"]

    records: List[Dict[str, Any]] = []
    by_name: Dict[str, Dict[str, Any]] = {}

    for r in range(2, ws.max_row + 1):
        stt = ws.cell(r, 1).value
        name = str(ws.cell(r, 2).value or "").strip()
        if not name:
            continue
        twa_3733 = _clean_num(ws.cell(r, 3).value)

        # Lấy TWA & STEL chuẩn từ bảng QCVN 03:2019/BYT (50 chất)
        if name in QCVN_03_2019_MAP:
            t03, s03 = QCVN_03_2019_MAP[name]
            twa_03_2019 = t03 if t03 is not None else "-"
            stel_03_2019 = s03 if s03 is not None else "-"
        else:
            twa_03_2019 = "-"
            stel_03_2019 = "-"

        # Lấy TWA & STEL chuẩn từ bảng QCVN 01:2025/BYT (70 chất)
        if name in QCVN_01_2025_MAP:
            t01, s01 = QCVN_01_2025_MAP[name]
            twa_01_2025 = t01 if t01 is not None else "-"
            stel_01_2025 = s01 if s01 is not None else "-"
        else:
            twa_01_2025 = "-"
            stel_01_2025 = "-"

        stel_3733 = QD3733_STEL_MAP.get(name, "-")
        twa_31_2026 = QCVN_31_2026_TWA_MAP.get(name, "-")

        rec = {
            "stt": stt,
            "name": name,
            "twa_3733": twa_3733,
            "twa_03_2019": twa_03_2019,
            "twa_01_2025": twa_01_2025,
            "stel_3733": stel_3733,
            "stel_03_2019": stel_03_2019,
            "stel_01_2025": stel_01_2025,
            "twa_31_2026": twa_31_2026,
        }
        records.append(rec)
        by_name[name] = rec

    # Bổ sung các chất mới (Butyl acrylat...) và toàn bộ tên đồng nghĩa / công thức / tên KQTN
    existing_lower = {k.lower(): k for k in by_name}
    next_stt = len(records) + 1

    for alias_name, ref_name, override_vals in NEW_AND_ALIAS_CHEMICALS:
        if alias_name in by_name:
            continue
        if override_vals is not None:
            t3733, t03, t01, s3733, s03, s01, t31 = override_vals
            rec = {
                "stt": next_stt,
                "name": alias_name,
                "twa_3733": t3733,
                "twa_03_2019": t03,
                "twa_01_2025": t01,
                "stel_3733": s3733,
                "stel_03_2019": s03,
                "stel_01_2025": s01,
                "twa_31_2026": t31,
            }
            records.append(rec)
            by_name[alias_name] = rec
            existing_lower[alias_name.lower()] = alias_name
            next_stt += 1
        elif ref_name and ref_name in by_name:
            src = by_name[ref_name]
            rec = {
                "stt": next_stt,
                "name": alias_name,
                "twa_3733": src["twa_3733"],
                "twa_03_2019": src["twa_03_2019"],
                "twa_01_2025": src["twa_01_2025"],
                "stel_3733": src["stel_3733"],
                "stel_03_2019": src["stel_03_2019"],
                "stel_01_2025": src["stel_01_2025"],
                "twa_31_2026": src["twa_31_2026"],
            }
            records.append(rec)
            by_name[alias_name] = rec
            existing_lower[alias_name.lower()] = alias_name
            next_stt += 1

    return records


def lookup_chemical_limit(
    name: str,
    db: Optional[List[Dict[str, Any]]] = None,
    prefer_stel: bool = False,
) -> Dict[str, Any]:
    """
    Tra cứu ngưỡng giới hạn tiếp xúc của một yếu tố hóa học theo tên bất kỳ.
    Ưu tiên: QCVN 01:2025/BYT -> QCVN 03:2019/BYT -> QĐ 3733/2002/QĐ-BYT.
    """
    if db is None:
        db = build_complete_chemical_db()
    q = unicodedata.normalize("NFC", name.strip()).lower()
    match = None
    for r in db:
        rn = unicodedata.normalize("NFC", r["name"]).lower()
        if rn == q:
            match = r
            break
    if match is None:
        for r in db:
            rn = unicodedata.normalize("NFC", r["name"]).lower()
            if q in rn or rn in q:
                match = r
                break
    if match is None:
        return {
            "found": False,
            "name": name,
            "effective_limit": None,
            "standard": None,
            "twa_3733": "-",
            "twa_03_2019": "-",
            "twa_01_2025": "-",
            "twa_31_2026": "-",
        }

    k01 = "stel_01_2025" if prefer_stel else "twa_01_2025"
    k03 = "stel_03_2019" if prefer_stel else "twa_03_2019"
    k37 = "stel_3733" if prefer_stel else "twa_3733"

    eff_limit = None
    std = None
    if isinstance(match.get(k01), (int, float)):
        eff_limit = match[k01]
        std = "QCVN 01:2025/BYT"
    elif isinstance(match.get(k03), (int, float)):
        eff_limit = match[k03]
        std = "QCVN 03:2019/BYT"
    elif isinstance(match.get(k37), (int, float)):
        eff_limit = match[k37]
        std = "Quyết định số 3733/2002/QĐ-BYT"

    return {
        "found": True,
        **match,
        "effective_limit": eff_limit,
        "standard": std,
    }


# ==============================================================================
# 2. SINH TEMPLATE CHUẨN: TINH TOÁN HƠI KHÍ ĐỘC
# ==============================================================================

def create_hoi_khi_doc_template(out_path: Optional[Path] = None) -> Path:
    """
    Sinh file template chuẩn Core/Templates/MTLD/TinhToan_HoiKhiDoc_template.xlsx:
    - Sheet 'hơi kim loai': giữ nguyên công thức động VLOOKUP, LET 2 chữ số có nghĩa,
      sửa lỗi #VALUE! khi nhân KQTN mg/L có dấu '<', và thêm cột Ngưỡng áp dụng, Quy chuẩn, Đánh giá.
    - Sheet 'Sheet1': cơ sở dữ liệu đầy đủ QĐ 3733/2002, QCVN 03:2019, QCVN 01:2025,
      STEL, QCVN 31:2026 và hơn 180 tên đồng nghĩa/công thức để VLOOKUP luôn khớp 100%.
    """
    if out_path is None:
        out_path = TEMPLATE_MTLD_DIR / "TinhToan_HoiKhiDoc_template.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    src_path = _find_sample("1019N2.25DG", "1019D-Tinh Toan Hoi Khi Doc.xlsx")
    wb = openpyxl.load_workbook(src_path, data_only=False)

    # 1. Cập nhật Sheet1 (Bảng dữ liệu tra cứu hóa chất)
    ws_data = wb["Sheet1"]
    db = build_complete_chemical_db()
    max_db_row = len(db) + 1

    headers_sheet1 = [
        "STT",
        "Tên hóa chất (Theo QĐ 3733 / QCVN / Tên thông dụng)",
        "TWA (QĐ 3733/2002) (mg/m³)",
        "TWA (QCVN 03:2019) (mg/m³)",
        "TWA (QCVN 01:2025) (mg/m³)",
        "STEL (QĐ 3733/2002) (mg/m³)",
        "STEL (QCVN 03:2019) (mg/m³)",
        "STEL (QCVN 01:2025) (mg/m³)",
        "TWA (QCVN 31:2026 - HL 09/04/2027) (mg/m³)",
        "Ngưỡng TWA hiện hành (mg/m³)",
        "Quy chuẩn / Tiêu chuẩn áp dụng (TWA)",
    ]

    font_hdr = Font(name="Times New Roman", size=11, bold=True)
    font_body = Font(name="Times New Roman", size=11, bold=False)
    fill_hdr = PatternFill("solid", fgColor="D9E1F2")
    thin_side = Side(style="thin", color="808080")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for c_idx, h_text in enumerate(headers_sheet1, 1):
        cell = ws_data.cell(1, c_idx, h_text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_all

    for idx, rec in enumerate(db, start=2):
        vals = [
            rec["stt"],
            rec["name"],
            rec["twa_3733"],
            rec["twa_03_2019"],
            rec["twa_01_2025"],
            rec["stel_3733"],
            rec["stel_03_2019"],
            rec["stel_01_2025"],
            rec["twa_31_2026"],
            f'=IF(ISNUMBER(E{idx}),E{idx},IF(ISNUMBER(D{idx}),D{idx},IF(ISNUMBER(C{idx}),C{idx},"-")))',
            f'=IF(ISNUMBER(E{idx}),"QCVN 01:2025/BYT",IF(ISNUMBER(D{idx}),"QCVN 03:2019/BYT",IF(ISNUMBER(C{idx}),"QĐ 3733/2002/QĐ-BYT","-")))',
        ]
        for c_idx, v in enumerate(vals, 1):
            cell = ws_data.cell(idx, c_idx, v)
            cell.font = font_body
            cell.border = border_all
            if c_idx == 2 or c_idx == 11:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths_s1 = [8, 44, 18, 18, 18, 18, 18, 18, 22, 20, 26]
    for c_idx, w in enumerate(col_widths_s1, 1):
        ws_data.column_dimensions[get_column_letter(c_idx)].width = w

    # 2. Cập nhật Sheet 'hơi kim loai'
    ws_calc = wb["hơi kim loai"]
    calc_headers = [
        "STT",
        "Vị trí lấy mẫu",
        "Lưu lượng lấy mẫu, l/ph",
        "Thời gian lấy mẫu, ph",
        "Thể tích mẫu, L",
        "Phiếu KQTN",
        "Tên chỉ tiêu",
        "KQTN,\nmg/L",
        "Thể tích dung dịch\nmL",
        "KQTN, µg/mẫu",
        "KQTT, mg/m3",
        "TWA (QĐ 3733/2002) (mg/m³)",
        "TWA (QCVN 03:2019) (mg/m³)",
        "TWA (QCVN 01:2025) (mg/m³)",
        "TWA (QCVN 31:2026) (mg/m³)",
        "Ngưỡng TWA áp dụng (mg/m³)",
        "Quy chuẩn áp dụng",
        "Đánh giá",
    ]
    for c_idx, h_text in enumerate(calc_headers, 1):
        cell = ws_calc.cell(1, c_idx, h_text)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_all

    # Thiết lập công thức mẫu chuẩn cho 30 dòng dữ liệu (dòng 2 -> 31)
    for r in range(2, 32):
        # Giữ nguyên dữ liệu mẫu ở dòng 2..16 của 1019D để làm ví dụ minh họa, xóa dòng 17..31
        if r > 16:
            for c_idx in range(1, 19):
                ws_calc.cell(r, c_idx, None)
            ws_calc.cell(r, 1, r - 1)
            ws_calc.cell(r, 3, 1)
            ws_calc.cell(r, 4, 30)
            ws_calc.cell(r, 8, "-")
            ws_calc.cell(r, 9, "-")

        # Cột E (5): Thể tích mẫu (L) = Lưu lượng * Thời gian
        e_val = ws_calc.cell(r, 5).value
        if e_val != "-":
            ws_calc.cell(
                r,
                5,
                f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}*D{r},"-")',
            )

        # Cột J (10): KQTN, µg/mẫu — nếu H (mg/L) và I (mL) có số thì tự nhân (an toàn dấu phẩy/chấm)
        h_val = ws_calc.cell(r, 8).value
        i_val = ws_calc.cell(r, 9).value
        if h_val not in (None, "-", "") and i_val not in (None, "-", ""):
            ws_calc.cell(
                r,
                10,
                f'=IF(ISNUMBER(H{r}),H{r}*I{r},_xlfn.TEXTBEFORE(H{r}," ")&" "&_xlfn.NUMBERVALUE(_xlfn.TEXTAFTER(H{r}," "),",",".")*I{r})',
            )

        # Cột K (11): KQTT, mg/m3 (làm tròn 2 chữ số có nghĩa theo chuẩn 1019D/Q326D033771)
        if ws_calc.cell(r, 5).value != "-":
            ws_calc.cell(
                r,
                11,
                f'=IF(OR(J{r}="",E{r}="-",E{r}=0),"-",_xlfn.LET(\n'
                f'_xlpm.v,IF(ISNUMBER(J{r}),J{r}/$E{r},_xlfn.NUMBERVALUE(_xlfn.TEXTAFTER(J{r}," "),",",".")/$E{r}),\n'
                f'_xlpm.rd,2-1-INT(LOG10(ABS(_xlpm.v))),\n'
                f'_xlpm.r,ROUND(_xlpm.v,_xlpm.rd),\n'
                f'_xlpm.dec,MAX(0,_xlpm.rd),\n'
                f'_xlpm.fmt,IF(_xlpm.dec=0,"0","0,"&REPT("0",_xlpm.dec)),\n'
                f'_xlpm.txt,TEXT(_xlpm.r,_xlpm.fmt),\n'
                f'IF(ISNUMBER(J{r}),_xlpm.txt,_xlfn.TEXTBEFORE(J{r}," ")&" "&_xlpm.txt)))',
            )

        # Cột L, M, N, O, P, Q: VLOOKUP từ Sheet1!$B$2:$K$<max_db_row>
        ws_calc.cell(
            r,
            12,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},2,FALSE),"-"))',
        )
        ws_calc.cell(
            r,
            13,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},3,FALSE),"-"))',
        )
        ws_calc.cell(
            r,
            14,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},4,FALSE),"-"))',
        )
        ws_calc.cell(
            r,
            15,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},8,FALSE),"-"))',
        )
        ws_calc.cell(
            r,
            16,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},9,FALSE),"-"))',
        )
        ws_calc.cell(
            r,
            17,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},10,FALSE),"-"))',
        )
        # Cột R (18): Đánh giá Đạt / Không đạt
        ws_calc.cell(
            r,
            18,
            f'=IF(OR(K{r}="-",K{r}="",P{r}="-",P{r}=""),"-",'
            f'IF(ISNUMBER(K{r}),IF(K{r}<=P{r},"Đạt","Không đạt"),'
            f'IF(LEFT(TRIM(K{r}),1)="<","Đạt",'
            f'IF(_xlfn.NUMBERVALUE(K{r},",",".")<=P{r},"Đạt","Không đạt"))))',
        )

        fill_none = PatternFill(fill_type=None)
        fill_kqtt = PatternFill("solid", fgColor="C6EFCE")
        for c_idx in range(1, 19):
            c_cell = ws_calc.cell(r, c_idx)
            c_cell.font = font_body
            c_cell.border = border_all
            c_cell.fill = fill_kqtt if c_idx == 11 else fill_none
            if c_idx in (2, 7, 17):
                c_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                c_cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths_calc = [6, 28, 12, 12, 12, 18, 24, 12, 12, 14, 15, 16, 16, 16, 16, 16, 22, 14]
    for c_idx, w in enumerate(col_widths_calc, 1):
        ws_calc.column_dimensions[get_column_letter(c_idx)].width = w

    wb.save(out_path)
    return out_path


# ==============================================================================
# 3. SINH TEMPLATE CHUẨN: VI KHÍ HẬU - VẬT LÝ - BỤI - CO2 - BỨC XẠ - RUNG - ĐTT
# ==============================================================================

def create_vikhihau_vatly_bui_template(out_path: Optional[Path] = None) -> Path:
    """
    Sinh file template chuẩn Core/Templates/MTLD/TinhToan_ViKhiHau_VatLy_Bui_template.xlsx:
    - Kết hợp mẫu 2026 Q326D033771 (Sheet1, NhietDo, DoAm, Gio, AnhSang, On, DataOn, Bui, CO2)
      với OnDaiTan, BucXaNhiet (từ 1019N2.25DG), Rung (từ 1126N2.25DG) và DienTuTruong (từ 334N2.25DG).
    - Sửa triệt để lỗi #REF! trong sheet CO2 (dòng 3..15) và lỗi công thức Không đạt của AnhSang
      (=IF(F3<G3,F3,"-") thay vì so D3 với G3).
    """
    if out_path is None:
        out_path = TEMPLATE_MTLD_DIR / "TinhToan_ViKhiHau_VatLy_Bui_template.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    src_2026 = _find_sample("Q326D033771", "Nhiet Am, CO2, Do sang, Gio.xlsx")
    wb = openpyxl.load_workbook(src_2026, data_only=False)

    # 1. Cắt bỏ các dòng trống thừa (> dòng 25) ở các sheet của Q326D033771 (tránh xuất PDF hàng trăm trang trắng)
    for sname in ("Sheet1", "NhietDo", "DoAm", "Gio", "AnhSang", "CO2", "Bui"):
        if sname in wb.sheetnames:
            ws_trim = wb[sname]
            if ws_trim.max_row > 25:
                ws_trim.delete_rows(26, ws_trim.max_row - 25)

    # 2. Sửa lỗi công thức ở Sheet AnhSang (Cột H & I: so sánh kết quả sau hiệu chỉnh F với ngưỡng G)
    if "AnhSang" in wb.sheetnames:
        ws_as = wb["AnhSang"]
        for r in range(3, 18):
            if ws_as.cell(r, 4).value is not None:
                ws_as.cell(r, 8, f'=IF(F{r}>=G{r},ROUND(F{r},0),"-")')
                ws_as.cell(r, 9, f'=IF(F{r}<G{r},ROUND(F{r},0),"-")')

    # 3. Sửa triệt để lỗi #REF! trong Sheet CO2 (tham chiếu đồng bộ Sheet1!B{r}, Sheet1!D{r}, Sheet1!J{r})
    if "CO2" in wb.sheetnames:
        from openpyxl.cell.cell import MergedCell
        ws_co2 = wb["CO2"]
        for idx, r in enumerate(range(3, 18), start=1):
            if not isinstance(ws_co2.cell(r, 1), MergedCell):
                ws_co2.cell(r, 1, idx)
            if not isinstance(ws_co2.cell(r, 2), MergedCell):
                ws_co2.cell(r, 2, f"=Sheet1!B{r}")
            if not isinstance(ws_co2.cell(r, 3), MergedCell):
                ws_co2.cell(r, 3, f"=Sheet1!D{r}")
            if not isinstance(ws_co2.cell(r, 4), MergedCell):
                ws_co2.cell(r, 4, f"=Sheet1!J{r}")

    # 3. Bổ sung Sheet Rung (theo QCVN 27:2016/BYT) và DienTuTruong (theo QCVN 25:2016/BYT)
    font_hdr = Font(name="Times New Roman", size=11, bold=True)
    font_body = Font(name="Times New Roman", size=11, bold=False)
    thin_side = Side(style="thin", color="808080")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    if "Rung" not in wb.sheetnames:
        ws_rung = wb.create_sheet("Rung")
        ws_rung.cell(1, 1, "=Sheet1!A1").font = font_hdr
        ws_rung.cell(1, 5, "=Sheet1!E1").font = font_hdr
        ws_rung.cell(2, 2, "KẾT QUẢ ĐO RUNG (QCVN 27:2016/BYT)").font = Font(
            name="Times New Roman", size=12, bold=True
        )
        headers_rung = [
            "STT",
            "Vị trí đo",
            "Lần 1 (m/s²)",
            "Lần 2 (m/s²)",
            "Lần 3 (m/s²)",
            "Trung bình (m/s²)",
            "Giới hạn QCVN 27:2016/BYT (m/s²)",
            "Mẫu đạt",
            "Mẫu không đạt",
        ]
        for c_idx, h_text in enumerate(headers_rung, 1):
            c = ws_rung.cell(4, c_idx, h_text)
            c.font = font_hdr
            c.border = border_all
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for idx, r in enumerate(range(5, 20), start=1):
            ws_rung.cell(r, 1, idx)
            ws_rung.cell(r, 2, f"=Sheet1!B{idx+2}")
            ws_rung.cell(r, 3, 0.020)
            ws_rung.cell(r, 4, 0.022)
            ws_rung.cell(r, 5, 0.021)
            ws_rung.cell(r, 6, f"=ROUND(AVERAGE(C{r}:E{r}),3)")
            ws_rung.cell(r, 7, 0.086)
            ws_rung.cell(r, 8, f'=IF(F{r}<=G{r},F{r},"-")')
            ws_rung.cell(r, 9, f'=IF(F{r}>G{r},F{r},"-")')
            for c_idx in range(1, 10):
                ws_rung.cell(r, c_idx).font = font_body
                ws_rung.cell(r, c_idx).border = border_all
        ws_rung.cell(20, 2, "Tổng cộng").font = font_hdr
        ws_rung.cell(20, 8, "=COUNT(H5:H19)").font = font_hdr
        ws_rung.cell(20, 9, "=COUNT(I5:I19)").font = font_hdr

    if "DienTuTruong" not in wb.sheetnames:
        ws_dtt = wb.create_sheet("DienTuTruong")
        ws_dtt.cell(1, 1, "=Sheet1!A1").font = font_hdr
        ws_dtt.cell(1, 5, "=Sheet1!E1").font = font_hdr
        ws_dtt.cell(2, 2, "KẾT QUẢ ĐO ĐIỆN TỪ TRƯỜNG TẦN SỐ CÔNG NGHIỆP (QCVN 25:2016/BYT)").font = Font(
            name="Times New Roman", size=12, bold=True
        )
        headers_dtt = [
            "STT",
            "Vị trí đo",
            "Cường độ điện trường E (V/m)",
            "Giới hạn E (V/m)",
            "Điện trường Đạt",
            "Điện trường KĐ",
            "Cường độ từ trường H (A/m)",
            "Giới hạn H (A/m)",
            "Từ trường Đạt",
            "Từ trường KĐ",
        ]
        for c_idx, h_text in enumerate(headers_dtt, 1):
            c = ws_dtt.cell(4, c_idx, h_text)
            c.font = font_hdr
            c.border = border_all
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for idx, r in enumerate(range(5, 15), start=1):
            ws_dtt.cell(r, 1, idx)
            ws_dtt.cell(r, 2, f"=Sheet1!B{idx+2}")
            ws_dtt.cell(r, 3, 15.0)
            ws_dtt.cell(r, 4, 5000)
            ws_dtt.cell(r, 5, f'=IF(C{r}<=D{r},C{r},"-")')
            ws_dtt.cell(r, 6, f'=IF(C{r}>D{r},C{r},"-")')
            ws_dtt.cell(r, 7, 0.08)
            ws_dtt.cell(r, 8, 80)
            ws_dtt.cell(r, 9, f'=IF(G{r}<=H{r},G{r},"-")')
            ws_dtt.cell(r, 10, f'=IF(G{r}>H{r},G{r},"-")')
            for c_idx in range(1, 11):
                ws_dtt.cell(r, c_idx).font = font_body
                ws_dtt.cell(r, c_idx).border = border_all
        ws_dtt.cell(15, 2, "Tổng cộng").font = font_hdr
        ws_dtt.cell(15, 5, "=COUNT(E5:E14)").font = font_hdr
        ws_dtt.cell(15, 6, "=COUNT(F5:F14)").font = font_hdr
        ws_dtt.cell(15, 9, "=COUNT(I5:I14)").font = font_hdr
        ws_dtt.cell(15, 10, "=COUNT(J5:J14)").font = font_hdr

    update_mtld_excel_headers(
        wb,
        dossier_code="Q326D033771",
        date_str="29/04/2026",
        file_suffix="Tính Toán Nhiệt Ẩm",
    )
    wb.save(out_path)
    return out_path


# ==============================================================================
# 4. SINH TEMPLATE CHUẨN: TÂM SINH LÝ LAO ĐỘNG & ECGONOMY (OWAS + VỊ TRÍ LĐ)
# ==============================================================================

def create_ecgonomy_template(out_path: Optional[Path] = None) -> Path:
    """
    Sinh file template chuẩn Core/Templates/MTLD/TinhToan_Ecgonomy_template.xlsx
    từ 1019N2.25DG - Ecgonomy.xlsx, kèm thêm sheet GanhNangTriOc (từ 334N2.25DG).
    """
    if out_path is None:
        out_path = TEMPLATE_MTLD_DIR / "TinhToan_Ecgonomy_template.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    src_eco = _find_sample("1019N2.25DG", "Ecgonomy.xlsx")
    wb = openpyxl.load_workbook(src_eco, data_only=False)

    if "GanhNangTriOc" not in wb.sheetnames:
        ws_gn = wb.create_sheet("GanhNangTriOc")
        font_hdr = Font(name="Times New Roman", size=11, bold=True)
        font_body = Font(name="Times New Roman", size=11, bold=False)
        thin_side = Side(style="thin", color="808080")
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        ws_gn.cell(1, 1, "ĐÁNH GIÁ GÁNH NẶNG LAO ĐỘNG TRÍ ÓC (TRÍ NHỚ SỐ & TRÍ NHỚ HÌNH)").font = Font(
            name="Times New Roman", size=12, bold=True
        )
        headers_gn = [
            "STT",
            "Phòng ban / Bộ phận",
            "Họ và tên nhân viên",
            "Ngày thực hiện",
            "Trí nhớ số - Đầu ca",
            "Trí nhớ số - Cuối ca",
            "Mức suy giảm Trí nhớ số (%)",
            "Trí nhớ hình - Đầu ca",
            "Trí nhớ hình - Cuối ca",
            "Mức suy giảm Trí nhớ hình (%)",
            "Đánh giá chung",
        ]
        for c_idx, h_text in enumerate(headers_gn, 1):
            c = ws_gn.cell(3, c_idx, h_text)
            c.font = font_hdr
            c.border = border_all
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for idx, r in enumerate(range(4, 14), start=1):
            ws_gn.cell(r, 1, idx)
            ws_gn.cell(r, 2, f"Bộ phận {idx}")
            ws_gn.cell(r, 3, f"Nhân viên {idx}")
            ws_gn.cell(r, 4, "06/11/2026")
            ws_gn.cell(r, 5, 10)
            ws_gn.cell(r, 6, 10)
            ws_gn.cell(
                r,
                7,
                f'=IF(E{r}<=0,"-",IF((E{r}-F{r})/E{r}<0.05,"< 5%",TEXT(ROUND((E{r}-F{r})/E{r}*100,0),"0")&"%"))',
            )
            ws_gn.cell(r, 8, 8)
            ws_gn.cell(r, 9, 8)
            ws_gn.cell(
                r,
                10,
                f'=IF(H{r}<=0,"-",IF((H{r}-I{r})/H{r}<0.05,"< 5%",TEXT(ROUND((H{r}-I{r})/H{r}*100,0),"0")&"%"))',
            )
            ws_gn.cell(
                r,
                11,
                f'=IF(AND( (E{r}-F{r})/E{r}<=0.2, (H{r}-I{r})/H{r}<=0.2 ),"Đạt","Cần lưu ý")',
            )
            for c_idx in range(1, 12):
                ws_gn.cell(r, c_idx).font = font_body
                ws_gn.cell(r, c_idx).border = border_all

    update_mtld_excel_headers(
        wb,
        dossier_code="Q326D033771",
        date_str="29/04/2026",
        file_suffix="Tính Toán Ecgonomy",
    )
    wb.save(out_path)
    return out_path


# ==============================================================================
# 5. SINH BỘ WORD TEMPLATE CHUẨN 2026 (TBKQ & BCKQ)
# ==============================================================================

def create_word_templates() -> Tuple[Path, Path]:
    """
    Sao chép và chuẩn hóa bộ Word template 2026 cho MTLD:
    - Core/Templates/MTLD/FORM - TBKQ - viet - mtld - 2026.docx (từ Q326D033771 chuẩn 2026)
    - Core/Templates/MTLD/FORM - TBKQ - viet - mtld (Full Factors) - 2026.docx (từ 1019N2.25DG nâng cấp header 2026)
    - Core/Templates/MTLD/FORM - BCKQ - mtld - 2026.docx (từ Q326D033771 - BCKQDG.docx chuẩn BM3030.03 HL:01/8/2025)
    Đồng thời tạo bản tại Core/Templates/TBKQ - viet - mtld.docx và Core/Templates/BCKQ - mtld.docx.
    """
    import docx

    TEMPLATE_MTLD_DIR.mkdir(parents=True, exist_ok=True)

    src_tbkq_2026 = _find_sample("Q326D033771", "BÁO CÁO KẾT QUẢ QTMTLD")
    src_tbkq_full = _find_sample("1126N2.25DG", "1126DG - BÁO CÁO KẾT QUẢ QTMTLĐ- 1")
    src_bckq_2026 = _find_sample("Q326D033771", "Q326D033771 - BCKQDG.docx")

    dst_tbkq_2026 = TEMPLATE_MTLD_DIR / "FORM - TBKQ - viet - mtld - 2026.docx"
    dst_tbkq_full = TEMPLATE_MTLD_DIR / "FORM - TBKQ - viet - mtld (Full Master) - 2026.docx"
    dst_bckq_2026 = TEMPLATE_MTLD_DIR / "FORM - BCKQ - mtld - 2026.docx"

    from Core.Builders.handlers.mtld import (
        ensure_mtld_table_headers_and_cantsplit,
        fix_mtld_header_and_page_fields,
    )

    doc_2026 = docx.Document(src_tbkq_2026)
    ensure_mtld_table_headers_and_cantsplit(doc_2026, detach_summary_rows=False)
    fix_mtld_header_and_page_fields(doc_2026, "Q326D033771", "29/04/2026", is_bckq=False)
    doc_2026.save(dst_tbkq_2026)

    doc_bckq = docx.Document(src_bckq_2026)
    ensure_mtld_table_headers_and_cantsplit(doc_bckq, detach_summary_rows=False)
    fix_mtld_header_and_page_fields(doc_bckq, "Q326D033771", "22/04/2026", is_bckq=True)
    doc_bckq.save(dst_bckq_2026)

    # Nâng cấp bản Full Master (1126N2.25DG) lên thông tin cơ quan chủ quản & địa chỉ 2026
    doc_full = docx.Document(src_tbkq_full)
    ensure_mtld_table_headers_and_cantsplit(doc_full, detach_summary_rows=False)
    # Table 0 trang bìa: cập nhật "BỘ KHOA HỌC VÀ CÔNG NGHỆ / TỔNG CỤC..." -> "ỦY BAN TIÊU CHUẨN ĐO LƯỜNG CHẤT LƯỢNG QUỐC GIA"
    if len(doc_full.tables) > 0:
        t0 = doc_full.tables[0]
        for cell in t0.rows[0].cells:
            for p in cell.paragraphs:
                if "BỘ KHOA HỌC" in p.text or "TỔNG CỤC" in p.text:
                    for r in p.runs:
                        r.text = ""
                    if p.runs:
                        p.runs[0].text = "ỦY BAN TIÊU CHUẨN ĐO LƯỜNG CHẤT LƯỢNG QUỐC GIA"
                    else:
                        p.add_run("ỦY BAN TIÊU CHUẨN ĐO LƯỜNG CHẤT LƯỢNG QUỐC GIA")
    # Table 2: cập nhật địa chỉ "phường Bến Nghé, Quận 1" -> "phường Sài Gòn"
    if len(doc_full.tables) > 2:
        t2 = doc_full.tables[2]
        for row in t2.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "49 Pasteur" in p.text:
                        for r in p.runs:
                            r.text = ""
                        if p.runs:
                            p.runs[0].text = "Địa chỉ: 49 Pasteur, phường Sài Gòn, Thành phố Hồ Chí Minh"
    doc_full.save(dst_tbkq_full)

    # Tạo bản chuẩn ở gốc Core/Templates/
    shutil.copy2(dst_tbkq_2026, ROOT_DIR / "Core" / "Templates" / "TBKQ - viet - mtld.docx")
    shutil.copy2(dst_bckq_2026, ROOT_DIR / "Core" / "Templates" / "BCKQ - mtld.docx")
    shutil.copy2(
        TEMPLATE_MTLD_DIR / "TinhToan_HoiKhiDoc_template.xlsx",
        ROOT_DIR / "Core" / "Templates" / "TinhToanHoiKhiDoc_template.xlsx",
    )
    shutil.copy2(
        TEMPLATE_MTLD_DIR / "TinhToan_ViKhiHau_VatLy_Bui_template.xlsx",
        ROOT_DIR / "Core" / "Templates" / "TinhToanMTLD_template.xlsx",
    )
    shutil.copy2(
        TEMPLATE_MTLD_DIR / "TinhToan_Ecgonomy_template.xlsx",
        ROOT_DIR / "Core" / "Templates" / "TinhToanEcgonomy_template.xlsx",
    )

    return dst_tbkq_2026, dst_bckq_2026


# ==============================================================================
# 6. HÀM ĐIỀN DỮ LIỆU TÍNH TOÁN TỪ TEMPLATE & XUẤT PDF (GIỮ 100% CÔNG THỨC ĐỘNG)
# ==============================================================================

def update_mtld_excel_headers(
    wb: openpyxl.Workbook,
    dossier_code: str,
    date_str: str = "",
    file_suffix: str = "Tính Toán Nhiệt Ẩm",
) -> None:
    """
    Chuẩn hóa PageSetup Header của tất cả các sheet trong file tính toán Excel:
    - Góc trái (`oddHeader.left.text`): `Số: <Mã hồ sơ>/<Đuôi file>` (vd `Số: Q326D999901/Tính Toán Nhiệt Ẩm`,
      `Số: Q326D999901/Tính Toán Hơi Khí Độc`, `Số: Q326D999901/Tính Toán Ecgonomy`).
    - Góc giữa (`oddHeader.center.text`): loại bỏ các ký tự xuống dòng thừa (`\\n\\n`) ở cuối tiêu đề sheet.
    - Góc phải (`oddHeader.right.text`): `Ngày <dd/mm/yyyy>`.
    """
    left_label = f"Số: {dossier_code}/{file_suffix}" if file_suffix else f"Số: {dossier_code}"
    right_label = (
        date_str
        if (not date_str or date_str.strip().lower().startswith("ngày"))
        else f"Ngày {date_str.strip()}"
    )

    for ws in wb.worksheets:
        # Không gắn header in cho Sheet1 tra cứu 678 chất của file Hơi khí độc
        if ws.title == "Sheet1" and "hơi kim loai" in wb.sheetnames:
            continue
        ws.oddHeader.left.text = left_label
        if ws.oddHeader.center.text:
            ws.oddHeader.center.text = ws.oddHeader.center.text.strip()
        if right_label:
            ws.oddHeader.right.text = right_label


def populate_vikhihau_vatly_bui_xlsx(
    out_path: Path,
    locations_main: List[Dict[str, Any]],
    company_name: str,
    dossier_code: str,
    inspection_date_str: str,
    template_path: Optional[Path] = None,
    file_suffix: str = "Tính Toán Nhiệt Ẩm",
) -> Path:
    """
    Sinh phiếu tính toán `<Mã> - TinhToanViKhiHauVatLyBui.xlsx` từ template chuẩn 2026,
    giữ nguyên 100% công thức động liên kết từ `Sheet1` sang `NhietDo`, `DoAm`, `Gio`,
    `AnhSang`, `On`, `Bui`, `CO2`, `Rung`, `DienTuTruong` và gắn Header chuẩn
    `Số: <Mã>/Tính Toán Nhiệt Ẩm`.
    """
    if template_path is None:
        template_path = TEMPLATE_MTLD_DIR / "TinhToan_ViKhiHau_VatLy_Bui_template.xlsx"
        if not template_path.exists():
            create_vikhihau_vatly_bui_template(template_path)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(template_path, data_only=False)
    ws1 = wb["Sheet1"]

    for idx, loc in enumerate(locations_main):
        r = 3 + idx
        ws1.cell(r, 1, idx + 1)
        ws1.cell(r, 2, company_name if idx == 0 else None)
        ws1.cell(r, 4, loc.get("code", f"T{idx+1}"))
        ws1.cell(r, 5, inspection_date_str)
        ws1.cell(r, 7, loc.get("temp"))
        ws1.cell(r, 8, loc.get("hum"))
        ws1.cell(r, 9, loc.get("lux_raw", loc.get("lux")))
        ws1.cell(r, 10, loc.get("co2_ppm"))
        ws1.cell(r, 11, loc.get("wind"))
        ws1.cell(r, 12, loc.get("noise"))
        if loc.get("dust_resp") is not None:
            ws1.cell(r, 17, float(loc["dust_resp"]) * 1000.0)

    update_mtld_excel_headers(
        wb,
        dossier_code=dossier_code,
        date_str=inspection_date_str,
        file_suffix=file_suffix,
    )
    wb.save(out_path)
    return out_path


def populate_ecgonomy_xlsx(
    out_path: Path,
    locations: List[Dict[str, Any]],
    dossier_code: str = "",
    date_str: str = "",
    template_path: Optional[Path] = None,
    file_suffix: str = "Tính Toán Ecgonomy",
) -> Path:
    """
    Sinh phiếu tính toán `<Mã> - TinhToanEcgonomy.xlsx` từ template chuẩn 2026
    và gắn Header chuẩn `Số: <Mã>/Tính Toán Ecgonomy`.
    """
    if template_path is None:
        template_path = TEMPLATE_MTLD_DIR / "TinhToan_Ecgonomy_template.xlsx"
        if not template_path.exists():
            create_ecgonomy_template(template_path)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(template_path, data_only=False)
    if "Tư thế LĐ" in wb.sheetnames:
        ws_tt = wb["Tư thế LĐ"]
        for idx, loc in enumerate(locations, start=1):
            r = idx + 3
            ws_tt.cell(r, 1, idx)
            ws_tt.cell(r, 2, loc.get("name", loc.get("location", f"Vị trí {idx}")))

    if dossier_code:
        update_mtld_excel_headers(
            wb,
            dossier_code=dossier_code,
            date_str=date_str,
            file_suffix=file_suffix,
        )
    wb.save(out_path)
    return out_path


def populate_hoi_khi_doc_xlsx(
    out_path: Path,
    samples: List[Dict[str, Any]],
    template_path: Optional[Path] = None,
    dossier_code: str = "",
    date_str: str = "",
    file_suffix: str = "Tính Toán Hơi Khí Độc",
) -> Path:
    """
    Điền dữ liệu mẫu hơi khí độc vào bản sao của TinhToan_HoiKhiDoc_template.xlsx,
    giữ nguyên 100% công thức động (VLOOKUP, LET 2 chữ số có nghĩa, đánh giá Đạt/Không đạt).

    Mỗi phần tử trong `samples` là dict:
      - location (str): Vị trí lấy mẫu
      - flow_lpm (float | str): Lưu lượng lấy mẫu (l/ph), hoặc "-" nếu đo trực tiếp/ống phát hiện
      - time_min (float | str): Thời gian lấy mẫu (ph), hoặc "-"
      - kqtn_no (str): Số phiếu KQTN (hoặc "Ống phát hiện nhanh" / "Đo trực tiếp")
      - parameter (str): Tên chỉ tiêu (tự động khớp VLOOKUP trên Sheet1)
      - kqtn_mg_l (float | str): KQTN mg/L (mặc định "-")
      - vol_sol_ml (float | str): Thể tích dung dịch mL (mặc định "-")
      - kqtn_ug (float | str | None): KQTN µg/mẫu (nếu đã có thẳng µg/mẫu)
      - direct_mg_m3 (float | str | None): Kết quả đo trực tiếp mg/m³ (nếu không qua hấp thụ)
    """
    if template_path is None:
        template_path = TEMPLATE_MTLD_DIR / "TinhToan_HoiKhiDoc_template.xlsx"
        if not template_path.exists():
            create_hoi_khi_doc_template(template_path)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(template_path, data_only=False)
    ws = wb["hơi kim loai"]
    ws_data = wb["Sheet1"]
    max_db_row = ws_data.max_row

    font_body = Font(name="Times New Roman", size=11, bold=False)
    thin_side = Side(style="thin", color="808080")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Xóa dữ liệu mẫu cũ từ dòng 2 trở đi
    max_clear = max(ws.max_row, len(samples) + 5)
    for r in range(2, max_clear + 1):
        for c in range(1, 19):
            ws.cell(r, c, None)

    for idx, item in enumerate(samples, start=1):
        r = idx + 1
        loc = item.get("location", "")
        flow = item.get("flow_lpm", 1)
        t_min = item.get("time_min", 30)
        kqtn_no = item.get("kqtn_no", "")
        param = item.get("parameter", "")
        kqtn_mg_l = item.get("kqtn_mg_l", "-")
        vol_sol_ml = item.get("vol_sol_ml", "-")
        kqtn_ug = item.get("kqtn_ug", None)
        direct_mg_m3 = item.get("direct_mg_m3", None)

        ws.cell(r, 1, idx)
        ws.cell(r, 2, loc)
        ws.cell(r, 3, flow)
        ws.cell(r, 4, t_min)

        if direct_mg_m3 is not None or flow == "-" or t_min == "-":
            ws.cell(r, 5, "-")
        else:
            ws.cell(r, 5, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}*D{r},"-")')

        ws.cell(r, 6, kqtn_no)
        ws.cell(r, 7, param)
        ws.cell(r, 8, kqtn_mg_l)
        ws.cell(r, 9, vol_sol_ml)

        if kqtn_mg_l not in (None, "-", "") and vol_sol_ml not in (None, "-", ""):
            ws.cell(
                r,
                10,
                f'=IF(ISNUMBER(H{r}),H{r}*I{r},_xlfn.TEXTBEFORE(H{r}," ")&" "&_xlfn.NUMBERVALUE(_xlfn.TEXTAFTER(H{r}," "),",",".")*I{r})',
            )
        elif kqtn_ug is not None:
            ws.cell(r, 10, kqtn_ug)
        else:
            ws.cell(r, 10, "-")

        if direct_mg_m3 is not None:
            ws.cell(r, 11, direct_mg_m3)
        else:
            ws.cell(
                r,
                11,
                f'=IF(OR(J{r}="",J{r}="-",E{r}="-",E{r}=0),"-",_xlfn.LET(\n'
                f'_xlpm.v,IF(ISNUMBER(J{r}),J{r}/$E{r},_xlfn.NUMBERVALUE(_xlfn.TEXTAFTER(J{r}," "),",",".")/$E{r}),\n'
                f'_xlpm.rd,2-1-INT(LOG10(ABS(_xlpm.v))),\n'
                f'_xlpm.r,ROUND(_xlpm.v,_xlpm.rd),\n'
                f'_xlpm.dec,MAX(0,_xlpm.rd),\n'
                f'_xlpm.fmt,IF(_xlpm.dec=0,"0","0,"&REPT("0",_xlpm.dec)),\n'
                f'_xlpm.txt,TEXT(_xlpm.r,_xlpm.fmt),\n'
                f'IF(ISNUMBER(J{r}),_xlpm.txt,_xlfn.TEXTBEFORE(J{r}," ")&" "&_xlpm.txt)))',
            )

        ws.cell(
            r,
            12,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},2,FALSE),"-"))',
        )
        ws.cell(
            r,
            13,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},3,FALSE),"-"))',
        )
        ws.cell(
            r,
            14,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},4,FALSE),"-"))',
        )
        ws.cell(
            r,
            15,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},8,FALSE),"-"))',
        )
        ws.cell(
            r,
            16,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},9,FALSE),"-"))',
        )
        ws.cell(
            r,
            17,
            f'=IF(G{r}="","-",IFERROR(VLOOKUP($G{r},Sheet1!$B$2:$K${max_db_row},10,FALSE),"-"))',
        )
        ws.cell(
            r,
            18,
            f'=IF(OR(K{r}="-",K{r}="",P{r}="-",P{r}=""),"-",'
            f'IF(ISNUMBER(K{r}),IF(K{r}<=P{r},"Đạt","Không đạt"),'
            f'IF(LEFT(TRIM(K{r}),1)="<","Đạt",'
            f'IF(_xlfn.NUMBERVALUE(K{r},",",".")<=P{r},"Đạt","Không đạt"))))',
        )

        fill_none = PatternFill(fill_type=None)
        fill_kqtt = PatternFill("solid", fgColor="C6EFCE")
        for c_idx in range(1, 19):
            c_cell = ws.cell(r, c_idx)
            c_cell.font = font_body
            c_cell.border = border_all
            c_cell.fill = fill_kqtt if c_idx == 11 else fill_none
            if c_idx in (2, 7, 17):
                c_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r].height = 34
    ws.column_dimensions["B"].width = 36

    # Xóa hẳn các dòng thừa phía dưới bảng để khi in PDF không hiện dòng trống
    last_data_row = len(samples) + 1
    if ws.max_row > last_data_row:
        ws.delete_rows(last_data_row + 1, ws.max_row - last_data_row)

    # Cập nhật PageSetup Header của sheet tính toán
    if dossier_code:
        update_mtld_excel_headers(
            wb,
            dossier_code=dossier_code,
            date_str=date_str,
            file_suffix=file_suffix,
        )

    wb.save(out_path)
    return out_path


def export_mtld_pdfs(
    files: List[Path],
    pdf_dir: Path,
    vkh_active_sheets: Optional[List[str]] = None,
) -> List[Path]:
    """
    Tính toán lại toàn bộ công thức Excel (lưu giá trị cache vào file .xlsx) và xuất PDF
    cho các file DOCX/XLSX của hồ sơ MTLD vào `pdf_dir` (an toàn máy in tem qua `avoid_label_printer`).
    """
    import gc
    import win32com.client
    from Core.Exporters.pdf_exporter import avoid_label_printer, convert_docx_to_pdf

    pdf_dir = Path(pdf_dir)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    avoid_label_printer()
    exported: List[Path] = []
    active_set = set(vkh_active_sheets or ["Sheet1", "NhietDo", "DoAm", "Gio", "AnhSang", "Bui"])

    for f in files:
        f = Path(f)
        out_pdf = pdf_dir / f"{f.stem}.pdf"
        try:
            if f.suffix.lower() == ".docx":
                # Báo cáo QTMTLĐ theo Phụ lục III NĐ 44/2016 có trang bìa riêng (add_bg=False)
                convert_docx_to_pdf(str(f), str(out_pdf), add_bg=False)
                if out_pdf.exists():
                    exported.append(out_pdf)
            elif f.suffix.lower() == ".xlsx":
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                try:
                    wb = excel.Workbooks.Open(str(f.resolve()), ReadOnly=False)
                    wb.Application.CalculateFull()
                    wb.Save()
                    if "HoiKhiDoc" in f.name and "hơi kim loai" in [s.Name for s in wb.Worksheets]:
                        ws_calc = wb.Worksheets("hơi kim loai")
                        ws_calc.PageSetup.Orientation = 2  # xlLandscape
                        ws_calc.PageSetup.Zoom = False
                        ws_calc.PageSetup.FitToPagesWide = 1
                        ws_calc.PageSetup.FitToPagesTall = False
                        ws_calc.ExportAsFixedFormat(0, str(out_pdf.resolve()))
                    elif "ViKhiHau" in f.name:
                        for s in wb.Worksheets:
                            if s.Name not in active_set:
                                s.Visible = 0  # xlSheetHidden
                        wb.ExportAsFixedFormat(0, str(out_pdf.resolve()))
                        for s in wb.Worksheets:
                            s.Visible = -1  # xlSheetVisible
                    else:
                        wb.ExportAsFixedFormat(0, str(out_pdf.resolve()))
                    wb.Close(SaveChanges=True)
                finally:
                    wb = None
                    excel.Quit()
                    del excel
                    gc.collect()
                if out_pdf.exists():
                    exported.append(out_pdf)
        except Exception as exc:
            print(f"[WARN] Không xuất được PDF cho {f.name}: {exc}")
    return exported


if __name__ == "__main__":
    p_hkd = create_hoi_khi_doc_template()
    p_vkh = create_vikhihau_vatly_bui_template()
    p_eco = create_ecgonomy_template()
    p_tbkq, p_bckq = create_word_templates()
    db = build_complete_chemical_db()
    print(f"[OK] Đã tạo {p_hkd} (Tổng số dòng tra cứu Sheet1: {len(db)})")
    print(f"[OK] Đã tạo {p_vkh}")
    print(f"[OK] Đã tạo {p_eco}")
    print(f"[OK] Đã tạo {p_tbkq} & {p_bckq}")
