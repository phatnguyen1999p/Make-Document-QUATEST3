# -*- coding: utf-8 -*-
"""Danh sách thông số nước thải (tên, đơn vị) dùng cho bản khung chưa có KQTN — phương pháp để "………" (lấy theo KQTN
khi về). Khóa: tên viết tắt trong Phụ lục I của biên bản hiện trường (Excel `… - Biên bản hiện trường.xlsx`)."""
from nuoc_common import PLACEHOLDER as P

UNITS = {
    "Tổng chất rắn lơ lửng (TSS)": "mg/L", "BOD₅": "mg/L", "COD": "mg/L", "Florua (F⁻)": "mg/L",
    "Amoni (NH₄⁺ tính theo N)": "mg/L", "Tổng nitơ": "mg/L", "Tổng phốt pho": "mg/L", "Xyanua (CN⁻)": "mg/L",
    "Thủy ngân (Hg)": "mg/L", "Đồng (Cu)": "mg/L", "Sắt (Fe)": "mg/L", "Kẽm (Zn)": "mg/L", "Chì (Pb)": "mg/L",
    "Asen (As)": "mg/L", "Niken (Ni)": "mg/L", "Tổng crôm (Cr)": "mg/L", "Chất hoạt động bề mặt anion": "mg/L",
    "Polyclo biphenyl (PCB)": "mg/L", "Tổng phenol": "mg/L", "Tổng Coliform": "MPN/ 100 mL", "Độ màu": "Pt-Co",
    "Cadimi (Cd)": "mg/L", "Mangan (Mn)": "mg/L", "Crôm VI (Cr⁶⁺)": "mg/L", "Dầu mỡ động thực vật": "mg/L",
    "Crôm III (Cr³⁺)": "mg/L", "Tổng hoạt độ phóng xạ α": "Bq/L", "Tổng hoạt độ phóng xạ β": "Bq/L",
    "Clorua (Cl⁻)": "mg/L", "Hóa chất BVTV phospho hữu cơ": "mg/L", "Sulfua (S²⁻)": "mg/L",
    "Tổng dầu mỡ khoáng": "mg/L",
}


def params(names):
    """[(tên, đơn vị, phương pháp '………')] theo thứ tự ``names``."""
    return [(n, UNITS[n], P) for n in names]
