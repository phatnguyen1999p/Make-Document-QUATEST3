# -*- coding: utf-8 -*-
"""
Hiệu chỉnh độ ồn theo ồn nền — dùng chung cho BSC, LAF, tủ cấy, tủ hút.

Nguồn: bảng "Correction chart for sound level readings" (dạng ảnh) trong quy trình
N2.TC.BSC.01, N2.TC.BSC.02, N2.TC.LAF.01; người dùng chốt áp dụng ngày 30/10/2026.
Chi tiết: QUY-TRINH-NGUONG-DANH-GIA.md mục 9.

    Chỉ hiệu chỉnh khi ồn nền > 55 dBA.  ΔN = ồn vận hành − ồn nền.
    ΔN (làm tròn về số nguyên)   Trừ vào ồn vận hành
    0 – 2                        không hiệu chỉnh — phải giảm ồn nền (cảnh báo, người dùng quyết)
    3                            3
    4 – 5                        2
    6 – 10                       1
    > 10                         0

Bảng dùng số nguyên, nên ΔN lẻ (vd 2,6; 5,4) được làm tròn nửa lên về số nguyên rồi mới tra
(2,6 → 3 → trừ 3; 5,4 → 5 → trừ 2; 5,5 → 6 → trừ 1). Không còn khe trống giữa các nhánh.
"""

import math
from typing import Any, Dict

AMBIENT_THRESHOLD_DBA = 55.0

STATUS_NOT_NEEDED = "khong_can"          # ồn nền ≤ 55 dBA
STATUS_CORRECTED = "da_hieu_chinh"        # tra bảng bình thường (kể cả ΔN > 10 → trừ 0)
STATUS_REDUCE_BACKGROUND = "giam_on_nen"  # ΔN ≤ 2 → quy trình yêu cầu giảm ồn nền, không tự hiệu chỉnh


def _round_half_up(x: float) -> int:
    return int(math.floor(x + 0.5)) if x >= 0 else -int(math.floor(-x + 0.5))


def background_noise_correction(ambient_dba: float, operating_dba: float,
                                threshold_dba: float = AMBIENT_THRESHOLD_DBA) -> Dict[str, Any]:
    """Trả về dict: delta_n, delta_n_int, correction, corrected_dba, status, warning (chuỗi rỗng nếu không có)."""
    amb = float(ambient_dba)
    op = float(operating_dba)
    delta_n = round(op - amb, 1)          # làm tròn 0,1 trước để tránh sai số dấu phẩy động (63,5 − 60,9)
    dn_int = _round_half_up(delta_n)
    warning = ""

    if amb <= threshold_dba:
        correction, status = 0.0, STATUS_NOT_NEEDED
    elif dn_int <= 2:
        correction, status = 0.0, STATUS_REDUCE_BACKGROUND
        warning = (f"Ồn nền {amb:.1f} dBA > {threshold_dba:.0f} dBA và ΔN = {delta_n:.1f} dB (≈ {dn_int}) ≤ 2 dB: "
                   "theo bảng hiệu chỉnh của quy trình phải giảm ồn nền, không tự hiệu chỉnh. "
                   "Đang giữ nguyên ồn vận hành — báo người dùng quyết.")
    elif dn_int == 3:
        correction, status = 3.0, STATUS_CORRECTED
    elif dn_int <= 5:
        correction, status = 2.0, STATUS_CORRECTED
    elif dn_int <= 10:
        correction, status = 1.0, STATUS_CORRECTED
    else:
        correction, status = 0.0, STATUS_CORRECTED

    return {
        "ambient": round(amb, 1),
        "operating": round(op, 1),
        "delta_n": delta_n,
        "delta_n_int": dn_int,
        "correction": correction,
        "corrected_dba": round(op - correction, 1),
        "status": status,
        "warning": warning,
    }


def noise_correction_value(ambient_dba: float, operating_dba: float) -> float:
    """Chỉ trả số dBA cần trừ (dùng khi ghi ô Excel); in cảnh báo nếu ΔN ≤ 2 dB."""
    res = background_noise_correction(ambient_dba, operating_dba)
    if res["warning"]:
        print(f"[WARNING] [Noise] {res['warning']}")
    return res["correction"]


def noise_correction_formula(ambient_ref: str, operating_ref: str,
                             threshold_dba: float = AMBIENT_THRESHOLD_DBA) -> str:
    """Công thức Excel cho ô hiệu chỉnh ồn nền (người dùng chốt 07/11/2026: ô tính ra là công thức, không số tĩnh):
    ồn nền ≤ 55 -> 0; ΔN làm tròn số nguyên (ROUND của Excel = nửa lên với số dương) ≤ 2 -> 0; 3 -> 3; 4–5 -> 2;
    6–10 -> 1; > 10 -> 0 — cùng bảng với background_noise_correction()."""
    dn = f"ROUND(ROUND({operating_ref}-{ambient_ref},1),0)"
    t = int(threshold_dba) if float(threshold_dba).is_integer() else threshold_dba
    return (f"=IF({ambient_ref}<={t},0,IF({dn}<=2,0,IF({dn}=3,3,IF({dn}<=5,2,IF({dn}<=10,1,0)))))")
