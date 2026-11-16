# -*- coding: utf-8 -*-
"""
Command Line Runner for QUATEST 3 Monitoring Document Generation.
Usage:
    python Core/run.py --sample-laf
    python Core/run.py --sample-khinen
    python Core/run.py --profile profile.json
"""

import os
import sys
import json
import shutil
import re
import argparse

sys.stdout.reconfigure(encoding='utf-8')

# Ensure workspace root is in sys.path
WORKSPACE_SOURCE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.getcwd() if not os.path.isdir(os.path.join(WORKSPACE_SOURCE, "Input")) else WORKSPACE_SOURCE
# Network drive root for synced deliverables; override with env var QUATEST3_NETWORK_ROOT.
NETWORK_ROOT = os.environ.get("QUATEST3_NETWORK_ROOT", r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026")
if WORKSPACE_SOURCE not in sys.path:
    sys.path.insert(0, WORKSPACE_SOURCE)

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation
from Core.Models.dossier import inspection_scope
from Core.Services import DossierRunner
from Core.Tools.bsc_calc import parse_bsc_excel
from Core.Tools.fumehood_calc import parse_fumehood_excel
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)



def create_sample_fumehood_dossier() -> Dossier:
    """Creates a sample dossier for Fume Hood (Tủ hút phá mẫu BL.GHA-18 - Phượng Hải JSC)."""
    input_dir = os.path.join(WORKSPACE, "Input", "043915")
    os.makedirs(input_dir, exist_ok=True)

    client = ClientInfo(
        name_vn="Công ty Cổ phần Khoa học Kỹ thuật Phượng Hải",
        name_en="Phuong Hai Technology Science Joint Stock Company",
        address_vn="103/23 Nguyễn Hữu Dật, phường Tây Thạnh, Thành phố Hồ Chí Minh",
        address_en="103/23 Nguyen Huu Dat Street, Tay Thanh Ward, Ho Chi Minh City"
    )
    personnel = PersonnelInfo(
        executor="Phát",
        sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
        sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
        sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
        sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
    )
    equipments = [
        eq("kanomax_6501"),
        eq("testo_440"),
        eq("rion_nl53"),
        eq("ati_6d")
    ]
    loc = MonitoringLocation(
        name_vn="Nhà xưởng của Công ty Cổ phần Khoa học Kỹ thuật Phượng Hải, 23/5 Đặng Thị Tám, xã Bà Điểm, Thành phố Hồ Chí Minh",
        name_en="The factory of Phuong Hai Technology Science Joint Stock Company, 23/5 Dang Thi Tam street, Ba Diem commune, Ho Chi Minh City"
    )
    results_data = {
        "cabinet": {
            "cabinet_name": "Tủ hút phá mẫu",
            "brand": "Phượng Hải JSC",
            "model": "BL.GHA-18",
            "serial": "BLGHA260418-01",
            "device_id": "N/A"
        },
        "openings": [
            {
                "width_mm": 1530,
                "height_mm": 585,
                "opening_pct": 100,
                "avg": 0.52,
                "limit_vn": "≥ 0,5 m/s (*)",
                "limit_en": "≥ 0.5 m/s (*)",
                "assessment_vn": "Đạt",
                "assessment_en": "Pass",
                "readings": [
                    [0.52, 0.53, 0.51, 0.52, 0.55],
                    [0.51, 0.53, 0.52, 0.51, 0.51]
                ]
            },
            {
                "width_mm": 1530,
                "height_mm": 293,
                "opening_pct": 50,
                "avg": 0.53,
                "limit_vn": "≥ 0,5 m/s (*)",
                "limit_en": "≥ 0.5 m/s (*)",
                "assessment_vn": "Đạt",
                "assessment_en": "Pass",
                "readings": [
                    [0.55, 0.51, 0.50, 0.55, 0.52]
                ]
            },
            {
                "width_mm": 1530,
                "height_mm": 147,
                "opening_pct": 25,
                "avg": 0.53,
                "limit_vn": "≥ 0,5 m/s (*)",
                "limit_en": "≥ 0.5 m/s (*)",
                "assessment_vn": "Đạt",
                "assessment_en": "Pass",
                "readings": [
                    [0.53, 0.52, 0.56, 0.53, 0.53]
                ]
            }
        ],
        "lux": {
            "avg": 863.0,
            "avg_rounded": 863,
            "limit_vn": "≥ 861 lux (**)",
            "limit_en": "≥ 861 lux (**)",
            "assessment_vn": "Đạt",
            "assessment_en": "Pass"
        },
        "noise": {
            "ambient": 51.6,
            "operating": 70.7,
            "corrected_dba": 70.7,
            "limit_vn": "≤ 85,0 dBA (**)",
            "limit_en": "≤ 85.0 dBA (**)",
            "assessment_vn": "Đạt",
            "assessment_en": "Pass"
        },
        "airflow_pattern": {
            "internal": {
                "result_vn": "- Dòng khí đi vào vách sau, hướng vào các khe hút của tủ.\n- Dòng khí sau cửa không thoát ra ngoài.",
                "result_en": "- Air flow entered the rear wall, directed towards the cabinet's intake vents.\n- Air flow behind the door did not escape.",
                "limit_vn": "- Dòng khí đi vào vách sau, hướng vào các khe hút của tủ.\n- Dòng khí sau cửa không thoát ra ngoài. (**)",
                "limit_en": "- The airflow enters the back wall, directed towards the cabinet's intake vents.\n- The airflow behind the door does not escape. (**)",
                "assessment_vn": "Đạt",
                "assessment_en": "Pass"
            },
            "perimeter": {
                "result_vn": "- Dòng khí đi vào trong tủ, không đi ngược ra ngoài.",
                "result_en": "- The airflow went into the cabinet, not out.",
                "limit_vn": "- Dòng khí đi vào trong tủ, không đi ngược ra ngoài. (**)",
                "limit_en": "- The airflow goes into the cabinet, not out. (**)",
                "assessment_vn": "Đạt",
                "assessment_en": "Pass"
            }
        }
    }
    dossier = Dossier(
        code="Q326D043915",
        cert_date="21/04/2026",
        sample_date="18/04/2026",
        sample_time_vn="Từ ngày 18/04/2026",
        sample_time_en="From 18/04/2026",
        sub_no=None,
        is_multi_cert=False,
        dossier_type="fumehood",
        scope_vn='Giám định tủ hút phá mẫu',
        scope_en='Measuring/ validating "Fume hood for sample preparation"',
        location_desc_vn="Nhà xưởng của Công ty Cổ phần Khoa học Kỹ thuật Phượng Hải, 23/5 Đặng Thị Tám, xã Bà Điểm, Thành phố Hồ Chí Minh",
        location_desc_en="The factory of Phuong Hai Technology Science Joint Stock Company, 23/5 Dang Thi Tam street, Ba Diem commune, Ho Chi Minh City",
        procedure_vn="QUATEST 3 3030:2025, N2.TC.T.01",
        procedure_en="QUATEST 3 3030:2025, N2.TC.T.01",
        standards_vn="ASHRAE 110, SEFA 1-2010\nTài liệu kỹ thuật công bố bởi nhà sản xuất",
        standards_en="ASHRAE 110, SEFA 1-2010\nTechnical documentation published by the manufacturer",
        basis_vn="SEFA 1-2010, ASHRAE 110",
        basis_en="SEFA 1-2010, ASHRAE 110",
        client=client,
        personnel=personnel,
        equipments=equipments,
        locations=[loc],
        results_data=results_data,
        assessment={
            "items_list_vn": [
                "tốc độ dòng khí qua cửa làm việc",
                "cường độ sáng bên trong tủ",
                "độ ồn khi vận hành",
                "hình thái dòng khí của tủ"
            ],
            "pass_vn": "- Tốc độ dòng khí qua cửa làm việc;\n- Cường độ sáng bên trong tủ;\n- độ ồn khi vận hành;\n- Hình thái dòng khí của tủ;",
            "fail_vn": "-",
            "basis_vn": "Tốc độ gió: Tài liệu kỹ thuật công bố bởi nhà sản xuất\nCác chỉ tiêu còn lại: SEFA 1-2010"
        },
        input_dir=input_dir,
        tbkq_dir=os.path.join(input_dir, "TBKQ"),
        output_dir=""
    )
    return dossier


def create_sample_bsc_dossier() -> Dossier:
    """Creates sample dossier for Q326D131638 - ABT Biosafety Cabinet Class II."""
    input_dir = os.path.join(WORKSPACE, "Input", "131638")
    if not os.path.isdir(input_dir):
        input_dir = os.path.join(WORKSPACE, "Input")

    client = ClientInfo(
        name_vn="Công ty TNHH Giải pháp Y sinh ABT",
        name_en="ABT Biomedical Solutions Company Limited",
        address_vn="Nhà xưởng 5.02, Lô L2, đường Long Hậu - Hiệp Phước, Khu công nghiệp Long Hậu, xã Cần Giuộc, tỉnh Long An",
        address_en="Factory 5.02, Lot L2, Long Hau - Hiep Phuoc street, Long Hau Industrial Park, Can Giuoc commune, Long An province"
    )

    personnel = PersonnelInfo(
        executor="Phát",
        sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
        sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
        sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
        sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
    )

    equipments = [
        eq("tsi_8380"),
        eq("kanomax_6501"),
        eq("ati_6d_2i"),
        eq("smoke"),
        eq("testo_440"),
        eq("extech_sdl470"),
        eq("rion_nl53")
    ]

    loc = MonitoringLocation(
        name_vn="Tủ an toàn sinh học cấp 2 (ISA2 - 1000)",
        name_en="Biosafety Cabinet Class II (ISA2 - 1000)",
        points_summary_vn=[
            "Tốc độ gió qua cửa làm việc (Inflow)",
            "Tốc độ gió cấp qua lọc HEPA (Downflow)",
            "Cường độ sáng bên trong tủ",
            "Đo độ ồn khi vận hành",
            "Kiểm tra độ rò rỉ màng lọc HEPA",
            "Kiểm tra hình thái dòng khí",
            "Cường độ tia UV dải C"
        ]
    )

    dossier = Dossier(
        code="Q326D131638",
        cert_date="22/10/2026",
        sample_date="22/10/2026",
        reg_date="16/10/2026",
        sample_time_vn="Ngày 22/10/2026 từ 09:30 đến 11:30",
        sample_time_en="From 22/10/2026",
        sub_no="",
        is_multi_cert=False,
        dossier_type="bsc",
        scope_vn="Giám định tủ an toàn sinh học cấp 2",
        scope_en="Inspection of Biosafety Cabinet Class II",
        location_desc_vn="Nhà xưởng của Công ty TNHH Giải pháp Y sinh ABT, Nhà xưởng 5.02, Lô L2, đường Long Hậu - Hiệp Phước, Khu công nghiệp Long Hậu, xã Cần Giuộc, tỉnh Long An",
        location_desc_en="Factory of ABT Biomedical Solutions Company Limited, Factory 5.02, Lot L2, Long Hau - Hiep Phuoc street, Long Hau Industrial Park, Can Giuoc commune, Long An province",
        standards_vn="QUATEST 3 3030:2025, N2.TC.BSC.02; EN 12469:2000",
        standards_en="QUATEST 3 3030:2025, N2.TC.BSC.02; EN 12469:2000",
        basis_vn="EN 12469:2000",
        basis_en="EN 12469:2000",
        client=client,
        personnel=personnel,
        equipments=equipments,
        locations=[loc],
        results_data={
            "cabinet": {
                "cabinet_name": "Tủ an toàn sinh học cấp 2",
                "brand": "ABT",
                "model": "ISA2 - 1000",
                "serial": "26080302-01",
                "device_id": "N/A"
            },
            "inflow": {
                "method": 2,
                "method_name": "Kanomax Climomaster 6501",
                "location": "Mép trên cửa kính",
                "width_mm": 930,
                "height_mm": 200,
                "readings": [
                    [0.24, 0.20, 0.20, 0.20, 0.22],
                    [0.60, 0.54, 0.62, 0.56, 0.64]
                ],
                "velocity": 0.40
            },
            "downflow": {
                "readings": [
                    [0.28, 0.30, 0.28, 0.26, 0.27],
                    [0.29, 0.29, 0.27, 0.29, 0.28],
                    [0.28, 0.28, 0.30, 0.30, 0.31]
                ],
                "avg": 0.28,
                "tolerance_pct": 20
            },
            "lux": {
                "readings": [
                    [914.9, 1164.0, 869.8],
                    [962.4, 1221.5, 920.1],
                    [827.0, 1096.6, 750.3]
                ],
                "avg": 969.6,
                "avg_rounded": 970
            },
            "noise": {
                "ambient": 47.7,
                "operating": 60.7,
                "corrected_dba": 60.7
            },
            "hepa": {
                "downflow_pct": 0.0006,
                "exhaust_pct": 0.0000,
                "aerosol_conc": 15
            },
            "uvc": {
                "readings": [148, 204, 203, 161],
                "avg": 179.0,
                "avg_rounded": 179
            }
        },
        assessment={
            "items_list_vn": [
                "tốc độ dòng khí qua cửa làm việc (inflow)",
                "tốc độ dòng khí cấp qua lọc HEPA (downflow)",
                "cường độ sáng bên trong tủ",
                "độ ồn khi vận hành",
                "mức rò rỉ lọc HEPA",
                "hình thái dòng khí",
                "cường độ tia UV dải C"
            ],
            "pass_vn": "Tốc độ dòng khí Inflow, Downflow;\nCường độ sáng;\nĐộ ồn khi vận hành;\nMức rò rỉ lọc HEPA;\nHình thái dòng khí;\nCường độ tia UV dải C",
            "fail_vn": "-",
            "basis_vn": "EN 12469:2000"
        },
        input_dir=input_dir,
        tbkq_dir=os.path.join(input_dir, "TBKQ"),
        output_dir=""
    )
    return dossier


def create_132918_dossier() -> Dossier:
    """Creates dossier for Q326D132918 - Phòng Thử nghiệm Công nghệ Sinh học - QUATEST 3 (BSC Class II ESI FLUFRANCE BIOCYT)."""
    input_dir = os.path.join(WORKSPACE, "Input", "132918")

    client = ClientInfo(
        name_vn="Phòng Thử nghiệm Công nghệ Sinh học - Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3",
        name_en="Biotechnology Testing Laboratory - Quality Assurance and Testing Center 3",
        address_vn="Lô C5, đường D1, Khu công nghiệp Cát Lái, phường Cát Lái, Thành phố Hồ Chí Minh",
        address_en="Lot C5, D1 Street, Cat Lai Industrial Zone, Cat Lai Ward, Ho Chi Minh City"
    )

    personnel = PersonnelInfo(
        executor="Phát",
        sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
        sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
        sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
        sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
    )

    equipments = [
        eq("tsi_8380"),
        eq("kanomax_6501"),
        eq("ati_6d_2i"),
        eq("smoke"),
        eq("testo_440"),
        eq("rion_nl53"),
        eq("aerotrak_9350")
    ]

    loc = MonitoringLocation(
        name_vn="Tủ an toàn sinh học cấp 2 (Tủ cấy BIOCYT)",
        name_en="Biosafety Cabinet Class II (BIOCYT)",
        points_summary_vn=[
            "Tốc độ gió qua cửa làm việc (Inflow)",
            "Tốc độ gió cấp qua lọc HEPA (Downflow)",
            "Cường độ sáng bên trong tủ",
            "Đo độ ồn khi vận hành",
            "Kiểm tra độ rò rỉ màng lọc HEPA",
            "Kiểm tra hình thái dòng khí",
            "Nồng độ hạt bụi trong tủ"
        ]
    )

    dossier = Dossier(
        code="Q326D132918",
        cert_date="22/10/2026",
        sample_date="22/10/2026",
        reg_date="20/10/2026",
        sample_time_vn="Từ ngày 22/10/2026",
        sample_time_en="From 22/10/2026",
        sub_no="",
        is_multi_cert=False,
        dossier_type="bsc",
        scope_vn="Giám định tủ an toàn sinh học cấp 2.",
        scope_en="Inspection of Biosafety Cabinet Class II.",
        location_desc_vn="Phòng Thử nghiệm Công nghệ Sinh học - Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3 - Lô C5, đường D1, Khu công nghiệp Cát Lái, phường Cát Lái, Thành phố Hồ Chí Minh",
        location_desc_en="Biotechnology Testing Laboratory - Quality Assurance and Testing Center 3 - Lot C5, D1 Street, Cat Lai Industrial Zone, Cat Lai Ward, Ho Chi Minh City",
        standards_vn="QUATEST 3 3030:2025; EN 12469:2000; N2.TC.BSC.02",
        standards_en="QUATEST 3 3030:2025; EN 12469:2000; N2.TC.BSC.02",
        basis_vn="EN 12469:2000; ISO 14644-1:2015",
        basis_en="EN 12469:2000; ISO 14644-1:2015",
        client=client,
        personnel=personnel,
        equipments=equipments,
        locations=[loc],
        results_data={
            "start_time": "11:00",
            "end_time": "16:00",
            "cabinet": {
                "cabinet_name": "Tủ cấy BIOCYT",
                "brand": "ESI FLUFRANCE",
                "model": "ESI BIO18 PC UV PF",
                "serial": "39809392",
                "device_id": "VS0012"
            },
            "inflow": {
                "width_mm": 1800,
                "height_mm": 180,
                "q_m3h": 934.4,
                "q_readings": [932, 936, 932, 936, 936],
                "area_m2": 0.324,
                "velocity": 0.80
            },
            "downflow": {
                "readings": [
                    [0.57, 0.53, 0.48, 0.45, 0.42, 0.42, 0.43, 0.47, 0.47],
                    [0.56, 0.52, 0.49, 0.44, 0.42, 0.43, 0.40, 0.49, 0.48],
                    [0.54, 0.52, 0.46, 0.44, 0.44, 0.44, 0.42, 0.48, 0.46]
                ],
                "row_sums": [4.24, 4.23, 4.20],
                "total_sum": 12.67,
                "avg": 0.47,
                "tolerance_pct": 20
            },
            "lux": {
                "readings": [
                    [956.7, 1122.7, 1160.9, 1084.6, 951.5],
                    [956.1, 1104.7, 1102.6, 1091.0, 878.8],
                    [937.4, 849.6, 1030.7, 990.1, 867.8]
                ],
                "avg": 1005.7,
                "avg_rounded": 1006
            },
            "noise": {
                "ambient": 44.7,
                "operating": 65.9,
                "corrected_dba": 65.9,
                "file_no": "0165",
                "assessment": "Không đạt",
                "assessment_en": "Fail"
            },
            "hepa": {
                "downflow_pct": 0.2578,
                "exhaust_pct": 0.0021,
                "aerosol_conc": 12
            },
            "particle": {
                "clean_class": "ISO Class 5",
                "sample_volume_l": 1000.0,
                "sample_time_min": 20,
                "readings": [0, 0, 0, 0],
                "sizes": [
                    ("Kích thước hạt d ≥ 0,3 µm", "Particle size d ≥ 0.3 µm", "0 hạt/m³", "0 particles/m³", "≤ 10 200\nhạt/m³ ⁽²⁾", "≤ 10,200\nparticles/m³ ⁽²⁾", "Đạt", "Pass"),
                    ("Kích thước hạt d ≥ 0,5 µm", "Particle size d ≥ 0.5 µm", "0 hạt/m³", "0 particles/m³", "≤ 3 520\nhạt/m³ ⁽²⁾", "≤ 3,520\nparticles/m³ ⁽²⁾", "Đạt", "Pass"),
                    ("Kích thước hạt d ≥ 1,0 µm", "Particle size d ≥ 1.0 µm", "0 hạt/m³", "0 particles/m³", "≤ 832\nhạt/m³ ⁽²⁾", "≤ 832\nparticles/m³ ⁽²⁾", "Đạt", "Pass"),
                    ("Kích thước hạt d ≥ 5,0 µm", "Particle size d ≥ 5.0 µm", "0 hạt/m³", "0 particles/m³", "≤ 29\nhạt/m³ ⁽²⁾", "≤ 29\nparticles/m³ ⁽²⁾", "Đạt", "Pass"),
                ]
            }
        },
        assessment={
            "items_list_vn": [
                "tốc độ dòng khí qua cửa làm việc (inflow)",
                "tốc độ dòng khí cấp qua lọc HEPA (downflow)",
                "cường độ sáng bên trong tủ",
                "độ ồn khi vận hành",
                "mức rò rỉ lọc HEPA",
                "hình thái dòng khí",
                "nồng độ hạt bụi trong tủ"
            ],
            "pass_vn": "Tốc độ dòng khí Inflow, Downflow;\nCường độ sáng;\nMức rò rỉ lọc khí thải HEPA;\nHình thái dòng khí;\nNồng độ hạt bụi trong tủ",
            "fail_vn": "Độ ồn khi vận hành (65,9 dBA > 65 dBA);\nMức rò rỉ lọc khí cấp HEPA (0,2578% > 0,01%)",
            "summary_vn": "- Đánh giá: Tủ cấy BIOCYT đạt các chỉ tiêu: tốc độ dòng khí Inflow, tốc độ dòng khí Downflow, cường độ sáng, mức rò rỉ lọc khí thải HEPA, hình thái dòng khí, nồng độ hạt bụi trong tủ; Không đạt các chỉ tiêu: độ ồn khi vận hành (65,9 dBA > 65 dBA), mức rò rỉ lọc khí cấp HEPA (0,2578% > 0,01%) theo tiêu chuẩn EN 12469:2000.",
            "basis_vn": "EN 12469:2000; ISO 14644-1:2015"
        },
        input_dir=input_dir,
        tbkq_dir=os.path.join(input_dir, "TBKQ"),
        output_dir=""
    )
    return dossier


def create_sample_laf_dossier() -> Dossier:
    """Creates sample dossier for Q326D126588 - Hulo Care LAF Cabinet."""
    input_dir = os.path.join(WORKSPACE, "Input", "126588")
    photo_bui = os.path.join(input_dir, "Photo_DoBui.jpg")
    photo_gio = os.path.join(input_dir, "Photo_DoGio.jpg")
    layout_bui = os.path.join(input_dir, "Layout_Bui.png")
    layout_gio = os.path.join(input_dir, "Layout_Gio.png")

    client = ClientInfo(
        name_vn="Chi nhánh Công ty TNHH Hulo Care",
        name_en="Branch of Hulo Care Company Limited",
        address_vn="Nhà xưởng D4, Lô G6, G7, G8, G9, đường N3, Khu công nghiệp Đông Nam, xã Bình Mỹ, huyện Củ Chi, Thành phố Hồ Chí Minh",
        address_en="Factory D4, Lot G6, G7, G8, G9, N3 street, Dong Nam Industrial Park, Binh My commune, Cu Chi district, Ho Chi Minh city"
    )

    personnel = PersonnelInfo(
        executor="Phát",
        sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
        sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
        sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
        sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
    )

    equipments = [
        eq("aerotrak_9350"),
        eq("kanomax_6501"),
        eq("smoke"),
        eq("testo_440")
    ]

    location = MonitoringLocation(
        name_vn="Tủ LAF (PM-026)",
        name_en="LAF Cabinet (PM-026)",
        area_m2="0,72 m2",
        clean_class="ISO 5",
        state_vn="Trạng thái nghỉ",
        state_en="At rest",
        points_summary_vn=[
            "Nồng độ hạt bụi\t: 01 vị trí",
            "Tốc độ gió qua lọc HEPA\t: 06 vị trí",
            "Cường độ ánh sáng\t: 06 vị trí",
            "Hướng dòng không khí\t: 01 vị trí"
        ],
        points_summary_en=[
            "Particle concentration\t: 01 point",
            "Air velocity passing HEPA filter\t: 06 points",
            "Illuminance\t: 06 points",
            "Airflow pattern\t: 01 point"
        ]
    )

    dossier = Dossier(
        code="Q326D126588",
        cert_date="21/10/2026",
        sample_date="16/10/2026",
        sample_time_vn="Ngày 16/10/2026 từ 09:00 đến 10:30",
        sample_time_en="From 16/10/2026",
        sub_no=1,
        is_multi_cert=True,
        dossier_type="laf",
        scope_vn="Đo đạc và đánh giá hệ thống LAF tại khu vực cân nguyên liệu.",
        scope_en="Measurement and evaluation of LAF system at the raw material weighing area.",
        location_desc_vn="Phòng Cân Nguyên Liệu - Chi nhánh Công ty TNHH Hulo Care - Nhà xưởng D4, Lô G6, G7, G8, G9, đường N3, Khu công nghiệp Đông Nam, xã Bình Mỹ, huyện Củ Chi, Thành phố Hồ Chí Minh",
        location_desc_en="Raw Material Weighing Room - Branch of Hulo Care Company Limited - Factory D4, Lot G6, G7, G8, G9, N3 street, Dong Nam Industrial Park, Binh My commune, Cu Chi district, Ho Chi Minh city",
        standards_vn="QUATEST 3 3030:2025; N2.TC.LAF.01; ISO 14644-1:2015",
        standards_en="QUATEST 3 3030:2025; N2.TC.LAF.01; ISO 14644-1:2015",
        basis_vn=(
            "- ISO 14644-1:2015 Cleanrooms and associated controlled environments — Part 1: Classification of air cleanliness by particle concentration;\n"
            "- Mức quy định được cung cấp bởi đại diện Chi nhánh Công ty TNHH Hulo Care."
        ),
        basis_en=(
            "- ISO 14644-1:2015 Cleanrooms and associated controlled environments — Part 1: Classification of air cleanliness by particle concentration;\n"
            "- Permissible limits provided by the representative of Branch of Hulo Care Company Limited."
        ),
        client=client,
        personnel=personnel,
        equipments=equipments,
        locations=[location],
        photos=[
            {"path": photo_bui, "caption_vn": "Đo nồng độ hạt bụi", "caption_en": "Particle concentration measurement"},
            {"path": photo_gio, "caption_vn": "Đo tốc độ gió qua lọc HEPA", "caption_en": "Air velocity measurement through HEPA filter"}
        ],
        layouts=[
            {"path": layout_bui, "caption_vn": "Sơ đồ đo nồng độ hạt bụi", "caption_en": "Air particulate measuring layout"},
            {"path": layout_gio, "caption_vn": "Sơ đồ đo tốc độ gió & ánh sáng", "caption_en": "Air velocity & illuminance measuring layout"}
        ],
        results_data={
            "particle": {
                "p1": ["0", "0", "0", "0"],
                "limits": ["10 200", "3 520", "832", "29"]
            },
            "velocity": {
                "points": [("1", "0,40"), ("2", "0,33"), ("3", "0,49"), ("4", "0,36"), ("5", "0,34"), ("6", "0,45")],
                "avg": "0,40",
                "range": "[0,36 - 0,54]"
            },
            "illuminance": {
                "points": [("1", "613,1"), ("2", "619,7"), ("3", "620,8"), ("4", "703,8"), ("5", "761,7"), ("6", "803,0")],
                "avg": "687,0",
                "ambient": "11,0",
                "limit": "≥ 480 (Ánh sáng nền: ≤ 160)"
            }
        },
        assessment={
            "items_list_vn": [
                "nồng độ hạt bụi",
                "tốc độ gió qua lọc HEPA",
                "cường độ ánh sáng",
                "hướng dòng không khí"
            ],
            "pass_vn": "Nồng độ hạt bụi;\nTốc độ gió qua lọc HEPA;\nCường độ ánh sáng;\nHướng dòng không khí",
            "fail_vn": "-",
            "basis_vn": "Mức quy định nồng độ hạt bụi tối đa đối với phòng sạch cấp ISO 5 nêu tại ISO 14644-1:2015 và Mức quy định cung cấp bởi Chi nhánh Công ty TNHH Hulo Care"
        },
        input_dir=input_dir,
        tbkq_dir=os.path.join(input_dir, "TBKQ"),
        output_dir="",
        network_dir=""
    )

    return dossier


def create_sample_khinen_dossier() -> Dossier:
    """Creates sample dossier for Q326D122266 - Nextern Vietnam Compressed Air."""
    input_dir = os.path.join(WORKSPACE, "Input", "122266 - Khí nén")
    if not os.path.isdir(input_dir):
        input_dir = os.path.join(WORKSPACE, "Input")

    kn1_img = os.path.join(input_dir, "Vi tri khi nen", "KN1.png")
    kn2_img = os.path.join(input_dir, "Vi tri khi nen", "KN2.png")

    client = ClientInfo(
        name_vn="Công ty TNHH Nextern Việt Nam",
        name_en="Nextern Vietnam Limited Liability Company",
        address_vn="Đường số 7, Khu công nghiệp VSIP II-A, thị xã Tân Uyên, tỉnh Bình Dương",
        address_en="No. 7 street, VSIP II-A Industrial Park, Tan Uyen Town, Binh Duong Province"
    )

    personnel = PersonnelInfo(
        executor="Phát",
        sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
        sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
        sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
        sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
    )

    equipments = [
        eq("lasair_iii_110"),
        eq("cs_ds400")
    ]

    loc1 = MonitoringLocation(
        name_vn="Đường ống khí nén tại phòng sản xuất",
        name_en="Pneumatic pipeline in manufacturing room",
        points_summary_vn=["KN1"],
        points_summary_en=["KN1"]
    )
    loc2 = MonitoringLocation(
        name_vn="Đường ống khí tại trạm máy nén",
        name_en="Pneumatic pipeline at the compressor station",
        points_summary_vn=["KN2"],
        points_summary_en=["KN2"]
    )

    dossier = Dossier(
        code="Q326D122266",
        cert_date="20/10/2026",
        sample_date="06/10/2026",
        sample_time_vn="Ngày 06/10/2026 từ 08:00 đến 17:00",
        sample_time_en="From 06/10/2026",
        sub_no=2,
        is_multi_cert=True,
        dossier_type="compressed_air",
        scope_vn="Quan trắc chất lượng hệ thống khí nén theo tiêu chuẩn ISO 8573-1:2010.",
        scope_en="Monitoring compressed air quality according to ISO 8573-1:2010.",
        location_desc_vn="Công ty TNHH Nextern Việt Nam - VSIP II-A, Tân Uyên, Bình Dương",
        location_desc_en="Nextern Vietnam Limited Liability Company - VSIP II-A, Tan Uyen, Binh Duong",
        standards_vn="ISO 8573-1:2010; ISO 8573-2:2018; ISO 8573-7:2014; QUATEST 3 3030:2025",
        standards_en="ISO 8573-1:2010; ISO 8573-2:2018; ISO 8573-7:2014; QUATEST 3 3030:2025",
        basis_vn="ISO 8573-1:2010 và Yêu cầu kỹ thuật do Công ty TNHH Nextern Việt Nam cung cấp.",
        basis_en="ISO 8573-1:2010 and Technical requirements provided by Nextern Vietnam Limited Liability Company.",
        client=client,
        personnel=personnel,
        equipments=equipments,
        locations=[loc1, loc2],
        photos=[
            {"path": kn1_img, "caption_vn": "Vị trí lấy mẫu KN1", "caption_en": "Sampling location KN1"},
            {"path": kn2_img, "caption_vn": "Vị trí lấy mẫu KN2", "caption_en": "Sampling location KN2"}
        ],
        results_data={
            "oil": {
                "rows": [
                    {"loc": "KN1", "val": "< 0,01", "class": "Class 1"},
                    {"loc": "KN2", "val": "< 0,01", "class": "Class 1"}
                ]
            },
            "dewpoint": {
                "rows": [
                    {"loc": "KN1\n(10:41 – 11:01)", "values": ["22,6", "7,2", "-36,3", "-36,6", "160", "Class 3"]},
                    {"loc": "KN2\n(14:38 – 14:58)", "values": ["24,5", "7,2", "-37,2", "-37,5", "140", "Class 3"]}
                ]
            },
            "particles": {
                "rows": [
                    {"loc": "KN1", "values": ["0", "0", "0", "0", "Class 1"]},
                    {"loc": "KN2", "values": ["0", "0", "0", "0", "Class 1"]}
                ]
            }
        },
        assessment={
            "items_list_vn": ["hàm lượng tổng dầu", "nhiệt độ điểm sương", "nồng độ hạt bụi"],
            "pass_vn": "Hàm lượng tổng dầu;\nNhiệt độ điểm sương;\nNồng độ hạt bụi",
            "fail_vn": "-",
            "basis_vn": "ISO 8573-1:2010 và Yêu cầu kỹ thuật cung cấp bởi Công ty TNHH Nextern Việt Nam"
        },
        input_dir=input_dir,
        tbkq_dir=os.path.join(input_dir, "TBKQ"),
        output_dir=""
    )
    return dossier


def create_sample_khinen_133266_dossier() -> Dossier:
    """Creates sample dossier for Q326D133266 - Nextern Vietnam Compressed Air from TinhToanKhiNen.xlsx."""
    from Core.Tools.khinen_calc import create_dossier_from_compressed_air_workbooks
    source_files = [
        os.path.join(WORKSPACE, "TinhToan KQ Khi nen", "Q326D133266- KN1 - TinhToanKhiNen.xlsx"),
        os.path.join(WORKSPACE, "TinhToan KQ Khi nen", "Q326D133266- KN2 - TinhToanKhiNen.xlsx")
    ]
    target_dir = os.path.join(WORKSPACE, "Input", "133266")
    tbkq_dir = os.path.join(target_dir, "TBKQ")
    os.makedirs(tbkq_dir, exist_ok=True)
    tbkq_files = []
    for f in source_files:
        dest = os.path.join(tbkq_dir, os.path.basename(f))
        shutil.copy2(f, dest)
        tbkq_files.append(dest)

    dossier = create_dossier_from_compressed_air_workbooks(
        tbkq_files,
        code="Q326D133266",
        cert_date="20/10/2026",
        sample_date="06/10/2026",
        sample_time_vn="Ngày 06/10/2026 từ 08:00 đến 17:00",
        sample_time_en="From 06/10/2026"
    )
    dossier.input_dir = target_dir
    dossier.tbkq_dir = tbkq_dir
    dossier.output_dir = ""
    return dossier


def main():
    parser = argparse.ArgumentParser(description="QUATEST 3 Automated Document Generator")
    parser.add_argument("--sample-bsc", action="store_true", help="Run sample Biosafety Cabinet Class II dossier (Q326D131638)")
    parser.add_argument("--sample-fumehood", action="store_true", help="Run sample Fume Hood dossier (Q326D043915)")
    parser.add_argument("--fumehood-excel", type=str, help="Generate full Fume Hood dossier from an existing Excel calculation file")
    parser.add_argument("--132918", dest="dossier_132918", action="store_true", help="Run dossier Q326D132918 (BSC Class II BIOCYT)")
    parser.add_argument("--all-input", action="store_true", help="Run all dossiers in Input/ (131638 and 132918)")
    parser.add_argument("--sample-laf", action="store_true", help="Run sample LAF dossier (Q326D126588)")
    parser.add_argument("--sample-khinen", action="store_true", help="Run sample Compressed Air dossier (Q326D122266)")
    parser.add_argument("--sample-khinen-133266", action="store_true", help="Run sample Compressed Air dossier Q326D133266 (KN1 & KN2)")
    parser.add_argument("--khinen-excel", type=str, help="Generate full Compressed Air dossier from Excel calculation file(s) or folder")
    parser.add_argument("--bsc-excel", type=str, help="Generate full BSC dossier from an existing Excel calculation file")
    parser.add_argument("--intake", type=str, help="Process raw scan PDFs in folder via 100%% Offline Local OCR (0 credit) and generate profile.json")
    parser.add_argument("--auto", type=str, help="Auto-detect and run dossier from folder (profile.json or Excel calculation)")
    parser.add_argument("--profile", type=str, help="Path to profile JSON configuration")
    parser.add_argument("--khong-danh-gia", "--no-assessment", dest="no_assessment", action="store_true", help="Generate TBKQ without assessment (4 columns: Mức khuyến nghị, no Pass/Fail evaluation)")
    parser.add_argument("--both-forms", "--ca-hai-form", dest="both_forms", action="store_true", help="Generate both Có đánh giá and Không đánh giá forms simultaneously")
    parser.add_argument("--split-scan", type=str, help="Split raw scan PDFs in dossier folder into BBLM.pdf, Ghi Chep.pdf, Khac.pdf in TBKQ/Edoc (trang trắng bỏ)")
    parser.add_argument("--split-all-scans", action="store_true", help="Split raw scan PDFs across all dossier folders in Input/")
    parser.add_argument("--upchungthu", type=str, nargs="?", const="all", help="Apply UpChungThu official background and run error checking on TBKQ PDF files for dossier or 'all'")
    parser.add_argument("--force-ocr", action="store_true", help="Bỏ qua checkpoint, OCR lại file scan (mặc định: dùng lại OCR đã lưu nếu file scan không đổi)")
    parser.add_argument("--resume", action="store_true", help="Chạy tiếp: bỏ qua giai đoạn đã xong và không đổi dữ liệu đầu vào (xem --status)")
    parser.add_argument("--force-stage", type=str, help="Bắt buộc làm lại giai đoạn (cách nhau bằng dấu phẩy): ocr,intake,split_scan,bckq,tbkq,pdf,qa hoặc all")
    parser.add_argument("--status", type=str, help="Xem tiến độ các giai đoạn của hồ sơ (đường dẫn thư mục hồ sơ hoặc tên thư mục trong Input/)")
    parser.add_argument("--no-pdf", action="store_true", help="Skip PDF conversion")
    parser.add_argument("--no-qa", action="store_true", help="Skip QA verification")
    parser.add_argument("--init", type=str, nargs="?", const=".", help="Initialize standard dossier folders and AGENTS.md rules in target directory")

    args = parser.parse_args()

    # Checkpoint / tiến độ
    from Core.Services.checkpoint import CheckpointStore, parse_force_stages
    if args.status is not None:
        cand = os.path.abspath(args.status)
        if not os.path.isdir(cand):
            cand = os.path.join(WORKSPACE_SOURCE, "Input", args.status)
        if not os.path.isdir(cand):
            print(f"Error: Không tìm thấy hồ sơ: {args.status}")
            return
        print(CheckpointStore(dossier_dir=cand).format_status())
        return
    try:
        force_stages = parse_force_stages(args.force_stage, args.force_ocr)
    except ValueError as _err:
        print(f"Error: {_err}")
        return

    if args.upchungthu is not None:
        from Core.Services.upchungthu_bridge import process_tbkq_pdf, is_tbkq_file
        print("\n========================================================")
        print("  UPCHUNGTHU INTEGRATION: BACKGROUND INSERTION & QA CHECK")
        print("========================================================")

        target = args.upchungthu
        dossier_targets = []
        if target == "all":
            input_dir = os.path.join(WORKSPACE_SOURCE, "Input")
            for item in sorted(os.listdir(input_dir)):
                dpath = os.path.join(input_dir, item)
                if os.path.isdir(dpath) and os.path.isdir(os.path.join(dpath, "pdf")):
                    dossier_targets.append(dpath)
        else:
            cand = os.path.abspath(target)
            if not os.path.isdir(cand):
                cand = os.path.join(WORKSPACE_SOURCE, "Input", target)
            if os.path.isdir(cand):
                dossier_targets.append(cand)
            else:
                print(f"Error: Target not found: {target}")
                return

        for dpath in dossier_targets:
            pdf_dir = os.path.join(dpath, "pdf")
            dname = os.path.basename(dpath)
            print(f"\n[Dossier: {dname}] Processing TBKQ in {pdf_dir}...")
            tbkq_found = 0
            for f in sorted(os.listdir(pdf_dir)):
                if is_tbkq_file(f) and f.lower().endswith(".pdf"):
                    tbkq_found += 1
                    pdf_f = os.path.join(pdf_dir, f)
                    process_tbkq_pdf(pdf_f, add_bg=True, print_log=True)
            if tbkq_found == 0:
                print("  No TBKQ PDF files found in folder.")
        return

    if args.init is not None:
        target = os.path.abspath(args.init)
        for sub in ["ThucHien", "TBKQ", os.path.join("TBKQ", "Edoc"), "pdf", "Photo"]:
            os.makedirs(os.path.join(target, sub), exist_ok=True)
        rule_candidates = [
            os.path.join(WORKSPACE_SOURCE, "AGENTS.md"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "Rule", "CLAUDE.md"),
        ]
        target_rule = os.path.join(target, "AGENTS.md")
        for rc in rule_candidates:
            if os.path.isfile(rc) and not os.path.isfile(target_rule):
                try:
                    shutil.copy2(rc, target_rule)
                    break
                except Exception as _err:
                    print(f"[Warning] run.py:886 {type(_err).__name__}: {_err}")
        print(f"[Init] Initialized standard dossier folders and rules at: {target}")
        return

    if args.intake:
        from Core.Tools.intake_helper import process_scan_intake
        process_scan_intake(args.intake, force_ocr=args.force_ocr)
        return
    elif args.auto:
        target_dir = os.path.abspath(args.auto)
        p_json = os.path.join(target_dir, "profile.json")
        if os.path.isfile(p_json):
            args.profile = p_json
        else:
            thuchien_dir = os.path.join(target_dir, "ThucHien")
            tbkq_dir = os.path.join(target_dir, "TBKQ")
            candidate_dirs = [d for d in [thuchien_dir, tbkq_dir, target_dir] if os.path.isdir(d)]
            found_excel = False
            for sdir in candidate_dirs:
                if found_excel:
                    break
                for f in sorted(os.listdir(sdir)):
                    if f.endswith(".xlsx") and not f.startswith("~$"):
                        if "TinhToanBSC" in f:
                            args.bsc_excel = os.path.join(sdir, f)
                            found_excel = True
                            break
                        elif "TinhToanFumehood" in f:
                            args.fumehood_excel = os.path.join(sdir, f)
                            found_excel = True
                            break
                        elif "TinhToanKhiNen" in f or "KN" in f:
                            args.khinen_excel = sdir
                            found_excel = True
                            break
            if not found_excel:
                print(f"[Info] Running intake on {target_dir} first...")
                from Core.Tools.intake_helper import process_scan_intake
                res = process_scan_intake(target_dir, force_ocr=args.force_ocr)
                args.profile = res["profile_path"]

    if args.split_all_scans:
        from Core.Tools.scan_splitter import split_all_dossiers_in_input
        split_all_dossiers_in_input(os.path.join(WORKSPACE, "Input"))
        return
    elif args.split_scan:
        from Core.Tools.scan_splitter import split_dossier_scans
        split_dossier_scans(args.split_scan, force_ocr=args.force_ocr)
        return
    elif args.sample_fumehood:
        dossier = create_sample_fumehood_dossier()
    elif args.fumehood_excel:
        parsed_data = parse_fumehood_excel(args.fumehood_excel)
        code_str = parsed_data.get("code") or "Q326D043915"
        cab = parsed_data.get("cabinet", {})
        sample_d = parsed_data.get("sample_date") or "21/04/2026"
        excel_dir = os.path.dirname(os.path.abspath(args.fumehood_excel))
        if os.path.basename(excel_dir).lower() in ("thuchien", "tbkq"):
            dossier_base_dir = os.path.dirname(excel_dir)
        else:
            dossier_base_dir = excel_dir

        c_name_vn = parsed_data.get("client_name_vn") or "Công ty Cổ phần Khoa học Kỹ thuật Phượng Hải"
        c_name_en = parsed_data.get("client_name_en") or "Phuong Hai Technology Science Joint Stock Company"
        c_addr_vn = parsed_data.get("client_address_vn") or "103/23 Nguyễn Hữu Dật, phường Tây Thạnh, Thành phố Hồ Chí Minh"
        c_addr_en = parsed_data.get("client_address_en") or "103/23 Nguyen Huu Dat Street, Tay Thanh Ward, Ho Chi Minh City"
        loc_desc_vn = parsed_data.get("location_vn") or f"Nhà xưởng của {c_name_vn}, 23/5 Đặng Thị Tám, xã Bà Điểm, Thành phố Hồ Chí Minh"
        loc_desc_en = parsed_data.get("location_en") or f"The factory of {c_name_en}, 23/5 Dang Thi Tam street, Ba Diem commune, Ho Chi Minh City"

        client = ClientInfo(
            name_vn=c_name_vn,
            name_en=c_name_en,
            address_vn=c_addr_vn,
            address_en=c_addr_en
        )
        personnel = PersonnelInfo(
            executor="Phát",
            sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
            sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
            sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
            sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
        )
        equipments = [
            eq("kanomax_6501"),
            eq("testo_440"),
            eq("rion_nl53"),
            eq("ati_6d")
        ]
        loc = MonitoringLocation(
            name_vn=f"{cab.get('cabinet_name', 'Tủ hút khí độc')} ({cab.get('model', '')})",
            name_en=f"Fume Hood ({cab.get('model', '')})"
        )
        dossier = Dossier(
            code=code_str,
            cert_date=sample_d,
            sample_date=sample_d,
            sample_time_vn=f"Từ ngày {sample_d}",
            sample_time_en=f"From {sample_d}",
            sub_no=None,
            is_multi_cert=False,
            dossier_type="fumehood",
            scope_vn=inspection_scope("Giám định " + cab.get("cabinet_name", "Tủ hút khí độc")),
            scope_en=f'Measuring/ validating "{cab.get("cabinet_name", "Fume Hood")}"',
            location_desc_vn=loc_desc_vn,
            location_desc_en=loc_desc_en,
            procedure_vn="QUATEST 3 3030:2025, N2.TC.T.01",
            procedure_en="QUATEST 3 3030:2025, N2.TC.T.01",
            standards_vn="ASHRAE 110, SEFA 1-2010\nTài liệu kỹ thuật công bố bởi nhà sản xuất",
            standards_en="ASHRAE 110, SEFA 1-2010\nTechnical documentation published by the manufacturer",
            basis_vn="SEFA 1-2010, ASHRAE 110",
            basis_en="SEFA 1-2010, ASHRAE 110",
            client=client,
            personnel=personnel,
            equipments=equipments,
            locations=[loc],
            results_data=parsed_data,
            input_dir=dossier_base_dir,
            thuchien_dir=os.path.join(dossier_base_dir, "ThucHien"),
            tbkq_dir=os.path.join(dossier_base_dir, "TBKQ"),
            edoc_dir=os.path.join(dossier_base_dir, "TBKQ", "Edoc"),
            pdf_dir=os.path.join(dossier_base_dir, "pdf"),
            output_dir=""
        )
    elif args.all_input:
        dossiers = [create_sample_fumehood_dossier(), create_sample_bsc_dossier(), create_132918_dossier()]
        for d in dossiers:
            if args.both_forms:
                d.both_forms = True
                d.has_assessment = True
            elif args.no_assessment:
                d.has_assessment = False
                d.both_forms = False
            runner = DossierRunner(d)
            runner.run_all(export_pdf=not args.no_pdf, run_qa=not args.no_qa, resume=args.resume, force_stages=force_stages)
        return
    elif args.sample_bsc:
        dossier = create_sample_bsc_dossier()
    elif args.dossier_132918:
        dossier = create_132918_dossier()
    elif args.bsc_excel:
        parsed_data = parse_bsc_excel(args.bsc_excel)
        meta = parsed_data.get("metadata", {})
        cab = parsed_data.get("cabinet", {})
        code_str = meta.get("code") or "Q326D131638"
        sample_d = meta.get("sample_date") or meta.get("reg_date") or "22/10/2026"
        sub = meta.get("sub_no") or ""

        # Default standard client & equipments for BSC
        b_name_vn = meta.get("client_name_vn") or "Công ty TNHH Giải pháp Y sinh ABT"
        b_name_en = meta.get("client_name_en") or "ABT Biomedical Solutions Company Limited"
        b_addr_vn = meta.get("client_address_vn") or "Nhà xưởng 5.02, Lô L2, đường Long Hậu - Hiệp Phước, Khu công nghiệp Long Hậu, xã Cần Giuộc, tỉnh Long An"
        b_addr_en = meta.get("client_address_en") or "Factory 5.02, Lot L2, Long Hau - Hiep Phuoc street, Long Hau Industrial Park, Can Giuoc commune, Long An province"
        b_loc_vn = meta.get("location_vn") or f"Nhà xưởng của {b_name_vn}, {b_addr_vn}"
        b_loc_en = meta.get("location_en") or f"Factory of {b_name_en}, {b_addr_en}"

        client = ClientInfo(
            name_vn=b_name_vn,
            name_en=b_name_en,
            address_vn=b_addr_vn,
            address_en=b_addr_en
        )
        personnel = PersonnelInfo(
            executor="Phát",
            sign_gdv_vn="GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam",
            sign_gdv_en="INSPECTOR\n\n\n\nNguyen Xuan Nam",
            sign_leader_vn="KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền",
            sign_leader_en="FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
        )
        equipments = [
            eq("tsi_8380"),
            eq("kanomax_6501"),
            eq("ati_6d_2i"),
            eq("smoke"),
            eq("testo_440"),
            eq("rion_nl53")
        ]
        if meta.get("uvc"):
            equipments.append(eq("extech_sdl470"))
        loc = MonitoringLocation(
            name_vn=f"{cab.get('cabinet_name', 'Tủ an toàn sinh học cấp 2')} ({cab.get('model', '')})",
            name_en=f"Biosafety Cabinet Class II ({cab.get('model', '')})"
        )
        excel_dir = os.path.dirname(os.path.abspath(args.bsc_excel))
        if os.path.basename(excel_dir).lower() in ("thuchien", "tbkq"):
            dossier_base_dir = os.path.dirname(excel_dir)
        else:
            dossier_base_dir = excel_dir

        dossier = Dossier(
            code=code_str,
            cert_date=sample_d,
            sample_date=sample_d,
            sample_time_vn=f"Ngày {sample_d} từ {meta.get('start_time', '09:30')} đến {meta.get('end_time', '11:30')}",
            sample_time_en=f"From {sample_d}",
            sub_no=sub,
            is_multi_cert=bool(sub),
            dossier_type="bsc",
            scope_vn="Giám định tủ an toàn sinh học cấp 2",
            scope_en="Inspection of Biosafety Cabinet Class II",
            location_desc_vn=b_loc_vn,
            location_desc_en=b_loc_en,
            standards_vn="QUATEST 3 3030:2025, N2.TC.BSC.02; EN 12469:2000",
            standards_en="QUATEST 3 3030:2025, N2.TC.BSC.02; EN 12469:2000",
            basis_vn="EN 12469:2000",
            basis_en="EN 12469:2000",
            client=client,
            personnel=personnel,
            equipments=equipments,
            locations=[loc],
            results_data=parsed_data,
            input_dir=dossier_base_dir,
            thuchien_dir=os.path.join(dossier_base_dir, "ThucHien"),
            tbkq_dir=os.path.join(dossier_base_dir, "TBKQ"),
            edoc_dir=os.path.join(dossier_base_dir, "TBKQ", "Edoc"),
            pdf_dir=os.path.join(dossier_base_dir, "pdf"),
            output_dir=""
        )
    elif args.sample_laf:
        dossier = create_sample_laf_dossier()
    elif args.sample_khinen:
        dossier = create_sample_khinen_dossier()
    elif args.sample_khinen_133266:
        dossier = create_sample_khinen_133266_dossier()
    elif args.khinen_excel:
        from Core.Tools.khinen_calc import create_dossier_from_compressed_air_workbooks
        target_path = os.path.abspath(args.khinen_excel)
        xlsx_files = []
        if os.path.isdir(target_path):
            tbkq_sub = os.path.join(target_path, "TBKQ")
            search_dir = tbkq_sub if os.path.isdir(tbkq_sub) else target_path
            for f in sorted(os.listdir(search_dir)):
                if f.endswith(".xlsx") and ("TinhToanKhiNen" in f or "KN" in f):
                    xlsx_files.append(os.path.join(search_dir, f))
            if not xlsx_files:
                for f in sorted(os.listdir(target_path)):
                    if f.endswith(".xlsx"):
                        xlsx_files.append(os.path.join(target_path, f))
        elif os.path.isfile(target_path):
            parent_dir = os.path.dirname(target_path)
            base_f = os.path.basename(target_path)
            m_code = re.search(r"(Q\d+D\d+|\d{6})", base_f)
            if m_code:
                code_pattern = m_code.group(1)
                for f in sorted(os.listdir(parent_dir)):
                    if f.endswith(".xlsx") and code_pattern in f and ("TinhToanKhiNen" in f or "KN" in f):
                        xlsx_files.append(os.path.join(parent_dir, f))
            if not xlsx_files:
                xlsx_files = [target_path]

        if not xlsx_files:
            print(f"[Error] No compressed air calculation Excel files found in {args.khinen_excel}")
            return

        first_fname = os.path.basename(xlsx_files[0])
        m_c = re.search(r"(Q\d+D\d+)", first_fname)
        code_str = m_c.group(1) if m_c else "Q326D133266"

        dossier = create_dossier_from_compressed_air_workbooks(xlsx_files, code=code_str)
        base_dir = os.path.dirname(os.path.abspath(xlsx_files[0]))
        if os.path.basename(base_dir).lower() in ("thuchien", "tbkq"):
            dossier.input_dir = os.path.dirname(base_dir)
        else:
            dossier.input_dir = base_dir
        dossier.thuchien_dir = os.path.join(dossier.input_dir, "ThucHien")
        dossier.tbkq_dir = os.path.join(dossier.input_dir, "TBKQ")
        dossier.edoc_dir = os.path.join(dossier.tbkq_dir, "Edoc")
        dossier.pdf_dir = os.path.join(dossier.input_dir, "pdf")
        dossier.output_dir = ""
    elif args.profile:
        with open(args.profile, "r", encoding="utf-8") as f:
            data = json.load(f)
        client = ClientInfo(**data.get("client", {}))
        personnel = PersonnelInfo(**data.get("personnel", {}))
        equipments = [Equipment(**eq) for eq in data.get("equipments", [])]
        locations = [MonitoringLocation(**loc) for loc in data.get("locations", [])]

        dossier = Dossier(
            code=data["code"],
            cert_date=data["cert_date"],
            sample_date=data.get("sample_date", ""),
            sample_time_vn=data.get("sample_time_vn", ""),
            sample_time_en=data.get("sample_time_en", ""),
            sub_no=data.get("sub_no", None),
            is_multi_cert=data.get("is_multi_cert", False),
            dossier_type=data.get("dossier_type", "cleanroom"),
            scope_vn=data.get("scope_vn", ""),
            scope_en=data.get("scope_en", ""),
            location_desc_vn=data.get("location_desc_vn", ""),
            location_desc_en=data.get("location_desc_en", ""),
            standards_vn=data.get("standards_vn", ""),
            standards_en=data.get("standards_en", ""),
            basis_vn=data.get("basis_vn", ""),
            basis_en=data.get("basis_en", ""),
            client=client,
            personnel=personnel,
            equipments=equipments,
            locations=locations,
            photos=data.get("photos", []),
            layouts=data.get("layouts", []),
            results_data=data.get("results_data", {}),
            assessment=data.get("assessment", {}),
            reg_date=data.get("reg_date", ""),
            input_dir=data.get("input_dir", os.path.dirname(os.path.abspath(args.profile))),
            tbkq_dir=data.get("tbkq_dir", os.path.join(os.path.dirname(os.path.abspath(args.profile)), "TBKQ")),
            output_dir=data.get("output_dir", ""),
            network_dir=data.get("network_dir", "")
        )
    else:
        parser.print_help()
        return

    if args.both_forms:
        dossier.both_forms = True
        dossier.has_assessment = True
    elif args.no_assessment:
        dossier.has_assessment = False
        dossier.both_forms = False

    runner = DossierRunner(dossier)
    runner.run_all(export_pdf=not args.no_pdf, run_qa=not args.no_qa, resume=args.resume, force_stages=force_stages)


if __name__ == "__main__":
    main()
