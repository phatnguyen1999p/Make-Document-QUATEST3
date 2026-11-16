# -*- coding: utf-8 -*-
"""
Hulo Care LAF Document Generator strictly matching QUATEST 3 Cleanroom Standard Form
Refactored to use the central QUATEST 3 Core Framework.
Profile: Q326D126588 - Chi nhánh Công ty TNHH Hulo Care
Equipment: Tủ LAF (PM-026) tại Phòng Cân Nguyên Liệu
Standards: ISO 14644-1:2015, QUATEST 3 3030:2025, N2.TC.LAF.01
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')

from Core.run import create_sample_laf_dossier
from Core.Services import DossierRunner


def main():
    print("=== Executing Hulo Care LAF Generation via QUATEST 3 Core Framework ===")
    dossier = create_sample_laf_dossier()
    runner = DossierRunner(dossier)
    runner.run_all(export_pdf=True, run_qa=True)


if __name__ == "__main__":
    main()
