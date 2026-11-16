# -*- coding: utf-8 -*-
"""
Compressed Air Document Generator strictly matching QUATEST 3 Standard Form (ISO 8573-1:2010).
Refactored to use the central QUATEST 3 Core Framework.
Profile: Q326D122266 - Công ty TNHH Nextern Việt Nam
Equipment: Đường ống khí nén phòng sản xuất & trạm máy nén
Standards: ISO 8573-1:2010, QUATEST 3 3030:2025
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')

from Core.run import create_sample_khinen_dossier
from Core.Services import DossierRunner


def main():
    print("=== Executing Nextern Compressed Air Generation via QUATEST 3 Core Framework ===")
    dossier = create_sample_khinen_dossier()
    runner = DossierRunner(dossier)
    runner.run_all(export_pdf=True, run_qa=True)


if __name__ == "__main__":
    main()
