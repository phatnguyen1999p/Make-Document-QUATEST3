# -*- coding: utf-8 -*-
"""
Client Inspection Certificate Builder (TBKQ - VN / EN) for QUATEST 3.
Orchestrates specialized handlers (Cleanroom, Compressed Air, LAF/BSC) to generate
high-fidelity, compliant client deliverables.
"""

from typing import Tuple
from ..Models.dossier import Dossier
from .handlers import (
    LAFBSCHandler,
    BSCHandler,
    FumeHoodHandler,
    CleanroomHandler,
    CompressedAirHandler,
    MTLDHandler,
)


class TBKQBuilder:
    """Orchestrating builder for TBKQ documents."""

    def __init__(self, dossier: Dossier):
        self.dossier = dossier
        dtype = dossier.dossier_type.lower()
        if dtype in ("fumehood", "fume_hood", "tu hut", "tu_hut", "tủ hút"):
            self.handler = FumeHoodHandler(dossier)
        elif dtype in ("bsc", "biological_safety_cabinet", "biosafety_cabinet"):
            self.handler = BSCHandler(dossier)
        elif dtype in ("laf", "laf_bsc"):
            self.handler = LAFBSCHandler(dossier)
        elif dtype in ("cleanroom", "phong sach"):
            self.handler = CleanroomHandler(dossier)
        elif dtype in ("compressed_air", "khi nen"):
            self.handler = CompressedAirHandler(dossier)
        elif dtype in ("mtld", "qtmtld", "moi truong lao dong", "môi trường lao động"):
            self.handler = MTLDHandler(dossier)
        else:
            # Default to Cleanroom handler if unknown
            self.handler = CleanroomHandler(dossier)

    def build_vn(self) -> str:
        """Builds Vietnamese TBKQ document."""
        return self.handler.build_tbkq(is_en=False, has_assessment=self.dossier.has_assessment)

    def build_en(self) -> str:
        """Builds English TBKQ document."""
        return self.handler.build_tbkq(is_en=True, has_assessment=self.dossier.has_assessment)

    def build_all(self) -> Tuple[str, ...]:
        """Builds all configured documents (VN and EN, and both evaluation modes if requested)."""
        if getattr(self.dossier, "both_forms", False) and hasattr(self.handler, "build_all_modes"):
            modes = self.handler.build_all_modes()
            return modes["vn_eval"], modes["en_eval"], modes["vn_no_eval"], modes["en_no_eval"]
        else:
            path_vn = self.build_vn()
            path_en = self.build_en()
            return path_vn, path_en
