# -*- coding: utf-8 -*-
"""
QUATEST 3 Environmental & Cleanroom Monitoring Framework (Room N2).
Provides modular architecture for high-precision, compliant report generation.
"""

from .Models import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation
from .Styling import (
    set_cell_content,
    set_cell_shading,
    set_cell_margins,
    set_paragraph_text,
    set_assessment_paragraph,
    set_table_autofit_window,
    set_table_row_height,
    format_table_structure,
    add_tbl_header,
    add_cant_split,
    strip_all_highlights,
    add_document_ending_line,
    update_headers,
    update_table0,
)
from .Builders import BCKQBuilder, TBKQBuilder
from .Builders.handlers import BSCHandler
from .Tools.bsc_calc import parse_bsc_excel, build_bsc_calculation_workbook
from .Tools.khinen_calc import (
    parse_compressed_air_workbook,
    build_compressed_air_workbook,
    create_dossier_from_compressed_air_workbooks,
)
from .Tools.scan_splitter import split_dossier_scans, split_all_dossiers_in_input
from .Exporters import PDFExporter, convert_docx_to_pdf
from .Services import DossierRunner

__all__ = [
    "Dossier",
    "ClientInfo",
    "Equipment",
    "PersonnelInfo",
    "MonitoringLocation",
    "set_cell_content",
    "set_cell_shading",
    "set_cell_margins",
    "set_paragraph_text",
    "set_assessment_paragraph",
    "set_table_autofit_window",
    "set_table_row_height",
    "format_table_structure",
    "add_tbl_header",
    "add_cant_split",
    "strip_all_highlights",
    "add_document_ending_line",
    "update_headers",
    "update_table0",
    "BCKQBuilder",
    "TBKQBuilder",
    "BSCHandler",
    "parse_bsc_excel",
    "build_bsc_calculation_workbook",
    "parse_compressed_air_workbook",
    "build_compressed_air_workbook",
    "create_dossier_from_compressed_air_workbooks",
    "split_dossier_scans",
    "split_all_dossiers_in_input",
    "PDFExporter",
    "convert_docx_to_pdf",
    "DossierRunner",
]
