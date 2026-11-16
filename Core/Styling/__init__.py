# -*- coding: utf-8 -*-
"""Word & OpenXML Styling and Formatting Package for QUATEST 3 Documents."""

from .styles import W_NS, set_cell_content, set_cell_shading, set_cell_margins
from .tables import (
    add_tbl_header,
    add_cant_split,
    set_table_row_height,
    set_table_col_widths,
    set_table_autofit_window,
    format_table_structure,
    format_bckq_table0,
    PHOTO_WIDTH_2COL,
    PHOTO_WIDTH_3COL,
    strip_image_crops,
    insert_image_to_cell,
    merge_stt_column_child_rows,
    get_bsc_row_heights,
    get_fumehood_row_heights,
)
from .paragraphs import set_paragraph_text, set_assessment_paragraph
from .elements import (
    strip_all_highlights,
    add_document_ending_line,
    drop_trailing_empty_paragraphs,
    drop_section_gaps,
    is_blank_paragraph,
    update_headers,
    update_table0,
    populate_cleanroom_table0,
    fix_unit_symbols,
)

__all__ = [
    "W_NS",
    "set_cell_content",
    "set_cell_shading",
    "set_cell_margins",
    "add_tbl_header",
    "add_cant_split",
    "set_table_row_height",
    "set_table_col_widths",
    "set_table_autofit_window",
    "format_table_structure",
    "format_bckq_table0",
    "PHOTO_WIDTH_2COL",
    "PHOTO_WIDTH_3COL",
    "strip_image_crops",
    "fix_unit_symbols",
    "insert_image_to_cell",
    "merge_stt_column_child_rows",
    "get_bsc_row_heights",
    "get_fumehood_row_heights",
    "set_paragraph_text",
    "set_assessment_paragraph",
    "strip_all_highlights",
    "add_document_ending_line",
    "drop_trailing_empty_paragraphs",
    "drop_section_gaps",
    "is_blank_paragraph",
    "update_headers",
    "update_table0",
    "populate_cleanroom_table0",
]
