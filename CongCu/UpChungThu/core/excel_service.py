"""
Excel automation service using MS Excel COM interface (win32com).
Handles scanning workbook sheets and exporting selected/merged sheets to PDF.
"""

import os
import re
from typing import List, Tuple, Dict, Any, Optional, Callable
import pythoncom
import win32com.client


class ExcelEngine:
    """
    Service for MS Excel COM automation tasks.
    """

    @staticmethod
    def scan_files(files: List[str], input_folder: str, 
                   progress_callback: Optional[Callable[[str, int, int], None]] = None) -> Tuple[List[Tuple[str, str, List[str]]], List[str], Dict[str, int]]:
        """
        Scan a list of Excel files in input_folder to discover sheets and column widths.
        Returns:
            temp_data: list of (full_path, filename, sheet_names)
            all_unique_sheets: list of distinct sheet names in order of discovery
            col_widths: dict mapping sheet name to column pixel width
        """
        pythoncom.CoInitialize()
        excel = None
        temp_data = []
        all_unique_sheets = []
        col_widths = {}

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            total = len(files)
            for i, file in enumerate(files):
                if progress_callback:
                    progress_callback(file, i + 1, total)

                full_path = os.path.abspath(os.path.join(input_folder, file))
                wb = None
                try:
                    wb = excel.Workbooks.Open(full_path, ReadOnly=True, UpdateLinks=0)
                    sheet_names = [s.Name for s in wb.Sheets]
                    temp_data.append((full_path, file, sheet_names))
                    for s in sheet_names:
                        if s not in all_unique_sheets:
                            all_unique_sheets.append(s)
                            pixel_w = max(len(f"Tất cả {s}"), len(s)) * 8 + 40
                            col_widths[s] = max(110, int(pixel_w))
                except Exception as e:
                    print(f"[ExcelEngine] Error scanning {file}: {e}")
                finally:
                    if wb:
                        try:
                            wb.Close(False)
                        except Exception:
                            pass

            return temp_data, all_unique_sheets, col_widths
        finally:
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    @staticmethod
    def export_sheets_to_pdf(excel_app, workbook_path: str, sheets_to_merge: List[str], output_pdf_path: str):
        """
        Export specified sheets from a workbook into a single PDF document.
        """
        full_pdf_path = os.path.abspath(output_pdf_path)
        pdf_dir = os.path.dirname(full_pdf_path)
        if not os.path.exists(pdf_dir):
            os.makedirs(pdf_dir, exist_ok=True)

        wb = None
        try:
            wb = excel_app.Workbooks.Open(os.path.abspath(workbook_path), ReadOnly=True, UpdateLinks=0)
            wb.Worksheets(sheets_to_merge).Select()
            # 0 = xlTypePDF
            excel_app.ActiveSheet.ExportAsFixedFormat(0, full_pdf_path)
        finally:
            if wb:
                try:
                    wb.Close(False)
                except Exception:
                    pass
