"""
QUATEST3 - HỆ THỐNG XỬ LÝ VĂN BẢN (Word, Excel, PDF)
Backwards-compatible entry point and legacy module wrapper.

For new development, please use:
- Entry point: main.py
- Backend logic: core/
- Frontend GUI: gui/
"""

import sys
import tkinter as tk
from gui.app import UnifiedDocumentSystem
from gui.theme import TM, ThemeManager, LIGHT_THEME, DARK_THEME, MODERN_FONTS
from gui.views.main_menu import MainMenu
from gui.views.word_view import WordProcessorView, WordProcessor
from gui.views.excel_view import ExcelProcessorView, ExcelProcessor
from gui.components.base_processor import MappingProcessor
from core.pdf_service import PdfService, detect_language
from core.word_service import WordEngine
from core.excel_service import ExcelEngine
from core.mapping_service import MappingService
from core.config_manager import ConfigManager
from core.utils import remove_accents, calculate_target_date


def main():
    root = tk.Tk()
    app = UnifiedDocumentSystem(root)
    root.mainloop()


if __name__ == "__main__":
    main()