"""
Core module for UpChungThu application.
Contains backend business logic, document processing, COM automation, and configuration management.
"""

from .config_manager import ConfigManager
from .pdf_service import PdfService, detect_language
from .word_service import WordEngine
from .excel_service import ExcelEngine
from .mapping_service import MappingService
from .utils import remove_accents, calculate_target_date

__all__ = [
    "ConfigManager",
    "PdfService",
    "detect_language",
    "WordEngine",
    "ExcelEngine",
    "MappingService",
    "remove_accents",
    "calculate_target_date"
]
