"""
GUI package for UpChungThu application.
Contains view screens, dialogs, components, themes, and main application controller.
"""

from .theme import TM, ThemeManager, LIGHT_THEME, DARK_THEME, MODERN_FONTS
from .app import UnifiedDocumentSystem

__all__ = [
    "TM",
    "ThemeManager",
    "LIGHT_THEME",
    "DARK_THEME",
    "MODERN_FONTS",
    "UnifiedDocumentSystem"
]
