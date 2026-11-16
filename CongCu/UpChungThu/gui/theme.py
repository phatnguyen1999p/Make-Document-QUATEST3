"""
Theme definitions, color palettes, fonts, and ThemeManager for Tkinter UI.
"""

from typing import Dict, Any

MODERN_FONTS = {
    "font_title": ("Inter", 26, "bold"),
    "font_tree": ("Inter", 13),
    "font_header": ("Inter", 12, "bold"),
    "font_normal": ("Inter", 11),
    "font_small": ("Inter", 10)
}

LIGHT_THEME = {
    "bg_main": "#f8fafc",
    "bg_card": "#ffffff",
    "text_main": "#0f172a",
    "text_muted": "#64748b",
    "primary": "#2563eb",
    "success": "#10b981",
    "danger": "#ef4444",
    "warning": "#f59e0b",
    "accent": "#6366f1",
    "border": "#e2e8f0",
    **MODERN_FONTS
}

DARK_THEME = {
    "bg_main": "#0f172a",
    "bg_card": "#1e293b",
    "text_main": "#f8fafc",
    "text_muted": "#94a3b8",
    "primary": "#3b82f6",
    "success": "#10b981",
    "danger": "#f43f5e",
    "warning": "#fbbf24",
    "accent": "#818cf8",
    "border": "#334155",
    **MODERN_FONTS
}


class ThemeManager:
    """
    Centralized theme manager supporting light/dark theme switching and persistence.
    """

    def __init__(self, config_manager=None):
        self.config_manager = config_manager
        self.is_dark_mode = False
        if self.config_manager:
            self.is_dark_mode = self.config_manager.get_dark_mode()
        self.theme: Dict[str, Any] = DARK_THEME if self.is_dark_mode else LIGHT_THEME

    def load_settings(self):
        if self.config_manager:
            self.is_dark_mode = self.config_manager.get_dark_mode()
        self.theme = DARK_THEME if self.is_dark_mode else LIGHT_THEME

    def toggle(self):
        self.is_dark_mode = not self.is_dark_mode
        self.theme = DARK_THEME if self.is_dark_mode else LIGHT_THEME
        if self.config_manager:
            self.config_manager.save_dark_mode(self.is_dark_mode)

    def __getitem__(self, key: str) -> Any:
        return self.theme.get(key)


# Global Theme Manager instance
TM = ThemeManager()
