"""
Configuration manager for handling local and LAN shared settings,
persistence, migration, and default resets.
"""

import os
import sys
import json
from typing import Dict, Any, List, Optional


class ConfigManager:
    """
    Manages application settings stored in JSON files.
    Supports local storage and shared LAN network path configuration.
    """

    DEFAULT_TIME_RULES = {
        "t1": "09:00",
        "t2": "16:00",
        "x1": 0,
        "x2": 1,
        "x3": 2
    }

    DEFAULT_WORD_SETTINGS = {
        "run_date_update": True,
        "run_conversion": True,
        "add_bg_var": True,
        "use_time_rules": False,
        "use_manual_date": False,
        "manual_date_str": ""
    }

    DEFAULT_EXCEL_SETTINGS = {
        "add_bg_var": True
    }

    def __init__(self, app_dir: Optional[str] = None):
        if app_dir:
            self.app_dir = app_dir
        elif getattr(sys, 'frozen', False):
            self.app_dir = os.path.dirname(sys.executable)
        else:
            self.app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        self.bootstrap_file = os.path.join(self.app_dir, "settings_source.json")
        self.local_settings_file = os.path.join(self.app_dir, "app_settings.json")
        self.settings_file = self.local_settings_file

        self.use_lan = False
        self.lan_path = ""
        self.data: Dict[str, Any] = {}

        self.inspector_library: List[str] = []
        self.boss_library: List[str] = []
        self.mapping_paths: List[str] = []

        self.init_settings_path()
        self.load_settings()

    def init_settings_path(self):
        """Determine whether to use local or LAN shared settings file."""
        self.settings_file = self.local_settings_file
        if os.path.exists(self.bootstrap_file):
            try:
                with open(self.bootstrap_file, "r", encoding="utf-8") as f:
                    b_data = json.load(f)
                    self.use_lan = bool(b_data.get("use_lan", False))
                    self.lan_path = str(b_data.get("lan_path", ""))
                    if self.use_lan and self.lan_path and os.path.exists(self.lan_path):
                        self.settings_file = self.lan_path
            except Exception:
                pass

    def load_settings(self):
        """Load all settings from the active JSON settings file."""
        self.data = {}
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
            except Exception as e:
                print(f"[ConfigManager] Error reading {self.settings_file}: {e}")

        # Load Signers
        signer_data = self.data.get("signer_settings", {})
        insp = signer_data.get("inspector", [])
        if isinstance(insp, str):
            self.inspector_library = [n.strip() for n in insp.split(";") if n.strip()]
        else:
            self.inspector_library = list(insp) if isinstance(insp, list) else []

        boss = signer_data.get("boss", [])
        if isinstance(boss, str):
            self.boss_library = [n.strip() for n in boss.split(";") if n.strip()]
        else:
            self.boss_library = list(boss) if isinstance(boss, list) else []

        # Load Mapping Paths
        raw_paths = self.data.get("mapping_paths", [])
        if isinstance(raw_paths, list):
            self.mapping_paths = [p for p in raw_paths if os.path.isdir(p)]
        else:
            self.mapping_paths = []

        # Migration fallback from legacy mapping_path.txt if empty
        if not self.mapping_paths:
            txt_path = os.path.join(self.app_dir, "mapping_path.txt")
            if os.path.exists(txt_path):
                try:
                    with open(txt_path, "r", encoding="utf-8") as f:
                        self.mapping_paths = [line.strip() for line in f if line.strip() and os.path.isdir(line.strip())]
                    self.save_mapping_paths(self.mapping_paths)
                except Exception:
                    pass

    def save_settings(self):
        """Save all current settings to the active JSON settings file."""
        self.data["signer_settings"] = {
            "inspector": self.inspector_library,
            "boss": self.boss_library
        }
        self.data["mapping_paths"] = self.mapping_paths

        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"[ConfigManager] Error saving to {self.settings_file}: {e}")

    def save_bootstrap(self, use_lan: bool, lan_path: str):
        """Save bootstrap information (Local vs LAN)."""
        self.use_lan = use_lan
        self.lan_path = lan_path
        data = {
            "use_lan": self.use_lan,
            "lan_path": self.lan_path
        }
        try:
            with open(self.bootstrap_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self.init_settings_path()
            self.load_settings()
        except Exception as e:
            raise RuntimeError(f"Could not save bootstrap settings: {e}")

    # --- Signer Management ---
    def add_signer(self, name: str, role: str) -> bool:
        name = name.strip()
        if not name:
            return False
        if role == "gdv":
            if name not in self.inspector_library:
                self.inspector_library.append(name)
                self.save_settings()
                return True
        elif role == "boss":
            if name not in self.boss_library:
                self.boss_library.append(name)
                self.save_settings()
                return True
        return False

    def remove_signer(self, name: str, role: str) -> bool:
        if role == "gdv" and name in self.inspector_library:
            self.inspector_library.remove(name)
            self.save_settings()
            return True
        elif role == "boss" and name in self.boss_library:
            self.boss_library.remove(name)
            self.save_settings()
            return True
        return False

    # --- Mapping Paths Management ---
    def add_mapping_path(self, path: str) -> bool:
        norm_path = os.path.normpath(path)
        if norm_path not in self.mapping_paths and os.path.isdir(norm_path):
            self.mapping_paths.append(norm_path)
            self.save_settings()
            return True
        return False

    def save_mapping_paths(self, paths: List[str]):
        self.mapping_paths = [os.path.normpath(p) for p in paths if os.path.isdir(p)]
        self.save_settings()

    # --- Time Rules ---
    def get_time_rules(self) -> Dict[str, Any]:
        rules = self.data.get("time_rules", {})
        merged = dict(self.DEFAULT_TIME_RULES)
        merged.update(rules)
        return merged

    def save_time_rules(self, rules: Dict[str, Any]):
        self.data["time_rules"] = rules
        self.save_settings()

    # --- Word Settings ---
    def get_word_settings(self) -> Dict[str, Any]:
        ws = self.data.get("word_settings", {})
        merged = dict(self.DEFAULT_WORD_SETTINGS)
        merged.update(ws)
        return merged

    def save_word_settings(self, ws: Dict[str, Any]):
        self.data["word_settings"] = ws
        self.save_settings()

    # --- Excel Settings ---
    def get_excel_settings(self) -> Dict[str, Any]:
        es = self.data.get("excel_settings", {})
        merged = dict(self.DEFAULT_EXCEL_SETTINGS)
        merged.update(es)
        return merged

    def save_excel_settings(self, es: Dict[str, Any]):
        self.data["excel_settings"] = es
        self.save_settings()

    # --- Dark Mode ---
    def get_dark_mode(self) -> bool:
        return bool(self.data.get("dark_mode", False))

    def save_dark_mode(self, is_dark: bool):
        self.data["dark_mode"] = is_dark
        self.save_settings()

    # --- Reset All Settings ---
    def reset_all_settings(self):
        """Delete config files and reset memory state to default."""
        files_to_delete = [
            self.local_settings_file,
            os.path.join(self.app_dir, "time_rules.json"),
            os.path.join(self.app_dir, "mapping_path.txt"),
            self.bootstrap_file
        ]
        for f in files_to_delete:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception:
                    pass

        self.use_lan = False
        self.lan_path = ""
        self.settings_file = self.local_settings_file
        self.data = {}
        self.inspector_library = []
        self.boss_library = []
        self.mapping_paths = []
