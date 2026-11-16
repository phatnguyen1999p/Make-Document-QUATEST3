"""
Directory mapping and indexing service.
Scans root archive directories, matches record folders (6-digit patterns),
and resolves subfolder structures (e.g. TBKQ).
"""

import os
import re
import time
from typing import List, Dict, Tuple


class MappingService:
    """
    Service for folder mapping, record searching, and TBKQ subfolder discovery.
    """

    @staticmethod
    def scan_base_directories(base_dirs: List[str]) -> Tuple[Dict[str, str], int]:
        """
        Scan base directories for subdirectories matching 6-digit record patterns.
        Returns:
            (mapping_dict, elapsed_milliseconds)
        """
        mapping: Dict[str, str] = {}
        if not base_dirs:
            return mapping, 0

        start_time = time.perf_counter()
        for base_dir in base_dirs:
            if os.path.exists(base_dir):
                try:
                    for name in os.listdir(base_dir):
                        path = os.path.join(base_dir, name)
                        if os.path.isdir(path):
                            if re.search(r'\d{6}', name):
                                mapping[name] = path
                except Exception as e:
                    print(f"[MappingService] Error reading {base_dir}: {e}")

        elapsed_ms = int((time.perf_counter() - start_time) * 1000)
        return mapping, elapsed_ms

    @staticmethod
    def get_tbkq_subfolder(parent_path: str) -> str:
        """
        Search for a subfolder containing 'tbkq' (case-insensitive) inside parent_path.
        Returns the TBKQ subfolder if found, otherwise returns parent_path.
        """
        try:
            for item in os.listdir(parent_path):
                item_path = os.path.join(parent_path, item)
                if os.path.isdir(item_path) and "tbkq" in item.lower():
                    return item_path
        except Exception:
            pass
        return parent_path

    @staticmethod
    def search_records(mapping: Dict[str, str], query: str) -> List[str]:
        """
        Filter mapping keys matching query string case-insensitively.
        """
        clean_q = query.strip().lower()
        if not clean_q:
            return []
        return [k for k in mapping.keys() if clean_q in k.lower()]
