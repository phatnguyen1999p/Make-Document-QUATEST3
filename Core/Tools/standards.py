# -*- coding: utf-8 -*-
"""
Trích dẫn đầy đủ tên tiêu chuẩn ở ô "Căn cứ" Table 0 TBKQ (người dùng chốt 01/11/2026):
"EN 12469:2000" -> "EN 12469:2000 - Biotechnology - Performance criteria for microbiological safety cabinets".
Tên giữ nguyên bản (tiếng Anh) ở cả TBKQ VN và EN. Mục đã có tên đầy đủ thì giữ nguyên.
Thêm tiêu chuẩn mới: thêm 1 dòng vào STANDARD_TITLES.
"""

import re
from typing import List

# (mẫu mã tiêu chuẩn — khớp trọn mục, tên nguyên bản)
STANDARD_TITLES = [
    (r"EN\s*12469(?::\s*\d{4})?",
     "Biotechnology - Performance criteria for microbiological safety cabinets"),
    (r"NSF\s*/\s*ANSI\s*49(?:[-:]\s*\d{4})?",
     "Biosafety Cabinetry: Design, Construction, Performance, and Field Certification"),
    (r"ISO\s*14644-1(?::\s*\d{4})?",
     "Cleanrooms and associated controlled environments — Part 1: Classification of air cleanliness by particle concentration"),
    (r"ISO\s*14644-3(?::\s*\d{4})?",
     "Cleanrooms and associated controlled environments — Part 3: Test methods"),
    (r"(?:ANSI\s*/\s*)?ASHRAE\s*(?:Standard\s*)?110(?:-\s*\d{4})?",
     "Methods of Testing Performance of Laboratory Fume Hoods"),
    (r"SEFA\s*1(?:-\s*\d{4})?",
     "Laboratory Fume Hoods - Recommended Practices"),
    (r"(?:CETA\s*)?CAG-003(?:[-:]\s*\d{4})?",
     "Certification Guide for Sterile Compounding Facilities"),
    (r"IEST-RP-CC002(?:\.\d+)?", "Unidirectional-Flow, Clean-Air Devices"),   # tủ cấy 139329/2 (10/11/2026)
    (r"(?:WHO\s*)?TRS\s*961(?:\s*\(\d{4}\))?\s*[-–]\s*(?:Annex|Phụ lục)\s*6",
     "WHO good manufacturing practices for sterile pharmaceutical products"),   # phòng sạch GMP 127215 (10/11/2026)
    (r"ISO\s*8573-1(?::\s*\d{4})?", "Compressed air — Part 1: Contaminants and purity classes"),
    (r"ISO\s*8573-2(?::\s*\d{4})?", "Compressed air — Part 2: Test methods for oil aerosol content"),
    (r"ISO\s*8573-3(?::\s*\d{4})?", "Compressed air — Part 3: Test methods for measurement of humidity"),
    (r"ISO\s*8573-4(?::\s*\d{4})?", "Compressed air — Part 4: Test methods for solid particle content"),
    (r"ISO\s*8573-7(?::\s*\d{4})?",
     "Compressed air — Part 7: Test method for viable microbiological contaminant content"),
]
_PATTERNS = [(re.compile(rf"^{p}$", re.IGNORECASE), t) for p, t in STANDARD_TITLES]


def standard_title(code: str) -> str:
    """Tên nguyên bản của mã tiêu chuẩn (khớp trọn chuỗi); không biết thì trả ''."""
    code = (code or "").strip()
    for pat, title in _PATTERNS:
        if pat.match(code):
            return title
    return ""


def _split_items(text: str) -> List[str]:
    items = []
    for part in re.split(r"[;\n]", text or ""):
        part = part.strip().lstrip("-").strip().rstrip(".;").strip()
        if not part:
            continue
        # "ASHRAE 110, SEFA 1-2010" -> 2 mục (chỉ tách khi mọi phần đều là mã tiêu chuẩn)
        sub = [s.strip() for s in part.split(",")]
        if len(sub) > 1 and all(standard_title(s) for s in sub):
            items.extend(sub)
        else:
            items.append(part)
    return items


def cite_standards(text: str) -> str:
    """Ô "Căn cứ": mỗi tiêu chuẩn kèm tên đầy đủ "<mã> - <tên>". Không có mã nào cần mở rộng -> trả nguyên văn.
    Từ 2 mục trở lên: mỗi mục 1 dòng "- ...;", dòng cuối kết thúc "." (như danh sách thiết bị, AGENTS mục 14)."""
    items = _split_items(text)
    if not any(standard_title(it) for it in items):
        return text
    items = [f"{it} - {standard_title(it)}" if standard_title(it) else it for it in items]
    if len(items) == 1:
        return items[0]
    lines = [f"- {it};" for it in items]
    lines[-1] = lines[-1][:-1] + "."
    return "\n".join(lines)
