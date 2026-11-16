"""
Text Standardization, Personnel Resolution, and Equipment Rules for QUATEST 3 (Room N2).
Implements rules from:
- QUY-TAC-FONT-VA-THIET-BI.md (Phần 2, Phần 3, Phần 7)
- QUY-TRINH-XU-LY-HO-SO.md (Mục 2, Mục 4)
"""
import re
import sys
import unicodedata

sys.stdout.reconfigure(encoding="utf-8")

CENTRAL_CITIES = [
    "Hồ Chí Minh",
    "Hà Nội",
    "Đà Nẵng",
    "Hải Phòng",
    "Cần Thơ",
    "Huế",
]


def to_nfc(text: str) -> str:
    """Normalize Unicode string to NFC (fixes Bẫy 10 NFD decomposed Vietnamese accents)."""
    if text is None:
        return ""
    return unicodedata.normalize("NFC", text)


def normalize_address(addr: str) -> str:
    """
    Standardize Vietnamese address string according to QUY-TAC-FONT-VA-THIET-BI.md Phần 7:
    1. Expand abbreviations (KCN, KCX, KDC, CCN, KĐT, KKT, KCNC, QL, TL, P., Q., H., TX., TT., X., KP., TDP., TP.HCM...)
    2. Remove trailing 'Việt Nam' / 'Vietnam' / 'VN'
    3. Normalize capitalization of administrative units (Thành phố TW vs thành phố thuộc tỉnh, tỉnh, quận/Quận, phường/Phường, Khu phố, đường...)
    """
    if not addr:
        return ""
    s = to_nfc(addr).strip()

    # Collapse multiple spaces
    s = re.sub(r"\s+", " ", s)

    # 1. Expand industrial/urban zone abbreviations
    zone_map = [
        (r"\bKCN\b", "Khu công nghiệp"),
        (r"\bKCX\b", "Khu chế xuất"),
        (r"\bKDC\b", "Khu dân cư"),
        (r"\bCCN\b", "Cụm công nghiệp"),
        (r"\bKĐT\b", "Khu đô thị"),
        (r"\bKKT\b", "Khu kinh tế"),
        (r"\bKCNC\b", "Khu công nghệ cao"),
        (r"\bkhu công nghiệp\b", "Khu công nghiệp"),
        (r"\bkhu chế xuất\b", "Khu chế xuất"),
        (r"\bkhu dân cư\b", "Khu dân cư"),
        (r"\bcụm công nghiệp\b", "Cụm công nghiệp"),
        (r"\bkhu đô thị\b", "Khu đô thị"),
        (r"\bkhu kinh tế\b", "Khu kinh tế"),
        (r"\bkhu công nghệ cao\b", "Khu công nghệ cao"),
    ]
    for pat, repl in zone_map:
        s = re.sub(pat, repl, s, flags=re.IGNORECASE)

    # 2. Expand TP.HCM / TP. Hà Nội etc.
    s = re.sub(
        r"\b(?:TP|Tp)\.?\s*HCM\b",
        "Thành phố Hồ Chí Minh",
        s,
    )
    s = re.sub(
        r"\b(?:TP|Tp)\.\s*",
        "Thành phố ",
        s,
    )

    # 3. Expand roads & quarters
    s = re.sub(r"\bQL\.?\s*(?=\d|[A-Z])", "Quốc lộ ", s)
    s = re.sub(r"\bTL\.?\s*(?=\d|[A-Z])", "Tỉnh lộ ", s)
    s = re.sub(r"\b(?:Đ|Đg)\.\s*", "đường ", s)
    s = re.sub(r"\bKP\.?\s*(?=\d|\w)", "Khu phố ", s)
    s = re.sub(r"\bTDP\.?\s*(?=\d|\w)", "Tổ dân phố ", s)

    # 4. Expand P., Q., H., TX., TT., X.
    # P. <number> -> Phường <number>; P. <word> -> phường <word>
    s = re.sub(r"\bP\.\s*(\d+)\b", r"Phường \1", s)
    s = re.sub(r"\bP\.\s*([^\d,\s])", r"phường \1", s)
    s = re.sub(r"\bQ\.\s*(\d+)\b", r"Quận \1", s)
    s = re.sub(r"\bQ\.\s*([^\d,\s])", r"quận \1", s)
    s = re.sub(r"\bH\.\s*", "huyện ", s)
    s = re.sub(r"\bTX\.\s*", "thị xã ", s)
    s = re.sub(r"\bTT\.\s*", "thị trấn ", s)
    s = re.sub(r"\bX\.\s*", "xã ", s)

    # 5. Fix capitalization of Phường/Quận/Huyện/Xã/Thị trấn/Thị xã/Tỉnh/Đường/Khu phố when not at start of string
    # Phường + number -> Phường, Phường + word -> phường
    s = re.sub(r"(?<!^)\b[Pp]hường\s+(\d+)\b", r"Phường \1", s)
    s = re.sub(r"(?<!^)\b[Pp]hường\s+([^\d,\s])", r"phường \1", s)
    s = re.sub(r"(?<!^)\b[Qq]uận\s+(\d+)\b", r"Quận \1", s)
    s = re.sub(r"(?<!^)\b[Qq]uận\s+([^\d,\s])", r"quận \1", s)
    s = re.sub(r"(?<!^)\b[Hh]uyện\s+", "huyện ", s)
    s = re.sub(r"(?<!^)\b[Tt]hị\s+[Xx]ã\s+", "thị xã ", s)
    s = re.sub(r"(?<!^)\b[Tt]hị\s+[Tt]rấn\s+", "thị trấn ", s)
    s = re.sub(r"(?<!^)\b[Xx]ã\s+", "xã ", s)
    s = re.sub(r"(?<!^)\b[Tt]ỉnh\s+", "tỉnh ", s)
    s = re.sub(r"(?<!^)\b[Đđ]ường\s+", "đường ", s)
    s = re.sub(r"\b[Kk]hu\s+[Pp]hố\s+", "Khu phố ", s)
    s = re.sub(r"\b[Tt]ổ\s+[Dd]ân\s+[Pp]hố\s+", "Tổ dân phố ", s)
    s = re.sub(r"\b[Ấấ]p\s+", "Ấp ", s)
    s = re.sub(r"\b[Tt]hôn\s+", "Thôn ", s)

    # 6. Central cities vs Provincial cities
    # First lower-case all non-initial 'Thành phố' -> 'thành phố', then capitalize Central Cities
    s = re.sub(r"(?<!^)\b[Tt]hành\s+[Pp]hố\s+", "thành phố ", s)
    for cc in CENTRAL_CITIES:
        s = re.sub(rf"\b[Tt]hành\s+[Pp]hố\s+{re.escape(cc)}\b", f"Thành phố {cc}", s)

    # 7. Strip trailing ', Việt Nam' / ', Vietnam' / ', VN' and trailing punctuation
    s = re.sub(r"[,;\s]+(?:Việt\s*Nam|Viet\s*Nam|Vietnam|VN)\.?\s*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+,", ",", s)
    s = re.sub(r",\s*", ", ", s)
    return s.strip().rstrip(".")


def resolve_personnel(implementer_names: list[str] | str) -> dict:
    """
    Resolve BCKQ 'Người thực hiện', 'Người báo cáo', and TBKQ 'Giám định viên' (signer)
    from the field implementer(s) recorded on BBLM (QUY-TRINH-XU-LY-HO-SO.md Mục 2 & Bước 2):
    - Implementer on BCKQ: only last name(s), joined by ' + ' (no 'Nam' added unless Nam was on BBLM)
    - Reporter on BCKQ:
        * If done by Toàn or Nghĩa alone -> '<tên> + Minh'
        * Otherwise -> '<tên> + Minh + Nam' (without duplicating Minh or Nam)
    - Signer on TBKQ:
        * Toàn (alone) -> 'Nguyễn Khánh Toàn'
        * Nghĩa (alone) -> 'Phạm Quốc Nghĩa'
        * All other cases (Huy, Nam, multiple people, etc.) -> 'Nguyễn Xuân Nam'
    """
    if isinstance(implementer_names, str):
        # Split by +, comma, semicolon, slash, 'và'
        raw_parts = re.split(r"[,;/+]|(?:\s+và\s+)", implementer_names)
    else:
        raw_parts = list(implementer_names)

    last_names = []
    for p in raw_parts:
        p = to_nfc(p).strip()
        if not p:
            continue
        words = p.split()
        last = words[-1]
        if last not in last_names:
            last_names.append(last)

    impl_str = " + ".join(last_names)

    # Check if single implementer is Toàn or Nghĩa
    is_toan_only = len(last_names) == 1 and last_names[0].lower() == "toàn"
    is_nghia_only = len(last_names) == 1 and last_names[0].lower() == "nghĩa"

    reporter_list = list(last_names)
    if "Minh" not in reporter_list:
        reporter_list.append("Minh")
    if not (is_toan_only or is_nghia_only):
        if "Nam" not in reporter_list:
            reporter_list.append("Nam")
    reporter_str = " + ".join(reporter_list)

    if is_toan_only:
        tbkq_signer = "Nguyễn Khánh Toàn"
    elif is_nghia_only:
        tbkq_signer = "Phạm Quốc Nghĩa"
    else:
        tbkq_signer = "Nguyễn Xuân Nam"

    return {
        "last_names": last_names,
        "bckq_implementer": impl_str,
        "bckq_reporter": reporter_str,
        "tbkq_signer": tbkq_signer,
    }


if __name__ == "__main__":
    # Quick self-test
    sample_addr = "A87, KDC Thới An, Đ. Lê Thị Riêng, KP. 1, P. Thới An, Q. 12, TP.HCM, Việt Nam"
    print("Input :", sample_addr)
    print("Output:", normalize_address(sample_addr))
    for person in ["Nguyễn Khánh Toàn", "Phạm Quốc Nghĩa", "Huy", "Nguyễn Trọng Phát, Trần Minh"]:
        print(person, "->", resolve_personnel(person))
