# -*- coding: utf-8 -*-
"""
Không ghi gọn (người dùng chốt 06/11/2026, hồ sơ 136247 — AGENTS mục 15):

- Tên đơn vị, địa chỉ, địa điểm giám định ghi ĐỦ như BBLM / PYC: không viết tắt (NGK, Cty, CP, PTN, QLCL,
  KCN, TP.HCM, BV, VN…) và không lược bỏ phần BBLM có ghi (vd BBLM "Phòng Vi sinh - Công ty Cổ phần Phân bón
  Bình Điền - <địa chỉ>" thì không được rút thành "Phòng Vi sinh - <địa chỉ>").
- "TNHH" là tên loại hình doanh nghiệp in trên GCN ĐKDN -> giữ nguyên (không phải ghi gọn).

Dùng:
- ``expand_abbrev(text)``      -> chuỗi đã mở rộng viết tắt (``Dossier.__post_init__`` tự gọi cho tên KH,
                                  địa chỉ, địa điểm tiếng Việt).
- ``normalize_address_en(text)`` -> địa chỉ tiếng Anh viết thường street / ward / province / city, trừ
                                  Ho Chi Minh City, Ha Noi City, Da Nang City, Can Tho City, Dong Nai City
                                  (người dùng chốt 20/11/2026; ``Dossier.__post_init__`` tự gọi).
- ``find_abbrev(text)``        -> danh sách viết tắt còn sót (``qa_check`` ``[Abbreviation]``).
- ``location_missing_org(loc, is_en)`` -> True khi địa điểm không nêu tên đơn vị nào
                                  (``qa_check`` ``[Location Short]``).
"""
import re
import unicodedata

# (mẫu, thay bằng, nhãn hiển thị) — chỉ áp cho văn bản tiếng Việt
ABBREV_VN = [
    (r"\bC[Tt][Yy]\b\.?", "Công ty", "Cty"),
    (r"\bNGK\b", "Nước giải khát", "NGK"),
    (r"(Công ty)\s+CP\b", r"\1 Cổ phần", "CP"),
    (r"\bCN\.?\s+(?=Công ty)", "Chi nhánh ", "CN"),
    (r"\bPTN\b", "Phòng thí nghiệm", "PTN"),
    (r"\bQLCL\b", "Quản lý chất lượng", "QLCL"),
    (r"\bKCNC\b", "Khu công nghệ cao", "KCNC"),
    (r"\bKCN\b", "Khu công nghiệp", "KCN"),
    (r"\bKCX\b", "Khu chế xuất", "KCX"),
    (r"\bCCN\b", "Cụm công nghiệp", "CCN"),
    (r"\bKDC\b", "Khu dân cư", "KDC"),
    (r"\bBV\b\.?", "Bệnh viện", "BV"),
    (r"\b(?:TP|Tp)\.?\s*HCM\b", "Thành phố Hồ Chí Minh", "TP.HCM"),
    (r"\bHCM\b", "Hồ Chí Minh", "HCM"),
    (r"\bVN\b", "Việt Nam", "VN"),
]

_ORG_VN = re.compile(r"(Công ty|Chi nhánh|Tổng công ty|Tập đoàn|Bệnh viện|Phòng khám|Trung tâm|Viện|Trường|"
                     r"Nhà máy|Xí nghiệp|Hợp tác xã|Doanh nghiệp|Cơ sở|Văn phòng|Cục|Sở|Ngân hàng|QUATEST)",
                     re.IGNORECASE)
_ORG_EN = re.compile(r"(Company|Branch|Limited|Corporation|Corp\.|Co\.|Group|Hospital|Clinic|Center|Centre|"
                     r"Institute|University|School|Factory|Plant|Enterprise|Bank|QUATEST|\bof\b|\bat\b)",
                     re.IGNORECASE)


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s or "")


def expand_abbrev(text: str) -> str:
    """Mở rộng viết tắt trong văn bản tiếng Việt (tên đơn vị, địa chỉ, địa điểm). Giữ 'TNHH'."""
    s = _nfc(text)
    for pat, repl, _ in ABBREV_VN:
        s = re.sub(pat, repl, s)
    return re.sub(r"[ \t]{2,}", " ", s)


# Địa chỉ tiếng Anh (người dùng chốt 20/11/2026, AGENTS mục 14): viết thường street, ward, province (và city);
# riêng 5 thành phố sau giữ "City" viết hoa.
CITY_KEEP_EN = ["Ho Chi Minh", "Ha Noi", "Da Nang", "Can Tho", "Dong Nai"]
_LOWER_EN = re.compile(r"\b(Street|Ward|Province|City)\b", re.IGNORECASE)
_CITY_KEEP_RE = re.compile(r"\b(" + "|".join(re.escape(c) for c in CITY_KEEP_EN) + r")\s+city\b", re.IGNORECASE)


def normalize_address_en(text: str) -> str:
    """Viết thường 'street', 'ward', 'province', 'city' trong địa chỉ / địa điểm tiếng Anh, trừ
    Ho Chi Minh City, Ha Noi City, Da Nang City, Can Tho City, Dong Nai City (giữ 'City').
    Vd 'No. 7 Street, Tay Thanh Ward, Binh Duong Province' -> 'No. 7 street, Tay Thanh ward, Binh Duong province'."""
    s = _nfc(text)
    if not s:
        return s
    s = _LOWER_EN.sub(lambda m: m.group(1).lower(), s)
    return _CITY_KEEP_RE.sub(lambda m: m.group(1) + " City", s)


def find_address_case_en(text: str) -> list:
    """Các từ street / ward / province / city viết sai hoa-thường trong địa chỉ tiếng Anh."""
    s = _nfc(text)
    return [] if normalize_address_en(s) == s else sorted({m.group(1) for m in _LOWER_EN.finditer(s)
                                                            if m.group(1) != m.group(1).lower()})


def find_abbrev(text: str) -> list:
    """Các viết tắt còn sót trong văn bản tiếng Việt (nhãn, không trùng)."""
    s = _nfc(text)
    return list(dict.fromkeys(lbl for pat, _, lbl in ABBREV_VN if re.search(pat, s)))


def location_missing_org(location: str, is_en: bool = False) -> bool:
    """Địa điểm giám định không nêu tên đơn vị nào (chỉ có tên phòng + địa chỉ) -> nghi ghi gọn,
    đối chiếu BBLM mục 5 (Địa điểm thực hiện)."""
    s = _nfc(location).strip()
    if not s:
        return False
    return not (_ORG_EN if is_en else _ORG_VN).search(s)


if __name__ == "__main__":
    for t in ["Phòng thí nghiệm tại Chi nhánh Cty TNHH NGK Swire Coca-Cola VN tại Tây Ninh - KCN Phú An Thạnh",
              "Phòng Vi sinh - Khu công nghiệp Long Định - Long Cang, xã Long Cang, tỉnh Tây Ninh",
              "Phòng Vi sinh - Công ty Cổ phần Phân bón Bình Điền - Khu công nghiệp Long Định - Long Cang"]:
        print(find_abbrev(t), location_missing_org(t), "->", expand_abbrev(t))
    for t in ["103/23 Nguyen Huu Dat Street, Tay Thanh Ward, Ho Chi Minh City",
              "No. 7 Street, VSIP II-A Industrial Park, Tan Uyen Ward, Binh Duong Province",
              "Lot 5, Long Duc Commune, Dong Nai city", "Thu Duc City, ha noi city"]:
        print(normalize_address_en(t))
