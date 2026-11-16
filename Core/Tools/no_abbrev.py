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
