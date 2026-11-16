"""Cách gọi thống nhất trong TBKQ / BCKQ / Excel (người dùng chốt 11/11/2026, hồ sơ 067460 — AGENTS mục 6).

- Gọi **"vị trí đo"** hoặc **"vị trí"**, không dùng **"điểm đo"** / "16 điểm" / "Số điểm đo"
  (EN: "location(s)", không dùng "point(s)").
- Chỉ tiêu độ sáng trong danh sách chỉ tiêu (bảng 9.1 phòng sạch, …) ghi **"Cường độ sáng trung bình"**,
  không ghi số lượng vị trí (EN "Average light intensity").

Không đụng tới thuật ngữ khác có chữ "điểm": "điểm sương", "Địa điểm", "thời điểm", "điểm tù", "điểm biên",
"điểm hiệu chuẩn".

    from Core.Tools.wording import normalize_points_vn, normalize_points_en, find_point_terms
"""
import re

NBSP = " "
_SP = r"[  ]+"

# "Cường độ sáng: 09 điểm;" / "Độ rọi: 9 vị trí" / "Cường độ sáng" -> "Cường độ sáng trung bình;"
_LUX_VN = re.compile(r"^(\s*-?\s*)(?:Cường độ sáng|Độ rọi|Độ sáng)(?:" + _SP + r"trung bình)?"
                     r"(?:\s*:\s*\d+" + _SP + r"(?:điểm|vị trí)(?:" + _SP + r"đo)?(?:\s*\([^)]*\))?)?\s*([;.,]?)\s*$",
                     re.IGNORECASE)
_LUX_EN = re.compile(r"^(\s*-?\s*)(?:Average" + _SP + r")?(?:Light intensity|Illuminance|Lighting level)"
                     r"(?:\s*:\s*\d+" + _SP + r"(?:points?|locations?)(?:\s*\([^)]*\))?)?\s*([;.,]?)\s*$",
                     re.IGNORECASE)

# "16 điểm" -> "16 vị trí" (cách không ngắt để Word không tách "vị / trí" trong ô hẹp)
_N_DIEM = re.compile(r"(\d+)" + _SP + r"điểm(?:" + _SP + r"đo)?(?![  ]*(?:sương|tù|biên|hiệu))")
_VT_DIEM = re.compile(r"([Vv])ị" + _SP + r"trí(?:" + _SP + r"các)?" + _SP + r"điểm" + _SP + r"đo")  # "vị trí các điểm đo"
_DIEM_DO = re.compile(r"(?<![Đđ]ịa )([Đđ])iểm" + _SP + r"đo")
_SO_DIEM = re.compile(r"\b([Ss])ố" + _SP + r"điểm(?![  ]*(?:sương|tù|biên|hiệu))")
_N_POINT = re.compile(r"\b(\d+)" + _SP + r"(?:measuring" + _SP + r"|measurement" + _SP + r")?points?\b", re.IGNORECASE)


def normalize_points_vn(text: str) -> str:
    """Một dòng / một câu tiếng Việt -> cách gọi thống nhất (vị trí đo, Cường độ sáng trung bình)."""
    if not text:
        return text
    m = _LUX_VN.match(text)
    if m:
        return f"{m.group(1)}Cường độ sáng trung bình{m.group(2)}"
    text = _VT_DIEM.sub(lambda m: ("C" if m.group(1) == "V" else "c") + "ác vị trí đo", text)
    text = _N_DIEM.sub(lambda m: f"{m.group(1)}{NBSP}vị{NBSP}trí", text)
    text = _DIEM_DO.sub(lambda m: ("V" if m.group(1) == "Đ" else "v") + "ị trí đo", text)
    text = _SO_DIEM.sub(lambda m: f"{m.group(1)}ố vị trí", text)
    return text


def normalize_points_en(text: str) -> str:
    """Bản tiếng Anh: "16 points" -> "16 locations", "Light intensity: 9 points" -> "Average light intensity"."""
    if not text:
        return text
    m = _LUX_EN.match(text)
    if m:
        return f"{m.group(1)}Average light intensity{m.group(2)}"
    return _N_POINT.sub(lambda m: f"{m.group(1)} location" + ("" if int(m.group(1)) == 1 else "s"), text)


_FIND_VN = re.compile(r"(?<![Đđ]ịa )[Đđ]iểm" + _SP + r"đo|\b\d+" + _SP + r"điểm(?![  ]*(?:sương|tù|biên|hiệu))"
                      r"|\b[Ss]ố" + _SP + r"điểm(?![  ]*(?:sương|tù|biên|hiệu))"
                      r"|(?:Cường độ sáng|Độ rọi)(?:" + _SP + r"trung bình)?\s*:\s*\d+" + _SP + r"(?:điểm|vị trí)")


def find_point_terms(text: str) -> list:
    """Các cụm sai cách gọi trong văn bản tiếng Việt (dùng cho qa_check final-check `[Point Term]`)."""
    return sorted({m.group(0).replace(NBSP, " ") for m in _FIND_VN.finditer(text or "")})


if __name__ == "__main__":
    for s in ("- Tốc độ dòng khí qua lọc HEPA: 16 điểm;", "Cường độ sáng: 09 điểm;", "Độ rọi: 9 vị trí.",
              "Số điểm đo", "Sơ đồ vị trí các điểm đo:", "Điểm đo", "Nhiệt độ điểm sương T1", "Địa điểm đo",
              "Nồng độ hạt bụi: 01 vị trí (2);", "Cường độ sáng bên trong tủ"):
        print(f"{s!r:45} -> {normalize_points_vn(s)!r}")
    for s in ("Downflow velocity: 16 points;", "Light intensity: 09 points;", "Noise level: 1 point;"):
        print(f"{s!r:45} -> {normalize_points_en(s)!r}")
    print(find_point_terms("Sơ đồ các điểm đo; 9 điểm; Nhiệt độ điểm sương; Địa điểm đo; Cường độ sáng: 9 vị trí"))
