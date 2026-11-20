# -*- coding: utf-8 -*-
"""Danh mục tên chuẩn của máy đo — MỘT tên VN + MỘT tên EN cho mỗi máy (AGENTS mục 14).

Mọi nơi ghi tên thiết bị (TBKQ VN/EN, BCKQ, tem, Excel, build_<6 số>.py, run.py, intake_helper)
lấy tên từ đây. Thêm máy mới / đổi cách gọi → sửa DUY NHẤT bảng ``EQUIPMENT`` bên dưới.

- ``eq("kanomax_6501")``       → ``Equipment`` với tên chuẩn VN/EN.
- ``canonical_name(text, lang)`` → tên chuẩn nếu ``text`` gọi đúng một máy trong danh mục, không thì giữ nguyên.
- ``Equipment.__post_init__`` tự gọi ``normalize_equipment`` → tên cũ trong script hồ sơ cũ vẫn ra tên chuẩn.
- ``nonstandard_lines(text)``  → dòng gọi máy trong danh mục bằng tên khác chuẩn (``qa_check`` ``[Equipment Name]``).
"""
from __future__ import annotations

import re
import unicodedata

# key: (tên VN, tên EN, regex nhận diện trên chữ thường). Thứ tự quan trọng: biến thể có phụ kiện đứng trước.
EQUIPMENT: dict[str, tuple[str, str, str]] = {
    # --- Tủ (BSC, LAF, tủ cấy, tủ hút) ---
    "kanomax_6501": (
        "Thiết bị đo tốc độ gió Kanomax Climomaster 6501, Japan",
        "Air velocity meter Kanomax Climomaster 6501, Japan",
        r"kanomax|climomaster|\b6501\b",
    ),
    "tsi_8380": (
        "Thiết bị đo lưu lượng TSI 8380, USA",
        "Air volume meter TSI 8380, USA",
        r"\b8380\b",
    ),
    "tsi_8375": (
        "Thiết bị đo lưu lượng và chênh áp TSI 8375, USA",
        "Air volume and differential pressure meter TSI 8375, USA",
        r"\b8375\b",
    ),
    "ati_6d_2i": (
        "Thiết bị tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA",
        "Aerosol generator ATI 6D and aerosol photometer ATI 2i, USA",
        r"photometer|quang kế|\b2i\b",
    ),
    "ati_6d": (  # chỉ máy tạo sol khí (tủ hút: thử khói), không có quang kế
        "Thiết bị tạo sol khí ATI aerosol generator, model 6D, USA",
        "Aerosol generator ATI 6D, USA",
        r"^(?!.*(?:photometer|quang kế|\b2i\b)).*\bati\b",
    ),
    "smoke": (
        "Thiết bị tạo khói khí dung (Smoke pencil / Aerosol generator)",
        "Smoke generator (Smoke pencil / Aerosol generator)",
        r"smoke|tạo khói",
    ),
    "testo_440": (
        "Thiết bị đo độ rọi Testo 440, Germany",
        "Digital luxmeter Testo 440, Germany",
        r"testo\s*(?:440|540)(?!.*(?:co2|tvoc))",
    ),
    "rion_nl53": (
        "Thiết bị đo độ ồn Rion NL-53, Japan",
        "Sound level meter Rion NL-53, Japan",
        r"nl[\s-]*53",
    ),
    "extech_sdl470": (
        "Thiết bị đo cường độ bức xạ UVC Extech SDL470, USA",
        "UVC light meter Extech SDL470, USA",
        r"extech|sdl\s*470",
    ),
    # --- Đếm hạt (tủ, phòng sạch, khí nén) ---
    "aerotrak_9350_hpd3": (
        "Thiết bị đếm hạt TSI AeroTrak 9350 và phụ kiện HPD III, USA",
        "Airborne particle counter TSI AeroTrak 9350 with HPD III accessory, USA",
        r"(?=.*(?:aerotrak|\b9350\b))(?=.*hpd\s*(?:iii|3)\b)",
    ),
    "aerotrak_9350": (
        "Thiết bị đếm hạt TSI AeroTrak 9350, USA",
        "Airborne particle counter TSI AeroTrak 9350, USA",
        r"aerotrak|\b9350\b",
    ),
    "lasair_iii_110_hpd2": (
        "Thiết bị đếm hạt PMS Lasair III 110 và phụ kiện HPD II, USA",
        "Airborne particle counter PMS Lasair III 110 with HPD II accessory, USA",
        r"(?=.*lasair)(?=.*hpd\s*(?:ii|2)\b)",
    ),
    "lasair_iii_110": (
        "Thiết bị đếm hạt PMS Lasair III 110, USA",
        "Airborne particle counter PMS Lasair III 110, USA",
        r"lasair",
    ),
    # --- Phòng sạch ---
    "testo_625": (
        "Thiết bị đo nhiệt độ, độ ẩm Testo 625, Germany",
        "Thermohygrometer Testo 625, Germany",
        r"testo\s*625",
    ),
    "testo_420": (
        "Thiết bị đo lưu lượng và chênh áp Testo 420, Germany",
        "Air volume and differential pressure meter Testo 420, Germany",
        r"testo\s*420",
    ),
    "buck_b30120": (
        "Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA",
        "Microbial air sampler Buck – Bio Culture B30120, USA",
        r"bio\s*culture|b30120",
    ),
    # --- Khí nén ---
    "cs_dp510": (
        "Thiết bị đo nhiệt độ điểm sương CS Instruments DP510, Germany",
        "Dew point meter CS Instruments DP510, Germany",
        r"dp\s*510",
    ),
    "cs_ds400": (
        "Thiết bị đo điểm sương và hàm lượng dầu CS Instruments DS 400, Germany",
        "Dew point and oil vapor measuring system CS Instruments DS 400, Germany",
        r"ds\s*400",
    ),
    "oil_kit": (   # người dùng chốt 15/11/2026 (hồ sơ 096547): lấy mẫu tổng dầu ghi theo vật tư thực dùng
        "Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®",
        "Glass fiber filter Pall life Sciences, carbon activated tube ORBO®",
        r"charcoal|orbo|than\s+hoạt\s+tính|carbon\s+activated",   # nhận cả tên cũ "Dụng cụ lấy mẫu tổng dầu (… Charcoal tube …)"
    ),
    "sas_pinocchio": (   # thêm 13/11/2026 (hồ sơ 134179 — BBLM "SAS Pinocchio Super II petri, Italia")
        "Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II, Italy",
        "Compressed air microbial sampler SAS Pinocchio Super II, Italy",
        r"pinocchio",
    ),
    # --- Nước (đo nhanh hiện trường) — người dùng chốt 16/11/2026: tên theo danh mục ---
    "mettler_s2": (
        "Thiết bị đo pH Mettler Toledo Seven2Go S2, USA",
        "pH meter Mettler Toledo Seven2Go S2, USA",
        r"seven\s*2\s*go|mettler",
    ),
    "hach_dr900": (
        "Thiết bị đo clo dư Hach DR900, USA",
        "Chlorine meter Hach DR900, USA",
        r"dr\s*900",
    ),
    # --- Nước thải / nước dưới đất (đo nhanh hiện trường) — thêm 16/11/2026 theo mẫu 100302, 106141, 106227 ---
    "hach_dr300": (
        "Thiết bị đo clo dư Hach DR300, USA",
        "Chlorine meter Hach DR300, USA",
        r"dr\s*300",
    ),
    "hanna_hi98192": (
        "Thiết bị đo TDS Hanna HI 98192, USA",
        "TDS meter Hanna HI 98192, USA",
        r"hi\s*98192",
    ),
    "hanna_hi98190": (
        "Thiết bị đo pH Hanna HI 98190, USA",
        "pH meter Hanna HI 98190, USA",
        r"hi\s*98190",
    ),
    "hanna_hi9142": (                     # 140604 (16/11/2026): DO nước mặt đo hiện trường
        "Thiết bị đo oxy hòa tan Hanna HI 9142, USA",
        "Dissolved oxygen meter Hanna HI 9142, USA",
        r"hi\s*9142\b",
    ),
    "hanna_hi93703": (                    # 141523 (16/11/2026): độ đục nước sinh hoạt đo hiện trường
        "Thiết bị đo độ đục Hanna HI 93703, USA",
        "Turbidity meter Hanna HI 93703, USA",
        r"hi\s*93703",
    ),
}

_RX = {k: re.compile(v[2], re.I) for k, v in EQUIPMENT.items()}
_ACCESSORY_BASE = {"aerotrak_9350_hpd3": "aerotrak_9350", "lasair_iii_110_hpd2": "lasair_iii_110"}


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s or "")


def match_key(text: str) -> str | None:
    """Mã máy nếu ``text`` gọi đúng MỘT máy trong danh mục; 0 hoặc ≥ 2 máy → None (giữ nguyên)."""
    t = _nfc(text).lower()
    if not t.strip():
        return None
    hits = [k for k, rx in _RX.items() if rx.search(t)]
    for acc, base in _ACCESSORY_BASE.items():  # bản có phụ kiện đã bao luôn máy gốc
        if acc in hits and base in hits:
            hits.remove(base)
    return hits[0] if len(hits) == 1 else None


def name(key: str, lang: str = "vn") -> str:
    vn, en, _ = EQUIPMENT[key]
    return en if lang == "en" else vn


def canonical_name(text: str, lang: str = "vn") -> str:
    """Tên chuẩn của máy mà ``text`` gọi tới (giữ nguyên ``text`` nếu không nhận ra đúng một máy)."""
    key = match_key(text)
    return name(key, lang) if key else (text or "")


def normalize_equipment(name_vn: str, name_en: str) -> tuple[str, str]:
    """Quy cặp tên VN/EN về tên chuẩn; EN trống thì điền theo VN (và ngược lại)."""
    kv, ke = match_key(name_vn), match_key(name_en)
    if kv and (not (name_en or "").strip() or ke == kv):
        return name(kv, "vn"), name(kv, "en")
    if ke and not (name_vn or "").strip():
        return name(ke, "vn"), name(ke, "en")
    return (name(kv, "vn") if kv else name_vn), (name(ke, "en") if ke else name_en)


def eq(key: str, **kw):
    """``Equipment`` có tên chuẩn: ``eq("rion_nl53")``; ``kw`` thêm model/serial/origin nếu cần."""
    from Core.Models.dossier import Equipment
    vn, en, _ = EQUIPMENT[key]
    return Equipment(name_vn=vn, name_en=en, **kw)


def as_dict(key: str) -> dict:
    """``{"name_vn": …, "name_en": …}`` cho chỗ dùng dict (intake_helper, profile JSON)."""
    vn, en, _ = EQUIPMENT[key]
    return {"name_vn": vn, "name_en": en}


def _core(line: str) -> str:
    s = _nfc(line).strip()
    s = re.sub(r"^[-–•]\s*", "", s)
    return s.rstrip(" ;.").strip()


def nonstandard_lines(text: str) -> list[tuple[str, str]]:
    """[(dòng đang ghi, tên chuẩn)] cho các dòng ``- …`` gọi máy trong danh mục bằng tên khác chuẩn."""
    out = []
    for line in _nfc(text).splitlines():
        if not line.strip().startswith(("-", "–", "•")):
            continue
        core = _core(line)
        key = match_key(core)
        if not key:
            continue
        vn, en, _ = EQUIPMENT[key]
        if core not in (vn, en):
            out.append((core, vn if re.search(r"[À-ỹĐđ]", core) else en))
    return out


if __name__ == "__main__":  # in bảng tên chuẩn
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    for k, (vn, en, _) in EQUIPMENT.items():
        print(f"{k:22} | {vn} | {en}")
