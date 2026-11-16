# -*- coding: utf-8 -*-
"""
Chế độ số liệu máy đếm hạt — vi phân (Δ / Diff) hay cộng dồn (Σ / Cumul) (người dùng chốt 11/11/2026, AGENTS mục 8C):

    Khí nén (ISO 8573-1, cấp theo DẢI cỡ hạt 0,1–0,5 / 0,5–1 / 1–5 µm)  -> số VI PHÂN:
        Lasair III  cột "Δ (N/m³)"      AeroTrak 9350  cột "Diff"
    Phòng sạch, tủ (ISO 14644-1, giới hạn d ≥ cỡ hạt)                 -> số CỘNG DỒN:
        Lasair III 110  cột "Σ (N/m³)"  AeroTrak 9350  cột "Cumul"

File xuất từ 9350 (`6ChannelFixed.xlsx`, TrakPro) ghi chế độ ở đầu file `Count Mode: Σ | Δ` và đơn vị
`Units: # / m³ | #` — PHẢI kiểm tra trước khi dùng (`aerotrak_meta`); `aerotrak_values` tự đổi về chế độ cần dùng
(và # -> #/m³ theo `Vol. (L)`). Database Lasair (`Counts n`) luôn là số đếm vi phân -> `lasair_values`.

    from Core.Tools.particle_mode import aerotrak_meta, aerotrak_values, DIFF, CUMUL
    meta = aerotrak_meta(path)                       # {"mode": "Σ", "units": "# / m³", "serial": ...}
    ch = aerotrak_values(path, 1113, want=DIFF)      # [0,3; 0,5; 1; 2; 3; 5 µm] hạt/m³ vi phân (khí nén)

Chạy kiểm tra: python Core/Tools/particle_mode.py <6ChannelFixed.xlsx> [số mẫu …]
"""
from __future__ import annotations

import sys
from typing import Dict, List, Sequence

import openpyxl

DIFF = "diff"     # Δ — số từng kênh (khí nén)
CUMUL = "cumul"   # Σ — cộng dồn từ kênh đó trở lên (phòng sạch, tủ)


def norm_mode(text) -> str:
    """'Σ', 'Cumul', 'Cumulative', 'sigma' -> CUMUL; 'Δ', 'Diff', 'Differential', 'Delta' -> DIFF; khác -> ''."""
    s = str(text or "").strip().lower()
    if not s:
        return ""
    if "σ" in s or s.startswith("cum") or "sigma" in s or s == "sum":
        return CUMUL
    if "δ" in s or "∆" in s or s.startswith("diff") or "delta" in s:
        return DIFF
    return ""


def to_diff(cumul: Sequence[float]) -> List[float]:
    """Σ -> Δ: kênh i = Σ_i − Σ_(i+1), kênh cuối giữ nguyên."""
    c = [float(v or 0) for v in cumul]
    return [c[i] - c[i + 1] for i in range(len(c) - 1)] + [c[-1]]


def to_cumul(diff: Sequence[float]) -> List[float]:
    """Δ -> Σ: kênh i = tổng các kênh từ i trở lên."""
    d = [float(v or 0) for v in diff]
    return [sum(d[i:]) for i in range(len(d))]


def check_cumul_shape(values: Sequence[float], what: str = "") -> bool:
    """Số cộng dồn phải giảm dần theo cỡ hạt; tăng lên là dấu hiệu dữ liệu thực ra là Δ -> in cảnh báo."""
    try:
        v = [float(x or 0) for x in values]
    except (TypeError, ValueError):
        return True
    bad =[i for i in range(len(v) - 1) if v[i] < v[i + 1]]
    if bad:
        print(f"[CRITICAL WARNING] [Particle Mode] {what}: số 'cộng dồn' không giảm dần {v} — kiểm tra lại có phải Δ")
    return not bad


# ---------------------------------------------------------------------------
# AeroTrak 9350 — 6ChannelFixed.xlsx
# ---------------------------------------------------------------------------
_CH = [f"Ch{i}" for i in range(1, 7)]


def _rows(path: str):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["6ChannelFixed"] if "6ChannelFixed" in wb.sheetnames else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def _after(row, label):
    vals = [str(v).strip() if v is not None else "" for v in row]
    if label in vals:
        for v in vals[vals.index(label) + 1:]:
            if v:
                return v
    return None


def aerotrak_meta(path: str) -> Dict[str, str]:
    """Đầu file 9350 -> {"mode": DIFF | CUMUL, "mode_raw": "Σ", "units": "# / m³", "per_m3": bool, "serial", "model"}.
    Không thấy dòng `Count Mode:` -> ValueError (không đoán)."""
    meta = {"mode": "", "mode_raw": "", "units": "", "serial": "", "model": ""}
    for row in _rows(path)[:40]:
        for key, lbl in (("mode_raw", "Count Mode:"), ("units", "Units:"), ("serial", "Serial:"), ("model", "Model:")):
            v = _after(row, lbl)
            if v and not meta[key]:
                meta[key] = v
    meta["mode"] = norm_mode(meta["mode_raw"])
    if not meta["mode"]:
        raise ValueError(f"[Particle Mode] {path}: không đọc được 'Count Mode' (Σ/Δ) — mở file kiểm tra tay")
    u = meta["units"].replace(" ", "").lower()
    meta["per_m3"] = "/m" in u or "m³" in u or "m3" in u
    return meta


def aerotrak_records(path: str) -> Dict[int, dict]:
    """{số mẫu #: {tiêu đề cột rút gọn: giá trị}} — giữ nguyên số máy (chưa đổi chế độ)."""
    hdr, out = None, {}
    for vals in _rows(path):
        if hdr is None:
            if "#" in vals and "Location" in vals:
                hdr = {str(v).split("\n")[0].strip(): i for i, v in enumerate(vals) if v not in (None, "")}
            continue
        n = vals[hdr["#"]] if "#" in hdr else None
        if isinstance(n, (int, float)):
            out[int(n)] = {k: vals[i] for k, i in hdr.items() if i < len(vals)}
    if hdr is None:
        raise ValueError(f"Không thấy dòng tiêu đề AeroTrak trong {path}")
    return out


def aerotrak_values(path: str, sample: int, want: str, meta: dict = None, recs: dict = None) -> List[float]:
    """6 kênh [0,3; 0,5; 1; 2; 3; 5 µm] hạt/m³ của mẫu `sample` theo chế độ `want` (DIFF khí nén, CUMUL phòng sạch).
    File khác chế độ -> tự đổi và in [INFO]; đơn vị '#' -> chia Vol. (L)/1000."""
    meta = meta or aerotrak_meta(path)
    recs = recs or aerotrak_records(path)
    if sample not in recs:
        raise KeyError(f"{path}: không có mẫu #{sample}")
    r = recs[sample]
    v = [float(r.get(c) or 0) for c in _CH]
    if not meta["per_m3"]:
        vol_l = float(r.get("Vol.") or 0)
        if vol_l <= 0:
            raise ValueError(f"{path} #{sample}: đơn vị '{meta['units']}' nhưng không có Vol. (L) để quy về hạt/m³")
        v = [x * 1000.0 / vol_l for x in v]
    if meta["mode"] == CUMUL:
        check_cumul_shape(v, f"{path} #{sample}")
    if meta["mode"] != want:
        print(f"[INFO] [Particle Mode] {path} #{sample}: file chế độ {meta['mode_raw']} -> đổi sang "
              f"{'Δ (vi phân)' if want == DIFF else 'Σ (cộng dồn)'}")
        v = to_diff(v) if want == DIFF else to_cumul(v)
    return v


def require_mode(path: str, want: str, purpose: str = "") -> dict:
    """Dừng nếu file 9350 không đúng chế độ dùng thẳng giá trị máy (vd bảng phòng sạch chép nguyên số Σ)."""
    meta = aerotrak_meta(path)
    if meta["mode"] != want or not meta["per_m3"]:
        raise ValueError(
            f"[Particle Mode] {path}: file là '{meta['mode_raw']}', đơn vị '{meta['units']}' — {purpose} cần "
            f"{'Σ (cộng dồn)' if want == CUMUL else 'Δ (vi phân)'}, # / m³. Xuất lại từ TrakPro đúng chế độ, hoặc "
            f"đọc bằng aerotrak_values(..., want=...) để tự đổi.")
    return meta


# ---------------------------------------------------------------------------
# Lasair III — database (Counts n = số đếm vi phân của kênh n)
# ---------------------------------------------------------------------------
def lasair_values(rec: dict, want: str) -> List[float]:
    """8 kênh [0,1 … 5 µm] hạt/m³ từ bản ghi database Lasair theo chế độ `want`."""
    vol = float(str(rec["Sample Volume"]).replace(",", "."))
    d = [float(str(rec[f"Counts {i}"]).replace(",", ".")) / vol for i in range(1, 9)]
    return d if want == DIFF else to_cumul(d)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    p = sys.argv[1]
    m = aerotrak_meta(p)
    print(f"{p}: Model {m['model']} S/N {m['serial']} | Count Mode {m['mode_raw']} ({m['mode']}) | Units {m['units']}")
    recs = aerotrak_records(p)
    for s in [int(x) for x in sys.argv[2:]] or sorted(recs)[:3]:
        print(f"  #{s}  Δ {[round(x) for x in aerotrak_values(p, s, DIFF, m, recs)]}"
              f"  Σ {[round(x) for x in aerotrak_values(p, s, CUMUL, m, recs)]}")
