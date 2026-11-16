"""Tra cứu thư viện giấy hiệu chuẩn thiết bị (HieuChuan/thu_vien_hieu_chuan.json).

Nguồn: Z:\\9.NV2 - MOI TRUONG\\3.1Giay hieu chuan thiet bi\\{2026,2027} (chỉ đọc). Dựng lại thư viện:
    python HieuChuan/extract_layout.py && python HieuChuan/parse_cert.py && python HieuChuan/build_library.py

Ví dụ:
    from Core.Tools.calib_library import find, valid_cert, cal_table, corrected
    c = valid_cert("MT33-3", "29/10/2026")                  # giấy còn hiệu lực ngày đo
    tbl = cal_table("MT33-3", "nhiệt độ", "29/10/2026")    # [(chỉ thị, chuẩn, sai số), ...]
    corrected(20.5, tbl)                                     # chỉ thị − sai số (nội suy; ngoài khoảng: điểm biên)

Hiệu chính (AGENTS mục 15, giống build_132809.calibrated): giá trị đúng = chỉ thị − sai số; sai số nội suy tuyến
tính giữa 2 điểm hiệu chuẩn kẹp hai bên (như FORECAST); ngoài khoảng hiệu chuẩn trừ sai số điểm biên gần nhất,
không ngoại suy.
"""
import datetime as dt
import json
import os
import re
import unicodedata

LIB_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "HieuChuan", "thu_vien_hieu_chuan.json")
_LIB = None


def _d(s):
    if isinstance(s, (dt.date, dt.datetime)):
        return s if isinstance(s, dt.date) and not isinstance(s, dt.datetime) else s.date()
    try:
        return dt.datetime.strptime(str(s).strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def _fold(s):
    s = unicodedata.normalize("NFD", str(s or "").lower().replace("đ", "d"))
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn")


def _qmatch(quantity, p):
    """'nhiệt độ' khớp đại lượng chuẩn (qty) hoặc tiêu đề bảng (quantity)."""
    q = _fold(quantity)
    return q == _fold(p.get("qty")) or q in _fold(p.get("quantity"))


def load(path=None, reload=False):
    global _LIB
    if _LIB is None or reload or path:
        with open(path or LIB_PATH, encoding="utf-8") as fh:
            _LIB = json.load(fh)
    return _LIB


def certificates():
    return load()["certificates"]


def find(code=None, serial=None, text=None, kind=None):
    """Lọc giấy theo mã MT (vd 'MT33-3', 'MT33' khớp cả MT33-1..3), serial, chữ trong tên/model, loại."""
    out = []
    for c in certificates():
        if code:
            cc = code.upper().replace(" ", "")
            if not (c["mt_code"] == cc or ("-" not in cc and c["mt_code"].split("-")[0] == cc)):
                continue
        if serial and _fold(serial) not in _fold(c.get("serial")):
            continue
        if text and _fold(text) not in _fold(" ".join(str(c.get(k, "")) for k in ("object", "model", "manufacturer", "file"))):
            continue
        if kind and _fold(kind) not in _fold(c.get("kind")):
            continue
        out.append(c)
    return out


def valid_cert(code, on_date, kind="hiệu chuẩn", serial=None, quantity=None):
    """Giấy có ngày HC ≤ ngày đo ≤ ngày HC lại; nhiều giấy thì lấy giấy HC gần nhất trước ngày đo."""
    d = _d(on_date)
    best = None
    for c in find(code=code, serial=serial, kind=kind):
        a, b = _d(c.get("cal_date")), _d(c.get("recal_date"))
        if not a or a > d or (b and b < d):
            continue
        if c.get("status") == "không đạt":
            continue
        if quantity and not any(_qmatch(quantity, p) for p in c.get("points", [])):
            continue
        if best is None or a > _d(best["cal_date"]):
            best = c
    return best


def cal_table(code, quantity, on_date, serial=None):
    """[(chỉ thị, chuẩn, sai số), ...] tăng dần theo chỉ thị, của đại lượng `quantity` trong giấy còn hiệu lực."""
    c = valid_cert(code, on_date, serial=serial, quantity=quantity)
    if not c:
        raise LookupError(f"Không có giấy hiệu chuẩn {code} ({quantity}) còn hiệu lực ngày {on_date}")
    pts = [p for p in c["points"] if _qmatch(quantity, p) and p.get("ind") is not None and p.get("err") is not None]
    return sorted(((p["ind"], p.get("ref"), p["err"]) for p in pts), key=lambda t: t[0])


def corrected(x, table):
    """Giá trị sau hiệu chính = chỉ thị − sai số (nội suy tuyến tính; ngoài khoảng: sai số điểm biên)."""
    if not table:
        raise ValueError("Bảng hiệu chuẩn rỗng")
    if x <= table[0][0]:
        return x - table[0][2]
    if x >= table[-1][0]:
        return x - table[-1][2]
    for (x1, _, e1), (x2, _, e2) in zip(table, table[1:]):
        if x1 <= x <= x2:
            return x - (e1 + (e2 - e1) * (x - x1) / (x2 - x1)) if x2 != x1 else x - e1
    raise ValueError(x)


def expiring(within_days=60, today=None):
    """Giấy đang dùng (không bị giấy mới thay) hết hạn trong `within_days` ngày tới hoặc đã hết hạn."""
    t = _d(today) if today else dt.date.today()
    out = []
    for c in certificates():
        b = _d(c.get("recal_date"))
        if b and not c.get("superseded_by") and (b - t).days <= within_days:
            out.append(c)
    return sorted(out, key=lambda c: _d(c["recal_date"]))


if __name__ == "__main__":
    import sys
    q = sys.argv[1:] or [None]
    for c in find(code=q[0]) if q[0] else certificates():
        print(f"{c['mt_code']:8} {c['cert_no']:12} {c['cal_date']:10} → {c['recal_date']:10} {c['status']:22} "
              f"{c['object'][:40]} | {c['model']} | SN {c['serial']}")
