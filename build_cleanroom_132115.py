# -*- coding: utf-8 -*-
"""
Builder script for Cleanroom Inspection Dossier Q326D132115.
Generates:
1. BCKQ - 1 - phong sach - Q326D132115.docx
2. TBKQ - viet - 1 - phong sach - Q326D132115.docx
3. TBKQ - en - 1 - phong sach - Q326D132115.docx
"""

import os
import docx


# Dossier Constants
CODE = "Q326D132115"
DATE_APP = "19/10/2026"
DATE_MEAS = "23/10/2026"
DATE_ISSUE = "23/10/2026"

CLIENT_VN = "Công ty TNHH Lâm Quang Đại"
CLIENT_ADDR_VN = "89 Lê Thị Riêng, Phường Thới An, Thành phố Hồ Chí Minh"
CLIENT_EN = "Lam Quang Dai Company Limited"
CLIENT_ADDR_EN = "89 Le Thi Rieng, Thoi An Ward, Ho Chi Minh City"

SITE_VN = "Công ty TNHH Thực phẩm Farina - Chi nhánh Hồ Chí Minh (Dự án Farina - Khu R0A, KCN Việt Nam - Singapore III, phường Tân Uyên, TP. Tân Uyên, Bình Dương)"
SITE_EN = "Farina Food Company Limited - Ho Chi Minh Branch (Farina Project - Zone R0A, Vietnam - Singapore III Industrial Park, Tan Uyen Ward, Tan Uyen City, Binh Duong Province)"

GDV_NAME_VN = "Nguyễn Xuân Nam"
GDV_NAME_EN = "Nguyen Xuan Nam"
BOSS_TITLE_VN = "KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC"
BOSS_TITLE_EN = "FOR DIRECTOR\nVICE DIRECTOR"
BOSS_NAME = "Lê Đăng Huyền"

# 10 Cleanroom Rooms info
ROOMS = [
    {
        "stt": 1,
        "name_vn": "Phòng đóng gói",
        "name_en": "Packaging Room",
        "area_vn": "57,0",
        "area_en": "57.0",
        "class": "ISO Class 8",
        "points_summary_vn": ["12 vị trí (từ vị trí 1 đến vị trí 12)"],
        "points_summary_en": ["12 points (from point 1 to point 12)"],
        "points_count": 12,
        "zero_sample": 887,
        "start_sample": 888,
    },
    {
        "stt": 2,
        "name_vn": "Phòng nướng",
        "name_en": "Baking Room",
        "area_vn": "45,0",
        "area_en": "45.0",
        "class": "ISO Class 8",
        "points_summary_vn": ["10 vị trí (từ vị trí 1 đến vị trí 10)"],
        "points_summary_en": ["10 points (from point 1 to point 10)"],
        "points_count": 10,
        "zero_sample": 900,
        "start_sample": 901,
    },
    {
        "stt": 3,
        "name_vn": "Phòng trộn",
        "name_en": "Mixing Room",
        "area_vn": "33,7",
        "area_en": "33.7",
        "class": "ISO Class 8",
        "points_summary_vn": ["09 vị trí (từ vị trí 1 đến vị trí 9)"],
        "points_summary_en": ["09 points (from point 1 to point 9)"],
        "points_count": 9,
        "zero_sample": 948,   # "Zero 911" + samples 912-920 are struck out in the field sheet
        "start_sample": 949,
    },
    {
        "stt": 4,
        "name_vn": "Phòng chuẩn bị nguyên liệu",
        "name_en": "Material Preparation Room",
        "area_vn": "17,1",
        "area_en": "17.1",
        "class": "ISO Class 8",
        "points_summary_vn": ["06 vị trí (từ vị trí 1 đến vị trí 6)"],
        "points_summary_en": ["06 points (from point 1 to point 6)"],
        "points_count": 6,
        "zero_sample": 921,
        "start_sample": 922,
    },
    {
        "stt": 5,
        "name_vn": "Phòng đóng thùng",
        "name_en": "Cartoning Room",
        "area_vn": "19,4",
        "area_en": "19.4",
        "class": "ISO Class 8",
        "points_summary_vn": ["06 vị trí (từ vị trí 1 đến vị trí 6)"],
        "points_summary_en": ["06 points (from point 1 to point 6)"],
        "points_count": 6,
        "zero_sample": 937,
        "start_sample": 938,
    },
    {
        "stt": 6,
        "name_vn": "Phòng thay đồ",
        "name_en": "Gowning Room",
        "area_vn": "11,1",
        "area_en": "11.1",
        "class": "ISO Class 8",
        "points_summary_vn": ["06 vị trí (từ vị trí 1 đến vị trí 6)"],
        "points_summary_en": ["06 points (from point 1 to point 6)"],
        "points_count": 6,
        "zero_sample": 962,
        "start_sample": 963,
    },
    {
        "stt": 7,
        "name_vn": "Phòng phụ liệu",
        "name_en": "Ingredient Room",
        "area_vn": "20,1",
        "area_en": "20.1",
        "class": "ISO Class 8",
        "points_summary_vn": ["03 vị trí (từ vị trí 1 đến vị trí 3)"],
        "points_summary_en": ["03 points (from point 1 to point 3)"],
        "points_count": 3,
        "zero_sample": 944,
        "start_sample": 945,
    },
    {
        "stt": 8,
        "name_vn": "Phòng đệm 1",
        "name_en": "Airlock 1",
        "area_vn": "4,1",
        "area_en": "4.1",
        "class": "ISO Class 8",
        "points_summary_vn": ["03 vị trí (từ vị trí 1 đến vị trí 3)"],
        "points_summary_en": ["03 points (from point 1 to point 3)"],
        "points_count": 3,
        "zero_sample": 928,   # sample 929 is struck out in the field sheet
        "start_sample": 930,
    },
    {
        "stt": 9,
        "name_vn": "Phòng đệm 2",
        "name_en": "Airlock 2",
        "area_vn": "3,9",
        "area_en": "3.9",
        "class": "ISO Class 8",
        "points_summary_vn": ["02 vị trí (từ vị trí 1 đến vị trí 2)"],
        "points_summary_en": ["02 points (from point 1 to point 2)"],
        "points_count": 2,
        "zero_sample": 933,   # sample 936 is struck out in the field sheet
        "start_sample": 934,
    },
    {
        "stt": 10,
        "name_vn": "Phòng rửa CCDC",
        "name_en": "Cleaning Room",
        "area_vn": "28,0",
        "area_en": "28.0",
        "class": "ISO Class 8",
        "points_summary_vn": ["03 vị trí (từ vị trí 1 đến vị trí 3)"],
        "points_summary_en": ["03 points (from point 1 to point 3)"],
        "points_count": 3,
        "zero_sample": 958,
        "start_sample": 959,
    }
]

# HEPA Filters H1 to H8
HEPA_FILTERS = [
    {"kh": "H1", "room_vn": "Phòng trộn", "room_en": "Mixing Room", "pos_vn": "H1", "pos_en": "H1", "leak": "< 0,0001", "leak_en": "< 0.0001"},
    {"kh": "H2", "room_vn": "Phòng trộn", "room_en": "Mixing Room", "pos_vn": "H2", "pos_en": "H2", "leak": "< 0,0001", "leak_en": "< 0.0001"},
    {"kh": "H3", "room_vn": "Phòng phụ liệu", "room_en": "Ingredient Room", "pos_vn": "H3", "pos_en": "H3", "leak": "< 0,0001", "leak_en": "< 0.0001"},
    {"kh": "H4", "room_vn": "Phòng nướng", "room_en": "Baking Room", "pos_vn": "H4", "pos_en": "H4", "leak": "< 0,0001", "leak_en": "< 0.0001"},
    {"kh": "H5", "room_vn": "Phòng đóng gói", "room_en": "Packaging Room", "pos_vn": "H5", "pos_en": "H5", "leak": "< 0,0001", "leak_en": "< 0.0001"},
    {"kh": "H6", "room_vn": "Phòng đóng gói", "room_en": "Packaging Room", "pos_vn": "H6", "pos_en": "H6", "leak": "< 0,0001", "leak_en": "< 0.0001"},
    {"kh": "H7", "room_vn": "Phòng đóng gói", "room_en": "Packaging Room", "pos_vn": "H7", "pos_en": "H7", "leak": "< 0,0001", "leak_en": "< 0.0001"},
    {"kh": "H8", "room_vn": "Phòng đóng gói", "room_en": "Packaging Room", "pos_vn": "H8", "pos_en": "H8", "leak": "< 0,0001", "leak_en": "< 0.0001"},
]

def load_particle_data():
    """Loads exact normalized particle concentrations from 6ChannelFixed.xlsx"""
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join('Input', '132115', '132115 - Data Do', '6ChannelFixed.xlsx'))
    ws = wb.active
    particles = {}
    for r in range(23, ws.max_row + 1):
        sid = ws.cell(r, 3).value
        if sid is not None:
            c05 = ws.cell(r, 20).value
            c10 = ws.cell(r, 22).value
            c50 = ws.cell(r, 26).value
            particles[int(sid)] = (c05, c10, c50)
    return particles

PARTICLE_DATA = load_particle_data()

def fmt_int_vn(val):
    if val is None:
        return "0"
    return f"{int(val):,}".replace(",", " ")

def fmt_int_en(val):
    if val is None:
        return "0"
    return f"{int(val):,}".replace(",", " ")

print(f"Loaded particle data for {len(PARTICLE_DATA)} samples.")


# ============================================================================================
# Temperature / humidity: every room has SEVERAL measuring points and each point is read 3 times
# (Testo 625, hand-written field sheet "Ghi chép giám định tại hiện trường", pages 4-6).
# Each entry = ((T1, T2, T3), (RH1, RH2, RH3)) raw readings of one point, in point order 1..n.
# ============================================================================================
import re
from decimal import Decimal, ROUND_HALF_UP


ROOM_BY_VN = {r["name_vn"]: r for r in ROOMS}

# ============================================================================================
# SO LIEU DO + MUC YEU CAU doc tu file Excel tinh toan do nguoi dung dien (26/10/2026):
#   Input\132115\132115 - Data Do\Q326D132115 - TinhToanPhongSach.xlsx
#   - "Bụi 0,3"   : cot "Cấp sạch yêu cầu" (cap sach tung phong). So dem hat lay tu 6ChannelFixed theo SO MAU
#                   cua phieu ghi chep (ROOMS zero_sample/start_sample), KHONG theo cac dong mau trong sheet.
#   - "Nhiệt Ẩm"  : 3 lan do/diem + "Mức đánh giá nhiệt" / "Mức đánh giá ẩm" (N/A = khong danh gia).
#   - "Lưu lượng" : 3 lan do/mieng, he so hieu chuan (cot HC), the tich phong; khong co cot muc -> khong danh gia ACH.
#   - "Chênh áp"  : cap vi tri, ky hieu, "Mức đánh giá" (N/A = khong danh gia), 3 lan do.
# ============================================================================================
#
# Tu 30/10/2026 nguoi dung sua so lieu truc tiep tren phieu tinh toan ThucHien\<mã> - TinhToanPhongSach.xlsx
# (vd sua do am, them diem chenh ap PA12) -> doc so lieu tu phieu do (CALC_XLSX); file Data Do (DATA_XLSX) chi con
# la nguon de fill_excel_132115.py dung phieu lan dau / khi chay --tao-lai.
# ============================================================================================
import unicodedata
import openpyxl as _openpyxl

DATA_XLSX = os.path.join("Input", "132115", "132115 - Data Do", f"{CODE} - TinhToanPhongSach.xlsx")
CALC_XLSX = os.path.join("Input", "132115", "ThucHien", f"{CODE} - TinhToanPhongSach.xlsx")

# ISO 14644-1:2015 Bang E.1 (hat/m3) cho co hat 0,5 / 1 / 5 um
ISO_CLASS_LIMITS = {
    5: {"0.5": 3_520, "1": 832, "5": None},
    6: {"0.5": 35_200, "1": 8_320, "5": 293},
    7: {"0.5": 352_000, "1": 83_200, "5": 2_930},
    8: {"0.5": 3_520_000, "1": 832_000, "5": 29_300},
    9: {"0.5": 35_200_000, "1": 8_320_000, "5": 293_000},
}


def _norm(txt):
    t = unicodedata.normalize("NFC", str(txt or "")).strip().lower()
    t = re.sub(r"^p\.\s*", "phòng ", t)
    t = re.sub(r"\bnl\b", "nguyên liệu", t)
    return re.sub(r"\s+", " ", t)


def _room_of(txt):
    """'P. Chuẩn bị NL' -> ROOMS entry 'Phòng chuẩn bị nguyên liệu' (None neu khong phai phong, vd 'Môi trường')."""
    n = _norm(txt)
    for r in ROOMS:
        if _norm(r["name_vn"]) == n:
            return r
    return None


def parse_spec(v):
    """'≤ 25' -> (None, 25); '≥ 5' -> (5, None); '20 - 25' -> (20, 25); 'N/A' / trong -> None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return (None, Decimal(str(v)))
    t = unicodedata.normalize("NFC", str(v)).strip().replace(",", ".")
    if not t or t.upper() in ("N/A", "NA", "-") or "không" in t.lower():
        return None
    nums = [Decimal(x) for x in re.findall(r"-?\d+(?:\.\d+)?", t)]
    if not nums:
        return None
    if "≤" in t or "<=" in t or "<" in t:
        return (None, nums[0])
    if "≥" in t or ">=" in t or ">" in t:
        return (nums[0], None)
    if len(nums) >= 2:
        return (nums[0], nums[1])
    return (None, nums[0])


def in_spec(v, spec):
    if spec is None:
        return True
    lo, hi = spec
    return (lo is None or v >= lo) and (hi is None or v <= hi)


def _num(x):
    return f"{x.normalize():f}" if isinstance(x, Decimal) else str(x)


def spec_txt(spec, is_en=False):
    if spec is None:
        return "Not specified" if is_en else "Không quy định"
    lo, hi = spec
    if lo is not None and hi is not None:
        s = f"{_num(lo)} – {_num(hi)}"
    elif hi is not None:
        s = f"≤ {_num(hi)}"
    else:
        s = f"≥ {_num(lo)}"
    return s if is_en else s.replace(".", ",")


def _formula_offset(f):
    """'=F3-0.1' -> Decimal('-0.1'); '=D2*0.981' -> Decimal('0.981')."""
    m = re.search(r"([-+*])\s*(\d+(?:\.\d+)?)\s*$", str(f or ""))
    if not m:
        return None
    v = Decimal(m.group(2))
    return -v if m.group(1) == "-" else v


def _eval_simple(f):
    """'=57*2.8' -> Decimal('159.6') (chi so va + - * / ( ))."""
    if isinstance(f, (int, float)):
        return Decimal(str(f))
    expr = str(f or "").lstrip("=").replace(",", ".")
    if not re.fullmatch(r"[\d.+\-*/() ]+", expr):
        raise ValueError(f"Khong tinh duoc cong thuc '{f}'")
    expr = re.sub(r"\d+(?:\.\d+)?", lambda m: f"Decimal('{m.group(0)}')", expr)
    return eval(expr, {"__builtins__": {}, "Decimal": Decimal}, {})


def _header_col(ws, *names):
    for c in range(1, ws.max_column + 1):
        v = _norm(ws.cell(1, c).value)
        if any(_norm(n) in v for n in names):
            return c
    return None


def _carry_rows(ws, room_col, first_row):
    """(room_text, row) cho moi dong du lieu; o phong trong = phong o tren (o gop)."""
    cur = None
    for r in range(first_row, ws.max_row + 1):
        if all(ws.cell(r, c).value in (None, "") for c in range(1, min(ws.max_column, 12) + 1)):
            continue
        v = ws.cell(r, room_col).value
        if v not in (None, ""):
            cur = str(v).strip()
        yield cur, r


def _merged_value(ws, r, c):
    """Gia tri o (r, c); o nam trong vung gop -> gia tri o goc tren trai (phieu ThucHien gop cot muc theo phong)."""
    v = ws.cell(r, c).value
    if v is None:
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                return ws.cell(rng.min_row, rng.min_col).value
    return v


def _load_excel():
    src = CALC_XLSX if os.path.isfile(CALC_XLSX) else DATA_XLSX
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Khong thay file so lieu {CALC_XLSX} / {DATA_XLSX}")
    print(f"Doc so lieu tu: {src}")
    wb = _openpyxl.load_workbook(src)                # cong thuc (khong data_only) de doc he so hieu chuan
    out = {}

    # --- Cap sach yeu cau (sheet Bui)
    ws = wb["Bụi 0,3"]
    col_cls = _header_col(ws, "Cấp sạch")
    classes = {}
    for room_txt, r in _carry_rows(ws, 1, 2):
        rr = _room_of(room_txt)
        v = ws.cell(r, col_cls).value if col_cls else None
        if rr and v:
            m = re.search(r"(\d+)", str(v))
            if m:
                classes.setdefault(rr["name_vn"], set()).add(int(m.group(1)))
    out["classes"] = classes

    # --- Nhiet do, do am
    ws = wb["Nhiệt Ẩm"]
    c_t = _header_col(ws, "Kết quả nhiệt")
    c_h = _header_col(ws, "Kết quả độ ẩm")
    c_st = _header_col(ws, "Mức đánh giá nhiệt")
    c_sh = _header_col(ws, "Mức đánh giá ẩm")
    c_cal = _header_col(ws, "Sau hiệu chuẩn")
    trh, specs, t_off, h_off = {}, {}, None, None
    for room_txt, r in _carry_rows(ws, 2, 3):
        rr = _room_of(room_txt)
        if rr is None:
            raise ValueError(f"Sheet Nhiệt Ẩm dong {r}: khong nhan ra phong '{room_txt}'")
        temps = tuple(ws.cell(r, c_t + i).value for i in range(3))
        rhs = tuple(ws.cell(r, c_h + i).value for i in range(3))
        if any(v is None for v in temps + rhs):
            raise ValueError(f"Sheet Nhiệt Ẩm dong {r}: thieu so do")
        trh.setdefault(rr["name_vn"], []).append((temps, rhs))
        specs.setdefault(rr["name_vn"], []).append((parse_spec(_merged_value(ws, r, c_st)), parse_spec(_merged_value(ws, r, c_sh))))
        if t_off is None and c_cal:
            t_off = _formula_offset(ws.cell(r, c_cal).value)
            h_off = _formula_offset(ws.cell(r, c_cal + 3).value)
    out.update(trh=trh, trh_specs=specs, t_off=t_off if t_off is not None else Decimal("0"),
               h_off=h_off if h_off is not None else Decimal("0"))

    # --- Luu luong + ACH
    ws = wb["Lưu lượng"]
    c_spec_ach = _header_col(ws, "Mức đánh giá", "Mức quy định", "Ngưỡng")
    flows = []          # [(room_vn, grille_vn, grille_en, avg_Decimal, volume or None, ach_spec)]
    cur_room = None
    for r in range(2, ws.max_row + 1):
        label = ws.cell(r, 1).value
        if label in (None, ""):
            continue
        label = str(label).strip()
        m = re.match(r"^(.*?)\s*-\s*(.+)$", label)
        room_txt, g_txt = (m.group(1), m.group(2)) if m else (re.sub(r"\s*\(.*\)\s*$", "", label), None)
        rr = _room_of(room_txt)
        if rr is None:
            raise ValueError(f"Sheet Lưu lượng dong {r}: khong nhan ra phong '{room_txt}'")
        code = re.search(r"\((H\d+)\)", label)
        if g_txt is None:
            g_vn = "Miệng cấp" + (f" ({code.group(1)})" if code else "")
            g_en = "Supply grille" + (f" ({code.group(1)})" if code else "")
        else:
            g_vn = g_txt
            g_en = re.sub(r"^Miệng cấp", "Supply grille", re.sub(r"^Miệng", "Grille", g_txt))
        raw = [ws.cell(r, c).value for c in (4, 5, 6)]
        fac = [_formula_offset(ws.cell(r, c).value) if isinstance(ws.cell(r, c).value, str) else None for c in (7, 8, 9)]
        vals = [Decimal(str(v)) * (f if f is not None else Decimal("1")) for v, f in zip(raw, fac)]
        avg = sum(vals) / len(vals)
        vol = ws.cell(r, 13).value
        flows.append({"room": rr["name_vn"], "g_vn": g_vn, "g_en": g_en, "avg": avg,
                      "vol": _eval_simple(vol) if vol not in (None, "") else None,
                      "ach_spec": parse_spec(ws.cell(r, c_spec_ach).value) if c_spec_ach else None})
    out["flows"] = flows

    # --- Chenh ap
    ws = wb["Chênh áp"]
    c_v1 = _header_col(ws, "Vị trí 1")
    c_v2 = _header_col(ws, "Vị trí 2")
    c_kh = _header_col(ws, "Kí hiệu", "Ký hiệu")
    c_sp = _header_col(ws, "Mức đánh giá", "Mức quy định")
    c_m1 = _header_col(ws, "Giá trị đo 1")
    pairs = []
    for r in range(2, ws.max_row + 1):
        kh = ws.cell(r, c_kh).value
        if kh in (None, ""):
            continue
        vals = [Decimal(str(ws.cell(r, c_m1 + i).value)) for i in range(3)]
        pairs.append({"kh": str(kh).strip(), "v1": str(ws.cell(r, c_v1).value).strip(),
                      "v2": str(ws.cell(r, c_v2).value).strip(), "avg": sum(vals) / 3,
                      "spec": parse_spec(_merged_value(ws, r, c_sp)) if c_sp else None})
    out["pairs"] = pairs

    # So hieu phieu tinh toan (header trai cua tung sheet, vd 'Số: Q326D132115/Chênh áp')
    out["sheet_no"] = {}
    for w in wb.worksheets:
        left = (w.oddHeader.left.text or "").strip()
        out["sheet_no"][w.title] = re.sub(r"^Số:\s*", "", left) or CODE
    return out


_X = _load_excel()

# Cap sach tung phong
_cls_all = set()
for _r in ROOMS:
    _c = _X["classes"].get(_r["name_vn"])
    if not _c:
        raise ValueError(f"Sheet Bụi 0,3: chua co 'Cấp sạch yêu cầu' cho {_r['name_vn']}")
    if len(_c) > 1:
        raise ValueError(f"Sheet Bụi 0,3: {_r['name_vn']} co nhieu cap sach {_c}")
    _r["iso_class"] = next(iter(_c))
    _r["class"] = f"ISO Class {_r['iso_class']}"
    _cls_all.add(_r["iso_class"])
if len(_cls_all) != 1:
    raise ValueError(f"Cac phong co cap sach khac nhau {_cls_all}: bang TBKQ hien chi ho tro 1 cap sach chung")
ISO_CLASS = next(iter(_cls_all))
ISO8_LIMITS = ISO_CLASS_LIMITS[ISO_CLASS]       # ten giu nguyen cho generate_all_reports (gioi han cua cap sach chung)

# Nhiet do / do am: moi diem 3 lan do ((T1,T2,T3),(RH1,RH2,RH3)), thu tu diem trong phong nhu sheet
TEMP_RH_READINGS = {r["name_vn"]: _X["trh"].get(r["name_vn"], []) for r in ROOMS}
TEMP_RH_SPECS = {r["name_vn"]: _X["trh_specs"].get(r["name_vn"], []) for r in ROOMS}
TEMP_CAL_OFFSET = _X["t_off"]
RH_CAL_OFFSET = _X["h_off"]
for _name, _pts in TEMP_RH_READINGS.items():
    if not _pts:
        raise ValueError(f"Sheet Nhiệt Ẩm: khong co diem do cho {_name}")

# Luu luong: gan mieng gio vao phong
for _r in ROOMS:
    _g = [f for f in _X["flows"] if f["room"] == _r["name_vn"]]
    if not _g:
        continue
    _tot = sum(f["avg"] for f in _g)
    _vol = next((f["vol"] for f in _g if f["vol"] is not None), None)
    if _vol is None:
        raise ValueError(f"Sheet Lưu lượng: thieu the tich phong {_r['name_vn']}")
    _ach = _tot / _vol
    _r["grilles"] = [{"name_vn": f["g_vn"], "name_en": f["g_en"], "flow": f["avg"]} for f in _g]
    _r["flow_tot"] = _tot
    _r["vol"] = _vol
    _r["ach"] = _ach
    _r["ach_spec"] = next((f["ach_spec"] for f in _g if f["ach_spec"] is not None), None)


def _place(txt):
    rr = _room_of(txt)
    if rr:
        return rr["name_vn"], rr["name_en"]
    n = _norm(txt)
    if n.startswith("môi trường"):
        return "Môi trường", "Ambient"
    if n.startswith("bên ngoài"):
        return "Bên ngoài", "Outside"
    return txt, txt


# Chenh ap
PRESSURE_PAIRS = []
for _p in _X["pairs"]:
    _v1, _v1e = _place(_p["v1"])
    _v2, _v2e = _place(_p["v2"])
    PRESSURE_PAIRS.append({"kh": _p["kh"], "v1_vn": _v1, "v1_en": _v1e, "v2_vn": _v2, "v2_en": _v2e,
                           "avg": _p["avg"], "spec": _p["spec"]})

SHEET_NO = _X["sheet_no"]


def _room_spec(name):
    pts = TEMP_RH_SPECS[name]
    t = pts[0][0] if pts else None
    h = pts[0][1] if pts else None
    if any(p != pts[0] for p in pts):
        print(f"[Warning] {name}: muc danh gia nhiet/am khac nhau giua cac diem; dung muc cua diem dau cho cau chu")
    r = ROOM_BY_VN.get(name) or next(x for x in ROOMS if x["name_vn"] == name)
    return {"t": t, "rh": h, "ach": r.get("ach_spec")}


ROOM_SPEC = {r["name_vn"]: _room_spec(r["name_vn"]) for r in ROOMS}


def _q1(x):
    """Round a Decimal to 1 decimal place (half-up, as Excel does)."""
    return Decimal(x).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)


def temp_rh_points(room_vn):
    """Returns [(T_avg, RH_avg), ...] (calibrated, rounded to 0.1) of every point of a room."""
    out = []
    for temps, rhs in TEMP_RH_READINGS[room_vn]:
        t = sum(Decimal(str(v)) for v in temps) / 3 + TEMP_CAL_OFFSET
        h = sum(Decimal(str(v)) for v in rhs) / 3 + RH_CAL_OFFSET
        out.append((_q1(t), _q1(h)))
    return out


def temp_rh_room_avg(room_vn):
    """(T_tb, RH_tb) ca phong = trung binh moi lan do sau hieu chuan cua moi diem, lam tron 0,1 (bang AVERAGE
    cot 'Trung bình phòng' cua phieu tinh toan). Nguoi dung chot 30/10/2026: danh gia nhiet do, do am theo trung
    binh ca phong (khong theo tung diem)."""
    rows = TEMP_RH_READINGS[room_vn]
    temps = [Decimal(str(v)) for ts, _ in rows for v in ts]
    rhs = [Decimal(str(v)) for _, hs in rows for v in hs]
    return _q1(sum(temps) / len(temps) + TEMP_CAL_OFFSET), _q1(sum(rhs) / len(rhs) + RH_CAL_OFFSET)


def trh_room_fail(room_vn):
    """(nhiet do khong dat, do am khong dat) cua trung binh phong so voi muc quy dinh cua phong."""
    t, h = temp_rh_room_avg(room_vn)
    sp = ROOM_SPEC[room_vn]
    return (sp["t"] is not None and not in_spec(t, sp["t"]),
            sp["rh"] is not None and not in_spec(h, sp["rh"]))


def trh_spec_groups():
    """Nhom phong theo muc quy dinh nhiet/am: [((muc_t, muc_rh), [phong...]), ..., (None, [phong khong co muc])].
    Moi nhom = 1 bang ket qua rieng trong TBKQ (dong cuoi ghi muc cua nhom, roi Ghi chu, roi Danh gia);
    nhom khong co muc (None) dat cuoi, khong co dong muc va khong co cau Danh gia (nguoi dung chot 30/10/2026)."""
    groups, no_spec = [], []
    for r in ROOMS:
        sp = ROOM_SPEC[r["name_vn"]]
        if sp["t"] is None and sp["rh"] is None:
            no_spec.append(r)
            continue
        key = (sp["t"], sp["rh"])
        for g in groups:
            if g[0] == key:
                g[1].append(r)
                break
        else:
            groups.append((key, [r]))
    if no_spec:
        groups.append((None, no_spec))
    return groups


def fmt_dec(x, is_en):
    s = f"{_q1(x):.1f}"
    return s if is_en else s.replace(".", ",")


# --------------------------------------------------------------------------------------------
# Evaluation per room / per position.  status in {"pass", "fail", "na"}
# --------------------------------------------------------------------------------------------
def eval_dust():
    res = {}
    for r in ROOMS:
        fails = set()
        for sid in range(r["start_sample"], r["start_sample"] + r["points_count"]):
            c05, c10, c50 = PARTICLE_DATA[sid]
            if (c05 or 0) > ISO8_LIMITS["0.5"]:
                fails.add("0.5")
            if (c10 or 0) > ISO8_LIMITS["1"]:
                fails.add("1")
            if (c50 or 0) > ISO8_LIMITS["5"]:
                fails.add("5")
        res[r["name_vn"]] = {"status": "fail" if fails else "pass", "fails": sorted(fails, key=float)}
    return res


def eval_temp_rh():
    """Danh gia trung binh ca phong theo muc cua phong (cot Mức đánh giá nhiệt/ẩm cua sheet Nhiệt Ẩm)."""
    res = {}
    for r in ROOMS:
        name = r["name_vn"]
        sp = ROOM_SPEC[name]
        t_fail, h_fail = trh_room_fail(name)
        t_status = "na" if sp["t"] is None else ("fail" if t_fail else "pass")
        rh_status = "na" if sp["rh"] is None else ("fail" if h_fail else "pass")
        if "fail" in (t_status, rh_status):
            status = "fail"
        elif "pass" in (t_status, rh_status):
            status = "pass"
        else:
            status = "na"
        res[name] = {"status": status, "t": t_status, "rh": rh_status}
    return res


def dp_value(p):
    return _q1(p["avg"])


def eval_dp():
    res = {}
    for p in PRESSURE_PAIRS:
        meas = dp_value(p)
        spec = p["spec"]
        status = "na" if spec is None else ("pass" if in_spec(meas, spec) else "fail")
        res[p["kh"]] = {
            "status": status, "measured": meas, "spec": spec,
            "from_vn": p["v1_vn"], "from_en": p["v1_en"], "to_vn": p["v2_vn"], "to_en": p["v2_en"],
        }
    return res


def fmt_flow(x):
    """Luu luong m3/h: so nguyen, tach hang nghin bang dau cach (giong nhau o ban Viet/Anh)."""
    v = int(Decimal(x).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    return f"{v:,}".replace(",", " ")


def ach_value(r):
    return _q1(r["ach"])


def eval_ach():
    res = {}
    for r in ROOMS:
        if not r.get("grilles"):
            continue
        spec = r.get("ach_spec")
        ach = ach_value(r)
        if spec is None:
            res[r["name_vn"]] = {"status": "na", "ach": ach}
        else:
            res[r["name_vn"]] = {"status": "pass" if in_spec(ach, spec) else "fail", "ach": ach}
    return res


# --------------------------------------------------------------------------------------------
# Assessment sentences (areas are split into compliant / non-compliant groups, one sentence at the
# bottom of every result section). Keywords between *** are rendered bold+italic by the styler.
# --------------------------------------------------------------------------------------------
def _quote(name):
    return f"“{name}”"


def _join(items, is_en):
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + (" and " if is_en else " và ") + items[-1]


def _areas(rooms, is_en, total, notes=None, label_all=True):
    """Phrase for a group of rooms ('the areas “A”, “B” and “C”'); notes = {vn_name: (note_vn, note_en)}."""
    notes = notes or {}
    if label_all and len(rooms) == total and not notes:
        return "the measurement areas" if is_en else "các khu vực thực hiện"
    names = []
    for r in rooms:
        n = _quote(r["name_en"] if is_en else r["name_vn"])
        if r["name_vn"] in notes:
            n += f" ({notes[r['name_vn']][1 if is_en else 0]})"
        names.append(n)
    if is_en:
        return ("the area " if len(names) == 1 else "the areas ") + _join(names, True)
    return ("khu vực " if len(names) == 1 else "các khu vực ") + _join(names, False)


def _positions(labels):
    """['AP1','AP2','AP5','AP6','AP7'] -> 'AP1, AP2, AP5 – AP7' (consecutive runs of >= 3 are collapsed)."""
    nums = [int(re.sub(r"\D", "", s)) for s in labels]
    prefix = re.sub(r"\d", "", labels[0]) if labels else ""
    out, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        if j - i >= 2:
            out.append(f"{prefix}{nums[i]} – {prefix}{nums[j]}")
        else:
            out.extend(f"{prefix}{n}" for n in nums[i:j + 1])
        i = j + 1
    return ", ".join(out)


def _not_evaluated_en(rooms, total):
    phrase = _areas(rooms, True, total, None, False)          # 'the area(s) “A” and “B”'
    phrase = phrase[0].upper() + phrase[1:]
    return phrase + (" was not evaluated" if len(rooms) == 1 else " were not evaluated")


def _paras(prefix, ok_s, bad_s, tail=""):
    """Assessment paragraphs. When the same indicator has BOTH compliant and non-compliant areas it is
    split in two paragraphs: the compliant areas on top, the non-compliant areas below.
    `prefix` ('Đánh giá: ' / 'Assessment: ') goes only in front of the first paragraph; `tail` (areas that
    are not evaluated) is appended to the last one."""
    # Chi tieu/khu vuc KHONG danh gia (khong co muc quy dinh): khong ghi cau "Khong danh gia..." (yeu cau nguoi dung
    # 26/10/2026); neu ca chi tieu khong danh gia thi tra ve [] -> bo han doan Danh gia duoi bang.
    paras = [x for x in (ok_s, bad_s) if x]
    if not paras:
        return []
    paras[0] = prefix + paras[0]
    return paras


def assess_dust(is_en):
    res = eval_dust()
    ok = [r for r in ROOMS if res[r["name_vn"]]["status"] == "pass"]
    bad = [r for r in ROOMS if res[r["name_vn"]]["status"] == "fail"]
    n = len(ROOMS)
    if is_en:
        crit = f"the specified limits for ISO Class {ISO_CLASS} cleanroom according to ISO 14644-1:2015"
        lead = "The airborne particle concentration results"
        ok_s = f"{lead} at {_areas(ok, True, n)} ***complied*** with {crit}." if ok else ""
        bad_s = f"{lead} at {_areas(bad, True, n)} ***did not comply*** with {crit}." if bad else ""
        return _paras("Assessment: ", ok_s, bad_s)
    crit = f"mức quy định đối với khu vực cấp sạch ISO {ISO_CLASS} theo ISO 14644-1:2015"
    lead = "Kết quả đo nồng độ hạt bụi trong không khí"
    ok_s = f"{lead} tại {_areas(ok, False, n)} ***phù hợp*** với {crit}." if ok else ""
    bad_s = f"{lead} tại {_areas(bad, False, n)} ***không phù hợp*** với {crit}." if bad else ""
    return _paras("Đánh giá: ", ok_s, bad_s)


def assess_temp_rh(is_en, rooms=None):
    """rooms: chi danh gia cac phong nay (1 nhom muc = 1 bang trong TBKQ); mac dinh moi phong (BCKQ)."""
    rooms = rooms or ROOMS
    res = eval_temp_rh()
    ok = [r for r in rooms if res[r["name_vn"]]["status"] == "pass"]
    bad = [r for r in rooms if res[r["name_vn"]]["status"] == "fail"]
    na = [r for r in rooms if res[r["name_vn"]]["status"] == "na"]
    n = len(rooms)
    ok_notes, bad_notes = {}, {}
    for r in ok:
        e = res[r["name_vn"]]
        if e["rh"] == "na":
            ok_notes[r["name_vn"]] = ("chỉ đánh giá nhiệt độ", "temperature only")
        elif e["t"] == "na":
            ok_notes[r["name_vn"]] = ("chỉ đánh giá độ ẩm", "humidity only")
    for r in bad:
        e = res[r["name_vn"]]
        vn, en = [], []
        if e["t"] == "fail":
            vn.append("nhiệt độ"); en.append("temperature")
        if e["rh"] == "fail":
            vn.append("độ ẩm"); en.append("humidity")
        bad_notes[r["name_vn"]] = (", ".join(vn), " and ".join(en))
    if is_en:
        crit = "the corresponding specifications provided by Lam Quang Dai Company Limited"
        lead = "The temperature and humidity test results"
        ok_s = f"{lead} at {_areas(ok, True, n, ok_notes, False)} ***complied*** with {crit}." if ok else ""
        bad_s = f"{lead} at {_areas(bad, True, n, bad_notes, False)} ***did not comply*** with {crit}." if bad else ""
        tail = f"{_not_evaluated_en(na, n)} because no temperature/humidity specification is given." if na else ""
        if na and len(na) == n:
            tail = "The temperature and humidity results were not evaluated because no specification is given."
        return _paras("Assessment: ", ok_s, bad_s, tail)
    crit = "mức quy định tương ứng được cung cấp bởi Công ty TNHH Lâm Quang Đại"
    lead = "Kết quả đo nhiệt độ, độ ẩm"
    ok_s = f"{lead} tại {_areas(ok, False, n, ok_notes, False)} ***phù hợp*** với {crit}." if ok else ""
    bad_s = f"{lead} tại {_areas(bad, False, n, bad_notes, False)} ***không phù hợp*** với {crit}." if bad else ""
    tail = f"Không đánh giá đối với {_areas(na, False, n, None, False)} do không có mức quy định về nhiệt độ, độ ẩm." if na else ""
    if na and len(na) == n:
        tail = "Không đánh giá kết quả đo nhiệt độ, độ ẩm do không có mức quy định."
    return _paras("Đánh giá: ", ok_s, bad_s, tail)


def assess_dp(is_en):
    res = eval_dp()
    ok = [k for k, v in res.items() if v["status"] == "pass"]
    bad = [k for k, v in res.items() if v["status"] == "fail"]
    na = [k for k, v in res.items() if v["status"] == "na"]
    if is_en:
        crit = "the corresponding specifications provided by Lam Quang Dai Company Limited"
        lead = "The differential pressure test results"
        ok_s = f"{lead} at {'locations' if len(ok) > 1 else 'location'} {_positions(ok)} ***complied*** with {crit}." if ok else ""
        bad_s = f"{lead} at {'locations' if len(bad) > 1 else 'location'} {_positions(bad)} ***did not comply*** with {crit}." if bad else ""
        if na and len(na) == len(res):
            tail = "The differential pressure results were not evaluated because no specification is given."
        else:
            tail = (f"{'Locations' if len(na) > 1 else 'Location'} {_positions(na)} {'were' if len(na) > 1 else 'was'} "
                    f"not evaluated because no specification is given.") if na else ""
        return _paras("Assessment: ", ok_s, bad_s, tail)
    crit = "mức quy định tương ứng được cung cấp bởi Công ty TNHH Lâm Quang Đại"
    lead = "Kết quả đo chênh lệch áp suất"
    ok_s = f"{lead} tại {'các vị trí' if len(ok) > 1 else 'vị trí'} {_positions(ok)} ***phù hợp*** với {crit}." if ok else ""
    bad_s = f"{lead} tại {'các vị trí' if len(bad) > 1 else 'vị trí'} {_positions(bad)} ***không phù hợp*** với {crit}." if bad else ""
    if na and len(na) == len(res):
        tail = "Không đánh giá kết quả đo chênh lệch áp suất do không có mức quy định."
    else:
        tail = (f"Không đánh giá đối với {'các vị trí' if len(na) > 1 else 'vị trí'} {_positions(na)} "
                f"do không có mức quy định.") if na else ""
    return _paras("Đánh giá: ", ok_s, bad_s, tail)


def assess_ach(is_en):
    res = eval_ach()
    flow_rooms = [r for r in ROOMS if r["name_vn"] in res]
    ok = [r for r in flow_rooms if res[r["name_vn"]]["status"] == "pass"]
    bad = [r for r in flow_rooms if res[r["name_vn"]]["status"] == "fail"]
    na = [r for r in flow_rooms if res[r["name_vn"]]["status"] == "na"]
    n = len(flow_rooms)
    if is_en:
        crit = "the technical specifications provided by Lam Quang Dai Company Limited"
        lead = "The calculation results of supply airflow rate and air change rate"
        ok_s = f"{lead} at {_areas(ok, True, n)} ***complied*** with {crit}." if ok else ""
        bad_s = (f"{lead} at {_areas(bad, True, n)} ***did not comply*** with {crit} "
                 f"(air change rate outside the specified range).") if bad else ""
        tail = f"{_not_evaluated_en(na, n)} because no air change rate specification is given." if na else ""
        if na and len(na) == n:
            tail = "The air change rate results were not evaluated because no specification is given."
        return _paras("Assessment: ", ok_s, bad_s, tail)
    crit = "thông số kỹ thuật được cung cấp bởi Công ty TNHH Lâm Quang Đại"
    lead = "Kết quả tính toán lưu lượng khí cấp và số lần trao đổi khí"
    ok_s = f"{lead} tại {_areas(ok, False, n)} ***phù hợp*** với {crit}." if ok else ""
    bad_s = (f"{lead} tại {_areas(bad, False, n)} ***không phù hợp*** với {crit} "
             f"(số lần trao đổi khí nằm ngoài mức quy định).") if bad else ""
    tail = f"Không đánh giá đối với {_areas(na, False, n, None, False)} do không có mức quy định về số lần trao đổi khí." if na else ""
    if na and len(na) == n:
        tail = "Không đánh giá kết quả tính toán lưu lượng khí cấp và số lần trao đổi khí do không có mức quy định."
    return _paras("Đánh giá: ", ok_s, bad_s, tail)


# --------------------------------------------------------------------------------------------
# Per-room summary used by the internal report (BCKQ) summary table: passed / failed indicators.
# --------------------------------------------------------------------------------------------
def points_detail(room_vn, is_en):
    """Cot 'So luong vi tri quan trac' (bang 9.1 TBKQ / bang khu vuc BCKQ): liet ke MOI chi tieu va diem do cua phong.
    Chenh ap: diem AP duoc ghi cho ca 2 phong cua cap (vi tri do nam giua 2 phong)."""
    r = ROOM_BY_VN[room_vn]
    lines = []
    n = r["points_count"]
    if is_en:
        lines.append(f"- Airborne particles: {n:02d} points (1 – {n})")
    else:
        lines.append(f"- Nồng độ hạt bụi: {n:02d} vị trí (1 – {n})")
    # Nhiet do, do am: NA danh so lien tuc toan ho so theo thu tu ROOMS (giong bang 9.3)
    start = 1
    for rr in ROOMS:
        if rr["name_vn"] == room_vn:
            break
        start += len(TEMP_RH_READINGS[rr["name_vn"]])
    k = len(TEMP_RH_READINGS[room_vn])
    if k:
        na = f"NA{start}" if k == 1 else f"NA{start} – NA{start + k - 1}"
        lines.append(f"- Temperature, humidity: {na}" if is_en else f"- Nhiệt độ, độ ẩm: {na}")
    # Chenh ap: ghi theo cap phong 'tu phong -> den phong' (giong bang 9.4), moi diem AP 1 dong
    aps = [p for p in PRESSURE_PAIRS if room_vn in (p["v1_vn"], p["v2_vn"])]
    if aps:
        lines.append("- Differential pressure:" if is_en else "- Chênh áp:")
        for p in aps:
            a_, b_ = (p["v1_en"], p["v2_en"]) if is_en else (p["v1_vn"], p["v2_vn"])
            lines.append(f"   {p['kh']}: {a_} → {b_}")
    grilles = r.get("grilles") or []
    if grilles:
        codes = [m.group(1) for g in grilles for m in [re.search(r"\((H\d+)\)", g["name_vn"])] if m]
        code_txt = f" ({_positions(codes)})" if codes and len(codes) == len(grilles) else ""
        g = len(grilles)
        if is_en:
            lines.append(f"- Supply airflow: {g} grille{'s' if g > 1 else ''}{code_txt}")
        else:
            lines.append(f"- Lưu lượng khí cấp: {g} miệng cấp{code_txt}")
    return lines


def room_summary():
    """Bang tom tat BCKQ theo phong: chi tieu dat / khong dat / khong danh gia.
    Dung CHUNG ham eval_* voi TBKQ nen ket luan BCKQ luon khop TBKQ (AGENTS.md muc 7)."""
    dust, trh, dp, ach = eval_dust(), eval_temp_rh(), eval_dp(), eval_ach()
    out = {}
    for r in ROOMS:
        name = r["name_vn"]
        spec = ROOM_SPEC[name]
        good, bad, na = [], [], []
        basis_customer = False

        d = dust[name]
        if d["status"] == "pass":
            good.append("Nồng độ hạt bụi")
        else:
            sizes = ", ".join(f"≥ {s} µm".replace(".", ",") for s in d["fails"])
            bad.append(f"Nồng độ hạt bụi (vượt giới hạn ISO {ISO_CLASS} ở cỡ hạt {sizes})")

        t = trh[name]
        t_avg, h_avg = temp_rh_room_avg(name)
        if t["t"] == "pass":
            good.append("Nhiệt độ"); basis_customer = True
        elif t["t"] == "fail":
            bad.append(f"Nhiệt độ (trung bình phòng {fmt_dec(t_avg, False)} °C; mức quy định {spec_txt(spec['t'])} °C)")
            basis_customer = True
        if t["rh"] == "pass":
            good.append("Độ ẩm"); basis_customer = True
        elif t["rh"] == "fail":
            bad.append(f"Độ ẩm (trung bình phòng {fmt_dec(h_avg, False)} %RH; mức quy định {spec_txt(spec['rh'])} %RH)")
            basis_customer = True
        if t["t"] == "na" and t["rh"] == "na":
            na.append("Nhiệt độ, độ ẩm")
        elif t["t"] == "na":
            na.append("Nhiệt độ")
        elif t["rh"] == "na":
            na.append("Độ ẩm")

        aps = [(k, v) for k, v in dp.items() if v["from_vn"] == name]
        if aps:
            oks = [k for k, v in aps if v["status"] == "pass"]
            fails = [f"{k}: {fmt_dec(v['measured'], False)} Pa; mức quy định {spec_txt(v['spec'])} Pa" for k, v in aps if v["status"] == "fail"]
            nas = [k for k, v in aps if v["status"] == "na"]
            if oks or fails:
                basis_customer = True
            if oks:
                good.append("Chênh áp (" + ", ".join(oks) + ")")
            if fails:
                bad.append("Chênh áp (" + "; ".join(fails) + ")")
            if nas:
                na.append("Chênh áp (" + ", ".join(nas) + ")")

        if name in ach:
            a = ach[name]
            if a["status"] == "pass":
                good.append("Số lần trao đổi khí"); basis_customer = True
            elif a["status"] == "fail":
                bad.append(f"Số lần trao đổi khí ({fmt_dec(a['ach'], False)} ACH; mức quy định {spec_txt(spec['ach'])} ACH)")
                basis_customer = True
            else:
                na.append("Số lần trao đổi khí")

        basis = f"ISO 14644-1:2015 (cấp ISO {ISO_CLASS})"
        if basis_customer:
            basis += "; mức quy định do Công ty TNHH Lâm Quang Đại cung cấp"
        out[name] = {"good": good, "bad": bad, "na": na, "basis": basis}
    return out


if __name__ == "__main__":
    for fn in (assess_dust, assess_temp_rh, assess_dp, assess_ach):
        print("\n".join(fn(False))); print("\n".join(fn(True))); print()
    for k, v in room_summary().items():
        print(k, v)
