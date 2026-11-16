# -*- coding: utf-8 -*-
"""
Phiếu tính toán phòng sạch (`<mã> - TinhToanPhongSach.xlsx`) — bộ dựng chung theo form của người dùng
(người dùng chốt 07/11/2026; form mẫu `1097 - TinhToanPhongSach.xlsx`, `078038 - TinhPhongSach.xlsx`):

    Bụi 0,3    bảng xuất AeroTrak 9350 (chế độ Σ, hạt/m³) + cột Phòng / Vị trí (Zero, 1…n)
    Bụi 0,1    database Lasair III 110 (số đếm vi phân) + công thức hạt/m³ vi phân và cộng dồn
    Nhiệt Ẩm   3 lần đọc nhiệt độ, độ ẩm -> sau hiệu chính -> trung bình
    Độ ồn      LAeq từng lần -> (sau hiệu chính) -> trung bình           (sheet riêng — người dùng chốt 07/11/2026)
    Lưu lượng  lưu lượng từng lần -> HC -> trung bình -> lưu lượng HEPA -> tổng phòng, thể tích, ACH
    Chênh áp   Vị trí 1 → Vị trí 2, 3 lần đo -> HC -> trung bình
    Scan lọc   Media, Gasket -> Max pen %

Khác form mẫu (người dùng chốt 07/11/2026): **giữ cột Mức quy định + Đánh giá** (công thức IF); **hiệu chỉnh kết quả
theo giấy hiệu chuẩn** cho mọi đại lượng có bảng hiệu chuẩn (nhiệt độ, độ ẩm, chênh áp, lưu lượng, độ ồn): giá trị
đúng = chỉ thị − sai số, sai số nội suy tuyến tính giữa 2 điểm kẹp hai bên (FORECAST), ngoài khoảng lấy sai số điểm
biên, không ngoại suy (AGENTS mục 8C, 15). Bảng hiệu chuẩn nằm ở sheet ẩn `HieuChuan` (không in, công thức vẫn động —
AGENTS mục 11). Không có bảng hiệu chuẩn cho thiết bị -> cột "sau hiệu chính" = giá trị đọc và ghi vào `notes`.

Phiếu chỉ gồm bảng số liệu (AGENTS mục 11): không dòng khách hàng / thiết bị / ghi chú.

    from Core.Tools.cleanroom_calc import CleanroomCalcBook, CalTable
    book = CleanroomCalcBook()
    cal_t = CalTable.from_library("MT33-3", "nhiệt độ", "29/10/2026", device="Testo 625 – máy 3")
    book.temp_rh([...], cal_t, cal_rh, so="Q326D138143/Nhiệt độ, độ ẩm", date="29/10/2026")
    book.save(path)
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Iterable, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter as L

from .excel_borders import thin_outside_borders

FONT = "Times New Roman"
SIZE = 11
THIN = Side(style="thin")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
CAL_SHEET = "HieuChuan"

# Tên sheet và tiêu đề phiếu (header giữa) — số phiếu "<mã>/<đuôi>" do script hồ sơ truyền vào (AGENTS mục 7)
TITLES = {
    "Bụi 0,3": "TÍNH TOÁN KẾT QUẢ NỒNG ĐỘ HẠT BỤI",
    "Bụi 0,1": "TÍNH TOÁN KẾT QUẢ NỒNG ĐỘ HẠT BỤI",
    "Nhiệt Ẩm": "TÍNH TOÁN KẾT QUẢ NHIỆT ĐỘ, ĐỘ ẨM",
    "Độ ồn": "TÍNH TOÁN KẾT QUẢ ĐỘ ỒN",
    "Lưu lượng": "TÍNH TOÁN VẬN TỐC GIÓ VÀ LƯU LƯỢNG QUA LỌC HEPA",
    "Chênh áp": "KẾT QUẢ ĐO CHÊNH LỆCH ÁP SUẤT",
    "Scan lọc": "TÍNH TOÁN KẾT QUẢ RÒ RỈ LỌC HEPA",
}
# Đuôi số phiếu mặc định (AGENTS mục 7)
SUFFIX = {"bui": "Bụi", "nhiet_am": "Nhiệt độ, độ ẩm", "do_on": "Độ ồn", "chenh_ap": "Chênh áp",
          "luu_luong": "Lưu lượng", "scan": "Scan lọc"}


# ---------------------------------------------------------------------------
# Giấy hiệu chuẩn
# ---------------------------------------------------------------------------
# CÔNG TẮC CHUNG — áp hiệu chỉnh theo giấy hiệu chuẩn cho mọi phiếu tính toán (tủ BSC/LAF/tủ cấy, phòng sạch).
# Người dùng chốt 09/11/2026: "Tạm thời không tính hiệu chuẩn, nhưng vẫn giữ khung chương trình sẵn sàng để áp
# hiệu chuẩn" -> False: CalTable.from_library trả None -> phiếu tủ 1 phần (không chèn phần b)), phiếu phòng sạch
# cột "sau hiệu chính" = ô đọc, số TBKQ/BCKQ = số đọc. Toàn bộ code hiệu chỉnh (thư viện HieuChuan/, CalTable,
# bsc_cal_blocks.finalize_bsc_workbook, sheet ẩn HieuChuan, công thức FORECAST) GIỮ NGUYÊN.
# Bật lại: đổi thành True (hoặc đặt biến môi trường LAMKQ_HIEU_CHINH=1) rồi dựng lại hồ sơ.
AP_HIEU_CHINH = os.environ.get("LAMKQ_HIEU_CHINH", "0").strip() in ("1", "true", "True")
_HC_OFF_SEEN: set = set()


@dataclass
class CalTable:
    """Bảng hiệu chuẩn một đại lượng: points = [(chỉ thị, chuẩn, sai số)] tăng dần theo chỉ thị."""
    device: str
    quantity: str
    unit: str
    points: List[Tuple[float, Optional[float], float]]
    cert_no: str = ""

    @classmethod
    def from_library(cls, code: str, quantity: str, on_date: str, device: str = "", unit: str = "",
                     serial: Optional[str] = None, fallback: Optional[Sequence] = None) -> Optional["CalTable"]:
        """Lấy từ thư viện `HieuChuan/` (giấy còn hiệu lực ngày đo). Không có -> `fallback` (bảng script) hoặc None.
        Có cả hai mà khác nhau -> in [WARNING] [Calibration] và dùng bảng thư viện.
        AP_HIEU_CHINH = False -> trả None (tạm không hiệu chỉnh), phần code bên dưới giữ nguyên để bật lại."""
        if not AP_HIEU_CHINH:
            if (code, quantity) not in _HC_OFF_SEEN:
                _HC_OFF_SEEN.add((code, quantity))
                print(f"[Calibration] Tạm không hiệu chỉnh theo giấy hiệu chuẩn ({code} {quantity}) — "
                      f"AP_HIEU_CHINH = False (Core/Tools/cleanroom_calc.py)")
            return None
        try:
            from .calib_library import cal_table, valid_cert
            pts = cal_table(code, quantity, on_date, serial=serial)
            cert = valid_cert(code, on_date, serial=serial, quantity=quantity) or {}
            pts = [(float(a), None if b is None else float(b), float(c)) for a, b, c in pts]
            if fallback and [tuple(round(x, 3) for x in (p[0], p[2])) for p in fallback] != \
                    [tuple(round(x, 3) for x in (p[0], p[2])) for p in pts]:
                print(f"[WARNING] [Calibration] {code} {quantity}: bảng trong script khác thư viện "
                      f"({cert.get('cert_no', '')}) — dùng bảng thư viện")
            return cls(device or code, quantity, unit, pts, cert.get("cert_no", "") or "")
        except Exception as e:      # noqa: BLE001 — thư viện thiếu/không có giấy
            if fallback:
                print(f"[WARNING] [Calibration] {code} {quantity}: {e} — dùng bảng trong script")
                return cls(device or code, quantity, unit, [tuple(p) for p in fallback])
            print(f"[WARNING] [Calibration] {code} {quantity}: {e} — không hiệu chính")
            return None

    def error_at(self, x: float) -> float:
        p = self.points
        if x <= p[0][0]:
            return p[0][2]
        if x >= p[-1][0]:
            return p[-1][2]
        for (x1, _, e1), (x2, _, e2) in zip(p, p[1:]):
            if x1 <= x <= x2:
                return e1 if x2 == x1 else e1 + (e2 - e1) * (x - x1) / (x2 - x1)
        return p[-1][2]

    def corrected(self, x: float) -> float:
        return x - self.error_at(x)


def apply_cal(x: float, cal: Optional[CalTable]) -> float:
    return x if cal is None else cal.corrected(x)


def round_half_up(x: float, nd: int = 1) -> float:
    from decimal import Decimal, ROUND_HALF_UP
    q = Decimal(1).scaleb(-nd)
    return float(Decimal(str(x)).quantize(q, rounding=ROUND_HALF_UP))


# ---------------------------------------------------------------------------
# Đọc file máy
# ---------------------------------------------------------------------------
AEROTRAK_COLS = ["#", "Zone", "Location", "Date/Time (dd/mm/yyyy)", "Device Status", "Sample Time",
                 "Ch1", "Ch2", "Ch3", "Ch4", "Ch5", "Ch6", "Flow Status", "Laser Status", "Vol.", "Temp",
                 "Humidity", "Velocity"]


def read_aerotrak(path: str, check_mode: bool = True):
    """6ChannelFixed.xlsx (AeroTrak 9350, chế độ Σ, # / m³) -> (nhãn cột {key: tiêu đề gốc}, {#: {key: giá trị}}).
    Phòng sạch chép nguyên số máy nên file PHẢI là `Count Mode: Σ`, `Units: # / m³` (AGENTS mục 8C) — sai thì dừng
    (`particle_mode.require_mode`); khí nén đọc bằng `particle_mode.aerotrak_values(..., want=DIFF)`."""
    if check_mode:
        from .particle_mode import require_mode, check_cumul_shape, CUMUL
        require_mode(path, CUMUL, "bảng bụi phòng sạch (ISO 14644-1)")
    ws = openpyxl.load_workbook(path, data_only=True).active
    hdr, labels, out = None, {}, {}
    for row in ws.iter_rows(values_only=True):
        vals = list(row)
        if hdr is None:
            if "#" in vals and "Location" in vals:
                hdr = {}
                for i, v in enumerate(vals):
                    if v in (None, ""):
                        continue
                    key = str(v).split("\n")[0].strip()
                    hdr[key] = i
                    labels[key] = str(v)
            continue
        n = vals[hdr["#"]] if "#" in hdr else None
        if isinstance(n, (int, float)):
            out[int(n)] = {k: vals[i] for k, i in hdr.items()}
    if hdr is None:
        raise ValueError(f"Không thấy dòng tiêu đề AeroTrak trong {path}")
    if check_mode:
        for n, rec in out.items():
            check_cumul_shape([rec.get(f"Ch{i}") for i in range(1, 7)], f"{os.path.basename(path)} #{n}")
    return labels, out


LASAIR_COLS = ["Data ID", "Sample Date", "Sample Time", "Sensor Name", "Sensor Type", "S/N", "Statistics",
               "Location Changed", "Location", "Manifold", "Laser OK", "Flow OK", "Coincidence Error <10",
               "Sample Valid", "Sample Interval", "Volume Units", "Sample Volume", "Number of Channels"] + \
              [x for i in range(1, 9) for x in (f"Size {i}", f"Counts {i}")] + ["Number Analogs"]


def _num(v):
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return v


def read_lasair_csv(path: str):
    """database CSV Lasair -> {Data ID: {cột: giá trị}}; CSV không có cột Data ID thì Data ID = số thứ tự dòng."""
    import csv
    out = {}
    with open(path, encoding="utf-8-sig") as fh:
        for i, r in enumerate(csv.DictReader(fh), 1):
            rec = {k.strip(): v for k, v in r.items() if k}
            did = int(_num(rec.get("Data ID") or i))
            rec["Data ID"] = did
            out[did] = rec
    return out


def read_lasair_sheet(path: str, sheet: str = "Bụi 0,1"):
    """Sheet database Lasair (dạng form 'Bụi 0,1' của người dùng) -> {Data ID: {cột: giá trị}}."""
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    hdr = {str(c.value).strip(): c.column - 1 for c in ws[1] if c.value}
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        did = row[hdr["Data ID"]] if "Data ID" in hdr else None
        if isinstance(did, (int, float)):
            out[int(did)] = {k: row[i] for k, i in hdr.items() if k in LASAIR_COLS}
    return out


def lasair_conc(rec: dict) -> List[float]:
    """[≥ Size1 … ≥ Size8] hạt/m³ cộng dồn = Σ counts từ kênh đó trở lên / Sample Volume (m³)."""
    vol = float(_num(rec["Sample Volume"]))
    c = [float(_num(rec[f"Counts {i}"])) for i in range(1, 9)]
    return [sum(c[k:]) / vol for k in range(8)]


# ---------------------------------------------------------------------------
# Sổ tính toán
# ---------------------------------------------------------------------------
def _fmt_num(v):
    return v if isinstance(v, (int, float)) else v


def _limit_cond(ref: str, lim) -> Optional[str]:
    """'≤ 25' / '≥ 5' / '20 - 25' / số (≤) -> điều kiện Excel cho ô `ref`; None khi không có mức."""
    if lim is None:
        return None
    if isinstance(lim, (int, float)):
        return f"{ref}<={lim}"
    s = str(lim).strip().replace(",", ".")
    if not s or s.upper() in ("N/A", "-"):
        return None
    m = re.match(r"^(-?[\d.]+)\s*[-–]\s*(-?[\d.]+)$", s)
    if m:
        return f"AND({ref}>={m.group(1)},{ref}<={m.group(2)})"
    m = re.match(r"^(≤|<=|<)\s*(-?[\d.]+)$", s)
    if m:
        return f"{ref}{'<' if m.group(1) == '<' else '<='}{m.group(2)}"
    m = re.match(r"^(≥|>=|>)\s*(-?[\d.]+)$", s)
    if m:
        return f"{ref}{'>' if m.group(1) == '>' else '>='}{m.group(2)}"
    return None


def _eval_formula(conds: Iterable[Optional[str]]) -> str:
    cs = [c for c in conds if c]
    if not cs:
        return "-"
    body = cs[0] if len(cs) == 1 else f"AND({','.join(cs)})"
    return f'=IF({body},"Đạt","Không đạt")'


class CleanroomCalcBook:
    def __init__(self):
        self.wb = openpyxl.Workbook()
        self._first = True
        self._cal_ws = None
        self._cal_next = 1
        self._cal_blocks = {}
        self.notes: List[str] = []          # điểm cần nêu trong báo cáo cuối (không in lên phiếu)

    # ---------------- tiện ích ô ----------------
    @staticmethod
    def put(ws, r, c, value, bold=False, align=CENTER, fmt=None, grid=True, italic=False):
        cell = ws.cell(r, c)
        cell.value = value
        cell.font = Font(name=FONT, size=SIZE, bold=bold, italic=italic)
        cell.alignment = align
        if grid:
            cell.border = GRID
        if fmt:
            cell.number_format = fmt
        return cell

    def merge(self, ws, r0, c0, r1, c1, value, **kw):
        self.put(ws, r0, c0, value, **kw)
        if (r0, c0) != (r1, c1):
            ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                ws.cell(r, c).border = GRID
        return ws.cell(r0, c0)

    def sheet(self, name: str, so: str, date: str, widths: Sequence[float], title_rows: str = "1:1",
              title: Optional[str] = None):
        ws = self.wb.active if self._first else self.wb.create_sheet()
        self._first = False
        ws.title = name
        ws.oddHeader.left.text = f"Số: {so}"
        ws.oddHeader.center.text = title or TITLES.get(name, name.upper())
        ws.oddHeader.right.text = f"Ngày {date}"
        for h in (ws.oddHeader.left, ws.oddHeader.center, ws.oddHeader.right):
            h.font, h.size = FONT, SIZE
        ws.oddFooter.center.text = "Trang &P/&N"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.horizontalCentered = True
        ws.print_title_rows = title_rows
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w
        return ws

    # ---------------- sheet ẩn HieuChuan ----------------
    def _cal_block(self, cal: CalTable) -> Tuple[int, int]:
        key = id(cal)
        if key in self._cal_blocks:
            return self._cal_blocks[key]
        if self._cal_ws is None:
            self._cal_ws = self.wb.create_sheet(CAL_SHEET)
            self._cal_ws.sheet_state = "hidden"
            for col, w in zip("ABC", (18, 18, 14)):
                self._cal_ws.column_dimensions[col].width = w
        ws, r = self._cal_ws, self._cal_next
        cert = f" (giấy số {cal.cert_no})" if cal.cert_no else ""
        self.merge(ws, r, 1, r, 3, f"Giấy hiệu chuẩn {cal.device}{cert} — {cal.quantity}", align=LEFT, bold=True,
                   grid=False)
        u = f", {cal.unit}" if cal.unit else ""
        for c, lbl in enumerate((f"Giá trị chỉ thị{u}", f"Giá trị chuẩn{u}", f"Sai số{u}"), 1):
            self.put(ws, r + 1, c, lbl, bold=True)
        first = r + 2
        for k, (ind, ref, err) in enumerate(cal.points):
            self.put(ws, first + k, 1, ind)
            self.put(ws, first + k, 2, ref if ref is not None else "-")
            self.put(ws, first + k, 3, err)
        last = first + len(cal.points) - 1
        self._cal_next = last + 3
        self._cal_blocks[key] = (first, last)
        return first, last

    def hc_formula(self, ref: str, value: float, cal: Optional[CalTable]) -> str:
        """Công thức 'sau hiệu chính' cho ô đọc `ref` có giá trị `value` (để chọn cặp điểm kẹp)."""
        if cal is None or not cal.points:
            return f"={ref}"
        first, last = self._cal_block(cal)
        cs = f"'{CAL_SHEET}'!"
        p = cal.points
        if value >= p[-1][0]:
            return f"={ref}-{cs}$C${last}"
        if value <= p[0][0]:
            return f"={ref}-{cs}$C${first}"
        for i in range(len(p) - 1):
            if p[i][0] <= value <= p[i + 1][0]:
                a, b = first + i, first + i + 1
                if p[i][0] == p[i + 1][0]:
                    return f"={ref}-{cs}$C${a}"
                return f"={ref}-FORECAST({ref},{cs}$C${a}:$C${b},{cs}$A${a}:$A${b})"
        return f"={ref}-{cs}$C${last}"

    def _cal_note(self, what: str, cal: Optional[CalTable]):
        if cal is None:
            self.notes.append(f"{what}: không có bảng hiệu chuẩn -> không hiệu chính (giá trị sau hiệu chính = giá trị đọc)")

    # ---------------- Bụi 0,3 (AeroTrak) ----------------
    def dust_aerotrak(self, groups: Sequence[dict], so: str, date: str, labels: Optional[dict] = None,
                      name: str = "Bụi 0,3"):
        """groups: [{"room": tên phòng, "cap": "ISO 8", "limits": {"Ch2": 3520000, "Ch3": 832000, "Ch6": 29300},
        "zero": bản ghi | None, "points": [bản ghi …], "labels": [nhãn vị trí …] (mặc định 1…n), "state": "" }]
        bản ghi = dict đọc bằng `read_aerotrak` (giá trị hạt/m³ chế độ Σ, giữ nguyên số máy)."""
        labels = labels or {}
        cols = [("Phòng", None, 18), ("Vị trí", None, 8)] + \
               [(labels.get(k, k), k, w) for k, w in zip(AEROTRAK_COLS, (7, 12, 11, 19, 9, 10, 11, 11, 11, 11, 11, 11,
                                                                       9, 9, 7, 7, 9, 9))] + \
               [("Mức quy định", None, 11), ("Đánh giá", None, 11)]
        ws = self.sheet(name, so, date, [c[2] for c in cols])
        for c, (lbl, _, _) in enumerate(cols, 1):
            self.put(ws, 1, c, lbl, bold=True)
        ws.row_dimensions[1].height = 32
        idx = {key: c for c, (_, key, _) in enumerate(cols, 1) if key}
        c_lim, c_ev = len(cols) - 1, len(cols)
        for hide in ("Location", "Device Status", "Flow Status"):
            ws.column_dimensions[L(idx[hide])].hidden = True
        r = 2
        for g in groups:
            lim = g.get("limits") or {}
            recs = ([("Zero", g["zero"])] if g.get("zero") else []) + \
                   list(zip(g.get("labels") or range(1, len(g["points"]) + 1), g["points"]))
            r0 = r
            lim_row = r0 + len(recs)
            for lbl, rec in recs:
                self.put(ws, r, 2, lbl, bold=True)
                for key, c in idx.items():
                    v = rec.get(key)
                    is_ch = key.startswith("Ch")
                    self.put(ws, r, c, v, bold=is_ch and key in lim, fmt="#,##0" if is_ch else None)
                if lbl == "Zero":
                    self.put(ws, r, c_ev, "Không tính", italic=True)
                else:
                    self.put(ws, r, c_ev, _eval_formula(
                        f"{L(idx[k])}{r}<={L(idx[k])}${lim_row}" for k in lim))
                r += 1
            # dòng giới hạn của phòng (mặc định ISO 14644-1:2015, Bảng 1; GMP: g["basis"] = "WHO TRS 961 (2011) – Phụ lục 6")
            basis = g.get("basis") or "ISO 14644-1:2015"
            self.merge(ws, r, 1, r, idx["Sample Time"], f"Giới hạn {g.get('cap', '')} ({basis})".strip(),
                       bold=True, align=LEFT)
            for key, c in idx.items():
                if key.startswith("Ch"):
                    self.put(ws, r, c, lim.get(key, "-"), bold=True, fmt="#,##0")
                elif c > idx["Sample Time"]:
                    self.put(ws, r, c, "")
            self.put(ws, r, c_ev, "")
            self.merge(ws, r0, 1, r - 1, 1, g["room"] + (f"\n({g['state']})" if g.get("state") else ""), bold=True)
            self.merge(ws, r0, c_lim, r, c_lim, g.get("cap", "-"))
            r += 1
        return ws

    # ---------------- Bụi 0,1 (Lasair) ----------------
    def dust_lasair(self, groups: Sequence[dict], so: str, date: str, name: str = "Bụi 0,1"):
        """groups: [{"room", "cap", "state": "", "limits": {chỉ số kênh 0..7: giới hạn}, "zero": bản ghi|None,
        "points": [bản ghi …], "labels": [...]}]; bản ghi đọc bằng `read_lasair_csv` / `read_lasair_sheet`.
        Kênh: Size 1..8 = 0,1; 0,15; 0,2; 0,25; 0,3; 0,5; 1,0; 5,0 µm (Counts vi phân)."""
        raw = LASAIR_COLS
        sizes = ["0,1", "0,15", "0,2", "0,25", "0,3", "0,5", "1,0", "5,0"]
        head = [(c, 12 if c in ("Sample Volume", "Sample Date", "Sample Interval") else 10) for c in raw] + [(f"Diff {s}", 9) for s in sizes] + \
               [("Phòng", 18), ("Trạng thái", 11), ("Vị trí", 8)] + [(f"≥ {s} µm", 11) for s in sizes] + \
               [("Mức quy định", 11), ("Đánh giá", 11)]
        ws = self.sheet(name, so, date, [w for _, w in head])
        for c, (lbl, _) in enumerate(head, 1):
            self.put(ws, 1, c, lbl, bold=True)
        ws.row_dimensions[1].height = 32
        ci = {lbl: c for c, (lbl, _) in enumerate(head, 1)}
        c_diff0 = ci["Diff 0,1"]
        c_cum0 = ci["≥ 0,1 µm"]
        c_lim, c_ev = ci["Mức quy định"], ci["Đánh giá"]
        visible = {"Data ID", "Sample Date", "Sample Time", "S/N", "Location", "Sample Interval", "Sample Volume",
                   "Phòng", "Trạng thái", "Vị trí", "Mức quy định", "Đánh giá"} | {f"≥ {s} µm" for s in sizes}
        for lbl, c in ci.items():
            if lbl not in visible:
                ws.column_dimensions[L(c)].hidden = True
        r = 2
        for g in groups:
            lim = g.get("limits") or {}
            recs = ([("Zero", g["zero"])] if g.get("zero") else []) + \
                   list(zip(g.get("labels") or range(1, len(g["points"]) + 1), g["points"]))
            r0 = r
            lim_row = r0 + len(recs)
            for lbl, rec in recs:
                for k in raw:
                    v = rec.get(k)
                    if k in ("Sample Volume",) or k.startswith("Counts") or k in ("Data ID",):
                        v = _num(v)
                    import datetime as _dtm
                    fmt = "0.000000" if k == "Sample Volume" else None
                    if isinstance(v, _dtm.datetime):
                        fmt = "dd/mm/yyyy" if (v.hour, v.minute, v.second) == (0, 0, 0) else "dd/mm/yyyy hh:mm:ss"
                    elif isinstance(v, _dtm.time):
                        fmt = "hh:mm:ss"
                    self.put(ws, r, ci[k], v, fmt=fmt)
                vol = f"${L(ci['Sample Volume'])}{r}"
                for i in range(8):
                    cnt = f"{L(ci[f'Counts {i + 1}'])}{r}"
                    self.put(ws, r, c_diff0 + i, f"={cnt}/{vol}", fmt="#,##0.0")
                for i in range(8):
                    self.put(ws, r, c_cum0 + i,
                             f"=ROUND(SUM({L(c_diff0 + i)}{r}:{L(c_diff0 + 7)}{r}),0)", fmt="#,##0", bold=i in lim)
                self.put(ws, r, ci["Vị trí"], lbl, bold=True)
                if lbl == "Zero":
                    self.put(ws, r, c_ev, "Không tính", italic=True)
                else:
                    self.put(ws, r, c_ev, _eval_formula(
                        f"{L(c_cum0 + i)}{r}<={L(c_cum0 + i)}${lim_row}" for i in sorted(lim) if lim[i] is not None))
                r += 1
            self.merge(ws, r, 1, r, ci["Vị trí"], f"Giới hạn {g.get('cap', '')} (ISO 14644-1:2015)".strip(),
                       bold=True, align=LEFT)
            for i in range(8):
                v = lim.get(i, "-")
                self.put(ws, r, c_cum0 + i, "-" if v is None else v, bold=True, fmt="#,##0")
            self.put(ws, r, c_ev, "")
            self.merge(ws, r0, ci["Phòng"], r - 1, ci["Phòng"], g["room"], bold=True)
            self.merge(ws, r0, ci["Trạng thái"], r - 1, ci["Trạng thái"], g.get("state") or "-")
            self.merge(ws, r0, c_lim, r, c_lim, g.get("cap", "-"))
            r += 1
        return ws

    # ---------------- Nhiệt Ẩm ----------------
    def temp_rh(self, rows: Sequence[dict], cal_t: Optional[CalTable], cal_rh: Optional[CalTable], so: str,
                date: str, n: int = 3):
        """rows: [{"room", "loc", "t": [..n], "rh": [..n], "lim_t": "20 - 24"|None, "lim_rh": "≤ 65"|None}].
        Phòng nhiều điểm -> thêm dòng 'Trung bình phòng', đánh giá theo trung bình phòng (AGENTS 8C)."""
        self._cal_note("Nhiệt độ", cal_t)
        self._cal_note("Độ ẩm", cal_rh)
        c_t, c_rh, c_tc, c_rhc = 4, 4 + n, 4 + 2 * n, 4 + 3 * n
        c_avt, c_avrh = 4 + 4 * n, 5 + 4 * n
        c_lt, c_et, c_lrh, c_erh = c_avrh + 1, c_avrh + 2, c_avrh + 3, c_avrh + 4
        ws = self.sheet("Nhiệt Ẩm", so, date, [5, 16, 14] + [7] * (4 * n) + [10, 10, 10, 10, 10, 10], "1:2")
        for c, lbl in ((1, "STT"), (2, "Phòng"), (3, "Vị trí")):
            self.merge(ws, 1, c, 2, c, lbl, bold=True)
        for c0, lbl in ((c_t, "Nhiệt độ đọc, °C"), (c_rh, "Độ ẩm đọc, %RH"), (c_tc, "Nhiệt độ sau hiệu chính, °C"),
                        (c_rhc, "Độ ẩm sau hiệu chính, %RH")):
            self.merge(ws, 1, c0, 1, c0 + n - 1, lbl, bold=True)
            for i in range(n):
                self.put(ws, 2, c0 + i, i + 1, bold=True)
        self.merge(ws, 1, c_avt, 1, c_avrh, "Trung bình", bold=True)
        self.put(ws, 2, c_avt, "Nhiệt độ, °C", bold=True)
        self.put(ws, 2, c_avrh, "Độ ẩm, %RH", bold=True)
        self.merge(ws, 1, c_lt, 1, c_et, "Nhiệt độ", bold=True)
        self.merge(ws, 1, c_lrh, 1, c_erh, "Độ ẩm", bold=True)
        for c, lbl in ((c_lt, "Mức quy định"), (c_et, "Đánh giá"), (c_lrh, "Mức quy định"), (c_erh, "Đánh giá")):
            self.put(ws, 2, c, lbl, bold=True)
        ws.row_dimensions[1].height = 30
        r, stt = 3, 1
        rooms = []
        for row in rows:
            if not rooms or rooms[-1][0] != row["room"]:
                rooms.append((row["room"], []))
            rooms[-1][1].append(row)
        for room, pts in rooms:
            r0 = r
            multi = len(pts) > 1
            for p in pts:
                self.put(ws, r, 1, stt)
                self.put(ws, r, 3, p.get("loc", ""))
                for i in range(n):
                    t, h = p["t"][i], p["rh"][i]
                    self.put(ws, r, c_t + i, t, fmt="0.0")
                    self.put(ws, r, c_rh + i, h, fmt="0.0")
                    self.put(ws, r, c_tc + i, self.hc_formula(f"{L(c_t + i)}{r}", t, cal_t), fmt="0.00")
                    self.put(ws, r, c_rhc + i, self.hc_formula(f"{L(c_rh + i)}{r}", h, cal_rh), fmt="0.00")
                self.put(ws, r, c_avt, f"=ROUND(AVERAGE({L(c_tc)}{r}:{L(c_tc + n - 1)}{r}),1)", fmt="0.0", bold=True)
                self.put(ws, r, c_avrh, f"=ROUND(AVERAGE({L(c_rhc)}{r}:{L(c_rhc + n - 1)}{r}),1)", fmt="0.0",
                         bold=True)
                if not multi:
                    self._limit_cells(ws, r, c_lt, c_et, f"{L(c_avt)}{r}", p.get("lim_t"))
                    self._limit_cells(ws, r, c_lrh, c_erh, f"{L(c_avrh)}{r}", p.get("lim_rh"))
                else:
                    for c in (c_lt, c_et, c_lrh, c_erh):
                        self.put(ws, r, c, "")
                r += 1
                stt += 1
            if multi:
                self.merge(ws, r, 1, r, c_avt - 1, "Trung bình phòng", bold=True, italic=True)
                self.put(ws, r, c_avt, f"=ROUND(AVERAGE({L(c_avt)}{r0}:{L(c_avt)}{r - 1}),1)", fmt="0.0", bold=True)
                self.put(ws, r, c_avrh, f"=ROUND(AVERAGE({L(c_avrh)}{r0}:{L(c_avrh)}{r - 1}),1)", fmt="0.0", bold=True)
                self._limit_cells(ws, r, c_lt, c_et, f"{L(c_avt)}{r}", pts[0].get("lim_t"))
                self._limit_cells(ws, r, c_lrh, c_erh, f"{L(c_avrh)}{r}", pts[0].get("lim_rh"))
                r += 1
            self.merge(ws, r0, 2, r - 1 - (1 if multi else 0), 2, room)
        return ws

    def _limit_cells(self, ws, r, c_lim, c_ev, ref, lim):
        self.put(ws, r, c_lim, lim if lim not in (None, "") else "-")
        self.put(ws, r, c_ev, _eval_formula([_limit_cond(ref, lim)]))

    # ---------------- Độ ồn ----------------
    def noise(self, rows: Sequence[dict], so: str, date: str, cal: Optional[CalTable] = None):
        """rows: [{"room", "loc", "reads": [LAeq …], "lim": "≤ 65"|None, "note": ""}] — mức None/'-' -> đánh giá '-'."""
        self._cal_note("Độ ồn", cal)
        n = max(len(x["reads"]) for x in rows)
        c_r, c_c = 4, 4 + n
        c_av, c_lim, c_ev = 4 + 2 * n, 5 + 2 * n, 6 + 2 * n
        ws = self.sheet("Độ ồn", so, date, [5, 18, 14] + [9] * (2 * n) + [11, 11, 11], "1:2")
        for c, lbl in ((1, "STT"), (2, "Phòng"), (3, "Vị trí")):
            self.merge(ws, 1, c, 2, c, lbl, bold=True)
        self.merge(ws, 1, c_r, 1, c_r + n - 1, "Độ ồn đọc LAeq, dBA", bold=True)
        self.merge(ws, 1, c_c, 1, c_c + n - 1, "Sau hiệu chính, dBA", bold=True)
        for i in range(n):
            self.put(ws, 2, c_r + i, i + 1, bold=True)
            self.put(ws, 2, c_c + i, i + 1, bold=True)
        for c, lbl in ((c_av, "Trung bình, dBA"), (c_lim, "Mức quy định"), (c_ev, "Đánh giá")):
            self.merge(ws, 1, c, 2, c, lbl, bold=True)
        r = 3
        for k, p in enumerate(rows, 1):
            self.put(ws, r, 1, k)
            self.put(ws, r, 2, p["room"])
            self.put(ws, r, 3, p.get("loc", ""))
            for i in range(n):
                if i < len(p["reads"]):
                    v = p["reads"][i]
                    self.put(ws, r, c_r + i, v, fmt="0.0")
                    self.put(ws, r, c_c + i, self.hc_formula(f"{L(c_r + i)}{r}", v, cal), fmt="0.0")
                else:
                    self.put(ws, r, c_r + i, "")
                    self.put(ws, r, c_c + i, "")
            m = len(p["reads"])
            self.put(ws, r, c_av, f"=ROUND(AVERAGE({L(c_c)}{r}:{L(c_c + m - 1)}{r}),1)", fmt="0.0", bold=True)
            self._limit_cells(ws, r, c_lim, c_ev, f"{L(c_av)}{r}", p.get("lim"))
            r += 1
        return ws

    # ---------------- Chênh áp ----------------
    def pressure(self, rows: Sequence[dict], so: str, date: str, cal: Optional[CalTable] = None, n: int = 3):
        """rows: [{"a": vị trí 1, "b": vị trí 2, "sign": "PA1", "reads": [..n], "lim": "3 - 7"|None}]."""
        self._cal_note("Chênh áp", cal)
        c_r, c_c = 6, 6 + n
        c_av, c_lim, c_ev = 6 + 2 * n, 7 + 2 * n, 8 + 2 * n
        ws = self.sheet("Chênh áp", so, date, [5, 18, 4, 18, 8] + [9] * (2 * n) + [11, 11, 11], "1:2")
        for c, lbl in ((1, "STT"), (2, "Vị trí 1"), (3, "→"), (4, "Vị trí 2"), (5, "Kí hiệu")):
            self.merge(ws, 1, c, 2, c, lbl, bold=True)
        self.merge(ws, 1, c_r, 1, c_r + n - 1, "Giá trị đo, Pa", bold=True)
        self.merge(ws, 1, c_c, 1, c_c + n - 1, "Sau hiệu chính, Pa", bold=True)
        for i in range(n):
            self.put(ws, 2, c_r + i, i + 1, bold=True)
            self.put(ws, 2, c_c + i, i + 1, bold=True)
        for c, lbl in ((c_av, "Trung bình, Pa"), (c_lim, "Mức quy định, Pa"), (c_ev, "Đánh giá")):
            self.merge(ws, 1, c, 2, c, lbl, bold=True)
        r = 3
        for k, p in enumerate(rows, 1):
            for c, v in ((1, k), (2, p["a"]), (3, "→"), (4, p["b"]), (5, p.get("sign", ""))):
                self.put(ws, r, c, v)
            for i, v in enumerate(p["reads"][:n]):
                self.put(ws, r, c_r + i, v, fmt="0.000")
                self.put(ws, r, c_c + i, self.hc_formula(f"{L(c_r + i)}{r}", v, cal), fmt="0.000")
            self.put(ws, r, c_av, f"=ROUND(AVERAGE({L(c_c)}{r}:{L(c_c + n - 1)}{r}),1)", fmt="0.0", bold=True)
            self._limit_cells(ws, r, c_lim, c_ev, f"{L(c_av)}{r}", p.get("lim"))
            r += 1
        return ws

    # ---------------- Lưu lượng ----------------
    def flow(self, rooms: Sequence[dict], so: str, date: str, cal: Optional[CalTable] = None,
             factor: Optional[float] = None, n: Optional[int] = None):
        """rooms: [{"room", "area": m², "height": m, "lim_ach": "≥ 20"|None,
                    "hepas": [{"name": "HEPA 1", "parts": [{"part": "1-1", "date": "", "time": "", "reads": [...]}]}]}]
        Hiệu chỉnh: `cal` (bảng hiệu chuẩn lưu lượng, trừ sai số nội suy) hoặc `factor` (hệ số nhân của giấy, vd 0,981)."""
        if cal is None and factor is None:
            self.notes.append("Lưu lượng: không có bảng hiệu chuẩn -> không hiệu chính (giá trị sau hiệu chính = giá trị đọc)")
        n = n or max(len(p["reads"]) for rm in rooms for h in rm["hepas"] for p in h["parts"])
        cols = ["Phòng", "Vị trí", "Phần", "Ngày", "Thời điểm"] + [f"Lần {i + 1}" for i in range(n)] + \
               [f"HC{i + 1}" for i in range(n)] + ["Trung bình, m³/h", "Lưu lượng HEPA, m³/h", "Tổng lưu lượng, m³/h",
                                                    "Diện tích, m²", "Chiều cao, m", "Thể tích phòng, m³", "ACH",
                                                    "Mức quy định", "Đánh giá"]
        widths = [16, 10, 7, 11, 10] + [8] * (2 * n) + [11, 11, 11, 9, 8, 10, 8, 10, 10]
        ws = self.sheet("Lưu lượng", so, date, widths)
        for c, lbl in enumerate(cols, 1):
            self.put(ws, 1, c, lbl, bold=True)
        ws.row_dimensions[1].height = 45
        c_r, c_h = 6, 6 + n
        c_av, c_hp, c_tot = 6 + 2 * n, 7 + 2 * n, 8 + 2 * n
        c_area, c_ht, c_vol, c_ach, c_lim, c_ev = (9 + 2 * n + i for i in range(6))
        r = 2
        for rm in rooms:
            r_room = r
            hepa_cells = []
            for hp in rm["hepas"]:
                r_h = r
                for p in hp["parts"]:
                    for c, v in ((3, p.get("part", "")), (4, p.get("date", "")), (5, p.get("time", ""))):
                        self.put(ws, r, c, v)
                    for i in range(n):
                        if i < len(p["reads"]):
                            v = p["reads"][i]
                            ref = f"{L(c_r + i)}{r}"
                            self.put(ws, r, c_r + i, v)
                            f = f"={ref}*{factor}" if factor is not None else self.hc_formula(ref, v, cal)
                            self.put(ws, r, c_h + i, f, fmt="0.0")
                        else:
                            self.put(ws, r, c_r + i, "")
                            self.put(ws, r, c_h + i, "")
                    m = len(p["reads"])
                    self.put(ws, r, c_av, f"=ROUND(AVERAGE({L(c_h)}{r}:{L(c_h + m - 1)}{r}),0)", fmt="0")
                    r += 1
                self.merge(ws, r_h, 2, r - 1, 2, hp["name"])
                self.merge(ws, r_h, c_hp, r - 1, c_hp, f"=SUM({L(c_av)}{r_h}:{L(c_av)}{r - 1})", fmt="#,##0")
                hepa_cells.append(f"{L(c_hp)}{r_h}")
            r1 = r - 1
            self.merge(ws, r_room, 1, r1, 1, rm["room"], bold=True)
            self.merge(ws, r_room, c_tot, r1, c_tot, "=" + "+".join(hepa_cells), fmt="#,##0", bold=True)
            self.merge(ws, r_room, c_area, r1, c_area, rm.get("area", ""), fmt="0.00")
            self.merge(ws, r_room, c_ht, r1, c_ht, rm.get("height", ""), fmt="0.00")
            vol = f"{L(c_vol)}{r_room}"
            self.merge(ws, r_room, c_vol, r1, c_vol, f"={L(c_area)}{r_room}*{L(c_ht)}{r_room}", fmt="0.0")
            ach = f"{L(c_ach)}{r_room}"
            self.merge(ws, r_room, c_ach, r1, c_ach, f"=ROUND({L(c_tot)}{r_room}/{vol},1)", fmt="0.0", bold=True)
            lim = rm.get("lim_ach")
            self.merge(ws, r_room, c_lim, r1, c_lim, lim if lim not in (None, "") else "-")
            self.merge(ws, r_room, c_ev, r1, c_ev, _eval_formula([_limit_cond(ach, lim)]))
        return ws

    # ---------------- Scan lọc ----------------
    def scan(self, rows: Sequence[dict], so: str, date: str, lim_default: str = "≤ 0,01"):
        """rows: [{"name": ký hiệu miệng cấp, "media": %, "gasket": %, "lim": "≤ 0,01" | "≤ 0,1" (H13)}]."""
        ws = self.sheet("Scan lọc", so, date, [6, 30, 14, 14, 14, 13, 11])
        for c, lbl in enumerate(("STT", "Ký hiệu miệng cấp", "Media, %", "Gasket, %", "Max pen, %",
                                 "Mức quy định, %", "Đánh giá"), 1):
            self.put(ws, 1, c, lbl, bold=True)
        r = 2
        for k, p in enumerate(rows, 1):
            for c, v in ((1, k), (2, p["name"]), (3, p.get("media", 0)), (4, p.get("gasket", 0))):
                self.put(ws, r, c, v, fmt="0.0000" if c in (3, 4) else None)
            self.put(ws, r, 5, f'=IF(MAX(C{r}:D{r})<>0,MAX(C{r}:D{r}),"< 0,0001")', fmt="0.0000", bold=True)
            lim = p.get("lim") or lim_default
            cond = _limit_cond(f"E{r}", lim)
            self.put(ws, r, 6, lim)
            self.put(ws, r, 7, f'=IF(ISNUMBER(E{r}),IF({cond},"Đạt","Không đạt"),"Đạt")' if cond else "-")
            r += 1
        return ws

    # ---------------- thêm sheet riêng của hồ sơ (vd Vi sinh) ----------------
    def extra_sheet(self, name: str, so: str, date: str, widths: Sequence[float], title: str, title_rows: str = "1:1"):
        return self.sheet(name, so, date, widths, title_rows, title=title)

    # ---------------- lưu ----------------
    def save(self, path: str) -> str:
        if self._cal_ws is not None:
            names = self.wb.sheetnames
            self.wb.move_sheet(CAL_SHEET, offset=len(names) - 1 - names.index(CAL_SHEET))
        thin_outside_borders(self.wb)
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.wb.save(path)
        for nt in dict.fromkeys(self.notes):
            print(f"[cleanroom_calc] {nt}")
        return path


# ---------------------------------------------------------------------------
# Sửa dòng viện dẫn phiếu trong BCKQ đã dựng (giữ nguyên ngày, định dạng) — dùng khi chỉ dựng lại phiếu
# ---------------------------------------------------------------------------
def replace_reference_lines(docx_path: str, startswith: str, new_lines: Sequence[str]) -> bool:
    """Đoạn (thân văn bản hoặc trong bảng) bắt đầu bằng `startswith` -> thay bằng các dòng `new_lines`
    (dòng đầu giữ đoạn cũ, dòng sau chép định dạng đoạn cũ). Trả True khi có thay."""
    import copy
    from docx import Document
    doc = Document(docx_path)

    def paragraphs():
        yield from doc.paragraphs
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    yield from cell.paragraphs

    done = False
    seen = set()
    for p in list(paragraphs()):
        if id(p._p) in seen or not p.text.strip().startswith(startswith):
            continue
        seen.add(id(p._p))
        anchor = p._p
        runs = p.runs
        if not runs:
            continue
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        runs[0].text = new_lines[0]
        for line in new_lines[1:]:
            el = copy.deepcopy(p._p)
            anchor.addnext(el)
            anchor = el
            from docx.text.paragraph import Paragraph
            q = Paragraph(el, p._parent)
            q.runs[0].text = line
        done = True
    if done:
        doc.save(docx_path)
    return done
