# -*- coding: utf-8 -*-
"""
Phiếu tính toán tủ (TinhToanBSC.xlsx — BSC, tủ cấy/LAF): mỗi phép đo có thiết bị đã có bảng hiệu chuẩn tách 2 phần
(người dùng chốt 07/11/2026):
    a) Kết quả đọc (chưa hiệu chỉnh)   — lưới số đọc gốc của form (giữ nguyên vị trí)
    b) Kết quả sau hiệu chỉnh          — lưới công thức ngay dưới lưới đọc: = đọc − FORECAST(đọc, sai số, chỉ thị)
và các ô kết quả (trung bình, lớn nhất/nhỏ nhất, Vtb, vận tốc DIM) trỏ vào phần b). Phép đo chưa có bảng hiệu chuẩn
(độ ồn Rion, rò rỉ HEPA ATI 2i, đếm hạt) giữ 1 phần.

Chèn dòng bằng Excel COM để Excel tự sửa mọi tham chiếu (cả sheet TBKQ-VN / TBKQ - ENG, vùng in, ảnh).
Bảng hiệu chuẩn ở sheet ẩn `HieuChuan` (không in — AGENTS mục 11). Mọi ô tính ra là công thức (không ghi số tĩnh):
hiệu chỉnh ồn nền I70 tra bảng bằng IF; bỏ dòng ghi chú bố trí điểm đo (A34, A51).

Dùng sau khi đã điền và hậu xử lý xong bằng openpyxl (địa chỉ ô form gốc):
    from Core.Tools.bsc_cal_blocks import finalize_bsc_workbook
    finalize_bsc_workbook(path, {"downflow": cal_kanomax, "inflow_kanomax": cal_kanomax, "lux": cal_lux,
                                 "uvc": cal_uvc, "inflow_dim": cal_tsi8380})
Số trên TBKQ/BCKQ tính bằng cùng bảng: `Core.Tools.cleanroom_calc.apply_cal` / `CalTable.corrected`.
"""
from __future__ import annotations

import gc
import os
from typing import Dict, Optional

from .cleanroom_calc import CalTable

CAL_SHEET = "HieuChuan"
SHEET = "Bảng tính"

# Khối lưới đo của form TinhToanBSC (dòng theo form gốc, trước khi chèn):
#   copy   = các dòng chép làm khung phần b)        grid = vùng ô số đọc
#   label  = ô ghi nhãn "a) Kết quả đọc …" (dòng trống trên lưới; "^" = chèn 1 dòng trên khối; "A10+" = nối vào ô)
#   results = ô kết quả cần trỏ sang phần b) {ô: mẫu công thức}
BLOCKS = {
    "uvc": {"copy": (113, 119), "grid": "D114:K118", "label": "^", "unit": "µW/cm²",
            "results": {"D122": "=AVERAGE({g})"}},
    "lux": {"copy": (52, 58), "grid": "C53:M57", "label": "A51", "unit": "lux",
            "results": {"D62": "=AVERAGE({g})"}},
    "downflow": {"copy": (35, 42), "grid": "C37:K41", "label": "A34", "unit": "m/s",
                 "results": {"D46": "=AVERAGE({g})", "D48": "=MAX({g})", "D49": "=MIN({g})"}},
    "inflow_kanomax": {"copy": (18, 25), "grid": "C21:J25", "label": "A20", "unit": "m/s",
                       "results": {"D31": "={vtb}"}},
    "inflow_dim": {"copy": (11, 13), "grid": "B12:F12", "label": "A10+", "unit": "m³/h",
                   "results": {"D31": "={vel}"}},
}
ORDER = ("uvc", "lux", "downflow", "inflow_kanomax", "inflow_dim")   # dưới lên: chèn khối dưới không làm lệch khối trên

NOISE_I70 = ("=IF(I68<=55,0,IF(ROUND(ROUND(I69-I68,1),0)<=2,0,IF(ROUND(ROUND(I69-I68,1),0)=3,3,"
             "IF(ROUND(ROUND(I69-I68,1),0)<=5,2,IF(ROUND(ROUND(I69-I68,1),0)<=10,1,0)))))")


def _retry(fn, *a, tries: int = 40, **kw):
    """Excel bận (RPC_E_CALL_REJECTED / SERVERCALL_RETRYLATER) -> chờ rồi gọi lại."""
    import time
    import pywintypes
    for i in range(tries):
        try:
            return fn(*a, **kw)
        except pywintypes.com_error as e:
            if e.args and e.args[0] in (-2147418111, -2147417846) and i < tries - 1:
                time.sleep(0.25)
                continue
            raise


def _set(rng, attr: str, value):
    _retry(setattr, rng, attr, value)


class _P:
    """Bọc đối tượng COM: mọi đọc/ghi thuộc tính, gọi hàm đều tự thử lại khi Excel bận."""
    __slots__ = ("_o",)

    def __init__(self, o):
        object.__setattr__(self, "_o", o)

    @staticmethod
    def _w(v):
        return _P(v) if hasattr(v, "_oleobj_") else v

    @staticmethod
    def _u(v):
        return v._o if isinstance(v, _P) else v

    def __getattr__(self, name):
        v = _retry(getattr, self._o, name)
        if hasattr(v, "_oleobj_"):
            return _P(v)
        if callable(v):
            return lambda *a, **k: _P._w(_retry(v, *[_P._u(x) for x in a], **{kk: _P._u(x) for kk, x in k.items()}))
        return v

    def __setattr__(self, name, value):
        _retry(setattr, self._o, name, _P._u(value))

    def __call__(self, *a, **k):
        return _P._w(_retry(self._o, *[_P._u(x) for x in a], **{kk: _P._u(x) for kk, x in k.items()}))

    def __iter__(self):
        for x in self._o:
            yield _P._w(x)


def _col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _write_cal_sheet(wb, cals: Dict[str, CalTable]):
    """Sheet ẩn HieuChuan: mỗi bảng 1 khối (tiêu đề, cột chỉ thị | chuẩn | sai số). Trả {key: (dòng đầu, dòng cuối)}."""
    ws = None
    for s in wb.Worksheets:
        if s.Name == CAL_SHEET:
            ws = s
    if ws is None:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = CAL_SHEET
    r = (ws.UsedRange.Row + ws.UsedRange.Rows.Count + 1) if ws.UsedRange.Value not in (None, "") else 1
    out, seen = {}, {}
    for key, cal in cals.items():
        if cal is None or not cal.points:
            continue
        if id(cal) in seen:
            out[key] = seen[id(cal)]
            continue
        cert = f" (giấy số {cal.cert_no})" if cal.cert_no else ""
        ws.Cells(r, 1).Value = f"Giấy hiệu chuẩn {cal.device}{cert} — {cal.quantity}"
        ws.Cells(r, 1).Font.Bold = True
        u = f", {cal.unit}" if cal.unit else ""
        for c, lbl in enumerate((f"Giá trị chỉ thị{u}", f"Giá trị chuẩn{u}", f"Sai số{u}"), 1):
            ws.Cells(r + 1, c).Value = lbl
            ws.Cells(r + 1, c).Font.Bold = True
        first = r + 2
        for k, (ind, ref, err) in enumerate(cal.points):
            ws.Cells(first + k, 1).Value = ind
            ws.Cells(first + k, 2).Value = "-" if ref is None else ref
            ws.Cells(first + k, 3).Value = err
        last = first + len(cal.points) - 1
        ws.Range(f"A{r + 1}:C{last}").Borders.LineStyle = 1
        ws.Columns("A:C").ColumnWidth = 18
        out[key] = seen[id(cal)] = (first, last)
        r = last + 3
    ws.Visible = 0          # xlSheetHidden
    return out


def _hc_formula(ref: str, value: float, cal: CalTable, rows) -> str:
    first, last = rows
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


def _add_block(ws, key: str, cal: CalTable, cal_rows):
    """Chèn phần b) dưới khối `key`; trả về số dòng đã chèn."""
    spec = BLOCKS[key]
    r0, r1 = spec["copy"]
    g = ws.Range(spec["grid"])
    g_r0, g_c0 = g.Row, g.Column
    nums = []
    for cell in g.Cells:
        v = cell.Value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            nums.append((cell.Row - g_r0, cell.Column - g_c0, float(v), cell.GetAddress(False, False)))
    if not nums:
        return 0
    n_copy = r1 - r0 + 1
    n_ins = n_copy + 1                                   # 1 dòng tiêu đề phần b)
    _retry(ws.Rows(f"{r1 + 1}:{r1 + n_ins}").Insert)
    title_row = r1 + 1
    dst0 = title_row + 1
    _retry(ws.Rows(f"{r0}:{r1}").Copy, ws.Rows(f"{dst0}:{dst0 + n_copy - 1}"))
    cert = f" theo giấy hiệu chuẩn {cal.cert_no}" if cal.cert_no else " theo giấy hiệu chuẩn"
    _title(ws, title_row, f"b) Kết quả sau hiệu chỉnh{cert}, {spec['unit']}")
    lab = f"a) Kết quả đọc (chưa hiệu chỉnh), {spec['unit']}"
    if spec["label"].endswith("+"):                      # nối vào dòng tiêu đề phương pháp (DIM: A10)
        c = ws.Range(spec["label"][:-1])
        c.Value = f"{str(c.Value or '').strip()} — {lab}"
    elif spec["label"] != "^":
        _title(ws, ws.Range(spec["label"]).Row, lab)
    shift = dst0 - r0
    # lưới phần b): xóa số chép sang, ghi công thức hiệu chỉnh cho đúng các ô có số đọc
    gb = ws.Range(ws.Cells(g_r0 + shift, g_c0), ws.Cells(g.Row + g.Rows.Count - 1 + shift, g_c0 + g.Columns.Count - 1))
    for cell in gb.Cells:
        if isinstance(cell.Value, (int, float)) and not isinstance(cell.Value, bool):
            cell.Value = None
    for dr, dc, v, raw_addr in nums:
        _set(_retry(ws.Cells, g_r0 + shift + dr, g_c0 + dc), "Formula", _hc_formula(raw_addr, v, cal, cal_rows))
    g_b = gb.GetAddress(False, False)
    extra = {"g": g_b}
    if key == "inflow_kanomax":
        # TB / k / Vtb của phần b) (cột K:L dòng 22–24 chép sang)
        tb, k, vtb = f"L{22 + shift}", f"L{23 + shift}", f"L{24 + shift}"
        ws.Range(tb).Formula = f"=AVERAGE({g_b})"
        if isinstance(ws.Range("L23").Value, (int, float)):        # có hệ số k theo nhãn tủ
            ws.Range(k).Formula = "=L23"
            ws.Range(vtb).Formula = f"={tb}*{k}"
            extra["vtb"] = vtb
        else:
            extra["vtb"] = tb
    if key == "inflow_dim":
        g_avg, vel = f"G{12 + shift}", f"J{12 + shift}"
        ws.Range(f"A{12 + shift}").Formula = "=A12"
        ws.Range(g_avg).Formula = f"=AVERAGE({g_b})"
        ws.Range(vel).Formula = f"={g_avg}/A12/3600"
        extra["vel"] = vel
    import re as _re
    for addr, tpl in spec["results"].items():
        col, row = _re.match(r"([A-Z]+)(\d+)$", addr).groups()
        ws.Range(f"{col}{int(row) + n_ins}").Formula = tpl.format(**extra)   # ô kết quả dưới khối dời n_ins dòng
    if spec["label"] == "^":                            # chèn dòng nhãn a) trên khối (Excel tự dời công thức)
        ws.Rows(f"{r0}:{r0}").Insert()
        _title(ws, r0, lab)
        n_ins += 1
    return n_ins


def _title(ws, row: int, text: str):
    """Dòng nhãn a) / b): bỏ định dạng chép từ dòng bên cạnh (gộp ô, căn giữa), ghi chữ nghiêng đậm căn trái."""
    rng = ws.Rows(f"{row}:{row}")
    rng.UnMerge()
    rng.ClearFormats()
    c = ws.Cells(row, 1)
    c.Value = text
    c.Font.Name = "Times New Roman"
    c.Font.Size = 12
    c.Font.Bold = True
    c.Font.Italic = True
    c.HorizontalAlignment = -4131        # xlLeft
    c.WrapText = False


def finalize_bsc_workbook(path: str, cals: Optional[Dict[str, Optional[CalTable]]] = None) -> str:
    """Thêm phần sau hiệu chỉnh + công thức ồn nền, tính lại, lưu. `cals` khóa: uvc, lux, downflow, inflow_kanomax,
    inflow_dim (None = thiết bị chưa có bảng hiệu chuẩn -> giữ 1 phần)."""
    import win32com.client
    cals = {k: v for k, v in (cals or {}).items() if v is not None and v.points}
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        xl = _P(xl)
        wb = xl.Workbooks.Open(os.path.abspath(path))
        xl.ScreenUpdating = False
        xl.Calculation = -4135                  # xlCalculationManual trong lúc chèn dòng
        ws = wb.Worksheets(SHEET)
        # Công thức thay số tĩnh; bỏ ghi chú bố trí điểm đo (AGENTS mục 11)
        if isinstance(ws.Range("I68").Value, (int, float)) and isinstance(ws.Range("I69").Value, (int, float)):
            ws.Range("I70").Formula = NOISE_I70
        for a in ("A34", "A51"):
            v = ws.Range(a).Formula
            if isinstance(v, str) and "Ghi chú" in v:
                ws.Range(a).Value = None
        cal_rows = _write_cal_sheet(wb, cals) if cals else {}
        for key in ORDER:
            if key in cals:
                n = _add_block(ws, key, cals[key], cal_rows[key])
                print(f"[bsc_cal_blocks] {key}: thêm phần sau hiệu chỉnh ({n} dòng)")
        xl.Calculation = -4105                  # xlCalculationAutomatic
        if cals:
            # Chèn dòng làm lưới ảnh bị ngắt trang cắt ngang (136247: nửa ảnh sang trang 5) -> ngắt trang trước mục ảnh
            try:
                hit = ws.Columns(1).Find("Hình ảnh quá trình")
                if hit is not None:
                    ws.HPageBreaks.Add(ws.Range(f"A{hit.Row}"))
            except Exception as e:  # noqa: BLE001 — không chặn việc lưu phiếu
                print(f"[bsc_cal_blocks] không đặt được ngắt trang trước mục ảnh: {e}")
        ws.Activate()
        _retry(xl.CalculateFull)
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        wb = None
        xl.Quit()
        del xl
        gc.collect()
    return path


def corrected_values(readings, cal: Optional[CalTable]):
    """Số đọc (danh sách hoặc lưới) -> số sau hiệu chỉnh cùng cấu trúc (giữ số đọc khi chưa có bảng hiệu chuẩn).
    Script hồ sơ dùng để tính số trên TBKQ/BCKQ khớp phần b) của phiếu."""
    if cal is None or not cal.points:
        return readings
    if readings and isinstance(readings[0], (list, tuple)):
        return [[cal.corrected(float(v)) for v in row] for row in readings]
    return [cal.corrected(float(v)) for v in readings]
