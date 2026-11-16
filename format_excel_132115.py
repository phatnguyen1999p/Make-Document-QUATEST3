# -*- coding: utf-8 -*-
"""Chuan hoa header + ke grid cho file tinh toan phong sach 132115."""
import copy
import re
import sys
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.dimensions import ColumnDimension

CODE = "Q326D132115"
DATE = "23/10/2026"
FONT = "Times New Roman"

thin = Side(style="thin", color="000000")
GRID = Border(left=thin, right=thin, top=thin, bottom=thin)

# So hieu tung phieu tinh toan (header trai): 'Số: <mã>/<ten phieu>' (nguoi dung chot 30/10/2026)
SHEET_NO = {
    "Bụi 0,3": "Tính Toán Bụi",
    "Nhiệt Ẩm": "Tính Toán Nhiệt Ẩm",
    "Lưu lượng": "Tính Toán Lưu Lượng",
    "Chênh áp": "Tính Toán Chênh Áp",
}
# So dong tieu de cot cua tung sheet (lap lai khi in sang trang + co dinh khi cuon)
HEADER_ROWS = {"Bụi 0,3": 1, "Nhiệt Ẩm": 2, "Lưu lượng": 1, "Chênh áp": 1}


def set_header(ws):
    hdr = ws.oddHeader
    title = hdr.center.text
    name = SHEET_NO.get(ws.title)
    hdr.left.text = f"Số: {CODE}/{name}" if name else f"Số: {CODE}"
    hdr.right.text = f"Ngày {DATE}"
    hdr.center.text = title
    for part in (hdr.left, hdr.center, hdr.right):
        part.font = part.font or f"{FONT},Regular"
        part.size = part.size or 13


def last_row(ws):
    for r in range(ws.max_row, 0, -1):
        if any(c.value not in (None, "") for c in ws[r]):
            return r
    return 1


def last_col(ws):
    mc = 1
    for row in ws.iter_rows():
        for c in row:
            if c.value not in (None, ""):
                mc = max(mc, c.column)
    for rng in ws.merged_cells.ranges:
        if rng.min_row == 1:
            mc = max(mc, rng.max_col)
    return mc


def merge_runs(ws, col, r0, r1, by_blank=False):
    """Gop doc cac o lien tiep cung gia tri (hoac o dau + cac o trong ben duoi)."""
    r = r0
    while r <= r1:
        v = ws.cell(r, col).value
        e = r
        while e + 1 <= r1:
            nv = ws.cell(e + 1, col).value
            if (by_blank and nv in (None, "")) or (not by_blank and nv == v):
                e += 1
            else:
                break
        if e > r and v not in (None, ""):
            ws.merge_cells(start_row=r, start_column=col, end_row=e, end_column=col)
        r = e + 1


def header_col(ws, text, row=1):
    for c in range(1, ws.max_column + 1):
        if text.lower() in str(ws.cell(row, c).value or "").lower():
            return c
    return None


def mark_class_sizes(ws, header_rows=1):
    """Bui: cot dem hat dinh dang so '#,##0' (Windows nguoi dung tach nhom bang dau cach -> '1 637 680');
    in dam cac co hat duoc danh gia theo cap sach yeu cau cua tung dong (ISO 14644-1, vd ISO 8 -> 0,5 / 1 / 5 µm)."""
    from build_cleanroom_132115 import ISO_CLASS_LIMITS
    size_col = {}
    for c in range(1, ws.max_column + 1):
        m = re.search(r"([\d.]+)\s*µm", str(ws.cell(1, c).value or ""))
        if m:
            size_col[m.group(1)] = c
    col_cls = header_col(ws, "Cấp sạch")
    evaluated = set()
    for r in range(header_rows + 1, last_row(ws) + 1):
        m = re.search(r"\d+", str(ws.cell(r, col_cls).value or ""))
        limits = ISO_CLASS_LIMITS.get(int(m.group())) if m else None
        for size, c in size_col.items():
            cell = ws.cell(r, c)
            cell.number_format = "#,##0"
            if limits and limits.get(size) is not None:
                cell.font = Font(name=cell.font.name, size=cell.font.sz, bold=True, color=cell.font.color)
                evaluated.add(c)
    # File nguon an cot Ch6 5µm (AA:AC): co hat duoc danh gia thi phai hien tren phieu
    for c in evaluated:
        span = next((r for r in ws.merged_cells.ranges if r.min_row == 1 and r.min_col <= c <= r.max_col), None)
        unhide_cols(ws, span.min_col if span else c, span.max_col if span else c)


def unhide_cols(ws, c0, c1):
    """Hien cot c0..c1; nhom cot an trum ra ngoai khoang nay van an (tach nhom)."""
    for key, dim in list(ws.column_dimensions.items()):
        a = dim.min or column_index_from_string(key)
        b = dim.max or a
        if not dim.hidden or b < c0 or a > c1:
            continue
        width = dim.width
        del ws.column_dimensions[key]
        for x, y, hidden in ((a, c0 - 1, True), (max(a, c0), min(b, c1), False), (c1 + 1, b, True)):
            if x <= y:
                L = get_column_letter(x)
                ws.column_dimensions[L] = ColumnDimension(ws, index=L, min=x, max=y, width=width, hidden=hidden)


def add_room_average(ws, blocks):
    """Nhiệt Ẩm: them 2 cot 'Trung bình phòng' (Nhiệt độ, Độ ẩm) sau cot 'Trung bình' tung diem, cong thuc
    =AVERAGE(...) gop theo phong. Nguoi dung chot 30/10/2026: danh gia nhiet do, do am theo trung binh ca phong.
    Chay lai nhieu lan khong sinh them cot (ghi de cung vi tri)."""
    c_avg = header_col(ws, "Trung bình")
    if not c_avg or header_col(ws, "Trung bình phòng") not in (None, c_avg + 2):
        return
    ct, ch = c_avg + 2, c_avg + 3
    L = get_column_letter
    span = f"{L(ct)}1:{L(ch)}1"
    if span not in [str(r) for r in ws.merged_cells.ranges]:
        ws.merge_cells(span)
    ws.cell(1, ct).value = "Trung bình phòng"
    ws.cell(2, ct).value = "Nhiệt độ"
    ws.cell(2, ch).value = "Độ ẩm"
    for r in (1, 2):
        for c in (ct, ch):
            ws.cell(r, c)._style = copy.copy(ws.cell(r, c_avg)._style)
    for c in (ct, ch):
        ws.column_dimensions[L(c)].width = ws.column_dimensions[L(c_avg)].width or 12
    for b0, b1 in blocks:
        for c, src in ((ct, c_avg), (ch, c_avg + 1)):
            ws.cell(b0, c).value = f"=AVERAGE({L(src)}{b0}:{L(src)}{b1})"
            ws.cell(b0, c)._style = copy.copy(ws.cell(b0, src)._style)
            if b1 > b0:
                ws.merge_cells(start_row=b0, start_column=c, end_row=b1, end_column=c)


def unmerge_from(ws, first_row):
    for rng in [r for r in ws.merged_cells.ranges if r.min_row >= first_row]:
        ws.unmerge_cells(str(rng))


def grid(ws, header_rows, data_size):
    lr, lc = last_row(ws), last_col(ws)
    for r in range(1, lr + 1):
        for c in range(1, lc + 1):
            cell = ws.cell(r, c)
            cell.border = GRID
            f = cell.font
            size = f.sz if r <= header_rows and f.sz else data_size
            cell.font = Font(name=FONT, size=size, bold=f.b if r <= header_rows else False,
                             italic=f.i, color=f.color)
            h = cell.alignment.horizontal or "center"
            cell.alignment = Alignment(horizontal=h, vertical="center",
                                       wrap_text=True if r <= header_rows else cell.alignment.wrap_text)
    # page setup: A4 ngang, vua 1 trang ngang, lap lai dong tieu de
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.print_title_rows = f"1:{header_rows}"       # lap lai dong tieu de cot o moi trang in
    ws.freeze_panes = f"A{header_rows + 1}"        # giu dong tieu de cot khi cuon
    ws.print_area = f"A1:{get_column_letter(lc)}{lr}"
    return lr, lc


def main(path, out=None):
    out = out or path
    wb = openpyxl.load_workbook(path)

    # --- Bụi 0,3: gop o theo dung cau truc dong tieu de cho moi dong du lieu
    ws = wb["Bụi 0,3"]
    lr = last_row(ws)
    groups = [(r.min_col, r.max_col) for r in ws.merged_cells.ranges
              if r.min_row == 1 and r.max_row == 1]
    for rng in [r for r in ws.merged_cells.ranges if r.min_row >= 2]:
        ws.unmerge_cells(str(rng))
    for r in range(2, lr + 1):
        for c0, c1 in groups:
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
    merge_runs(ws, 1, 2, lr, by_blank=True)   # cot Phong
    grid(ws, header_rows=HEADER_ROWS[ws.title], data_size=11)
    mark_class_sizes(ws, HEADER_ROWS[ws.title])

    # --- Nhiệt Ẩm: tieu de STT/Phong/Vi tri gop 2 dong, gop ten phong
    ws = wb["Nhiệt Ẩm"]
    for col in (1, 2, 3):
        if ws.cell(1, col).value in (None, "") and ws.cell(2, col).value:
            ws.cell(1, col).value = ws.cell(2, col).value
            ws.cell(1, col).font = copy.copy(ws.cell(2, col).font)
            ws.cell(2, col).value = None
            ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
    # gop ten phong (o dau + cac o trong ben duoi) va cot muc danh gia trong pham vi tung phong
    lr_na = last_row(ws)
    unmerge_from(ws, 3)
    blocks, start = [], None
    for r in range(3, lr_na + 1):
        if ws.cell(r, 2).value not in (None, ""):
            if start:
                blocks.append((start, r - 1))
            start = r
    if start:
        blocks.append((start, lr_na))
    for b0, b1 in blocks:
        if b1 > b0:
            ws.merge_cells(start_row=b0, start_column=2, end_row=b1, end_column=2)
        for col in [c for c in (header_col(ws, "Mức đánh giá nhiệt"), header_col(ws, "Mức đánh giá ẩm")) if c]:
            # phieu da gop o (chay lai tren phieu ThucHien): o duoi trong sau khi tach gop -> lay muc o tren
            for r in range(b0 + 1, b1 + 1):
                if ws.cell(r, col).value in (None, ""):
                    ws.cell(r, col).value = ws.cell(r - 1, col).value
            merge_runs(ws, col, b0, b1)
    add_room_average(ws, blocks)
    grid(ws, header_rows=HEADER_ROWS[ws.title], data_size=13)

    # --- Lưu lượng, Chênh áp
    for name in ("Lưu lượng", "Chênh áp"):
        grid(wb[name], header_rows=HEADER_ROWS[name], data_size=13)

    # --- Hien thi: do rong cot, xuong dong, dinh dang so (chi hien thi, cong thuc giu nguyen)
    def fmt(ws, cols, f, r0):
        for col in cols:
            for r in range(r0, last_row(ws) + 1):
                ws[f"{col}{r}"].number_format = f

    def wrap(ws, col, r0):
        for r in range(r0, last_row(ws) + 1):
            a = ws[f"{col}{r}"].alignment
            ws[f"{col}{r}"].alignment = Alignment(horizontal=a.horizontal, vertical="center", wrap_text=True)

    ws = wb["Bụi 0,3"]
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["K"].width = 19
    wrap(ws, "A", 2)
    ws = wb["Nhiệt Ẩm"]
    wrap(ws, "B", 3)
    c_cal = header_col(ws, "Sau hiệu chuẩn")
    if c_cal:
        fmt(ws, [get_column_letter(c) for c in range(c_cal, last_col(ws) + 1)], "0.0", 3)
    ws = wb["Lưu lượng"]
    ws.column_dimensions["A"].width = 27
    ws.column_dimensions["B"].width = 14
    fmt(ws, "B", "dd/mm/yyyy", 2)
    fmt(ws, "GHIJLMN", "0.0", 2)
    fmt(ws, "K", "0.00", 2)
    ws = wb["Chênh áp"]
    c_avg = header_col(ws, "Trung bình")
    if c_avg:
        fmt(ws, get_column_letter(c_avg), "0.0", 2)

    for ws in wb.worksheets:
        set_header(ws)
    wb.save(out)


if __name__ == "__main__":
    main(*sys.argv[1:3])
