# -*- coding: utf-8 -*-
"""
Viền ô dùng chung cho các bảng tính Excel tính toán (TinhToanBSC / TuCay / Fumehood / KhiNen ...).

- thin_outside_borders(wb): đổi mọi viền dày (medium / thick) thành viền mảnh (thin), giữ nguyên màu.
  Người dùng chốt 30/10/2026: chỉ dùng "Outside Borders" thường, không dùng "Thick Outside Borders".
- strike_box(ws, "C114:I118"): khung chỉ tiêu KHÔNG ĐO -> xóa số, gộp ô, viền ngoài mảnh + gạch chéo
  (như khung Phương pháp 2 khi đo inflow bằng phương pháp 1, AGENTS mục 8B).
"""

from openpyxl.styles import Border, Side
from openpyxl.utils.cell import range_boundaries

THIN = Side(style="thin", color="000000")
_THICK_STYLES = {"medium", "thick"}


def _thin(side: Side) -> Side:
    if side is not None and side.style in _THICK_STYLES:
        return Side(style="thin", color=side.color)
    return side


def thin_outside_borders(wb) -> int:
    """Đổi viền medium/thick -> thin trên mọi sheet. Trả về số ô đã đổi."""
    n = 0
    for ws in wb.worksheets:
        for cell in list(ws._cells.values()):
            b = cell.border
            sides = (b.left, b.right, b.top, b.bottom, b.diagonal)
            if not any(s is not None and s.style in _THICK_STYLES for s in sides):
                continue
            cell.border = Border(left=_thin(b.left), right=_thin(b.right), top=_thin(b.top),
                                 bottom=_thin(b.bottom), diagonal=_thin(b.diagonal),
                                 diagonalUp=b.diagonalUp, diagonalDown=b.diagonalDown,
                                 outline=b.outline, vertical=b.vertical, horizontal=b.horizontal)
            n += 1
    return n


def strike_box(ws, cell_range: str) -> None:
    """Chỉ tiêu không đo: gộp vùng lưới đo, xóa số, viền ngoài mảnh và gạch chéo từ góc trên trái xuống góc dưới phải."""
    min_c, min_r, max_c, max_r = range_boundaries(cell_range)
    for m in list(ws.merged_cells.ranges):
        if not (m.max_row < min_r or m.min_row > max_r or m.max_col < min_c or m.min_col > max_c):
            ws.unmerge_cells(str(m))
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            ws.cell(r, c).value = None
    ws.merge_cells(cell_range)
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            first = (r == min_r and c == min_c)
            ws.cell(r, c).border = Border(
                top=THIN if r == min_r else None, bottom=THIN if r == max_r else None,
                left=THIN if c == min_c else None, right=THIN if c == max_c else None,
                diagonal=THIN if first else None, diagonalDown=first)
