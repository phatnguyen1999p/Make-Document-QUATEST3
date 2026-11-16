# -*- coding: utf-8 -*-
"""Tao phieu tinh toan ThucHien\\Q326D132115 - TinhToanPhongSach.xlsx tu file so lieu nguoi dung dien.

Nguon (KHONG ghi de): Input\\132115\\132115 - Data Do\\Q326D132115 - TinhToanPhongSach.xlsx
  - Giu nguyen sheet "Nhiệt Ẩm", "Lưu lượng", "Chênh áp" (so do + cot muc danh gia nguoi dung dien, cong thuc).
  - Sheet "Bụi 0,3": dung lai cot "Cấp sạch yêu cầu" cua tung phong, nhung cac dong mau duoc DUNG LAI theo
    so mau cua phieu ghi chep hien truong (ROOMS trong build_cleanroom_132115: zero_sample / start_sample),
    so dem lay tu 6ChannelFixed.xlsx. (File nguon gan nham mau da gach bo, vd P. Tron 912-920.)
Dich: Input\\132115\\ThucHien\\Q326D132115 - TinhToanPhongSach.xlsx (ghi file tam roi os.replace).
Tu 30/10/2026: phieu ThucHien da co thi KHONG dung lai tu file nguon (nguoi dung sua so lieu truc tiep tren phieu
ThucHien, vd do am, diem chenh ap PA12) -> chi chuan hoa dinh dang (format_excel_132115). Dung lai tu dau: --tao-lai.
"""
import copy
import os
import shutil
import sys

import openpyxl

from build_cleanroom_132115 import ROOMS, DATA_XLSX, CODE, _room_of

RAW_XLSX = os.path.join("Input", "132115", "132115 - Data Do", "6ChannelFixed.xlsx")
DEST_DIR = os.path.join("Input", "132115", "ThucHien")
DEST = os.path.join(DEST_DIR, f"{CODE} - TinhToanPhongSach.xlsx")

if os.path.isfile(DEST) and "--tao-lai" not in sys.argv:
    from format_excel_132115 import main as _format_excel
    tmp = DEST + ".tmp.xlsx"
    shutil.copy2(DEST, tmp)
    _format_excel(tmp)
    os.replace(tmp, DEST)          # file tam + os.replace: khong ghi xuyen qua hard link
    print("Phieu ThucHien da co (nguoi dung sua tren phieu): chi chuan hoa dinh dang ->", DEST)
    sys.exit(0)

wb = openpyxl.load_workbook(DATA_XLSX)
raw_ws = openpyxl.load_workbook(RAW_XLSX).active

samples_by_id = {}
for r in range(23, raw_ws.max_row + 1):
    sid = raw_ws.cell(r, 3).value
    if sid is not None:
        samples_by_id[int(sid)] = [raw_ws.cell(r, c).value for c in range(3, 39)]
print(f"Total samples in 6ChannelFixed: {len(samples_by_id)}")

# ---- Sheet Bụi 0,3
ws = wb["Bụi 0,3"]
hdr = {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
col_cls = next(c for k, c in hdr.items() if k.startswith("Cấp sạch"))
col_sid = hdr["#"]                      # so mau; cac cot sau la du lieu 6ChannelFixed theo dung thu tu
max_col = ws.max_column

# cap sach + style tung cot (lay tu dong du lieu dau tien) truoc khi xoa
cls_by_room, cur = {}, None
for r in range(2, ws.max_row + 1):
    if ws.cell(r, 1).value not in (None, ""):
        cur = ws.cell(r, 1).value
    rr = _room_of(cur)
    if rr and ws.cell(r, col_cls).value and rr["name_vn"] not in cls_by_room:
        cls_by_room[rr["name_vn"]] = ws.cell(r, col_cls).value
styles = {c: copy.copy(ws.cell(3, c)._style) for c in range(1, max_col + 1)}
row_h = ws.row_dimensions[3].height

for rng in [x for x in ws.merged_cells.ranges if x.min_row >= 2]:
    ws.unmerge_cells(str(rng))
ws.delete_rows(2, ws.max_row - 1)

cur_row = 2
missing = []
for room in ROOMS:
    pts = [("Zero", room["zero_sample"])] + [(str(i), room["start_sample"] + i - 1) for i in range(1, room["points_count"] + 1)]
    for idx, (label, sid) in enumerate(pts):
        for c in range(1, max_col + 1):
            ws.cell(cur_row, c)._style = copy.copy(styles[c])
        if row_h:
            ws.row_dimensions[cur_row].height = row_h
        if idx == 0:
            ws.cell(cur_row, 1, room["name_vn"])
        ws.cell(cur_row, 2, label)
        ws.cell(cur_row, col_cls, cls_by_room.get(room["name_vn"]))
        data = samples_by_id.get(sid)
        if data is None:
            missing.append(sid)
            ws.cell(cur_row, col_sid, sid)
        else:
            for k, v in enumerate(data):
                ws.cell(cur_row, col_sid + k, v)
        cur_row += 1
if missing:
    raise SystemExit(f"[FATAL] 6ChannelFixed thieu so mau {missing}")
print(f"Bụi 0,3: {cur_row - 2} dong theo so mau phieu ghi chep (cap sach giu theo file nguon: {cls_by_room})")

# ---- Luu vao ThucHien (file tam + os.replace: khong ghi de len file nguon neu hai file la hard link)
os.makedirs(DEST_DIR, exist_ok=True)
tmp = DEST + ".tmp.xlsx"
wb.save(tmp)

from format_excel_132115 import main as _format_excel
_format_excel(tmp)
os.replace(tmp, DEST)
print("Saved calculation sheet to ThucHien:", DEST)
