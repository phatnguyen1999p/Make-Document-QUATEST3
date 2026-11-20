# -*- coding: utf-8 -*-
r"""
Q326D131873 — Công ty Cổ phần Công nghệ PPJ-HNC — bùn thải, 03 ngày lấy mẫu (B1 29/10, B2 30/10, B3 31/10/2026),
01 chứng thư (mã trần) + BCKQ. CHƯA CÓ KQTN -> bản khung, kết quả "………" tô vàng.

Nguồn (chép từ Z `0.VIMCERTS078\Q326D131873 - CONG NGHE PPJ-HNC - BUN` về Input\131873, chỉ đọc):
  - `0.  BBGDHT .pdf` (6 trang = 3 BBLM số Q326D131873/1, /2, /3): PYC Q326D131873 ngày 16/10/2026; "CÔNG TY CỔ PHẦN CÔNG
    NGHỆ PPJ-HNC", "Cụm Công nghiệp Tam Quan, phường Tam Quan, tỉnh Gia Lai"; "Lấy mẫu và phân tích bùn thải"; địa điểm:
    nhà máy của Công ty tại địa chỉ trên; thiết bị "-"; QUATEST3 3030:2025, N2.TC.B.01; nhà máy hoạt động sản xuất;
    "Bùn thải sau khi ép từ máy ép bùn", mỗi ngày 03 mẫu đơn (1 máy ép + 2 bao chứa) trộn đều gộp 01 mẫu:
      B1 29/10/2026 16:00 – 16:30 (16:00 máy ép, 16:15 bao chứa, 16:30 bao chứa), niêm phong KT372818;
      B2 30/10/2026 11:15 – 11:45 (11:15, 11:30, 11:45), KT372819;
      B3 31/10/2026 07:45 – 08:15 (07:45, 08:00, 08:15), KT372820;
    ghi chú: đại diện Công ty vận hành máy ép bùn để lấy mẫu; trang 2: "Hệ thống xử lý nước thải của Công ty"; đại diện
    KH ký "Mai Thị Thân" (báo giá: người liên hệ Mai Thân); người thực hiện Phạm Quốc Nghĩa.
  - `0.  BAO GIA  .pdf` (QONV26016234, 15/10/2026): chỉ tiêu theo QCVN 07:2025/BNNMT — pH, độ ẩm; ngâm chiết Cd, Ni, Ba,
    Co, Hg, As, Pb, Se, phenol, benzen, clo benzen; hàm lượng tuyệt đối tổng xyanua; theo yêu cầu KH: hàm lượng tuyệt đối
    và ngâm chiết Crom VI, tổng dầu; 01 mẫu x 3 ngày = 03 mẫu; chứng thư tiếng Việt.
  - `3. TBKQ  .docx`, `2. BCKQ  .docx`, `1. tinh toan chat thai  .xlsx`: khung chép từ hồ sơ Q326D112119 (Tân Đức) —
    phương pháp, ngưỡng QCVN 07:2025 lấy từ đây.
  - Ảnh `0.  03 VI TRI - 29-10/30-10/31-10  .jpg`, `0.  MAU B1/B2/B3  .jpg`.

Quyết định của người dùng (17/11/2026): "chỉ làm 1 bộ KQ chứng thư thôi chứ không tách ra làm 3 … thể hiện 03 ngày lấy
mẫu, mỗi ngày 01 mẫu B1-B2-B3/ kèm số niêm. Kết quả gồm 03 cột KQTN + 01 cột tính trung bình + 01 cột ngưỡng QC
07:2025; 02 chỉ tiêu Crom VI + tổng dầu tách riêng thành 01 bảng - vì không thuộc QC 07:2025".

Quyết định của người dùng (17/11/2026, lần 2):
  - Chưa có KQTN -> để trống. Trung bình (để trống tới khi có KQTN): ngày '< LOD' tính 0, cộng các ngày có số chia 3;
    nhỏ hơn LOD ghi '< LOD', lớn hơn ghi trị số (hàm `trung_binh`). Đánh giá: so giá trị trung bình với ngưỡng.
  - Ngưỡng: "File ngưỡng có trong thư mục, kiếm nó và áp vào … thiếu thì tô vàng và để trống". Trên máy / Z không có
    file QCVN 07:2025 (chỉ QCVN 07:2009 trên Z `TIÊU CHUẨN VIỆT NAM. Tong hop cac quy chuan moi truong`) -> AI lấy
    ngưỡng theo các chứng thư QCVN 07:2025 đã làm năm 2026 (017976, 020756, 069814, 107520…): clo benzen ≥ 100, selen
    > 1; tổng xyanua ≥ 590 (khớp QCVN 07:2009 bảng 3; khung 112119 ghi ≥ 15 là sai).
  - Bảng Crom VI + tổng dầu: thêm câu VIMCERTS 078, tô vàng. Người chứng kiến: Mai Thị Thân (đúng).
  - Phương pháp (17/11/2026, lần 3): "phần phương pháp trong TBKQ hãy để trống chung với kết quả, sau này có Kết quả
    rồi thì dựa theo phương pháp và kết quả để thay thế" -> cột phương pháp "………" tô vàng (hàm `meth`, điền `METHODS`
    theo phiếu KQTN khi về); phương pháp trong LEACH / ABSOLUTE / EXTRA chỉ để tham khảo.

  - Bản TBKQ người dùng sửa tay trên Z (17/11/2026 18:02) — "Giữ sửa của tôi" (18/11/2026): bảng Ghi chú 9.2.3 câu
    VIMCERTS cùng dòng "Ghi chú:", bỏ dòng "●", bảng 5103 pct (hàm `user_note_923`).

Suy luận của AI (cần người dùng xác nhận):
  1. (bỏ — phương pháp để trống tới khi có KQTN, xem quyết định lần 3)
  2. Bảng Crom VI + tổng dầu: 03 cột KQTN + cột trung bình, không cột ngưỡng, không đánh giá.

Chạy:  python build_131873.py [--pdf]
"""
import argparse
import datetime
import os
import nuoc_common as N
from nuoc_common import WORKSPACE, PLACEHOLDER as P
import build_nuoc_mau as M
from build_nuoc_mau import Body, cell_text, note_table, fix_cover, appendix_protos, make_head, _keep, _keep_rows
from docx import Document
from docx.oxml.ns import qn
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import merge_same_vertical, style_note_tables, appendix_total_field
from Core.Styling.elements import add_document_ending_line
from Core.Styling.tables import set_table_autofit_window, insert_image_to_cell, add_cant_split

IN = os.path.join(WORKSPACE, "Input", "131873")
CLIENT = "Công ty Cổ phần Công nghệ PPJ-HNC"
ADDR = "Cụm Công nghiệp Tam Quan, phường Tam Quan, tỉnh Gia Lai"
STD = "QCVN 07:2025/BNNMT - Quy chuẩn kỹ thuật quốc gia về ngưỡng chất thải nguy hại"
QC = "QCVN 07:2025/BNNMT"
WITNESS = f"Đại diện {CLIENT}\n(chị Mai Thị Thân)"
SAMPLES = [  # ký hiệu, ngày, khung giờ, số niêm phong, ảnh vị trí, ảnh mẫu
    ("B1", "29/10/2026", "16:00 – 16:30", "KT372818", "0.  03 VI TRI - 29-10  .jpg", "0.  MAU B1  .jpg"),
    ("B2", "30/10/2026", "11:15 – 11:45", "KT372819", "0.  03 VI TRI - 30-10  .jpg", "0.  MAU B2  .jpg"),
    ("B3", "31/10/2026", "07:45 – 08:15", "KT372820", "0.  03 VI TRI - 31-10  .jpg", "0.  MAU B3  .jpg"),
]
VI_TRI = ("Bùn thải sau khi ép từ máy ép bùn của hệ thống xử lý nước thải (mỗi mẫu trộn đều từ 03 mẫu đơn: 01 mẫu tại "
          "máy ép bùn, 02 mẫu tại bao chứa bùn)")
CHI_TIEU = ("- Tính chất nguy hại: tính kiềm/ tính axit (pH);\n"
            "- Thành phần nguy hại theo nồng độ ngâm chiết: cadimi, niken, bari, coban, thủy ngân, asen, chì, selen, "
            "phenol, benzen, clo benzen;\n"
            "- Thành phần nguy hại theo hàm lượng tuyệt đối: tổng xyanua;\n"
            "- Crom VI, tổng dầu (nồng độ ngâm chiết và hàm lượng tuyệt đối) theo yêu cầu khách hàng.")
ICP = "US EPA 1311 & US EPA 200.7"
LEACH = [  # (tên, phương pháp, ngưỡng CTNH mg/L, tô vàng ngưỡng?)
    ("Cadimi (Cd)", ICP, "≥ 1,0", False), ("Niken (Ni)", ICP, "≥ 70", False), ("Bari (Ba)", ICP, "≥ 100", False),
    ("Coban (Co)", ICP, "≥ 80", False), ("Thủy ngân (Hg)", "US EPA 1311 & SMEWW 3112B:2023", "≥ 0,2", False),
    ("Asen (As)", ICP, "≥ 5,0", False), ("Chì (Pb)", ICP, "≥ 5,0", False), ("Selen (Se)", ICP, "> 1", False),
    ("Phenol", "US EPA 1311 & US EPA 3510C & US EPA 3650B & US EPA 8041A", "≥ 1 000", False),
    ("Benzen", "US EPA 1311 & US EPA 5021A", "≥ 0,5", False),
    ("Clo benzen", "US EPA 1311 & US EPA 5021A", "≥ 100", False),     # theo 069814, 107520 (QCVN 07:2025)
]
# tổng xyanua ≥ 590 ppm: QCVN 07:2009 bảng 3 + 017976, 069814, 069820, 070191 (khung 112119 ghi ≥ 15 — bỏ)
ABSOLUTE = [("Tổng xyanua (CN⁻)", "US EPA 9013A & US EPA 9010C & US EPA 9014", "≥ 590", False)]
EXTRA = [  # Crom VI, tổng dầu — theo yêu cầu khách hàng (không thuộc QCVN 07:2025); phương pháp ngâm chiết theo 005526
    ("Crom VI (Cr⁶⁺) – nồng độ ngâm chiết", "mg/L", "US EPA 1311 & SMEWW 3500-Cr B:2023"),
    ("Crom VI (Cr⁶⁺) – hàm lượng tuyệt đối", "ppm", P),
    ("Tổng dầu – nồng độ ngâm chiết", "mg/L", "US EPA 1311 & SMEWW 5520B:2023"),
    ("Tổng dầu – hàm lượng tuyệt đối", "ppm", P),
]
# Kết quả KQTN (chưa có — để trống): {tên chỉ tiêu: [B1, B2, B3]}, vd {"Cadimi (Cd)": ["< 0,01", "0,02", "< 0,01"]}
RESULTS = {}


def trung_binh(vals):
    """Trung bình 3 ngày (người dùng chốt 17/11/2026): giá trị '< LOD' tính là 0, cộng các ngày có số rồi chia 3;
    trung bình nhỏ hơn LOD -> ghi '< LOD', lớn hơn -> ghi trị số đó. Cả 3 ngày '< LOD' -> '< LOD'.
    Chưa đủ 3 kết quả -> '………'."""
    if len(vals) != 3 or any(not v or P in v for v in vals):
        return P
    lods, nums, dec = [], [], 0
    for v in vals:
        s = v.replace(" ", "").replace(" ", "")
        if s.startswith("<"):
            lods.append(float(s[1:].replace(",", ".")))
        else:
            nums.append(float(s.replace(",", ".")))
            dec = max(dec, len(s.split(",")[1]) if "," in s else 0)
    lod = max(lods) if lods else None
    avg = sum(nums) / 3
    if lod is not None and avg < lod:
        return "< " + (f"{lod:g}".replace(".", ","))
    return f"{avg:.{dec}f}".replace(".", ",")


def user_note_923(nt):
    """Bảng Ghi chú 9.2.3 theo bản người dùng sửa tay trên Z (17/11/2026 18:02, người dùng "Giữ sửa của tôi" 18/11/2026):
    câu VIMCERTS đưa lên cùng dòng "Ghi chú:", bỏ dòng "●"; bảng rộng 5103 pct, lưới cột 970 / 8578."""
    trs = nt.findall(qn("w:tr"))
    vim_tc = trs[1].findall(qn("w:tc"))[1]
    head_tc = trs[0].findall(qn("w:tc"))[1]
    for p in head_tc.findall(qn("w:p")):
        head_tc.remove(p)
    for p in vim_tc.findall(qn("w:p")):
        head_tc.append(p)
    nt.remove(trs[1])
    pr = nt.find(qn("w:tblPr"))
    w = pr.find(qn("w:tblW"))
    if w is not None:
        w.set(qn("w:w"), "5103")
        w.set(qn("w:type"), "pct")
    for g, val in zip(nt.find(qn("w:tblGrid")), ("970", "8578")):
        g.set(qn("w:w"), val)


def meth(name, known):
    """Phương pháp: chưa có KQTN -> '………' tô vàng cùng kết quả (người dùng chốt 17/11/2026: "phần phương pháp trong
    TBKQ hãy để trống chung với kết quả, sau này có kết quả rồi thì dựa theo phương pháp và kết quả để thay thế").
    ``known`` (phương pháp theo hồ sơ cũ) chỉ để tham khảo — khi có KQTN ghi theo phiếu vào ``METHODS``."""
    return METHODS.get(name, P)


METHODS = {}   # phương pháp theo phiếu KQTN khi về: {tên chỉ tiêu: phương pháp}


def res(name):
    r = RESULTS.get(name, [P, P, P])
    return list(r) + [trung_binh(r)]

CONFIG = dict(
    code6="131873", loai="bun thai", kind="bt",
    cert_date=datetime.date.today().strftime("%d/%m/%Y"),   # chưa KQTN -> ngày dựng (AGENTS mục 2)
    client_vn=CLIENT, address_vn=ADDR, reg_date="16/10/2026",
    dates=[s[1] for s in SAMPLES],
    time_bckq="; ".join(f"Ngày {d} từ {t.replace(' – ', ' đến ')}" for _k, d, t, *_ in SAMPLES),
    site_vn=f"Nhà máy của {CLIENT}, {ADDR}",
    scope_vn="Lấy mẫu và giám định chất lượng bùn thải", procedure="QUATEST 3 3030:2025, N2.TC.B.01",
    equip=[], executor="Nghĩa", witness_vn="chị Mai Thị Thân – đại diện " + CLIENT, req_date=P,
    samples=[dict(ky_hieu=k, vi_tri=VI_TRI, time=f"{d}\n{t}", seal=s, chi_tieu=CHI_TIEU,
                  photo=os.path.join(IN, ph)) for k, d, t, s, ph, _m in SAMPLES],
)


def hl_cell(cell):
    for r_el in cell._tc.iter(qn("w:r")):
        N._set_hl(r_el)


def result_table(b, first_cols, widths, rows, last_head=None):
    """Bảng 2 dòng tiêu đề: first_cols | Kết quả (B1, B2, B3) | Trung bình | [last_head]."""
    ncol = len(first_cols) + 4 + (1 if last_head else 0)
    t = b.table(2 + len(rows), ncol, widths, header_rows=2)
    k = len(first_cols)
    for j, h in enumerate(first_cols):
        t.cell(0, j).merge(t.cell(1, j))
        cell_text(t.cell(0, j), h)
    t.cell(0, k).merge(t.cell(0, k + 2))
    cell_text(t.cell(0, k), "Kết quả")
    for j, (key, *_r) in enumerate(SAMPLES):
        cell_text(t.cell(1, k + j), key)
    t.cell(0, k + 3).merge(t.cell(1, k + 3))
    cell_text(t.cell(0, k + 3), "Trung bình")
    if last_head:
        t.cell(0, k + 4).merge(t.cell(1, k + 4))
        cell_text(t.cell(0, k + 4), last_head)
    for i, vals in enumerate(rows, start=2):
        for j, v in enumerate(vals):
            cell_text(t.cell(i, j), v, align="left" if j == (1 if first_cols[0] == "STT" else 0) else "center")
    return t


def build_tbkq(cfg):
    d = N.make_dossier(cfg, dict(standard_vn=STD), sub_no=None, multi=False)
    d.ensure_output_directories()
    src = CleanroomHandler(d).build_tbkq(is_en=False)
    doc = Document(src)
    fix_cover(doc, False, header_colon=False)
    text_proto, note_proto, head_proto = appendix_protos(doc)
    b = Body(doc, text_proto)
    head = make_head(b, head_proto)

    head("Nội dung thực hiện")
    b.para(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm theo yêu cầu của {CLIENT}, cụ thể như sau:", keep_next=True)
    t = b.table(4, 4, [30, 12, 16, 42])
    for j, h in enumerate(["Vị trí lấy mẫu", "Ký hiệu mẫu", "Ngày lấy mẫu", "Chỉ tiêu thử nghiệm"]):
        cell_text(t.rows[0].cells[j], h)
    for i, (k, dd, *_r) in enumerate(SAMPLES, start=1):
        cell_text(t.cell(i, 1), k)
        cell_text(t.cell(i, 2), dd)
    t.cell(1, 0).merge(t.cell(3, 0))
    cell_text(t.cell(1, 0), VI_TRI)
    t.cell(1, 3).merge(t.cell(3, 3))
    cell_text(t.cell(1, 3), CHI_TIEU, align="left")
    _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:-1])

    b.para("- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong:", keep_next=True, space_before=6)
    t = b.table(4, 4, [36, 14, 26, 24])
    for j, h in enumerate(["Người chứng kiến", "Ký hiệu mẫu", "Thời gian lấy mẫu", "Số niêm phong"]):
        cell_text(t.rows[0].cells[j], h)
    for i, (k, dd, tt, seal, *_r) in enumerate(SAMPLES, start=1):
        cell_text(t.cell(i, 1), k)
        cell_text(t.cell(i, 2), f"{dd}\n{tt}")
        cell_text(t.cell(i, 3), seal)
    t.cell(1, 0).merge(t.cell(3, 0))
    cell_text(t.cell(1, 0), WITNESS)
    _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:-1])

    b.para("- Hình ảnh vị trí lấy mẫu và mẫu niêm phong:", keep_next=True, space_before=6)
    photos = [os.path.join(IN, s[4]) for s in SAMPLES] + [os.path.join(IN, s[5]) for s in SAMPLES]
    pt = doc.add_table(rows=2, cols=3)
    b.emit(pt._tbl)
    set_table_autofit_window(pt)
    for n, ph in enumerate(photos):
        insert_image_to_cell(pt.rows[n // 3].cells[n % 3], ph, width=M.Cm(4.8))
    for r in pt.rows:
        add_cant_split(r)

    head("Kết quả giám định")
    vim = ("●", "Các thông số được chứng nhận VIMCERTS 078.")
    # 9.2.1
    M.sub_head(b, "9.2.1. Tính chất nguy hại")
    t = result_table(b, ["Tính chất nguy hại", "Phương pháp"], [22, 15, 12, 12, 12, 12, 15],
                     [["Tính kiềm/ tính axit (pH)", meth("pH", "US EPA 9045D")] + res("pH") + ["pH ≥ 12,5 hoặc\npH ≤ 2,0"]],
                     last_head="Ngưỡng CTNH ^(1)")
    _keep_rows(t._tbl, t._tbl.findall(qn("w:tr")))
    nt = note_table(b, note_proto, [vim, ("(1)", f"Ngưỡng chất thải nguy hại tại {QC}, bảng 1.")])
    _keep_rows(nt, nt.findall(qn("w:tr"))[:-1])

    def comp(title, rows, unit, mark, note):
        M.sub_head(b, title)
        data = [[str(i), n, meth(n, m)] + res(n) + [thr] for i, (n, m, thr, _h) in enumerate(rows, 1)]
        t = result_table(b, ["STT", "Thành phần nguy hại", "Phương pháp"], [6, 18, 26, 10, 10, 10, 9, 11], data,
                         last_head=f"Ngưỡng CTNH ^({mark}), {unit}")
        cell_text(t.cell(0, 3), f"Kết quả, {unit}")
        if not all(r[2] == P for r in data):        # cột phương pháp còn trống: không gộp dọc (tách trang được)
            merge_same_vertical(t, 2, 2)
        for i, (_n, m, _t, hl) in enumerate(rows, start=2):
            if hl:
                hl_cell(t.cell(i, 7))
            if m == P:
                hl_cell(t.cell(i, 2))
        trs = t._tbl.findall(qn("w:tr"))
        _keep_rows(t._tbl, trs[:5])
        _keep_rows(t._tbl, trs[-2:])
        nt = note_table(b, note_proto, [vim, (f"({mark})", note)])
        if mark == 3:
            user_note_923(nt)
        _keep_rows(nt, nt.findall(qn("w:tr"))[:-1])

    comp("9.2.2. Nồng độ ngâm chiết của các thành phần nguy hại", LEACH, "mg/L", 2,
         f"Ngưỡng chất thải nguy hại tính theo nồng độ ngâm chiết tại {QC}, bảng 2.")
    comp("9.2.3. Hàm lượng tuyệt đối của các thành phần nguy hại", ABSOLUTE, "ppm", 3,
         f"Ngưỡng chất thải nguy hại tính theo hàm lượng tuyệt đối tại {QC}, bảng 3.")

    # 9.2.4 Crom VI, tổng dầu — theo yêu cầu khách hàng
    M.sub_head(b, "9.2.4. Crom VI và tổng dầu (theo yêu cầu khách hàng)")
    data = [[str(i), n, u, meth(n, m)] + res(n) for i, (n, u, m) in enumerate(EXTRA, 1)]
    t = result_table(b, ["STT", "Chỉ tiêu", "Đơn vị", "Phương pháp"], [6, 26, 9, 23, 9, 9, 9, 9], data)
    trs = t._tbl.findall(qn("w:tr"))
    _keep_rows(t._tbl, trs[:-1])
    nt = note_table(b, note_proto, [vim, ("●", f"Chỉ tiêu thực hiện theo yêu cầu khách hàng, không thuộc {QC}; "
                                               "không đánh giá.")])
    for r_el in nt.findall(qn("w:tr"))[1].iter(qn("w:r")):   # VIMCERTS cho Crom VI, dầu: người dùng "thêm vào và
        N._set_hl(r_el)                                     # tô vàng" (17/11/2026)
    _keep_rows(nt, nt.findall(qn("w:tr")))         # Ghi chú đi liền phần Đánh giá (không để Đánh giá lẻ trang)

    b.para("Đánh giá:", bold=True, keep_next=True, space_before=6)
    b.para(f"So sánh giá trị trung bình kết quả thử nghiệm của 03 mẫu (B1, B2, B3) với các ngưỡng chất thải nguy hại "
           f"tương ứng nêu tại {QC}:", keep_next=True)
    for k, key in enumerate(("cao hơn mức quy định", "thấp hơn mức quy định")):
        q = b.para(parts=[("+ Chỉ tiêu ", False, False), (key, True, True), ("\t: " + P, False, False)],
                   indent_cm=0.75, tab_cm=7.5, keep_next=k == 0)
        _keep(q._p, lines=True)

    style_note_tables(doc)
    N._highlight_ph(doc)
    add_document_ending_line(doc)
    appendix_total_field(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - viet - bun thai - {d.code}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    print(f"[131873] TBKQ -> {os.path.relpath(dst, WORKSPACE)}")
    return dst


def bckq_site_block(cfg, samples, all_rows, b, doc):
    """Mục 4.5 BCKQ: bảng vị trí 3 ngày + niêm phong + ảnh."""
    b.para(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm được chỉ định bởi {cfg['witness_vn']}, cụ thể như sau:",
           keep_next=True)
    t = b.table(4, 5, [27, 10, 17, 14, 32])
    for j, h in enumerate(["Vị trí thực hiện", "Kí hiệu mẫu", "Thời gian lấy mẫu", "Số niêm phong", "Chỉ tiêu thử nghiệm"]):
        cell_text(t.rows[0].cells[j], h)
    for i, (k, dd, tt, seal, *_r) in enumerate(SAMPLES, start=1):
        cell_text(t.cell(i, 1), k)
        cell_text(t.cell(i, 2), f"{dd}\n{tt}")
        cell_text(t.cell(i, 3), seal)
    t.cell(1, 0).merge(t.cell(3, 0))
    cell_text(t.cell(1, 0), VI_TRI, align="left")
    t.cell(1, 4).merge(t.cell(3, 4))
    cell_text(t.cell(1, 4), CHI_TIEU, align="left")
    b.para("- Thời gian lấy 03 mẫu đơn của từng mẫu gộp: B1 16:00 (máy ép), 16:15, 16:30 (bao chứa); B2 11:15, 11:30, "
           "11:45; B3 07:45, 08:00, 08:15. Đại diện Công ty vận hành máy ép bùn để lấy mẫu.", space_before=6)
    b.para("- Hình ảnh vị trí lấy mẫu và mẫu niêm phong:", keep_next=True, space_before=6)
    photos = [os.path.join(IN, s[4]) for s in SAMPLES] + [os.path.join(IN, s[5]) for s in SAMPLES]
    pt = doc.add_table(rows=2, cols=3)
    b.emit(pt._tbl)
    set_table_autofit_window(pt)
    for n, ph in enumerate(photos):
        insert_image_to_cell(pt.rows[n // 3].cells[n % 3], ph, width=M.Cm(4.8))
    for r in pt.rows:
        add_cant_split(r)


def run(pdf=False):
    cfg = CONFIG
    docs = [build_tbkq(cfg)]
    N._bckq_site_block = bckq_site_block             # khuôn 4.5 riêng của hồ sơ (3 ngày, niêm phong)
    docs.insert(0, N.build_bckq(cfg, [[] for _ in cfg["samples"]], [None] * 3, []))
    if pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(N.make_dossier(cfg, None, None, False)).export_files(docs)
    print("\n===== QA =====")
    for p in docs:
        res = N.check_single_docx(p, code=f"Q326D{cfg['code6']}")
        print(f"  {res.get('status')}\t{os.path.relpath(p, WORKSPACE)}")
        for i in res.get("issues", []):
            if "……" not in str(i) and "bôi vàng" not in str(i):
                print(f"      - {i}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    run(ap.parse_args().pdf)
