# -*- coding: utf-8 -*-
"""
Q326D131873 — Công ty Cổ phần Công nghệ PPJ-HNC — lấy mẫu và giám định bùn thải (QCVN 07:2025/BNNMT).

Nguồn (Input\\131873, chép từ Z ...\\2026\\0.VIMCERTS078\\Q326D131873 - CONG NGHE PPJ-HNC - BUN):
- `0.  BBGDHT .pdf` (6 trang): BBLM /1 (29/10/2026, 2 trang), /2 (30/10/2026, 2 trang), /3 (31/10/2026, 2 trang).
  Mỗi ngày 1 mẫu gộp từ 3 lần lấy (máy ép / bao chứa / bao chứa): B1 16:00–16:30 niêm phong KT372818;
  B2 11:15–11:45 KT372819; B3 07:45–08:15 KT372820. Người thực hiện Phạm Quốc Nghĩa; chứng kiến: đại diện công ty
  (chữ ký "Mai Thị Thân"); ghi tay: "Đại diện công ty vận hành máy ép bùn từ hệ thống xử lý nước thải của công ty".
  Quy trình `QUATEST3 3030:2025, N2.TC.B.01`; thiết bị `-`; PYC số Q326D131873 ngày 16/10/2026.
- `0.  BAO GIA  .pdf` (3 trang): báo giá QONV26016234 ngày 15/10/2026 — chỉ tiêu QCVN 07:2025/BNNMT: pH, độ ẩm;
  ngâm chiết Cd, Ni, Ba, Co, Hg, As, Pb, Se, Phenol, Benzen, Clo benzen; hàm lượng tuyệt đối tổng xyanua; theo yêu cầu
  KH: hàm lượng tuyệt đối và ngâm chiết của Crom VI, Tổng dầu; 01 mẫu × 3 ngày = 03 mẫu; chứng thư tiếng Việt.
- Ảnh `THUC HIEN\\D131873 - PHOTO\\PHOTO BAO CAO`: `MAU B1/B2/B3` (niêm phong), `03 VI TRI - 29-10/30-10/31-10`.
- `1. tinh toan chat thai.xlsx`, `2. BCKQ.docx`, `3. TBKQ.docx`: mẫu của hồ sơ Q326D112119 (Tân Đức) — chỉ tham khảo khuôn.
- CHƯA CÓ phiếu KQTN, chưa có PYC, phiếu ghi chép trong scan.

Quyết định của người dùng (17/11/2026):
1. "Tách 3 chứng thư trước, có gì sửa sau" -> /1 B1 (29/10), /2 B2 (30/10), /3 B3 (31/10); BCKQ gộp 1 file (mã trần).
2. Ngưỡng clo benzen, xyanua (≥ 590 theo 107520; mẫu 112119 ghi ≥ 15) -> "highlight chỗ đó" (tô vàng, HL_TEXTS).
3. Crom (VI), tổng dầu (ngưỡng `-`, "không đánh giá", phương pháp `………`) -> "giữ nguyên highlight chỗ đó".
4. Phương pháp theo 112119 -> "oke". 5. Độ ẩm không đưa vào TBKQ -> "oke".
6. Người chứng kiến: "Mai Thị Thân".
7. Ngày gửi mẫu, ngày yêu cầu cấp KQTN (chưa có PYC / KQTN) -> "để trống, highlight lên".

Suy luận của AI (cần người dùng xác nhận):
1. Chỉ dựng TBKQ tiếng Việt (báo giá: chi phí chứng thư tiếng Việt; không có PYC mục 15).

Chạy:  python build_131873.py          (DOCX + XLSX + Edoc, QA)
       python build_131873.py --pdf    (thêm PDF + giấy tiêu đề UpChungThu)
"""

import os
import sys
import argparse
import datetime

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ["QUATEST3_NETWORK_ROOT"] = r"C:\__khong_dong_bo_Z__"

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn

from Core.Models import Dossier, ClientInfo, PersonnelInfo
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER as P, merge_same_vertical, style_note_tables, appendix_total_field
from Core.Styling.elements import add_document_ending_line
from Core.Styling.tables import set_table_autofit_window, insert_image_to_cell, add_cant_split
from Core.Tools.qa_check import check_single_docx
from Core.Tools.kqtn_late import tre_han_text
from build_nuoc_mau import (Body, cell_text, note_table, fix_cover, appendix_protos, make_head, sub_head, _keep,
                            _keep_rows, _num, _ptext)

CODE6 = "131873"
CODE = f"Q326D{CODE6}"
REG_DATE = "16/10/2026"
DAYS = ["29/10/2026", "30/10/2026", "31/10/2026"]
TIMES = [("16:00", "16:30"), ("11:15", "11:45"), ("07:45", "08:15")]
SAMPLES = ["B1", "B2", "B3"]
SEALS = ["KT372818", "KT372819", "KT372820"]

# KQTN — điền khi có phiếu (None = chưa có -> ………)
KQTN_SO = None          # vd "MTTN26xxxxxx"
KQTN_NGAY = None        # dd/mm/yyyy
SEND_DATE = None        # "Ngày nhận mẫu" trên phiếu KQTN
REQ_DATE = None         # "Ngày dự kiến trả kết quả" trên PYC
CERT_DATE = KQTN_NGAY or P

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần Công nghệ PPJ-HNC",
    name_en="PPJ-HNC Technology Joint Stock Company",
    address_vn="Cụm công nghiệp Tam Quan, phường Tam Quan, tỉnh Gia Lai",
    address_en="Tam Quan Industrial Cluster, Tam Quan ward, Gia Lai province",
)
SITE_VN = "Nhà máy của Công ty Cổ phần Công nghệ PPJ-HNC, Cụm công nghiệp Tam Quan, phường Tam Quan, tỉnh Gia Lai"
WITNESS = "Đại diện Công ty Cổ phần Công nghệ PPJ-HNC\n(chị Mai Thị Thân)"
VI_TRI = "Bùn thải sau khi ép từ máy ép bùn của hệ thống xử lý nước thải"
CHI_TIEU = ("Tính chất nguy hại: tính kiềm/ tính axit (pH);\n"
            "Thành phần nguy hại theo nồng độ ngâm chiết: cadimi, nicken, bari, coban, thủy ngân, asen, chì, selen, "
            "phenol, benzen, clo benzen;\n"
            "Thành phần nguy hại theo hàm lượng tuyệt đối: tổng xyanua;\n"
            "Theo yêu cầu khách hàng: crom (VI), tổng dầu (nồng độ ngâm chiết và hàm lượng tuyệt đối).")

M_METAL = "US EPA 1311 & US EPA 200.7"
# kết quả theo mẫu B1, B2, B3 — None -> ………
PH = ([None, None, None], "pH ≥ 12,5 hoặc\npH ≤ 2,0")
LEACH = [  # (tên, phương pháp, [B1, B2, B3], ngưỡng mg/L)
    ("Cadimi (Cd)", M_METAL, [None] * 3, "≥ 1,0"),
    ("Nicken (Ni)", M_METAL, [None] * 3, "≥ 70"),
    ("Bari (Ba)", M_METAL, [None] * 3, "≥ 100"),
    ("Coban (Co)", M_METAL, [None] * 3, "≥ 80"),
    ("Thủy ngân (Hg)", "US EPA 1311 & SMEWW 3112B:2023", [None] * 3, "≥ 0,2"),
    ("Asen (As)", M_METAL, [None] * 3, "≥ 5,0"),
    ("Chì (Pb)", M_METAL, [None] * 3, "≥ 5,0"),
    ("Selen (Se)", M_METAL, [None] * 3, "> 1"),
    ("Phenol", "US EPA 1311 & US EPA 3510C & US EPA 3650B & US EPA 8041A", [None] * 3, "≥ 1 000"),
    ("Benzen", "US EPA 1311 & US EPA 5021A", [None] * 3, "≥ 0,5"),
    ("Clo benzen", P, [None] * 3, P),
    ("Crom (VI) ^(3)", P, [None] * 3, "-"),
    ("Tổng dầu ^(3)", P, [None] * 3, "-"),
]
ABS = [  # kết quả đã quy đổi mẫu rắn khô (ppm) — lấy từ phiếu Excel
    ("Tổng xyanua (CN⁻)", "US EPA 9013A & US EPA 9010C & US EPA 9014", [None] * 3, "≥ 590"),
    ("Crom (VI) ^(5)", P, [None] * 3, "-"),
    ("Tổng dầu ^(5)", P, [None] * 3, "-"),
]
MOISTURE = [None, None, None]
CUSTOMER_NOTE = "Chỉ tiêu thực hiện theo yêu cầu của khách hàng, không có ngưỡng tại QCVN 07:2025/BNNMT."

INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
PHOTO_DIR = os.path.join(INPUT_DIR, "THUC HIEN", "D131873 - PHOTO", "PHOTO BAO CAO")
SEAL_PHOTOS = [os.path.join(PHOTO_DIR, f"MAU {s}  .jpg") for s in SAMPLES]
LOC_PHOTOS = [os.path.join(PHOTO_DIR, f"03 VI TRI - {d[:5].replace('/', '-')}  .jpg") for d in DAYS]


def v(x):
    return P if x is None else x


def make_cert_dossiers():
    """3 chứng thư /1 /2 /3 — mỗi chứng thư 1 mẫu, 1 ngày (người dùng chốt 17/11/2026)."""
    from Core.Models import assign_cert_numbers
    ds = []
    for k, (day, (a, e)) in enumerate(zip(DAYS, TIMES)):
        d = make_dossier()
        d.sample_date, d.inspection_dates = day, []
        d.sample_time_vn, d.sample_time_en = f"Ngày {day} từ {a} đến {e}", f"From {day}"
        d.sub_no, d.is_multi_cert = None, True
        ds.append(d)
    return assign_cert_numbers(ds)


def make_dossier():
    time_vn = "; ".join(f"{'Ngày' if i == 0 else 'ngày'} {d} từ {a} đến {b}" for i, (d, (a, b)) in enumerate(zip(DAYS, TIMES)))
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=DAYS[0], reg_date=REG_DATE,
        sample_time_vn=time_vn, sample_time_en=f"From {DAYS[0]}", inspection_dates=list(DAYS),
        sub_no=None, is_multi_cert=False, dossier_type="cleanroom",
        scope_vn="Lấy mẫu và giám định chất lượng bùn thải", scope_en="Sampling and quality inspection of sludge",
        location_desc_vn=SITE_VN, location_desc_en=SITE_VN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.B.01", procedure_en="QUATEST 3 3030:2025, N2.TC.B.01",
        standards_vn="QCVN 07:2025/BNNMT - Quy chuẩn kỹ thuật quốc gia về ngưỡng chất thải nguy hại",
        standards_en="QCVN 07:2025/BNNMT - Quy chuẩn kỹ thuật quốc gia về ngưỡng chất thải nguy hại",
        client=CLIENT, personnel=PersonnelInfo(executor="Nghĩa"),
        equipments=[], locations=[], photos=[], results_data={},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
def photo_row(body, paths, width_cm=4.8):
    t = body.doc.add_table(rows=1, cols=len(paths))
    body.emit(t._tbl)
    set_table_autofit_window(t)
    for c, p in zip(t.rows[0].cells, paths):
        insert_image_to_cell(c, p, width=__import__("docx.shared", fromlist=["Cm"]).Cm(width_cm))
    add_cant_split(t.rows[0])
    return t


def sample_table(body, lead, rows, unit_head, limit_head, widths, samples=SAMPLES):
    """Bảng: lead cột | Kết quả (1 mẫu: 1 dòng tiêu đề; nhiều mẫu: dòng 2 ghi ký hiệu) | ngưỡng.
    rows = [(cells_before…, [kết quả theo samples], ngưỡng)]."""
    nl, ns = len(lead), len(samples)
    ncol = nl + ns + 1
    hr = 1 if ns == 1 else 2
    t = body.table(hr + len(rows), ncol, widths, header_rows=hr)
    for j, h in enumerate(lead):
        cell_text(t.rows[0].cells[j], h)
    cell_text(t.rows[0].cells[nl], unit_head)
    if ns > 1:
        for k, s in enumerate(samples):
            cell_text(t.rows[1].cells[nl + k], s)
    cell_text(t.rows[0].cells[ncol - 1], limit_head)
    for i, row in enumerate(rows, start=hr):
        *before, res, lim = row
        c = t.rows[i].cells
        for j, x in enumerate(before):
            cell_text(c[j], x, align="left" if j == nl - 2 and nl > 2 or (nl == 2 and j == 0) else "center")
        for k, r in enumerate(res):
            op, val = _num(r or "")
            _, lv = _num((lim or "").lstrip("≥> "))
            hi = r is not None and val is not None and lv is not None and op != "<" and val >= lv
            cell_text(c[nl + k], v(r), bold=hi)
        cell_text(c[ncol - 1], lim)
        if lim in HL_TEXTS:                      # ngưỡng chưa chắc -> tô vàng (người dùng 17/11/2026)
            highlight_cell(c[ncol - 1])
    if ns > 1:                                   # gộp tiêu đề
        t.rows[0].cells[nl].merge(t.rows[0].cells[nl + ns - 1])
        for j in list(range(nl)) + [ncol - 1]:
            m = t.rows[0].cells[j].merge(t.rows[1].cells[j])
            for p in m.paragraphs[1:]:
                if not p.text.strip():
                    p._p.getparent().remove(p._p)
    for j in range(1, nl):                       # gộp dọc ô trùng chữ, trừ ô chưa có (………)
        r = hr
        while r < len(t.rows):
            txt = t.rows[r].cells[j].text.strip()
            e = r
            while e + 1 < len(t.rows) and t.rows[e + 1].cells[j].text.strip() == txt and txt != P:
                e += 1
            if e > r:
                merge_same_vertical(t, j, r, e)
            r = e + 1
    trs = t._tbl.findall(qn("w:tr"))
    _keep_rows(t._tbl, trs[:6])
    _keep_rows(t._tbl, trs[-2:])
    return t


HL_TEXTS = {"≥ 590"}   # xyanua: hồ sơ 107520 ghi ≥ 590, mẫu 112119 ghi ≥ 15 — chờ người dùng xác nhận


def highlight_cell(cell):
    for r in cell._tc.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = r.makeelement(qn("w:rPr"), {})
            r.insert(0, rpr)
        if rpr.find(qn("w:highlight")) is None:
            rpr.append(rpr.makeelement(qn("w:highlight"), {qn("w:val"): "yellow"}))


def is_hi(res, lim):
    out = []
    for r in res:
        op, val = _num(r or "")
        _, lv = _num((lim or "").lstrip("≥> "))
        out.append(None if r is None or lv is None else (op != "<" and val is not None and val >= lv))
    return out


def eval_block(rows):
    """-> (cao hơn, thấp hơn, không đánh giá) dạng chữ; chưa có KQTN -> ………"""
    base = lambda n: (lambda s: s[:1].lower() + s[1:])(n.split(" ^(")[0])
    noeval = [base(n) for n, _m, _r, lim in rows if lim == "-"]
    ev = [(base(n), is_hi(r, lim)) for n, _m, r, lim in rows if lim != "-"]
    if any(x is None for _n, h in ev for x in h):
        return P, P, ", ".join(noeval)
    hi = [n for n, h in ev if any(h)]
    return (", ".join(hi) or "-", "Tất cả các chỉ tiêu" if not hi else "các chỉ tiêu còn lại"), None, ", ".join(noeval)


def highlight_placeholders(doc):
    parts = [doc.element.body] + [s.header._element for s in doc.sections] + [s.footer._element for s in doc.sections]
    for root in parts:
        for r in root.iter(qn("w:r")):
            t = "".join(x.text or "" for x in r.iter(qn("w:t")))
            if P in t:
                rpr = r.find(qn("w:rPr"))
                if rpr is None:
                    rpr = r.makeelement(qn("w:rPr"), {})
                    r.insert(0, rpr)
                if rpr.find(qn("w:highlight")) is None:
                    h = rpr.makeelement(qn("w:highlight"), {qn("w:val"): "yellow"})
                    rpr.append(h)


def assessment_lines(b, intro, rows, tab=7.5, last=False):
    b.para(intro, keep_next=True)
    res = eval_block(rows)
    if res[0] == P:
        hi, lo = P, P
    else:
        hi, lo = res[0]
    noeval = res[2]
    lines = [("cao hơn mức quy định", hi), ("thấp hơn mức quy định", lo)]
    if noeval:
        lines.append(("không đánh giá", noeval))
    for k, (key, val) in enumerate(lines):
        q = b.para(parts=[("+ Chỉ tiêu ", False, False), (key, True, True), ("\t: " + val, False, False)],
                   indent_cm=0.75, tab_cm=tab, keep_next=not (last and k == len(lines) - 1))
        _keep(q._p, lines=True)


def ph_sentence(vals):
    if any(x is None for x in vals):
        return f"Độ pH {P} ngưỡng nguy hại tương ứng nêu tại QCVN 07:2025/BNNMT."
    ok = all(2.0 < _num(x)[1] < 12.5 for x in vals)
    return "Độ pH " + ("không vượt" if ok else "vượt") + " ngưỡng nguy hại tương ứng nêu tại QCVN 07:2025/BNNMT."


def pick(rows, k):
    """Kết quả của riêng mẫu k: (tên, phương pháp, [kết quả], ngưỡng)."""
    return [(n, m, [r[k]], lim) for n, m, r, lim in rows]


# ---------------------------------------------------------------------------
def build_tbkq(d, k):
    """Chứng thư /k+1: mẫu SAMPLES[k], ngày DAYS[k] (người dùng chốt 17/11/2026: tách 3 chứng thư)."""
    s, (a, e), seal = SAMPLES[k], TIMES[k], SEALS[k]
    src = CleanroomHandler(d).build_tbkq(is_en=False)
    doc = Document(src)
    fix_cover(doc, False, header_colon=False)
    text_proto, note_proto, head_proto = appendix_protos(doc)
    b = Body(doc, text_proto)
    head = make_head(b, head_proto)

    head("Nội dung thực hiện")
    b.para(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm theo yêu cầu của {CLIENT.name_vn}, cụ thể như sau:", keep_next=True)
    t = b.table(2, 3, [26, 14, 60])
    for j, h in enumerate(["Vị trí lấy mẫu", "Ký hiệu mẫu", "Chỉ tiêu thử nghiệm"]):
        cell_text(t.rows[0].cells[j], h)
    cell_text(t.rows[1].cells[0], VI_TRI)
    cell_text(t.rows[1].cells[1], s)
    cell_text(t.rows[1].cells[2], CHI_TIEU, align="left")
    _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:1])

    b.para("- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong:", keep_next=True, space_before=6)
    t = b.table(2, 4, [38, 16, 24, 22])
    for j, h in enumerate(["Người chứng kiến", "Ký hiệu mẫu", "Thời điểm lấy mẫu", "Số niêm phong"]):
        cell_text(t.rows[0].cells[j], h)
    for j, x in enumerate([WITNESS, s, f"{a} – {e}", seal]):
        cell_text(t.rows[1].cells[j], x)
    _keep_rows(t._tbl, t._tbl.findall(qn("w:tr")))
    b.photo(SEAL_PHOTOS[k])
    b.para("- Hình ảnh vị trí lấy mẫu:", keep_next=True, space_before=6)
    b.photo(LOC_PHOTOS[k])

    head("Kết quả giám định")
    vim = ("●", "Các thông số được chứng nhận VIMCERTS 078.")
    sub_head(b, "9.2.1. Tính chất nguy hại")
    sample_table(b, ["Tính chất nguy hại", "Phương pháp"],
                 [("Tính kiềm/ tính axit (pH)", "US EPA 9045D", [PH[0][k]], PH[1])],
                 "Kết quả", "Ngưỡng CTNH ^(1)", [32, 22, 16, 30], samples=[s])
    nt = note_table(b, note_proto, [vim, ("(1)", "Ngưỡng chất thải nguy hại tại QCVN 07:2025/BNNMT, bảng 1.")])
    _keep_rows(nt, nt.findall(qn("w:tr"))[:-1])

    def comp(title, rows, unit, mark, note, cust_mark):
        sub_head(b, title)
        sample_table(b, ["STT", "Thành phần nguy hại", "Phương pháp"],
                     [(str(i), n, m.replace(" & ", " &\n"), r, lim) for i, (n, m, r, lim) in enumerate(rows, 1)],
                     f"Kết quả, {unit}", f"Ngưỡng CTNH ^({mark}), {unit}", [8, 28, 30, 16, 18], samples=[s])
        nt = note_table(b, note_proto, [vim, (f"({mark})", note), (f"({cust_mark})", CUSTOMER_NOTE)])
        _keep_rows(nt, nt.findall(qn("w:tr"))[:-1])

    leach, absr = pick(LEACH, k), pick(ABS, k)
    comp("9.2.2. Nồng độ ngâm chiết của các thành phần nguy hại", leach, "mg/L", 2,
         "Ngưỡng chất thải nguy hại tính theo nồng độ ngâm chiết tại QCVN 07:2025/BNNMT, bảng 2.", 3)
    comp("9.2.3. Hàm lượng tuyệt đối của các thành phần nguy hại", absr, "ppm", 4,
         "Ngưỡng chất thải nguy hại tính theo hàm lượng tuyệt đối tại QCVN 07:2025/BNNMT, bảng 3.", 5)

    b.para("Đánh giá:", bold=True, keep_next=True, space_before=6)
    b.para(ph_sentence([PH[0][k]]), keep_next=True)
    assessment_lines(b, "So sánh kết quả thử nghiệm nồng độ ngâm chiết của các thành phần nguy hại với các ngưỡng nguy "
                        "hại tương ứng nêu tại QCVN 07:2025/BNNMT:", leach)
    assessment_lines(b, "So sánh kết quả thử nghiệm hàm lượng tuyệt đối của các thành phần nguy hại với các ngưỡng nguy "
                        "hại tương ứng nêu tại QCVN 07:2025/BNNMT:", absr, last=True)
    style_note_tables(doc)
    add_document_ending_line(doc)
    appendix_total_field(doc)
    highlight_placeholders(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - viet - {k + 1} - bun thai - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    print(f"[131873] TBKQ /{k + 1} -> {dst}")
    return dst


# ---------------------------------------------------------------------------
def build_bckq(d):
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - bun thai - {CODE}.docx")
    doc = Document(out)
    for row in doc.tables[0].rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs:
                    if P in r.text and "thiết bị" in row.cells[0].text.lower():
                        r.text = r.text.replace(P, "-")
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    h1 = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Nội dung thực hiện"))
    h2 = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Kết quả đo đạc"))
    gdv = next(el for el in kids if el.tag == qn("w:p") and _ptext(el).strip().startswith("Báo cáo của GĐV"))
    text_proto = __import__("copy").deepcopy(h1.getnext())
    i1, i2, ig = kids.index(h1), kids.index(h2), kids.index(gdv)
    for el in kids[i1 + 1:i2] + kids[i2 + 1:ig]:
        body_el.remove(el)
    b = Body(doc, text_proto)
    b.anchor = h2
    b.para(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm được chỉ định bởi đại diện {CLIENT.name_vn}, cụ thể như sau:",
           keep_next=True)
    t = b.table(4, 3, [26, 14, 60])
    for j, h in enumerate(["Vị trí thực hiện", "Kí hiệu mẫu", "Chỉ tiêu thử nghiệm"]):
        cell_text(t.rows[0].cells[j], h)
    for i, s in enumerate(SAMPLES, start=1):
        cell_text(t.rows[i].cells[0], VI_TRI)
        cell_text(t.rows[i].cells[1], s)
        cell_text(t.rows[i].cells[2], CHI_TIEU.replace("pH);", "pH), độ ẩm;"), align="left")
    merge_same_vertical(t, 0, 1)
    merge_same_vertical(t, 2, 1)
    b.para("- Mỗi mẫu được gộp từ 03 lần lấy mẫu trong ngày (01 lần tại máy ép bùn, 02 lần tại bao chứa), trộn đều "
           "thành 01 mẫu.", space_before=6)
    b.para("- Thông tin về quá trình lấy mẫu niêm phong:", keep_next=True, space_before=6)
    t = b.table(4, 4, [38, 16, 24, 22])
    for j, h in enumerate(["Người chứng kiến", "Ký hiệu mẫu", "Thời điểm lấy mẫu", "Số niêm phong"]):
        cell_text(t.rows[0].cells[j], h)
    for i, (s, dd, (a, e), seal) in enumerate(zip(SAMPLES, DAYS, TIMES, SEALS), start=1):
        for j, x in enumerate([WITNESS, s, f"{a} – {e}\n{dd}", seal]):
            cell_text(t.rows[i].cells[j], x)
    merge_same_vertical(t, 0, 1)
    photo_row(b, SEAL_PHOTOS)
    b.para("- Hình ảnh vị trí lấy mẫu:", keep_next=True, space_before=6)
    photo_row(b, LOC_PHOTOS)

    b.anchor = gdv
    _keep(h2)
    b.para(f"- KQTN tính chất nguy hại, nồng độ ngâm chiết và hàm lượng tuyệt đối của mẫu chất thải B1, B2, B3: xem tại "
           f"phiếu KQTN số {v(KQTN_SO)} ngày {v(KQTN_NGAY)};")
    b.para(f"- Tính toán hàm lượng tuyệt đối của chỉ tiêu phân tích theo mẫu rắn khô: xem tại phiếu Tính toán hàm "
           f"lượng tuyệt đối theo mẫu rắn khô số {CODE} ngày {v(KQTN_NGAY)}.")
    b.para("Đánh giá:", bold=True, keep_next=True, space_before=6)
    b.para(ph_sentence(PH[0]), keep_next=True)
    assessment_lines(b, "So sánh kết quả thử nghiệm các thành phần nguy hại theo nồng độ ngâm chiết của chất thải với "
                        "các ngưỡng nguy hại tương ứng nêu tại QCVN 07:2025/BNNMT:", LEACH)
    assessment_lines(b, "So sánh kết quả thử nghiệm các thành phần nguy hại theo hàm lượng tuyệt đối của chất thải với "
                        "các ngưỡng nguy hại tương ứng nêu tại QCVN 07:2025/BNNMT:", ABS, last=True)
    b.para("Gửi mẫu thử nghiệm:", bold=True, keep_next=True, space_before=6)
    for line in (f"Phiếu KQTN số {v(KQTN_SO)} ngày {v(KQTN_NGAY)}", f"Ngày gửi mẫu\t: {v(SEND_DATE)}",
                 f"Ngày yêu cầu cấp KQTN\t: {v(REQ_DATE)}", f"Ngày nhận KQTN\t: {v(KQTN_NGAY)}"):
        b.para(line, keep_next=True, tab_cm=5.5)
    b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre_han_text(KQTN_NGAY, SEND_DATE, REQ_DATE, P)}")
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm")):
            p.paragraph_format.keep_with_next = True
    highlight_placeholders(doc)
    doc.save(out)
    print(f"[131873] BCKQ -> {out}")
    return out


# ---------------------------------------------------------------------------
def build_excel(d):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from Core.Tools.excel_borders import thin_outside_borders
    wb = Workbook()
    ws = wb.active
    ws.title = "Tinh toan"
    f = Font(name="Times New Roman", size=12)
    fb = Font(name="Times New Roman", size=12, bold=True)
    yellow = PatternFill("solid", fgColor="FFFF00")
    ws["A1"], ws["F1"] = f"Số: {CODE}", f"Ngày {v(KQTN_NGAY)}"
    ws["A3"] = "TÍNH TOÁN QUY ĐỔI HÀM LƯỢNG TUYỆT ĐỐI THEO MẪU RẮN KHÔ"
    ws.merge_cells("A3:F3")
    for j, h in enumerate(["Kí hiệu mẫu", "Chỉ tiêu", "Kết quả, mg/kg", "Độ ẩm, %", "Hàm lượng chất rắn khô, %",
                           "Kết quả quy đổi về khối lượng khô, mg/kg"], start=1):
        ws.cell(5, j, h)
    items = [n.split(" ^(")[0] for n, *_ in ABS]
    r = 6
    thin = Side(style="thin")
    for k, s in enumerate(SAMPLES):
        r0 = r
        for i, name in enumerate(items):
            ws.cell(r, 2, name)
            res = ABS[i][2][k]
            if res is None:
                ws.cell(r, 3, P).fill = yellow
            elif res.strip().startswith("<"):
                ws.cell(r, 3, res)
            else:
                ws.cell(r, 3, float(res.replace(",", ".")))
            if i == 0:
                m = MOISTURE[k]
                c = ws.cell(r, 4, float(m.replace(",", ".")) if m else None)
                if not m:
                    c.fill = yellow
            else:
                ws.cell(r, 4, f'=IF(ISNUMBER($D${r0}),$D${r0},"")')
            ws.cell(r, 5, f'=IF(ISNUMBER(D{r}),100-D{r},"")')
            ws.cell(r, 6, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r})),ROUND(C{r}*100/E{r},2),C{r})')
            for col in (4, 5):
                ws.cell(r, col).number_format = "0.0"
            r += 1
        ws.cell(r0, 1, s)
        ws.merge_cells(start_row=r0, start_column=1, end_row=r - 1, end_column=1)
    for row in ws.iter_rows(min_row=5, max_row=r - 1, max_col=6):
        for c in row:
            c.font = fb if c.row == 5 else f
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws["A1"].font = ws["F1"].font = f
    ws["A3"].font = Font(name="Times New Roman", size=13, bold=True)
    ws["A3"].alignment = Alignment(horizontal="center")
    for col, w in zip("ABCDEF", (14, 18, 14, 11, 16, 22)):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[5].height = 48
    thin_outside_borders(wb)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - TinhToanChatThai.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    print(f"[131873] Excel -> {out}")
    return out


def build_edoc():
    import fitz
    from Core.Tools.scan_splitter import finish_edoc
    out_dir = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(out_dir, exist_ok=True)
    for src, name in (("0.  BBGDHT .pdf", "BBLM"), ("0.  BAO GIA  .pdf", "Khac")):
        doc = fitz.open()
        doc.insert_pdf(fitz.open(os.path.join(INPUT_DIR, src)))
        finish_edoc(doc, INPUT_DIR, name)
        dst = os.path.join(out_dir, f"{CODE6} - {name}.pdf")
        doc.save(dst)
        print(f"[131873] Edoc -> {dst} ({doc.page_count} trang)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    args = ap.parse_args()
    d = make_dossier()
    d.ensure_output_directories()
    build_edoc()
    certs = make_cert_dossiers()
    docs = [build_bckq(d)] + [build_tbkq(c, k) for k, c in enumerate(certs)]
    xls = [build_excel(d)]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(d).export_files(docs[:1])
        for c, p in zip(certs, docs[1:]):
            PDFExporter(c).export_files([p])
        PDFExporter(d).export_excel_files(xls)
    print("\n===== QA =====")
    for p in docs:
        res = check_single_docx(p, code=d.code)
        print(f"  {res.get('status')}\t{os.path.relpath(p, WORKSPACE)}")
        for i in res.get("issues", []):
            print(f"      - {i}")


if __name__ == "__main__":
    main()
