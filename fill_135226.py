# -*- coding: utf-8 -*-
r"""
Q326D135226 — TBKQ /1 – /11 ĐIỀN VÀO FORM GỐC của người dùng (thay bản dựng bằng nuoc_common), cùng cách với 131229.

Người dùng (18/11/2026, 131229): "giữ lại form cũ … điền kết quả dựa theo form cũ chứ không phải lấy từ KQTN mà sửa lại
phần chỉ tiêu … Bạn chỉ điền thôi không cần sửa"; "Tương tự kiểm tra với hồ sơ IDICO 135226".

Form (chỉ đọc): `Input\135226\3. TBKQ - <n> - Q326D135226.docx` (= bản trên Z, 16/11/2026).
KQTN: `Q326D135226 - TBKQ\Edoc\KQ-Q326D135226-MTTN26020080.01 … .11.pdf` (17/11/2026).

Chỉ điền: ngày ban hành ở header (ngày KQTN), ô Kết quả "………" theo KQTN (KPH -> "< LOD", LOQ -> "< LOQ" + "(*)" chỉ
số trên như ghi chú của form; mẫu không có LOQ bỏ dòng ghi chú "(*)"). Giữ nguyên thứ tự, tên, đơn vị, phương pháp của form.
Áp lại các quyết định người dùng đã chốt sáng 18/11/2026 cho hồ sơ này: địa điểm theo BBLM ("Khu xử lý nước thải " +
form); Căn cứ "-" tô vàng; "NL" (N5) tô vàng; Ni N9 "Error" (KQTN) tô vàng. Dấu "-" mục 2 -> đánh số tự động
(như 131229).

Chạy:  python fill_135226.py [--pdf] [--only 1,3]
"""
import argparse
import glob
import os
import re
import sys

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

WS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, WS)
from Core.Tools.kqtn_reader import read_kqtn           # noqa: E402
from fill_131229 import fill_runs, dash_numbering, txt  # noqa: E402

IN = os.path.join(WS, "Input", "135226")
OUT = os.path.join(IN, "TBKQ")
CODE = "Q326D135226"
PH = "………"
# tên chỉ tiêu trên form -> khóa tìm trong tên KQTN (chữ thường)
KEYS = [("độ màu", "độ màu"), ("sinh ho", "sinh hóa"), ("oxy hóa học", "oxy hóa học"), ("lơ lửng", "lơ lửng"),
        ("asen", "asen"), ("thủy ngân", "thủy ngân"), ("chì", "chì"), ("cadimi", "cadmi"), ("xyanua", "xyanua"),
        ("crom (vi)", "crom vi"), ("crom (iii)", "crom iii"), ("đồng", "đồng"), ("kẽm", "kẽm"), ("niken", "niken"),
        ("mangan", "mangan"), ("sắt", "sắt"), ("phenol", "phenol"), ("dầu mỡ", "dầu mỡ"), ("sunfua", "sunfua"),
        ("florua", "florua"), ("clorua", "clorua"), ("amoni", "amoni"), ("nitơ", "nitơ"), ("photpho", "phospho"),
        ("coliform", "coliform"), ("phóng xạ α", "alpha"), ("phóng xạ β", "beta")]


def highlight(el, words):
    """Tô vàng đúng cụm chữ (tách run nếu cần)."""
    for r in list(el.iter(qn("w:r"))):
        t = r.find(qn("w:t"))
        if t is None or not t.text:
            continue
        for w in words:
            if w in t.text:
                pre, post = t.text.split(w, 1)
                import copy
                parts = [(pre, False), (w, True), (post, False)]
                parent, idx = r.getparent(), list(r.getparent()).index(r)
                for j, (s, hl) in enumerate(p for p in parts if p[0]):
                    r2 = copy.deepcopy(r)
                    r2.find(qn("w:t")).text = s
                    r2.find(qn("w:t")).set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    if hl:
                        rp = r2.find(qn("w:rPr"))
                        if rp is None:
                            rp = OxmlElement("w:rPr"); r2.insert(0, rp)
                        h = OxmlElement("w:highlight"); h.set(qn("w:val"), "yellow"); rp.append(h)
                    parent.insert(idx + j, r2)
                parent.remove(r)
                break


def fill(n, k, date):
    d = docx.Document(os.path.join(IN, f"3. TBKQ - {n} - {CODE}.docx"))
    for s in d.sections:
        for hp in (s.header, s.first_page_header):
            for p in hp.paragraphs:
                fill_runs(p._p, date)
    for row in d.tables[0].rows:
        lab, val = row.cells[0].text.strip(), row.cells[-1]._tc
        if lab.startswith("Địa điểm") and not txt(val).startswith("Khu xử lý"):
            t = next(t for t in val.iter(qn("w:t")) if t.text)
            t.text = "Khu xử lý nước thải " + t.text            # theo BBLM (người dùng 18/11/2026)
        if lab.startswith("Thiết bị") and not txt(val).strip():   # /6: ô giá trị trống, "………" lạc xuống 1 dòng riêng
            p = val.find(qn("w:p"))
            r = OxmlElement("w:r"); rp = OxmlElement("w:rPr")
            for tag in ("w:sz", "w:szCs"):
                e = OxmlElement(tag); e.set(qn("w:val"), "26"); rp.append(e)
            r.append(rp); t = OxmlElement("w:t"); t.text = "Máy đo pH Seven2Go S2, Mettler Toledo – USA"; r.append(t)
            p.append(r)                                    # như form /3 – /11
            stray = row._tr.getnext()
            if stray is not None and txt(stray).strip() == PH:
                stray.getparent().remove(stray)
        if lab.startswith("Căn cứ"):
            highlight(val, ["-"])
        if lab.startswith("-") and "Giấy yêu cầu" in lab:
            dash_numbering(d, row.cells[0]._tc.find(qn("w:p")))
    for t in d.tables[1:2]:
        highlight(t._tbl, ["NL"])
    rt = next(t for t in d.tables if t.rows[0].cells[0].text.strip() in ("Stt", "STT"))
    missing, any_loq, used = [], False, set()
    kq_methods, report = [], []
    for tr in rt._tbl.findall(qn("w:tr"))[1:]:
        tcs = tr.findall(qn("w:tc"))
        name, cell = txt(tcs[1]).strip(), tcs[-1]
        if PH not in txt(cell):
            kq_methods.append(None)
            continue                                       # đo hiện trường: form đã có số
        key = next((kk for f, kk in KEYS if f in name.lower()), None)
        r = next((r for r in k["rows"] if key and key in r.name.lower()), None)
        kq_methods.append(r.method if r else None)
        if r is None:
            missing.append(name)
            continue
        used.add(r.name)
        v, loq = r.value()
        fill_runs(cell, v)
        for h in list(cell.iter(qn("w:highlight"))):
            h.getparent().remove(h)
        if v == "Error":
            highlight(cell, ["Error"])
        if loq:
            any_loq = True
            p = cell.find(qn("w:p"))
            r2 = OxmlElement("w:r")
            rp = OxmlElement("w:rPr")
            for tag, val in (("w:sz", "26"), ("w:szCs", "26"), ("w:vertAlign", "superscript")):
                e = OxmlElement(tag); e.set(qn("w:val"), val); rp.append(e)
            r2.append(rp)
            t2 = OxmlElement("w:t"); t2.text = " (*)"
            t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r2.append(t2)
            p.append(r2)
    from fill_form_nuoc import fix_methods
    fix_methods(rt._tbl, kq_methods, report)
    for x in report:
        print("   -", x)
    extra = [r.name for r in k["rows"] if r.name not in used]
    if not any_loq:
        nt = next(t for t in d.tables if t.rows[0].cells[0].text.strip().startswith("Ghi chú"))
        for tr in nt._tbl.findall(qn("w:tr")):
            if txt(tr).strip().startswith("(*)"):
                nt._tbl.remove(tr)
    out = os.path.join(OUT, f"TBKQ - viet - {n} - nuoc thai - {CODE}.docx")
    d.save(out)
    print(f"[OK] /{n}  LOQ: {any_loq}  form thiếu KQTN: {missing or '-'}  KQTN không có trong form: {extra or '-'}")
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    only = {int(x) for x in a.only.split(",") if x.strip()}
    kq = {}
    for f in glob.glob(os.path.join(IN, f"{CODE} - TBKQ", "Edoc", "KQ-*.pdf")):
        k = read_kqtn(f)
        m = re.search(r"- N(\d+)$", k["ten_mau"])
        if m:
            kq[int(m.group(1))] = k
    date = max(k["ngay"] for k in kq.values())
    docs = [fill(n, kq[n], date) for n in range(1, 12) if not only or n in only]
    if a.pdf:
        from build_135226 import CONFIG
        from nuoc_common import make_dossier
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(make_dossier(dict(CONFIG, cert_date=date), None, None, False)).export_files(docs)
    from Core.Tools.qa_check import check_single_docx
    for p in docs:
        r = check_single_docx(p, code=CODE)
        print(r.get("status"), os.path.basename(p))
        for i in r.get("issues", []):
            print("    -", i)
