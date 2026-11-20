# -*- coding: utf-8 -*-
r"""
Điền TBKQ nước vào FORM GỐC của người dùng (một chứng thư) — 137037, 137102 (người dùng 18/11/2026: "kiểm tra lại luôn
3 hồ sơ 137037, 135226, 137102 … nếu chưa sửa lại"; quy tắc AGENTS mục 8C nước uống điểm 0, mẫu `fill_131229.py`).

Giữ nguyên thứ tự, tên, đơn vị, mức của form; chỉ điền / thay ô kết quả theo KQTN, ngày, thiết bị, Đánh giá. Phương pháp
form khác hẳn KQTN -> sửa chữ theo KQTN, giữ khung ô (người dùng 18/11/2026: "sửa phương pháp nhưng giữ form").

137037 (form người dùng đã điền sẵn, còn sót số của hồ sơ cũ 104561): bảng người chứng kiến (ký hiệu 104561-NSH, 09:40 ngày
21/09/2026 -> 137037-NSH, 09:35 ngày 09/11/2026 theo BBLM), permanganat "< 1,5 ⁽²⁾" -> "< 1" (KQTN KPH, LOD 1); địa điểm
theo BBLM ("Phòng thí nghiệm của …", người dùng chốt 18/11/2026); quy trình N2.TC.N.01 của form tô vàng (BBLM ghi
N2.TC.N.02 — người dùng 18/11/2026: tô vàng).
137102 (form để ………): thiết bị lạc dòng như 135226/6 -> điền ô (pH Mettler, clo dư Hach DR900 theo cách viết form
135226); form không có ghi chú ⁽²⁾ mà KQTN có kết quả LOQ (nitơ tổng, COD) -> thêm dòng ghi chú (2); form không có đoạn
Đánh giá -> thêm như các form nước khác (AGENTS 8C: luôn có Đánh giá khi có mức).

Chạy:  python fill_form_nuoc.py 137037|137102 [--pdf]
"""
import argparse
import copy
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
from fill_135226 import highlight                       # noqa: E402

PH = "………"
N = lambda s: re.sub(r"[\s\-\.–&⁻]|\(.*?\)|method", "", s.lower().translate(str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")))   # noqa: E731  bỏ cách viết, ngoặc, "Method", "&"
CFG = {
    "137037": dict(loai="nuoc sinh hoat", kq="KQ-*.pdf", word="Thông số",
                   keys=[("màu", "độ màu"), ("đục", "độ đục"), ("asen", "asen"), ("permanganat", "permanganat"),
                         ("amoni", "amoni"), ("coliform", "coliforms"), ("escherichia", "escherichia")],
                   site_prefix="Phòng thí nghiệm của ", hl_proc=True,
                   witness=("137037-NSH", "09:35 ngày 09/11/2026")),
    "137102": dict(loai="nuoc thai", kq="KQTN/*.01.pdf", word="Thông số",
                   keys=[("lơ lửng", "lơ lửng"), ("sinh ho", "sinh hóa"), ("oxy hóa học", "oxy hóa học"),
                         ("dầu mỡ", "dầu mỡ"), ("amoni", "amoni"), ("đồng", "đồng"), ("tổng nitơ", "nitơ"),
                         ("phốt pho", "phospho"), ("chì", "chì"), ("kẽm", "kẽm"), ("crom (iii)", "crom iii"),
                         ("crom (vi)", "crom vi"), ("thủy ngân", "thủy ngân"), ("asen", "asen"), ("cadimi", "cadmi"),
                         ("coliform", "coliforms")],
                   equip=["Máy đo pH Seven2Go S2, Mettler Toledo – USA;", "Máy đo Chlorine DR900, Hach – USA."],
                   std="QCVN 40:2011/BTNMT (cột B, Kq = 1, Kf = 1)"),
}


def set_text(cell, value):
    """Thay toàn bộ chữ trong ô bằng value (giữ định dạng run đầu), bỏ tô vàng."""
    p = cell.find(qn("w:p"))
    for extra in cell.findall(qn("w:p"))[1:]:
        cell.remove(extra)
    runs = p.findall(qn("w:r"))
    base = runs[0] if runs else None
    for r in runs:
        p.remove(r)
    r = copy.deepcopy(base) if base is not None else OxmlElement("w:r")
    for ch in list(r):
        if ch.tag != qn("w:rPr"):
            r.remove(ch)
    rp = r.find(qn("w:rPr"))
    if rp is not None:
        for tag in ("w:highlight", "w:vertAlign", "w:b"):
            for e in rp.findall(qn(tag)):
                rp.remove(e)
    t = OxmlElement("w:t"); t.text = value
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    p.append(r)
    for h in list(cell.iter(qn("w:highlight"))):          # cả dấu đoạn
        h.getparent().remove(h)
    return r


def sup_run(cell, mark):
    p = cell.find(qn("w:p"))
    r = copy.deepcopy(p.findall(qn("w:r"))[0])
    for ch in list(r):
        if ch.tag != qn("w:rPr"):
            r.remove(ch)
    va = OxmlElement("w:vertAlign"); va.set(qn("w:val"), "superscript")
    r.find(qn("w:rPr")).append(va)
    t = OxmlElement("w:t"); t.text = f" {mark}"
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    p.append(r)


def fix_methods(rt_tbl, kq_methods, report):
    """Sửa cột Phương pháp (cột 4) theo KQTN, giữ form (người dùng 18/11/2026: "sửa phương pháp nhưng giữ form").
    kq_methods: phương pháp KQTN theo từng dòng dữ liệu (None = đo hiện trường, giữ của form). Ô gộp dọc: cả nhóm cùng
    phương pháp mới -> sửa chữ ô đầu; nhóm lẫn -> tách thành các nhóm liền nhau cùng phương pháp."""
    trs = rt_tbl.findall(qn("w:tr"))[1:]
    cells = [tr.findall(qn("w:tc"))[3] for tr in trs]

    def vm(tc):
        pr = tc.find(qn("w:tcPr"))
        e = pr.find(qn("w:vMerge")) if pr is not None else None
        return None if e is None else (e.get(qn("w:val")) or "continue")

    groups, cur = [], []
    for i, tc in enumerate(cells):
        if vm(tc) == "continue" and cur:
            cur.append(i)
        else:
            if cur:
                groups.append(cur)
            cur = [i]
    groups.append(cur)
    for g in groups:
        form_m = txt(cells[g[0]]).strip()
        want = [kq_methods[i] or form_m for i in g]
        same = lambda a, b: N(a) == N(b) or N(a) in N(b) or N(b) in N(a)   # noqa: E731
        if all(same(w, form_m) for w in want):
            continue
        runs, start = [], 0                               # nhóm liền nhau cùng phương pháp
        for j in range(1, len(g) + 1):
            if j == len(g) or not same(want[j], want[start]):
                runs.append((start, j))
                start = j
        for a, b in runs:
            head = cells[g[a]]
            pr = head.find(qn("w:tcPr"))
            for e in pr.findall(qn("w:vMerge")):
                pr.remove(e)
            if b - a > 1:
                e = OxmlElement("w:vMerge"); e.set(qn("w:val"), "restart"); pr.append(e)
            if a > 0 and not head.findall(qn("w:p")):
                head.append(copy.deepcopy(cells[g[0]].find(qn("w:p"))))
            if not head.findall(qn("w:r")) and cells[g[0]].findall(qn("w:r")):
                head.find(qn("w:p")).append(copy.deepcopy(cells[g[0]].findall(qn("w:r"))[0]))
            text = want[a] if not same(want[a], form_m) else form_m
            set_text(head, text)
            for i in range(a + 1, b):
                pr2 = cells[g[i]].find(qn("w:tcPr"))
                for e in pr2.findall(qn("w:vMerge")):
                    pr2.remove(e)
                e = OxmlElement("w:vMerge"); pr2.append(e)
            if not same(want[a], form_m):
                report.append(f"PHƯƠNG PHÁP dòng {g[a] + 1}–{g[b - 1] + 1}: '{form_m}' -> '{want[a]}'")


def clean(v):
    v = re.sub(r"⁽[^⁾]*⁾", "", v).strip()
    return re.sub(r"^<\s*", "< ", v)


def lim_ok(v, lim):
    nums = re.findall(r"[\d][\d\s]*(?:,\d+)?", lim)
    m = re.match(r"^<\s*([\d,]+)$", v)
    if m and len(nums) == 2:                       # "< 0,1" với khoảng (0,2 – 1,0): mọi giá trị đều dưới cận dưới -> không đạt
        return float(m.group(1).replace(",", ".")) > float(nums[0].replace(" ", "").replace(",", "."))
    if v.startswith("<") or not nums or not re.match(r"^[\d,]+$", v):
        return True
    x = float(v.replace(",", "."))
    vals = [float(n.replace(" ", "").replace(",", ".")) for n in nums]
    return vals[0] <= x <= vals[1] if len(vals) == 2 else x <= vals[0]


def main(h6, pdf):
    c = CFG[h6]
    IN = os.path.join(WS, "Input", h6)
    code = f"Q326D{h6}"
    k = read_kqtn(glob.glob(os.path.join(IN, c["kq"]))[0])
    date = k["ngay"]
    d = docx.Document(os.path.join(IN, f"3. TBKQ - {code}.docx"))
    for s in d.sections:
        for hp in (s.header, s.first_page_header):
            for p in hp.paragraphs:
                fill_runs(p._p, date)
    t0 = d.tables[0]
    for row in t0.rows:
        lab, val = row.cells[0].text.strip(), row.cells[-1]._tc
        if lab.startswith("Địa điểm") and c.get("site_prefix") and not txt(val).startswith(c["site_prefix"]):
            t = next(t for t in val.iter(qn("w:t")) if t.text)
            t.text = c["site_prefix"] + t.text[0].lower() + t.text[1:]
        if lab.startswith("Quy trình") and c.get("hl_proc"):
            highlight(val, [txt(val).strip()])
        if lab.startswith("Thiết bị") and c.get("equip") and not txt(val).strip():
            stray = row._tr.getnext()
            if stray is not None and txt(stray).strip() == PH:
                stray.getparent().remove(stray)
            p0 = val.find(qn("w:p"))
            for i, line in enumerate(c["equip"]):
                p = p0 if i == 0 else copy.deepcopy(p0)
                for r in p.findall(qn("w:r")):
                    p.remove(r)
                r = OxmlElement("w:r"); rp = OxmlElement("w:rPr")
                for tag in ("w:sz", "w:szCs"):
                    e = OxmlElement(tag); e.set(qn("w:val"), "26"); rp.append(e)
                r.append(rp); t = OxmlElement("w:t"); t.text = "- " + line; r.append(t); p.append(r)
                if i:
                    val.append(p)
        if lab.startswith("-") and "Giấy yêu cầu" in lab:
            dash_numbering(d, row.cells[0]._tc.find(qn("w:p")))
    if c.get("witness"):
        wt = next(t for t in d.tables if t.rows[0].cells[0].text.strip().startswith("Người chứng kiến"))
        tcs = wt._tbl.findall(qn("w:tr"))[1].findall(qn("w:tc"))
        set_text(tcs[1], c["witness"][0])
        set_text(tcs[2], c["witness"][1])
    rt = next(t for t in d.tables if t.rows[0].cells[0].text.strip() in ("Stt", "STT"))
    meth, any_loq, fails, report = "", False, [], []
    kq_methods = []
    for tr in rt._tbl.findall(qn("w:tr"))[1:]:
        tcs = tr.findall(qn("w:tc"))
        name = txt(tcs[1]).strip()
        if txt(tcs[3]).strip():
            meth, meth_tc = txt(tcs[3]).strip(), tcs[3]
        key = next((kk for f, kk in c["keys"] if f in name.lower()), None)
        r = next((r for r in k["rows"] if key and key in r.name.lower()), None)
        kq_methods.append(r.method if r else None)
        if r is None:
            continue                                       # đo / cảm quan hiện trường: giữ số của form
        v, loq = r.value()
        v = clean(v)
        old = txt(tcs[4]).strip()
        set_text(tcs[4], v)
        if loq:
            any_loq = True
            sup_run(tcs[4], "(2)")
        if old != v:
            report.append(f"{name}: '{old}' -> '{v}'")
        if not lim_ok(v, txt(tcs[5]).strip()):
            fails.append(name)
    fix_methods(rt._tbl, kq_methods, report)
    nt = next(t for t in d.tables if t.rows[0].cells[0].text.strip().startswith("Ghi chú"))
    rows2 = [tr for tr in nt._tbl.findall(qn("w:tr")) if txt(tr).strip().startswith("(2)")]
    if not any_loq:
        for tr in rows2:
            nt._tbl.remove(tr)
    elif not rows2:
        r1 = next(tr for tr in nt._tbl.findall(qn("w:tr")) if txt(tr).strip().startswith("(1)"))
        r2 = copy.deepcopy(r1)
        tcs = r2.findall(qn("w:tc"))
        for t in tcs[0].iter(qn("w:t")):
            t.text = t.text.replace("(1)", "(2)")
        set_text(tcs[1], "Giới hạn định lượng của phương pháp thử.")
        r1.addnext(r2)
        report.append("thêm dòng ghi chú (2) Giới hạn định lượng")
    body_txt = " ".join(p.text for p in d.paragraphs)
    if "Đánh giá" not in body_txt:
        last = nt._tbl
        paras = [("Đánh giá:", True),
                 (f"So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên với mức quy định tương ứng tại "
                  f"{c['std']}:", False),
                 (f"+ Thông số không đạt: {', '.join(fails) if fails else '-'}", False),
                 (f"+ Thông số đạt: {'các thông số còn lại' if fails else 'tất cả các thông số'}", False)]
        anchor = last
        for text, bold in paras:
            p = OxmlElement("w:p"); ppr = OxmlElement("w:pPr")
            sp = OxmlElement("w:spacing"); sp.set(qn("w:before"), "120" if bold else "40"); sp.set(qn("w:after"), "40")
            ppr.append(sp)
            if text.startswith("+"):
                ind = OxmlElement("w:ind"); ind.set(qn("w:left"), "426"); ppr.append(ind)
            elif not bold:
                jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "both"); ppr.append(jc)
            if bold:
                kn = OxmlElement("w:keepNext"); ppr.insert(0, kn)
            p.append(ppr)
            r = OxmlElement("w:r"); rp = OxmlElement("w:rPr")
            if bold:
                rp.append(OxmlElement("w:b"))
            for tag in ("w:sz", "w:szCs"):
                e = OxmlElement(tag); e.set(qn("w:val"), "26"); rp.append(e)
            r.append(rp); t = OxmlElement("w:t"); t.text = text
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); r.append(t); p.append(r)
            anchor.addnext(p)
            anchor = p
        report.append("thêm đoạn Đánh giá")
    out = os.path.join(IN, "TBKQ", f"TBKQ - viet - {c['loai']} - {code}.docx")
    d.save(out)
    print(f"[OK] {os.path.relpath(out, WS)}  không đạt: {fails or '-'}")
    for x in report:
        print("   -", x)
    if pdf:
        mod = __import__(f"build_{h6}")
        from nuoc_common import make_dossier
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(make_dossier(dict(mod.CONFIG, cert_date=date), None, None, False)).export_files([out])
    from Core.Tools.qa_check import check_single_docx
    r = check_single_docx(out, code=code)
    print(r.get("status"))
    for i in r.get("issues", []):
        print("    -", i)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("ho_so")
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    main(a.ho_so, a.pdf)
