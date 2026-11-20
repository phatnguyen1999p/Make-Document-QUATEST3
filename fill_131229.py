# -*- coding: utf-8 -*-
r"""
Q326D131229 — TBKQ VN/EN /1 – /6 ĐIỀN VÀO FORM GỐC của người dùng (thay bản dựng bằng nuoc_common).

Người dùng (18/11/2026): "bạn đã sửa đổi thứ tự của bản gốc và thay đổi chỉ tiêu, quá dài và không cần thiết, tôi cần
bạn giữ lại form cũ (nằm ở bên ngoài thư mục TBKQ mà bạn đã làm), điền kết quả dựa theo form cũ chứ không phải lấy từ
KQTN mà sửa lại phần chỉ tiêu, form gốc là form theo quy chuẩn rồi. Bạn chỉ điền thôi không cần sửa nhé" + "dấu gạch đầu
dòng ở mục 2 ... không được ngay với số 7".

Form (chỉ đọc): `Input\131229\3. TBKQ - <n> - vn|en - Q326D131229.docx` (= bản trên Z, 09/11/2026); riêng /1 vn lấy bản
người dùng sửa trên Z 18/11/2026 15:29 (`Input\131229\_Z_18-11\`, đã điền ngày ban hành, nén ảnh).
KQTN: `Q326D131229 - TBKQ\77708 - edoc\KQ-…-MTTN26019574.01 … .06.pdf` (16/11/2026).

Chỉ điền: ngày ban hành ở header (ngày KQTN 16/11/2026), "Thiết bị sử dụng" = "-" (BBLM: không đo hiện trường), ô
Kết quả (KPH -> "< LOD", LOQ -> "< LOQ" + ⁽²⁾, vi sinh "<1⁽ᶜ⁾" -> "< 1"; EN dấu chấm), Đánh giá. Giữ nguyên thứ tự, tên,
đơn vị, phương pháp, mức của form. Dấu ⁽²⁾ của form ở ô nitrat chuyển theo KQTN (chỉ tiêu nào là LOQ thì gắn); mẫu không
có kết quả LOQ (LSP2) bỏ dòng ghi chú ⁽²⁾. Dòng "- Giấy yêu cầu…" (mục 2): dấu "-" gõ tay -> đánh số tự động kiểu "-"
cùng thụt lề với số 1. … 8. để thẳng hàng.

Chạy:  python fill_131229.py [--pdf] [--only 1,3]
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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from Core.Tools.kqtn_reader import read_kqtn           # noqa: E402

WS = os.path.dirname(os.path.abspath(__file__))
IN = os.path.join(WS, "Input", "131229")
OUT = os.path.join(IN, "TBKQ")
CODE = "Q326D131229"
PH = "………"
# thứ tự chỉ tiêu của form (giống VN / EN) -> khóa tìm trong tên KQTN
ORDER = ["stibi", "asen", "bari", "bor (", "bromat", "cadmi", "tổng clo", "clorate", "clorit", "crom", "đồng",
         "cyanid", "fluorid", "chì", "mangan", "thủy ngân", "molybden", "niken", "nitrat", "nitrit", "selen",
         "escherichia", "coliforms", "enterococci", "pseudomonas", "bào tử"]


def form_path(n, lang):
    z = os.path.join(IN, "_Z_18-11", f"3. TBKQ - {n} - {lang} - {CODE}.docx")
    return z if os.path.exists(z) else os.path.join(IN, f"3. TBKQ - {n} - {lang} - {CODE}.docx")


def kqtn_by_sample():
    out = {}
    for f in glob.glob(os.path.join(IN, f"{CODE} - TBKQ", "77708 - edoc", "*.pdf")):
        k = read_kqtn(f)
        m = re.search(r"LSP\s*(\d)", k["ten_mau"])
        out[int(m.group(1))] = k
    return out


def results(k):
    rows = []
    for key in ORDER:
        r = next(r for r in k["rows"] if key in r.name.lower())
        if r.result.replace(" ", "").startswith("<1⁽"):            # vi sinh "<1⁽ᶜ⁾" = không phát hiện
            v, loq = "< 1", False
        else:
            v, loq = r.value()
        rows.append((r.name, v, loq))
    return rows


def passes(v, lim):
    if not re.match(r"^[\d,]+$", lim) or v.startswith("<"):
        return True
    return float(v.replace(",", ".")) <= float(lim.replace(",", "."))


def txt(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def unhl(run):
    rp = run.find(qn("w:rPr"))
    if rp is not None:
        for h in rp.findall(qn("w:highlight")):
            rp.remove(h)


def fill_runs(el, value):
    """Thay '………' (run tô vàng) trong el bằng value, bỏ tô vàng."""
    done = False
    for r in el.iter(qn("w:r")):
        for t in r.findall(qn("w:t")):
            if t.text and PH in t.text:
                t.text = t.text.replace(PH, value)
                unhl(r)
                done = True
    return done


def dash_numbering(doc, para):
    """Dấu '-' gõ tay -> đánh số tự động dạng '-' cùng thụt lề với danh sách 1. … 8. (numId 10)."""
    num = doc.part.numbering_part.element
    ppr = para.find(qn("w:pPr"))
    src = None
    for p in doc.element.body.iter(qn("w:p")):
        np_ = p.find(qn("w:pPr"))
        if np_ is not None and np_.find(qn("w:numPr")) is not None:
            src = np_.find(qn("w:numPr")).find(qn("w:numId")).get(qn("w:val"))
            break
    n = next(x for x in num.findall(qn("w:num")) if x.get(qn("w:numId")) == src)
    aid = n.find(qn("w:abstractNumId")).get(qn("w:val"))
    ab = next(a for a in num.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == aid)
    new_aid = str(max(int(a.get(qn("w:abstractNumId"))) for a in num.findall(qn("w:abstractNum"))) + 1)
    new_nid = str(max(int(x.get(qn("w:numId"))) for x in num.findall(qn("w:num"))) + 1)
    ab2 = copy.deepcopy(ab)
    ab2.set(qn("w:abstractNumId"), new_aid)
    for ch in list(ab2):
        if ch.tag in (qn("w:nsid"), qn("w:tmpl"), qn("w:numStyleLink"), qn("w:styleLink")):
            ab2.remove(ch)
    for lvl in ab2.findall(qn("w:lvl"))[1:]:
        ab2.remove(lvl)
    lvl = ab2.find(qn("w:lvl"))
    lvl.find(qn("w:numFmt")).set(qn("w:val"), "bullet")
    lvl.find(qn("w:lvlText")).set(qn("w:val"), "-")
    rp = lvl.find(qn("w:rPr"))
    f = OxmlElement("w:rFonts")
    for a in ("ascii", "hAnsi", "cs"):
        f.set(qn(f"w:{a}"), "Times New Roman")
    rp.insert(0, f)
    num.insert(list(num).index(num.findall(qn("w:abstractNum"))[-1]) + 1, ab2)
    n2 = OxmlElement("w:num")
    n2.set(qn("w:numId"), new_nid)
    a2 = OxmlElement("w:abstractNumId")
    a2.set(qn("w:val"), new_aid)
    n2.append(a2)
    num.append(n2)
    npr = OxmlElement("w:numPr")
    il = OxmlElement("w:ilvl"); il.set(qn("w:val"), "0")
    ni = OxmlElement("w:numId"); ni.set(qn("w:val"), new_nid)
    npr.append(il); npr.append(ni)
    ppr.insert(0, npr)
    # bỏ dấu "-" gõ tay đầu đoạn
    for t in para.iter(qn("w:t")):
        if t.text:
            t.text = re.sub(r"^\s*-\s*", "", t.text)
            break
    for r in para.findall(qn("w:r")):
        if r.find(qn("w:tab")) is not None and not txt(r):
            r.getparent().remove(r)
            break


def fill(n, lang, k, date):
    en = lang == "en"
    d = docx.Document(form_path(n, lang))
    for s in d.sections:
        for hp in (s.header, s.first_page_header):
            for p in hp.paragraphs:
                fill_runs(p._p, date)
    t0 = d.tables[0]
    for row in t0.rows:
        lab = row.cells[0].text.strip()
        if lab.startswith(("Thiết bị", "Equipment")):
            fill_runs(row.cells[-1]._tc, "-")
        if lab.startswith("-") and ("Giấy yêu cầu" in lab or "application" in lab):
            dash_numbering(d, row.cells[0]._tc.find(qn("w:p")))
    res = results(k)
    rt = next(t for t in d.tables if t.rows[0].cells[0].text.strip() in ("Stt", "No."))
    trs = rt._tbl.findall(qn("w:tr"))[1:]
    assert len(trs) == len(res), (len(trs), len(res))
    fails = []
    any_loq = False
    for tr, (name, v, loq) in zip(trs, res):
        tcs = tr.findall(qn("w:tc"))
        cell, lim = tcs[4], txt(tcs[5]).strip()
        vv = v.replace(",", ".") if en else v
        fill_runs(cell, vv)
        sup = [r for r in cell.iter(qn("w:r")) if "(2)" in txt(r)]
        if loq:
            any_loq = True
            if not sup:
                p = cell.find(qn("w:p"))
                base = p.findall(qn("w:r"))[0]
                r2 = copy.deepcopy(base)
                for tt in r2.findall(qn("w:t")):
                    r2.remove(tt)
                rp = r2.find(qn("w:rPr"))
                va = OxmlElement("w:vertAlign"); va.set(qn("w:val"), "superscript")
                rp.append(va)
                t2 = OxmlElement("w:t"); t2.text = " (2)"
                t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                r2.append(t2)
                p.append(r2)
        else:
            for r in sup:
                r.getparent().remove(r)
        if not passes(v, lim.replace(".", ",") if en else lim):
            fails.append(txt(tcs[1]).strip())
            for r in cell.iter(qn("w:r")):
                rp = r.find(qn("w:rPr"))
                if rp is not None and rp.find(qn("w:b")) is None:
                    rp.insert(0, OxmlElement("w:b"))
    for h in list(rt._tbl.iter(qn("w:highlight"))):    # tô vàng sót ở dấu đoạn / run rỗng của ô form
        h.getparent().remove(h)
    if not any_loq:                                    # không có kết quả LOQ -> bỏ dòng ghi chú (2)
        nt = next(t for t in d.tables if t.rows[0].cells[0].text.strip().startswith(("Ghi chú", "Note")))
        for tr in nt._tbl.findall(qn("w:tr")):
            if txt(tr).strip().startswith("(2)"):
                nt._tbl.remove(tr)
    for p in d.paragraphs:
        tx = p.text
        if PH not in tx:
            continue
        if ("không đạt" in tx) or ("Non-compliant" in tx):
            fill_runs(p._p, ", ".join(fails) if fails else "-")
        elif ("đạt" in tx) or ("Compliant" in tx):
            if fails:
                ok = [txt(tr.findall(qn("w:tc"))[1]).strip() for tr in trs
                      if txt(tr.findall(qn("w:tc"))[1]).strip() not in fails]
                fill_runs(p._p, ", ".join(ok))
            else:
                fill_runs(p._p, "all parameters" if en else "tất cả các chỉ tiêu")
    out = os.path.join(OUT, f"TBKQ - {'en' if en else 'viet'} - {n} - nuoc uong - {CODE}.docx")
    d.save(out)
    left = [p.text for p in d.paragraphs if PH in p.text]
    print(f"[OK] {os.path.relpath(out, WS)}  không đạt: {fails or '-'}  LOQ: {any_loq}  còn '………': {left}")
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    only = {int(x) for x in a.only.split(",") if x.strip()}
    kq = kqtn_by_sample()
    date = max(k["ngay"] for k in kq.values())
    docs = []
    for n in range(1, 7):
        if only and n not in only:
            continue
        for lang in ("vn", "en"):
            docs.append(fill(n, lang, kq[n], date))
    if a.pdf:
        from build_131229 import CONFIG
        from nuoc_common import make_dossier
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(make_dossier(dict(CONFIG, cert_date=date), None, None, False)).export_files(docs)
    from Core.Tools.qa_check import check_single_docx
    for p in docs:
        r = check_single_docx(p, code=CODE)
        print(r.get("status"), os.path.basename(p))
        for i in r.get("issues", []):
            print("    -", i)
