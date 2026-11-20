# -*- coding: utf-8 -*-
r"""
Q326D142705 — Công ty TNHH On Semiconductor Việt Nam — nước uống (máy lọc nước) quý 4, 15 mẫu NU1 – NU15, 15 chứng thư
/1 – /15, 5 chỉ tiêu vi sinh phụ lục III QCVN 06-1:2010/BYT. ĐIỀN VÀO FORM GỐC của người dùng (AGENTS 8C nước uống điểm 0).

Nguồn (chép từ Z `2026\Q326D142705 - ON SEMICONDUCTOR - Nước uống (Chị Loan) mẫu nước uống Quý 4` về Input\142705, chỉ đọc):
  - `3. Thông báo kết quả\3. TBKQ - <n> - Q326D142705.docx` (form TBKQ người dùng, 12/11/2026).
  - `2. Báo cáo kết quả\BCKQ - Q326D142705.docx` (form BCKQ người dùng, đã có 15 ảnh).
  - `4. E.doc\142705 - BBLM.pdf` (3 trang scan): ngày 12/11/2026, bắt đầu 09h30 (chữ viết tay, chưa chắc "09"/"05"),
    kết thúc 12h00; không ghi thiết bị; QTGĐ 3030.2025; N2.TC.N.07; NU1 – NU15, niêm phong "-"; dòng 16 gạch bỏ; đại diện
    KH Lê Hồng Tuyền (?); đại diện TTKT3 ký "Nguyễn Đức Tuấn" (?).
  - KQTN `4. E.doc\KQ-Q326D142705-VSTN26029833.01 – .15.pdf` (mẫu BM105.02.VIL), ngày 19/11/2026, nhận mẫu 12/11/2026.
    NU1 – NU14: cả 5 chỉ tiêu "<1⁽ᶜ⁾"; NU15: Pseudomonas aeruginosa 5,0 × 10² CFU/250 mL (KHÔNG ĐẠT), còn lại "<1".

Quyết định của người dùng (19/11/2026): "hồ sơ 142705 đã có KQTN, hãy cập nhật và dựng hồ sơ, ngoài ra ngay chỗ tôi chụp,
  nếu không có thông tin gì thì xóa luôn bảng đó đi, để dòng phía trên là đủ rồi (thông tin về quá trình lấy mẫu và niêm
  phong)" -> bảng Người chứng kiến toàn "-" bị xóa (TBKQ và BCKQ), giữ dòng "- Thông tin về quá trình lấy mẫu …: -".

Suy luận của AI (cần người dùng xác nhận):
  1. Kết quả "<1⁽ᶜ⁾" ghi "< 1" (như 131229 đã duyệt, AGENTS mục 13); hồ sơ quý trước 050153 ghi "Không phát hiện".
     Người dùng 19/11/2026: "mục 1 cứ giữ <1 trước, tô vàng lên là được", sau đó "bỏ tô vàng tất cả đi" -> giữ "< 1",
     không tô vàng. Mục 2, 3, 4 người dùng
     xác nhận đúng; xóa bảng người chứng kiến trống: "áp dụng chỉ hồ sơ này" (không thành quy tắc).
  2. Tên chỉ tiêu giữ theo form, chỉ sửa lỗi chính tả "Streptococci feacal" -> "Streptococci faecal" (AGENTS 8C điểm 3).
  3. Thiết bị "-" (BBLM không ghi thiết bị, chỉ lấy mẫu vi sinh — như 050153).
  4. Ngày gửi mẫu: không có biên bản gửi mẫu -> tạm ngày nhận mẫu trên KQTN 12/11/2026 (AGENTS mục 13).
  5. BCKQ "QCVN 6-1:2010/BYT" -> "QCVN 06-1:2010/BYT" (AGENTS 8C điểm 3).

Chạy:  python fill_142705.py [--pdf] [--only 1,15]
"""
import argparse
import copy
import glob
import os
import re
import sys

import docx
from docx.oxml.ns import qn

WS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, WS)
from Core.Tools.kqtn_reader import read_kqtn           # noqa: E402
from fill_131229 import fill_runs, txt                  # noqa: E402
from fill_form_nuoc import set_text, fix_methods        # noqa: E402
from fill_form_nhieu import fill_equipment              # noqa: E402

IN = os.path.join(WS, "Input", "142705")
CODE = "Q326D142705"
PH = "………"
KEYS = [("escherichia", "escherichia"), ("coliform", "coliforms"), ("streptococ", "enterococci"),
        ("pseudomonas", "pseudomonas"), ("bào tử", "bào tử")]
SEND_DATE = "12/11/2026"          # không có biên bản gửi mẫu -> ngày nhận mẫu trên KQTN
CONFIG = dict(code6="142705", client_vn="Công ty TNHH On Semiconductor Việt Nam",
              address_vn="Số 10, đường 17A, khu công nghiệp Biên Hòa 2, phường Trấn Biên, Thành phố Đồng Nai",
              reg_date="09/11/2026", dates=["12/11/2026"], time_bckq="Ngày 12/11/2026 từ 09:30 đến 12:00",
              site_vn="Nhà máy của Công ty TNHH On Semiconductor Việt Nam, số 10, đường 17A, khu công nghiệp Biên Hòa 2, "
                      "phường Trấn Biên, Thành phố Đồng Nai",
              scope_vn="Lấy mẫu và giám định chất lượng nước uống", procedure="QUATEST 3 3030:2025, N2.TC.N.07",
              equip=[], executor="Tuấn")


def clean(v):
    v = re.sub(r"⁽[^⁾]*⁾", "", v).strip()
    return re.sub(r"^<\s*", "< ", v)


def drop_empty_witness(d):
    """Bảng Người chứng kiến chỉ có "-" -> xóa (người dùng 19/11/2026); giữ dòng chữ phía trên."""
    for t in d.tables:
        if t.rows[0].cells[0].text.strip().startswith("Người chứng kiến"):
            data = {c.text.strip() for r in t.rows[1:] for c in r.cells}
            if data <= {"-", ""}:
                t._tbl.getparent().remove(t._tbl)
                return True
    return False


def kqtn():
    out = {}
    for f in glob.glob(os.path.join(IN, "4. E.doc", "KQ-*.pdf")):
        k = read_kqtn(f)
        out[k["ten_mau"].split("-")[-1].strip()] = k
    return out


def bold(cell):
    for r in cell.iter(qn("w:r")):
        rp = r.find(qn("w:rPr"))
        if rp is not None and rp.find(qn("w:b")) is None:
            from docx.oxml import OxmlElement
            rp.insert(0, OxmlElement("w:b"))


def fill_tbkq(n, kqs, date):
    d = docx.Document(os.path.join(IN, "3. Thông báo kết quả", f"3. TBKQ - {n} - {CODE}.docx"))
    ky = txt(d.tables[1]._tbl.findall(qn("w:tr"))[1].findall(qn("w:tc"))[1]).strip()
    k = kqs[ky]
    for s in d.sections:
        for hp in (s.header, s.first_page_header):
            for p in hp.paragraphs:
                fill_runs(p._p, date)
    for row in d.tables[0].rows:
        lab, val = row.cells[0].text.strip(), row.cells[-1]._tc
        if lab.startswith("Thiết bị") and txt(val).strip() in ("", PH):
            fill_equipment(row, val, ["-"])
    rt = next(t for t in d.tables if t.rows[0].cells[0].text.strip() in ("Stt", "STT"))
    fails, passed, kq_methods = [], [], []
    for tr in rt._tbl.findall(qn("w:tr"))[1:]:
        tcs = tr.findall(qn("w:tc"))
        name = txt(tcs[1]).strip()
        if "feacal" in name:                                    # lỗi chính tả của form
            for t in tcs[1].iter(qn("w:t")):
                t.text = (t.text or "").replace("feacal", "faecal")
            name = name.replace("feacal", "faecal")
        key = next(kk for f, kk in KEYS if f in name.lower())
        r = next(r for r in k["rows"] if key in r.name.lower())
        kq_methods.append(r.method)
        v = clean(r.result)
        set_text(tcs[4], v)
        if v.startswith("<"):
            passed.append(name)
        else:                                                   # mức "Không phát hiện"
            fails.append(name)
            bold(tcs[4])
    report = []
    fix_methods(rt._tbl, kq_methods, report)
    for p in d.paragraphs:
        if "Chỉ tiêu không đạt" in p.text and PH in p.text:
            fill_runs(p._p, ", ".join(fails) if fails else "-")
        elif "Chỉ tiêu đạt" in p.text and PH in p.text:
            fill_runs(p._p, ("các chỉ tiêu còn lại" if fails else "tất cả các chỉ tiêu"))
    for h in list(d.element.body.iter(qn("w:highlight"))):     # tô vàng của ô đã điền
        h.getparent().remove(h)
    nt = next(t for t in d.tables if t.rows[0].cells[0].text.strip().startswith("Ghi chú"))
    gap = nt._tbl.getprevious()                                # đoạn trống trước Ghi chú -> 1pt (AGENTS [Note Gap])
    if gap is not None and gap.tag == qn("w:p") and not txt(gap).strip():
        from docx.oxml import OxmlElement
        ppr = gap.find(qn("w:pPr"))
        for e in ppr.findall(qn("w:spacing")):
            ppr.remove(e)
        sp = OxmlElement("w:spacing")
        for k_, v_ in (("w:before", "0"), ("w:after", "0"), ("w:line", "20"), ("w:lineRule", "exact")):
            sp.set(qn(k_), v_)
        ppr.append(sp)
    w = drop_empty_witness(d)
    os.makedirs(os.path.join(IN, "TBKQ"), exist_ok=True)
    out = os.path.join(IN, "TBKQ", f"TBKQ - viet - {n} - nuoc uong - {CODE}.docx")
    d.save(out)
    print(f"[OK] /{n} {ky}  không đạt: {fails or '-'}  xóa bảng người chứng kiến: {w}")
    for x in report:
        print("   -", x)
    return out, ky, fails


def fill_bckq(kqs, date, results):
    d = docx.Document(os.path.join(IN, "2. Báo cáo kết quả", f"BCKQ - {CODE}.docx"))
    for row in d.tables[0].rows:
        lab, val = row.cells[0].text.strip(), row.cells[-1]._tc
        if lab.startswith("Thiết bị") and txt(val).strip() in ("", PH):
            set_text(val, "-")
    drop_empty_witness(d)
    fail_keys = [(ky, f) for _, ky, f in results if f]
    et = next(t for t in d.tables if t.rows[0].cells[0].text.strip() == "Ký hiệu mẫu"
              and "đạt" in t.rows[0].cells[1].text)
    r1 = et._tbl.findall(qn("w:tr"))[1]
    tcs = r1.findall(qn("w:tc"))
    for t in tcs[3].iter(qn("w:t")):
        t.text = (t.text or "").replace("QCVN 6-1", "QCVN 06-1")
    all_names = "Escherichia coli, Coliform tổng số, Streptococci faecal, Pseudomonas aeruginosa, Bào tử vi khuẩn kị khí khử sulfit"
    if fail_keys:
        ok_keys = [ky for _, ky, f in results if not f]
        last_ok = ok_keys[-1]
        ts = [t for t in tcs[0].iter(qn("w:t")) if t.text]     # "NU1 → NU15" (chữ chia run 'NU1 ', ' NU1', '5')
        joined = "".join(t.text for t in ts)
        i = joined.rfind("NU15")
        if i >= 0:
            new = joined[:i] + last_ok + joined[i + 4:]
            pos = 0
            for t in ts:                                        # trả lại chữ theo độ dài từng run
                L = len(t.text)
                t.text, pos = new[pos:pos + L] if t is not ts[-1] else new[pos:], pos + L
        set_text(tcs[1], "Tất cả các chỉ tiêu")
        set_text(tcs[2], "-")
        anchor = r1
        for ky, f in fail_keys:
            nr = copy.deepcopy(r1)
            c = nr.findall(qn("w:tc"))
            set_text(c[0], ky)
            set_text(c[1], ", ".join(x for x in all_names.split(", ") if x not in f))
            set_text(c[2], ", ".join(f))
            bold(c[2])
            anchor.addnext(nr)
            anchor = nr
        for tr in et._tbl.findall(qn("w:tr")):                 # dòng không chẻ sang trang
            trp = tr.find(qn("w:trPr"))
            if trp is None:
                from docx.oxml import OxmlElement
                trp = OxmlElement("w:trPr"); tr.insert(0, trp)
            if trp.find(qn("w:cantSplit")) is None:
                from docx.oxml import OxmlElement
                trp.append(OxmlElement("w:cantSplit"))
    else:
        set_text(tcs[1], "Tất cả các chỉ tiêu")
        set_text(tcs[2], "-")
    from datetime import date as _d
    today = _d.today().strftime("%d/%m/%Y")
    fills = {"Phiếu KQTN số": date, "Ngày gửi mẫu": SEND_DATE, "Ngày nhận KQTN": date,
             "Ngày kết thúc hồ sơ thực tế": today, "Ngày báo cáo": today}
    import unicodedata
    for p in d.paragraphs:
        pt = unicodedata.normalize("NFC", p.text).strip()        # form gõ chữ tổ hợp (NFD)
        for lab, v in fills.items():
            if pt.startswith(lab) and PH in pt:
                fill_runs(p._p, v)
    for h in list(d.element.body.iter(qn("w:highlight"))):
        h.getparent().remove(h)
    os.makedirs(os.path.join(IN, "ThucHien"), exist_ok=True)
    out = os.path.join(IN, "ThucHien", f"BCKQ - nuoc uong - {CODE}.docx")
    d.save(out)
    print(f"[OK] BCKQ  không đạt: {fail_keys or '-'}")
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    only = {int(x) for x in a.only.split(",") if x.strip()}
    kqs = kqtn()
    date = max(k["ngay"] for k in kqs.values())
    res = [fill_tbkq(n, kqs, date) for n in range(1, 16) if not only or n in only]
    docs = [r[0] for r in res]
    if not only:
        docs.append(fill_bckq(kqs, date, res))
    if a.pdf:
        from nuoc_common import make_dossier
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(make_dossier(dict(CONFIG, cert_date=date), None, None, False)).export_files(docs)
    from Core.Tools.qa_check import check_single_docx
    for p in docs:
        r = check_single_docx(p, code=CODE)
        print(r.get("status"), os.path.basename(p))
        for i in r.get("issues", []):
            print("    -", i)
