# -*- coding: utf-8 -*-
r"""
Điền KQTN vào FORM GỐC TBKQ của người dùng — hồ sơ nước nhiều chứng thư (`3. TBKQ - <n> - <mã>.docx` ở gốc thư mục
hồ sơ). Cùng cách với `fill_135226.py` (người dùng 18/11/2026, 131229: "giữ lại form cũ … điền kết quả dựa theo form cũ
… Bạn chỉ điền thôi không cần sửa"; "sửa phương pháp nhưng giữ form"; AGENTS mục 8C nước uống điểm 0).

Chỉ điền: ngày ban hành ở header (ngày KQTN muộn nhất), ô "………" (kết quả theo KQTN; KPH -> "< LOD"; LOQ -> "< LOQ" +
ký hiệu của dòng ghi chú form "(*)" / "(2)" chỉ số trên; mẫu không có LOQ bỏ dòng ghi chú đó), thiết bị (theo BBLM, cách
viết của form), Đánh giá (chỉ form có cột mức + Căn cứ). Nhóm chất (PCBs, BVTV clo / phospho hữu cơ): điền từng dòng
"+ <cấu tử>" theo cấu tử KQTN; cấu tử KQTN không có trên form -> chỉ báo, không thêm dòng (AGENTS mục 13).
Giữ nguyên thứ tự, tên, đơn vị, số đo hiện trường của form. Phương pháp khác KQTN -> sửa chữ, giữ khung (fix_methods).

Quyết định cũ của người dùng áp lên form (xem docstring `build_<6 số>.py`):
  134486 (16/11/2026): quy trình nước sinh hoạt để trống, tô vàng -> giữ "N2.TC.N.02" của form, tô vàng (như 137037).
  135228 (16/11/2026): thiết bị theo BBLM Hanna HI 98190 (tô vàng); ngày 27/10/2026 tô vàng; địa điểm "IDICO – TSC đường
         N5, …" theo BBLM (tô vàng).
  135225: không có quyết định riêng cho TBKQ (bảng mức GPMT 156 chưa có -> không đánh giá, như form).

Chạy:  python fill_form_nhieu.py 134486|135225|135228 [--pdf] [--only 1,3]
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
from Core.Tools.kqtn_reader import read_kqtn_all       # noqa: E402
from fill_131229 import fill_runs, dash_numbering, txt  # noqa: E402
from fill_135226 import highlight                       # noqa: E402
from fill_form_nuoc import set_text, fix_methods, clean, lim_ok   # noqa: E402

PH = "………"
METTLER = "Máy đo pH Seven2Go S2, Mettler Toledo – USA"      # cách viết các form IDICO (135226, 135228)
HACH = "Máy đo Chlorine {}, Hach – USA"                      # như 137102
# tên trên form (chữ thường) -> khóa tìm trong tên KQTN; nhóm chất đứng trước (tránh "photpho" của "photpho hữu cơ")
KEYS = [("biphenyl", "biphenyl"), ("lân hữu cơ", "phospho hữu cơ"), ("photpho hữu cơ", "phospho hữu cơ"),
        ("clo hữu cơ", "chlor hữu cơ"), ("phóng xạ α", "alpha"), ("phóng xạ β", "beta"),
        ("màu", "độ màu"), ("đục", "độ đục"), ("permanganat", "permanganat"), ("sinh ho", "sinh hóa"),
        ("oxy hóa học", "oxy hóa học"), ("lơ lửng", "lơ lửng"), ("asen", "asen"), ("thủy ngân", "thủy ngân"),
        ("chì", "chì"), ("cadimi", "cadmi"), ("xyanua", "xyanua"), ("crom (vi)", "crom vi"), ("crom (iii)", "crom iii"),
        ("đồng", "đồng"), ("kẽm", "kẽm"), ("niken", "niken"), ("mangan", "mangan"), ("sắt", "sắt"),
        ("phenol", "phenol"), ("dầu mỡ", "dầu mỡ"), ("sunfua", "sunfua"), ("florua", "florua"), ("clorua", "clorua"), ("mặn", "mặn"),
        ("amoni", "amoni"), ("nitơ", "nitơ"), ("photpho", "phospho"), ("escherichia", "escherichia"),
        ("coliform", "coliform")]
ALIAS = {"hexaclorobenzen": "hexaclobenzen", "lindanγhch": "γhchlindan", "endosulfansulfate": "endosulfansulfat",
         "methoxychlor": "metoxyclo", "parathionethyl": "parathionethylparathion"}
CFG = {
    "134486": dict(loai="nuoc", kq_dir="KQTN_20-11", n=range(1, 12),
                   proc_fix={n: ("N2.TC.N.02", "N2.TC.N.01") for n in (1, 2, 3)},   # NSH = N2.TC.N.01 (người dùng 20/11)
                   equip={n: [METTLER, HACH.format("DR900")] if n <= 5 else [METTLER] for n in range(1, 12)},
                   std="QCVN 01-1:2024/BYT", water="nước sinh hoạt", photo_cm={1: 4.0, 2: 4.0, 3: 3.2},
                   # người dùng 19/11/2026: thêm chỉ tiêu KQTN có mà form không có — CHỈ hồ sơ này, không thành quy tắc
                   them={n: [("florua", "Florua", "sunfua")] for n in (6, 7, 8, 10)}, them_cau_tu={4, 5},
                   kph_lod=True, cert_date="20/11/2026"),          # người dùng sửa trên Z 20/11: ngày ban hành 20/11                                   # điền lại 20/11/2026: KPH ghi giới hạn phát hiện
    # 135229 (20/11/2026, người dùng: "Phú Mỹ 2 có kết quả rồi, điền vào luôn đi, điền kết quả theo form"): thiết bị theo
    # BBLM (canonical, như 135228); NSH quy trình N2.TC.N.01; clo dư dưới 0,1 -> "< 0,1" (quyết định 16/11); KPH = LOD
    "135229": dict(loai="nuoc", kq_dir="KQTN_TBKQ_Edoc", n=range(1, 17), kph_lod=True,
                   proc_fix={n: ("N2.TC.N.02", "N2.TC.N.01") for n in range(1, 6)},
                   equip={n: (["Thiết bị đo pH Mettler Toledo Seven2Go S2, USA", "Thiết bị đo clo dư Hach DR900, USA"]
                              if n in (1, 2, 3, 4, 5, 7, 12) else
                              ["Thiết bị đo pH Mettler Toledo Seven2Go S2, USA"] if n in (6, 13, 14, 15, 16) else ["-"])
                          for n in range(1, 17)},
                   std="QCVN 01-1:2024/BYT", water="nước sinh hoạt", clo_lt01=True, keep_tail=True,
                   drop={n: ["độ mặn"] for n in (8, 9, 10, 11)}),           # người dùng 20/11: "xóa độ mặn"
    "135225": dict(loai="nuoc thai", kq_dir="Q326D135225 - TBKQ", n=range(1, 7),
                   equip={n: [METTLER, HACH.format("DR300")] if n == 1 else [METTLER] for n in range(1, 7)}),
    "135228": dict(loai="nuoc thai", kq_dir="Q326D135228 - TBKQ", n=range(1, 8),
                   # người dùng 20/11/2026: thiết bị như BBLM, bỏ "IDICO – TSC đường N5", bỏ tô vàng, ngày ban hành 20/11
                   equip={n: ["Thiết bị đo clo dư Hach DR300, USA", "Thiết bị đo pH Hanna HI 98190, USA"] for n in range(1, 8)},
                   equip_force=True, no_hl=True, cert_date="20/11/2026",
                   them={3: [("dầu mỡ", "Tổng dầu mỡ khoáng", "phenol")]}, them_cau_tu={1, 2}),   # người dùng 19/11/2026
}


def kv(c, x):
    """KPH: quy tắc 19/11/2026 (AGENTS mục 13) -> giới hạn phát hiện, không "<" (cờ kph_lod); cũ -> "< LOD"."""
    if c.get("kph_lod"):
        v, loq, _ = x.ket_qua()
        return v, loq
    return x.value()


def norm(s):
    s = s.lower().replace("hàm lượng", "").replace("’", "'").lstrip("+ ")
    s = re.sub(r"[^0-9a-zà-ỹα-ω]", "", s)
    return ALIAS.get(s, s)


def leaves(k, group):
    """Cấu tử lá của nhóm chất (bỏ dòng nhóm con DDTs, heptachlor + epoxide)."""
    rows, i = k["rows"], k["rows"].index(group)
    out = []
    for r in rows[i + 1:]:
        if not r.stt.startswith(group.stt + "."):
            break
        if not any(x.stt.startswith(r.stt + ".") for x in rows):
            out.append(r)
    return out


def add_sup(p, mark):
    rs = p.findall(qn("w:r"))
    r = copy.deepcopy(rs[-1]) if rs else OxmlElement("w:r")
    for ch in list(r):
        if ch.tag != qn("w:rPr"):
            r.remove(ch)
    rp = r.find(qn("w:rPr"))
    if rp is None:
        rp = OxmlElement("w:rPr"); r.insert(0, rp)
    va = OxmlElement("w:vertAlign"); va.set(qn("w:val"), "superscript"); rp.append(va)
    t = OxmlElement("w:t"); t.text = f" {mark}"
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t); p.append(r)


def set_para(p, value):
    """Thay chữ "………" trong đoạn p bằng value, bỏ tô vàng."""
    ts = [t for t in p.iter(qn("w:t")) if t.text]
    ts[0].text = value
    for t in ts[1:]:
        t.text = ""
    for h in list(p.iter(qn("w:highlight"))):
        h.getparent().remove(h)


def fill_equipment(row, val, lines):
    stray = row._tr.getnext()
    if stray is not None and txt(stray).strip() == PH:
        stray.getparent().remove(stray)
    p0 = val.find(qn("w:p"))
    for extra in val.findall(qn("w:p"))[1:]:
        val.remove(extra)
    many = len(lines) > 1
    for i, line in enumerate(lines):
        p = p0 if i == 0 else copy.deepcopy(p0)
        for r in p.findall(qn("w:r")):
            p.remove(r)
        r = OxmlElement("w:r"); rp = OxmlElement("w:rPr")
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag); e.set(qn("w:val"), "26"); rp.append(e)
        r.append(rp); t = OxmlElement("w:t")
        t.text = (f"- {line}{';' if i < len(lines) - 1 else '.'}" if many else line)
        r.append(t); p.append(r)
        if i:
            val.append(p)


def add_assessment(d, nt_tbl, std, water, fails):
    # "Đánh giá:" in đậm đứng đầu câu so sánh (cùng đoạn) — form 1 trang phụ lục không đủ chỗ cho dòng riêng
    paras = [(f"So sánh kết quả thử nghiệm với mức quy định tại {std}:",       # 1 dòng cùng "Đánh giá:"
              "head"),
             (f"+ Thông số không đạt: {', '.join(fails) if fails else '-'}", False),
             (f"+ Thông số đạt: {'các thông số còn lại' if fails else 'tất cả các thông số'}", False)]
    anchor = nt_tbl
    for text, bold in paras:
        p = OxmlElement("w:p"); ppr = OxmlElement("w:pPr")
        if bold:
            ppr.append(OxmlElement("w:keepNext"))
        sp = OxmlElement("w:spacing"); sp.set(qn("w:before"), "60" if bold else "0"); sp.set(qn("w:after"), "0")
        ppr.append(sp)
        if text.startswith("+"):
            ind = OxmlElement("w:ind"); ind.set(qn("w:left"), "426"); ppr.append(ind)
        else:
            jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "both"); ppr.append(jc)
        p.append(ppr)
        r = OxmlElement("w:r"); rp = OxmlElement("w:rPr")
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag); e.set(qn("w:val"), "26"); rp.append(e)
        r.append(rp)
        if bold == "head":
            rh = copy.deepcopy(r); rh.find(qn("w:rPr")).insert(0, OxmlElement("w:b"))
            th = OxmlElement("w:t"); th.text = "Đánh giá: "
            th.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); rh.append(th); p.append(rh)
        t = OxmlElement("w:t"); t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); r.append(t); p.append(r)
        anchor.addnext(p)
        anchor = p


def fill(h6, c, n, kqs, date):
    IN = os.path.join(WS, "Input", h6)
    code = f"Q326D{h6}"
    d = docx.Document(os.path.join(IN, f"3. TBKQ - {n} - {code}.docx"))
    report = []
    ky = txt(d.tables[1]._tbl.findall(qn("w:tr"))[1].findall(qn("w:tc"))[1]).strip()
    k = kqs.get(ky)
    if k is None:
        print(f"[BỎ] /{n} ({ky}): không có KQTN")
        return None
    for s in d.sections:
        for hp in (s.header, s.first_page_header):
            for p in hp.paragraphs:
                fill_runs(p._p, date)
    for row in d.tables[0].rows:
        lab, val = row.cells[0].text.strip(), row.cells[-1]._tc
        if lab.startswith("Địa điểm") and c.get("site_prefix") and not txt(val).startswith(c["site_prefix"]):
            t = next(t for t in val.iter(qn("w:t")) if t.text)
            t.text = c["site_prefix"] + t.text
            highlight(val, [c["site_prefix"].rstrip(", ")])
        if lab.startswith("Quy trình") and n in c.get("proc_fix", {}):
            old, new = c["proc_fix"][n]
            ts = [t for t in val.iter(qn("w:t")) if t.text]
            joined = "".join(t.text for t in ts)
            kk_ = joined.find(old)
            if kk_ >= 0:                                     # thay trong chữ chia nhiều run, giữ định dạng form
                pos = 0
                for t in ts:
                    a, b = pos, pos + len(t.text); pos = b
                    lo, hi = max(a, kk_), min(b, kk_ + len(old))
                    if lo < hi:
                        t.text = t.text[:lo - a] + (new if a <= kk_ < b else "") + t.text[hi - a:]
            for h in list(val.iter(qn("w:highlight"))):
                h.getparent().remove(h)
        if lab.startswith("Quy trình") and n in c.get("hl_proc", []):
            m = re.search(r"N2\.TC\.N\.\d+", txt(val))
            if m:
                highlight(val, [m.group(0)])
        if lab.startswith("Thiết bị"):
            v = txt(val).strip()
            if c.get("equip") and (not v or v == PH or c.get("equip_force")):
                fill_equipment(row, val, c["equip"][n])
            elif c.get("equip_replace") and c["equip_replace"][0] in v:
                set_text(val, v.replace(*c["equip_replace"]))     # chữ form chia nhiều run -> ghi lại cả ô
                highlight(val, [c["equip_replace"][1]])
        for w in c.get("hl", []):
            if w in txt(val):
                highlight(val, [w])
        if lab.startswith("-") and "Giấy yêu cầu" in lab:
            dash_numbering(d, row.cells[0]._tc.find(qn("w:p")))
    nt = next(t for t in d.tables if t.rows[0].cells[0].text.strip().startswith("Ghi chú"))
    note_row = next((tr for tr in nt._tbl.findall(qn("w:tr")) if re.match(r"^\((\*|2)\)", txt(tr).strip())), None)
    mark = re.match(r"^\((\*|2)\)", txt(note_row).strip()).group(0) if note_row is not None else "(*)"
    rt = next(t for t in d.tables if t.rows[0].cells[0].text.strip() in ("Stt", "STT"))
    for drop in c.get("drop", {}).get(n, []):             # bỏ dòng thông số theo người dùng (135229 độ mặn, 20/11)
        for tr in rt._tbl.findall(qn("w:tr"))[1:]:
            tcs = tr.findall(qn("w:tc"))
            if not txt(tcs[1]).strip().lower().startswith(drop):
                continue
            nxt = tr.getnext()
            if nxt is not None and nxt.tag == qn("w:tr"):          # ô gộp dọc bắt đầu ở dòng bỏ -> dòng sau nhận chữ
                for i, tc in enumerate(tcs):
                    pr = tc.find(qn("w:tcPr"))
                    vm = pr.find(qn("w:vMerge")) if pr is not None else None
                    if vm is None or vm.get(qn("w:val")) != "restart":
                        continue
                    ntc = nxt.findall(qn("w:tc"))[i]
                    npr = ntc.find(qn("w:tcPr"))
                    nvm = npr.find(qn("w:vMerge")) if npr is not None else None
                    if nvm is not None and nvm.get(qn("w:val")) != "restart":
                        nn = nxt.getnext()
                        nn_cont = nn is not None and nn.tag == qn("w:tr") and                             (lambda e: e is not None and e.get(qn("w:val")) != "restart")(
                                (nn.findall(qn("w:tc"))[i].find(qn("w:tcPr")) or OxmlElement("w:tcPr")).find(qn("w:vMerge")))
                        if nn_cont:
                            nvm.set(qn("w:val"), "restart")
                        else:
                            npr.remove(nvm)
                        for p_old in ntc.findall(qn("w:p")):
                            ntc.remove(p_old)
                        for p_src in tc.findall(qn("w:p")):
                            ntc.append(copy.deepcopy(p_src))
            rt._tbl.remove(tr)
            report.append(f"BỎ dòng: {txt(tcs[1]).strip()}")
        for i, tr in enumerate(rt._tbl.findall(qn("w:tr"))[1:], 1):
            set_text(tr.findall(qn("w:tc"))[0], str(i))
        st = d.tables[1]._tbl.findall(qn("w:tr"))[1].findall(qn("w:tc"))[2]   # "02 thông số" -> "01 thông số"
        m_ = re.match(r"^\s*(\d+) thông số", txt(st))
        if m_:
            set_text(st, txt(st).replace(m_.group(1) + " thông số", f"{int(m_.group(1)) - 1:02d} thông số", 1).strip())
    has_limit = len(rt.rows[0].cells) >= 6
    used, kq_methods, fails, any_loq, fail_cells = set(), [], [], False, []
    for tr in rt._tbl.findall(qn("w:tr"))[1:]:
        tcs = tr.findall(qn("w:tc"))
        name_ps = tcs[1].findall(qn("w:p"))
        name = txt(name_ps[0]).strip() if name_ps else ""
        cell = tcs[4]
        if c.get("clo_lt01") and name.lower().startswith("clo dư"):        # đo nhanh < 0,1 -> "< 0,1" (AGENTS 8C)
            m_ = re.match(r"^\s*(\d+(?:,\d+)?)\s*$", txt(cell))
            if m_ and float(m_.group(1).replace(",", ".")) < 0.1:
                set_text(cell, "< 0,1")
                report.append(f"{name}: {m_.group(1)} -> < 0,1")
        if PH not in txt(cell):
            kq_methods.append(None)
            if has_limit and not lim_ok(clean(txt(cell).strip()), txt(tcs[5]).strip()):
                fails.append(re.sub(r"\s*\(", " (", txt(tcs[1]).strip().replace(chr(10), " "))); fail_cells.append(cell)
            continue
        key = next((kk for f, kk in KEYS if f in name.lower()), None)
        r = next((r for r in k["rows"] if r.level == 0 and key and key in r.name.lower()), None) or             next((r for r in k["rows"] if key in ("alpha", "beta") and key in r.name.lower()), None)   # cấu tử phóng xạ
        kq_methods.append(r.method if r else None)
        if r is None:
            report.append(f"FORM có, KQTN không: {name} (giữ ……… tô vàng)")
            continue
        used.add(r.stt)
        comps = [txt(p).strip() for p in name_ps if txt(p).strip().startswith("+")]
        if comps:                                          # nhóm chất: từng dòng "+ cấu tử"
            lv = leaves(k, r)
            used.update(x.stt for x in lv)
            phs = [p for p in cell.findall(qn("w:p")) if PH in txt(p)]
            for comp, p in zip(comps, phs):
                x = next((x for x in lv if norm(x.name) == norm(comp)), None) or \
                    next((x for x in lv if norm(comp) in norm(x.name) or norm(x.name) in norm(comp)), None)
                if x is None:
                    report.append(f"  cấu tử form không có trên KQTN: {comp}")
                    continue
                lv.remove(x)
                v, loq = kv(c, x)
                set_para(p, clean(v))
                if loq:
                    any_loq = True
                    add_sup(p, mark)
            for p in cell.findall(qn("w:p")):               # đoạn trống của ô nhóm: bỏ tô vàng sót của form
                if not txt(p).strip():
                    for h in list(p.iter(qn("w:highlight"))):
                        h.getparent().remove(h)
            if len(comps) != len(phs):
                report.append(f"  {name}: {len(comps)} cấu tử / {len(phs)} ô kết quả")
            if lv and n in c.get("them_cau_tu", ()):         # thêm cấu tử KQTN có mà form không có (19/11/2026)
                np_, rp_ = [p for p in name_ps if txt(p).strip().startswith("+")][-1], phs[-1] if phs else None
                for x in lv:
                    nm = "+ " + re.sub(r"^hàm lượng\s+", "", x.name, flags=re.I)
                    nm = nm[:2] + nm[2].upper() + nm[3:]
                    a = copy.deepcopy(np_); set_para(a, nm); np_.addnext(a); np_ = a
                    if rp_ is not None:
                        v, loq = kv(c, x)
                        b = copy.deepcopy(rp_)
                        for va in list(b.iter(qn("w:vertAlign"))):
                            va.getparent().getparent().getparent().remove(va.getparent().getparent())
                        set_para(b, clean(v)); rp_.addnext(b); rp_ = b
                        if loq:
                            any_loq = True
                            add_sup(b, mark)
                    report.append(f"  THÊM cấu tử: {r.name} — {nm} = {kv(c, x)[0]}")
                lv = []
                for rowx in rt._tbl.findall(qn("w:tr")):     # chiều cao ép của form làm dòng dài ra, chẻ trang trống
                    if "hữu cơ" in txt(rowx) or "biphenyl" in txt(rowx).lower():
                        trp = rowx.find(qn("w:trPr"))
                        for h in (trp.findall(qn("w:trHeight")) if trp is not None else []):
                            trp.remove(h)
            for x in lv:
                report.append(f"  KQTN có, form không: {r.name} — {x.name} = {kv(c, x)[0]}")
            continue
        v, loq = kv(c, r)
        v = clean(v)
        set_text(cell, v)
        if loq:
            any_loq = True
            add_sup(cell.find(qn("w:p")), mark)
        if has_limit and not r.kph and not lim_ok(v, txt(tcs[5]).strip()):   # KPH luôn dưới mức
            fails.append(re.sub(r"\s*\(", " (", txt(tcs[1]).strip().replace(chr(10), " "))); fail_cells.append(cell)
    fix_methods(rt._tbl, kq_methods, report)
    for key, form_name, after in c.get("them", {}).get(n, []):   # thêm dòng chỉ tiêu (người dùng 19/11/2026)
        r = next(r for r in k["rows"] if r.level == 0 and key in r.name.lower())
        if r.stt in used:
            continue
        trs = rt._tbl.findall(qn("w:tr"))[1:]
        ref = next(tr for tr in trs if after in txt(tr.findall(qn("w:tc"))[1]).lower())
        nxt = ref.getnext()
        new_tr = copy.deepcopy(ref)
        tcs = new_tr.findall(qn("w:tc"))
        pr = tcs[3].find(qn("w:tcPr"))
        for e in (pr.findall(qn("w:vMerge")) if pr is not None else []):
            pr.remove(e)
        if not list(tcs[3].iter(qn("w:r"))):                  # ô phương pháp "tiếp nối" không có chữ
            tcs[3].find(qn("w:p")).append(copy.deepcopy(next(tcs[2].iter(qn("w:r")))))
        v, loq = kv(c, r)
        for tc, val in ((tcs[1], form_name), (tcs[2], "mg/l"), (tcs[3], r.method), (tcs[4], clean(v))):
            set_text(tc, val)
        if loq:
            any_loq = True
            add_sup(tcs[4].find(qn("w:p")), mark)
        ref.addnext(new_tr)
        if nxt is not None:                                    # dòng sau đang "tiếp nối" ô gộp của ref -> tách
            npr = nxt.findall(qn("w:tc"))[3].find(qn("w:tcPr"))
            if npr is not None and npr.find(qn("w:vMerge")) is not None and not npr.find(qn("w:vMerge")).get(qn("w:val")):
                report.append("  CẢNH BÁO: chèn giữa ô phương pháp gộp — soát lại")
        used.add(r.stt)
        report.append(f"THÊM dòng: {form_name} | {r.method} | {clean(v)}")
    if c.get("them", {}).get(n):                               # đánh lại STT
        for i, tr in enumerate(rt._tbl.findall(qn("w:tr"))[1:], 1):
            set_text(tr.findall(qn("w:tc"))[0], str(i))
    for r in k["rows"]:
        if r.stt not in used and not any(r.stt.startswith(u + ".") for u in used):
            report.append(f"KQTN có, form không: {r.name} = {kv(c, r)[0]}")
    if not any_loq and note_row is not None:
        nt._tbl.remove(note_row)
    for cell_ in fail_cells:                                   # giá trị không đạt in đậm (AGENTS mục 3)
        for r_ in cell_.iter(qn("w:r")):
            rp_ = r_.find(qn("w:rPr"))
            if rp_ is None:
                rp_ = OxmlElement("w:rPr"); r_.insert(0, rp_)
            if rp_.find(qn("w:b")) is None:
                rp_.insert(0, OxmlElement("w:b"))
    if has_limit:                                              # đã đánh giá theo KQTN -> bỏ tô vàng đề mục "Đánh giá:" của form
        for p in d.paragraphs:
            if p.text.strip().startswith("Đánh giá"):
                for h in list(p._p.iter(qn("w:highlight"))):
                    h.getparent().remove(h)
    if has_limit and fails and "Đánh giá" in " ".join(p.text for p in d.paragraphs):   # form có sẵn câu "tất cả đạt"
        for p in d.paragraphs:
            t_ = p.text
            if re.match(r"^\s*\+ (Thông số|Chỉ tiêu) không đạt", t_):
                set_para(p._p, t_.split(":")[0] + ": " + ", ".join(fails))
                report.append(f"Đánh giá: không đạt -> {', '.join(fails)}")
            elif re.match(r"^\s*\+ (Thông số|Chỉ tiêu) đạt", t_):
                set_para(p._p, t_.split(":")[0] + ": các " + ("thông số" if "Thông số" in t_ else "chỉ tiêu") + " còn lại")
    if has_limit and c.get("std") and "Đánh giá" not in " ".join(p.text for p in d.paragraphs):
        add_assessment(d, nt._tbl, c["std"], c["water"], fails)
        report.append(f"thêm đoạn Đánh giá (không đạt: {fails or '-'})")
        if c.get("keep_tail"):                     # 2 dòng cuối bảng + Ghi chú + Đánh giá cùng trang (AGENTS mục 3)
            def _kn(el):
                for p_ in el.iter(qn("w:p")):
                    ppr_ = p_.find(qn("w:pPr"))
                    if ppr_ is None:
                        ppr_ = OxmlElement("w:pPr"); p_.insert(0, ppr_)
                    if ppr_.find(qn("w:keepNext")) is None:
                        ppr_.insert(0, OxmlElement("w:keepNext"))
            for tr_ in rt._tbl.findall(qn("w:tr"))[-3:]:
                _kn(tr_)
            _kn(nt._tbl)
            gp_ = nt._tbl.getprevious()
            if gp_ is not None and gp_.tag == qn("w:p"):
                _kn(gp_)
        # form 1 trang phụ lục đã kín: thêm Đánh giá thì thu ảnh + đoạn trống trước Ghi chú để vẫn 1 trang (đoạn
        # cuối của form chứa đường kết thúc neo nhiều hình — rớt sang trang mới là vỡ bố cục)
        if c.get("photo_cm"):
            from docx.shared import Cm
            from docx.shared import Pt
            w = Cm(c["photo_cm"][n])                     # /3: ô vị trí 4 dòng -> ảnh nhỏ hơn
            for sh in d.inline_shapes:
                f = w / sh.width
                sh.width, sh.height = w, int(sh.height * f)
            for p in d.paragraphs:                      # đoạn ảnh sát dưới, tiêu đề 9.2 bớt khoảng trên
                if p._p.findall(".//" + qn("w:drawing")) and not p.text.strip():
                    p.paragraph_format.space_after = Pt(0)
                if p.text.strip() == "Kết quả giám định":
                    p.paragraph_format.space_before = Pt(3)
                    p.paragraph_format.space_after = Pt(3)
        last = d.element.body.findall(qn("w:p"))[-1]       # đoạn chứa đường kết thúc: cao 1pt exact
        if not txt(last).strip():
            lp = last.find(qn("w:pPr"))
            if lp is None:
                lp = OxmlElement("w:pPr"); last.insert(0, lp)
            for e in lp.findall(qn("w:spacing")):
                lp.remove(e)
            sp = OxmlElement("w:spacing")
            for k_, v_ in (("w:before", "0"), ("w:after", "0"), ("w:line", "20"), ("w:lineRule", "exact")):
                sp.set(qn(k_), v_)
            lp.append(sp)
        gap = nt._tbl.getprevious()
        if gap is not None and gap.tag == qn("w:p") and not txt(gap).strip():
            ppr = gap.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr"); gap.insert(0, ppr)
            for e in ppr.findall(qn("w:spacing")):
                ppr.remove(e)
            sp = OxmlElement("w:spacing")
            for k_, v_ in (("w:before", "0"), ("w:after", "0"), ("w:line", "20"), ("w:lineRule", "exact")):
                sp.set(qn(k_), v_)
            ppr.append(sp)
    if c.get("no_hl"):
        for h in list(d.element.body.iter(qn("w:highlight"))):
            h.getparent().remove(h)
    out = os.path.join(IN, "TBKQ", f"TBKQ - viet - {n} - {c['loai']} - {code}.docx")
    d.save(out)
    print(f"[OK] /{n} {ky}  LOQ: {any_loq}  {os.path.basename(out)}")
    for x in report:
        print("   -", x)
    return out


def main(h6, pdf, only):
    c = CFG[h6]
    IN = os.path.join(WS, "Input", h6)
    kqs = {}
    for f in glob.glob(os.path.join(IN, c["kq_dir"], "**", f"KQ-*{h6}*.pdf"), recursive=True):
        for k in read_kqtn_all(f):
            kqs[k["ten_mau"].split("-")[-1].strip()] = k
    date = c.get("cert_date") or max((k["ngay"] for t, k in kqs.items() if not t.upper().startswith("QC")),
               key=lambda s: tuple(reversed(s.split("/"))))
    docs = [x for x in (fill(h6, c, n, kqs, date) for n in c["n"] if not only or n in only) if x]
    if pdf:
        mod = __import__(f"build_{h6}")
        from nuoc_common import make_dossier
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(make_dossier(dict(mod.CONFIG, cert_date=date), None, None, False)).export_files(docs)
    from Core.Tools.qa_check import check_single_docx
    for p in docs:
        r = check_single_docx(p, code=f"Q326D{h6}")
        print(r.get("status"), os.path.basename(p))
        for i in r.get("issues", []):
            print("    -", i)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("ho_so")
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    main(a.ho_so, a.pdf, {int(x) for x in a.only.split(",") if x.strip()})
