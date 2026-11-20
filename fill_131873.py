# -*- coding: utf-8 -*-
r"""
Q326D131873 — bùn thải PPJ-HNC: ĐIỀN KQTN VÀO FORM TBKQ / BCKQ ĐANG CÓ (bản trên Z, gồm sửa tay của người dùng).

Người dùng (19/11/2026): "đầu tiên xác định đúng mã số niêm của từng mẫu, đúng tên mẫu, rồi điền vào KQ và phương pháp thử
có sẵn theo form nhé, chỉ điền KQ và phương pháp thôi, đừng sửa form … Áp dụng cho các hồ sơ về sau" (AGENTS mục 13).

Ghép mẫu: cột B1 / B2 / B3 của form <-> phiếu KQTN có tên mẫu "Q326D131873 - Bn" VÀ số niêm trên phiếu trùng số niêm
của Bn trong bảng niêm phong của form (B1 KT372818, B2 KT372819, B3 KT372820) — lệch thì dừng, báo.
Điền: ô "………" của cột Phương pháp (theo KQTN), Kết quả B1–B3 (KPH -> "< LOD"), Trung bình (cách tính người dùng chốt
17/11/2026, `trung_binh`), 2 dòng Đánh giá, ngày ban hành ở header (= ngày KQTN, AGENTS mục 2). Không đổi gì khác.
Độ ẩm có trên KQTN nhưng form không có -> không thêm. Hàm lượng tuyệt đối ghi như KQTN (chưa quy đổi rắn khô).

Form gốc (lưu 1 lần trước khi điền): ThucHien\_checkpoint\form_TBKQ.docx, form_BCKQ.docx.
Chạy:  python fill_131873.py [--pdf]
"""
import argparse
import glob
import os
import re
import shutil
import sys

import docx
from docx.oxml.ns import qn

WS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, WS)
from Core.Tools.kqtn_reader import read_kqtn           # noqa: E402
from fill_131229 import fill_runs, txt                  # noqa: E402
from fill_form_nuoc import set_text                     # noqa: E402

IN = os.path.join(WS, "Input", "131873")
CODE = "Q326D131873"
PH = "………"
TBKQ = os.path.join(IN, "TBKQ", f"TBKQ - viet - bun thai - {CODE}.docx")
BCKQ = os.path.join(IN, "ThucHien", f"BCKQ - bun thai - {CODE}.docx")
CK = os.path.join(IN, "ThucHien", "_checkpoint")
SUP = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
SUP_R = str.maketrans("0123456789", "⁰¹²³⁴⁵⁶⁷⁸⁹")


def key_of(name):
    """Tên chỉ tiêu trên form -> (khóa tên KQTN, đơn vị KQTN mong đợi hoặc None)."""
    n = name.lower()
    unit = "mg/kg" if "tuyệt đối" in n else ("mg/L" if "ngâm chiết" in n else None)
    for f, k in (("ph", "độ ph"), ("cadimi", "cadimi"), ("niken", "niken"), ("bari", "bari"), ("coban", "coban"),
                 ("thủy ngân", "thủy ngân"), ("asen", "arsen"), ("chì", "chì"), ("selen", "selen"),
                 ("phenol", "phenol"), ("clo benzen", "clo benzen"), ("benzen", "hàm lượng benzen"),
                 ("xyanua", "xyanua"), ("crom vi", "crom vi"), ("tổng dầu", "tổng dầu")):
        if (f == "ph" and n.startswith("tính kiềm")) or (f != "ph" and f in n):
            return k, unit
    return None, None


def find_row(k, name):
    key, unit = key_of(name)
    rows = [r for r in k["rows"] if key and key in r.name.lower()]
    if unit:
        rows = [r for r in rows if r.unit.startswith(unit)]
    return rows[0] if len(rows) == 1 else None


def num(v):
    """'4,91 x 10³' -> 4910.0; '7,3' -> 7.3; trả (số, số chữ số lẻ, mũ hoặc None)."""
    s = v.replace(" ", "").replace(" ", "").replace("×", "x")
    m = re.match(r"^([\d,]+)x10([⁰¹²³⁴⁵⁶⁷⁸⁹]+)$", s)
    if m:
        a = m.group(1)
        return float(a.replace(",", ".")) * 10 ** int(m.group(2).translate(SUP)), len(a.split(",")[1]) if "," in a else 0, int(m.group(2).translate(SUP))
    return float(s.replace(",", ".")), (len(s.split(",")[1]) if "," in s else 0), None


def trung_binh(vals):
    """Người dùng chốt 17/11/2026: KPH tính 0, cộng các ngày có số chia 3; trung bình nhỏ hơn giới hạn phát hiện ->
    ghi giới hạn phát hiện. Từ 19/11/2026 KPH ghi giới hạn phát hiện KHÔNG dấu "<" (ket_qua()).
    vals: [(chuỗi ghi, kph, lod)]. Trả (chuỗi trung bình, dưới_LOD)."""
    lods, nums, dec, exp = [], [], 0, None
    for v, kph, lod in vals:
        if kph:
            lods.append(float(lod.replace(",", ".")))
            continue
        s_ = v.replace(" ", "")
        if s_.startswith("<"):                                 # < LOQ: coi như LOD (không có ở 131873)
            lods.append(float(s_[1:].replace(",", ".")))
            continue
        x, d, e = num(v)
        nums.append(x)
        dec, exp = max(dec, d), (e if e is not None else exp)
    lod = max(lods) if lods else None
    avg = sum(nums) / 3
    if lod is not None and avg < lod:
        return f"{lod:g}".replace(".", ","), True
    if exp is not None:                                      # giữ dạng a x 10ⁿ như KQTN
        return f"{avg / 10 ** exp:.{dec}f}".replace(".", ",") + " x 10" + str(exp).translate(SUP_R), False
    return f"{avg:.{dec}f}".replace(".", ","), False


def limit_exceeded(avg, lim):
    if avg.startswith("<") or not lim:
        return False
    x = num(avg)[0]
    m = re.search(r"(≥|>|≤)\s*([\d\s]+(?:,\d+)?)", lim.replace(" ", " "))
    if "pH" in lim:
        return x >= 12.5 or x <= 2.0
    if not m:
        return False
    t = float(m.group(2).replace(" ", "").replace(",", "."))
    return x >= t if m.group(1) == "≥" else (x > t if m.group(1) == ">" else False)


def samples(d):
    """Cột mẫu -> số niêm theo bảng niêm phong của form."""
    t = next(t for t in d.tables if t.rows[0].cells[0].text.strip().startswith("Người chứng kiến"))
    out = {}
    for r in t.rows[1:]:
        cs = [c.text.strip() for c in r.cells]
        out[cs[1]] = cs[-1]
    return out


def match_kqtn(seals):
    by = {}
    for f in glob.glob(os.path.join(IN, "TBKQ", "Edoc", f"KQ-{CODE}-*.pdf")):
        k = read_kqtn(f)
        ky = k["ten_mau"].split("-")[-1].strip()
        if ky in seals:
            if k["niem"] != seals[ky]:
                raise SystemExit(f"[DỪNG] {ky}: niêm trên KQTN {k['niem']} khác form {seals[ky]}")
            by[ky] = k
            print(f"   {ky}: phiếu {k['so_phieu']}  tên mẫu '{k['ten_mau']}'  niêm {k['niem']} = form {seals[ky]}")
    missing = set(seals) - set(by)
    if missing:
        raise SystemExit(f"[DỪNG] thiếu KQTN cho {missing}")
    return by


methods = {}                                                # tên chỉ tiêu -> {mẫu: phương pháp KQTN}


def fill_tables(d, by, report):
    high, low = [], []
    for t in d.tables:
        trs = t._tbl.findall(qn("w:tr"))
        if len(trs) < 3 or "Phương pháp" not in txt(trs[0]):
            continue
        sub = [txt(tc).strip() for tc in trs[1].findall(qn("w:tc"))]
        cols = [i for i, s in enumerate(sub) if s in by]                # vị trí ô B1, B2, B3 (theo hàng 2)
        head = [txt(tc).strip() for tc in trs[0].findall(qn("w:tc"))]
        for tr in trs[2:]:
            tcs = tr.findall(qn("w:tc"))
            if PH not in txt(tr):
                continue
            name = txt(tcs[1 if head[0] == "STT" else 0]).strip()
            mcol = next(i for i, h in enumerate(head) if h == "Phương pháp")
            vals, meth = [], None
            for ky, ci in zip(sub[cols[0]:cols[-1] + 1], range(len(cols))):
                r = find_row(by[ky], name)
                if r is None:
                    report.append(f"KHÔNG ghép được '{name}' với KQTN {ky} — để ………")
                    vals = None
                    break
                v, _, kph = r.ket_qua()                          # KPH -> giới hạn phát hiện, không "<"
                v = re.sub(r"⁽[^⁾]*⁾", "", v).strip()
                vals.append((v, kph, r.lod))
                m_ = re.sub(r"⁽[^⁾]*⁾", "", r.method).strip()
                methods.setdefault(name, {})[ky] = m_
                meth = meth or m_
            if vals is None:
                continue
            # vị trí ô trong hàng dữ liệu: hàng dữ liệu không có ô gộp ngang -> cột theo hàng 2
            data_cols = [mcol + 1 + i for i in range(3)] if len(tcs) >= mcol + 5 else []
            set_text(tcs[mcol], meth)
            for ci, (v, _, _) in zip(data_cols, vals):
                set_text(tcs[ci], v)
            avg, below = trung_binh(vals)
            set_text(tcs[mcol + 4], avg)
            lim = txt(tcs[mcol + 5]).strip() if len(tcs) > mcol + 5 else ""
            if lim:
                hi = (not below) and limit_exceeded(avg, lim)
                (high if hi else low).append(name.split(" (")[0] if "pH" not in name else "pH")
            report.append(f"{name}: {meth} | {' / '.join(v for v, _, _ in vals)} | TB {avg}")
    return high, low


def fill_first(p, value):
    """Thay "………" ĐẦU TIÊN trong đoạn (đoạn có 2 chỗ trống: "Phiếu KQTN số ……… ngày ………")."""
    for t in p.iter(qn("w:t")):
        if t.text and PH in t.text:
            t.text = t.text.replace(PH, value, 1)
            return True
    return False


def fill_bckq(by, date):
    """BCKQ: chỉ điền các ô "………" theo dõi KQTN (số phiếu từng mẫu, ngày). Ngày gửi mẫu: không có biên bản gửi mẫu ->
    ngày nhận mẫu trên KQTN; ngày yêu cầu cấp: không có PYC thử nghiệm -> tạm ngày lấy mẫu cuối + 8 (AGENTS mục 13)."""
    import datetime
    import unicodedata
    from Core.Tools.kqtn_late import tre_han_text
    d = docx.Document(os.path.join(CK, "form_BCKQ.docx"))
    send = max((k["ngay_nhan_mau"] for k in by.values()), key=lambda s: tuple(reversed(s.split("/"))))
    last = datetime.datetime.strptime("31/10/2026", "%d/%m/%Y")
    req = (last + datetime.timedelta(days=8)).strftime("%d/%m/%Y")
    so = "; ".join(by[k]["so_phieu"] for k in sorted(by))
    for p in d.element.body.iter(qn("w:p")):
        t = unicodedata.normalize("NFC", "".join(x.text or "" for x in p.iter(qn("w:t"))))
        if PH not in t:
            continue
        m = re.search(r"KQTN mẫu (B\d)", t)
        if m:
            fill_runs(p, by[m.group(1)]["so_phieu"])
        elif t.startswith("Phiếu KQTN số"):
            fill_first(p, so)
            fill_first(p, date)
        elif t.startswith("Ngày gửi mẫu"):
            fill_runs(p, send)
        elif t.startswith("Ngày nhận KQTN"):
            fill_runs(p, date)
        # người dùng 19/11/2026: "không đánh ô trễ hạn, Ngày yêu cầu cấp để trống" -> giữ "………" 2 ô này
    d.save(BCKQ)
    print(f"[OK] BCKQ  KQTN {so} ngày {date}; gửi mẫu {send}; yêu cầu cấp + trễ hạn để trống")
    return BCKQ


def set_widths(t, pct):
    """Đặt lại độ rộng cột (đơn vị 1/50 %, tổng 5000) cho lưới + mọi ô (ô gộp ngang cộng các cột)."""
    tbl = t._tbl
    grid = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    total = sum(int(g.get(qn("w:w"))) for g in grid)
    for g, w in zip(grid, pct):
        g.set(qn("w:w"), str(round(total * w / 5000)))
    for tr in tbl.findall(qn("w:tr")):
        i = 0
        for tc in tr.findall(qn("w:tc")):
            pr = tc.find(qn("w:tcPr"))
            gs = pr.find(qn("w:gridSpan"))
            n = int(gs.get(qn("w:val"))) if gs is not None else 1
            pr.find(qn("w:tcW")).set(qn("w:w"), str(sum(pct[i:i + n])))
            pr.find(qn("w:tcW")).set(qn("w:type"), "pct")
            i += n


# người dùng 19/11/2026: "Nới ra giúp tôi đi" — cột Trung bình (và B1–B3 bảng 9.2.4) hẹp làm "< 0,01" xuống dòng
WIDTHS = {8: [300, 900, 1200, 500, 500, 500, 550, 550],          # 9.2.2 / 9.2.3: TB 450 -> 550, PP 1300 -> 1200
          "crom": [300, 1100, 450, 950, 550, 550, 550, 550]}     # 9.2.4: B1–B3, TB 450 -> 550; tên, PP bớt


def widen(d):
    for t in d.tables:
        trs = t._tbl.findall(qn("w:tr"))
        if len(trs) < 3 or "Phương pháp" not in txt(trs[0]):
            continue
        ncol = len(t._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")))
        if "Crom VI" in txt(t._tbl):
            set_widths(t, WIDTHS["crom"])
        elif ncol == 8:
            set_widths(t, WIDTHS[8])


def main(pdf):
    os.makedirs(CK, exist_ok=True)
    for src, dst in ((TBKQ, "form_TBKQ.docx"), (BCKQ, "form_BCKQ.docx")):
        if not os.path.exists(os.path.join(CK, dst)):
            shutil.copy2(src, os.path.join(CK, dst))          # form gốc trước khi điền (lần đầu)
    d = docx.Document(os.path.join(CK, "form_TBKQ.docx"))
    seals = samples(d)
    by = match_kqtn(seals)
    date = max(k["ngay"] for k in by.values())
    for s in d.sections:                                      # ngày ban hành = ngày KQTN (AGENTS mục 2)
        for hp in (s.header, s.first_page_header):
            for p in hp.paragraphs:
                for t in p._p.iter(qn("w:t")):
                    if t.text and re.search(r"\d{2}/\d{2}/\d{4}", t.text) and "Ngày" in p.text:
                        t.text = re.sub(r"\d{2}/\d{2}/\d{4}", date, t.text)
                fill_runs(p._p, date)
    report = []
    high, low = fill_tables(d, by, report)
    widen(d)
    for p in d.paragraphs:
        if "cao hơn mức quy định" in p.text and PH in p.text:
            fill_runs(p._p, ", ".join(high) if high else "-")
        elif "thấp hơn mức quy định" in p.text and PH in p.text:
            fill_runs(p._p, "các chỉ tiêu còn lại" if high else "tất cả các chỉ tiêu")
    for h in list(d.element.body.iter(qn("w:highlight"))):
        h.getparent().remove(h)
    d.save(TBKQ)
    print(f"[OK] TBKQ ngày {date}; cao hơn ngưỡng: {high or '-'}")
    # kiểm phương pháp: ô Phương pháp trong TBKQ = phương pháp KQTN của CẢ 3 mẫu
    d2 = docx.Document(TBKQ)
    for t in d2.tables:
        trs = t._tbl.findall(qn("w:tr"))
        if len(trs) < 3 or "Phương pháp" not in txt(trs[0]):
            continue
        head = [txt(tc).strip() for tc in trs[0].findall(qn("w:tc"))]
        mcol = head.index("Phương pháp")
        for tr in trs[2:]:
            tcs = tr.findall(qn("w:tc"))
            name = txt(tcs[1 if head[0] == "STT" else 0]).strip()
            got = txt(tcs[mcol]).strip()
            want = methods.get(name, {})
            ok = want and all(w == got for w in want.values())
            print(f"   PP {'OK ' if ok else 'LỆCH'} {name}: TBKQ '{got}' | KQTN " +
                  "; ".join(f"{k} '{w}'" for k, w in want.items()))
    for x in report:
        print("   -", x)
    docs = [TBKQ, fill_bckq(by, date)]
    if pdf:
        import build_131873 as B
        from nuoc_common import make_dossier
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(make_dossier(dict(B.CONFIG, cert_date=date), None, None, False)).export_files(docs)
    from Core.Tools.qa_check import check_single_docx
    r = check_single_docx(TBKQ, code=CODE)
    print(r.get("status"))
    for i in r.get("issues", []):
        print("    -", i)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    main(ap.parse_args().pdf)
