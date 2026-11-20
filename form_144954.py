# -*- coding: utf-8 -*-
r"""
Q326D144954 — Orro Home: DỰNG TBKQ /1–/5 + BCKQ TỪ KHUÔN CỦA NGƯỜI DÙNG (thay bản dựng bằng nuoc_common 17/11/2026).

Người dùng (19/11/2026): "hồ sơ 144954 đã có form sẵn trong thư mục, dựa vào form đó dựng hồ sơ 144954 cho tôi".
Khuôn (chỉ đọc, gốc thư mục hồ sơ trên Z = Input\144954):
  - `3. TBKQ - 1.docx` (TBKQ /1 của 124254, cùng khách, nước uống QCVN 06-1) -> /1, /2, /3 (N1 – N3).
  - `3. TBKQ - 4.docx` (TBKQ /3.1 của 096118, nước sinh hoạt QCVN 01-1:2024) -> /4, /5 (N4, N5).
  - `2. BCKQTD.docx` (BCKQ 124254, người dùng đã thêm dòng N4, N5) -> BCKQ.
Giữ nguyên bố cục, thứ tự / tên / đơn vị chỉ tiêu, mức, ghi chú của khuôn; chỉ thay thông tin hồ sơ 144954 (số chứng thư,
ngày, khách, địa chỉ, ngày PYC, nội dung, thời gian, địa điểm, thiết bị, GĐV, vị trí, ký hiệu, người chứng kiến, ảnh, số đo
hiện trường) theo BBLM 16/11/2026 và các quyết định người dùng 17/11/2026 (docstring `build_144954.py`).
CHƯA CÓ KQTN: kết quả + phương pháp phòng thí nghiệm, Đánh giá "………" tô vàng (AGENTS mục 13); ngày ban hành = ngày dựng.
KQTN về: ghép theo tên mẫu + số niêm, chỉ điền KQ + phương pháp (AGENTS mục 13, mẫu `fill_131873.py`).

Chạy:  python form_144954.py [--pdf]
"""
import argparse
import copy
import datetime
import io
import os
import re
import sys

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageOps

WS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, WS)
from fill_131229 import txt                             # noqa: E402

IN = os.path.join(WS, "Input", "144954")
CODE = "Q326D144954"
PH = "………"
TODAY = datetime.date.today().strftime("%d/%m/%Y")
CLIENT = "Công ty TNHH Điện máy Orro Home"
ADDR = "Số 27, đường số 4C, khu nhà ở Rạch Bà Tánh, xã Bình Hưng, Thành phố Hồ Chí Minh"      # theo BBLM (17/11)
SITE = ("Văn phòng Orro Home, số 21-23 Phạm Hùng (nối dài), khu dân cư Đại Phúc Green Villas, xã Bình Hưng, "
        "Thành phố Hồ Chí Minh")
WITNESS = "Lê Thị Hồng Anh – Nhân viên"
DATE = "16/11/2026"
SAMPLES = {  # n: (vị trí theo BBLM, số đo hiện trường)
    1: ("Máy lọc và cung cấp nước nóng Philips AUT9124/90", {"ph": "8,73", "tds": "9"}),
    2: ("Máy lọc và cung cấp nước nóng, lạnh Philips ADD6996WH/90", {"ph": "7,37", "tds": "6"}),
    3: ("Máy lọc và cung cấp nước nóng Philips AUT7005/9", {"ph": "6,67", "tds": "3"}),
    4: ("Bộ lọc nước tại vòi – Philips AWP3704/97", {"ph": "7,25", "clo": "< 0,1", "mùi": "Không có mùi lạ"}),
    5: ("Bộ lọc nước đầu nguồn – Philips AWP1830", {"ph": "7,27", "clo": "0,67", "mùi": "Không có mùi lạ"}),
}


def hl(r):
    rp = r.find(qn("w:rPr"))
    if rp is None:
        rp = OxmlElement("w:rPr"); r.insert(0, rp)
    if rp.find(qn("w:highlight")) is None:
        h = OxmlElement("w:highlight"); h.set(qn("w:val"), "yellow"); rp.append(h)


def unhl(el):
    for h in list(el.iter(qn("w:highlight"))):
        h.getparent().remove(h)


def para_set(p, text, yellow=False):
    """Ghi lại cả chữ của đoạn p (giữ định dạng run đầu có chữ)."""
    runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
    if not runs:
        runs = p.findall(qn("w:r"))
    base = runs[0] if runs else None
    for r in p.findall(qn("w:r")):
        if r.find(qn("w:drawing")) is None and not r.findall(qn("w:fldChar")):
            p.remove(r)
    r = copy.deepcopy(base) if base is not None else OxmlElement("w:r")
    for ch in list(r):
        if ch.tag != qn("w:rPr"):
            r.remove(ch)
    rp = r.find(qn("w:rPr"))
    if rp is not None:
        for e in rp.findall(qn("w:highlight")) + rp.findall(qn("w:vertAlign")):
            rp.remove(e)
    t = OxmlElement("w:t"); t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    p.append(r)
    unhl(p)
    if yellow:
        hl(r)


def replace_span(p, pattern, new):
    """Thay đoạn khớp pattern trong chữ của p dù bị chia nhiều run; giữ nguyên các run / tab khác."""
    ts = [t for t in p.iter(qn("w:t")) if t.text]
    joined = "".join(t.text for t in ts)
    m = re.search(pattern, joined)
    if not m:
        return False
    pos = 0
    for t in ts:
        a, b = pos, pos + len(t.text)
        pos = b
        lo, hi = max(a, m.start()), min(b, m.end())
        if lo >= hi:
            continue
        rep = new if a <= m.start() < b else ""
        t.text = t.text[:lo - a] + rep + t.text[hi - a:]
    return True


def label_ph(p, label):
    """Đoạn = nhãn (thường) + "………" tô vàng."""
    para_set(p, label)
    r = copy.deepcopy(p.findall(qn("w:r"))[-1])
    r.find(qn("w:t")).text = PH
    hl(r)
    p.append(r)


def cell_set(tc, text, yellow=False):
    ps = tc.findall(qn("w:p"))
    for p in ps[1:]:
        tc.remove(p)
    lines = text.split("\n")
    para_set(ps[0], lines[0], yellow)
    prev = ps[0]
    for ln in lines[1:]:
        p = copy.deepcopy(ps[0]); para_set(p, ln, yellow); prev.addnext(p); prev = p


def sub_para(p, old, new):
    import unicodedata
    t = unicodedata.normalize("NFC", txt(p))                 # form có chữ tổ hợp (NFD)
    old = unicodedata.normalize("NFC", old)
    if old in t:
        para_set(p, t.replace(old, new))
        return True
    return False


def photo_bytes(path, width_px=1200):
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    if im.height > im.width:                               # ảnh dọc -> xoay ngang (AGENTS mục 3)
        im = im.rotate(90, expand=True)
    im.thumbnail((width_px, width_px))
    b = io.BytesIO(); im.save(b, "JPEG", quality=85)
    return b.getvalue(), im.width, im.height


def replace_images(d, paths):
    """Thay ảnh theo thứ tự xuất hiện; giữ chiều rộng, chiều cao theo tỉ lệ ảnh mới, bỏ crop."""
    blips = list(d.element.body.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"))
    for blip, path in zip(blips, paths):
        rid = blip.get(qn("r:embed"))
        data, w, h = photo_bytes(path)
        part = d.part.related_parts[rid]
        part._blob = data
        inline = blip
        while inline is not None and inline.tag.split("}")[1] not in ("inline", "anchor"):
            inline = inline.getparent()
        ext = inline.find("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
        cx = int(ext.get("cx")); cy = int(cx * h / w)
        ext.set("cy", str(cy))
        for x in inline.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}ext"):
            if x.get("cx"):
                x.set("cx", str(cx)); x.set("cy", str(cy))
        for sr in inline.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect"):
            sr.attrib.clear()


def fix_header(d, n_label):
    for s in d.sections:
        for hp in (s.header, s.first_page_header):
            for p in hp.paragraphs:
                t = txt(p._p)
                if "Số:" in t and re.search(r"Q3\d{2}[A-D]\d{6}", t):
                    t2 = re.sub(r"Q3\d{2}[A-D]\d{6}(/[\d.]+)?", n_label, t)
                    t2 = re.sub(r"\d{2}/\d{2}/\d{4}", TODAY, t2)
                    para_set(p._p, t2)


def add_tds_row(rt):
    """Thêm hàng TDS phòng thí nghiệm (SMEWW 2540C theo PYC thử nghiệm) sau Selen, đánh lại STT; hàng mới tô vàng."""
    trs = rt._tbl.findall(qn("w:tr"))
    sel = next(tr for tr in trs if txt(tr.findall(qn("w:tc"))[1]).strip().startswith("Selen"))
    new = copy.deepcopy(sel)
    for vm in list(new.iter(qn("w:vMerge"))):
        vm.getparent().remove(vm)
    tcs = new.findall(qn("w:tc"))
    for tc, v in zip(tcs, ("", "Tổng chất rắn hòa tan (TDS)", "mg/L", PH, PH, "Không quy định")):
        cell_set(tc, v, yellow=True)
    sel.addnext(new)
    for i, tr in enumerate(rt._tbl.findall(qn("w:tr"))[1:], 1):
        tc0 = tr.findall(qn("w:tc"))[0]
        cell_set(tc0, str(i), yellow=(tr is new))


def add_assessment(d):
    """/4–/5: khuôn 096118 không có Đánh giá -> thêm (người dùng 20/11: "thêm đánh giá trước đi, tô vàng lên"),
    chép khối Đánh giá của khuôn /1 rồi viết lại cho QCVN 01-1:2024/BYT, tô vàng."""
    src = docx.Document(os.path.join(IN, "3. TBKQ - 1.docx")).element.body
    ps = list(src.iter(qn("w:p")))
    i = next(k for k, p in enumerate(ps) if txt(p).strip().startswith("Đánh giá"))
    block = [copy.deepcopy(p) for p in ps[i:i + 4]]
    texts = ["Đánh giá:", "So sánh kết quả thử nghiệm các thông số của mẫu nước cấp nêu trên với ngưỡng giới hạn "
             "cho phép tại QCVN 01-1:2024/BYT:", "+ Thông số không đạt: ", "+ Thông số đạt: "]
    note = [t for t in d.tables if t.rows[0].cells[0].text.strip().startswith("Ghi")][-1]._tbl
    prev = note
    for p, t in zip(block, texts):
        para_set(p, t, yellow=True)
        if t.startswith("+"):
            label_ph(p, t)
            for r in p.findall(qn("w:r")):
                hl(r)
        prev.addnext(p); prev = p


def tbkq(n):
    form = "3. TBKQ - 1.docx" if n <= 3 else "3. TBKQ - 4.docx"
    nu = n <= 3
    d = docx.Document(os.path.join(IN, form))
    vi_tri, field = SAMPLES[n]
    fix_header(d, f"{CODE}/{n}")
    for row in d.tables[0].rows:
        lab = row.cells[0].text.strip()
        val = row.cells[-1]._tc
        v = txt(val).strip()
        if lab.startswith("Đơn vị yêu cầu"):
            cell_set(val, CLIENT)
        elif lab.startswith("Địa chỉ"):
            cell_set(val, ADDR)
        elif "Giấy yêu cầu" in lab:
            cell_set(val, "12/11/2026")
            for pp in row._tr.iter(qn("w:p")):                 # mã PYC ở ô giữa
                for old_code in re.findall(r"Q3\d{2}[A-D]\d{6}", txt(pp)):
                    sub_para(pp, old_code, CODE)
        elif lab.startswith("Nội dung"):
            cell_set(val, "Lấy mẫu và giám định chất lượng " + ("nước uống" if nu else "nước cấp"))
        elif lab.startswith("Thời gian"):
            cell_set(val, f"Từ {DATE}")
        elif lab.startswith("Địa điểm"):
            cell_set(val, SITE)
        if "DR900" in txt(row._tr):                       # máy đo clo theo BBLM: Hach DR300 (người dùng 17/11)
            for p in row._tr.iter(qn("w:p")):
                sub_para(p, "DR900", "DR300")
        if "Nguyễn Khánh Toàn" in txt(row._tr):           # người thực hiện Minh -> GĐV Nguyễn Xuân Nam
            for p in row._tr.iter(qn("w:p")):
                sub_para(p, "Nguyễn Khánh Toàn", "Nguyễn Xuân Nam")
    for p in list(d.element.body.iter(qn("w:p"))):
        sub_para(p, "Công ty TNHH Jones Lang LaSalle (Việt Nam)", CLIENT)
    st = d.tables[1]._tbl.findall(qn("w:tr"))[1].findall(qn("w:tc"))
    cell_set(st[0], vi_tri)
    cell_set(st[1], f"N{n}")
    if nu:                                                  # người dùng 20/11: TDS PTN (PYC thử nghiệm) -> 22 chỉ tiêu
        for p in st[2].iter(qn("w:p")):
            sub_para(p, "21 chỉ tiêu hóa lý", "22 chỉ tiêu hóa lý")
    wt = next(t for t in d.tables if t.rows[0].cells[0].text.strip().startswith("Người chứng kiến"))
    wtc = wt._tbl.findall(qn("w:tr"))[1].findall(qn("w:tc"))
    for tc, v in zip(wtc, (WITNESS, f"N{n}", DATE, "-")):
        cell_set(tc, v)
    rt = next(t for t in d.tables if t.rows[0].cells[0].text.strip() in ("Stt", "STT"))
    group = ""
    for tr in rt._tbl.findall(qn("w:tr"))[1:]:
        tcs = tr.findall(qn("w:tc"))
        name = txt(tcs[1]).strip().lower()
        m = txt(tcs[3]).strip()
        if m:
            group = m
        if "hiện trường" in group:                         # số đo / cảm quan hiện trường theo BBLM
            key = "ph" if name == "ph" else ("tds" if "tds" in name else ("clo" if "clo" in name else
                                                                           ("mùi" if "mùi" in name else None)))
            if key in field:
                cell_set(tcs[4], field[key])
            continue
        if m:
            cell_set(tcs[3], PH, yellow=True)              # phương pháp phòng thí nghiệm: chờ KQTN
        cell_set(tcs[4], PH, yellow=True)
    if not nu:                                              # người dùng 20/11: "xóa dòng TDS đi" -> không thêm hàng TDS
        add_assessment(d)
    for p in d.paragraphs:
        t = p.text
        if nu and (t.strip().startswith("+ Chỉ tiêu") or t.strip().startswith("+ Thông số")):
            label_ph(p._p, t.split(":")[0] + ": ")
    replace_images(d, [os.path.join(IN, "0. vi tri", f"N{n}.jpg")])
    out = os.path.join(IN, "TBKQ", f"TBKQ - viet - {n} - {'nuoc uong' if nu else 'nuoc cap'} - {CODE}.docx")
    d.save(out)
    print(f"[OK] /{n} N{n} <- {form}")
    return out


def bckq():
    d = docx.Document(os.path.join(IN, "2. BCKQTD.docx"))
    for row in d.tables[0].rows:
        lab = row.cells[0].text.strip()
        val = row.cells[-1]._tc
        if any("Số:" in c.text for c in row.cells):
            for pp in row._tr.iter(qn("w:p")):
                for old_code in re.findall(r"Q3\d{2}[A-D]\d{6}", txt(pp)):
                    sub_para(pp, old_code, CODE)
        elif lab.startswith("Địa chỉ"):
            cell_set(val, ADDR)
        elif lab.startswith("Nội dung"):
            cell_set(val, "Lấy mẫu, giám định chất lượng nước uống, nước cấp")
        elif lab.startswith("Thời gian"):
            cell_set(val, "Ngày 16/11/2026 từ 14:15 đến 15:30")
        elif lab.startswith("Người thực hiện"):
            cell_set(val, "Minh")
        if "DR900" in txt(row._tr):
            for p in row._tr.iter(qn("w:p")):
                sub_para(p, "DR900", "DR300")
    trs = d.tables[1]._tbl.findall(qn("w:tr"))
    for n, tr in zip(range(1, 6), trs[1:6]):
        tc = tr.findall(qn("w:tc"))[0]
        cell_set(tc, SAMPLES[n][0])
        for b in list(tc.iter(qn("w:b"))) + list(tc.iter(qn("w:bCs"))):   # dòng N4, N5 người dùng thêm: chữ đậm
            b.getparent().remove(b)
    replace_images(d, [os.path.join(IN, "0. vi tri", f"N{n}.jpg") for n in range(1, 6)])
    from docx.shared import Cm
    for sh in d.inline_shapes:                               # 4,3 cm (khung 3 cột 4–5 cm) để BCKQ còn 2 trang
        f = Cm(4.3) / sh.width
        sh.width, sh.height = Cm(4.3), int(sh.height * f)
    a_b = [("MTTN26022136", "30/11/2026"), ("MTTN26022134", "24/11/2026")]   # a: N1–N3, b: N4–N5 (PYC thử nghiệm)
    k = -1
    import unicodedata
    for p in d.element.body.iter(qn("w:p")):
        t = unicodedata.normalize("NFC", txt(p))             # form có chữ tổ hợp (NFD)
        if "21 chỉ tiêu hóa lý" in t:
            sub_para(p, "21 chỉ tiêu hóa lý", "22 chỉ tiêu hóa lý")
        if t.strip().startswith("So sánh kết quả thử nghiệm"):    # người dùng 20/11: nhắc cả QCVN 01-1, tô vàng
            para_set(p, "So sánh kết quả thử nghiệm các chỉ tiêu của mẫu nước uống N1, N2, N3 (trừ pH, TDS) với mức "
                        "quy định tương ứng tại QCVN 06-1:2010/BYT, phụ lục II và III; mẫu nước cấp N4, N5 với ngưỡng "
                        "giới hạn cho phép tại QCVN 01-1:2024/BYT:", yellow=True)
        if "Kết quả phân tích mẫu N1, N2, N3" in t:
            para_set(p, "- Kết quả phân tích mẫu N1, N2, N3: xem tại phiếu KQTN số MTTN26022136.01-03; mẫu N4, N5: "
                        "xem tại phiếu KQTN số MTTN26022134.01-02")
        elif re.match(r"^\s*[ab]\. Phiếu KQTN số", t):
            k += 1
            label_ph(p, t.split("Phiếu")[0] + f"Phiếu KQTN số {a_b[k][0]} ngày ")
        elif t.strip().startswith("Ngày gửi mẫu"):
            para_set(p, f"Ngày gửi mẫu: {DATE}")                # người dùng 20/11: theo ngày trên BBLM
        elif t.strip().startswith("Ngày yêu cầu cấp KQTN") and k >= 0:
            para_set(p, f"Ngày yêu cầu cấp KQTN: {a_b[k][1]}")
        elif t.strip().startswith("Ngày nhận KQTN"):
            label_ph(p, "Ngày nhận KQTN: ")
        elif t.strip().startswith(("Ngày kết thúc hồ sơ thực tế", "Ngày báo cáo")):
            replace_span(p, r"\d{1,2}/\d{1,2}/\d{4}", TODAY)   # chỉ thay ngày (chữ chia nhiều run), giữ tab
    et = next(t for t in d.tables if t.rows[0].cells[0].text.strip() == "Ký hiệu mẫu" and len(t.rows[0].cells) == 3)
    for tr in et._tbl.findall(qn("w:tr"))[1:]:
        for tc in tr.findall(qn("w:tc")):
            cell_set(tc, PH, yellow=True)
    out = os.path.join(IN, "ThucHien", f"BCKQ - nuoc - {CODE}.docx")
    d.save(out)
    print("[OK] BCKQ <- 2. BCKQTD.docx")
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    docs = [tbkq(n) for n in range(1, 6)] + [bckq()]
    if a.pdf:
        import build_144954 as B
        from nuoc_common import make_dossier
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(make_dossier(dict(B.CONFIG, cert_date=TODAY), None, None, False)).export_files(docs)
    from Core.Tools.qa_check import check_single_docx
    for p in docs:
        r = check_single_docx(p, code=CODE)
        print(r.get("status"), os.path.basename(p))
        for i in r.get("issues", []):
            if "Placeholder" not in i and "Yellow" not in i:
                print("    -", i)
