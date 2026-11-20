# -*- coding: utf-8 -*-
"""
Khắc phục hồ sơ Q326D122266 /2 (khí nén KN1, KN2) — chứng thư thay thế Q326D122266/2/R.

Nguồn (chép từ Z "Q326D122266 - NEXTERN - Phòng sạch" về Input\\122266\\_nguon, chỉ đọc):
  - TBKQ\\TBKQ - vn/en - 2 - khi nen - Q326D122266.docx (bản ban hành 06/11/2026, làm tay / Antigravity;
    PDF ban hành ở TBKQ\\Khắc Phục\\KQ Ban Hanh).
  - Thuc hien\\Tính Toan\\122266 - Bụi KN 0.1.xlsx (database Lasair III 110 S/N 137662).
  - Thuc hien\\Tính Toan\\Khí nén\\Q326D122266- KN1 - TinhToanKhiNen.xlsx (phiếu tính toán KN1, N* = 1,008737).
  - Ảnh giấy in nhanh Lasair người dùng gửi 19/11/2026: "Final Sample Report" #12, 06/10/2026 10:58:19 → 11:33:38,
    Location "122266 - do", Δ (N/m³) 29242 / 30770 / 19743 / 33391 / 19559 / 9940 / 6261 / 3 — khớp dòng #12 database.

Sự không phù hợp:
  /2 ban hành lấy bụi KN1 từ mẫu #42 (07/10/2026 ~11:00 → 11:35:25), không phải mẫu #12 ngày 06/10 đã in giấy giao
  khách (KN1 đo điểm sương 06/10 10:41–11:01; chứng thư ghi "Từ 06/10/2026"). Đổi sang mẫu #12:
  KN1 0,1–0,5 / 0,5–1 / 1–5 / > 5 µm = 133 864 / 10 027 / 6 316 / 3 hạt/m³; d > 0,5 µm 45 742 → 16 346 (đạt < 27 615).

Quyết định của người dùng (19/11/2026):
  - "ok, vậy làm khắc phục giúp t hồ sơ này" — sau khi AI báo KN1 dùng nhầm mẫu 07/10, giấy in giao khách là #12.

Suy luận của AI (cần người dùng xác nhận):
  - KN2 giữ mẫu #23 (06/10 16:15:46) như bản ban hành (người dùng chưa gửi giấy in KN2).
  - Số EN dạng /2/E/R (như 096547, 139329). Ngày ban hành /R = ngày dựng.
  - Không sửa /3 (KN3), BCKQ khí nén (chỉ viện dẫn phiếu tính toán).

Chạy:  python khac_phuc_122266.py [--pdf]
Kết quả: Input\\122266\\TBKQ\\Khắc Phục\\ (TBKQ /2/R VN, EN; phiếu KN1 /R; phiếu Xử lý sự KPH) -> robocopy lên Z cùng tên.
"""
import os
import sys
import copy
import shutil
import argparse
import datetime as _dt

sys.stdout.reconfigure(encoding="utf-8")
WS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, WS)

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

CODE = "Q326D122266"
OLD_DATE = "06/11/2026"
NEW_DATE = _dt.date.today().strftime("%d/%m/%Y")
SRC = os.path.join(WS, "Input", "122266", "_nguon", "TBKQ")
OUT = os.path.join(WS, "Input", "122266", "TBKQ", "Khắc Phục")

KN1 = ("133 864", "10 027", "6 316", "3", "16 346")      # = phiếu KN1 /R (Excel COM)


def _set_run_text(p, text):
    rs = [r for r in p.runs]
    rs[0].text = text
    for r in rs[1:]:
        r.text = ""


def _cell_set(cell, text, bold=None):
    p = cell.paragraphs[0]
    rs = [r for r in p.runs if r.text]
    rs[0].text = text
    for r in rs[1:]:
        r.text = ""
    if bold is not None:
        rs[0].bold = bold


def patch(lang):
    is_en = lang == "en"
    src = os.path.join(SRC, f"TBKQ - {lang} - 2 - khi nen - {CODE}.docx")
    dst = os.path.join(OUT, f"TBKQ - {lang} - 2 - R - khi nen - {CODE}.docx")
    shutil.copy2(src, dst)
    d = docx.Document(dst)
    old = f"{CODE}/2/E" if is_en else f"{CODE}/2"
    new = old + "/R"
    n = 0
    for s in d.sections:
        h = s.header
        if h.is_linked_to_previous:
            continue
        for t in h._element.iter(qn("w:t")):
            if t.text and old in t.text and new not in t.text:
                t.text = t.text.replace(old, new); n += 1
            if t.text and OLD_DATE in t.text:
                t.text = t.text.replace(OLD_DATE, NEW_DATE)
    assert n >= 2, f"{lang}: chỉ đổi {n} số chứng thư"
    # dòng thay thế dưới số ở trang bìa
    hdr0 = d.sections[0].header._element
    num_p = next(p for p in hdr0.iter(qn("w:p")) if new in "".join(x.text or "" for x in p.iter(qn("w:t"))))
    tpl = [r for r in num_p.findall(qn("w:r")) if r.find(qn("w:t")) is not None][0]
    p2 = copy.deepcopy(num_p)
    for r in p2.findall(qn("w:r")):
        p2.remove(r)
    r2 = copy.deepcopy(tpl)
    rpr = r2.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr"); r2.insert(0, rpr)
    for tag in ("w:b", "w:bCs", "w:i", "w:iCs"):
        e = rpr.find(qn(tag))
        if e is not None:
            rpr.remove(e)
    rpr.append(OxmlElement("w:i"))
    r2.find(qn("w:t")).text = (
        f"(This inspection certificate replaces the Inspection certificate No. {CODE}/2/E dated {OLD_DATE})"
        if is_en else
        f"(Chứng thư giám định này thay thế Chứng thư giám định số {CODE}/2 ngày {OLD_DATE})")
    p2.append(r2)
    num_p.addnext(p2)

    # bảng nồng độ hạt (KN1)
    t8 = next(t for t in d.tables if len(t.rows) >= 6 and t.rows[2].cells[0].text.strip() == "KN1"
              and "884" in t.rows[2].cells[1].text)
    cells = list(dict.fromkeys(t8.rows[2].cells))
    for c, v in zip(cells[1:4], KN1[:3]):
        _cell_set(c, v)
    c5 = list(dict.fromkeys(t8.rows[3].cells))[1]
    _cell_set(c5, f"Concentration of particle size d > 5.0 µm: {KN1[3]} particles/m3" if is_en
              else f"Nồng độ hạt bụi có kích thước d > 5,0 µm: {KN1[3]} hạt/m3")
    # bảng d > 0,5 µm
    t10 = next(t for t in d.tables if len(t.rows) == 4 and len(t.columns) >= 2
               and t.rows[1].cells[0].text.strip() == "KN1" and "45" in t.rows[1].cells[1].text)
    _cell_set(t10.rows[1].cells[1], KN1[4], bold=False)          # đạt -> gỡ đậm (mục 3)
    # đánh giá
    p = next(p for p in d.paragraphs if ("particle size d > 0.5" in p.text if is_en
                                         else "nồng độ hạt bụi có kích thước d > 0,5" in p.text))
    runs = [r for r in p.runs if r.text]
    base, phrase, tail = runs[0], runs[1], runs[2]
    for r in p.runs:   # chỉ xóa chữ, không đụng run chứa đường kết thúc (drawing/AlternateContent — AGENTS mục 5)
        if r._r not in (base._r, phrase._r, tail._r):
            for t in r._r.findall(qn("w:t")):
                t.text = ""
    if is_en:
        base.text = ("The measurement result of concentration of particle size d > 0.5 µm in compressed air at "
                     "location KN1 ")
        p1, mid, p2t = "complied", ", at location KN2 ", "did not comply"
    else:
        base.text = "Kết quả đo nồng độ hạt bụi có kích thước d > 0,5 µm trong khí nén tại vị trí KN1 "
        p1, mid, p2t = "phù hợp", ", tại vị trí KN2 ", "không phù hợp"
    phrase.text = p1
    r_mid = copy.deepcopy(base._r); r_mid.find(qn("w:t")).text = mid
    r_ph2 = copy.deepcopy(phrase._r); r_ph2.find(qn("w:t")).text = p2t
    phrase._r.addnext(r_mid); r_mid.addnext(r_ph2)
    for r in (r_mid,):
        r.find(qn("w:t")).set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    # câu mới dài hơn -> dòng cuối chạy tới giữa trang, đường kết thúc (neo ở đoạn này) đè chữ -> hạ thêm 6pt
    import re as _re
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    for pv in p._p.iter(WP + "positionV"):
        off = pv.find(WP + "posOffset")
        if off is not None:
            off.text = str(int(off.text) + 76200)
    for el in p._p.iter():
        st = el.get("style")
        if st and "margin-top:" in st:
            el.set("style", _re.sub(r"margin-top:([-\d.]+)pt",
                                    lambda m: f"margin-top:{float(m.group(1)) + 6:g}pt", st))
    d.save(dst)
    print(f"[122266] {os.path.basename(dst)}: {old} -> {new}, KN1 = {KN1}, ngày {NEW_DATE}")
    return dst


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    os.makedirs(OUT, exist_ok=True)
    outs = [patch("vn"), patch("en")]
    from Core.Tools.qa_check import check_single_docx
    for o in outs:
        print(check_single_docx(o, code=CODE).get("status"), os.path.basename(o))
    if a.pdf:
        from Core.Exporters.pdf_exporter import convert_docx_to_pdf, convert_xlsx_to_pdf
        for o in outs:
            print("PDF", convert_docx_to_pdf(o, os.path.splitext(o)[0] + ".pdf"))
        x = os.path.join(OUT, f"{CODE}- KN1 - TinhToanKhiNen - R.xlsx")
        if os.path.isfile(x):   # chỉ sheet chính như PDF gốc (3 trang), không xuất sheet DiemSuong
            import win32com.client as w
            from Core.Exporters.pdf_exporter import avoid_label_printer
            avoid_label_printer()
            xl = w.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
            try:
                wb = xl.Workbooks.Open(os.path.abspath(x), ReadOnly=True)
                wb.Worksheets("Tinh_Toan_Khi_Co_Ap_V6").ExportAsFixedFormat(0, os.path.abspath(x)[:-5] + ".pdf")
                wb.Close(False)
            finally:
                xl.Quit()
            print("PDF", x[:-5] + ".pdf")


if __name__ == "__main__":
    main()
