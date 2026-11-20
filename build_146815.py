# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D146815 - Công ty Cổ phần Mỹ phẩm Sài Gòn — phòng sạch ISO 8, chỉ đo nồng độ hạt bụi, 1 chứng thư (mã trần).

Nguồn (Input/146815, chép từ Z "…\\2026\\Q326D146815 - Mỹ phẩm Sài Gòn - PS" 17/11/2026 — chỉ đọc):
  - 0. BBLM.pdf (1 trang, BM3030.02) ngày 17/11/2026: PYC Q326D146815 ngày 16/11/2026; Công Ty Cổ Phần Mỹ Phẩm Sài Gòn,
      930 Nguyễn Thị Định, Khu công nghiệp Cát Lái (cụm II), Phường Cát Lái, Thành phố Hồ Chí Minh (địa điểm cùng địa chỉ);
      "Đo nồng độ hạt bụi trong khu vực phòng sạch"; 9h15 – 9h50 ngày 17/11/2026; TSI 9350; N2.TC.PS.01;
      "Khu làm khô chai thủy tinh, khu để chai thủy tinh, Airlock" – nồng độ hạt bụi – 15; ghi tay "Tại thời điểm đo bụi
      nhà máy đang hoạt động"; người thực hiện Nguyễn Hoàng Việt.
  - 0. Ghi chép.pdf (4 trang, BM3030.01): tr.1 thông tin chung (PYC ngày 17/11/2026); tr.2 mặt bằng đánh số + bảng file;
      tr.3 bảng file tiếp; tr.4 trắng. AeroTrak 9350, 1 phút. Khu làm khô chai thủy tinh: zero 1333, 1–6 = 1334–1339;
      Khu để chai thủy tinh: zero 1322, 1–6 = 1323–1328 (1 dòng gạch bỏ); Airlock: zero 1329, 1–3 = 1330–1332.
      Mặt bằng ghi "Khu để chai sạch".
  - 0. Phiếu kiểm soát.pdf (10 trang): tr.1 phiếu kiểm soát công việc; tr.3, 5 PYC (2 bản) — MST 0300767984,
      chứng thư Việt – Anh, theo báo giá QONV26017679; tr.7–9 báo giá 10/11/2026: ISO 8; Khu làm khô chai thủy tinh
      10,78 m² 06 vị trí; Khu để chai sạch 16,67 m² 06 vị trí; Airlock 5,40 m² 03 vị trí; tr.2, 4, 6, 10 trắng.
  - 0. KQ bui - Q326D146815.xlsx (sheet 6ChannelFixed, Σ, # / m³) + 6ChannelFixed.pdf: AeroTrak 9350-02
      S/N 93502239001, file 1322 – 1339 ngày 17/11/2026 (đồng hồ máy 09:01 – 09:26).
  - 0. Layout.png: sơ đồ người dùng vẽ (số vị trí theo từng khu) -> dùng nguyên, không chú thích (AGENTS mục 8C).
  - 0. Photo/*.jpg: 7 ảnh (không có thư mục Use -> dùng cả 7).
  - 1. / 2. / 3. … Q326D120423 (đợt 1, khách khác): bản người dùng làm tay, chỉ tham khảo hình thức — không dùng số liệu.

Quyết định của người dùng (18/11/2026):
  - Tên khu thứ 2: "khu để chai sạch" (theo báo giá, mặt bằng; BBLM / ghi chép ghi "Khu để chai thủy tinh").
  - Ngày PYC 16/11/2026; trạng thái "Hoạt động"; giờ đo "theo BBLM" (09:15 – 09:50).
  - "oke" các điểm (4), (5), (7) dưới đây.

Suy luận của AI (người dùng đã đồng ý 18/11/2026):
  (1) [đã chốt: Khu để chai sạch]
  (2) Ngày PYC 16/11/2026 theo BBLM và PYC (phiếu ghi chép ghi 17/11/2026).
  (3) Trạng thái phòng: hoạt động (ghi chú tay BBLM); giới hạn ISO 8 (ISO 14644-1:2015) như báo giá.
  (4) Diện tích theo báo giá; tên EN khách theo tên quốc tế MST "Saigon Cosmetics Corporation"; tên khu EN do AI dịch.
  (5) Nội dung yêu cầu theo BBLM "Đo nồng độ hạt bụi trong khu vực phòng sạch".
  (6) Thời gian theo BBLM 09:15 – 09:50 (đồng hồ máy đếm hạt 09:01 – 09:26).
  (7) Sơ đồ người dùng vẽ chữ tiếng Việt -> dùng cho cả TBKQ EN.

Không có KQTN -> ngày ban hành = ngày dựng (AGENTS mục 2).

Chạy:  python build_146815.py          (DOCX + XLSX + Edoc)
       python build_146815.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ Z trong code — chép bằng `python Core/Tools/ai_files.py chep 146815 --ghi` (AGENTS mục 0).
"""

import os
import sys
import argparse
from datetime import date

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_dong_bo_Z"))

from docx import Document

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, verdict, style_note_tables
from Core.Tools.qa_check import check_single_docx
from Core.Tools.cleanroom_calc import CleanroomCalcBook, read_aerotrak
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)


# --- helper hậu xử lý DOCX (chép từ build_127215.py — cùng form phòng sạch) ---
import copy
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _optimized(path: str, max_w: int = 1200) -> str:
    """Ảnh nén trong <thư mục ảnh>/_optimized (AGENTS mục 1, 7): xoay EXIF, ảnh dọc xoay ngang, không crop."""
    from PIL import Image, ImageOps
    out_dir = os.path.join(os.path.dirname(path), "_optimized")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".jpg")
    im = ImageOps.exif_transpose(Image.open(path))
    if im.height > im.width:
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=88)
    return out


def _replace_in_paragraph(p_el, old: str, new: str) -> bool:
    """Thay chuỗi trong một đoạn (kể cả khi trải qua nhiều run): chữ mới nằm ở run chứa ký tự đầu."""
    ts = [t for t in p_el.iter(qn("w:t"))]
    full = "".join(t.text or "" for t in ts)
    i = full.find(old)
    if i < 0:
        return False
    j = i + len(old)
    pos = 0
    done = False
    for t in ts:
        s, e = pos, pos + len(t.text or "")
        pos = e
        if e <= i or s >= j:
            continue
        txt = t.text or ""
        a, b = max(i, s) - s, min(j, e) - s
        t.text = txt[:a] + ("" if done else new) + txt[b:]
        t.set(qn("xml:space"), "preserve")
        done = True
    return True


def _replace_in_doc(doc, old, new) -> int:
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        while _replace_in_paragraph(p, old, new):
            n += 1
    return n


def _fix_cover_lists(doc):
    """Thiết bị / Căn cứ (Table 0): bỏ numPr (tránh '- - '), thụt 0 để mọi dòng '- ' thẳng hàng (AGENTS mục 14)."""
    for row in doc.tables[0].rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if not p.text.lstrip().startswith("- "):
                    continue
                pPr = p._p.get_or_add_pPr()
                for el in pPr.findall(qn("w:numPr")) + pPr.findall(qn("w:ind")):
                    pPr.remove(el)
                ind = OxmlElement("w:ind")
                ind.set(qn("w:left"), "0")
                ind.set(qn("w:hanging"), "0")
                pPr.append(ind)


def _keep_headings_with_tables(doc):
    """Tiêu đề mục ngay trước bảng kết quả -> keepNext; tiêu đề cột + vài dòng đầu đi liền (AGENTS mục 3)."""
    for el in doc.element.body.iterchildren(qn("w:tbl")):
        prev = el.getprevious()
        if prev is not None and prev.tag == qn("w:p") and "".join(t.text or "" for t in prev.iter(qn("w:t"))).strip():
            pPr = prev.get_or_add_pPr()
            if pPr.find(qn("w:keepNext")) is None:
                pPr.insert(0, OxmlElement("w:keepNext"))
            rows = el.findall(qn("w:tr"))
            for tr in rows[:min(5, len(rows) - 1)]:
                for p_el in tr.iter(qn("w:p")):
                    ppr = p_el.get_or_add_pPr()
                    if ppr.find(qn("w:keepNext")) is None:
                        ppr.insert(0, OxmlElement("w:keepNext"))


def _highlight_placeholders(doc):
    """Tô vàng riêng phần '………' (thân văn bản, bảng, header/footer)."""
    roots = [doc.element.body]
    for s in doc.sections:
        roots += [s.header._element, s.footer._element, s.first_page_header._element]
    for root in roots:
        for r in list(root.iter(qn("w:r"))):
            t = r.find(qn("w:t"))
            if t is None or PLACEHOLDER not in (t.text or "") or len(r.findall(qn("w:t"))) != 1:
                continue
            parts = t.text.split(PLACEHOLDER)
            pieces = []
            for i, part in enumerate(parts):
                if part:
                    pieces.append((part, False))
                if i < len(parts) - 1:
                    pieces.append((PLACEHOLDER, True))
            parent, idx = r.getparent(), r.getparent().index(r)
            extras = [c for c in r if c.tag not in (qn("w:rPr"), qn("w:t"))]
            if extras:
                lead = copy.deepcopy(r)
                for c in list(lead):
                    if c.tag == qn("w:t"):
                        lead.remove(c)
                parent.insert(idx + 1, lead)
                idx += 1
            for k, (txt, hl) in enumerate(pieces):
                nr = copy.deepcopy(r)
                for c in list(nr):
                    if c.tag not in (qn("w:rPr"), qn("w:t")):
                        nr.remove(c)
                nt = nr.find(qn("w:t"))
                nt.text = txt
                nt.set(qn("xml:space"), "preserve")
                if hl:
                    rpr = nr.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr")
                        nr.insert(0, rpr)
                    h = OxmlElement("w:highlight")
                    h.set(qn("w:val"), "yellow")
                    rpr.append(h)
                parent.insert(idx + 1 + k, nr)
            parent.remove(r)


def _keep_result_blocks(doc):
    """Bảng kết quả ngắn (≤ 10 hàng) + đoạn trống + bảng Ghi chú + 'Đánh giá:' + câu kết luận cùng trang;
    câu kết luận chỉ keepLines, không keepNext (AGENTS mục 3, 15)."""
    def kn(el, on=True):
        for p_el in ([el] if el.tag == qn("w:p") else list(el.iter(qn("w:p")))):
            ppr = p_el.get_or_add_pPr()
            for x in ppr.findall(qn("w:keepNext")):
                ppr.remove(x)
            if on:
                ppr.insert(0, OxmlElement("w:keepNext"))

    body = doc.element.body
    tbls = list(body.iterchildren(qn("w:tbl")))[1:]          # bỏ Table 0 (bìa)
    for tbl in tbls:
        nxt = tbl.getnext()
        gaps = []
        while nxt is not None and nxt.tag == qn("w:p") and not "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip():
            gaps.append(nxt)
            nxt = nxt.getnext()
        if nxt is None or nxt.tag != qn("w:tbl"):
            continue
        first = "".join(t.text or "" for t in nxt.iter(qn("w:t"))).strip()
        if not first.startswith(("Ghi chú", "Note")):
            continue
        rows = tbl.findall(qn("w:tr"))
        for tr in (rows if len(rows) <= 12 else rows[-2:]):
            kn(tr)
        for g in gaps:
            kn(g)
        head = nxt.getnext()
        if head is None or head.tag != qn("w:p"):
            continue
        if "".join(t.text or "" for t in head.iter(qn("w:t"))).strip() not in ("Đánh giá:", "Assessment:"):
            kn(nxt, False)      # Ghi chú không có Đánh giá theo sau -> không keepNext (E32)
            continue
        kn(nxt)
        kn(head)
        sent = head.getnext()
        if sent is not None and sent.tag == qn("w:p"):
            kn(sent, False)
            ppr = sent.get_or_add_pPr()
            if ppr.find(qn("w:keepLines")) is None:
                ppr.insert(0, OxmlElement("w:keepLines"))


CODE = "Q326D146815"
CODE6 = "146815"
REG_DATE = "16/11/2026"
SAMPLE_DATE = "17/11/2026"
START, END = "09:15", "09:50"
CERT_DATE = date.today().strftime("%d/%m/%Y")
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
RAW_XLSX = os.path.join(INPUT_DIR, "0. KQ bui - Q326D146815.xlsx")
PHOTO_DIR = os.path.join(INPUT_DIR, "0. Photo")
LAYOUT = os.path.join(INPUT_DIR, "0. Layout.png")

CLIENT = ClientInfo(
    name_vn="Công ty Cổ phần Mỹ phẩm Sài Gòn",
    name_en="Saigon Cosmetics Corporation",
    address_vn="930 Nguyễn Thị Định, Khu công nghiệp Cát Lái (cụm II), phường Cát Lái, Thành phố Hồ Chí Minh",
    address_en="930 Nguyen Thi Dinh Street, Cat Lai Industrial Park (Cluster II), Cat Lai Ward, Ho Chi Minh City",
)
LOCATION_VN = f"Phòng sạch của {CLIENT.name_vn} - {CLIENT.address_vn}"
LOCATION_EN = f"Cleanroom of {CLIENT.name_en} - {CLIENT.address_en}"

ISO_CLASS = 8
ISO8_LIMITS = (3520000, 832000, 29300)          # ISO 14644-1:2015 Bảng 1, d ≥ 0,5 / 1 / 5 µm, hạt/m³
TRANG_SO_DO = "2"                               # trang phụ lục chứa sơ đồ — kiểm tra lại trên PDF

ROOMS = [
    {"vn": "Khu làm khô chai thủy tinh", "en": "Glass bottle drying area", "area": "10,78", "zero": 1333,
     "dust": list(zip(range(1, 7), range(1334, 1340)))},
    {"vn": "Khu để chai sạch", "en": "Clean bottle storage area", "area": "16,67", "zero": 1322,
     "dust": list(zip(range(1, 7), range(1323, 1329)))},
    {"vn": "Airlock", "en": "Airlock", "area": "5,40", "zero": 1329,
     "dust": list(zip(range(1, 4), range(1330, 1333)))},
]

LABELS, RECS = read_aerotrak(RAW_XLSX)
_need = [f for r in ROOMS for f in [r["zero"]] + [x[1] for x in r["dust"]]]
if [f for f in _need if f not in RECS]:
    raise SystemExit(f"[FATAL] file đếm hạt thiếu {[f for f in _need if f not in RECS]}")
for r in ROOMS:
    for _, f in r["dust"]:
        if str(RECS[f].get("Date/Time (dd/mm/yyyy)", ""))[:10] != SAMPLE_DATE:
            raise SystemExit(f"[FATAL] file {f} không đo ngày {SAMPLE_DATE}")
    if any(RECS[r["zero"]][k] for k in ("Ch1", "Ch2", "Ch3", "Ch6")):
        print(f"[WARNING] file zero {r['zero']} khác 0")


def dust_vals(f):
    return [RECS[f]["Ch2"], RECS[f]["Ch3"], RECS[f]["Ch6"]]


def dust_ok():
    return all(v <= lim for r in ROOMS for _, f in r["dust"] for v, lim in zip(dust_vals(f), ISO8_LIMITS))


def build_dossier() -> Dossier:
    photos = [{"path": _optimized(os.path.join(PHOTO_DIR, f)), "caption_vn": "", "caption_en": ""}
              for f in sorted(os.listdir(PHOTO_DIR)) if f.lower().endswith(".jpg")]
    layouts = [{"path": LAYOUT, "caption_vn": "", "caption_en": ""}]   # sơ đồ người dùng vẽ -> không chú thích
    locations = []
    for r in ROOMS:
        n = len(r["dust"])
        locations.append(MonitoringLocation(
            name_vn=r["vn"], name_en=r["en"], area_m2=r["area"], clean_class=f"ISO {ISO_CLASS}",
            state_vn="Hoạt động", state_en="In operation",
            points_summary_vn=[f"Nồng độ hạt bụi: {n:02d} vị trí (1 – {n})."],
            points_summary_en=[f"Particle concentration: {n:02d} locations (1 – {n})."],
        ))
    ok = dust_ok()
    bui = [{"cap": f"ISO {ISO_CLASS}", "gioi_han": list(ISO8_LIMITS), "trang_so_do": TRANG_SO_DO,
            "diem": [{"khu_vuc": r["vn"], "khu_vuc_en": r["en"], "vi_tri": n, "gia_tri": dust_vals(f)}
                     for r in ROOMS for n, f in r["dust"]],
            "ket_luan_vn": verdict(ok, False), "ket_luan_en": verdict(ok, True)}]
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False,
        dossier_type="cleanroom",
        scope_vn="Đo nồng độ hạt bụi trong khu vực phòng sạch",
        scope_en="Measurement of particle concentration in cleanroom area",
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01",
        procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01",
        standards_vn="ISO 14644-1:2015",
        standards_en="ISO 14644-1:2015",
        client=CLIENT,
        personnel=PersonnelInfo(executor="Việt"),     # BBLM: Nguyễn Hoàng Việt -> GĐV ký Nguyễn Xuân Nam
        equipments=[eq("aerotrak_9350")],
        locations=locations,
        photos=photos, layouts=layouts,
        results_data={"bui_3kenh": bui, "layout_cols": 1, "layout_width_cm": 12.0},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


def finish_tbkq(src: str, is_en: bool) -> str:
    doc = Document(src)
    _fix_cover_lists(doc)
    _replace_in_doc(doc, " (3 particle sizes)" if is_en else " (3 kênh hạt)", "")
    for t in doc.tables:                     # bảng khu vực: hàng ép cao -> 560 atLeast (AGENTS 15 [Row Too Tall])
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if ("Cấp sạch" in hdr or "Class" in hdr) and len(t.columns) >= 3:
            for tr in t._tbl.findall(qn("w:tr"))[1:]:
                for h in tr.iter(qn("w:trHeight")):
                    if int(h.get(qn("w:val"), "0")) > 760:
                        h.set(qn("w:val"), "560")
    _keep_headings_with_tables(doc)
    _keep_result_blocks(doc)
    _keep_room_groups(doc)
    style_note_tables(doc)      # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3)
    from Core.Builders.tag_context import appendix_total_field
    appendix_total_field(doc)   # footer phụ lục {= {NUMPAGES} - 1} (AGENTS mục 4)
    _highlight_placeholders(doc)
    lang = "en" if is_en else "viet"
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {lang} - phong sach - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    return dst


SO_BUI = f"{CODE}/Bụi"


def _th(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["tham_chieu"] = [
        f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {SO_BUI} ngày {SAMPLE_DATE}."]
    ok = dust_ok()
    ctx["danh_gia"] = [{
        "khu_vuc": "; ".join(f"“{r['vn']}”" for r in ROOMS),
        "dat": "Nồng độ hạt bụi" if ok else "",
        "khong_dat": "Không" if ok else "Nồng độ hạt bụi",
        "can_cu": (f"ISO 14644-1:2015 (cấp sạch ISO {ISO_CLASS}: d ≥ 0,5 µm ≤ {_th(ISO8_LIMITS[0])} hạt/m³; "
                   f"d ≥ 1 µm ≤ {_th(ISO8_LIMITS[1])} hạt/m³; d ≥ 5 µm ≤ {_th(ISO8_LIMITS[2])} hạt/m³)"),
    }]
    out = b.build(ctx=ctx, out_name=f"BCKQ - 1 - phong sach - {CODE}.docx")
    doc = Document(out)
    _keep_headings_with_tables(doc)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Hình ảnh", "Sơ đồ", "Báo cáo của GĐV")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    return out


def build_excel() -> str:
    book = CleanroomCalcBook()
    groups = [{"room": r["vn"], "cap": f"ISO {ISO_CLASS}",
               "limits": {"Ch2": ISO8_LIMITS[0], "Ch3": ISO8_LIMITS[1], "Ch6": ISO8_LIMITS[2]},
               "zero": RECS[r["zero"]], "points": [RECS[f] for _, f in r["dust"]],
               "labels": [n for n, _ in r["dust"]]} for r in ROOMS]
    book.dust_aerotrak(groups, so=SO_BUI, date=SAMPLE_DATE, labels=LABELS)
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - TinhToanPhongSach.xlsx")
    book.save(out)
    for n in book.notes:
        print(f"[build_146815] {n}")
    print(f"[build_146815] Excel -> {out}")
    return out


def _keep_room_groups(doc):
    """Bảng kết quả có cột khu vực gộp dọc: các dòng cùng một khu vực đi liền nhau (chép từ build_120423.py)."""
    for tbl in list(doc.element.body.iterchildren(qn("w:tbl")))[1:]:
        rows = tbl.findall(qn("w:tr"))

        def cont(tr):
            tcs = tr.findall(qn("w:tc"))
            vm = tcs[0].find(f"{qn('w:tcPr')}/{qn('w:vMerge')}") if tcs else None
            return vm is not None and vm.get(qn("w:val")) in (None, "continue")
        size, group, k = {}, {}, None
        for i, tr in enumerate(rows):
            k = k if cont(tr) and k is not None else i
            size[k] = size.get(k, 0) + 1
            group[i] = k
        for i, (tr, nxt) in enumerate(zip(rows[2:], rows[3:]), 2):          # bỏ 2 dòng tiêu đề
            if cont(nxt) and size[group[i]] <= 8:
                for p_el in tr.iter(qn("w:p")):
                    ppr = p_el.get_or_add_pPr()
                    if ppr.find(qn("w:keepNext")) is None:
                        ppr.insert(0, OxmlElement("w:keepNext"))


def split_edoc() -> list:
    """BBLM = '0. BBLM.pdf'; Ghi chép = '0. Ghi chép.pdf' tr.1–3 + phiếu in '6ChannelFixed.pdf';
    Khác = '0. Phiếu kiểm soát.pdf' (PKSCV, PYC, báo giá). Trang trắng tự bỏ (finish_edoc, AGENTS mục 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    outs = []
    for name, srcs in ((f"{CODE6} - BBLM.pdf", ["0. BBLM.pdf"]),
                       (f"{CODE6} - Ghi Chep.pdf", ["0. Ghi chép.pdf", "6ChannelFixed.pdf"]),
                       (f"{CODE6} - Khac.pdf", ["0. Phiếu kiểm soát.pdf"])):
        dst = fitz.open()
        for s in srcs:
            dst.insert_pdf(fitz.open(os.path.join(INPUT_DIR, s)))
        p = os.path.join(edoc, name)
        finish_edoc(dst, INPUT_DIR, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_146815] Edoc -> {p} ({dst.page_count} trang)")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()
    d = build_dossier()
    d.ensure_output_directories()
    split_edoc()
    h = CleanroomHandler(d)
    vn = finish_tbkq(h.build_tbkq(is_en=False), is_en=False)
    en = finish_tbkq(h.build_tbkq(is_en=True), is_en=True)
    bckq = build_bckq(d)
    xlsx = build_excel()
    docs = [bckq, vn, en]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(d)
        ex.export_files(docs)
        ex.export_excel_files([xlsx])
    print("\n===== QA =====")
    print(f"  Bụi: {'phù hợp' if dust_ok() else 'KHÔNG phù hợp'} ISO {ISO_CLASS}")
    for p in docs:
        print(f"  {check_single_docx(p, code=d.code).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
