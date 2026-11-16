# -*- coding: utf-8 -*-
"""
Build standard Cleanroom (Phòng sạch) form templates from the 2026 sample results
stored in `New From 2026\\PhongSach`.

Learned structure (from samples 123613, 108033, 078038 + current template):
  Table 0 (cover)            -> kept 100% untouched (copied from the newest sample 123613)
  Mục "Nội dung thực hiện"   -> 4 fixed paragraphs (compatible with CleanroomHandler)
  T1 Khu vực / Diện tích / Cấp sạch / Số vị trí, T2 Ghi chú (1)
  T3 Ảnh hiện trường (grid 2 cột), T4 Sơ đồ vị trí
  Library of result blocks: Heading -> Bảng KQ -> Ghi chú -> Đánh giá -> Kết luận

Output: Core\\Templates\\PhongSach\\
    FORM - TBKQ - viet - phong sach - 2026.docx
    FORM - TBKQ - en - phong sach - 2026.docx
    FORM - BCKQ - phong sach - 2026.docx

Usage:  python Core\\Tools\\build_cleanroom_forms.py
"""
import copy
import glob
import os
import sys
import unicodedata

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
from Core.Styling import add_document_ending_line  # noqa: E402

SAMPLE_DIR = os.path.join(ROOT, "New From 2026", "PhongSach")
TPL_DIR = os.path.join(ROOT, "Core", "Templates")
OUT_DIR = os.path.join(TPL_DIR, "PhongSach")

PH = "………"
TOKEN = "{PH}"

RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
             "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
             "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
             "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
             "specVanish", "oMath"]


# --------------------------------------------------------------------------- helpers
def find_file(*parts):
    for f in glob.glob(os.path.join(SAMPLE_DIR, "*.docx")) + glob.glob(os.path.join(TPL_DIR, "*.docx")):
        name = os.path.basename(f)
        if "(2)" in name:
            continue
        if all(p in name for p in parts):
            return f
    raise FileNotFoundError(parts)


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def rpr_set(rpr, tag, attrs=None):
    """Insert/replace a child of rPr respecting the schema order."""
    old = rpr.find(qn("w:" + tag))
    if old is not None:
        rpr.remove(old)
    el = OxmlElement("w:" + tag)
    for k, v in (attrs or {}).items():
        el.set(qn("w:" + k), v)
    idx = RPR_ORDER.index(tag)
    for i, ch in enumerate(list(rpr)):
        local = ch.tag.split("}")[1]
        if local in RPR_ORDER and RPR_ORDER.index(local) > idx:
            rpr.insert(i, el)
            return el
    rpr.append(el)
    return el


def rpr_del(rpr, *tags):
    for t in tags:
        for el in rpr.findall(qn("w:" + t)):
            rpr.remove(el)


def segs(text, base_style=None):
    """'abc {PH} def' -> [('abc ', base), ('………', 'ph'), (' def', base)]."""
    out = []
    parts = text.split(TOKEN)
    for i, part in enumerate(parts):
        if part:
            out.append((part, base_style))
        if i < len(parts) - 1:
            out.append((PH, "ph"))
    return out


def set_par(p, segments, keep_bold=False):
    """Replace all runs of paragraph `p` by `segments` [(text, style)], reusing the first run's rPr.
    style: None | 'b' | 'i' | 'bi' | 'ph' (yellow highlight) | 'sup'."""
    tpl = None
    for r in p.iter(qn("w:r")):
        if r.find(qn("w:t")) is not None:
            tpl = r.find(qn("w:rPr"))
            break
    if tpl is None:
        for r in p.iter(qn("w:r")):
            if r.find(qn("w:rPr")) is not None:
                tpl = r.find(qn("w:rPr"))
                break
    for ch in list(p):
        if ch.tag != qn("w:pPr"):
            p.remove(ch)
    for txt, style in segments:
        r = OxmlElement("w:r")
        rpr = copy.deepcopy(tpl) if tpl is not None else OxmlElement("w:rPr")
        rpr_del(rpr, "highlight", "vertAlign", "rStyle")
        if not keep_bold:
            rpr_del(rpr, "b", "bCs", "i", "iCs")
        if style in ("b", "bi"):
            rpr_set(rpr, "b"); rpr_set(rpr, "bCs")
        if style in ("i", "bi"):
            rpr_set(rpr, "i"); rpr_set(rpr, "iCs")
        if style == "ph":
            rpr_set(rpr, "highlight", {"val": "yellow"})
        if style == "sup":
            rpr_set(rpr, "vertAlign", {"val": "superscript"})
        r.append(rpr)
        lines = txt.split("\n")
        for li, line in enumerate(lines):
            if li:
                r.append(OxmlElement("w:br"))
            t = OxmlElement("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = line
            r.append(t)
        p.append(r)


def set_cell(tc, lines, keep_bold=False):
    """lines: str | list[str|list[segment]]  -> one paragraph per line."""
    if isinstance(lines, str):
        lines = [lines]
    ps = tc.findall(qn("w:p"))
    p0 = ps[0]
    for p in ps[1:]:
        tc.remove(p)
    for li, line in enumerate(lines):
        p = p0 if li == 0 else copy.deepcopy(p0)
        if li:
            tc.append(p)
        set_par(p, segs(line) if isinstance(line, str) else line, keep_bold=keep_bold)


def is_vcont(tc):
    pr = tc.find(qn("w:tcPr"))
    vm = pr.find(qn("w:vMerge")) if pr is not None else None
    return vm is not None and vm.get(qn("w:val")) in (None, "continue")


def fill_row(tr, values, keep_bold=False):
    tcs = [tc for tc in tr.findall(qn("w:tc")) if not is_vcont(tc)]
    for tc, v in zip(tcs, values):
        if v is not None:
            set_cell(tc, v, keep_bold=keep_bold)


def strip_numpr(el):
    for n in list(el.iter(qn("w:numPr"))):
        n.getparent().remove(n)


def strip_drawings(el):
    for tag in ("w:drawing", "w:pict", "w:object"):
        for d in list(el.iter(qn(tag))):
            r = d.getparent()
            while r is not None and r.tag != qn("w:r"):
                r = r.getparent()
            if r is not None and r.getparent() is not None:
                r.getparent().remove(r)


def clear_cell(tc, hint=""):
    ps = tc.findall(qn("w:p"))
    for p in ps[1:]:
        tc.remove(p)
    strip_drawings(ps[0])
    set_par(ps[0], segs(hint) if hint else [])


def ensure_header_rows(tbl, n):
    for tr in tbl.findall(qn("w:tr"))[:n]:
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = OxmlElement("w:trPr")
            tr.insert(0, trpr)
        if trpr.find(qn("w:tblHeader")) is None:
            trpr.append(OxmlElement("w:tblHeader"))


def ensure_row_height(tbl, start, twips=440):
    for tr in tbl.findall(qn("w:tr"))[start:]:
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = OxmlElement("w:trPr")
            tr.insert(0, trpr)
        h = trpr.find(qn("w:trHeight"))
        if h is None:
            h = OxmlElement("w:trHeight")
            trpr.append(h)
        try:
            cur = int(h.get(qn("w:val")) or 0)
        except ValueError:
            cur = 0
        if cur < 340:
            h.set(qn("w:val"), str(twips))
        h.set(qn("w:hRule"), "atLeast")


def trim_table(tbl, header_n, keep_data, has_last):
    trs = tbl.findall(qn("w:tr"))
    last = trs[-1] if has_last else None
    data = trs[header_n:len(trs) - (1 if has_last else 0)]
    for tr in data[keep_data:]:
        tbl.remove(tr)
    trs = tbl.findall(qn("w:tr"))
    return trs[header_n:header_n + keep_data], last


def clone_table(doc, idx):
    t = copy.deepcopy(doc.tables[idx]._tbl)
    strip_numpr(t)
    strip_drawings(t)
    return t


def find_par(doc, startswith, with_num=None):
    for p in doc.element.body.iter(qn("w:p")):
        t = unicodedata.normalize("NFC", text_of(p)).strip()
        if t.startswith(unicodedata.normalize("NFC", startswith)):
            if with_num is None:
                return p
            has = p.find(qn("w:pPr")) is not None and p.find(qn("w:pPr")).find(qn("w:numPr")) is not None
            if has == with_num:
                return p
    raise ValueError(startswith)


# --------------------------------------------------------------------------- language packs
L = {
    "vn": dict(
        base=("TBKQ - 1 - viet - 123613",),
        p3=(("viet - 1 - phong sach - Q326D108033",), 5),
        p6=(("TBKQ - 1 - viet - 123613",), 5),
        thn=(("TBKQ - 1 - viet - 123613",), 7),
        pa=(("TBKQ - 1 - viet - 123613",), 9),
        hepa=(("TBKQ - viet - phong sach.docx",), 10),
        ach=(("TBKQ - viet - phong sach.docx",), 18),
        vs=(("Khu Vi Sinh",), 8),
        scan=(("Scanloc",), 2),
        h_content="Nội dung thực hiện",
        area_by="Khu vực thực hiện theo sự chỉ định của {PH}:",
        area_list="“{PH}”; “{PH}”",
        state="Trạng thái phòng sạch: “{PH}”",
        photos="Hình ảnh quá trình thực hiện:",
        layout="Sơ đồ các vị trí quan trắc:",
        t1_points=["- Nồng độ hạt bụi: {PH} vị trí", "- Nhiệt độ - độ ẩm: {PH} vị trí",
                   "- Chênh lệch áp suất: {PH} vị trí", "- Độ ồn: {PH} vị trí"],
        n1=[("(1) …", "Thông số kỹ thuật phòng sạch do {PH} cung cấp."),
            ("●", "Số lượng vị trí đo nồng độ hạt bụi được xác định theo bảng A1, ISO 14644-1:2015.")],
        accr="Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.",
        n_point="Vị trí đo xem sơ đồ minh họa tại trang {PH}.",
        n_sample="Vị trí lấy mẫu xem sơ đồ minh họa tại trang {PH}.",
        n_limit="Mức quy định được cung cấp bởi {PH}.",
        n_volume="Thể tích phòng được cung cấp bởi đại diện {PH}.",
        n_vs=["Kết quả được biểu thị “< 10 CFU/m³” khi không phát hiện vi sinh mục tiêu trên thể tích không khí được lấy mẫu.",
              "Phương pháp thử nghiệm được công nhận ISO 17025 – VILAS 004."],
        assess_lbl="Đánh giá:",
        ok="phù hợp",
        h_p="Kết quả đo nồng độ hạt bụi",
        h_vs="Kết quả xác định mức độ nhiễm bẩn vi sinh vật",
        h_thn="Kết quả đo nhiệt độ - độ ẩm và độ ồn",
        h_pa="Kết quả đo chênh lệch áp suất giữa các khu vực",
        h_hepa="Kết quả đo độ rò rỉ lọc HEPA",
        h_scan="Kết quả quét rò rỉ màn lọc HEPA",
        scan_layout="Sơ đồ vị trí các lọc HEPA khu vực {PH}:",
        h_ach="Kết quả đo lưu lượng khí cấp qua lọc HEPA",
        lim_p="Giới hạn tối đa cho phép đối với phòng sạch cấp ISO {PH} quy định tại ISO 14644-1:2015",
        lim_vs="Giới hạn mức độ ô nhiễm vi sinh vật đối với cấp sạch {PH} nêu tại WHO TRS 961 (2011) – Phụ lục 6",
        lim_by="Mức quy định cung cấp bởi {PH}",
        vs_hdr=None,
        scan_hdr=["Khu vực đo", "Vị trí đo (2)", "Độ rò rỉ lọc HEPA, %", "Độ rò rỉ gasket, %"],
        a_p=("Kết quả đo nồng độ hạt bụi trong không khí tại “{PH}” ", " với mức quy định đối với khu vực cấp sạch ISO {PH} theo ISO 14644-1:2015."),
        a_thn=("Kết quả đo nhiệt độ, độ ẩm và độ ồn tại “{PH}” ", " với mức quy định tương ứng được cung cấp bởi {PH}."),
        a_pa=("Kết quả đo chênh lệch áp suất giữa các khu vực nêu trên ", " với mức quy định tương ứng được cung cấp bởi {PH}."),
        a_hepa=("Kết quả đo độ rò rỉ lọc HEPA tại “{PH}” ", " với mức quy định tương ứng nêu tại ISO 14644-3:2019."),
        a_ach=("Kết quả tính toán số lần trao đổi khí tại “{PH}” ", " với mức quy định tương ứng được cung cấp bởi {PH}."),
    ),
    "en": dict(
        base=("TBKQ - 1 - en - 123613",),
        p3=(("en - 1 - phong sach - Q326D108033",), 5),
        p6=(("TBKQ - 1 - en - 123613",), 5),
        thn=(("TBKQ - 1 - en - 123613",), 7),
        pa=(("TBKQ - 1 - en - 123613",), 9),
        hepa=(("TBKQ - en - phong sach.docx",), 10),
        ach=(("TBKQ - en - phong sach.docx",), 18),
        vs=(("Khu Vi Sinh",), 8),
        scan=(("Scanloc",), 2),
        h_content="Performance content",
        area_by="The performance location was indicated by {PH}:",
        area_list="“{PH}”; “{PH}”",
        state="Cleanroom occupancy state: “{PH}”",
        photos="Photos of the performance process:",
        layout="Layout of monitoring points:",
        t1_points=["- Particle concentration: {PH} points", "- Temperature - Humidity: {PH} points",
                   "- Pressure differential: {PH} points", "- Noise: {PH} points"],
        n1=[("(1) …", "The cleanroom specifications provided by {PH}."),
            ("●", "Number of particle-concentration measurement points determined according to Table A.1, ISO 14644-1:2015.")],
        accr="The performance field is accredited to ISO 17020 - VIAS 004.",
        n_point="See measuring points in the diagram on page {PH}.",
        n_sample="See sampling points in the diagram on page {PH}.",
        n_limit="Permissible limit provided by {PH}.",
        n_volume="Room volume provided by the representative of {PH}.",
        n_vs=["The result is expressed as “< 10 CFU/m³” when the target microorganism is not detected in the sampled air volume.",
              "The test method is accredited to ISO 17025 – VILAS 004."],
        assess_lbl="Assessment:",
        ok="complied",
        h_p="Measurement result of particle concentration",
        h_vs="Determination result of microbial contamination level",
        h_thn="Measurement result of temperature – humidity and noise",
        h_pa="Measurement result of pressure differential between areas",
        h_hepa="Measurement result of HEPA filter leakage",
        h_scan="Measurement result of HEPA filter scan leak test",
        scan_layout="Layout of HEPA filter locations in {PH}:",
        h_ach="Measurement result of airflow rate supply through HEPA filter",
        lim_p="The maximum permissible limit for ISO Class {PH} cleanroom specified in ISO 14644-1:2015",
        lim_vs="Limit of microbial contamination for Grade {PH} specified in WHO TRS 961 (2011) – Annex 6",
        lim_by="Permissible limit provided by {PH}",
        vs_hdr=[["Sampling area", "Sampling point (2)", [[("Airborne colony count, CFU/m", None), ("3", "sup")]]],
                ["Total aerobic microbial count"]],
        scan_hdr=["Measurement room", "Measurement point (2)", "HEPA filter leakage, %", "Gasket leakage, %"],
        a_p=("The measurement result of airborne particle concentration in “{PH}” ", " with the specified limit for ISO Class {PH} cleanroom according to ISO 14644-1:2015."),
        a_thn=("The measurement results of temperature, humidity and noise in “{PH}” ", " with the corresponding permissible limit provided by {PH}."),
        a_pa=("The measurement results of pressure differential between areas mentioned above ", " with the corresponding permissible limit provided by {PH}."),
        a_hepa=("The measurement result of HEPA filter leakage in “{PH}” ", " with the corresponding permissible limit specified in ISO 14644-3:2019."),
        a_ach=("The calculation result of air changes per hour in “{PH}” ", " with the corresponding permissible limit provided by {PH}."),
    ),
}

AREA = "“{PH}”"


# --------------------------------------------------------------------------- TBKQ builder
def build_tbkq(lang):
    cfg = L[lang]
    base_path = find_file(*cfg["base"])
    doc = docx.Document(base_path)
    body = doc.element.body
    cache = {}

    def src(key):
        parts, idx = cfg[key]
        f = find_file(*parts)
        if f not in cache:
            cache[f] = docx.Document(f)
        return clone_table(cache[f], idx)

    # ---- prototypes taken from the base document (keeps exact 2026 formatting)
    tbls = doc.tables
    t1, t2, t3, t4 = tbls[1]._tbl, tbls[2]._tbl, tbls[3]._tbl, tbls[4]._tbl
    p_head = copy.deepcopy(find_par(doc, cfg["h_p"], with_num=True))
    p_assess_lbl = copy.deepcopy(find_par(doc, cfg["assess_lbl"]))
    note_proto = copy.deepcopy(tbls[6]._tbl)  # 3-row note (Ghi chú / (2) … / ●)
    p_layout_lbl = copy.deepcopy(find_par(doc, cfg["layout"]))
    # assessment sentence prototype (first one after the particle heading)
    ok_word = cfg["ok"]
    p_assess = None
    for p in body.iter(qn("w:p")):
        t = text_of(p)
        if ok_word in t and len(t) > 40:
            p_assess = copy.deepcopy(p)
            break
    strip_drawings(p_assess)
    # empty spacer
    p_empty = copy.deepcopy(p_head)
    pPr = p_empty.find(qn("w:pPr"))
    if pPr is not None:
        strip_numpr(pPr)
        st = pPr.find(qn("w:pStyle"))
        if st is not None:
            pPr.remove(st)
    set_par(p_empty, [])

    # ---- 1) intro paragraphs: Nội dung / Khu vực ... : / “list” / Trạng thái
    p_content = find_par(doc, cfg["h_content"])
    p_area = p_content.getnext()
    while p_area.tag != qn("w:p") or not text_of(p_area).strip():
        p_area = p_area.getnext()
    p_state = find_par(doc, cfg["state"].split(":")[0])
    set_par(p_content, [(cfg["h_content"], None)], keep_bold=True)
    set_par(p_area, segs(cfg["area_by"]))
    p_list = copy.deepcopy(p_area)
    set_par(p_list, segs(cfg["area_list"]))
    p_area.addnext(p_list)
    set_par(p_state, segs(cfg["state"]))

    # ---- 2) T1 area table: header + 1 placeholder row
    rows, _ = trim_table(t1, 1, 1, False)
    fill_row(rows[0], [AREA, [[(PH, "ph"), (" m", None), ("2", "sup")]], "ISO {PH}", cfg["t1_points"]])
    ensure_header_rows(t1, 1)
    ensure_row_height(t1, 1)

    # ---- 3) T2 note (1)
    rebuild_note(t2, note_proto, cfg["n1"])

    # ---- 4) photos (T3) & layout (T4): empty grids
    for tr in t3.findall(qn("w:tr"))[1:]:
        t3.remove(tr)
    for tc in t3.iter(qn("w:tc")):
        clear_cell(tc)
    for tc in t4.iter(qn("w:tc")):
        clear_cell(tc)

    # ---- 5) remove all old result blocks after T4 (keep bookmarkEnd + sectPr)
    started = False
    for el in list(body):
        if el is t4:
            started = True
            continue
        if started and el.tag in (qn("w:p"), qn("w:tbl")):
            body.remove(el)
    sectPr = body.find(qn("w:sectPr"))

    def add(el):
        sectPr.addprevious(el)

    def heading(text):
        h = copy.deepcopy(p_head)
        set_par(h, segs(text), keep_bold=True)
        add(h)

    def spacer():
        add(copy.deepcopy(p_empty))

    def note(entries):
        t = copy.deepcopy(note_proto)
        rebuild_note(t, note_proto, entries)
        add(t)

    def assessment(pair):
        lbl = copy.deepcopy(p_assess_lbl)
        add(lbl)
        a = copy.deepcopy(p_assess)
        set_par(a, segs(pair[0]) + [(ok_word, "bi")] + segs(pair[1]))
        add(a)

    def result_table(tbl, header_n, data_vals, last_vals=None, header_texts=None):
        has_last = last_vals is not None
        rows, last = trim_table(tbl, header_n, len(data_vals), has_last)
        for tr, vals in zip(rows, data_vals):
            fill_row(tr, vals)
        if has_last and last is not None:
            fill_row(last, last_vals)
        if header_texts:
            for tr, vals in zip(tbl.findall(qn("w:tr"))[:header_n], header_texts):
                fill_row(tr, vals, keep_bold=True)
        ensure_header_rows(tbl, header_n)
        ensure_row_height(tbl, header_n)
        add(tbl)

    pt = (cfg["n_point"])
    accr = ("●", cfg["accr"])

    # Block 1 — particle concentration, 3 particle sizes (≥0.5 / ≥1 / ≥5 µm)
    heading(cfg["h_p"])
    result_table(src("p3"), 2, [[AREA, "1", PH_, PH_, PH_], ["2", PH_, PH_, PH_]],
                 [cfg["lim_p"], PH_, PH_, PH_])
    spacer(); note([("(2) …", pt), accr]); assessment(cfg["a_p"]); spacer()

    # Block 2 — particle concentration, 6 particle sizes (≥0.1 ... ≥5 µm)
    heading(cfg["h_p"])
    result_table(src("p6"), 2, [[AREA, "1"] + [PH_] * 6, ["2"] + [PH_] * 6],
                 [cfg["lim_p"]] + [PH_] * 6)
    spacer(); note([("(2) …", pt), accr]); assessment(cfg["a_p"]); spacer()

    # Block 3 — microbial contamination (VS)
    heading(cfg["h_vs"])
    result_table(src("vs"), 2, [[AREA, "VS1", PH_], ["VS2", PH_]],
                 [cfg["lim_vs"], "≤ {PH}"], header_texts=cfg["vs_hdr"])
    spacer(); note([("(2) …", cfg["n_sample"])] + [("●", x) for x in cfg["n_vs"]] + [accr]); spacer()

    # Block 4 — temperature / humidity / noise
    heading(cfg["h_thn"])
    result_table(src("thn"), 1, [[AREA, PH_, PH_, PH_]], [None, PH_, PH_, PH_])
    spacer(); note([("(3) …", cfg["n_limit"]), accr]); assessment(cfg["a_thn"]); spacer()

    # Block 5 — pressure differential
    heading(cfg["h_pa"])
    result_table(src("pa"), 1, [[AREA, None, AREA, PH_, PH_]])
    spacer(); note([("(3) …", cfg["n_limit"]), accr]); assessment(cfg["a_pa"]); spacer()

    # Block 6 — HEPA filter leakage (aerosol photometer)
    heading(cfg["h_hepa"])
    result_table(src("hepa"), 1, [[AREA, "1", PH_], ["2", PH_]], [None, None])
    spacer(); note([("(2) …", pt), accr]); assessment(cfg["a_hepa"]); spacer()

    # Block 7 — HEPA filter scan leak test (filter + gasket)
    heading(cfg["h_scan"])
    lbl = copy.deepcopy(p_layout_lbl)
    set_par(lbl, segs(cfg["scan_layout"]))
    add(lbl)
    lay = copy.deepcopy(t4)
    add(lay)
    spacer()
    result_table(src("scan"), 1, [[AREA, "1", PH_, PH_], ["2", PH_, PH_]],
                 header_texts=[cfg["scan_hdr"]] if cfg["scan_hdr"] else None)
    spacer(); note([("(2) …", pt), accr]); spacer()

    # Block 8 — airflow rate & air changes per hour
    heading(cfg["h_ach"])
    result_table(src("ach"), 1, [[AREA, "1", PH_, PH_, PH_, PH_], ["2", PH_]],
                 [cfg["lim_by"], "≥ {PH}"])
    spacer(); note([("(2) …", pt), ("(3) …", cfg["n_volume"]), accr]); assessment(cfg["a_ach"])

    add_document_ending_line(doc)
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, f"FORM - TBKQ - {'viet' if lang == 'vn' else 'en'} - phong sach - 2026.docx")
    doc.save(out)
    return out


PH_ = "{PH}"


def rebuild_note(tbl, proto, entries):
    """Rebuild a 3-row-style note table: row0 'Ghi chú:' + one row per entry (label, text)."""
    p_trs = proto.findall(qn("w:tr"))
    ref_row, bul_row = p_trs[1], p_trs[-1]
    trs = tbl.findall(qn("w:tr"))
    # single-row legacy note -> convert to the 2026 3-row layout
    head = copy.deepcopy(p_trs[0]) if len(trs) < 2 else trs[0]
    for tr in trs:
        tbl.remove(tr)
    tbl.append(head)
    for label, text in entries:
        row = copy.deepcopy(bul_row if label == "●" else ref_row)
        fill_row(row, [label, text])
        tbl.append(row)


# --------------------------------------------------------------------------- BCKQ builder
def build_bckq():
    doc = docx.Document(find_file("BCKQ - 1 - phong sach"))
    body = doc.element.body
    tbls = doc.tables
    t1, t2, t3, t4, t5 = (tbls[i]._tbl for i in range(1, 6))

    p_area = find_par(doc, "Khu vực thực hiện theo")
    p_list = p_area.getnext()
    p_state = find_par(doc, "Trạng thái phòng sạch")
    set_par(p_area, segs("Khu vực thực hiện theo sự chỉ định của {PH}:"))
    set_par(p_list, segs("“{PH}”; “{PH}”"))
    set_par(p_state, segs("Trạng thái phòng sạch: “{PH}”"))

    rows, _ = trim_table(t1, 1, 1, False)
    fill_row(rows[0], [AREA, [[(PH, "ph"), (" m", None), ("2", "sup")]], "ISO {PH}",
                       ["- Đo nồng độ hạt bụi: {PH} vị trí", "- Đo nhiệt độ - độ ẩm: {PH} vị trí",
                        "- Đo chênh lệch áp suất: {PH} vị trí"]])
    ensure_header_rows(t1, 1)
    ensure_row_height(t1, 1)

    tc_note = t2.findall(qn("w:tr"))[0].findall(qn("w:tc"))[-1]
    set_cell(tc_note, ["(1) … Thông số kỹ thuật phòng sạch do đại diện {PH} cung cấp.",
                       "Số lượng vị trí đo nồng độ hạt bụi được xác định theo bảng A1, ISO 14644-1:2015."])

    for tr in t3.findall(qn("w:tr"))[1:]:
        t3.remove(tr)
    for tc in t3.iter(qn("w:tc")):
        clear_cell(tc)
    for tc in t4.iter(qn("w:tc")):
        clear_cell(tc)
    set_par(find_par(doc, "Sơ đồ vị trí đo"), segs("Sơ đồ các vị trí đo:"))

    # result reference lines
    p_ref = find_par(doc, "- Kết quả đo nồng độ hạt bụi")
    refs = [
        "- Kết quả đo nồng độ hạt bụi: xem tại phiếu Kết quả đo nồng độ hạt bụi số Q326D{PH}/Bụi ngày {PH}",
        "- Kết quả xác định mức độ nhiễm bẩn vi sinh vật: xem tại phiếu kết quả thử nghiệm số {PH} ngày {PH}",
        "- Kết quả đo nhiệt độ - độ ẩm, độ ồn, chênh lệch áp suất, độ rò rỉ lọc HEPA, lưu lượng gió: xem tại TBKQ số Q326D{PH}",
    ]
    set_par(p_ref, segs(refs[0]))
    anchor = p_ref
    for line in refs[1:]:
        np_ = copy.deepcopy(p_ref)
        set_par(np_, segs(line))
        anchor.addnext(np_)
        anchor = np_

    # assessment table
    rows, _ = trim_table(t5, 1, 1, False)
    fill_row(rows[0], [AREA, ["Nồng độ hạt bụi", "Nhiệt độ - độ ẩm", "Chênh lệch áp suất"], "-",
                       ["Mức quy định nồng độ hạt bụi tối đa đối với phòng sạch cấp ISO {PH} nêu tại ISO 14644-1:2015",
                        "Mức quy định được cung cấp bởi {PH}"]])
    for tc in rows[0].findall(qn("w:tc")):
        pr = tc.find(qn("w:tcPr"))
        if pr is not None:
            for vm in pr.findall(qn("w:vMerge")):
                pr.remove(vm)
    ensure_header_rows(t5, 1)
    ensure_row_height(t5, 1)

    # GĐV report
    for start, txt in [("Ngày kết thúc hồ sơ thực tế", "Ngày kết thúc hồ sơ thực tế: {PH}"),
                       ("Thời gian trễ hạn", "Thời gian trễ hạn: Không"),
                       ("Nguyên nhân trễ hạn", "Nguyên nhân trễ hạn: Không")]:
        set_par(find_par(doc, start), segs(txt))
    for start in ("Ngày báo cáo", "Người báo cáo"):
        # keep the sample's leading tabs (single-line layout); only the value becomes a placeholder
        p = find_par(doc, start)
        ts = list(p.iter(qn("w:t")))
        k = next(i for i, t in enumerate(ts) if ":" in (t.text or ""))
        ts[k].text = ts[k].text.split(":")[0] + ": "
        ts[k].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for t in ts[k + 1:]:
            r = t.getparent()
            r.getparent().remove(r)
        r_ph = copy.deepcopy(ts[k].getparent())
        for ch in list(r_ph):
            if ch.tag != qn("w:rPr"):
                r_ph.remove(ch)
        rpr = r_ph.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r_ph.insert(0, rpr)
        rpr_set(rpr, "highlight", {"val": "yellow"})
        t_ph = OxmlElement("w:t")
        t_ph.text = PH
        r_ph.append(t_ph)
        ts[k].getparent().addnext(r_ph)

    out = os.path.join(OUT_DIR, "FORM - BCKQ - phong sach - 2026.docx")
    os.makedirs(OUT_DIR, exist_ok=True)
    doc.save(out)
    return out


if __name__ == "__main__":
    for lang in ("vn", "en"):
        print("[OK]", build_tbkq(lang))
    print("[OK]", build_bckq())
