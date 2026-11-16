"""
QA & Final Verification Tool for QUATEST 3 (Room N2).
Replaces final-check.ps1, test-open.ps1, and rename-done.ps1.
Checks:
- Leftover old strings (-Old), foreign Q3xx codes, malformed 7-digit codes (Bẫy 5)
- NFC-normalized Vietnamese search across document.xml + all header*.xml + footer*.xml (Bẫy 3, 10, 36)
- Remaining TAG {{ }} / {% %} from tagged templates (Core/Templates/TAGS.md)
- Remaining '………' / '...' placeholders and <w:highlight> yellow highlights (Bẫy 23, 52, 56)
- Duplicate wp:docPr IDs (Bẫy 34) & missing [Content_Types].xml image extensions (Bẫy 6)
- Leftover Word comments (Bẫy 50)
- Accreditation sentence (ISO 17020 - VIAS 004) in the note table under every TBKQ result-data table
- keepNext / keepLines / cantSplit on 'Đánh giá:' / 'Assessment:' blocks (QUY-TAC Phần 5)
- Word COM test-open + page count verification (Bẫy 11, 24, 27, 36)
"""
import os
import re
import sys
import time
import zipfile
import argparse
import unicodedata
from lxml import etree

sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Cau cong nhan bat buoc trong bang 'Ghi chú:' / 'Note:' duoi moi bang ket qua TBKQ (nguoi dung chot 30/10/2026)
ACCREDIT_VN = "Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004"
ACCREDIT_EN = "The performance field is accredited to ISO 17020 - VIAS 004"


def _note_norm(note_text: str) -> str:
    return re.sub(r"\s+", " ", note_text.lower()).replace("–", "-")   # VIAS viết gạch '–' cũng nhận


_BYT_WATER_RE = re.compile(r"nước uống|nước sinh hoạt|nước sạch|nước cấp|drinking water|domestic water|clean water")
_ENV_WATER_RE = re.compile(r"nước|water|bùn|sludge|chất thải|waste")


def is_vimcerts(table0_text: str) -> bool:
    """Câu công nhận theo QUY CHUẨN (người dùng chốt 16/11/2026, AGENTS mục 3, 8C): mẫu nước dùng QCVN …/BYT
    (nước uống, nước sinh hoạt, nước sạch, nước cấp) -> VIAS 004; các mẫu nước / bùn / chất thải còn lại
    (QCVN …/BTNMT, …/BNNMT, mức KH) -> VIMCERTS 078; mọi loại hồ sơ khác (tủ, phòng sạch, khí nén…) -> VIAS 004.
    ``table0_text``: chữ của bảng bìa (Table 0) — xét dòng 'Nội dung yêu cầu' và 'Căn cứ'."""
    t = to_nfc(table0_text).lower()
    if "/byt" in t.replace(" ", ""):
        return False
    if _BYT_WATER_RE.search(t):
        return False
    m = re.search(r"(nội dung yêu cầu giám định|inspection request content)(.{0,200})", t, re.S)
    scope = m.group(2) if m else ""
    return bool(_ENV_WATER_RE.search(scope))


VIMCERTS_KEY = "vimcerts 078"


# Bang ket qua = dong tieu de co cot ket qua / vi tri do (khong tinh bang thong tin khu vuc, bang anh/so do)
RESULT_TABLE_RE = re.compile(r"kết quả|result|vị trí đo|measuring point|measurement point")

def to_nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s or "")


def _effective_size(run, para, default_pt):
    """Cỡ chữ hiệu lực của run: rPr của run -> style ký tự -> style đoạn (theo chuỗi base) -> docDefaults."""
    if run.font.size is not None:
        return run.font.size.pt
    for st in (run.style if run.style is not None and run.style.name != "Default Paragraph Font" else None, para.style):
        while st is not None:
            if st.font is not None and st.font.size is not None:
                return st.font.size.pt
            st = st.base_style
    return default_pt


def check_note_label(docx_path: str) -> list:
    """[Note Label] — nhãn 'Ghi chú' / 'Note' của bảng Ghi chú phải gạch chân đủ cả nhãn (AGENTS mục 3)."""
    from docx import Document as _Doc
    from docx.oxml.ns import qn as _qn
    out = []
    doc = _Doc(docx_path)
    for i, tbl in enumerate(doc.element.body.iter(_qn("w:tbl"))):
        for p in tbl.iter(_qn("w:p")):
            txt = "".join(t.text or "" for t in p.iter(_qn("w:t")))
            lab = next((lb for lb in ("Ghi chú", "Note") if txt.lstrip().startswith(lb + ":")
                        or txt.lstrip().startswith(lb + " :")), None)
            if not lab:
                continue
            flags, done = [], False
            for r in p.iter(_qn("w:r")):
                u = r.find(_qn("w:rPr") + "/" + _qn("w:u"))
                ul = u is not None and u.get(_qn("w:val")) not in ("none", None)
                for ch in "".join(t.text or "" for t in r.iter(_qn("w:t"))):
                    if ch == ":":
                        done = True
                        break
                    if ch.strip():
                        flags.append(ul)
                if done:
                    break
            if any(flags) and not all(flags):   # gạch một phần (form không gạch nhãn thì bỏ qua)
                out.append(f"[Note Label] bảng Ghi chú #{i}: nhãn '{lab}' gạch chân chưa đủ")
            break
    return out


def check_table0_align(docx_path: str, tol_twips: int = 20) -> list:
    """[Table0 Align] — hàng 'nhãn | : | giá trị' của Table 0 TBKQ: dòng đầu ô giá trị phải cùng một cột
    (thụt trái − treo + dòng đầu lệch nhau quá `tol_twips` là lỗi; AGENTS mục 3)."""
    from docx import Document as _Doc
    doc = _Doc(docx_path)
    if not doc.tables:
        return []
    offs = {}
    for row in doc.tables[0].rows:
        cells, seen = [], set()
        for c in row.cells:
            if c._tc not in seen:
                seen.add(c._tc)
                cells.append(c)
        if len(cells) != 3 or cells[1].text.strip() != ":" or not cells[2].text.strip():
            continue
        ind = cells[2].paragraphs[0]._p.find(f".//{{{W_NS}}}ind")
        g = (lambda k: int(ind.get(f"{{{W_NS}}}{k}", 0) or 0)) if ind is not None else (lambda k: 0)
        offs[to_nfc(cells[0].text.strip())[:30]] = g("left") - g("hanging") + g("firstLine")
    if offs and max(offs.values()) - min(offs.values()) > tol_twips:
        base = max(set(offs.values()), key=list(offs.values()).count)
        odd = [f"'{k}' {v:+d}" for k, v in offs.items() if v != base]
        return [f"[Table0 Align] Table 0: ô giá trị không thẳng cột ({', '.join(odd)} twips so với {base:+d}) "
                f"— dựng lại (tag_context.align_table0_values)"]
    return []


def check_layout_rules(docx_path: str, is_bckq: bool) -> tuple[list, list]:
    """Quy tắc rút ra từ hồ sơ 132809 (05/11/2026, AGENTS mục 3, 4, 7):
    - [Font Size] TBKQ: đoạn văn, Table 0, bảng kết quả/thông tin 13pt; bảng 'Ghi chú:'/'Note:' 11pt in nghiêng (09/11/2026). BCKQ: mọi chữ 13pt.
    - [Row Too Tall] (cảnh báo) hàng dữ liệu 1 dòng cao > 760 twips (1,34 cm) trong bảng kết quả.
    - [Note Gap] (cảnh báo) đoạn trống giữa bảng kết quả và bảng Ghi chú không thu về <= 2pt exact (Ghi chú phải sát bảng)."""
    import docx
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    issues, warnings = [], []
    doc = docx.Document(docx_path)
    dd = doc.styles.element.find(f"{{{W_NS}}}docDefaults")
    sz = dd.find(f".//{{{W_NS}}}rPrDefault//{{{W_NS}}}sz") if dd is not None else None
    default_pt = int(sz.get(f"{{{W_NS}}}val")) / 2 if sz is not None else 10.0

    def is_note_tbl(el):
        return el is not None and el.tag == f"{{{W_NS}}}tbl" and \
            to_nfc("".join(el.itertext())).strip().startswith(("Ghi chú", "Note"))

    bad = {}
    def scan(paras, where, want):
        for p in paras:
            for r in p.runs:
                if not r.text.strip():
                    continue
                s = _effective_size(r, p, default_pt)
                if s != want:
                    bad.setdefault((where, s, want), to_nfc(p.text.strip())[:50])

    t_idx = 0
    for el in doc.element.body.iterchildren():
        if el.tag == f"{{{W_NS}}}p":
            scan([Paragraph(el, doc)], "đoạn văn", 13)
        elif el.tag == f"{{{W_NS}}}tbl":
            note = is_note_tbl(el)
            want = 13 if (is_bckq or not note) else 11
            tbl = Table(el, doc)
            seen = set()
            for row in tbl.rows:
                for c in row.cells:
                    if id(c._tc) not in seen:
                        seen.add(id(c._tc))
                        scan(c.paragraphs, f"bảng #{t_idx}" + (" (Ghi chú)" if note else ""), want)
            if not is_bckq and not note and t_idx > 0 and el.find(f".//{{{W_NS}}}drawing") is None:
                for r_i, tr in enumerate(el.findall(f"{{{W_NS}}}tr")[1:], start=1):
                    th = tr.find(f"{{{W_NS}}}trPr/{{{W_NS}}}trHeight")
                    if th is None or int(th.get(f"{{{W_NS}}}val", "0")) <= 760:
                        continue
                    one_line = all(len(tc.findall(f".//{{{W_NS}}}p")) <= 1 and tc.find(f".//{{{W_NS}}}br") is None
                                   and len(to_nfc("".join(tc.itertext()))) <= 25 for tc in tr.findall(f"{{{W_NS}}}tc"))
                    if one_line and "".join(tr.itertext()).strip():
                        warnings.append(f"[Row Too Tall] bảng #{t_idx} hàng #{r_i}: hàng 1 dòng cao {th.get(f'{{{W_NS}}}val')} twips "
                                        f"(> 760) — hạ về ~440 (AGENTS mục 3)")
            if note and not is_bckq:
                # Bảng Ghi chú in nghiêng toàn bộ (người dùng chốt 09/11/2026)
                not_it = [to_nfc("".join(t.text or "" for t in r.iter(f"{{{W_NS}}}t")))
                          for r in el.iter(f"{{{W_NS}}}r")
                          if "".join(t.text or "" for t in r.iter(f"{{{W_NS}}}t")).strip()
                          and r.find(f"{{{W_NS}}}rPr/{{{W_NS}}}i") is None]
                if not_it:
                    issues.append(f"[Note Style] bảng Ghi chú #{t_idx}: chữ chưa in nghiêng — vd '{not_it[0][:40]}' "
                                  f"(AGENTS mục 3: 11pt in nghiêng)")
                gap = el.getprevious()
                if gap is not None and gap.tag == f"{{{W_NS}}}p" and not "".join(gap.itertext()).strip() \
                        and gap.getprevious() is not None and gap.getprevious().tag == f"{{{W_NS}}}tbl":
                    sp = gap.find(f"{{{W_NS}}}pPr/{{{W_NS}}}spacing")
                    ok = sp is not None and sp.get(f"{{{W_NS}}}lineRule") == "exact" and \
                        int(sp.get(f"{{{W_NS}}}line", "999")) <= 40
                    if not ok:
                        warnings.append(f"[Note Gap] bảng Ghi chú #{t_idx} cách bảng kết quả một đoạn trống cao bình thường "
                                        f"— thu về 1pt exact để Ghi chú nằm sát bảng (AGENTS mục 3)")
            t_idx += 1
    for (where, s, want), sample in bad.items():
        issues.append(f"[Font Size] {where}: chữ {s:g}pt (phải {want}pt) — vd '{sample}'")
    return issues, warnings


def check_layout_stale(docx_path: str) -> list:
    """[Layout Stale] — sơ đồ draw.io (<hồ sơ>/TBKQ/_diagrams/<tên>.drawio) chưa vào DOCX (AGENTS 8C):
      (a) .drawio lưu sau PNG xuất ra (người dùng kéo lại điểm, chưa chạy lại script), hoặc
      (b) PNG hiện tại (<tên>_vn.png cho TBKQ VN/BCKQ, <tên>_en.png cho TBKQ EN) không có trong word/media của DOCX
          (đã xuất lại sơ đồ nhưng chưa dựng lại DOCX). DOCX ở TBKQ\\ hoặc ThucHien\\."""
    import glob
    import hashlib
    d = os.path.dirname(os.path.abspath(docx_path))
    if os.path.basename(d).lower() not in ("tbkq", "thuchien"):
        return []
    lang = "en" if re.search(r"(^|[\s-])en([\s-]|$)", os.path.basename(docx_path).lower()) else "vn"
    fmt = lambda t: time.strftime("%d/%m/%Y %H:%M", time.localtime(t))      # noqa: E731
    media = None
    out = []
    for f in glob.glob(os.path.join(os.path.dirname(d), "TBKQ", "_diagrams", "*.drawio")):
        png = f"{os.path.splitext(f)[0]}_{lang}.png"
        name = os.path.basename(f)
        if not os.path.isfile(png):
            continue
        if os.path.getmtime(f) > os.path.getmtime(png) + 2:
            out.append(f"[Layout Stale] {name} sửa lúc {fmt(os.path.getmtime(f))}, sau lần xuất "
                       f"{os.path.basename(png)} ({fmt(os.path.getmtime(png))}) — chạy lại script hồ sơ.")
            continue
        if media is None:
            with zipfile.ZipFile(docx_path) as zf:
                media = {hashlib.sha256(zf.read(n)).hexdigest() for n in zf.namelist() if n.startswith("word/media/")}
        if hashlib.sha256(open(png, "rb").read()).hexdigest() not in media:
            out.append(f"[Layout Stale] DOCX chưa chứa {os.path.basename(png)} hiện tại (xuất {fmt(os.path.getmtime(png))}) "
                       f"— dựng lại DOCX (chạy script hồ sơ không kèm --chi-so-do).")
    return out


def check_single_docx(docx_path: str, code: str = None, old_terms: list[str] = None) -> dict:
    issues = []
    warnings = []
    highlights = []
    dots = []

    # Auto-infer dossier code from filename if not explicitly provided
    if not code:
        m_code = re.search(r"(Q3\d{2}[A-Z]\d{5,6})", os.path.basename(docx_path))
        if m_code:
            code = m_code.group(1)

    code_6 = re.search(r"(\d{6})$", code).group(1) if code and re.search(r"(\d{6})$", code) else None
    base_lower = os.path.basename(docx_path).lower()
    is_bckq = "bckq" in base_lower
    is_tem = "tem" in base_lower

    with zipfile.ZipFile(docx_path, "r") as zf:
        names = zf.namelist()

        # 1. Check [Content_Types].xml vs word/media/* (Bẫy 6)
        ct_root = etree.fromstring(zf.read("[Content_Types].xml"))
        declared_exts = {
            (el.get("Extension") or "").lower()
            for el in ct_root.findall(f"{{{CT_NS}}}Default")
        }
        for n in names:
            if n.startswith("word/media/"):
                ext = os.path.splitext(n)[1].lstrip(".").lower()
                if ext and ext not in declared_exts:
                    issues.append(f"[Bẫy 6 - Content_Types]: Media '{n}' extension '{ext}' not in [Content_Types].xml")

        # 1b. Check Header Certificate Numbers consistency across headers
        header_certs = []
        for h_ent in sorted(n for n in names if n.startswith("word/header") and n.endswith(".xml")):
            h_root = etree.fromstring(zf.read(h_ent))
            h_txt = to_nfc("".join(h_root.itertext()))
            found_certs = re.findall(r"(?:Số|No)\s*:\s*([A-Za-z0-9/]+)", h_txt)
            for fc in found_certs:
                header_certs.append((h_ent, fc))
                base = fc.split("/")[0]
                if code and base != code:
                    issues.append(f"[Header Code Mismatch] in {h_ent}: cert '{fc}' base '{base}' != expected '{code}'")

        distinct_certs = sorted(set(c for _, c in header_certs))
        if len(distinct_certs) > 1:
            issues.append(f"[Header Inconsistency]: Conflicting certificate numbers across headers: {distinct_certs}")

        # 2. Check document.xml + all header*.xml + footer*.xml
        xml_entries = ["word/document.xml"]
        xml_entries += sorted(n for n in names if n.startswith("word/header") and n.endswith(".xml"))
        xml_entries += sorted(n for n in names if n.startswith("word/footer") and n.endswith(".xml"))

        for ent in xml_entries:
            if ent not in names:
                continue
            root = etree.fromstring(zf.read(ent))
            full_text = to_nfc(" ".join(t.strip() for t in root.itertext() if t.strip()))

            # Check malformed 7-digit code (Bẫy 5)
            bad_7 = re.findall(r"Q3\d{2}[A-Z]\d{7,}", full_text)
            if bad_7:
                issues.append(f"[Bẫy 5 - 7-digit code] in {ent}: {bad_7}")

            # Check foreign Q3 codes
            found_codes = re.findall(r"Q3\d{2}[A-Z]\d{5,6}", full_text)
            for fc in set(found_codes):
                if code and fc != code:
                    issues.append(f"[Foreign Code] in {ent}: found '{fc}' (expected '{code}')")

            # Check old terms
            if old_terms:
                for term in old_terms:
                    t_nfc = to_nfc(term).strip()
                    if t_nfc and t_nfc.lower() in full_text.lower():
                        issues.append(f"[Leftover Old Term] in {ent}: found '{t_nfc}'")

            # Check highlights and dots per run/paragraph
            for p_i, p_el in enumerate(root.findall(f".//{{{W_NS}}}p")):
                p_txt = to_nfc("".join(p_el.itertext())).strip()
                if "……" in p_txt or "..." in p_txt:
                    dots.append(f"{ent} P#{p_i}: '{p_txt[:80]}'")
                    issues.append(f"[Placeholder Found] in {ent} P#{p_i}: Placeholder '……'/'...' chưa điền: '{p_txt[:80]}'")
                if "{{" in p_txt or "{%" in p_txt:
                    issues.append(f"[Leftover TAG] in {ent} P#{p_i}: TAG template chưa được điền: '{p_txt[:80]}'")
                hl_els = p_el.findall(f".//{{{W_NS}}}highlight")
                if hl_els:
                    highlights.append(f"{ent} P#{p_i} ({len(hl_els)} hl): '{p_txt[:80]}'")
                    issues.append(f"[Yellow Highlight] in {ent} P#{p_i}: Còn {len(hl_els)} đoạn bôi vàng: '{p_txt[:80]}'")

            if ent == "word/document.xml":
                # Check that dossier code in body does NOT contain suffixes (/1, /2, /E, etc.)
                suffixed = re.findall(r"(Q3\d{2}[A-Z]\d{5,6}(?:/\w+)+)", full_text)   # \w: cả chữ có dấu
                if is_bckq or is_tem:
                    # BCKQ: số báo cáo và số phiếu tính toán theo số chứng thư; tem tủ: số tem = số chứng thư
                    # -> cho phép '<mã>/<số>' (AGENTS mục 4, 7), vẫn báo lỗi nếu có /E hoặc đuôi khác.
                    # Phiếu tính toán đánh số theo tên phiếu ('<mã>/Bui', '<mã>/Lưu Lượng'… — người dùng chốt
                    # 05/11/2026, hồ sơ 138143) -> cho phép hậu tố chữ từ 2 ký tự (vẫn bắt '/E').
                    # Hồ sơ nhiều chứng thư: phiếu tính toán '<mã>/<n>/Bụi' (120423 đợt 7) cũng hợp lệ.
                    suffixed = [s for s in suffixed
                                if not re.fullmatch(r"Q3\d{2}[A-Z]\d{5,6}/(\d+(?:/[^\W\d_]{2,})?|[^\W\d_]{2,})", s)]
                if suffixed:
                    issues.append(f"[Body Dossier Code With Suffix] in document.xml: found {set(suffixed)} (Dossier code in body text/tables must be bare without /1, /2, /E, etc.)")

                # Check Table 0 Row 4 label cell vs code cell
                tbls = root.findall(f".//{{{W_NS}}}tbl")
                if tbls:
                    t0 = tbls[0]
                    rows = t0.findall(f".//{{{W_NS}}}tr")
                    if len(rows) > 4:
                        r4 = rows[4]
                        cells = r4.findall(f".//{{{W_NS}}}tc")
                        if cells:
                            c0_txt = to_nfc("".join(cells[0].itertext())).strip()
                            if re.search(r"Q3\d{2}[A-Z]\d{5,6}", c0_txt):
                                issues.append(f"[Table 0 Label Overwritten] in Table 0 row 4: cell 0 contains code '{c0_txt}' instead of application label")

                    # Dòng '- ' trong Table 0 (Thiết bị, Căn cứ) phải thẳng hàng: không thụt treo / thụt dòng đầu (AGENTS mục 14)
                    for p_el in t0.iter(f"{{{W_NS}}}p"):
                        if not to_nfc("".join(p_el.itertext())).lstrip().startswith("-"):
                            continue
                        for ind in p_el.findall(f"{{{W_NS}}}pPr/{{{W_NS}}}ind"):
                            if any(ind.get(f"{{{W_NS}}}{k}") not in (None, "0") for k in ("hanging", "firstLine")):
                                issues.append(f"[Table 0 Dash Misaligned] dòng '{to_nfc(''.join(p_el.itertext()))[:50]}' còn thụt treo/thụt dòng đầu -> các gạch '- ' không thẳng hàng")
                                break

                    # Một máy một tên (AGENTS mục 14): dòng thiết bị phải đúng tên chuẩn trong Core/Tools/equipment_names.py
                    try:
                        from Core.Tools.equipment_names import nonstandard_lines
                        t0_txt = "\n".join(
                            "".join("\n" if el.tag in (f"{{{W_NS}}}br", f"{{{W_NS}}}cr") else (el.text or "")
                                    for el in p_el.iter(f"{{{W_NS}}}t", f"{{{W_NS}}}br", f"{{{W_NS}}}cr"))
                            for p_el in t0.iter(f"{{{W_NS}}}p"))
                        for cur, std in nonstandard_lines(t0_txt):
                            issues.append(f"[Equipment Name] Table 0 ghi '{cur}' -> tên chuẩn '{std}'")
                    except ImportError:
                        pass

                    # Không ghi gọn (AGENTS mục 15): Table 0 tiếng Việt không còn viết tắt; Địa điểm phải nêu tên đơn vị
                    try:
                        from Core.Tools.no_abbrev import find_abbrev, location_missing_org
                        is_en_doc = bool(re.search(r"(^|[ _-])en([ _-]|$)", base_lower)) and not is_bckq
                        for tr in rows:
                            tcs = tr.findall(f"{{{W_NS}}}tc")
                            if len(tcs) < 2:
                                continue
                            label = to_nfc("".join(tcs[0].itertext())).strip()
                            val = to_nfc(" ".join("".join(p.itertext()) for p in tcs[-1].findall(f".//{{{W_NS}}}p"))).strip()
                            if not is_en_doc and re.search(r"(Đơn vị|Địa chỉ|Địa điểm)", label):
                                ab = find_abbrev(val)
                                if ab:
                                    issues.append(f"[Abbreviation] Table 0 '{label[:30]}' còn viết tắt {ab}: '{val[:80]}'")
                            if re.search(r"(Địa điểm|Inspection address)", label) and location_missing_org(val, is_en_doc):
                                issues.append(f"[Location Short] Table 0 '{label[:30]}' không nêu tên đơn vị: '{val[:80]}'"
                                              " -> ghi đủ như BBLM mục 5 (Địa điểm thực hiện), không rút gọn")
                    except ImportError:
                        pass

                    # Check Table 0 signature horizontal alignment (AGENTS.md Section 3)
                    # When right cell has 2 title lines (KT. GIÁM ĐỐC / FOR DIRECTOR on line 1, PHÓ GIÁM ĐỐC / VICE DIRECTOR on line 2),
                    # left cell MUST have line 1 as blank spacer and line 2 as GIÁM ĐỊNH VIÊN / INSPECTOR
                    # so that INSPECTOR aligns horizontally with VICE DIRECTOR, and GIÁM ĐỊNH VIÊN with PHÓ GIÁM ĐỐC.
                    if rows:
                        last_row = rows[-1]
                        cells = last_row.findall(f".//{{{W_NS}}}tc")
                        if len(cells) >= 2:
                            c_left, c_right = cells[0], cells[-1]
                            p_left = [to_nfc("".join(p.itertext())).strip() for p in c_left.findall(f".//{{{W_NS}}}p")]
                            p_right = [to_nfc("".join(p.itertext())).strip() for p in c_right.findall(f".//{{{W_NS}}}p")]
                            p0_r = p_right[0].upper() if len(p_right) > 0 else ""
                            p1_r = p_right[1].upper() if len(p_right) > 1 else ""
                            if ("KT. GIÁM ĐỐC" in p0_r or "FOR DIRECTOR" in p0_r) and ("PHÓ GIÁM ĐỐC" in p1_r or "VICE DIRECTOR" in p1_r):
                                p0_l = p_left[0].upper() if len(p_left) > 0 else ""
                                p1_l = p_left[1].upper() if len(p_left) > 1 else ""
                                if "GIÁM ĐỊNH VIÊN" in p0_l or "INSPECTOR" in p0_l:
                                    issues.append(f"[Table 0 Signature Misaligned] Left title '{p_left[0]}' is on line 1 (aligned with '{p_right[0]}'). It MUST be on line 2 to align horizontally with '{p_right[1]}'.")
                                elif not ("GIÁM ĐỊNH VIÊN" in p1_l or "INSPECTOR" in p1_l):
                                    issues.append(f"[Table 0 Signature Missing] Line 2 of left signature cell must be GIÁM ĐỊNH VIÊN or INSPECTOR to align horizontally with '{p_right[1]}'.")

                                # Check 3.0 cm signature space (w:before >= 1650 twips ~ 3.0 cm)
                                paras_l = c_left.findall(f".//{{{W_NS}}}p")
                                if paras_l:
                                    last_p = paras_l[-1]
                                    sp = last_p.find(f".//{{{W_NS}}}spacing")
                                    if sp is not None:
                                        before_val = int(sp.get(f"{{{W_NS}}}before", "0"))
                                        if before_val < 1650:
                                            cm_val = before_val / 567.0
                                            issues.append(f"[Table 0 Signature Space Too Small] Left signature space is only {cm_val:.2f} cm ({before_val} twips). Must be 3.0 cm (~1701 twips).")

                    # Check Table 0 Align Top in BCKQ (AGENTS.md Section 7)
                    if is_bckq and rows:
                        for r_i, row in enumerate(rows):
                            for c_i, cell in enumerate(row.findall(f"{{{W_NS}}}tc")):
                                tcPr = cell.find(f"{{{W_NS}}}tcPr")
                                if tcPr is not None:
                                    v = tcPr.find(f"{{{W_NS}}}vAlign")
                                    if v is not None and v.get(f"{{{W_NS}}}val") not in ("top", None):
                                        warnings.append(f"[BCKQ Table 0 vAlign] Row {r_i} Cell {c_i} has vAlign='{v.get(f'{{{W_NS}}}val')}' instead of 'top'")

                # Check duplicate wp:docPr IDs (Bẫy 34)
                docPr_ids = [
                    dp.get("id")
                    for dp in root.findall(f".//{{{WP_NS}}}docPr")
                    if dp.get("id") is not None
                ]
                dup_ids = sorted(set(x for x in docPr_ids if docPr_ids.count(x) > 1))
                if dup_ids:
                    issues.append(f"[Bẫy 34 - Duplicate docPr id]: {dup_ids}")

                # Check leftover comments (Bẫy 50)
                comments_ref = root.findall(f".//{{{W_NS}}}commentRangeStart") + root.findall(f".//{{{W_NS}}}commentReference")
                if comments_ref:
                    warnings.append(f"[Bẫy 50 - Word Comments]: {len(comments_ref)} comment tag(s) in document.xml")

                # Check keepNext on 'Đánh giá:' / 'Assessment:' (Phần 5)
                for p_el in root.findall(f".//{{{W_NS}}}p"):
                    pt = to_nfc("".join(p_el.itertext())).strip()
                    if pt.startswith("Đánh giá") or pt.startswith("Assessment"):
                        pPr = p_el.find(f"{{{W_NS}}}pPr")
                        if pPr is None or pPr.find(f"{{{W_NS}}}keepNext") is None:
                            warnings.append(f"[Phần 5 - Missing keepNext] on paragraph: '{pt[:50]}'")

                # Check repeat header row (<w:tblHeader/>) on data tables (Phần 5)
                for t_idx, tbl_el in enumerate(root.findall(f".//{{{W_NS}}}tbl")):
                    if t_idx == 0:
                        continue
                    tbl_rows = tbl_el.findall(f"{{{W_NS}}}tr")
                    if len(tbl_rows) <= 1:
                        continue
                    first_txt = to_nfc("".join(tbl_rows[0].itertext())).strip()
                    if first_txt.startswith("Ghi chú") or first_txt.startswith("Note"):
                        continue
                    if tbl_rows[0].find(f".//{{{W_NS}}}drawing") is not None:
                        continue
                    trPr0 = tbl_rows[0].find(f"{{{W_NS}}}trPr")
                    if trPr0 is None or trPr0.find(f"{{{W_NS}}}tblHeader") is None:
                        warnings.append(f"[Phần 5 - Missing tblHeader] Table #{t_idx} header row missing <w:tblHeader/> (Repeat Header Row)")

                # Check vertical STT merge for indicators with child rows (AGENTS.md Section 3)
                if not is_tem and not is_bckq:
                    for t_idx, tbl_el in enumerate(root.findall(f".//{{{W_NS}}}tbl")):
                        tbl_rows = tbl_el.findall(f"{{{W_NS}}}tr")
                        if len(tbl_rows) <= 2:
                            continue
                        row0_cells = tbl_rows[0].findall(f"{{{W_NS}}}tc")
                        if not row0_cells:
                            continue
                        c0_hdr = to_nfc("".join(row0_cells[0].itertext())).strip().lower()
                        if c0_hdr not in ["stt", "no.", "no", "tt"]:
                            continue

                        curr_stt = None
                        curr_start_r = None
                        for r_idx in range(1, len(tbl_rows)):
                            row_el = tbl_rows[r_idx]
                            cells = row_el.findall(f"{{{W_NS}}}tc")
                            if not cells:
                                continue
                            if len(cells) == 1 or row_el.find(f".//{{{W_NS}}}drawing") is not None:
                                curr_stt = None
                                continue
                            c0 = cells[0]
                            c0_txt = "".join(c0.itertext()).strip()
                            vMerge = c0.find(f".//{{{W_NS}}}vMerge")
                            row_txt = "".join(row_el.itertext()).strip()
                            if not row_txt:
                                curr_stt = None
                                continue

                            if c0_txt:
                                curr_stt = c0_txt
                                curr_start_r = r_idx
                            elif curr_stt is not None:
                                if vMerge is None:
                                    issues.append(
                                        f"[Unmerged STT Child Row] Table #{t_idx} row #{r_idx} is a child row under STT '{curr_stt}' (row #{curr_start_r}) but STT cell is not vertically merged (<w:vMerge>)."
                                    )

                # Check accreditation sentence in the 'Ghi chú:' / 'Note:' table right under every result-data table
                # (AGENTS.md Section 3; area-info / photo / layout tables are not checked)
                if not is_tem and not is_bckq:
                    is_en_doc = " - en - " in base_lower
                    accredit = ACCREDIT_EN if is_en_doc else ACCREDIT_VN
                    acc_norm = re.sub(r"\s+", " ", accredit.lower())
                    body = root.find(f"{{{W_NS}}}body")
                    body_tbls = body.findall(f"{{{W_NS}}}tbl") if body is not None else []
                    tbl_txt = [to_nfc("".join(t.itertext())).strip() for t in body_tbls]
                    # Câu công nhận theo quy chuẩn: nước QCVN …/BYT -> VIAS 004; nước / bùn / chất thải còn lại -> VIMCERTS 078
                    # (người dùng chốt 16/11/2026, AGENTS mục 3, 8C)
                    vim = is_vimcerts(tbl_txt[0]) if tbl_txt else False
                    vim_txt = "Parameters are accredited to VIMCERTS 078" if is_en_doc else "Các thông số được chứng nhận VIMCERTS 078"
                    is_note = [t.startswith(("Ghi chú", "Note")) for t in tbl_txt]
                    for t_idx, tbl_el in enumerate(body_tbls):
                        if t_idx == 0 or is_note[t_idx]:
                            continue
                        tbl_rows = tbl_el.findall(f"{{{W_NS}}}tr")
                        if tbl_el.find(f".//{{{W_NS}}}drawing") is not None:
                            continue    # bảng ảnh / sơ đồ (chú thích "Vị trí đo …" không phải bảng kết quả)
                        hdr = to_nfc(" ".join("".join(r.itertext()) for r in tbl_rows[:2])).lower()
                        if not RESULT_TABLE_RE.search(hdr):
                            continue
                        want = vim_txt if vim else accredit
                        if not (t_idx + 1 < len(body_tbls) and is_note[t_idx + 1]):
                            issues.append(f"[Missing Accreditation Note] Table #{t_idx} (result table) has no 'Ghi chú:/Note:' table right below it (must contain '{want}').")
                            continue
                        nt = _note_norm(tbl_txt[t_idx + 1])
                        has_vias, has_vim = acc_norm in nt, VIMCERTS_KEY in nt
                        if vim:
                            if not has_vim:
                                issues.append(f"[Missing Accreditation Note] Note table under result table #{t_idx}: mẫu nước / bùn / chất thải không theo QCVN …/BYT phải ghi '● {vim_txt}.' ở dòng đầu (AGENTS mục 3, 8C).")
                            else:
                                note_rows = body_tbls[t_idx + 1].findall(f"{{{W_NS}}}tr")
                                first = _note_norm("".join(note_rows[1].itertext())) if len(note_rows) > 1 else ""
                                if VIMCERTS_KEY not in first:
                                    issues.append(f"[Accreditation Order] Note table under result table #{t_idx}: câu VIMCERTS 078 phải ở dòng đầu bảng Ghi chú.")
                            if has_vias:
                                issues.append(f"[Wrong Accreditation Note] Note table under result table #{t_idx}: hồ sơ VIMCERTS không ghi câu VIAS 004 — chỉ ghi câu VIMCERTS 078.")
                        else:
                            if has_vim:
                                issues.append(f"[Wrong Accreditation Note] Note table under result table #{t_idx}: câu VIMCERTS 078 chỉ dùng cho mẫu nước / bùn / chất thải không theo QCVN …/BYT — hồ sơ này ghi '{accredit}.'")
                            if not has_vias:
                                issues.append(f"[Missing Accreditation Note] Note table under result table #{t_idx} lacks '{accredit}.'")

                # Check 2-column / 3-column image tables and image crop (AGENTS.md Section 3)
                for t_idx, tbl_el in enumerate([] if is_tem else root.findall(f".//{{{W_NS}}}tbl")):
                    tbl_rows = tbl_el.findall(f"{{{W_NS}}}tr")
                    if not tbl_rows:
                        continue
                    drawings = tbl_el.findall(f".//{{{W_NS}}}drawing")
                    if not drawings:
                        continue
                    # Check if table is a 2-column image grid
                    row0_cells = tbl_rows[0].findall(f"{{{W_NS}}}tc")
                    if len(row0_cells) == 2:
                        for drw in drawings:
                            ext = drw.find(f".//{{{WP_NS}}}extent")
                            if ext is not None:
                                cx = int(ext.get("cx", 0))
                                cx_cm = cx / 360000.0
                                # Rule: horizontal width 7 - 8 cm (tolerate 6.8 - 8.2 cm)
                                if cx > 0 and (cx < 2448000 or cx > 2952000):
                                    warnings.append(f"[Image Width in 2-Col Grid] Table #{t_idx} image width is {cx_cm:.2f} cm (expected 7 - 8 cm per rule).")
                    elif len(row0_cells) == 3:
                        for drw in drawings:
                            ext = drw.find(f".//{{{WP_NS}}}extent")
                            if ext is not None:
                                cx = int(ext.get("cx", 0))
                                cx_cm = cx / 360000.0
                                # Rule: horizontal width 4 - 5 cm (tolerate 3.8 - 5.2 cm)
                                if cx > 0 and (cx < 1368000 or cx > 1872000):
                                    warnings.append(f"[Image Width in 3-Col Grid] Table #{t_idx} image width is {cx_cm:.2f} cm (expected 4 - 5 cm per rule).")

                # Check strictly NO crop on images in reports
                if not is_tem:
                    for drw in root.findall(f".//{{{W_NS}}}drawing"):
                        sr = drw.find(f".//{{{A_NS}}}srcRect")
                        if sr is not None and sr.attrib:
                            crop_vals = {k: v for k, v in sr.attrib.items() if int(v) > 0}
                            if crop_vals:
                                issues.append(f"[Image Crop Detected] in {ent}: Image has active crop {crop_vals} (Quy tắc: tuyệt đối không crop ảnh).")

                # Check for unnumbered KN when numbered KN (e.g. KN2) is present in the document
                if "KN2" in full_text:
                    bad_kn_captions = re.findall(r"(?:lấy\s+mẫu|vị\s+trí)\s+KN(?![0-9])", full_text, re.IGNORECASE)
                    if bad_kn_captions:
                        issues.append(f"[Unnumbered KN in Multi-Point Dossier] found {bad_kn_captions} when dossier has KN2. First point must be 'KN1', not generic 'KN'.")

                # Cách gọi "vị trí đo" (không "điểm đo"), độ sáng không ghi số lượng — AGENTS mục 6 (11/11/2026)
                if ent == "word/document.xml" and " - en - " not in base_lower and not is_tem:
                    from Core.Tools.wording import find_point_terms
                    bad_terms = find_point_terms(full_text)
                    if bad_terms:
                        issues.append(f"[Point Term] còn {bad_terms} — gọi 'vị trí đo' / 'vị trí', độ sáng ghi "
                                      f"'Cường độ sáng trung bình' không kèm số lượng (Core/Tools/wording.py)")

                # Check standard QUATEST 3 ending double line (thinThick, 4.5pt, centered) - TBKQ only, not required for BCKQ or tem tu
                if is_bckq or is_tem:
                    # Rule: BCKQ is created BEFORE TBKQ / Chứng thư giám định.
                    # BCKQ MUST NOT cite "TBKQ" or "Thông báo kết quả" or "Chứng thư giám định".
                    # It must ONLY cite calculation sheets ("phiếu tính toán").
                    bad_cites = re.findall(r"(?:TBKQ\s+số|Thông\s+báo\s+kết\s+quả\s+số|Chứng\s+thư\s+(?:giám\s+định\s+)?số)", full_text, re.IGNORECASE)
                    if bad_cites:
                        issues.append(f"[BCKQ Cites Certificate/TBKQ] in {ent}: found {bad_cites}. BCKQ is created before TBKQ and must ONLY cite calculation sheets ('phiếu tính toán').")
                else:
                    body_paras = root.findall(f".//{{{W_NS}}}body/{{{W_NS}}}p")
                    total_body_paras = len(body_paras)
                    ending_lines = []
                    for p_idx, p_el in enumerate(body_paras):
                        p_xml = etree.tostring(p_el).decode("utf-8")
                        if "thinThick" in p_xml or "Straight Connector" in p_xml:
                            is_centered = (
                                ("<wp:align>center</wp:align>" in p_xml)
                                or ("mso-position-horizontal:center" in p_xml)
                                or ("1094105" in p_xml)
                                or ("1050290" in p_xml)
                                or ("1121728" in p_xml)
                            )
                            has_tt = "thinThick" in p_xml
                            from_end = total_body_paras - p_idx
                            ending_lines.append({
                                "p_idx": p_idx,
                                "from_end": from_end,
                                "is_centered": is_centered,
                                "has_thin_thick": has_tt
                            })

                    if not ending_lines:
                        issues.append("[Missing Ending Line] in document.xml: Document is missing the standard QUATEST 3 ending double line (thinThick, 4.5pt, centered) at the end of the document.")
                    else:
                        valid_end_lines = [el for el in ending_lines if el["from_end"] <= 3 and el["has_thin_thick"] and el["is_centered"]]
                        if not valid_end_lines:
                            issues.append(f"[Invalid Ending Line] in document.xml: Ending line found at paragraph(s) {[el['p_idx'] for el in ending_lines]} but is not properly configured at the end with thinThick and center alignment.")
                        # Đoạn trống ngay trước đường kết thúc -> đường kẻ xa dòng cuối, dễ sinh khoảng trắng/trang chỉ có
                        # đường kẻ (AGENTS mục 3, 5; 127215/120423 10/11/2026 có 6 đoạn trống)
                        last_el = body_paras[ending_lines[-1]["p_idx"]]
                        gap_n, prev_el = 0, last_el.getprevious()
                        while prev_el is not None and prev_el.tag == f"{{{W_NS}}}p":
                            if "".join(t.text or "" for t in prev_el.iter(f"{{{W_NS}}}t")).strip():
                                break
                            if any(True for x in prev_el.iter() if isinstance(x.tag, str) and
                                   x.tag.split("}")[-1] in ("drawing", "pict", "object", "AlternateContent", "sectPr")):
                                break
                            gap_n += 1
                            prev_el = prev_el.getprevious()
                        if gap_n:
                            issues.append(f"[Ending Gap] in document.xml: {gap_n} đoạn trống ngay trước đường kết thúc — đường kẻ phải sát dòng cuối (dựng lại bằng Core mới: add_document_ending_line tự xóa).")
                        orphaned = [el for el in ending_lines if el["from_end"] > 3]
                        if orphaned:
                            warnings.append(f"[Orphaned Connector] in document.xml: Found misplaced connector line at paragraph(s) {[el['p_idx'] for el in orphaned]} (should only be at the end of the document).")

        # Check BCKQ 'Người báo cáo' tab count to prevent line wrapping (AGENTS.md Section 7)
        if is_bckq and "word/document.xml" in names:
            root = etree.fromstring(zf.read("word/document.xml"))
            for p_el in root.findall(f".//{{{W_NS}}}p"):
                pt = to_nfc("".join(p_el.itertext())).strip()
                if "Người báo cáo" in pt or "Người  báo cáo" in pt:
                    tab_count = len(p_el.findall(f".//{{{W_NS}}}tab"))
                    if tab_count > 7:
                        issues.append(f"[Line Wrap Risk] 'Người báo cáo' has {tab_count} tabs (maximum 7 tabs allowed to prevent line wrapping)")

    # Cỡ chữ, hàng cao thừa, Ghi chú sát bảng (TBKQ, BCKQ; không áp cho tem tủ)
    if not is_tem:
        try:
            li, lw = check_layout_rules(docx_path, is_bckq)
            issues.extend(li)
            warnings.extend(lw)
        except Exception as e:
            warnings.append(f"[Layout Check Skipped] {e}")

    # Ô giá trị Table 0 thẳng cột (người dùng báo 137035, 09/11/2026 — 'From …' lùi trái so với các dòng khác)
    if not is_tem and not is_bckq:
        try:
            issues.extend(check_table0_align(docx_path))
        except Exception as e:
            warnings.append(f"[Table0 Align Skipped] {e}")
        try:
            issues.extend(sorted(set(check_note_label(docx_path))))
        except Exception as e:
            warnings.append(f"[Note Label Skipped] {e}")

    # Sơ đồ draw.io sửa sau khi dựng DOCX -> DOCX đang chứa sơ đồ cũ (AGENTS 8C)
    if not is_tem:
        warnings.extend(check_layout_stale(docx_path))

    status = "CLEAN" if not issues else "ISSUES_FOUND"
    print(f"\n=== QA Check: {os.path.basename(docx_path)} [{status}] ===")
    if issues:
        for iss in issues:
            print(f"  [ERROR] {iss}")
    if warnings:
        for w in warnings:
            print(f"  [WARN]  {w}")
    print(f"  Placeholders ('……'/'...'): {len(dots)} paragraph(s)")
    for d in dots[:10]:
        print(f"    - {d}")
    print(f"  Yellow Highlights (<w:highlight>): {len(highlights)} paragraph(s)")
    for h in highlights[:10]:
        print(f"    - {h}")

    return {
        "file": docx_path,
        "status": status,
        "issues": issues,
        "warnings": warnings,
        "dots": dots,
        "highlights": highlights,
    }


def test_open_word(targets: list[str], export_pdf: bool = False) -> bool:
    """Open each .docx in Word COM to check readability and actual page counts (Bẫy 11, 24, 27, 36). Returns True if all passed."""
    try:
        from Core.Exporters.pdf_exporter import avoid_label_printer
        avoid_label_printer()
    except Exception as _e:
        pass

    import win32com.client

    all_ok = True
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        for fpath in targets:
            abs_p = os.path.abspath(fpath)
            try:
                doc = w.Documents.Open(abs_p, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
                doc.Repaginate()
                pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
                secs = doc.Sections.Count
                print(f"  [WORD OPEN OK] {os.path.basename(fpath)} -> Pages={pages}, Sections={secs}")
                if "bckq" in os.path.basename(fpath).lower():
                    for wp in doc.Paragraphs:
                        if "Người báo cáo" in wp.Range.Text:
                            lines = wp.Range.ComputeStatistics(1)
                            if lines > 1:
                                print(f"    [FAIL] 'Người báo cáo' wrapped into {lines} lines (expected 1 line)!")
                                all_ok = False
                            else:
                                print(f"    [PASS] 'Người báo cáo' is properly on a single line.")
                if export_pdf:
                    pdf_out = os.path.splitext(abs_p)[0] + ".check.pdf"
                    doc.SaveAs2(pdf_out, FileFormat=17)
                    print(f"    Exported PDF -> {pdf_out}")
                doc.Close(SaveChanges=False)
            except Exception as e:
                print(f"  [WORD OPEN FAILED] {os.path.basename(fpath)} -> {e}")
                all_ok = False
    finally:
        w.Quit()
    return all_ok




def check_pdf_with_upchungthu(targets: list[str], cert_date: str = None, add_bg: bool = False) -> bool:
    """Run UpChungThu comprehensive verification on PDF file(s). Returns True if all passed."""
    try:
        from Core.Services.upchungthu_bridge import process_tbkq_pdf, verify_tbkq_pdf, is_tbkq_file
    except Exception as e:
        print(f"  [Warning] Cannot load UpChungThu bridge: {e}")
        return False

    print("\n--- [UpChungThu QA] Verifying TBKQ Deliverable PDFs ---")
    pdf_targets = []
    for t in targets:
        if os.path.isdir(t):
            for root, _, files in os.walk(t):
                for f in sorted(files):
                    if f.lower().endswith(".pdf") and is_tbkq_file(f):
                        pdf_targets.append(os.path.join(root, f))
        elif t.lower().endswith(".pdf"):
            pdf_targets.append(t)

    if not pdf_targets:
        print("  No TBKQ PDF files found to verify.")
        return True

    all_ok = True
    for p in pdf_targets:
        res = process_tbkq_pdf(p, cert_date=cert_date, add_bg=add_bg, print_log=True)
        if res.get("status") != "OK":
            all_ok = False
    return all_ok


def main():
    parser = argparse.ArgumentParser(description="QA & Final Check Tool for QUATEST 3")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_fc = sub.add_parser("final-check", help="Run XML & rule verification on .docx file(s)")
    p_fc.add_argument("target", help="Path to .docx file or directory")
    p_fc.add_argument("--code", help="Expected profile code e.g. Q326D108033")
    p_fc.add_argument("--old", help="Comma-separated list of old strings to check against")

    p_to = sub.add_parser("test-open", help="Open .docx file(s) in Word COM and verify page counts")
    p_to.add_argument("target", help="Path to .docx file or directory")
    p_to.add_argument("--export-pdf", action="store_true", help="Export .check.pdf for visual inspection")

    p_cp = sub.add_parser("check-pdf", help="Run UpChungThu error-checking on exported TBKQ PDF file(s)")
    p_cp.add_argument("target", help="Path to PDF file or directory")
    p_cp.add_argument("--date", help="Expected certificate date (dd/mm/yyyy)")
    p_cp.add_argument("--add-bg", action="store_true", help="Also apply background template if missing")

    args = parser.parse_args()
    targets = []
    if os.path.isdir(args.target):
        for root, _, files in os.walk(args.target):
            for f in sorted(files):
                if args.cmd == "check-pdf":
                    if f.lower().endswith(".pdf") and not f.startswith("~$"):
                        try:
                            from ..Services.upchungthu_bridge import is_tbkq_file
                        except ImportError:
                            from Core.Services.upchungthu_bridge import is_tbkq_file
                        if is_tbkq_file(f):
                            targets.append(os.path.join(root, f))
                else:
                    if f.lower().endswith(".docx") and not f.startswith("~$"):
                        targets.append(os.path.join(root, f))
    else:
        targets.append(args.target)

    has_issues = False
    if args.cmd == "final-check":
        old_list = [x.strip() for x in args.old.split(",") if x.strip()] if args.old else []
        for t in targets:
            res = check_single_docx(t, args.code, old_list)
            if res.get("status") != "CLEAN" or res.get("issues"):
                has_issues = True
        # Also auto-check TBKQ PDFs if target is a dossier directory
        if os.path.isdir(args.target):
            pdf_dir = os.path.join(args.target, "pdf")
            if os.path.isdir(pdf_dir):
                pdf_ok = check_pdf_with_upchungthu([pdf_dir])
                if not pdf_ok:
                    has_issues = True
        if has_issues:
            sys.exit(1)
    elif args.cmd == "test-open":
        ok = test_open_word(targets, args.export_pdf)
        if not ok:
            sys.exit(1)
    elif args.cmd == "check-pdf":
        ok = check_pdf_with_upchungthu(targets, cert_date=args.date, add_bg=args.add_bg)
        if not ok:
            sys.exit(1)


if __name__ == "__main__":
    main()
