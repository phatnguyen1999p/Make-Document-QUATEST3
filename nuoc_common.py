# -*- coding: utf-8 -*-
"""
Bộ dựng chung hồ sơ nước NHIỀU MẪU (nước thải, nước dưới đất, nước uống / sinh hoạt) — mỗi mẫu một chứng thư /n,
BCKQ gộp 1 file (AGENTS mục 8C "Nước thải / …", "Nước uống / NSH"). Khuôn bìa / phụ lục / Ghi chú lấy từ
`build_nuoc_mau.py` (hồ sơ ảo người dùng đã xem 16/11/2026); số liệu phòng thí nghiệm đọc thẳng từ phiếu KQTN PDF
(`Core/Tools/kqtn_reader.py`) — tên chỉ tiêu, phương pháp, kết quả, giới hạn phát hiện ghi theo phiếu KQTN.

Script hồ sơ (`build_<6 số>.py`) chỉ khai báo CONFIG rồi gọi ``run(CONFIG)``. Xem `build_135228.py` làm mẫu.

Quy tắc áp:
  - KQTN ``KPH`` -> ``< <giới hạn phát hiện>``; ``< x (LOQ)`` -> ``< x`` + ghi chú "Giới hạn định lượng của phương pháp
    thử." (AGENTS mục 13, 8C).
  - Nhóm chất (PCBs, thuốc BVTV): 1 STT, cấu tử ``+ …`` trong cùng ô (form nước thải).
  - Nước không theo QCVN …/BYT: Ghi chú dòng đầu ``● Các thông số được chứng nhận VIMCERTS 078.``; ``mg/L``; chữ
    "thông số"; không có mức -> bảng 5 cột, Căn cứ ``-``, không Đánh giá (AGENTS mục 3, 8C).
  - Chưa có KQTN -> ô kết quả ``………`` tô vàng, ngày ban hành ``………`` (AGENTS mục 2, 13).
"""
import os
import re
import sys
import copy
import glob
import datetime

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ["QUATEST3_NETWORK_ROOT"] = r"C:\__khong_dong_bo_Z__"     # không đồng bộ Z từ code (AGENTS mục 0, E21)

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_COLOR_INDEX

import build_nuoc_mau as M
from build_nuoc_mau import (Body, cell_text, note_table, fix_cover, appendix_protos, make_head, _keep, _keep_rows,
                            merge_text_runs, evaluate, en_num)
from Core.Models import Dossier, ClientInfo, PersonnelInfo
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, merge_same_vertical, style_note_tables, appendix_total_field
from Core.Styling.elements import add_document_ending_line
from Core.Styling.tables import set_table_autofit_window, insert_image_to_cell, add_cant_split
from Core.Tools.equipment_names import eq
from Core.Tools.kqtn_reader import read_kqtn, read_kqtn_all
from Core.Tools.qa_check import check_single_docx

FIELD = "Đo tại hiện trường"
TODAY = datetime.date.today().strftime("%d/%m/%Y")

# Tên gọn cho cột "Chỉ tiêu thử nghiệm" (bảng 9.1) — khóa: chuỗi con trong tên chỉ tiêu KQTN (chữ thường)
SHORT = [
    ("nhiệt độ", "nhiệt độ"), ("ph", "pH"), ("độ màu", "độ màu"), ("màu sắc", "độ màu"), ("bod", "BOD₅"), ("cod", "COD"),
    ("tss", "TSS"), ("chất rắn lơ lửng", "TSS"), ("asen", "As"), ("thủy ngân", "Hg"), ("chì", "Pb"), ("cadmi", "Cd"),
    ("xyanua", "CN⁻"), ("crom vi", "Cr⁶⁺"), ("crom iii", "Cr³⁺"), ("đồng", "Cu"), ("kẽm", "Zn"), ("niken", "Ni"),
    ("mangan", "Mn"), ("sắt", "Fe"), ("phenol", "tổng phenol"), ("dầu mỡ khoáng", "dầu mỡ khoáng"),
    ("sunfua", "S²⁻"), ("florua", "F⁻"), ("clorua", "Cl⁻"), ("amoni", "amoni"), ("nitơ", "tổng nitơ"),
    ("phospho", "tổng phốt pho"), ("photpho", "tổng phốt pho"), ("clo dư", "clo dư"), ("coliform", "Coliform"),
    ("polychlorinated", "PCBs"), ("chlor hữu cơ", "hóa chất BVTV clo hữu cơ"),
    ("phospho hữu cơ", "hóa chất BVTV phốt pho hữu cơ"), ("alpha", "tổng hoạt độ phóng xạ α"),
    ("beta", "tổng hoạt độ phóng xạ β"), ("tds", "TDS"), ("tổng chất rắn hòa tan", "TDS"),
]


def short_name(name):
    low = name.lower()
    for key in ("phospho hữu cơ", "chlor hữu cơ", "polychlorinated"):
        if key in low:
            return dict(SHORT)[key]
    if low == "ph":
        return "pH"
    for k, v in SHORT:
        if k == "ph":
            continue
        if k in low:
            return v
    return re.sub(r"^Hàm lượng\s+", "", name)


def unit_fmt(u):
    u = (u or "").strip()
    if u.lower() in ("mg/l",):
        return "mg/L"
    if u.replace(" ", "").upper().startswith("MPN/100"):
        return "MPN/ 100 mL"
    m = re.match(r"^CFU/\s*(\d+)\s*ml$", u.replace(" ", ""), re.I)
    if m:                                           # CFU/250mL -> CFU/ 250 mL (như khung người dùng)
        return f"CFU/ {m.group(1)} mL"
    u = re.sub(r"/l$", "/L", u)
    return u or "-"


# ---------------------------------------------------------------------------
# KQTN -> hàng bảng kết quả
# ---------------------------------------------------------------------------
KPH_LOD = False   # cfg["kph_lod"]: KPH -> giới hạn phát hiện, không "<" (AGENTS mục 13, 19/11/2026); đặt trong run()


def _kv(r):
    if KPH_LOD:
        v, loq, _ = r.ket_qua()
        return v, loq
    return r.value()


def kqtn_rows(kq):
    """Trả list dict(name, unit, method, result, loq, vim) — nhóm chất gộp 1 hàng, cấu tử '+ …'."""
    rows, i, src = [], 0, kq["rows"]
    while i < len(src):
        r = src[i]
        if r.level == 0 and r.result.strip() == "-" and i + 1 < len(src) and src[i + 1].level >= 1:
            names, res, loq, vim = [r.name], [""], False, bool(r.marks)
            j = i + 1
            while j < len(src) and src[j].level >= 1:
                c = src[j]
                if c.result.strip() != "-":
                    v, lq = _kv(c)
                    names.append("+ " + re.sub(r"^Hàm lượng\s+", "", c.name))
                    res.append(v)
                    loq = loq or lq
                j += 1
            rows.append(dict(name="\n".join(names), unit=unit_fmt(r.unit), method=r.method,
                             result="\n".join(res), loq=loq, vim=vim, group=True))
            i = j
            continue
        v, lq = _kv(r)
        v = re.sub(r"\s*⁽[ᵃ-ᶻa-z]⁾$", "", v)         # '<1⁽ᶜ⁾' (KQTN: '< 1 CFU' = không phát hiện) -> '< 1'
        v = re.sub(r"^<\s*(?=\d)", "< ", v)
        rows.append(dict(name=r.name, unit=unit_fmt(r.unit), method=r.method, result=v, loq=lq, vim=bool(r.marks),
                         group=False))
        i += 1
    return rows


def field_rows(sample):
    """Hàng đo / cảm quan tại hiện trường của TBKQ (theo thứ tự ``field_tbkq``) khi đã có KQTN."""
    out = []
    for key in sample.get("field_tbkq", ()):
        out.append(dict(name=FIELD_NAMES.get(key, key), unit=FIELD_UNITS.get(key, "-"),
                        method="Cảm quan" if key == "Mùi" else FIELD, result=sample["field"][key], loq=False,
                        vim=True, group=False))
    return out


def clo(v):
    """Clo dư đo nhanh Hach DR900 / DR300 đọc dưới 0,1 mg/L -> '< 0,1' (AGENTS mục 8C, nước uống / NSH điểm 5)."""
    try:
        return "< 0,1" if float(v.replace(",", ".")) < 0.1 else v
    except ValueError:
        return v


def pyc_index(path):
    """{đuôi tên mẫu PYC (N1, NT, M3…): [(tên, phương pháp)]} + dict PYC (ngày dự kiến trả…)."""
    from Core.Tools.pyc_reader import read_pyc
    p = read_pyc(path)
    return {re.split(r"\s*-\s*", s["ten"])[-1].replace(" ", ""): s["params"] for s in p["samples"]}, p


def placeholder_rows(sample):
    """Chưa có KQTN: hàng theo danh sách chỉ tiêu (CONFIG ``params`` hoặc đọc từ khung làm tay ``draft``), kết quả
    phòng thí nghiệm '………'; hàng đo / cảm quan tại hiện trường lấy số BBLM khai trong ``field``."""
    params = sample.get("params") or params_from_draft(sample["draft"])
    out = []
    for n, u, m in params:
        if m.startswith(("Đo tại hiện trường", "Cảm quan")):
            val = sample.get("field", {}).get(_field_key(n), PLACEHOLDER)
            out.append(dict(name=n, unit=unit_fmt(u), method=m, result=val, loq=False, vim=True, group=False))
        else:
            grp = "\n" in n
            res = "\n".join([""] + [PLACEHOLDER] * n.count("\n")) if grp else PLACEHOLDER
            # chưa KQTN: phương pháp để trống cùng kết quả, thay theo phiếu KQTN khi về (người dùng chốt 17/11/2026,
            # hồ sơ 131873 — AGENTS mục 13)
            out.append(dict(name=n, unit=unit_fmt(u), method=PLACEHOLDER, result=res, loq=False, vim=True, group=grp))
    return out


def _field_key(name):
    low = name.lower()
    for k in ("Nhiệt độ", "pH", "Clo dư tự do", "Clo dư", "Mùi", "TDS", "DO", "Oxy hòa tan"):
        if k.lower() in low:
            return k
    return name


def params_from_draft(path, with_limit=False):
    """Đọc bảng kết quả (STT | tên | đơn vị | phương pháp | kết quả …) của khung TBKQ làm tay -> [(tên, đơn vị,
    phương pháp)]; ô phương pháp gộp dọc (trống) lấy theo hàng trên; sửa lỗi gõ của khung (SWEWW, BOD5…)."""
    d = Document(path)
    for t in d.tables:
        head = [c.text.strip().lower() for c in t.rows[0].cells]
        if not head or head[0] != "stt" or "phương pháp" not in " ".join(head):
            continue
        out, last_m, last_field = [], "", False
        for r in t.rows[1:]:
            cells = [c.text.strip() for c in r.cells]
            if not cells[0].isdigit():
                continue
            name = re.sub(r"[ \t]*\n[ \t]*", "\n", cells[1]).strip()
            if "\n" in name and not name.split("\n")[1].startswith("+"):
                name = " ".join(name.split("\n"))
            for a, b in (("BOD5", "BOD₅"), ("O2", "O₂"), ("NH4+", "NH₄⁺"), ("NH3", "NH₃"), ("CN-", "CN⁻"),
                         ("S2-", "S²⁻"), ("F-)", "F⁻)"), ("Cl-)", "Cl⁻)"), ("PCB-", "PCB ")):
                name = name.replace(a, b)
            m = " ".join(cells[3].split())
            if not m:
                m = last_m
            last_m = m
            for a, b in (("SWEWW", "SMEWW"), ("4500-S2- ", "4500-S²⁻ "), ("4500-S2 ", "4500-S²⁻ "), ("3500-Cr. B", "3500-Cr B")):
                m = m.replace(a, b)
            unit = " ".join(cells[2].split())
            if with_limit:
                lim = cells[5] if len(cells) > 5 else "-"
                out.append((name, unit, m, " ".join(lim.split()).replace(" - ", " – ")))
            else:
                out.append((name, unit, m))
        return out
    return []


FIELD_UNITS = {"Nhiệt độ": "°C", "pH": "-", "Clo dư": "mg/L", "Clo dư tự do": "mg/L", "TDS": "mg/L",
               "DO": "mg/L", "Mùi": "-"}
FIELD_NAMES = {"DO": "Oxy hòa tan (DO)"}
# Gói chỉ tiêu trên PYC -> chỉ tiêu con (tạm, KQTN về thì ghi theo KQTN)
PKG = {
    "QCVN 40:2025/BTNMT - Vi sinh": [("Coliforms", "MPN/100mL", "SMEWW 9221B:2023")],
    "Hoạt độ phóng xạ": [("Tổng hoạt độ phóng xạ α", "Bq/L", PLACEHOLDER),
                         ("Tổng hoạt độ phóng xạ β", "Bq/L", PLACEHOLDER)],
}


def unit_guess(name):
    """Đơn vị tạm cho chỉ tiêu PYC (PYC không ghi đơn vị) — khi KQTN về thì ghi theo KQTN."""
    low = name.lower()
    if "coli" in low:
        return "MPN/100mL"
    if "độ màu" in low or "ðộ màu" in low:
        return "Pt-Co"
    if "độ đục" in low or "ðộ đục" in low:
        return "NTU"
    if "phóng xạ" in low:
        return "Bq/L"
    if "pcb" in low or "bảo vệ thực vật" in low or "voc" in low:
        return "µg/L"
    return "mg/L"


def pyc_params(pyc_list, field=(), packages=None):
    """[(tên, phương pháp)] của PYC -> [(tên, đơn vị, phương pháp)] cho khung chứng thư chưa có KQTN.
    ``field``: chỉ tiêu đo hiện trường (thứ tự TBKQ) đặt đầu bảng; ``packages``: {chuỗi con tên gói PYC: [(tên, đơn vị,
    phương pháp), …]} — gói (vd "… Vi sinh", "Hoạt độ phóng xạ") bung thành các chỉ tiêu con."""
    out = [(FIELD_NAMES.get(k, k), FIELD_UNITS.get(k, "-"), "Cảm quan" if k == "Mùi" else FIELD) for k in field]
    packages = PKG if packages is None else packages
    for name, meth in pyc_list:
        name = name.replace("Ð", "Đ")
        pk = next((v for k, v in (packages or {}).items() if k.lower() in name.lower()), None)
        if pk is not None:
            out += pk
            continue
        out.append((name, unit_guess(name), meth))
    return out


def pyc_sample(code6, key, pyc, time, vi_tri, field=None, tbkq=("pH",), equip=("mettler_s2",), photo=None, **kw):
    """Một mẫu khung (chưa KQTN) với chỉ tiêu theo PYC thử nghiệm (``pyc`` = dict của ``pyc_index``); hàng đo hiện
    trường ``tbkq`` (chỉ những đại lượng có số trong ``field``) đặt đầu bảng; không đo hiện trường -> không thiết bị."""
    field = dict(field or {})
    tb = [k for k in tbkq if k in field]
    return dict(dict(ky_hieu=f"{code6}-{key}", kqtn_key=key, time=time, vi_tri=vi_tri, field=field, field_tbkq=tb,
                     params=pyc_params(pyc[key], field=tb), equip=list(equip) if field else [], photo=photo), **kw)


NSH_STD ="QCVN 01-1:2024/BYT - Quy chuẩn kỹ thuật quốc gia về chất lượng nước sạch sử dụng cho mục đích sinh hoạt"
NSH_DRAFT = os.path.join(WORKSPACE, "Input", "134486", "3. TBKQ - 1 - Q326D134486.docx")


def nsh_base(draft=None):
    """Khung mẫu nước sinh hoạt (10 thông số nhóm A, QCVN 01-1:2024/BYT) — danh sách thông số, phương pháp, mức lấy
    theo khung TBKQ nước sinh hoạt người dùng đã làm (mặc định 134486/1)."""
    draft = draft or NSH_DRAFT
    return dict(kind="nsh", scope_vn="Lấy mẫu và giám định chất lượng nước sinh hoạt", standard_vn=NSH_STD,
                chi_tieu="10 thông số nhóm A nêu tại QCVN 01-1:2024/BYT",
                params=params_from_draft(draft), limits=limits_from_draft(draft),
                field_tbkq=["pH", "Clo dư tự do", "Mùi"],
                limit_head="Ngưỡng giới hạn cho phép ^(1)",
                limit_note="Ngưỡng giới hạn tối đa cho phép quy định tại QCVN 01-1:2024/BYT.",
                intro="So sánh kết quả thử nghiệm các thông số của mẫu nước nêu trên so với mức quy định tương ứng tại "
                      "QCVN 01-1:2024/BYT:")


def limits_from_draft(path):
    """{tên thông số: mức} từ cột thứ 6 của bảng kết quả khung TBKQ (chứng thư cũ cùng khách) — suy luận, cần hỏi."""
    return {n: lim for n, _u, _m, lim in params_from_draft(path, with_limit=True)}


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def make_dossier(cfg, sample=None, sub_no=None, multi=True):
    sd = (sample or {}).get("date")                 # mẫu lấy ở một ngày riêng (hồ sơ nhiều đợt, vd 138952)
    d0 = sd or cfg["dates"][0]
    dates = [sd] if sd else cfg["dates"]
    loc =(sample or {}).get("site_vn") or (cfg.get("site_bckq") if sample is None else None) or cfg["site_vn"]
    equip = (sample or {}).get("equip", cfg["equip"]) if sample else cfg["equip"]
    return Dossier(
        code=f"Q326D{cfg['code6']}", cert_date=cfg.get("cert_date") or PLACEHOLDER, sample_date=d0,
        reg_date=cfg["reg_date"], inspection_dates=dates if len(dates) > 1 else [],
        sample_time_vn=cfg["time_bckq"] if sample is None else f"Ngày {d0}", sample_time_en=f"From {d0}",
        sub_no=sub_no, is_multi_cert=multi, dossier_type="cleanroom",
        scope_vn=(sample or {}).get("scope_vn") or cfg["scope_vn"],
        scope_en=(sample or {}).get("scope_en") or cfg.get("scope_en") or (sample or {}).get("scope_vn") or cfg["scope_vn"],
        location_desc_vn=loc, location_desc_en=(sample or {}).get("site_en") or cfg.get("site_en") or loc,
        procedure_vn=(sample or {}).get("procedure") or cfg["procedure"],
        procedure_en=(sample or {}).get("procedure") or cfg["procedure"],
        standards_vn=(sample or {}).get("standard_vn") or "-", standards_en=(sample or {}).get("standard_en") or (sample or {}).get("standard_vn") or "-",
        client=ClientInfo(name_vn=cfg["client_vn"], name_en=cfg.get("client_en", cfg["client_vn"]),
                          address_vn=cfg["address_vn"], address_en=cfg.get("address_en", cfg["address_vn"])),
        personnel=PersonnelInfo(executor=cfg["executor"]),
        equipments=[eq(k) for k in equip], locations=[], photos=[], results_data={},
        input_dir=os.path.join(WORKSPACE, "Input", cfg["code6"]), output_dir="", network_dir="",
    )


def _set_hl(r_el):
    from docx.oxml import OxmlElement
    rpr = r_el.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r_el.insert(0, rpr)
    if rpr.find(qn("w:highlight")) is None:
        h = OxmlElement("w:highlight")
        h.set(qn("w:val"), "yellow")
        rpr.append(h)


def highlight_text(doc, needles, headers=False):
    """Tô vàng các run chứa (một phần) chuỗi cần người dùng soát lại — CONFIG ``highlight`` (người dùng yêu cầu
    16/11/2026: chỗ AI theo nguồn chưa chốt thì tô vàng để soát). ``headers``: xét cả header các section."""
    needles = [n for n in needles if n]
    if not needles:
        return
    roots = [doc.element.body]
    if headers:
        for sec in doc.sections:
            for part in (sec.header, sec.first_page_header, sec.even_page_header):
                if part is not None and not part.is_linked_to_previous:
                    roots.append(part._element)
    for p_el in (p for root in roots for p in root.iter(qn("w:p"))):
        runs = [r for r in p_el.iter(qn("w:r"))]
        texts = ["".join(x.text or "" for x in r.iter(qn("w:t"))) for r in runs]
        full = "".join(texts)
        for nd in needles:
            start = 0
            while True:
                k = full.find(nd, start)
                if k < 0:
                    break
                pos = 0
                for r, t in zip(runs, texts):
                    if t and pos < k + len(nd) and pos + len(t) > k:
                        _set_hl(r)
                    pos += len(t)
                start = k + len(nd)


def _keep_proc_placeholder(doc, procedure):
    """Quy trình còn để trống ('………', người dùng chọn 16/11/2026) — bộ chuẩn hóa quy trình của Core bỏ mất '………'
    -> ghi lại vào ô Quy trình của Table 0 (sau đó _highlight_ph tô vàng)."""
    if PLACEHOLDER not in (procedure or ""):
        return
    for row in doc.tables[0].rows:
        cells = row.cells
        if cells[0].text.strip().startswith("Quy trình") and PLACEHOLDER not in cells[-1].text:
            p = cells[-1].paragraphs[-1]
            rpr = p.runs[0]._r.find(qn("w:rPr")) if p.runs else None
            r = p.add_run(", " + PLACEHOLDER if p.text.strip() else PLACEHOLDER)
            if rpr is not None:                     # cùng font / cỡ chữ với phần trước
                r._r.insert(0, copy.deepcopy(rpr))


def _highlight_ph(doc):
    """Tô vàng mọi run còn '………' (AGENTS mục 13: thiếu dữ liệu)."""
    for r_el in doc.element.body.iter(qn("w:r")):
        t = "".join(x.text or "" for x in r_el.iter(qn("w:t")))
        if PLACEHOLDER in t:
            rpr = r_el.find(qn("w:rPr"))
            if rpr is None:
                from docx.oxml import OxmlElement
                rpr = OxmlElement("w:rPr")
                r_el.insert(0, rpr)
            if rpr.find(qn("w:highlight")) is None:
                from docx.oxml import OxmlElement
                h = OxmlElement("w:highlight")
                h.set(qn("w:val"), "yellow")
                rpr.append(h)


# ---------------------------------------------------------------------------
# TBKQ một mẫu
# ---------------------------------------------------------------------------
def build_tbkq(cfg, n, sample, rows, is_en=False):
    """``is_en``: bản tiếng Anh (CONFIG ``en=True``; tên chỉ tiêu EN theo ``names_en``, câu chữ EN theo khung
    người dùng 131229 — người dùng yêu cầu 17/11/2026)."""
    d = make_dossier(cfg, sample, sub_no=n, multi=len(cfg["samples"]) > 1)
    d.ensure_output_directories()
    src = CleanroomHandler(d).build_tbkq(is_en=is_en)
    doc = Document(src)
    kind = sample.get("kind", cfg.get("kind"))
    L = (lambda vn, en: en if is_en else vn)
    fix_cover(doc, is_en, header_colon=kind in ("nu", "nsh"))
    if kind in ("nu", "nsh") and not is_en:                       # header nước BYT 'Ngày: ' cả khi ngày còn '………' (AGENTS mục 2)
        for sec in doc.sections:
            for part in (sec.header, sec.first_page_header):
                for t_el in part._element.iter(qn("w:t")):
                    if t_el.text and re.search(r"Ngày\s*" + PLACEHOLDER, t_el.text):
                        t_el.text = re.sub(r"Ngày\s*" + PLACEHOLDER, "Ngày: " + PLACEHOLDER, t_el.text)
                    elif t_el.text and t_el.text.rstrip().endswith("Ngày") and "Ngày:" not in t_el.text:
                        t_el.text = t_el.text.rstrip() + ": "
    _keep_proc_placeholder(doc, (sample or {}).get("procedure") or cfg["procedure"])
    for row in doc.tables[0].rows:                  # không có mức -> Căn cứ "-" (AGENTS mục 8C nước thải)
        cells = row.cells
        if (cells[0].text.strip().startswith(("Căn cứ", "Applicable")) and not cells[-1].text.strip()
                and ((sample or {}).get("standard_vn") or "-").strip() in ("-", "")):   # nhiều căn cứ: danh sách ở dòng dưới
            ps = cells[-1].paragraphs
            r = ps[0].runs[0] if ps[0].runs else ps[0].add_run()
            r.text = "-"
            r.font.size = Pt(13)
            r.font.name = "Times New Roman"
    trs = doc.tables[0]._tbl.findall(qn("w:tr"))    # Căn cứ "-": bỏ hàng danh sách căn cứ chỉ còn "-" (in lẻ dưới bìa)
    for i, tr in enumerate(trs[:-1]):
        tcs = tr.findall(qn("w:tc"))
        if tcs and "".join(x.text or "" for x in tcs[0].iter(qn("w:t"))).strip().startswith(("Căn cứ", "Applicable")):
            nxt = trs[i + 1].findall(qn("w:tc"))
            if len(nxt) == 1 and "".join(x.text or "" for x in nxt[0].iter(qn("w:t"))).strip() in ("-", ""):
                trs[i + 1].getparent().remove(trs[i + 1])
            break
    if cfg.get("hl_can_cu") or sample.get("hl_can_cu"):   # Căn cứ chưa chốt -> tô vàng ô giá trị (138139, 17/11/2026)
        for row in doc.tables[0].rows:
            if row.cells[0].text.strip().startswith(("Căn cứ", "Applicable")):
                for r_el in row.cells[-1]._tc.iter(qn("w:r")):
                    _set_hl(r_el)
    text_proto, note_proto, head_proto = appendix_protos(doc)
    b = Body(doc, text_proto)
    head = make_head(b, head_proto)

    head(L("Nội dung thực hiện", "Performance content"))
    if sample.get("gui_mau"):                       # mẫu khách gửi (vd 145103) — khuôn theo khung người dùng
        info_table(b, [("Tên mẫu theo khai báo", sample["ten_mau"]), ("Số lượng mẫu", sample.get("so_luong", "01")),
                       ("Ngày nhận mẫu", cfg["ngay_nhan_mau"]), ("Người gửi mẫu", cfg["nguoi_gui"]),
                       ("Mô tả tình trạng mẫu", sample["mo_ta"]),
                       ("Chỉ tiêu yêu cầu thử nghiệm", sample.get("chi_tieu") or chi_tieu_text(rows))])
        if sample.get("photo"):
            b.para("- Hình ảnh mẫu gửi:", keep_next=True, space_before=6)
            b.photo(sample["photo"])
        return _tbkq_results(cfg, n, sample, rows, d, src, doc, b, head, kind, note_proto, is_en)
    b.para(L(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm {'được chỉ định bởi' if cfg.get('chi_dinh') else 'theo yêu cầu của'} "
             f"{cfg['client_vn']}, cụ thể như sau:",          # chi_dinh: theo form 120476 (người dùng 19/11/2026, 138689)
             f"- Sampling location and parameters were indicated by {cfg.get('client_en', cfg['client_vn'])}, "
             f"detailed as follows:"), keep_next=True)
    members = sample.get("members") or [sample]   # chứng thư gộp nhiều mẫu (vd 141523 nước thải NT1 – NT3)
    if sample.get("members") and sample.get("member_times"):   # cùng vị trí, khác giờ (vd 138139): gộp ô vị trí
        t = b.table(1 + len(members), 4, [30, 14, 18, 38])
        for j, h in enumerate(L(["Vị trí lấy mẫu", "Ký hiệu mẫu", "Thời điểm lấy mẫu", "Chỉ tiêu thử nghiệm"],
                                ["Sampling location", "Sampling code", "Sampling time", "Parameter"])):
            cell_text(t.rows[0].cells[j], h)
        ct = (sample.get("chi_tieu_en") if is_en else None) or sample.get("chi_tieu") or chi_tieu_text(rows)
        for i, m in enumerate(members, start=1):
            cell_text(t.cell(i, 1), m["ky_hieu"])
            cell_text(t.cell(i, 2), m["time"])
        t.cell(1, 0).merge(t.cell(len(members), 0))
        cell_text(t.cell(1, 0), (sample.get("vi_tri_en") if is_en else None) or sample["vi_tri"])
        t.cell(1, 3).merge(t.cell(len(members), 3))
        cell_text(t.cell(1, 3), ct, align="left" if len(ct) > 40 else "center")
        _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:-1])
        members = []                                # đã ghi xong bảng vị trí
    else:
        t = b.table(1 + len(members), 3, [32, 18, 50])
    if members:
        for j, h in enumerate(L(["Vị trí lấy mẫu", "Ký hiệu mẫu", "Chỉ tiêu thử nghiệm"],
                                ["Sampling location", "Sampling code", "Parameter"])):
            cell_text(t.rows[0].cells[j], h)
    ct = (sample.get("chi_tieu_en") if is_en else None) or sample.get("chi_tieu") or chi_tieu_text(rows)
    for i, m in enumerate(members, start=1):
        cell_text(t.rows[i].cells[0], (m.get("vi_tri_en") if is_en else None) or m["vi_tri"])
        cell_text(t.rows[i].cells[1], m["ky_hieu"])
        cell_text(t.rows[i].cells[2], ct, align="left" if len(ct) > 40 else "center")
    if len(members) > 1:
        merge_same_vertical(t, 2, 1)
    if members:
        _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:1])
    members = sample.get("members") or [sample]
    if sample.get("seal"):
        b.para(L("- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong:",
                 "- Information on the sampling process and photo of sealed sample:"), keep_next=True, space_before=6)
        t = b.table(2, 4, [40, 18, 20, 22])
        for j, h in enumerate(L(["Người chứng kiến", "Ký hiệu mẫu", "Thời điểm lấy mẫu", "Số niêm phong"],
                                ["Witness", "Sampling code", "Sampling time", "Sealing number"])):
            cell_text(t.rows[0].cells[j], h)
        for j, v in enumerate([L(cfg["witness_vn"], cfg.get("witness_en", cfg["witness_vn"])), sample["ky_hieu"], sample["time"], sample["seal"]]):
            cell_text(t.rows[1].cells[j], v)
        _keep_rows(t._tbl, t._tbl.findall(qn("w:tr")))
        if sample.get("seal_photo"):
            b.photo(sample["seal_photo"])
    else:
        b.para(L("- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong: -",
                 "- Information on the sampling process and photo of sealed sample: -"), space_before=6)
    if sample.get("members"):
        photos = [m["photo"] for m in members if m.get("photo")]
        if photos:
            b.para(L("- Hình ảnh vị trí lấy mẫu:", "- Photo of sampling location:"), keep_next=True, space_before=6)
            photo_grid(b, doc, photos, width_cm=sample.get("photo_cm", 7.5))
    elif sample.get("photo"):
        b.para(L("- Hình ảnh vị trí lấy mẫu:", "- Photo of sampling location:"), keep_next=True, space_before=6)
        b.photo(sample["photo"], width_cm=sample.get("photo_cm", 7.0))   # cỡ ảnh để bỏ trang thừa (135226)
    elif sample.get("photo_note"):                  # ảnh chưa đặt tên -> để trống, ghi chú (người dùng chốt 16/11/2026)
        b.para(f"- Hình ảnh vị trí lấy mẫu: {sample['photo_note']}", space_before=6)
    return _tbkq_results(cfg, n, sample, rows, d, src, doc, b, head, kind, note_proto, is_en)


def info_table(b, pairs, widths=(36, 3, 61)):
    """Bảng 'nhãn | : | giá trị' không viền (mẫu khách gửi, 9.1)."""
    t = b.table(len(pairs), 3, list(widths), header_rows=0, row_height=0)
    tb = t._tbl.tblPr
    for old in tb.findall(qn("w:tblBorders")):
        tb.remove(old)
    for tcb in list(t._tbl.iter(qn("w:tcBorders"))):
        tcb.getparent().remove(tcb)
    from docx.oxml import OxmlElement
    brd = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "none")
        brd.append(e)
    tb.append(brd)
    for row, (k, v) in zip(t.rows, pairs):
        cell_text(row.cells[0], k, align="left")
        cell_text(row.cells[1], ":")
        cell_text(row.cells[2], v, align="left")
        for c in row.cells:
            c.vertical_alignment = M.WD_CELL_VERTICAL_ALIGNMENT.TOP
    return t


EN_LIMIT = {"Không phát hiện": "Not detected", "Không có mùi lạ": "No strange odor"}
EN_METHOD = {FIELD: "On-site measurement", "Cảm quan tại hiện trường": "On-site sensory", "Cảm quan": "Sensory"}


def _tbkq_results(cfg, n, sample, rows, d, src, doc, b, head, kind, note_proto, is_en=False):
    members = sample.get("members") or [sample]
    L = (lambda vn, en: en if is_en else vn)
    num = (lambda x: en_num(x) if is_en and x else x)
    names_en = dict(cfg.get("names_en", {}), **sample.get("names_en", {}))
    head(L("Kết quả giám định", "Inspection result"))
    word = L("Thông số" if sample.get("word", cfg.get("word", "thông số")) == "thông số" else "Chỉ tiêu", "Parameter")
    has_lim = bool(sample.get("limits"))
    has_lim2 = has_lim and bool(sample.get("limits2"))   # so 2 quy chuẩn cùng lúc (141523: QCVN 40 + QCVN 12)
    sample2 = dict(sample, limits=sample.get("limits2") or {})
    hdr = L(["STT", word, "Đơn vị", "Phương pháp", "Kết quả"], ["No.", word, "Unit", "Method", "Result"]) + (
        [L(sample.get("limit_head", "Giới hạn tối đa ^(1)"), sample.get("limit_head_en", "Maximum limit ^(1)"))]
        if has_lim else []) + ([sample.get("limit_head2", "Giới hạn tối đa ^(2)")] if has_lim2 else [])
    widths = ([7, 23, 13, 20, 15, 22] if kind == "nsh" else [7, 31, 12, 22, 13, 15]) if has_lim else [8, 40, 13, 24, 15]
    if has_lim2:
        widths = [6, 26, 9, 20, 13, 13, 13]
    if sample.get("col_widths"):                    # nới cột riêng một mẫu để bỏ trang thừa (135226)
        widths = sample["col_widths"]
    if sample.get("headers"):
        hdr, widths = sample["headers"], sample["widths"]
    if sample.get("members"):                       # mỗi mẫu một cột kết quả
        many = len(members) > 3                     # từ 4 mẫu: tiêu đề chỉ ký hiệu, nới cột kết quả (138139)
        hdr = ["STT", word, "Đơn vị", "Phương pháp"] + [(m['ky_hieu'] if many else
                                                         f"Kết quả {m['ky_hieu'].split('-')[-1].strip()}")
                                                        for m in members]
        widths = ([6, 27, 8, 16] if many else [7, 30, 11, 20])
        widths = widths + [(100 - sum(widths)) / len(members)] * len(members)
    t = b.table(1 + len(rows), len(hdr), widths)
    for j, h in enumerate(hdr):
        cell_text(t.rows[0].cells[j], h)
    loq_mark = "(3)" if has_lim2 else "(2)" if has_lim else "(1)"
    for i, r in enumerate(rows, start=1):
        c = t.rows[i].cells
        res = r["result"]
        if r["loq"] and not r["group"]:
            res = f"{res} ^{loq_mark}"
        lim = lim_of(sample, r["name"]) if has_lim else None
        ok = evaluate(r["result"], lim) if has_lim and not r["group"] and PLACEHOLDER not in r["result"] else None
        lim2 = lim_of(sample2, r["name"]) if has_lim2 else None
        if has_lim2 and not r["group"] and PLACEHOLDER not in r["result"] and evaluate(r["result"], lim2) is False:
            ok = False                              # không đạt theo một trong hai quy chuẩn -> in đậm
        cols = sample.get("cols") or (["name", "unit", "method", "result"] + (["limit"] if has_lim else [])
                                  + (["limit2"] if has_lim2 else []))
        vals = {"name": r["name"], "unit": r["unit"], "method": r["method"], "result": res, "limit": lim,
                "limit2": lim2}
        if is_en:
            vals.update(name=names_en.get(r["name"], r["name"]), method=EN_METHOD.get(r["method"], r["method"]),
                        result=num(res), limit=EN_LIMIT.get(lim, num(lim)) if lim else lim)
        if sample.get("members"):
            cols = ["name", "unit", "method"] + [f"res{k}" for k in range(len(members))]
            vals.update({f"res{k}": (num(v) or "").replace("< ", "< ")   # "< 0,5" không gãy dòng
                         for k, v in enumerate(r["results"])})
        cell_text(c[0], str(i))
        for j, key in enumerate(cols, start=1):
            cell_text(c[j], vals[key], align="left" if key == "name" else "center",
                      bold=(key == "result" and ok is False))
    cols = sample.get("cols") or (["name", "unit", "method", "result"] + (["limit"] if has_lim else [])
                                  + (["limit2"] if has_lim2 else []))
    if not any(r["method"] == PLACEHOLDER for r in rows):   # phương pháp còn "………" (chưa KQTN): không gộp dọc —
        merge_same_vertical(t, cols.index("method") + 1, 1)  # ô gộp cả bảng không tách trang được (138167)
    else:                                           # chưa KQTN: vẫn gộp các ô phương pháp giống nhau, khác "………"
        i = 0                                       # (vd "Đo tại hiện trường" pH + clo dư — người dùng 19/11/2026, 138689)
        while i < len(rows):
            j = i
            while j + 1 < len(rows) and rows[j + 1]["method"] == rows[i]["method"]:
                j += 1
            if j > i and rows[i]["method"] != PLACEHOLDER:
                merge_same_vertical(t, cols.index("method") + 1, i + 1, j + 1)
            i = j + 1
    if has_lim:
        merge_text_runs(t, cols.index("limit") + 1, 1)
        if has_lim2:
            merge_text_runs(t, cols.index("limit2") + 1, 1)
        if sample.get("hl_limits"):                 # mức lấy theo hồ sơ cũ — tô vàng để người dùng soát
            hl = sample["hl_limits"]                # True: mọi ô mức; list: chỉ các thông số nêu tên
            for row, rr in zip(t.rows[1:], rows):
                if isinstance(hl, (list, tuple)) and rr["name"] not in hl:
                    continue
                for r_el in row.cells[cols.index("limit") + 1]._tc.iter(qn("w:r")):
                    _set_hl(r_el)
    if sample.get("hl_field") and not sample.get("members"):   # số đo hiện trường cần người dùng soát -> tô vàng
        for row, rr in zip(t.rows[1:], rows):
            if rr["method"].startswith(("Đo tại hiện trường", "Cảm quan")) or rr["method"] == FIELD:
                for r_el in row.cells[cols.index("result") + 1]._tc.iter(qn("w:r")):
                    _set_hl(r_el)
    for row in t.rows[1:]:
        for cc in row.cells:
            for p in cc.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(13)
    trs = t._tbl.findall(qn("w:tr"))
    mcol = cols.index("method") + 1                 # hàng chung ô phương pháp gộp dọc không tách trang
    for i in range(1, len(trs) - 1):
        nxt = trs[i + 1].findall(qn("w:tc"))
        vm = nxt[mcol].find(qn("w:tcPr") + "/" + qn("w:vMerge")) if len(nxt) > mcol else None
        if vm is not None and vm.get(qn("w:val")) in (None, "continue"):
            _keep_rows(t._tbl, [trs[i]])
    _keep_rows(t._tbl, trs[:5])
    _keep_rows(t._tbl, trs[-2:])
    byt = kind in ("nu", "nsh")
    notes = [] if byt else [("●", L("Các thông số được chứng nhận VIMCERTS 078.",
                                     "Parameters are accredited to VIMCERTS 078."))]
    if has_lim:
        notes.append(("(1)", L(sample["limit_note"], sample.get("limit_note_en", sample["limit_note"]))))
    if has_lim2:
        notes.append(("(2)", sample["limit_note2"]))
    if any(r["loq"] for r in rows):
        notes.append((loq_mark, L("Giới hạn định lượng của phương pháp thử.",
                                  "Limit of quantification of test method.")))
    if byt:
        notes += [("●", L("Phương pháp thử nghiệm được công nhận ISO 17025 - VILAS 004.",
                          "The test method is accredited to ISO 17025 - VILAS 004.")),
                  ("●", L("Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.",
                          "The performance field is accredited to ISO 17020 - VIAS 004."))]
    nt = note_table(b, note_proto, notes)
    if has_lim:
        _keep_rows(nt, nt.findall(qn("w:tr")))
        if is_en:
            M.assessment(b, sample.get("intro_en", sample["intro"]), eval_lines_en(rows, sample, names_en),
                         is_en=True)
        else:
            M.assessment(b, sample["intro"], eval_lines(rows, sample))
            if has_lim2:                            # đánh giá riêng theo quy chuẩn thứ hai
                b.para(sample["intro2"], keep_next=True, space_before=3)
                lines2 = eval_lines(rows, sample2)
                for k, (pre, key, val) in enumerate(lines2):
                    q = b.para(parts=[("+ " + (f"{pre} " if pre else ""), False, False), (key, True, True),
                                      ("\t: " + val, False, False)],
                               indent_cm=0.75, tab_cm=7.0, keep_next=k < len(lines2) - 1)
                    _keep(q._p, lines=True)
    else:
        _keep_rows(nt, nt.findall(qn("w:tr"))[:-1])
    if sample.get("after_results"):                 # phần thêm sau bảng kết quả (vd đo ồn 138167)
        sample["after_results"](b, doc, note_proto, is_en)
    style_note_tables(doc)
    _highlight_ph(doc)
    highlight_text(doc, L(list(cfg.get("highlight", [])) + list(sample.get("highlight", [])),
                          list(cfg.get("highlight_en", [])) + list(sample.get("highlight_en", []))))
    if sample.get("_hl_cert") and n:                # số chứng thư theo số mẫu, nhảy cóc -> tô vàng (16/11/2026)
        highlight_text(doc, [f"{d.code}/{n}"], headers=True)
    add_document_ending_line(doc)
    appendix_total_field(doc)
    mid = f"{n} - " if n else ""
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {L('viet', 'en')} - {mid}{sample.get('loai', cfg['loai'])} - {d.code}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    print(f"[nuoc_common] TBKQ {L('VN', 'EN')} /{n} {sample['ky_hieu']} -> {os.path.relpath(dst, WORKSPACE)}")
    return dst


def photo_grid(b, doc, photos, width_cm=7.5):
    """Lưới ảnh 2 cột (AGENTS mục 3: 7 – 8 cm, không crop); 1 ảnh -> bảng 1 cột, ảnh giữa trang (138139, 20/11/2026)."""
    cols = 1 if len(photos) == 1 else 2
    pt = doc.add_table(rows=(len(photos) + cols - 1) // cols, cols=cols)
    b.emit(pt._tbl)
    set_table_autofit_window(pt)
    for k, ph in enumerate(photos):
        insert_image_to_cell(pt.rows[k // cols].cells[k % cols], ph, width=M.Cm(width_cm))
    for r in pt.rows:
        add_cant_split(r)


def member_rows(sample, kqs):
    """Hàng bảng kết quả chứng thư gộp nhiều mẫu khi đã có KQTN từng mẫu: đo hiện trường (``field_tbkq`` của chứng
    thư) rồi chỉ tiêu KQTN theo thứ tự phiếu mẫu đầu; mỗi mẫu một giá trị trong ``results``."""
    per = []
    for m, kq in zip(sample["members"], kqs):
        mm = dict(m, field_tbkq=m.get("field_tbkq", sample.get("field_tbkq", [])))
        per.append({r["name"]: r for r in field_rows(mm) + kqtn_rows(kq)})
    out = []
    for name, r0 in per[0].items():
        vals = [p.get(name, {}).get("result", "-") for p in per]
        out.append(dict(r0, result=vals[0], results=vals, loq=any(p.get(name, {}).get("loq") for p in per)))
    return out


def group_rows(sample):
    """Hàng bảng kết quả của chứng thư gộp nhiều mẫu (chưa KQTN): đo hiện trường lấy số từng mẫu, còn lại '………'."""
    out = []
    for n, u, m in sample["params"]:
        if m.startswith(("Đo tại hiện trường", "Cảm quan")):
            vals = [mem.get("field", {}).get(_field_key(n), "-") for mem in sample["members"]]
        else:
            vals = [PLACEHOLDER] * len(sample["members"])
            m = PLACEHOLDER                         # chưa KQTN: phương pháp để trống (AGENTS mục 13, 17/11/2026)
        out.append(dict(name=n, unit=unit_fmt(u), method=m, result=vals[0], results=vals, loq=False, vim=True,
                        group=False))
    return out


def lim_of(sample, name):
    lims = sample.get("limits") or {}
    if name in lims:
        return lims[name]
    sn = short_name(name)
    for k, v in lims.items():
        if short_name(k) == sn:
            return v
    return "-"


def eval_lines_en(rows, sample, names_en):
    """Đánh giá bản EN (câu chữ theo khung người dùng 131229)."""
    if any(PLACEHOLDER in r["result"] for r in rows):
        return [("", "Non-compliant parameter", PLACEHOLDER), ("", "Compliant parameter", PLACEHOLDER)]
    no_eval, fail = [], []
    for r in rows:
        if r["group"]:
            continue
        v = evaluate(r["result"], lim_of(sample, r["name"]))
        if v is None:
            no_eval.append(names_en.get(r["name"], r["name"]))
        elif not v:
            fail.append(names_en.get(r["name"], r["name"]))
    lines = []
    if no_eval:
        lines.append(("", "Non-evaluated parameter", ", ".join(no_eval)))
    lines.append(("", "Non-compliant parameter", ", ".join(fail) or "-"))
    lines.append(("", "Compliant parameter", "all parameters" if not (fail or no_eval) else "the remaining parameters"))
    return lines


def eval_lines(rows, sample):
    w, bad, good = sample.get("eval_words", ("Thông số", "không đạt", "đạt"))
    if any(PLACEHOLDER in r["result"] for r in rows):
        return [(w, bad, PLACEHOLDER), (w, good, PLACEHOLDER)]
    no_eval, fail, ok = [], [], []
    for r in rows:
        if r["group"]:
            continue
        v = evaluate(r["result"], lim_of(sample, r["name"]))
        (no_eval if v is None else ok if v else fail).append(short_name(r["name"]))
    lines = []
    if no_eval:
        lines.append((w, "không đánh giá", ", ".join(no_eval)))
    lines.append((w, bad, ", ".join(fail) or "-"))
    lines.append((w, good, f"tất cả các {w.lower()}" if not (fail or no_eval) else f"các {w.lower()} còn lại"))
    return lines


def chi_tieu_text(rows):
    seen = []
    for r in rows:
        s = short_name(r["name"].split("\n")[0])
        if s not in seen:
            seen.append(s)
    txt = ", ".join(seen)
    return txt if txt.startswith("pH") else txt[:1].upper() + txt[1:]


# ---------------------------------------------------------------------------
# BCKQ gộp
# ---------------------------------------------------------------------------
def build_bckq(cfg, all_rows, kqs, qc):
    d = make_dossier(cfg, None, sub_no=None, multi=False)
    d.equipments = [eq(k) for k in cfg["equip"]]
    if cfg.get("scope_bckq"):
        d.scope_vn = d.scope_en = cfg["scope_bckq"]
    if cfg.get("procedure_bckq"):
        d.procedure_vn = d.procedure_en = cfg["procedure_bckq"]
    out = BCKQBuilder(d).build(out_name=f"BCKQ - {cfg['loai']} - {d.code}.docx")
    doc = Document(out)
    for row in doc.tables[0].rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs:
                    if PLACEHOLDER in r.text and "thiết bị" in row.cells[0].text.lower():
                        r.text = r.text.replace(PLACEHOLDER, "-")
    _keep_proc_placeholder(doc, cfg.get("procedure_bckq") or cfg["procedure"])
    body_el = doc.element.body
    kids = list(body_el.iterchildren())
    h1 = next(el for el in kids if el.tag == qn("w:p") and M._ptext(el).strip().startswith("Nội dung thực hiện"))
    h2 = next(el for el in kids if el.tag == qn("w:p") and M._ptext(el).strip().startswith("Kết quả đo đạc"))
    gdv = next(el for el in kids if el.tag == qn("w:p") and M._ptext(el).strip().startswith("Báo cáo của GĐV"))
    text_proto = copy.deepcopy(h1.getnext())
    i1, i2, ig = kids.index(h1), kids.index(h2), kids.index(gdv)
    for el in kids[i1 + 1:i2] + kids[i2 + 1:ig]:
        body_el.remove(el)
    b = Body(doc, text_proto)
    samples = cfg["samples"]

    b.anchor = h2
    if cfg.get("gui_mau"):                          # mẫu khách gửi (vd 145103) — khuôn BCKQ theo khung người dùng
        for line in (f"Số lượng mẫu\t: {len(samples):02d}", f"Ngày nhận mẫu\t: {cfg['ngay_nhan_mau']}",
                     f"Người gửi mẫu\t: {cfg['nguoi_gui']}"):
            b.para(line, tab_cm=5.5)
        b.para("Tên mẫu, tình trạng mẫu và các chỉ tiêu thực hiện cụ thể như sau:", keep_next=True, space_before=3)
        t = b.table(1 + len(samples), 3, [28, 44, 28])
        for j, h in enumerate(["Tên mẫu theo khai báo", "Mô tả tình trạng mẫu", "Chỉ tiêu"]):
            cell_text(t.rows[0].cells[j], h)
        for i, s in enumerate(samples, start=1):
            cell_text(t.rows[i].cells[0], s["ten_mau"], align="left")
            cell_text(t.rows[i].cells[1], s["mo_ta"], align="left")
            cell_text(t.rows[i].cells[2], s.get("chi_tieu") or chi_tieu_text(all_rows[i - 1]), align="left")
        for extra in cfg.get("bckq_notes", []):
            b.para(extra, space_before=6)
        photos = [s["photo"] for s in samples if s.get("photo")]
        if photos:
            b.para("Hình ảnh mẫu gửi:", keep_next=True, space_before=6)
            photo_grid(b, doc, photos)
    else:
        _bckq_site_block(cfg, samples, all_rows, b, doc)
    _bckq_results(cfg, samples, all_rows, kqs, qc, b, doc, h2, gdv, out)
    return out


def _bckq_site_block(cfg, samples, all_rows, b, doc):
    if True:
        b.para(f"- Vị trí lấy mẫu và chỉ tiêu thử nghiệm được chỉ định bởi {cfg['witness_vn']}, cụ thể như sau:",
                   keep_next=True)
        t = b.table(1 + len(samples), 4, [33, 15, 12, 40])
        for j, h in enumerate(["Vị trí thực hiện", "Kí hiệu mẫu", "Thời điểm lấy mẫu", "Chỉ tiêu thử nghiệm"]):
            cell_text(t.rows[0].cells[j], h)
        for i, s in enumerate(samples, start=1):
            cell_text(t.rows[i].cells[0], s["vi_tri"], align="left")
            cell_text(t.rows[i].cells[1], s["ky_hieu"])
            cell_text(t.rows[i].cells[2], s.get("time", "-"))
            cell_text(t.rows[i].cells[3], s.get("chi_tieu") or chi_tieu_text(all_rows[i - 1]), align="left")
        for extra in cfg.get("bckq_notes", []):
            b.para(extra, space_before=6)
        b.para("- Thông tin về quá trình lấy mẫu niêm phong: -", space_before=6)
        photos = [s["photo"] for s in samples if s.get("photo")]
        if photos:
            b.para("- Hình ảnh vị trí lấy mẫu:", keep_next=True, space_before=6)
            ncol = 1 if len(photos) == 1 else 2         # 1 ảnh: bảng 1 cột -> ảnh giữa trang (người dùng 19/11/2026)
            nrow = (len(photos) + ncol - 1) // ncol
            pt = doc.add_table(rows=nrow, cols=ncol)
            b.emit(pt._tbl)
            set_table_autofit_window(pt)
            for k, ph in enumerate(photos):
                insert_image_to_cell(pt.rows[k // ncol].cells[k % ncol], ph, width=M.Cm(7.5))
            for r in pt.rows:
                add_cant_split(r)
        nophoto = [s["ky_hieu"] for s in samples if not s.get("photo") and s.get("photo_note")]
        if nophoto:
            b.para(f"- Hình ảnh vị trí lấy mẫu {', '.join(nophoto)}: "
                   f"{next(s['photo_note'] for s in samples if not s.get('photo') and s.get('photo_note'))}",
                   space_before=6)


def _bckq_results(cfg, samples, all_rows, kqs, qc, b, doc, h2, gdv, out):
    b.anchor = gdv
    _keep(h2)
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    for s, kq in zip(samples, kqs):
        so = f"{kq['so_phieu']} ngày {kq['ngay']}" if kq else PLACEHOLDER
        b.para(f"- KQTN mẫu {s['ky_hieu']}: xem tại phiếu KQTN số {so};").alignment = WD_ALIGN_PARAGRAPH.LEFT
    for q in qc:
        qn_ = re.sub(r"^(?i:mẫu)\s+", "", q['ten_mau'].split('-')[-1].strip())   # "Mẫu QC" -> "QC"
        b.para(f"- KQTN mẫu {qn_} (mẫu kiểm soát chất lượng): xem tại phiếu KQTN "
               f"số {q['so_phieu']} ngày {q['ngay']}.").alignment = WD_ALIGN_PARAGRAPH.LEFT
    fcols = [k for k in ("pH", "Nhiệt độ", "Clo dư", "Clo dư tự do", "TDS", "DO", "Mùi") if any(k in s.get("field", {}) for s in samples)]
    if fcols:
        b.para("- Kết quả đo tại hiện trường:", keep_next=True, space_before=3)
        heads = {"pH": "pH", "Nhiệt độ": "Nhiệt độ [°C]", "Clo dư": "Clo dư [mg/L]", "Clo dư tự do": "Clo dư tự do [mg/L]",
                 "TDS": "TDS [mg/L]", "DO": "DO [mg/L]", "Mùi": "Mùi"}
        fs = [s for s in samples if s.get("field")]
        t = b.table(1 + len(fs) + 1, 1 + len(fcols), [20] + [80 / len(fcols)] * len(fcols))
        cell_text(t.rows[0].cells[0], "Kí hiệu mẫu")
        for j, k in enumerate(fcols, start=1):
            cell_text(t.rows[0].cells[j], heads[k])
        for i, s in enumerate(fs, start=1):
            cell_text(t.rows[i].cells[0], s["ky_hieu"])
            for j, k in enumerate(fcols, start=1):
                cell_text(t.rows[i].cells[j], s["field"].get(k, "-"))
                if s.get("hl_field") and k in s["field"]:
                    for r_el in t.rows[i].cells[j]._tc.iter(qn("w:r")):
                        _set_hl(r_el)
        cell_text(t.rows[-1].cells[0], "Thiết bị sử dụng")
        for j, k in enumerate(fcols, start=1):
            dev = cfg.get("field_devices", {}).get(k, "-")
            cell_text(t.rows[-1].cells[j], dev, italic=dev != "-")
        _keep_rows(t._tbl, t._tbl.findall(qn("w:tr"))[:-1])
    ev = [(sm, rw) for sm, rw, kq in zip(samples, all_rows, kqs) if kq and sm.get("limits")]
    if ev and len(ev) == len(samples):              # đủ KQTN + có mức -> Đánh giá theo từng mẫu, khớp TBKQ (AGENTS mục 7)
        b.para("Đánh giá:", keep_next=True, space_before=6)
        for sm, rw in ev:
            b.para(f"- Mẫu {sm['ky_hieu']} so với {sm.get('standard_vn', '').split(' - ')[0] or 'mức quy định'}:",
                   keep_next=True)
            for w, kind_, txt in eval_lines(rw, sm):
                b.para(f"	+ {w} {kind_}	: {txt}", tab_cm=5.5)
    else:
        b.para("Đánh giá: -", keep_next=True, space_before=6)
    from Core.Tools.kqtn_late import tre_han_text
    b.para("Gửi mẫu thử nghiệm:", bold=True, keep_next=True, space_before=6)
    got = [k for k in kqs if k] + list(qc)
    if got:
        roots = sorted({k["so_phieu"].split(".")[0] for k in got})
        subs = ", ".join("." + k["so_phieu"].split(".")[1] for k in sorted(got, key=lambda x: x["so_phieu"])
                         if "." in k["so_phieu"])   # phiếu 1 mẫu không có đuôi .01 (137037)
        kq_date = max(got, key=lambda k: datetime.datetime.strptime(k["ngay"], "%d/%m/%Y"))["ngay"]
        send = min((k["ngay_nhan_mau"] for k in got if k["ngay_nhan_mau"]),
                   key=lambda x: datetime.datetime.strptime(x, "%d/%m/%Y"))
        so = f"{', '.join(roots)} ({subs})" if subs else ", ".join(roots)
    else:
        kq_date = send = so = PLACEHOLDER
    if got and cfg.get("send_date"):                # ngày trên BIÊN BẢN GỬI MẪU (AGENTS mục 13, người dùng chốt 18/11/2026); không có biên bản mới lấy ngày nhận mẫu KQTN
        send = cfg["send_date"]
    req = cfg.get("req_date") or PLACEHOLDER
    for line in (f"Phiếu KQTN số {so} ngày {kq_date}", f"Ngày gửi mẫu\t: {send}",
                 f"Ngày yêu cầu cấp KQTN\t: {req}", f"Ngày nhận KQTN\t: {kq_date}"):
        b.para(line, keep_next=True, tab_cm=5.5)
    if "tre_han" in cfg:                            # người dùng chốt riêng (vd để trống, tô vàng)
        b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {cfg['tre_han']}")
    elif got and PLACEHOLDER not in req:
        b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre_han_text(kq_date, send, req, PLACEHOLDER)}")
    else:
        b.para(f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {PLACEHOLDER}")
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm")):
            p.paragraph_format.keep_with_next = True
    _highlight_ph(doc)
    highlight_text(doc, list(cfg.get("highlight", [])) + [h for s in samples for h in s.get("highlight", [])]
                   + list(cfg.get("highlight_bckq", [])))   # chỉ BCKQ (vd ngày gửi mẫu, trễ hạn — 131229)
    if cfg.get("no_highlight"):                     # người dùng bỏ tô vàng tất cả (135228, 20/11/2026)
        for h in list(doc.element.body.iter(qn("w:highlight"))):
            h.getparent().remove(h)
    doc.save(out)
    print(f"[nuoc_common] BCKQ -> {os.path.relpath(out, WORKSPACE)}")
    return out


# ---------------------------------------------------------------------------
def find_kqtn(cfg):
    """{ký hiệu mẫu (đuôi tên mẫu KQTN): dict}, [mẫu QC / blank]."""
    files = sorted(set(glob.glob(os.path.join(WORKSPACE, "Input", cfg["code6"], "**", "KQ-*.pdf"), recursive=True)))
    by, qc = {}, []
    seen, seen_qc = set(), set()
    for f in files:
        if cfg["code6"] not in os.path.basename(f) or os.path.basename(f) in seen:
            continue                                      # cùng file KQTN chép ở 2 thư mục (134486, 20/11/2026)
        seen.add(os.path.basename(f))
        for k in read_kqtn_all(f):                        # file gộp nhiều phiếu (134486) -> từng phiếu
            if k["so_phieu"] in cfg.get("kqtn_map", {}):    # tên mẫu KQTN không có ký hiệu (vd 138139 "Nước thải")
                by[cfg["kqtn_map"][k["so_phieu"]]] = k
                continue
            tail = k["ten_mau"].split("-")[-1].strip()
            if tail.upper().startswith(("QC", "BLANK", "TB")) or re.search(r"\bQC\b", tail.upper()):   # "MẪU QC"
                if k["so_phieu"] not in seen_qc:
                    seen_qc.add(k["so_phieu"])
                    qc.append(k)
            else:
                by[tail] = k
    return by, qc


def run(cfg, pdf=False, only=None):
    global KPH_LOD
    KPH_LOD = bool(cfg.get("kph_lod"))
    by, qc = find_kqtn(cfg)
    multi = len(cfg["samples"]) > 1
    kqs, all_rows, docs = [], [], []
    dates = []
    for s in cfg["samples"]:
        if "field_tbkq" not in s and (s.get("params") or s.get("draft")):
            ps = s.get("params") or params_from_draft(s["draft"])
            s["field_tbkq"] = [_field_key(n) for n, _u, m in ps if m.startswith(("Đo tại hiện trường", "Cảm quan"))
                               and _field_key(n) in s.get("field", {})]
        if s.get("members"):                        # chứng thư gộp nhiều mẫu: mỗi mẫu một cột kết quả
            mk = [by.get(m.get("kqtn_key", m["ky_hieu"])) for m in s["members"]]
            if all(mk):
                kqs.append(mk)
                all_rows.append(member_rows(s, mk))
                dates += [k["ngay"] for k in mk]
            else:                                   # chưa đủ KQTN -> khung
                kqs.append(None)
                all_rows.append(group_rows(s))
            continue
        kq = by.get(s["kqtn_key"])
        kqs.append(kq)
        rows = (field_rows(s) + kqtn_rows(kq)) if kq else placeholder_rows(s)
        if s.get("order"):                          # thứ tự chỉ tiêu chung cho mọi chứng thư (phiếu KQTN xếp khác nhau)
            pos_ = {nm: i for i, nm in enumerate(s["order"])}
            rows = sorted(rows, key=lambda r: pos_.get(r["name"], len(pos_)))
        all_rows.append(rows)
        if kq:
            dates.append(kq["ngay"])
    if cfg.get("cert_date") is None and dates and all(kqs):
        cfg["cert_date"] = max(dates + [q["ngay"] for q in qc], key=lambda x: datetime.datetime.strptime(x, "%d/%m/%Y"))
    for pos, (s, rows) in enumerate(zip(cfg["samples"], all_rows), start=1):
        # cfg cert_by_key: số chứng thư = số của ký hiệu mẫu (N9 -> /9, bỏ trống số của mẫu không lấy được — người
        # dùng chốt 16/11/2026, 139925 / 140067); số khác thứ tự -> tô vàng
        n = s.get("cert_no") or (int(re.sub(r"\D", "", s["kqtn_key"])) if cfg.get("cert_by_key") else pos)
        s["_hl_cert"] = s.get("hl_cert") or n != pos
        if only and n not in only:
            continue
        docs.append(build_tbkq(cfg, n if multi else None, s, rows))
        if cfg.get("en") or s.get("en"):           # chứng thư Việt - Anh (PYC mục 15 / người dùng chốt)
            docs.append(build_tbkq(cfg, n if multi else None, s, rows, is_en=True))
    fs, fr, fk = [], [], []                         # BCKQ liệt kê từng mẫu (bung chứng thư gộp)
    for s, rows, kq in zip(cfg["samples"], all_rows, kqs):
        for m in (s.get("members") or [s]):
            fs.append(dict(m, chi_tieu=m.get("chi_tieu") or s.get("chi_tieu")) if s.get("members") else m)
            fr.append(rows)
            fk.append(kq[s["members"].index(m)] if isinstance(kq, list) else kq)
    docs.insert(0, build_bckq(dict(cfg, samples=fs), fr, fk, qc))
    if pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        PDFExporter(make_dossier(cfg, None, None, False)).export_files(docs)
    report = []
    for p in docs:
        res = check_single_docx(p, code=f"Q326D{cfg['code6']}")
        report.append((res.get("status"), os.path.relpath(p, WORKSPACE), res.get("issues", [])))
    print("\n===== QA =====")
    for st, p, issues in report:
        print(f"  {st}\t{p}")
        for i in issues:
            print(f"      - {i}")
    return docs
