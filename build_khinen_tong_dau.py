# -*- coding: utf-8 -*-
"""
Bộ dựng chung cho hồ sơ khí nén CHỈ LẤY MẪU TỔNG DẦU (1 vị trí, 1 chứng thư) — dùng cho 139348, 139353, 139363
(cùng nhà máy Phenikaa Hòa Lạc, 12/11/2026). Script hồ sơ (`build_<6 số>.py`) khai `Case` rồi gọi `run(case)`.

- Phiếu tính toán: template TinhToanKhiNen (khinen_calc) giữ công thức; khối bụi (Lasair, AeroTrak), vi sinh, điểm
  sương không đo -> gạch chéo (strike_box, AGENTS 8B); N* lấy nhiệt độ khí ghi trên phiếu ghi chép (đọc từ DP510).
- Chưa có KQTN -> ô KQTN, kết quả tổng dầu, phân loại, số phiếu, ngày ban hành = `………` tô vàng (AGENTS mục 2, 13).
  Có KQTN: điền `kqtn` trong Case (µg/mẫu như phiếu, vd "< 10" / "< 5"; số phiếu, ngày, ngày nhận mẫu) rồi dựng lại.
- TBKQ / BCKQ: CompressedAirHandler + BCKQBuilder (template TAG), chỉ mục Hàm lượng tổng dầu.
"""

import os
import re
import sys
import copy
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Optional

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_ton_tai_Z"))

import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from Core.Models import ClientInfo, PersonnelInfo, assign_cert_numbers
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.compressed_air import CompressedAirHandler
from Core.Builders.tag_context import PLACEHOLDER, style_note_tables
from Core.Styling.tables import add_tbl_header
from Core.Tools.khinen_calc import build_compressed_air_workbook, create_dossier_from_compressed_air_workbooks
from Core.Tools.excel_borders import thin_outside_borders, strike_box
from Core.Tools.kqtn_late import tre_han_text
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq


@dataclass
class Case:
    code6: str
    client: ClientInfo
    location_vn: str                  # Địa điểm giám định (BBLM mục 5, ghi đủ — AGENTS 15)
    location_en: str
    sym: str                          # ký hiệu vị trí theo BBLM (KN1, KN2, KN3)
    name_vn: str                      # tên vị trí theo BBLM / phiếu ghi chép
    name_en: str
    reg_date: str                     # ngày PYC (BBLM mục 1)
    sample_date: str
    start: str                        # BBLM mục 6
    end: str
    oil_from: str                     # phiếu ghi chép: thời gian lấy mẫu
    oil_to: str
    gas_temp: float                   # phiếu ghi chép: nhiệt độ khí lấy mẫu (°C) -> N*
    photo: str                        # ảnh vị trí (người dùng đã đánh dấu ký hiệu)
    scan_new: str                     # scan BBLM tr. 1–2, ghi chép tr. 3 (tr. 4 trắng)
    scan_khac: str                    # scan PKSCV, PYC, điều kiện chung, báo giá
    executor: str = "Phát"
    flow_aerosol: float = 35          # L/phút (phiếu ghi chép)
    flow_vapor: float = 1
    factory_status_vn: str = "Không hoạt động sản xuất (trạng thái nghỉ)"
    factory_status_en: str = "Not in operation"
    kqtn: Optional[dict] = None       # {"aerosol": "< 10", "vapor": "< 5", "so_hat": "HDTN…", "so_hoi": "MTTN…",
    #                                    "ngay": "dd/mm/yyyy", "ngay_nhan_mau": "dd/mm/yyyy"}
    ngay_yc_kqtn: str = ""            # PYC trống "Ngày dự kiến trả kết quả" -> ngày lấy mẫu + 8 (AGENTS 13)
    system_chain_vn: list = field(default_factory=list)
    system_chain_en: list = field(default_factory=list)

    @property
    def code(self):
        return f"Q326D{self.code6}"

    @property
    def input_dir(self):
        return os.path.join(WORKSPACE, "Input", self.code6)

    @property
    def minutes(self):
        t = lambda s: datetime.strptime(s, "%H:%M")
        return int((t(self.oil_to) - t(self.oil_from)).total_seconds() // 60)


def _yc_date(c: Case) -> str:
    if c.ngay_yc_kqtn:
        return c.ngay_yc_kqtn
    d = datetime.strptime(c.sample_date, "%d/%m/%Y") + timedelta(days=8)
    return d.strftime("%d/%m/%Y")


# ---------------------------------------------------------------------------
# Tiện ích DOCX
# ---------------------------------------------------------------------------
def _ptext(el):
    import unicodedata
    return unicodedata.normalize("NFC", "".join(t.text or "" for t in el.iter(qn("w:t"))))


def _keep(p_el):
    pPr = p_el.get_or_add_pPr()
    if pPr.find(qn("w:keepNext")) is None:
        pPr.insert(0, OxmlElement("w:keepNext"))


def _highlight_placeholders(doc):
    roots = [doc.element.body] + [part._element for s in doc.sections for part in
                                  (s.header, s.first_page_header, s.even_page_header) if not part.is_linked_to_previous]
    for r in [r for root in roots for r in root.iter(qn("w:r"))]:
        if not any(PLACEHOLDER in (t.text or "") for t in r.findall(qn("w:t"))):
            continue
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        if rpr.find(qn("w:highlight")) is None:
            h = OxmlElement("w:highlight")
            h.set(qn("w:val"), "yellow")
            rpr.append(h)


def _fix_cover_lists(doc):
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
    for el in doc.element.body.iterchildren(qn("w:tbl")):
        prev = el.getprevious()
        if prev is not None and prev.tag == qn("w:p") and _ptext(prev).strip():
            _keep(prev)
            rows = el.findall(qn("w:tr"))
            for tr in rows[:min(5, len(rows) - 1)]:
                for p_el in tr.iter(qn("w:p")):
                    _keep(p_el)


def _superscript_note_marks(doc):
    from docx.text.run import Run
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        if not _ptext(tbl).strip().startswith(("Ghi chú", "Note")):
            continue
        for r_el in list(tbl.iter(qn("w:r"))):
            ts = r_el.findall(qn("w:t"))
            m = re.match(r"^\s*(\(\d\))(.*)$", ts[0].text or "") if len(ts) == 1 else None
            if not m:
                continue
            run = Run(r_el, None)
            if run.font.superscript:
                continue
            if m.group(2):
                rest = copy.deepcopy(r_el)
                r_el.addnext(rest)
                Run(rest, None).text = m.group(2)
            run.text = m.group(1)
            run.font.superscript = True


def _replace_in_paragraphs(root, mapping: dict) -> int:
    n = 0
    for p in root.iter(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        txt = "".join(t.text or "" for t in ts)
        if not any(a in txt for a in mapping):
            continue
        for a, b in mapping.items():
            while a in txt:
                start = txt.index(a)
                pos, first = 0, None
                for t in ts:
                    L_ = len(t.text or "")
                    if pos + L_ > start:
                        first = t
                        break
                    pos += L_
                new = txt.replace(a, b, 1)
                before = sum(len(t.text or "") for t in ts[:ts.index(first)])
                first.text = new[before:]
                first.set(qn("xml:space"), "preserve")
                for t in ts[ts.index(first) + 1:]:
                    t.text = ""
                txt = new
                n += 1
    return n


def _optimized(path: str, out_dir: str, max_w: int = 1200) -> str:
    from PIL import Image, ImageOps
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".jpg")
    im = ImageOps.exif_transpose(Image.open(path))
    if im.height > im.width:
        im = im.rotate(90, expand=True)
    if im.width > max_w:
        im = im.resize((max_w, round(im.height * max_w / im.width)), Image.Resampling.LANCZOS)
    im.convert("RGB").save(out, "JPEG", quality=88)
    return out


# ---------------------------------------------------------------------------
# Phiếu tính toán
# ---------------------------------------------------------------------------
def build_workbook(c: Case, cert_date: str) -> str:
    k = c.kqtn or {}
    point_data = {
        "code": c.code, "point_label": c.sym, "cert_date": cert_date,
        "location_name": c.name_vn, "location_symbol": c.sym,
        "location_photo": c.photo,
        # điểm sương không đo: chỉ khai nhiệt độ khí (phiếu ghi chép) để N* — khối điểm sương gạch chéo sau
        "dewpoint": {"temp_start": c.gas_temp, "temp_end": c.gas_temp},
        "oil": {"flow_aerosol": c.flow_aerosol, "flow_vapor": c.flow_vapor,
                "time_aerosol": c.minutes, "time_vapor": c.minutes,
                # chưa có KQTN: dựng tạm bằng số của template để đọc được, rồi thay bằng ……… ở finalize
                "kqtn_aerosol": k.get("aerosol", "< 10"), "kqtn_vapor": k.get("vapor", "< 5"),
                "kqtn_no_aerosol": k.get("so_hat", ""), "kqtn_no_vapor": k.get("so_hoi", "")},
    }
    path = os.path.join(c.input_dir, "ThucHien", f"{c.code} - {c.sym} - TinhToanKhiNen.xlsx")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    build_compressed_air_workbook(point_data, path)
    return path


def finalize_workbook(c: Case, path: str):
    from openpyxl.styles import PatternFill, Alignment
    wb = openpyxl.load_workbook(path)
    ws = wb["Tinh_Toan_Khi_Co_Ap_V6"]
    for rng in ("C52:J59", "C61:J68", "C77:J83", "C86:J111"):     # bụi Lasair / AeroTrak, vi sinh, điểm sương: không đo
        strike_box(ws, rng)
    ws._images = [im for im in ws._images if im.anchor._from.row < 84]   # bỏ ảnh DP510 / CS HumCalc
    ws["B91"].value = ws["B102"].value = None
    ws["C36"].value = c.gas_temp                  # N*: nhiệt độ khí phiếu ghi chép (đọc từ DP510)
    for ref in ("F3", "D4", "D5", "D6", "D7", "D8", "F8", "H6", "H8", "I6", "I8", "C85"):
        ws[ref].value = "-"
    for r in range(11, 16):                       # tổng hợp vi sinh: không đo (công thức trỏ khối đã gạch -> 0)
        ws[f"C{r}"].value = ws[f"D{r}"].value = "-"
    ws["H9"].value = "=D9"
    a = ws["H2"].alignment
    ws["H2"].alignment = Alignment(horizontal=a.horizontal, vertical=a.vertical, wrap_text=True, shrink_to_fit=True)
    if not c.kqtn:                                # chưa có KQTN -> không để số mẫu của template (AGENTS 13)
        yellow = PatternFill("solid", fgColor="FFFF00")
        for ref in ("C46", "D46", "I43", "J43", "I44", "J44", "C49", "F9"):
            ws[ref].value = PLACEHOLDER
            ws[ref].fill = yellow
    if "DiemSuong" in wb.sheetnames:
        wb["DiemSuong"].sheet_state = "hidden"
    thin_outside_borders(wb)
    wb.save(path)
    wb.close()
    try:
        import win32com.client
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            x = xl.Workbooks.Open(os.path.abspath(path))
            xl.CalculateFull()
            s = x.Worksheets("Tinh_Toan_Khi_Co_Ap_V6")
            print(f"[{c.code6}] phiếu {c.sym}: N* = {s.Range('C38').Value:.4f}; tổng dầu = {s.Range('D9').Value}; "
                  f"{s.Range('F9').Value}")
            x.Save()
            x.Close()
        finally:
            xl.Quit()
    except ImportError as e:
        print(f"[{c.code6}] Excel COM warning: {e}")


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def make_dossier(c: Case, xlsx: str, cert_date: str):
    d = create_dossier_from_compressed_air_workbooks(
        [xlsx], code=c.code, cert_date=cert_date, sample_date=c.sample_date,
        sample_time_vn=f"Ngày {c.sample_date} từ {c.start} đến {c.end}", sample_time_en=f"From {c.sample_date}",
        sub_no=None, client=c.client, personnel=PersonnelInfo(executor=c.executor), equipments=[eq("oil_kit")],
        scope_vn="Giám định chất lượng khí nén", scope_en="Compressed air quality inspection",   # BBLM mục 4
        standards_vn="ISO 8573-1:2010", standards_en="ISO 8573-1:2010",                       # PYC: "Iso 8573"
        basis_vn="ISO 8573-1:2010", basis_en="ISO 8573-1:2010",
    )
    d.is_multi_cert = False
    d.reg_date = c.reg_date
    d.procedure_vn = d.procedure_en = "QUATEST 3 3030:2025, N2.TC.KN.01"      # BBLM mục 8
    d.location_desc_vn, d.location_desc_en = c.location_vn, c.location_en
    loc = d.locations[0]
    loc.name_vn, loc.name_en = c.name_vn, c.name_en
    loc.points_summary_vn = loc.points_summary_en = [c.sym]
    opt = os.path.join(c.input_dir, "Photo", "_optimized")
    d.photos = [{"path": _optimized(c.photo, opt), "caption_vn": f"Vị trí lấy mẫu {c.sym}",
                 "caption_en": f"Sampling location {c.sym}"}]
    rd = d.results_data
    rd["dewpoint"]["rows"], rd["particles"]["rows"] = [], []     # chỉ lấy mẫu tổng dầu (BBLM)
    for pp in rd.get("parsed_points", []):
        pp["dewpoint"]["requirement"] = pp["particles"]["requirement"] = ""
        pp["total_oil"]["requirement"] = ""
    k = c.kqtn
    if not k:
        rd["oil"]["rows"] = [{"loc": c.sym, "val": PLACEHOLDER, "class": PLACEHOLDER}]
    rd["system_status_vn"], rd["system_status_en"] = "Đang hoạt động", "In operation"     # phiếu ghi chép
    rd["factory_status_vn"], rd["factory_status_en"] = c.factory_status_vn, c.factory_status_en
    if c.system_chain_vn:
        rd["system_chain_vn"], rd["system_chain_en"] = c.system_chain_vn, c.system_chain_en
    so_hoi = (k or {}).get("so_hoi") or PLACEHOLDER
    so_hat = (k or {}).get("so_hat") or PLACEHOLDER
    rd["kqtn_by_location"] = {c.sym: f"- Dầu hơi: xem tại phiếu {so_hoi}\n- Dầu hạt: xem tại phiếu {so_hat}"}
    yc = _yc_date(c)
    if k:
        ent = lambda so: {"so": so, "ngay": k["ngay"], "ngay_gui": k["ngay_nhan_mau"], "ngay_nhan": k["ngay"],
                          "ngay_yc": yc, "tre_han": tre_han_text(k["ngay"], k["ngay_nhan_mau"], yc, PLACEHOLDER)}
        rd["kqtn"] = [ent(so_hoi), ent(so_hat)]
    else:
        rd["kqtn"] = [{"so": PLACEHOLDER, "ngay": PLACEHOLDER, "ngay_gui": PLACEHOLDER, "ngay_nhan": PLACEHOLDER,
                       "ngay_yc": yc, "tre_han": PLACEHOLDER}] * 2
    rd["bckq_label"] = "khi nen"
    d.input_dir, d.output_dir, d.network_dir = c.input_dir, "", ""
    d.thuchien_dir = os.path.join(c.input_dir, "ThucHien")
    d.tbkq_dir = os.path.join(c.input_dir, "TBKQ")
    d.edoc_dir = os.path.join(d.tbkq_dir, "Edoc")
    d.pdf_dir = os.path.join(c.input_dir, "pdf")
    print(f"[{c.code6}] {c.sym}: dầu {rd['oil']['rows']}")
    return d


# ---------------------------------------------------------------------------
# TBKQ / BCKQ
# ---------------------------------------------------------------------------
def build_tbkq(c: Case, d, is_en: bool) -> str:
    src = CompressedAirHandler(d).build_tbkq(is_en=is_en)
    doc = Document(src)
    _fix_cover_lists(doc)
    if not c.system_chain_vn:     # người dùng 16/11/2026: không có sơ đồ hệ thống -> "Bỏ mục mô tả" (TBKQ)
        body = doc.element.body
        for p in list(body.iterchildren(qn("w:p"))):
            if _ptext(p).strip().startswith(("Mô tả hệ thống khí nén", "Description of compressed air system")):
                nxt = p.getnext()
                body.remove(p)
                if nxt is not None and nxt.tag == qn("w:p") and _ptext(nxt).strip() == PLACEHOLDER:
                    body.remove(nxt)
    _keep_headings_with_tables(doc)
    # "9.2 Kết quả giám định" + đoạn trống + "9.2.1 Hàm lượng tổng dầu" đi liền bảng (không rớt lẻ cuối trang)
    head = next((p for p in doc.element.body.iterchildren(qn("w:p"))
                 if _ptext(p).strip() in ("Kết quả giám định", "Inspection result")), None)
    el = head
    while el is not None and el.tag == qn("w:p"):
        _keep(el)
        el = el.getnext()
    _superscript_note_marks(doc)
    _highlight_placeholders(doc)
    style_note_tables(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - khi nen - {c.code}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)               # tên handler (file do code tạo, E41)
    print(f"[{c.code6}] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


def build_bckq(c: Case, d) -> str:
    out = BCKQBuilder(d).build(out_name=f"BCKQ - 1 - khi nen - {c.code}.docx")
    doc = Document(out)
    body = doc.element.body
    _replace_in_paragraphs(body, {
        # template BCKQ ghi cứng trạng thái nhà máy -> theo phiếu ghi chép
        "nhà máy đang hoạt động sản xuất.": f"nhà máy {c.factory_status_vn[0].lower() + c.factory_status_vn[1:]}; "
                                            "hệ thống khí nén đang hoạt động.",
        f"phiếu Tính toán kết quả nồng độ hạt bụi số {c.code} ngày": f"phiếu Kết quả tính toán số {c.code} - {c.sym} ngày",
        f"phiếu Tính toán kết quả nhiệt độ điểm sương số {c.code} ngày": f"phiếu Kết quả tính toán số {c.code} - {c.sym} ngày",
    })
    if not c.system_chain_vn:     # người dùng 16/11/2026: BCKQ ghi "không có thông tin"
        for p in list(body.iterchildren(qn("w:p"))):
            if _ptext(p).strip().startswith("- Sơ đồ hệ thống khí nén"):
                nxt = p.getnext()
                if nxt is not None and nxt.tag == qn("w:p") and _ptext(nxt).strip() == PLACEHOLDER:
                    body.remove(nxt)
                ts = list(p.iter(qn("w:t")))
                ts[0].text = re.sub(r"\s*:\s*$", "", _ptext(p).rstrip()) + ": Không có thông tin."
                ts[0].set(qn("xml:space"), "preserve")
                for t in ts[1:]:
                    t.text = ""
                break
    _keep_headings_with_tables(doc)
    for t in doc.tables[1:]:
        if len(t.rows) > 1 and t.rows[0].cells[0].text.strip().startswith("Vị trí"):
            add_tbl_header(t.rows[0])
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Báo cáo của GĐV", "Gửi mẫu thử nghiệm", "- Hình ảnh", "- Sơ đồ",
                                      "Kết quả đo đạc", "Hàm lượng tổng dầu")):
            p.paragraph_format.keep_with_next = True
    _highlight_placeholders(doc)
    doc.save(out)
    print(f"[{c.code6}] BCKQ -> {out}")
    return out


def split_edoc(c: Case) -> list:
    """BBLM = scan mới tr. 1–2; Ghi chép = tr. 3 (tr. 4 trắng bỏ); Khác = scan PKSCV, PYC, điều kiện chung, báo giá."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc
    edoc = os.path.join(c.input_dir, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(c.scan_new)
    parts = []
    bblm = fitz.open(); bblm.insert_pdf(src, from_page=0, to_page=1); parts.append(("BBLM", bblm))
    gc = fitz.open(); gc.insert_pdf(src, from_page=2, to_page=src.page_count - 1); parts.append(("Ghi Chep", gc))
    kh = fitz.open(); kh.insert_pdf(fitz.open(c.scan_khac)); parts.append(("Khac", kh))
    outs = []
    for tag, pdf in parts:
        name = f"{c.code6} - {tag}.pdf"
        finish_edoc(pdf, c.input_dir, name)
        p = os.path.join(edoc, name)
        pdf.save(p)
        outs.append(p)
        print(f"[{c.code6}] Edoc -> {p} ({pdf.page_count} tr.)")
    return outs


def run(c: Case, pdf: bool = False):
    k = c.kqtn
    cert_date = k["ngay"] if k else PLACEHOLDER          # ngày ban hành = ngày KQTN (AGENTS mục 2)
    xlsx = build_workbook(c, cert_date)
    d = make_dossier(c, xlsx, cert_date)
    assign_cert_numbers([d])
    d.ensure_output_directories()
    finalize_workbook(c, xlsx)
    split_edoc(c)
    docs = [build_bckq(c, d), build_tbkq(c, d, False), build_tbkq(c, d, True)]
    if pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(d)
        ex.export_files(docs)
        ex.export_excel_files([xlsx])
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=c.code).get('status')}\t{os.path.basename(p)}")
    return docs, xlsx
