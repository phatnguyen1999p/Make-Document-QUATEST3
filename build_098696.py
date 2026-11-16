# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D098696 - Công ty TNHH Fujikura Fiber Optics Việt Nam (phòng sạch "Medical Workshop", 10/11/2026) — 1 chứng thư.

Khu vực (như TBKQ 2025 00994/N2.25/DG và sơ đồ người dùng vẽ): "Main room" 345 m² (PAR 7–26, HEPA 7–16, API 3–5),
"Incoming warehouse" 84 m² (PAR 3–6, HEPA 3–6, API 1), "Hydrophilic" 38 m² (PAR 1–2, HEPA 1–2, API 2).
Chỉ tiêu (BBLM): nồng độ hạt (0,3 ÷ 5,0) µm 26 vị trí; nhiệt độ, độ ẩm 26 vị trí; lưu lượng gió cấp qua lọc HEPA 16;
chênh áp 05. Không có mức quy định (TBKQ 2025: Căn cứ "-") -> không đánh giá.

Nguồn (Input/098696/_Z — chép nguyên thư mục hồ sơ trên Z 15/11/2026, chỉ đọc):
  0.  BBLM .pdf — BM3030.02 2 trang: PYC Q326D098696 ngày 11/08/2026; địa chỉ "Số 9, đường số 6, Khu công nghiệp Việt Nam
      – Singapore, phường Bình Hòa, thành phố Hồ Chí Minh"; "Đo chất lượng phòng sạch" (hạt 0,3 ÷ 5,0 µm; nhiệt độ, độ ẩm
      tương đối; chênh lệch áp suất; lưu lượng gió cấp qua lọc HEPA); địa điểm "tại địa chỉ nêu tại mục 3"; 10/11/2026
      8h45 – 12h00; AeroTrak 9350, Testo 625, Capture Hood Testo 420; N2.TC.PS.01; nhà máy hoạt động sản xuất; báo giá
      QONV26012585; 26 / 16 / 26 / 05 vị trí; "Khu vực đo: Medical Workshop"; ký: Phạm Quốc Nghĩa.
  0.  so do vi tri do + .pdf — tr. 1 sơ đồ người dùng vẽ; tr. 3 mặt bằng KH (M5-DRW-000002) ghi tay số file bụi + nhiệt
      ẩm từng PAR (= phiếu ghi chép); tr. 5 ghi tay lưu lượng; "Tuấn + Nghĩa 10/11/26".
  0.  SO DO VI TRI DO .docx — sơ đồ người dùng vẽ (EMF): image1 = VN, image2 = EN (dùng nguyên, không chú thích).
  THUC HIEN -/98696 - KQ DO/D98696 - bui/6ChannelFixed.xlsx — AeroTrak 9350-02 S/N 93501302003, Σ, # / m³, 1 phút
      (50 L): zero 534, 537, 542; PAR 1–2 = 535–536, PAR 3–6 = 538–541, PAR 7–26 = 543–562 (số file ghi tay tr. 3).
  THUC HIEN -/98696 - tinh toan KQ do/1.  NHIET & AM (DA HIEU CHUAN) .xlsx — người dùng nhập 15/11/2026: 1 lần đọc /
      vị trí, hiệu chỉnh TREND theo giấy hiệu chuẩn Testo 625 Q326C032208 (T 15,2/20/24,9/29,8/39,7 -> 15/20/25/30/40;
      RH 45,8/54,5/64,8/74,3/84 -> 40/50/60/70/80).
  THUC HIEN -/98696 - tinh toan KQ do/1.  TINH LUU LUONG + CHENH AP  .xlsx — người dùng: lưu lượng 3 lần đọc / HEPA,
      TB × 0,981 (hệ số giấy hiệu chuẩn Testo 420); chênh áp API 1–5 3 lần đọc, TB (không hệ số).
  3. TBKQ - viet / ENG - 2025 .docx — kết quả cũ cùng khách (tên khu vực, cách trình bày, tên chênh áp).
  2.  BM3030.03- BCKQGD - 98696 .docx — BCKQ người dùng soạn dở 15/11/2026 (người thực hiện "Nghĩa + Tuấn", ghi chú không
      chụp hình).

Quyết định của người dùng (15/11/2026):
  - "Chạy 98696, hồ sơ này không có hình, ghi trong báo cáo là khu vực thực hiện không được phép ghi hình, người thực
    hiện là Nghĩa và Tuấn".
  - (giao ở hồ sơ nhầm số 096547) "lưu ý có kết quả cũ năm 2025, đọc theo nội dung ghi trên kết quả, và thể hiện
    tương tự theo form kết quả mới" -> tên khu vực / vị trí / cách trình bày theo TBKQ 2025, dựng bằng form TAG mới.

Người dùng 15/11/2026: "6 chép luôn lên Z, còn lại giữ" -> các điểm dưới GIỮ NGUYÊN (đã xác nhận); đã chép lên Z.

Suy luận của AI (người dùng đã xác nhận 15/11/2026):
  - Nhiệt độ, độ ẩm lấy theo phiếu người dùng đã hiệu chỉnh (TREND, hồi quy cả bảng) — AGENTS 8C chốt nội suy 2 điểm kẹp
    và công tắc AP_HIEU_CHINH đang tắt; nội suy lệch TREND 0,1 – 0,4 %RH. Phiếu Excel dựng lại dùng công thức TREND.
  - PAR 1 độ ẩm: phiếu người dùng TREND trên 4 điểm ($I$6:$I$9), các vị trí khác 5 điểm -> dựng lại theo 5 điểm
    (cùng kết quả làm tròn 44,6).
  - Lưu lượng = TB 3 lần × 0,981 (như phiếu người dùng); chênh áp = TB 3 lần.
  - Tên khu vực "Incoming warehouse" (TBKQ 2025, sơ đồ người dùng); mặt bằng KH và BCKQ soạn dở ghi "ICM&WH".
  - API 2: "Hydrophilic -> Hành lang" theo sơ đồ người dùng (TBKQ 2025 ghi "Môi trường").
  - Địa chỉ theo BBLM (phường Bình Hòa); BCKQ soạn dở ghi "Số 9 VSIP …, Phường Thuận Giao".
  - GĐV ký Phạm Quốc Nghĩa (người thực hiện Nghĩa + Tuấn); người báo cáo "Nghĩa + Tuấn + Minh".
  - Tên tiếng Anh khách hàng theo TBKQ EN 2026 hồ sơ 096547 ("Fujikura Fiber Optics Vietnam Company Limited").

Chạy:  python build_098696.py          (DOCX + XLSX + Edoc)
       python build_098696.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: trong code (network_dir để trống) — chép lên Z: python Core/Tools/ai_files.py chep 098696 --ghi
"""

import os
import re
import sys
import copy
import argparse
import datetime as _dt

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_ton_tai_Z"))

import numpy as np
import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import PLACEHOLDER, style_note_tables, appendix_total_field
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq
from Core.Tools.cleanroom_calc import CleanroomCalcBook, read_aerotrak, round_half_up

CODE = "Q326D098696"
CODE6 = "098696"
REG_DATE = "11/08/2026"            # BBLM mục 1 (PYC)
SAMPLE_DATE = "10/11/2026"         # BBLM
START, END = "08:45", "12:00"      # BBLM mục 6
CERT_DATE = _dt.date.today().strftime("%d/%m/%Y")    # không có KQTN -> ngày dựng (AGENTS mục 2)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SRC = os.path.join(INPUT_DIR, "_Z")
SCAN_BBLM = os.path.join(SRC, "0.  BBLM .pdf")
SCAN_GC = os.path.join(SRC, "0.  so do vi tri do + .pdf")
LAYOUT_DOCX = os.path.join(SRC, "0.  SO DO VI TRI DO .docx")
DUST_XLSX = os.path.join(SRC, "THUC HIEN -", "98696 - KQ DO", "D98696 - bui", "6ChannelFixed.xlsx")
USER_TRH = os.path.join(SRC, "THUC HIEN -", "98696 - tinh toan KQ do", "1.  NHIET & AM (DA HIEU CHUAN) .xlsx")
USER_FLOW = os.path.join(SRC, "THUC HIEN -", "98696 - tinh toan KQ do", "1.  TINH LUU LUONG + CHENH AP  .xlsx")
DIAGRAM_DIR = os.path.join(INPUT_DIR, "TBKQ", "_diagrams")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Fujikura Fiber Optics Việt Nam",
    name_en="Fujikura Fiber Optics Vietnam Company Limited",
    address_vn="Số 9, đường số 6, Khu công nghiệp Việt Nam – Singapore, phường Bình Hòa, Thành phố Hồ Chí Minh",  # BBLM
    address_en="No. 9, street No. 6, Vietnam – Singapore Industrial Park, Binh Hoa ward, Ho Chi Minh City",
)
LOCATION_VN = f"Phòng sạch của {CLIENT.name_vn} - {CLIENT.address_vn}"     # AGENTS 15 (phòng sạch)
LOCATION_EN = f"Cleanroom of {CLIENT.name_en} - {CLIENT.address_en}"
EXECUTOR = "Nghĩa + Tuấn"          # người dùng 15/11/2026
REPORTER = "Nghĩa + Tuấn + Minh"
NO_PHOTO_VN = "Theo quy định của công ty, khu vực thực hiện không được phép ghi hình."
NO_PHOTO_EN = "According to the company's regulations, recording images in the performance area is not permitted."

# (tên VN/EN, diện tích m², PAR, HEPA, số vị trí chênh áp) — thứ tự như TBKQ 2025
ROOMS = [
    ("Hydrophilic", "Hydrophilic", 38, range(1, 3), range(1, 3), 1),
    ("Incoming warehouse", "Incoming warehouse", 84, range(3, 7), range(3, 7), 1),
    ("Main room", "Main room", 345, range(7, 27), range(7, 17), 3),
]
ROOM_OF_PAR = {p: r[0] for r in ROOMS for p in r[3]}
ROOM_OF_HEPA = {h: r[0] for r in ROOMS for h in r[4]}
ZERO_FILES = {"Hydrophilic": 534, "Incoming warehouse": 537, "Main room": 542}   # file zero đầu mỗi khu vực
PAR_FILE = {p: (534 + p if p <= 2 else 535 + p if p <= 6 else 536 + p) for p in range(1, 27)}  # tr. 3 phiếu ghi chép
# chênh áp: (ký hiệu, khu vực 1, khu vực 2 VN, EN) — TBKQ 2025 + sơ đồ người dùng (API 2 -> Hành lang)
DP_ROWS = [("API 1", "Incoming warehouse", "Material delivery", "Material delivery"),
           ("API 2", "Hydrophilic", "Hành lang", "Corridor"),
           ("API 3", "Main room", "Locker room", "Locker room"),
           ("API 4", "Main room", "Carton packing", "Carton packing"),
           ("API 5", "Main room", "Shipping room", "Shipping room")]
CAL_DEVICE = "Testo 625 (MT33-3)"
CAL_CERT = "Q326C032208"
CAL_T = [(15.2, 15.0), (20.0, 20.0), (24.9, 25.0), (29.8, 30.0), (39.7, 40.0)]       # (chỉ thị, chuẩn)
CAL_RH = [(45.8, 40.0), (54.5, 50.0), (64.8, 60.0), (74.3, 70.0), (84.0, 80.0)]
LAYOUT_W_CM = 12.0                                    # sơ đồ: VN = EN cùng nằm trang phụ lục 1 (soát PDF)
ROW_TWIPS = 400                                       # hàng bảng kết quả 0,7 cm (AGENTS mục 3: tối thiểu 0,6 cm)
FLOW_FACTOR = 0.981                                   # hệ số giấy hiệu chuẩn Testo 420 (phiếu người dùng)
SO = {k: f"{CODE}/{v}" for k, v in (("bui", "Bụi"), ("nhiet_am", "Nhiệt độ, độ ẩm"), ("chenh_ap", "Chênh áp"),
                                    ("luu_luong", "Lưu lượng"))}


def vn(x, nd=1):
    return f"{x:.{nd}f}".replace(".", ",")


def trend(x, tab):
    a, b = np.polyfit([p[0] for p in tab], [p[1] for p in tab], 1)
    return a * x + b


# ---------------------------------------------------------------------------
# Đọc số liệu
# ---------------------------------------------------------------------------
LABELS, RECS = read_aerotrak(DUST_XLSX)
for f in list(ZERO_FILES.values()) + list(PAR_FILE.values()):
    if f not in RECS:
        raise SystemExit(f"[FATAL] 6ChannelFixed.xlsx thiếu file {f}")
for z in ZERO_FILES.values():
    if any(RECS[z][f"Ch{i}"] for i in range(1, 7)):
        raise SystemExit(f"[FATAL] file {z} không phải zero")
DUST = {p: [RECS[f][f"Ch{i}"] for i in range(1, 7)] for p, f in PAR_FILE.items()}


def read_user_trh():
    ws = openpyxl.load_workbook(USER_TRH, data_only=True).active
    out = {}
    for r in range(5, 31):
        m = re.match(r"PAR\s*(\d+)", str(ws.cell(r, 1).value or ""))
        out[int(m.group(1))] = {"t": ws.cell(r, 2).value, "rh": ws.cell(r, 3).value,
                                "t_user": ws.cell(r, 4).value, "rh_user": ws.cell(r, 5).value}
    assert sorted(out) == list(range(1, 27)), sorted(out)
    return out


def read_user_flow():
    ws = openpyxl.load_workbook(USER_FLOW, data_only=True).active
    flows, dps = {}, {}
    for r in range(1, ws.max_row + 1):
        lab = str(ws.cell(r, 7).value or "").strip()
        m = re.match(r"HEPA\s*(\d+)", lab)
        if m:
            flows[int(m.group(1))] = [ws.cell(r + i, 3).value for i in range(3)]
        m = re.match(r"API\s*-\s*(\d+)", lab)
        if m:
            dps[int(m.group(1))] = [ws.cell(r + i, 1).value for i in range(3)]
    assert sorted(flows) == list(range(1, 17)) and sorted(dps) == list(range(1, 6)), (sorted(flows), sorted(dps))
    return flows, dps


TRH = read_user_trh()
FLOW_READS, DP_READS = read_user_flow()
T_VAL = {p: round_half_up(trend(v["t"], CAL_T), 1) for p, v in TRH.items()}
RH_VAL = {p: round_half_up(trend(v["rh"], CAL_RH), 1) for p, v in TRH.items()}
for p, v in TRH.items():                                  # đối chiếu phiếu người dùng (AGENTS mục 11)
    if (round_half_up(v["t_user"], 1), round_half_up(v["rh_user"], 1)) != (T_VAL[p], RH_VAL[p]):
        print(f"[WARNING] [Data Mismatch] PAR {p}: phiếu người dùng {v['t_user']:.2f} / {v['rh_user']:.2f}, "
              f"code {T_VAL[p]} / {RH_VAL[p]}")
FLOW = {h: round(sum(r) / 3 * FLOW_FACTOR) for h, r in FLOW_READS.items()}
DP = {n: round_half_up(sum(r) / 3, 1) for n, r in DP_READS.items()}


# ---------------------------------------------------------------------------
# Sơ đồ (người dùng vẽ sẵn — EMF trong 0. SO DO VI TRI DO.docx)
# ---------------------------------------------------------------------------
def build_layouts() -> dict:
    import zipfile
    from PIL import Image
    os.makedirs(DIAGRAM_DIR, exist_ok=True)
    out = {}
    z = zipfile.ZipFile(LAYOUT_DOCX)
    for lang, name in (("vn", "image1.emf"), ("en", "image2.emf")):
        emf = os.path.join(DIAGRAM_DIR, f"_so_do_{lang}.emf")
        open(emf, "wb").write(z.read(f"word/media/{name}"))
        im = Image.open(emf)
        im.load(dpi=150)
        png = os.path.join(DIAGRAM_DIR, f"so_do_vi_tri_{lang}.png")
        im.convert("RGB").save(png)
        os.remove(emf)
        out[lang] = [{"path": png, "caption_vn": "", "caption_en": ""}]
    return out


# ---------------------------------------------------------------------------
# Dossier
# ---------------------------------------------------------------------------
def build_dossier(layouts) -> Dossier:
    locs = []
    for name_vn, name_en, area, pars, hepas, n_dp in ROOMS:
        n_p, n_h = len(pars), len(hepas)
        locs.append(MonitoringLocation(
            name_vn=name_vn, name_en=name_en, area_m2=str(area), clean_class="-",
            state_vn="Đang hoạt động", state_en="In operation",
            points_summary_vn=[f"Nồng độ hạt bụi: {n_p:02d} vị trí;", f"Nhiệt độ - độ ẩm: {n_p:02d} vị trí;",
                               f"Lưu lượng gió qua lọc HEPA: {n_h:02d} vị trí;", f"Chênh lệch áp suất: {n_dp:02d} vị trí."],
            points_summary_en=[f"Particle concentration: {n_p:02d} location{'s' if n_p > 1 else ''};",
                               f"Temperature - Humidity: {n_p:02d} location{'s' if n_p > 1 else ''};",
                               f"Airflow rate through HEPA filter: {n_h:02d} location{'s' if n_h > 1 else ''};",
                               f"Pressure differential: {n_dp:02d} location{'s' if n_dp > 1 else ''}."],
        ))
    pts = [(name, p) for name, *_r in ROOMS for p in _r[2]]
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}", sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False, dossier_type="cleanroom",
        scope_vn="Giám định chất lượng phòng sạch", scope_en="Cleanroom performance inspection",
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.PS.01", procedure_en="QUATEST 3 3030:2025, N2.TC.PS.01",
        standards_vn="-", standards_en="-",
        client=CLIENT, personnel=PersonnelInfo(executor=EXECUTOR),
        equipments=[eq("aerotrak_9350"), eq("testo_625"), eq("testo_420")],     # BBLM mục 7
        locations=locs, photos=[], layouts=layouts["vn"],
        results_data={
            "layout_width_cm": LAYOUT_W_CM,
            "bui_6kenh": [{"cap": "-", "gioi_han": [None] * 6, "trang_so_do": "2",
                           "diem": [{"khu_vuc": room, "vi_tri": f"PAR {p}", "gia_tri": DUST[p]} for room, p in pts]}],
            "nhiet_am": {"dong": [{"khu_vuc": room, "nhiet_do": vn(T_VAL[p]), "do_am": vn(RH_VAL[p])} for room, p in pts],
                         "muc_nhiet_do": "-", "muc_do_am": "-", "muc_do_on": "-"},
            "chenh_ap": {"dong": [{"khu_vuc_1": a, "khu_vuc_2": b, "khu_vuc_2_en": b_en, "gia_tri": vn(DP[int(s[-1])]),
                                   "muc": "-"} for s, a, b, b_en in DP_ROWS]},
            "luu_luong": [{"muc_ach": None, "trang_so_do": "2",
                           "diem": [{"khu_vuc": ROOM_OF_HEPA[h], "vi_tri": f"HEPA {h}", "luu_luong": FLOW[h]}
                                    for name, *_r in ROOMS for h in _r[3]]}],
        },
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
# Tiện ích DOCX (theo build_134179 / build_138143)
# ---------------------------------------------------------------------------
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.text.paragraph import Paragraph
from Core.Styling.styles import set_cell_shading
from Core.Styling.tables import format_table_structure, add_tbl_header, add_cant_split
from Core.Styling.elements import add_document_ending_line


def _ptext(el):
    import unicodedata
    return unicodedata.normalize("NFC", "".join(t.text or "" for t in el.iter(qn("w:t"))))


def _keep(p_el):
    pPr = p_el.get_or_add_pPr()
    if pPr.find(qn("w:keepNext")) is None:
        pPr.insert(0, OxmlElement("w:keepNext"))


def _keep_rows(rows):
    for tr in rows:
        for p_el in tr.iter(qn("w:p")):
            _keep(p_el)


def _is_note_table(el) -> bool:
    return el is not None and el.tag == qn("w:tbl") and _ptext(el).strip().startswith(("Ghi chú", "Note"))


def _set_text(p_el, text: str):
    runs = p_el.findall(qn("w:r"))
    proto = copy.deepcopy(runs[0])
    for ch in [c for c in p_el if c.tag != qn("w:pPr")]:
        p_el.remove(ch)
    for c in [c for c in proto if c.tag != qn("w:rPr")]:
        proto.remove(c)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    proto.append(t)
    p_el.append(proto)


def _set_cell(tc, text: str):
    ps = tc.findall(qn("w:p"))
    base = next((p for p in ps if p.findall(qn("w:r"))), ps[0])
    for p in ps:
        if p is not base:
            tc.remove(p)
    _set_text(base, text)


def note_rows(proto, lines):
    el = copy.deepcopy(proto)
    rows = el.findall(qn("w:tr"))
    item = rows[-1]
    for r in rows[1:]:
        el.remove(r)
    for marker, text in lines:
        tr = copy.deepcopy(item)
        el.append(tr)
        tcs = tr.findall(qn("w:tc"))
        _set_cell(tcs[0], marker)
        _set_cell(tcs[1], text)
    return el


def _superscript_note_marks(doc):
    from docx.text.run import Run
    for tbl in doc.element.body.iterchildren(qn("w:tbl")):
        if not _is_note_table(tbl):
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


def add_runs(p, text, size=13, bold=False, italic=False, sup=False):
    r = p.add_run(text)
    r.font.size, r.bold, r.italic = Pt(size), bold, italic
    r.font.name = "Times New Roman"
    if sup:
        r.font.superscript = True
    return r


def cell_text(cell, text, bold=False, italic=False, align="center", sup="", size=13):
    for p in list(cell.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = cell.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    for i, line in enumerate(text.split("\n")):
        p = p0 if i == 0 else cell.add_paragraph()
        p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
        pf = p.paragraph_format
        pf.space_before, pf.space_after, pf.line_spacing = Pt(1), Pt(1), 1.0
        pf.left_indent = pf.first_line_indent = Pt(0)
        add_runs(p, line, size, bold, italic)
    if sup:
        add_runs(cell.paragraphs[-1], sup, size, bold, italic, sup=True)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def make_table(doc, anchor_el, rows, cols, widths_pct, header_rows=1, row_height=440):
    t = doc.add_table(rows=rows, cols=cols)
    anchor_el.addnext(t._tbl)
    format_table_structure(t, row_height=row_height)
    for mar in list(t._tbl.iter(qn("w:tcMar"))):      # E51
        mar.getparent().remove(mar)
    total = 9355
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths_pct):
        gc.set(qn("w:w"), str(round(total * w / 100)))
    for row in t.rows:
        for c, w in zip(row._tr.findall(qn("w:tc")), widths_pct):
            tcPr = c.get_or_add_tcPr()
            old = tcPr.find(qn("w:tcW"))
            if old is not None:
                tcPr.remove(old)
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(round(50 * w)))
            tcW.set(qn("w:type"), "pct")
            tcPr.insert(0, tcW)
        add_cant_split(row)
    for r in t.rows[:header_rows]:
        add_tbl_header(r)
        for c in r.cells:
            set_cell_shading(c, "D9D9D9")
    return t


def thousands(v):
    return f"{int(v):,}".replace(",", " ")


def dec(x, is_en):
    s = vn(x)
    return s.replace(",", ".") if is_en else s


# ---------------------------------------------------------------------------
# TBKQ: bìa + 9.1 theo template phòng sạch; 9.2 dựng lại theo bố cục TBKQ 2025 (bảng theo từng vị trí, không mức)
# ---------------------------------------------------------------------------
def _drop_class_column(doc):
    """Bảng khu vực: bỏ cột "Cấp sạch" / "Class" (hồ sơ không có cấp sạch — như TBKQ 2025)."""
    for t in doc.tables:
        if not t.rows:
            continue
        head = [c.text.strip() for c in t.rows[0].cells]
        idx = next((i for i, h in enumerate(head) if re.match(r"(Cấp sạch|Class|Clean)", h)), None)
        if idx is None:
            continue
        grid = t._tbl.tblGrid.findall(qn("w:gridCol"))
        w_rm = int(grid[idx].get(qn("w:w")))
        t._tbl.tblGrid.remove(grid[idx])
        for tr in t._tbl.findall(qn("w:tr")):
            tcs = tr.findall(qn("w:tc"))
            rm = tcs[idx].find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
            tr.remove(tcs[idx])
            last = tr.findall(qn("w:tc"))[-1].find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
            if rm is not None and last is not None and rm.get(qn("w:type")) == last.get(qn("w:type")):
                last.set(qn("w:w"), str(int(last.get(qn("w:w"))) + int(rm.get(qn("w:w")))))   # dồn độ rộng (dxa / pct)
        grid = t._tbl.tblGrid.findall(qn("w:gridCol"))
        grid[-1].set(qn("w:w"), str(int(grid[-1].get(qn("w:w"))) + w_rm))
        break


def _fix_91(doc, is_en):
    body = doc.element.body
    _drop_class_column(doc)
    # ghi chú bảng khu vực
    rep = {
        r"Thông số kỹ thuật phòng sạch.*": f"Diện tích các phòng do đại diện {CLIENT.name_vn} cung cấp.",
        r".*bảng A1.*": f"Số lượng vị trí đo do đại diện {CLIENT.name_vn} chỉ định.",
        r"(The )?([Cc]leanroom|[Tt]echnical) (specification|parameter)s?.*": f"The room area was provided by the representative of {CLIENT.name_en}.",
        r".*[Tt]able A\.?1.*": f"The number of measuring locations was indicated by the representative of {CLIENT.name_en}.",
    }
    for tbl in body.iterchildren(qn("w:tbl")):
        if not _is_note_table(tbl):
            continue
        for tc in tbl.iter(qn("w:tc")):
            txt = _ptext(tc).strip()
            for pat, new in rep.items():
                if re.fullmatch(pat, txt):
                    _set_cell(tc, new)
    for p in body.iterchildren(qn("w:p")):
        txt = _ptext(p).strip()
        if txt.startswith("Sơ đồ các vị trí quan trắc"):
            _set_text(p, "Sơ đồ các vị trí đo:")
        elif re.match(r"(Layout|Diagram) of (the )?monitoring", txt):
            _set_text(p, "Layout of measuring locations:")


def _results_vn_en(is_en):
    L = (lambda vn_, en_: en_) if is_en else (lambda vn_, en_: vn_)
    note_pos = ("(1) …", L("Vị trí đo xem sơ đồ minh họa tại trang 2.", "See measuring locations in the diagram on page 2."))
    note_acc = ("●", L("Lĩnh vực thực hiện được công nhận ISO 17020 - VIAS 004.",
                       "The performance field is accredited to ISO 17020 - VIAS 004."))
    area, loc = L("Khu vực đo", "Area"), L("Vị trí đo ", "Location ")     # EN ngắn: tiêu đề 1 dòng như VN (VN = EN ngắt trang)
    pts = [(name, p) for name, *_r in ROOMS for p in _r[2]]
    hep = [(name, h) for name, *_r in ROOMS for h in _r[3]]
    sizes = ["0,3", "0,5", "1", "2", "3", "5"]
    return [
        {"head": L("Kết quả đo nồng độ hạt bụi", "Measuring result of particle concentration"),
         "widths": [20, 12] + [68 / 6] * 6, "hdr2": True,
         "top": [(area, 0, 0, False), (loc, 1, 1, True),
                 (L("Nồng độ hạt bụi trong không khí, hạt/m³", "Airborne particle concentration, particles/m³"), 2, 7, False)],
         "sub": [f"d ≥ {s.replace(',', '.') if is_en else s} µm" for s in sizes],
         "rows": [(room, f"PAR {p}", [thousands(v) for v in DUST[p]]) for room, p in pts],
         "notes": [note_pos, note_acc]},
        {"head": L("Kết quả đo nhiệt độ - độ ẩm", "Measuring result of temperature - humidity"),
         "widths": [32, 20, 24, 24], "hdr2": False,
         "top": [(area, 0, 0, False), (loc, 1, 1, True), (L("Nhiệt độ, °C", "Temperature, °C"), 2, 2, False),
                 (L("Độ ẩm, %RH", "Humidity, %RH"), 3, 3, False)],
         "rows": [(room, f"PAR {p}", [dec(T_VAL[p], is_en), dec(RH_VAL[p], is_en)]) for room, p in pts],
         "notes": [note_pos, note_acc]},
        {"head": L("Kết quả đo lưu lượng gió cấp qua lọc HEPA", "Measuring result of airflow rate through HEPA filter"),
         "widths": [38, 27, 35], "hdr2": False,
         "top": [(area, 0, 0, False), (loc, 1, 1, True), (L("Lưu lượng gió, m³/h", "Airflow rate, m³/h"), 2, 2, False)],
         "rows": [(room, f"HEPA {h}", [thousands(FLOW[h])]) for room, h in hep],
         "notes": [note_pos, note_acc]},
        {"head": L("Kết quả đo chênh lệch áp suất giữa các khu vực",
                   "Measuring result of pressure differential between areas"),
         "widths": [28, 28, 20, 24], "hdr2": False, "pressure": True,
         "top": [(L("Khu vực đo 1", "Area 1"), 0, 0, False), (L("Khu vực đo 2", "Area 2"), 1, 1, False),
                 (loc, 2, 2, True), (L("Mức chênh áp, Pa", "Pressure differential, Pa"), 3, 3, False)],
         "rows": [(a, b_en if is_en else b, s, [dec(DP[int(s[-1])], is_en)]) for s, a, b, b_en in DP_ROWS],
         "notes": [note_pos, note_acc]},
    ]


def _build_results(doc, is_en):
    body = doc.element.body
    kids = list(body.iterchildren())
    sect_i = next(i for i, el in enumerate(kids)
                  if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None)
    first_head = None
    for el in kids[sect_i + 1:]:
        if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None \
                and re.search(r"bụi|particle", _ptext(el), re.I):
            first_head = el
            break
    assert first_head is not None, "không thấy tiêu đề mục kết quả bụi"
    head_proto = copy.deepcopy(first_head)
    nxt, gap_proto, note_proto = first_head.getnext(), None, None
    while nxt is not None:
        if nxt.tag == qn("w:tbl") and _is_note_table(nxt):
            note_proto = copy.deepcopy(nxt)
            gp = nxt.getprevious()
            if gp is not None and gp.tag == qn("w:p") and not _ptext(gp).strip():
                gap_proto = copy.deepcopy(gp)
            break
        nxt = nxt.getnext()
    assert note_proto is not None and gap_proto is not None
    # xóa từ tiêu đề bụi tới hết (kể cả đường kết thúc — thêm lại sau)
    started = False
    for el in list(body.iterchildren()):
        if el is first_head:
            started = True
        if started and el.tag != qn("w:sectPr"):
            body.remove(el)
    prev = body.find(qn("w:sectPr")).getprevious()      # đoạn trống sau sơ đồ (template) -> bỏ, kẻo đẩy bảng xuống
    while prev is not None and prev.tag == qn("w:p") and not _ptext(prev).strip() \
            and prev.find(f".//{qn('w:drawing')}") is None and prev.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is None:
        nxt_prev = prev.getprevious()
        body.remove(prev)
        prev = nxt_prev
    anchor = [body.find(qn("w:sectPr")).getprevious()]

    def put(el):
        anchor[0].addnext(el)
        anchor[0] = el
        return el

    for sec in _results_vn_en(is_en):
        h = put(copy.deepcopy(head_proto))
        for r in h.findall(qn("w:r"))[1:]:
            h.remove(r)
        _set_text(h, sec["head"])
        _keep(h)
        pf = Paragraph(h, None).paragraph_format
        pf.space_before = Pt(6)
        n_hdr = 2 if sec["hdr2"] else 1
        ncol = len(sec["widths"])
        t = make_table(doc, anchor[0], n_hdr + len(sec["rows"]), ncol, sec["widths"], header_rows=n_hdr,
                       row_height=ROW_TWIPS)
        anchor[0] = t._tbl
        for text, c0, c1, mark in sec["top"]:
            cell = t.cell(0, c0)
            if c1 > c0:
                cell = cell.merge(t.cell(0, c1))
            elif n_hdr == 2:
                cell = cell.merge(t.cell(1, c0))
            cell_text(cell, text.rstrip(), sup="(1)" if mark else "")
        if n_hdr == 2:
            for j, s in enumerate(sec["sub"]):
                cell_text(t.cell(1, 2 + j), s, italic=True)
        rows = sec["rows"]
        for i, row in enumerate(rows):
            r = n_hdr + i
            if sec.get("pressure"):
                a, b, s, vals = row
                cell_text(t.cell(r, 0), a)
                cell_text(t.cell(r, 1), b)
                cell_text(t.cell(r, 2), s)
                cell_text(t.cell(r, 3), vals[0])
            else:
                room, label, vals = row
                cell_text(t.cell(r, 1), label)
                for j, v in enumerate(vals):
                    cell_text(t.cell(r, 2 + j), v)
        if not sec.get("pressure"):                    # gộp dọc ô khu vực theo phòng
            i = 0
            while i < len(rows):
                k = i
                while k + 1 < len(rows) and rows[k + 1][0] == rows[i][0]:
                    k += 1
                cell = t.cell(n_hdr + i, 0)
                if k > i:
                    cell = cell.merge(t.cell(n_hdr + k, 0))
                cell_text(cell, rows[i][0])
                i = k + 1
        trs = t._tbl.findall(qn("w:tr"))
        _keep_rows(trs[:min(n_hdr + 3, len(trs) - 1)])
        _keep_rows(trs[-2:])
        g = put(copy.deepcopy(gap_proto))
        _keep(g)
        note = put(note_rows(note_proto, sec["notes"]))
        _keep_rows(note.findall(qn("w:tr"))[:-1])


def finish_tbkq(src: str, is_en: bool) -> str:
    doc = Document(src)
    _fix_cover_lists(doc)
    _fix_91(doc, is_en)
    _build_results(doc, is_en)
    add_document_ending_line(doc)
    appendix_total_field(doc)
    _superscript_note_marks(doc)
    style_note_tables(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - {'en' if is_en else 'viet'} - phong sach - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)
    print(f"[build_098696] TBKQ {'EN' if is_en else 'VN'} -> {dst}")
    return dst


# ---------------------------------------------------------------------------
# BCKQ
# ---------------------------------------------------------------------------
REFS = [
    f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {SO['bui']} ngày {SAMPLE_DATE}.",
    f"- Kết quả đo nhiệt độ, độ ẩm: xem tại phiếu Tính toán kết quả nhiệt độ, độ ẩm số {SO['nhiet_am']} ngày {SAMPLE_DATE}.",
    f"- Kết quả đo lưu lượng gió cấp qua lọc HEPA: xem tại phiếu Tính toán vận tốc gió và lưu lượng qua lọc HEPA "
    f"số {SO['luu_luong']} ngày {SAMPLE_DATE}.",
    f"- Kết quả đo chênh lệch áp suất: xem tại phiếu Kết quả đo chênh lệch áp suất số {SO['chenh_ap']} ngày {SAMPLE_DATE}.",
]


def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["tham_chieu"] = REFS
    ctx["danh_gia"] = [{"khu_vuc": "; ".join(f"“{r[0]}”" for r in ROOMS),
                        "dat": "Không đánh giá", "khong_dat": "-",
                        "can_cu": f"{CLIENT.name_vn} không cung cấp mức quy định"}]
    out = b.build(ctx=ctx, out_name=f"BCKQ - 1 - phong sach - {CODE}.docx")
    doc = Document(out)
    body = doc.element.body
    _drop_class_column(doc)
    for tbl in body.iterchildren(qn("w:tbl")):
        if _is_note_table(tbl):
            for tc in tbl.iter(qn("w:tc")):
                txt = _ptext(tc).strip()
                if txt.startswith("Thông số kỹ thuật phòng sạch"):
                    _set_cell(tc, f"Diện tích các phòng do đại diện {CLIENT.name_vn} cung cấp.")
                elif "bảng A1" in txt:
                    _set_cell(tc, f"(1) Diện tích các phòng do đại diện {CLIENT.name_vn} cung cấp. "
                                  f"Số lượng vị trí đo do đại diện {CLIENT.name_vn} chỉ định.")
    # Ảnh: người dùng 15/11/2026 — ghi "khu vực thực hiện không được phép ghi hình" (AGENTS mục 3: sau đầu mục ảnh)
    paras = list(body.iter(qn("w:p")))
    hp = next((p for p in paras if _ptext(p).strip().startswith(("- Hình ảnh", "Hình ảnh"))), None)
    if hp is not None:
        nxt = copy.deepcopy(hp)
        hp.addnext(nxt)
        _set_text(nxt, NO_PHOTO_VN)
    else:
        sd = next(p for p in paras if _ptext(p).strip().startswith(("Sơ đồ", "- Sơ đồ")))
        new = copy.deepcopy(sd)
        sd.addprevious(new)
        _set_text(new, f"- Hình ảnh quá trình thực hiện: {NO_PHOTO_VN}")
    # sơ đồ BCKQ 11 cm (giữ tỉ lệ, không crop) để báo cáo gọn 2 trang — không để trang cuối chỉ có "Người báo cáo"
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for inl in body.iter(f"{ns}inline"):
        ext = inl.find(f"{ns}extent")
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
        new = int(11.0 * 360000)
        if cx > new:
            ncy = round(cy * new / cx)
            ext.set("cx", str(new)); ext.set("cy", str(ncy))
            for e in inl.iter(f"{a_ns}ext"):
                if e.get("cx"):
                    e.set("cx", str(new)); e.set("cy", str(ncy))
    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith("Ngày báo cáo"):
            p.paragraph_format.keep_with_next = True
        if t.startswith(("Sơ đồ", "- Sơ đồ", "Báo cáo của GĐV", "Kết quả đo đạc")):
            p.paragraph_format.keep_with_next = True
        if t.startswith("Người báo cáo") and REPORTER not in t:
            print(f"[WARNING] BCKQ người báo cáo: {t}")
    doc.save(out)
    print(f"[build_098696] BCKQ -> {out}")
    return out


# ---------------------------------------------------------------------------
# Excel phiếu tính toán (form phòng sạch chung — cleanroom_calc)
# ---------------------------------------------------------------------------
def build_excel() -> str:
    from openpyxl.utils import get_column_letter as L_
    book = CleanroomCalcBook()
    groups = []
    for name, _en, _a, pars, _h, _n in ROOMS:
        groups.append({"room": name, "cap": "-", "limits": {}, "zero": RECS[ZERO_FILES[name]],
                       "points": [RECS[PAR_FILE[p]] for p in pars], "labels": [f"PAR {p}" for p in pars]})
    ws_d = book.dust_aerotrak(groups, so=SO["bui"], date=SAMPLE_DATE, labels=LABELS)
    for row in ws_d.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("Giới hạn"):
                c.value = "Không có mức quy định"
    book.temp_rh([{"room": ROOM_OF_PAR[p], "loc": f"PAR {p}", "t": [TRH[p]["t"]], "rh": [TRH[p]["rh"]],
                   "lim_t": None, "lim_rh": None} for p in range(1, 27)], None, None,
                 so=SO["nhiet_am"], date=SAMPLE_DATE, n=1)
    book.notes = [n for n in book.notes if not n.startswith(("Nhiệt độ", "Độ ẩm"))]
    # hiệu chỉnh nhiệt độ / độ ẩm: TREND theo giấy hiệu chuẩn (như phiếu người dùng) — bảng ở sheet ẩn HieuChuan
    ws_c = book.wb.create_sheet("HieuChuan")
    ws_c.sheet_state = "hidden"
    ranges = {}
    r = 1
    for key, tab, unit in (("t", CAL_T, "°C"), ("rh", CAL_RH, "%RH")):
        book.merge(ws_c, r, 1, r, 2, f"Giấy hiệu chuẩn {CAL_DEVICE} số {CAL_CERT} — "
                   f"{'nhiệt độ' if key == 't' else 'độ ẩm'}", bold=True, grid=False)
        book.put(ws_c, r + 1, 1, f"Giá trị chỉ thị, {unit}", bold=True)
        book.put(ws_c, r + 1, 2, f"Giá trị chuẩn, {unit}", bold=True)
        for k, (a, b) in enumerate(tab):
            book.put(ws_c, r + 2 + k, 1, a)
            book.put(ws_c, r + 2 + k, 2, b)
        ranges[key] = (r + 2, r + 1 + len(tab))
        r += len(tab) + 4
    ws_c.column_dimensions["A"].width = ws_c.column_dimensions["B"].width = 22
    ws_t = book.wb["Nhiệt Ẩm"]
    for row in ws_t.iter_rows(min_row=3):
        for c in row:
            if isinstance(c.value, str) and re.fullmatch(r"=[A-Z]+\d+", c.value):
                ref = c.value[1:]
                key = "t" if c.column == 6 else "rh"
                a, b = ranges[key]
                c.value = f"=TREND(HieuChuan!$B${a}:$B${b},HieuChuan!$A${a}:$A${b},{ref})"
    # chênh áp, lưu lượng: không có bảng hiệu chuẩn — lưu lượng nhân hệ số 0,981 (giấy hiệu chuẩn Testo 420)
    book.pressure([{"a": a, "b": b, "sign": s, "reads": DP_READS[int(s[-1])], "lim": None} for s, a, b, _e in DP_ROWS],
                  so=SO["chenh_ap"], date=SAMPLE_DATE)
    rooms = []
    for name, _en, area, _p, hepas, _n in ROOMS:
        rooms.append({"room": name, "area": area, "height": "-", "lim_ach": None,
                      "hepas": [{"name": f"HEPA {h}", "parts": [{"part": "", "date": SAMPLE_DATE, "time": "",
                                                                  "reads": FLOW_READS[h]}]} for h in hepas]})
    ws_f = book.flow(rooms, so=SO["luu_luong"], date=SAMPLE_DATE, factor=FLOW_FACTOR)
    head = {ws_f.cell(1, c).value: c for c in range(1, ws_f.max_column + 1)}
    for lbl in ("Phần", "Thời điểm", "Chiều cao, m", "Thể tích phòng, m³", "ACH", "Mức quy định", "Đánh giá"):
        ws_f.column_dimensions[L_(head[lbl])].hidden = True
    for row in ws_f.iter_rows(min_row=2):
        for c in row:
            if c.column in (head["Thể tích phòng, m³"], head["ACH"]) and isinstance(c.value, str) and c.value.startswith("="):
                c.value = "-"
    for ws in book.wb.worksheets:                       # mỗi phiếu in vừa 1 trang (26 vị trí)
        ws.page_setup.fitToHeight = 1
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - TinhToanPhongSach.xlsx")
    book.save(out)
    # tính lại bằng Excel COM, đối chiếu số TBKQ (AGENTS mục 11)
    try:
        import win32com.client
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        try:
            x = xl.Workbooks.Open(os.path.abspath(out))
            xl.CalculateFull()
            w = x.Worksheets("Nhiệt Ẩm")
            bad = []
            for rr in range(3, 3 + 40):
                loc = w.Cells(rr, 3).Value
                m = re.match(r"PAR (\d+)", str(loc or ""))
                if m:
                    p = int(m.group(1))
                    got = (w.Cells(rr, 8).Value, w.Cells(rr, 9).Value)
                    if (round_half_up(got[0], 1), round_half_up(got[1], 1)) != (T_VAL[p], RH_VAL[p]):
                        bad.append((p, got))
            x.Save()
            x.Close()
        finally:
            xl.Quit()
        if bad:
            raise SystemExit(f"[FATAL] [Data Mismatch] phiếu nhiệt ẩm khác TBKQ: {bad}")
    except ImportError as e:
        print(f"[build_098696] Excel COM warning: {e}")
    print(f"[build_098696] Excel -> {out}")
    return out


def split_edoc() -> list:
    """BBLM 2 trang; Ghi chép = '0. so do vi tri do + .pdf' (sơ đồ người dùng + mặt bằng KH ghi tay số liệu) — bỏ trang
    trắng, xoay đúng hướng (AGENTS 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    outs = []
    for name, src in ((f"{CODE6} - BBLM.pdf", SCAN_BBLM), (f"{CODE6} - Ghi Chep.pdf", SCAN_GC)):
        pdf = fitz.open()
        pdf.insert_pdf(fitz.open(src))
        finish_edoc(pdf, INPUT_DIR, name)
        p = os.path.join(edoc, name)
        pdf.save(p)
        outs.append(p)
        print(f"[build_098696] Edoc -> {p} ({pdf.page_count} tr.)")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--thu", action="store_true", help="chỉ dựng TBKQ VN thô để xem template")
    args = ap.parse_args()
    print(f"[build_098696] T {T_VAL}\nRH {RH_VAL}\nFLOW {FLOW}\nDP {DP}")
    layouts = build_layouts()
    d = build_dossier(layouts)
    d.personnel.reporter = REPORTER
    d.ensure_output_directories()
    if args.thu:
        print(CleanroomHandler(d).build_tbkq(is_en=False))
        return
    split_edoc()
    h = CleanroomHandler(d)
    vn_doc = finish_tbkq(h.build_tbkq(is_en=False), is_en=False)
    d.layouts = layouts["en"]
    en_doc = finish_tbkq(h.build_tbkq(is_en=True), is_en=True)
    d.layouts = layouts["vn"]
    bckq = build_bckq(d)
    xlsx = build_excel()
    docs = [bckq, vn_doc, en_doc]
    if args.pdf:
        from Core.Exporters.pdf_exporter import PDFExporter
        ex = PDFExporter(d)
        ex.export_files(docs)
        ex.export_excel_files([xlsx])
    print("\n===== QA =====")
    for p in docs:
        print(f"  {check_single_docx(p, code=CODE).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
