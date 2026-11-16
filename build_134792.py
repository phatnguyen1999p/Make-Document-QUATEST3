# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D134792 - Công ty TNHH Nippon Chemiphar Việt Nam (03/11/2026).

1 chứng thư (mã trần): Tủ cấy vi sinh - Clean Bench, pH Scientific, Model N/A, Serial N/A, mã QC-EQ-052
  -> BSCHandler: bỏ inflow, lọc khí thải, hình thái dòng khí, UVC (phiếu gạch chéo);
     có downflow, lux, ồn, HEPA lọc cấp, nồng độ hạt bụi, vi sinh không khí 6 vị trí (chờ KQTN).

Quyết định của người dùng (03/11/2026):
  - BBLM: "Theo yêu cầu của khách hàng: Không thể hiện đánh giá tủ" -> bảng kết quả 3 cột
    (STT | Chỉ tiêu | Kết quả đo), bỏ cả cột mức và cột đánh giá (dossier.has_assessment=False).
  - Vi sinh không khí (tổng số vi sinh vật hiếu khí, VT1 - VT6 theo sơ đồ trên phiếu): thêm vào bảng,
    kết quả + ngày ban hành để "………" tô vàng chờ KQTN (AGENTS mục 2, 13).
  - Downflow hàng 1: điểm 4 = 0,50, điểm 5 = 0,49 -> trung bình tính từ 21 điểm 0,43 m/s
    (phiếu ghi tay 0,41 không dùng, AGENTS mục 11).
  - Model: N/A theo BBLM (phiếu thông số tủ ghi "BSC12", báo giá ghi "BSC16").
  - Nội dung yêu cầu "Giám định tủ cấy" theo PYC/báo giá (BBLM + phiếu ghi chép in sẵn "tủ an toàn sinh học").
Quyết định của người dùng (12/11/2026):
  - "KQTN 134792 có rồi": KQTN VSTN26028703.01 - .07 ngày 12/11/2026 (Z `TBKQ/Edoc/KQ-Q326D134792-VSTN…`,
    chép về Input/134792/KQTN) -> điền vi sinh, ngày ban hành = 12/11/2026 (AGENTS mục 2, 13).
  - "Giữ như vậy" (12/11/2026): KQTN .01 - .06 = VT1 - VT6 theo thứ tự trên sơ đồ phiếu ghi chép, .07 = blank
    (phiếu KQTN không ghi ký hiệu mẫu). Kết quả: VT1 - VT5 < 1, VT6 6,0 × 10¹ CFU/đĩa, blank < 1 CFU/đĩa.
    Ngày yêu cầu cấp KQTN 11/11/2026 (PYC trống ngày dự kiến trả KQ -> ngày lấy mẫu + 8, AGENTS mục 13);
    KQTN 12/11/2026 = 7 ngày làm việc từ ngày gửi mẫu -> trễ hạn "không" (người dùng 12/11/2026: "Không đánh
    trễ kết quả thử nghiệm dưới 10 ngày làm việc" — trước đó ghi "có"). Ngày gửi mẫu = ngày nhận mẫu trên KQTN 03/11/2026.
Suy luận của AI (cần người dùng xác nhận):
  - Giờ thực hiện 09:30 - 11:30 theo BBLM.
  - Quy trình theo BBLM: QUATEST 3 3030:2025, N2.TC.BSC.02 (bỏ EN 12469:2000 - đã ở Căn cứ).
  - HEPA lọc cấp: 2 lần quét ATI 2i (0,0054 % và 0,0059 %) -> lấy 0,0059 % như phiếu ghi chép.

Số liệu đọc từ phiếu ghi chép (Input/134792/20261103140954604.pdf tr. 3-6) và phiếu in nhiệt
(Input/134792/20261103141009066.pdf: HEPA ATI 2i, AeroTrak 9350 sample 529/530).

Chạy:  python build_134792.py            (DOCX + XLSX, không PDF)
       python build_134792.py --pdf      (thêm xuất PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
Không đồng bộ ổ mạng Z: (network_dir để trống; chép lên Z bằng robocopy theo AGENTS mục 0).
"""

import os
import sys
import copy
import argparse

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)

import docx
import openpyxl
import fitz
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl.styles import PatternFill

from Core.Models import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.bsc import BSCHandler
from Core.Tools.bsc_calc import build_bsc_calculation_workbook
from Core.Tools.noise_corr import noise_correction_value
from Core.Tools.qa_check import check_single_docx
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

CODE = "Q326D134792"
REG_DATE = "29/10/2026"
SAMPLE_DATE = "03/11/2026"
PLACEHOLDER = "………"
KQTN_SO = "VSTN26028703.01 - .07"   # số phiếu KQTN vi sinh (.01-.06 = VT1-VT6, .07 = blank)
KQTN_SO_VT = "VSTN26028703.01 - .06"
KQTN_SO_BLANK = "VSTN26028703.07"
KQTN_NGAY = "12/11/2026"     # ngày phiếu KQTN dd/mm/yyyy -> ngày ban hành TBKQ (AGENTS mục 2)
NGAY_GUI_MAU = "03/11/2026"  # = "Ngày nhận mẫu" trên KQTN
NGAY_YC_KQTN = "11/11/2026"  # PYC trống ngày dự kiến trả KQ -> ngày lấy mẫu + 8 (AGENTS mục 13)
NGAY_NHAN_KQTN = "12/11/2026"
MICRO_BLANK = "< 1"
CERT_DATE = KQTN_NGAY or PLACEHOLDER
START, END = "09:30", "11:30"
INPUT_DIR = os.path.join(WORKSPACE, "Input", "134792")
PHOTO_DIR = os.path.join(INPUT_DIR, "Photo", "Use")
TBKQ_DIR = os.path.join(INPUT_DIR, "TBKQ")

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Nippon Chemiphar Việt Nam",
    name_en="Nippon Chemiphar Vietnam Co., Ltd.",
    address_vn="Số 76 Đại lộ Độc Lập, Khu công nghiệp Việt Nam - Singapore, phường An Phú, Thành phố Hồ Chí Minh",
    address_en="No. 76 Doc Lap Avenue, Vietnam - Singapore Industrial Park, An Phu Ward, Ho Chi Minh City",
)
LOCATION_VN = ("Nhà máy của Công ty TNHH Nippon Chemiphar Việt Nam - Số 76 Đại lộ Độc Lập, "
               "Khu công nghiệp Việt Nam - Singapore, phường An Phú, Thành phố Hồ Chí Minh")
LOCATION_EN = ("Factory of Nippon Chemiphar Vietnam Co., Ltd. - No. 76 Doc Lap Avenue, "
               "Vietnam - Singapore Industrial Park, An Phu Ward, Ho Chi Minh City")

EQUIPMENTS = [
    eq("kanomax_6501"),
    eq("ati_6d_2i"),
    eq("testo_440"),
    eq("rion_nl53"),
    eq("aerotrak_9350"),
]

# ---------------------------------------------------------------------------
# Số liệu phiếu ghi chép
# ---------------------------------------------------------------------------
DOWNFLOW = [[0.45, 0.39, 0.44, 0.50, 0.49, 0.46, 0.48],     # vách trong
            [0.35, 0.45, 0.42, 0.43, 0.44, 0.46, 0.55],
            [0.42, 0.40, 0.36, 0.36, 0.33, 0.38, 0.49]]     # cửa
DF_WIDTH, DF_DEPTH = 1200, 680
LUX = [[751.9, 731.2, 671.3, 624.7],
       [870.1, 830.3, 757.1, 673.5],
       [899.3, 945.6, 885.6, 761.5]]
NOISE_AMB, NOISE_OP = 54.4, 59.0          # file 0184 / 0183; ồn nền ≤ 55 dBA -> không hiệu chỉnh
HEPA_DF = 0.0059                          # ATI 2i: 0,0054 % và 0,0059 %, nồng độ 28 µg/L
PARTICLE = [1, 0, 0, 0]                   # AeroTrak sample 530 (zero 529), 1000 L, cộng dồn d ≥ 0,3 / 0,5 / 1,0 / 5,0 µm
MICRO_POINTS = ["VT1", "VT2", "VT3", "VT4", "VT5", "VT6"]
# Kết quả vi sinh theo KQTN (CFU/đĩa), ghi như KQTN (AGENTS mục 13); None -> "………" tô vàng
MICRO_RESULTS = {"VT1": "< 1", "VT2": "< 1", "VT3": "< 1", "VT4": "< 1", "VT5": "< 1", "VT6": "6,0 × 10¹"}
# Vị trí lấy mẫu trên sơ đồ phiếu (x: trái -> phải, y: cửa -> vách trong), tỉ lệ 0..1
MICRO_XY = {"VT1": (0.18, 0.78), "VT2": (0.18, 0.28), "VT3": (0.40, 0.56),
            "VT4": (0.60, 0.53), "VT5": (0.80, 0.78), "VT6": (0.80, 0.18)}


def _avg(grid):
    vals = [v for row in grid for v in row]
    return sum(vals) / len(vals)


def _photos():
    if not os.path.isdir(PHOTO_DIR):
        return []
    return [{"path": os.path.join(PHOTO_DIR, f), "caption_vn": "", "caption_en": ""}
            for f in sorted(os.listdir(PHOTO_DIR)) if f.lower().endswith((".jpg", ".jpeg", ".png"))]


def build_micro_diagram(is_en: bool) -> str:
    """Sơ đồ 6 vị trí lấy mẫu vi sinh (theo hình vẽ trên phiếu ghi chép), chữ Times New Roman (AGENTS mục 8C)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle
    plt.rcParams["font.family"] = "Times New Roman"
    out_dir = os.path.join(TBKQ_DIR, "_diagrams")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"vi_tri_vi_sinh_{'en' if is_en else 'vn'}.png")
    fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=200)
    ax.add_patch(Rectangle((0, 0), 1, 1, fill=False, lw=2, ec="black"))
    for name, (x, y) in MICRO_XY.items():
        ax.plot(x, y, marker="+", ms=14, mew=2, color="#1F4E9E")
        ax.text(x + 0.025, y - 0.02, name, fontsize=14, va="top", ha="left")
    ax.text(0.5, 1.04, "Back wall" if is_en else "Vách trong", ha="center", va="bottom", fontsize=14, style="italic")
    ax.text(0.5, -0.04, "Front opening" if is_en else "Cửa", ha="center", va="top", fontsize=14, style="italic")
    ax.set_xlim(-0.05, 1.05)
    ax.set_ylim(-0.18, 1.18)
    ax.set_aspect("auto")
    ax.axis("off")
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out


def dossier() -> Dossier:
    df_avg = _avg(DOWNFLOW)
    lux_avg = _avg(LUX)
    corr = noise_correction_value(NOISE_AMB, NOISE_OP)
    p_vn = [f"{v} hạt/m³" for v in PARTICLE]
    p_en = [f"{v} particles/m³" for v in PARTICLE]
    sizes = [(f"Kích thước hạt d ≥ {s.replace('.', ',')} µm", f"Particle size d ≥ {s} µm", p_vn[i], p_en[i],
              lim_vn, lim_en, "-", "-")
             for i, (s, lim_vn, lim_en) in enumerate((("0.3", "≤ 10 200 hạt/m³", "≤ 10,200 particles/m³"),
                                                       ("0.5", "≤ 3 520 hạt/m³", "≤ 3,520 particles/m³"),
                                                       ("1.0", "≤ 832 hạt/m³", "≤ 832 particles/m³"),
                                                       ("5.0", "≤ 29 hạt/m³", "≤ 29 particles/m³")))]
    micro_pts = [{"name_vn": p, "name_en": p,
                  "result_vn": MICRO_RESULTS[p] or PLACEHOLDER,
                  "result_en": (MICRO_RESULTS[p] or PLACEHOLDER).replace(",", ".")}
                 for p in MICRO_POINTS]
    rd = {
        "start_time": START, "end_time": END,
        "file_label": "Tu cay - KhongDanhGia", "bckq_label": "tu cay",
        "cabinet": {
            "cabinet_name": "Tủ cấy vi sinh",
            "cabinet_name_en": "Clean bench",
            "cap_tu": "-",
            "brand": "pH Scientific", "model": "N/A", "serial": "N/A", "device_id": "QC-EQ-052",
            "seal_item_name": "“Tủ cấy vi sinh”",
        },
        "inflow": {"skip": True},
        "downflow": {"readings": DOWNFLOW, "avg": round(df_avg, 2), "tolerance_pct": 20,
                     "width_mm": DF_WIDTH, "depth_mm": DF_DEPTH},
        "lux": {"readings": LUX, "avg": lux_avg, "avg_rounded": int(round(lux_avg))},
        "noise": {"ambient": NOISE_AMB, "operating": NOISE_OP, "corrected_dba": round(NOISE_OP - corr, 1),
                  "file_ambient": "0184", "file_operating": "0183"},
        "hepa": {"downflow_pct": HEPA_DF, "exhaust_pct": None, "skip_exhaust": True, "aerosol_conc": 28},
        "airflow_pattern": {"skip": True},
        "particle": {"clean_class": "ISO Class 5", "sample_volume_l": 1000.0, "sample_time_min": 20,
                     "readings": PARTICLE, "file_zero": "529", "file_ghi": "530", "sizes": sizes},
        "microbial": {
            "title_vn": "Thử khả năng lọc vô trùng của tủ",
            "title_en": "Sterile filtration test of the cabinet",
            "subtitle_vn": "Tổng số vi sinh vật hiếu khí (CFU/đĩa)",
            "subtitle_en": "Total aerobic microbial count (CFU/plate)",
            "task_vn": "Thử khả năng lọc vô trùng của tủ;",
            "task_en": "Sterile filtration test of the cabinet;",
            "points": micro_pts,
            "diagram": build_micro_diagram(False),
            "diagram_en": build_micro_diagram(True),
            "caption_vn": "Sơ đồ 1: Vị trí lấy mẫu thử khả năng lọc vô trùng trong tủ",
            "caption_en": "Diagram 1: Sampling locations for cabinet sterile filtration test",
            "heading_vn": "Sơ đồ vị trí lấy mẫu thử khả năng lọc vô trùng:",
            "heading_en": "Sterile filtration test sampling locations diagram:",
        },
        "tasks_vn": ["Đo tốc độ dòng khí cấp qua lọc HEPA;", "Đo cường độ sáng bên trong tủ;",
                     "Đo độ ồn của tủ khi vận hành;", "Kiểm tra độ rò rỉ của lọc HEPA khí cấp;",
                     "Đo nồng độ hạt bụi trong tủ;", "Thử khả năng lọc vô trùng của tủ."],
        "tasks_en": ["Downflow air velocity test;", "Illuminance test in the cabinet workspace;",
                     "Noise test of the cabinet in operation;", "HEPA filter leakage test;",
                     "Particle concentration measurement in the workspace;", "Sterile filtration test of the cabinet."],
    }
    v = lambda x: str(x).replace(".", ",")
    micro_txt = PLACEHOLDER if not all(MICRO_RESULTS.values()) else \
        "; ".join(f"{p}: {MICRO_RESULTS[p]} CFU/đĩa" for p in MICRO_POINTS)
    # BCKQ: khách hàng yêu cầu không thể hiện đánh giá -> mọi chỉ tiêu vào "Chỉ tiêu không đánh giá" (AGENTS mục 7)
    khong_dg = [
        "Theo yêu cầu của khách hàng không thể hiện đánh giá tủ (BBLM)",
        f"Tốc độ trung bình dòng khí cấp qua lọc HEPA {round(df_avg, 2):.2f} m/s".replace(".", ","),
        f"Cường độ sáng trung bình bên trong tủ {int(round(lux_avg))} lux",
        f"Độ ồn khi vận hành {v(rd['noise']['corrected_dba'])} dBA (độ ồn nền {v(NOISE_AMB)} dBA)",
        f"Mức rò rỉ của lọc khí cấp {HEPA_DF:.4f} %".replace(".", ","),
        "Nồng độ hạt bụi trong tủ (d ≥ 0,3 µm: " + p_vn[0] + "; d ≥ 0,5 µm: " + p_vn[1]
        + "; d ≥ 1,0 µm: " + p_vn[2] + "; d ≥ 5,0 µm: " + p_vn[3] + ")",
        f"Thử khả năng lọc vô trùng của tủ (tổng số vi sinh vật hiếu khí {micro_txt})",
    ]
    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=f"Ngày {SAMPLE_DATE} từ {START} đến {END}",
        sample_time_en=f"From {SAMPLE_DATE}",
        sub_no=None, is_multi_cert=False,
        location_desc_vn=LOCATION_VN, location_desc_en=LOCATION_EN,
        client=CLIENT, personnel=PersonnelInfo(executor="Phát"),
        input_dir=INPUT_DIR, tbkq_dir=TBKQ_DIR, output_dir="", network_dir="",
        dossier_type="bsc",
        has_assessment=False,
        scope_vn="Giám định tủ cấy", scope_en="Inspection of Clean bench",
        procedure_vn="QUATEST 3 3030:2025, N2.TC.BSC.02", procedure_en="QUATEST 3 3030:2025, N2.TC.BSC.02",
        standards_vn="EN 12469:2000; ISO 14644-1:2015", standards_en="EN 12469:2000; ISO 14644-1:2015",
        basis_vn="EN 12469:2000; ISO 14644-1:2015", basis_en="EN 12469:2000; ISO 14644-1:2015",
        equipments=EQUIPMENTS,
        locations=[MonitoringLocation(name_vn="Tủ cấy vi sinh pH Scientific (QC-EQ-052)",
                                      name_en="Clean bench pH Scientific (QC-EQ-052)")],
        photos=_photos(),
        results_data=rd,
        assessment={"basis_vn": "EN 12469:2000; ISO 14644-1:2015",
                    "phu_hop_vn": [], "khong_phu_hop_vn": [], "khong_danh_gia_vn": khong_dg},
    )


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX: tô vàng "………" (theo build_131671._highlight_placeholders)
# ---------------------------------------------------------------------------
def highlight_placeholders(path: str):
    doc = docx.Document(path)
    roots = [doc.element.body]
    for s in doc.sections:
        roots += [s.header._element, s.footer._element, s.first_page_header._element]
    n = 0
    for root in roots:
        for r in list(root.iter(qn("w:r"))):
            t = r.find(qn("w:t"))
            if t is None or PLACEHOLDER not in (t.text or "") or len(r.findall(qn("w:t"))) != 1:
                continue
            parts = t.text.split(PLACEHOLDER)
            parent, idx = r.getparent(), r.getparent().index(r)
            k = 0
            for i, part in enumerate(parts):
                for txt, hl in ([(part, False)] if part else []) + ([(PLACEHOLDER, True)] if i < len(parts) - 1 else []):
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
                        n += 1
                    parent.insert(idx + 1 + k, nr)
                    k += 1
            # giữ tab/br của run gốc ở đầu
            extras = [c for c in r if c.tag not in (qn("w:rPr"), qn("w:t"))]
            if extras:
                lead = copy.deepcopy(r)
                for c in list(lead):
                    if c.tag == qn("w:t"):
                        lead.remove(c)
                parent.insert(idx, lead)
            parent.remove(r)
    doc.save(path)
    print(f"[134792] Tô vàng {n} ô '{PLACEHOLDER}' -> {os.path.basename(path)}")


def _ptext(p_el) -> str:
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


def _para_with_text(proto, text: str):
    """Chép đoạn mẫu, giữ pPr + định dạng run đầu, thay nội dung (tab giữ bằng \\t)."""
    p = copy.deepcopy(proto)
    runs = p.findall(qn("w:r"))
    rpr = runs[0].find(qn("w:rPr")) if runs else None
    for r in runs:
        p.remove(r)
    for k, part in enumerate(text.split("\t")):
        if k:
            r = OxmlElement("w:r")
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            r.append(OxmlElement("w:tab"))
            p.append(r)
        if part:
            r = OxmlElement("w:r")
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            t = OxmlElement("w:t")
            t.text = part
            t.set(qn("xml:space"), "preserve")
            r.append(t)
            p.append(r)
    return p


def add_kqtn_block(path: str):
    """BCKQ: khối "Gửi mẫu thử nghiệm" (số / ngày KQTN, ngày gửi mẫu, yêu cầu cấp, nhận, trễ hạn) đặt trước
    "Báo cáo của GĐV" — theo BCKQ 131671 người dùng sửa (AGENTS mục 7, 13)."""
    import datetime
    dd = lambda s: datetime.datetime.strptime(s, "%d/%m/%Y").date()
    doc = docx.Document(path)
    body = doc.element.body
    gdv = next(p for p in body.iterchildren(qn("w:p")) if _ptext(p).strip().startswith("Báo cáo của GĐV"))
    proto = next(p for p in body.iterchildren(qn("w:p")) if _ptext(p).strip().startswith("Ngày kết thúc hồ sơ"))
    from Core.Tools.kqtn_late import tre_han_text   # ≤ 10 ngày làm việc từ ngày gửi mẫu: không trễ (AGENTS 13)
    tre = tre_han_text(KQTN_NGAY, NGAY_GUI_MAU, NGAY_YC_KQTN, PLACEHOLDER)
    lines = [_para_with_text(gdv, "Gửi mẫu thử nghiệm:")] + [_para_with_text(proto, s) for s in (
        f"Phiếu KQTN số {KQTN_SO or PLACEHOLDER} ngày {KQTN_NGAY or PLACEHOLDER}",
        f"Ngày gửi mẫu: {NGAY_GUI_MAU or PLACEHOLDER}",
        f"Ngày yêu cầu cấp KQTN: {NGAY_YC_KQTN or PLACEHOLDER}",
        f"Ngày nhận KQTN: {NGAY_NHAN_KQTN or PLACEHOLDER}",
        f"KQTN có trễ hạn so với yêu cầu của GĐV hay không? {tre}")]
    for el in lines:
        gdv.addprevious(el)
    from docx.text.paragraph import Paragraph
    for el in lines[:-1]:
        Paragraph(el, None).paragraph_format.keep_with_next = True   # python-docx giữ đúng thứ tự pPr
    doc.save(path)


# ---------------------------------------------------------------------------
# Phiếu tính toán Excel
# ---------------------------------------------------------------------------
def _calc_common(d: Dossier) -> dict:
    c = dict(d.results_data)
    c.update({
        "code": d.code, "sub_no": d.sub_no, "reg_date": d.reg_date, "sample_date": d.sample_date,
        "cert_date": d.cert_date, "start_time": START, "end_time": END, "photos": d.photos, "input_dir": d.input_dir,
        "client_name_vn": d.client.name_vn, "client_name_en": d.client.name_en,
        "client_address_vn": d.client.address_vn, "client_address_en": d.client.address_en,
        "scope_vn": d.scope_vn, "scope_en": d.scope_en,
        "location_vn": d.location_desc_vn, "location_en": d.location_desc_en,
        "leader_name_vn": "Lê Đăng Huyền", "gdv_name_vn": "Nguyễn Xuân Nam",
        "leader_name_en": "Le Dang Huyen", "gdv_name_en": "Nguyen Xuan Nam",
        "equipments": [e.name_vn for e in d.equipments], "equipments_en": [e.name_en for e in d.equipments],
    })
    cab = d.results_data.get("cabinet", {})
    for k in ("cabinet_name", "brand", "model", "serial", "device_id"):
        c[k] = cab.get(k, "")
    return c


def build_excel(d: Dossier) -> str:
    out = os.path.join(d.thuchien_dir, f"{d.code} - TinhToanTuCay.xlsx")
    os.makedirs(d.thuchien_dir, exist_ok=True)
    calc = _calc_common(d)
    # Inflow không thực hiện: truyền giá trị giả cho bộ dựng (phương pháp 1), sau đó gạch "-" (như 132488/1)
    calc["inflow"] = {"method": 1, "width_mm": 1000, "height_mm": 1000, "q_readings": [3600]}
    build_bsc_calculation_workbook(calc, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Bảng tính"]
    ws["B4"] = "Tủ cấy (Clean bench)"
    ws["H5"] = d.results_data["cabinet"]["device_id"]
    # 1. Inflow: không thực hiện
    for ref in ("A12", "B12", "C12", "D12", "E12", "F12", "G12", "J12", "A14", "A15", "D31", "F31", "H31", "J31"):
        ws[ref] = "-"
    ws["A10"] = "Phương pháp 1:  DIM - TSI 8380"   # mặc định TSI 8380 (AGENTS mục 8C, 11/11/2026)
    ws["A28"] = "Đánh giá (không thực hiện chỉ tiêu này)"
    # Khách hàng yêu cầu không thể hiện đánh giá: ô đánh giá ghi "-" (mức giữ để tham khảo)
    for ref in ("J46", "J48", "J62", "J76", "H83", "D136"):
        ws[ref] = "-"
    ws["A79"] = None
    # 5. HEPA: chỉ có lọc khí cấp
    ws["D87"], ws["F87"], ws["H87"] = "-", "-", "-"
    # 6. Hình thái dòng khí: không thực hiện
    for row in (101, 106):
        for col in "DHL":
            ws[f"{col}{row}"] = "-"
    ws["A97"] = "6. Kiểm tra hình thái dòng khí (không thực hiện)"
    # Sheet TBKQ-VN / TBKQ - ENG: bỏ kết luận Đạt/Pass
    for name in ("TBKQ-VN", "TBKQ - ENG"):
        if name in wb.sheetnames:
            for row in wb[name].iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.strip() in ("Đạt", "Pass", "Không đạt", "Fail"):
                        cell.value = "-"
    wb.save(out)
    _recalc(out)
    print(f"[134792] Excel -> {out}")
    return out


def _recalc(path: str):
    """Tính lại công thức bằng Excel COM (openpyxl lưu lại làm mất giá trị đã tính)."""
    try:
        import gc
        import win32com.client
    except ImportError:
        print("[134792] Không có pywin32 - bỏ qua tính lại Excel")
        return
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(path))
        excel.CalculateFull()
        wb.Close(SaveChanges=True)
        wb = None
    finally:
        excel.Quit()
        del excel
        gc.collect()


# ---------------------------------------------------------------------------
# Edoc: tách scan (AGENTS mục 8A)
# ---------------------------------------------------------------------------
def build_edoc():
    edoc = os.path.join(TBKQ_DIR, "Edoc")
    os.makedirs(edoc, exist_ok=True)
    src = fitz.open(os.path.join(INPUT_DIR, "20261103140954604.pdf"))       # tr.1-2 BBLM, tr.3-6 ghi chép
    thermal = fitz.open(os.path.join(INPUT_DIR, "20261103141009066.pdf"))   # phiếu in nhiệt HEPA + AeroTrak
    bblm = fitz.open()
    bblm.insert_pdf(src, from_page=0, to_page=1)
    from Core.Tools.scan_splitter import finish_edoc   # bỏ trang trắng + xoay đúng hướng (AGENTS 8A)
    finish_edoc(bblm, INPUT_DIR, "134792 - BBLM.pdf")
    bblm.save(os.path.join(edoc, "134792 - BBLM.pdf"))
    gc_ = fitz.open()
    gc_.insert_pdf(src, from_page=2, to_page=5)
    gc_.insert_pdf(thermal)
    finish_edoc(gc_, INPUT_DIR, "134792 - Ghi Chep.pdf")
    gc_.save(os.path.join(edoc, "134792 - Ghi Chep.pdf"))
    print(f"[134792] Edoc: BBLM {bblm.page_count} tr., Ghi Chep {gc_.page_count} tr.")


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    args = ap.parse_args()

    build_edoc()
    d = dossier()
    d.ensure_output_directories()
    handler = BSCHandler(d)
    vn, en = handler.build_tbkq(is_en=False), handler.build_tbkq(is_en=True)
    tem = handler.build_tem_tu()
    xlsx = build_excel(d)
    # BCKQ: thêm viện dẫn phiếu KQTN vi sinh (AGENTS mục 7)
    b = BCKQBuilder(d)
    ctx = b.build_context()
    kq_ngay = KQTN_NGAY or PLACEHOLDER
    ctx["tham_chieu"].append(f"- Kết quả thử khả năng lọc vô trùng của tủ (tổng số vi sinh vật hiếu khí tại VT1 - VT6): "
                             f"xem tại phiếu Kết quả thử nghiệm số {KQTN_SO_VT if KQTN_SO else PLACEHOLDER} "
                             f"ngày {kq_ngay}.")
    ctx["tham_chieu"].append("+ KQ được biểu thị < 1 CFU/đĩa → không phát hiện vi sinh vật mục tiêu trên đĩa lấy mẫu.")
    ctx["tham_chieu"].append(f"+ KQTN của mẫu Blank (phiếu số {KQTN_SO_BLANK if KQTN_SO else PLACEHOLDER}): "
                             f"{MICRO_BLANK or PLACEHOLDER} CFU/đĩa → môi trường lấy mẫu đảm bảo kết quả.")
    bckq = b.build(ctx=ctx, out_name=f"BCKQ - 1 - tu cay - {CODE}.docx")
    add_kqtn_block(bckq)
    docs = [bckq, vn, en, tem]
    for p in docs:
        highlight_placeholders(p)
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
