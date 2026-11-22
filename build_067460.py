# -*- coding: utf-8 -*-
"""
Hồ sơ Q326D067460 - Công ty TNHH Thương mại Khoa học Kỹ thuật Lâm Việt — 2 buồng LAF tại nhà máy Công ty TNHH TM-SX
Việt Thọ: đợt 18/06/2026 (downflow, độ sáng, độ ồn) + đợt 11/11/2026 (quét rò rỉ HEPA, nồng độ bụi).
Template PHÒNG SẠCH (người dùng chọn 11/11/2026), mỗi buồng LAF 1 chứng thư (/1, /2), mỗi chứng thư gồm số liệu cả
2 biên bản; chỉ TBKQ tiếng Việt (PYC mục 15 chỉ chọn Tiếng Việt).

Nguồn (Input/067460, chép từ Z 11/11/2026 — chỉ đọc):
  - KQDo/20261111161732829.pdf (scan 26 trang):
      tr.1 phiếu kiểm soát công việc; tr.3 PYC Q326D067460 ngày 09/06/2026 (2 LAF: "LAF cân mẫu 1 mặt nhựa"
      W1600×D1600×H2400 mm, "LAF lấy mẫu 1 mặt nhựa" W1200×D1200×H2400 mm; downflow, rò rỉ ULPA, ồn, độ sáng, bụi
      Class A; chứng thư Tiếng Việt; MST 0304125427 — đối chiếu tax_helper 20/11/2026: EXACT_MATCH tên + địa chỉ
      477/21 Phạm Văn Chiêu, phường An Hội Đông, TP Hồ Chí Minh, tên quốc tế LAM VIET SCIENTIFIC TECHNOLOGY TRADING
      COMPANY LIMITED, đang hoạt động); tr.5 bản PYC thứ 2; tr.7–9 báo giá QONV26009397;
      tr.11–12 BBLM 11/11/2026 (BM3030.02): "Quét rò rỉ màng lọc HEPA và đo nồng độ bụi khu vực phòng sạch",
        09:00 – 11:30, AeroTrak 9350 + ATI 6D/2i, QUATEST 3 3030:2025; N2.TC.PS.01; "Buồng LAF tại phòng vi sinh bao bì
        và tại phòng cân nguyên liệu": 04 HEPA, bụi 02 vị trí; người thực hiện Nguyễn Tiến Phát;
      tr.13–15 ghi chép 11/11/2026 (BM3030.01): bụi DPC 9350, 1 m³, zero 575, "LAF cân NL" 1 -> 576,
        "LAF vi sinh bao bì" 2 -> 577; HEPA: Cân NL - Phải 0,0030 %, Cân NL - Trái 0,0029 %, Vi sinh bao bì - Trái
        0,0042 %, Vi sinh bao bì - Phải 0,0023 %; cấp gió trên xuống, hồi tường, nhà xưởng không sản xuất (nghỉ);
      tr.17–18 BBLM 18/06/2026 (BM3030.02, N2.TC.LAF.01, 2 buồng LAF /1 vi sinh bao bì, /2 cân NL);
        10:00 – 16:00, Testo 440, Kanomax 6501, Rion NL-53, ATI, AeroTrak; người thực hiện Nguyễn Tiến Phát;
      tr.19–21 ghi chép thông số tủ /2 (cân NL), tr.23–25 tủ /1 (vi sinh bao bì): downflow, độ sáng, ồn ngày 18/06;
        mục HEPA và bụi ghi "Thực hiện 11/11/2026", downflow + bụi "Theo GMP Grade A" (downflow 0,36 – 0,54 m/s);
        độ sáng / ồn mức in sẵn EN 12469:2000 (≥ 750 lux; ≤ 65 dBA); inflow, hình thái dòng khí, UVC gạch bỏ.
        Downflow 16 điểm (Kanomax, cách lưới phân phối khí 30 cm), độ sáng 9 điểm, ồn: cân NL nền 47,2 (file 0008) /
        vận hành 76,0 dBA (file 0005); vi sinh bao bì vận hành 74,8 dBA (file 0004), không ghi ồn nền.
        Tủ /2 mục HEPA ghi "HEPA cấp - Phải 0,0029; HEPA cấp - Trái 0,0030" (ngược phiếu ghi chép 11/11).
      tr.2, 4, 6, 10, 16, 22, 26 trắng.
  - KQDo/20261111161811921.pdf (3 trang): phiếu in AeroTrak 9350 (575 zero 250 L; 576 Phòng cân nguyên liệu
    10:12–10:32 1000 L; 577 Phòng vi sinh bao bì 10:55–11:15 1000 L — mọi kênh Diff / Cumul = 0 hạt/m³);
    phiếu in ATI 2i (mỗi lọc 2 lần quét: trên = gasket, dưới = media; kết quả = lớn nhất — khớp phiếu ghi chép):
      Cân NL trái 0,0029 / 0,0014; phải 0,0030 / 0,0017; Vi sinh bao bì trái 0,0021 / 0,0042; phải 0,0023 / 0,0022.
  - KQDo/Data.xlsx + Auto_0018..0021 (Rion NL-53, ngày máy 06/08/2026, 58,7 – 63,8 dBA) — số file không khớp phiếu
    (0004, 0005, 0008) -> KHÔNG dùng, độ ồn lấy theo số ghi tay trên phiếu.
  - Photos-1-001 (2)/1/Use, /2/Use: 6 ảnh người dùng chọn (gồm 2 ảnh 18/06 đo downflow bằng Kanomax).
  - Thuc hien/*.docx|pdf: form BBLM / ghi chép trống (do công cụ khác tạo) — không dùng.

Quyết định của người dùng (11/11/2026):
  - Dùng template phòng sạch.
  - Phiếu in ATI: giá trị trên là gasket, giá trị dưới là media.
  - Phạm vi chứng thư gồm cả 2 biên bản (18/06 và 11/11). Các điểm (2) – (5) dưới: "ok".
  - Kết quả độ sáng và độ ồn tạm thời không đánh giá (công tắc DANH_GIA_SANG_ON = False).
  - Vẽ sơ đồ 16 điểm đo tốc độ gió; mục 9.3 TBKQ là bảng số liệu dẫn chứng từng điểm.
  - Mỗi LAF là 1 kết quả: /1 buồng LAF phòng vi sinh bao bì, /2 buồng LAF phòng cân nguyên liệu (theo ký hiệu ghi
    chép BBLM 18/06); BCKQ, phiếu tính toán tách theo /1, /2. Bản gộp (mã trần) dựng trước đó bỏ — file đã xóa ở
    Input, code chỉ còn dựng từng buồng (14/11/2026).
  - (20/11/2026) Dựng lại theo quy tắc mới; ảnh: "nếu 3 hình thì vẫn để 3 cột" (giữ 1 hàng 3 cột).
  - (17/11/2026) Sơ đồ vị trí đo tốc độ gió: vị trí đo cân với tấm lọc HEPA (cách đều mép lọc); dựng bằng draw.io để
    sửa được (trước đó vẽ PIL, không có file .drawio).
  - (16/11/2026) Dòng "Nồng độ hạt bụi: 01 vị trí (2)" ở bảng 9.1 khó hiểu ("(2)" đọc như ghi chú) -> bỏ số vị trí;
    ký hiệu "(1)" ở dòng Ghi chú phải in nhỏ (chỉ số trên) — sửa chung trong Core (superscript_note_markers).
  - Thống nhất gọi "vị trí đo" / "vị trí", không dùng "điểm đo". Mục cường độ sáng ghi "Cường độ sáng trung bình",
    không ghi số lượng vị trí.

Suy luận của AI:
  (1) Downflow 0,36 – 0,54 m/s theo WHO TRS 961 (GMP Grade A, "Theo GMP Grade A" trên phiếu). Sơ đồ điểm đo: lưới 4 × 4,
      đánh số theo hàng từ vách trong, trái -> phải (thứ tự ghi trên phiếu), lọc trái cột 1–2, phải cột 3–4.
  (2) Cấp sạch GMP Grade A ("Theo GMP Grade A" trên phiếu), giới hạn bụi WHO TRS 961 (2011) – Phụ lục 6 trạng thái nghỉ
      (d ≥ 0,5 µm ≤ 3 520; d ≥ 5 µm ≤ 20 hạt/m³) như hồ sơ 127215; HEPA ≤ 0,01 % (ISO 14644-3:2019).
  (3) Diện tích theo PYC: "LAF cân mẫu" 1,6 × 1,6 = 2,56 m² ↔ buồng tại phòng cân nguyên liệu;
      "LAF lấy mẫu" 1,2 × 1,2 = 1,44 m² ↔ buồng tại phòng vi sinh bao bì.
  (4) "TM-SX" (Việt Thọ) mở rộng thành "Thương mại - Sản xuất" (chưa có MST Việt Thọ để tra).
  (5) Lọc trái / phải của buồng cân NL theo phiếu ghi chép 11/11 (nhãn viết tay trên phiếu in ATI bị sửa, khó đọc;
      phiếu thông số tủ /2 ghi ngược lại).

Chạy:  python build_067460.py          (DOCX + XLSX + Edoc)
       python build_067460.py --pdf    (thêm PDF qua Word/Excel COM + UpChungThu, chỉ Windows)
       --tao-lai-so-do                 (dựng lại TBKQ/_diagrams/so_do_vi_tri.drawio từ tọa độ code — bỏ chỉnh tay)
Sơ đồ 9.3: TBKQ/_diagrams/so_do_vi_tri.drawio — mở bằng draw.io, kéo vị trí đo ở tab so_do_vi_tri_vn, Ctrl+S,
chạy lại script (file đã có thì code giữ nguyên chỉnh tay).
Không đồng bộ Z trong code — chép bằng `python Core/Tools/ai_files.py chep 067460 --ghi` (AGENTS mục 0).
"""

import os
import re
import sys
import argparse
from datetime import date

sys.stdout.reconfigure(encoding="utf-8")
WORKSPACE = os.path.dirname(os.path.abspath(__file__))
if WORKSPACE not in sys.path:
    sys.path.insert(0, WORKSPACE)
os.environ.setdefault("QUATEST3_NETWORK_ROOT", os.path.join(WORKSPACE, "_khong_dong_bo_Z"))

from docx import Document
from docx.oxml.ns import qn

from Core.Models import Dossier, ClientInfo, PersonnelInfo, MonitoringLocation
from Core.Builders.bckq_builder import BCKQBuilder
from Core.Builders.handlers.cleanroom import CleanroomHandler
from Core.Builders.tag_context import verdict, style_note_tables
from Core.Tools.qa_check import check_single_docx
from Core.Tools.cleanroom_calc import CleanroomCalcBook
from Core.Tools.equipment_names import eq  # tên máy chuẩn (AGENTS mục 14)

import build_127215 as b127   # dùng lại các hàm hậu xử lý DOCX (bảng bụi GMP, giữ khối kết quả…)

CODE = "Q326D067460"
CODE6 = "067460"
REG_DATE = "09/06/2026"
SAMPLE_DATE = "11/11/2026"
START, END = "09:00", "11:30"
CERT_DATE = date.today().strftime("%d/%m/%Y")      # không có KQTN -> ngày dựng (AGENTS mục 2)
INPUT_DIR = os.path.join(WORKSPACE, "Input", CODE6)
SCAN_A = os.path.join(INPUT_DIR, "KQDo", "20261111161732829.pdf")
SCAN_B = os.path.join(INPUT_DIR, "KQDo", "20261111161811921.pdf")
PHOTO_DIRS = [os.path.join(INPUT_DIR, "Photos-1-001 (2)", n, "Use") for n in ("1", "2")]

CLIENT = ClientInfo(
    name_vn="Công ty TNHH Thương mại Khoa học Kỹ thuật Lâm Việt",
    name_en="Lam Viet Scientific Technology Trading Company Limited",      # tên quốc tế theo MST 0304125427
    address_vn="477/21 Phạm Văn Chiêu, phường An Hội Đông, Thành phố Hồ Chí Minh",
    address_en="477/21 Pham Van Chieu street, An Hoi Dong ward, Ho Chi Minh City",   # AGENTS mục 14 (20/11/2026)
)
SITE_VN = ("Nhà máy của Công ty TNHH Thương mại - Sản xuất Việt Thọ - Lô A1, Khu công nghiệp Sông Mây, đường số 4, "
           "xã Bình Minh, Thành phố Đồng Nai")
SITE_EN = ("Factory of Viet Tho Trading - Manufacturing Company Limited - Lot A1, Song May Industrial Park, Road No. 4, "
           "Binh Minh commune, Dong Nai City")

GRADE = "A"
BASIS_VN, BASIS_EN = "WHO TRS 961 (2011) – Phụ lục 6", "WHO TRS 961 (2011) – Annex 6"
LIM05, LIM5 = 3520, 20                     # GMP Grade A, trạng thái nghỉ, hạt/m³ (d ≥ 0,5 µm; d ≥ 5 µm)
HEPA_LIM = "≤ 0,01"

# Phiếu in AeroTrak 9350-02 S/N 93501302003 (Σ = Cumul, hạt/m³; mọi kênh 0)
def _rec(no, zone, start, stime, vol, vals):
    keys = ("Ch1", "Ch2", "Ch3", "Ch4", "Ch5", "Ch6")
    r = {"#": no, "Zone": zone, "Location": "Location01", "Date/Time (dd/mm/yyyy)": f"{SAMPLE_DATE} {start}",
         "Device Status": "OK", "Sample Time": stime, "Flow Status": "OK", "Laser Status": "OK", "Vol.": vol}
    r.update(dict(zip(keys, vals)))
    return r


ZERO = _rec(575, "067460 - zero", "09:44:52", "00:05:00", 250.0, [0] * 6)
# Buồng LAF: (tên VN, tên EN, diện tích m², vị trí bụi, bản ghi AeroTrak, [(lọc, gasket %, media %)] trái, phải)
# Thứ tự = số chứng thư (người dùng chốt 11/11/2026: mỗi LAF 1 kết quả) theo ký hiệu ghi chép BBLM 18/06:
# /1 buồng LAF phòng vi sinh bao bì, /2 buồng LAF phòng cân nguyên liệu. Ảnh: Photos-1-001 (2)/<n>/Use.
LAFS = [
    ("Buồng LAF tại phòng vi sinh bao bì", "LAF booth in packaging microbiology room", "1,44", 2,
     _rec(577, "067460 do", "10:55:15", "00:20:00", 1000.0, [0] * 6),
     [("Lọc trái", "0,0021", "0,0042"), ("Lọc phải", "0,0023", "0,0022")]),
    ("Buồng LAF tại phòng cân nguyên liệu", "LAF booth in raw material weighing room", "2,56", 1,
     _rec(576, "067460 do", "10:12:17", "00:20:00", 1000.0, [0] * 6),
     [("Lọc trái", "0,0029", "0,0014"), ("Lọc phải", "0,0030", "0,0017")]),
]
AREA = LAFS[0]             # buồng của chứng thư đang dựng (main() đặt lại)
CUR_N = 1                  # số chứng thư đang dựng (/1, /2)
# HEPA: (vị trí, gasket %, media %) — phiếu in ATI 2i: lần quét TRÊN = gasket, lần DƯỚI = media (người dùng chốt
# 11/11/2026). Kết quả trên chứng thư = max(gasket, media), khớp phiếu ghi chép.


def _max_pct(f) -> str:
    return max(f[1], f[2], key=lambda s: float(s.replace(",", ".")))


# --- Đợt 18/06/2026 (phiếu thông số tủ; lưới đọc theo hàng, trái -> phải) ---
JUNE_DATE = "18/06/2026"
J_START, J_END = "10:00", "16:00"
JUNE = {
    "Buồng LAF tại phòng cân nguyên liệu": {
        "df": [0.41, 0.42, 0.40, 0.56, 0.42, 0.39, 0.44, 0.45, 0.42, 0.42, 0.43, 0.39, 0.36, 0.42, 0.38, 0.43],
        "lux": [1090.5, 1200.2, 1001.4, 1025.5, 1248.1, 1111.0, 866.8, 1005.0, 892.8],
        "bg": 47.2, "op": 76.0},
    "Buồng LAF tại phòng vi sinh bao bì": {
        "df": [0.43, 0.42, 0.42, 0.49, 0.45, 0.40, 0.35, 0.48, 0.34, 0.32, 0.43, 0.44, 0.53, 0.37, 0.41, 0.46],
        "lux": [652.5, 660.2, 740.6, 678.2, 740.2, 658.8, 615.7, 600.6, 605.1],
        "bg": None, "op": 74.8},      # phiếu không ghi ồn nền
}
DF_LIM, LUX_LIM, NOISE_LIM = "0,36 - 0,54", "≥ 750", "≤ 65"
DF_LO, DF_HI, LUX_MIN, NOISE_MAX = 0.36, 0.54, 750, 65


def _vn(x: float, nd: int) -> str:
    return f"{x:.{nd}f}".replace(".", ",")


def df_avg(name):
    from Core.Tools.cleanroom_calc import round_half_up
    v = JUNE[name]["df"]
    return round_half_up(sum(v) / len(v), 2)


def lux_avg(name):
    from Core.Tools.cleanroom_calc import round_half_up
    v = JUNE[name]["lux"]
    return round_half_up(sum(v) / len(v), 1)


def noise_val(name):
    from Core.Tools.noise_corr import noise_correction_value
    j = JUNE[name]
    return round(j["op"] - (noise_correction_value(j["bg"], j["op"]) if j["bg"] is not None else 0), 1)


# Người dùng chốt 11/11/2026: kết quả độ sáng và độ ồn TẠM THỜI KHÔNG ĐÁNH GIÁ -> bỏ dòng mức + câu Đánh giá (TBKQ),
# không liệt kê ở Đánh giá BCKQ, Excel mức / đánh giá "-". Bật lại (mức EN 12469: ≥ 750 lux, ≤ 65 dBA) -> True.
DANH_GIA_SANG_ON = False
STANDARDS = ("ISO 14644-1:2015; WHO TRS 961 (2011) – Annex 6; ISO 14644-3:2019"
             + ("; EN 12469:2000" if DANH_GIA_SANG_ON else ""))


def june_checks(name):
    """[(tên chỉ tiêu, giá trị in, đạt? | None = không đánh giá)] của một buồng."""
    d, lx, n = df_avg(name), lux_avg(name), noise_val(name)
    return [("Tốc độ trung bình dòng khí cấp qua lọc HEPA", f"{_vn(d, 2)} m/s", DF_LO <= d <= DF_HI),
            ("Cường độ sáng trung bình", f"{_vn(lx, 1)} lux", (lx >= LUX_MIN) if DANH_GIA_SANG_ON else None),
            ("Độ ồn", f"{_vn(n, 1)} dBA", (n <= NOISE_MAX) if DANH_GIA_SANG_ON else None)]


AEROTRAK_LABELS = {"Ch1": "0,3 µm\n(hạt/m³)", "Ch2": "0,5 µm\n(hạt/m³)", "Ch3": "1,0 µm\n(hạt/m³)",
                   "Ch4": "2,0 µm\n(hạt/m³)", "Ch5": "3,0 µm\n(hạt/m³)", "Ch6": "5,0 µm\n(hạt/m³)",
                   "#": "Sample #", "Vol.": "Volume (L)"}

# cấu hình cho các hàm hậu xử lý dùng lại của build_127215
b127.GRADE_LIMITS = {GRADE: (LIM05, LIM5)}
b127.GROUP_ORDER = (GRADE,)
b127.BASIS_VN, b127.BASIS_EN = BASIS_VN, BASIS_EN


def _pct(s: str) -> float:
    return float(s.replace(",", "."))


def dust_ok() -> bool:
    return AREA[4]["Ch2"] <= LIM05 and AREA[4]["Ch6"] <= LIM5


def hepa_ok() -> bool:
    return all(_pct(_max_pct(f)) <= 0.01 for f in AREA[5])


def build_dossier() -> Dossier:
    photos = [{"path": b127._optimized(os.path.join(dp, f)), "caption_vn": "", "caption_en": ""}
              for dp in [PHOTO_DIRS[CUR_N - 1]] for f in sorted(os.listdir(dp)) if f.lower().endswith(".jpg")]
    locations = [MonitoringLocation(
        name_vn=vn, name_en=en, area_m2=area, clean_class=GRADE,
        state_vn="Trạng thái nghỉ", state_en="At rest",
        points_summary_vn=["Tốc độ dòng khí qua lọc HEPA: 16 vị trí;",   # cách không ngắt: không tách "vị / trí"
                           "Cường độ sáng trung bình;",
                           "Độ ồn: 01 vị trí;",
                           "Nồng độ hạt bụi: 01 vị trí;",   # bỏ "(2)": dễ đọc nhầm thành ghi chú (16/11/2026)
                           "Độ rò rỉ lọc HEPA: 02 lọc (trái, phải)."],
        points_summary_en=["Downflow velocity: 16 locations;", "Average light intensity;", "Noise level: 01 location;",
                           "Particle concentration: 01 location;",
                           "HEPA filter leakage: 02 filters (left, right)."],
    ) for vn, en, area, pos, _, _ in [AREA]]

    name, rec = AREA[0], AREA[4]
    bui = [{"cap": GRADE, "gioi_han": [LIM05, "", LIM5], "trang_so_do": "",
            "diem": [{"khu_vuc": name, "vi_tri": AREA[3], "gia_tri": [rec["Ch2"], rec["Ch3"], rec["Ch6"]]}],
            "ket_luan_vn": verdict(dust_ok(), False), "ket_luan_en": verdict(dust_ok(), True)}]
    # Mục downflow / độ sáng / độ ồn dựng bằng khung bảng HEPA của template (3 cột Khu vực | Vị trí | Giá trị),
    # mỗi mục một nhóm "hepa"; tiêu đề, tên cột, dòng mức, câu đánh giá sửa ở _fix_extra_sections.
    hepa = [{"muc": f"≤ {900 + i}", "trang_so_do": "", "ket_luan_vn": "phù hợp",
             "diem": [{"khu_vuc": name, "vi_tri": vt, "gia_tri": val}]}
            for i, (vt, val) in enumerate((("16 vị trí", _vn(df_avg(name), 2)),
                                           ("Bên trong buồng", _vn(lux_avg(name), 1)),
                                           ("Bên trong buồng", _vn(noise_val(name), 1))), 1)]
    hepa.append({"muc": HEPA_LIM, "trang_so_do": "",
                 "diem": [{"khu_vuc": name, "vi_tri": f[0], "gia_tri": _max_pct(f)} for f in AREA[5]]})

    return Dossier(
        code=CODE, cert_date=CERT_DATE, sample_date=SAMPLE_DATE, reg_date=REG_DATE,
        sample_time_vn=(f"Ngày {JUNE_DATE} từ {J_START} đến {J_END}; "
                        f"ngày {SAMPLE_DATE} từ {START} đến {END}"),
        sample_time_en=f"From {JUNE_DATE} and {SAMPLE_DATE}",
        inspection_dates=[JUNE_DATE, SAMPLE_DATE],
        sub_no=CUR_N, is_multi_cert=True,
        dossier_type="cleanroom",
        # theo BBLM 18/06 ("Đo kiểm tra các yếu tố môi trường của LAF Cân mẫu") — bao cả HEPA + bụi của BBLM 11/11;
        # câu gộp 2 BBLM dài 3 dòng làm bìa tràn sang trang 2
        scope_vn="Đo kiểm tra các yếu tố môi trường của buồng LAF",
        scope_en="Measurement of environmental parameters of LAF booths",
        location_desc_vn=SITE_VN, location_desc_en=SITE_EN,
        procedure_vn="QUATEST 3 3030:2025, N2.TC.LAF.01, N2.TC.PS.01",
        procedure_en="QUATEST 3 3030:2025, N2.TC.LAF.01, N2.TC.PS.01",
        # Căn cứ: bụi theo ISO 14644-1 + mức WHO TRS 961 (GMP Grade A), downflow WHO TRS 961, HEPA ISO 14644-3.
        # Độ sáng / ồn tạm không đánh giá (11/11/2026) -> không ghi EN 12469:2000.
        standards_vn=STANDARDS,
        standards_en=STANDARDS,
        client=CLIENT,
        personnel=PersonnelInfo(executor="Phát"),     # BBLM: Nguyễn Tiến Phát -> GĐV ký Nguyễn Xuân Nam
        equipments=[eq("kanomax_6501"), eq("testo_440"), eq("rion_nl53"), eq("aerotrak_9350"), eq("ati_6d_2i")],
        locations=locations,
        photos=photos, layouts=[],
        # 3 ảnh / buồng -> 1 hàng 3 cột 5 cm (người dùng chốt 20/11/2026: "nếu 3 hình thì vẫn để 3 cột")
        results_data={"bui_3kenh": bui, "hepa": hepa, "photo_cols": 3, "photo_width_cm": 5.0},
        input_dir=INPUT_DIR, output_dir="", network_dir="",
    )


# ---------------------------------------------------------------------------
# Hậu xử lý DOCX
# ---------------------------------------------------------------------------
def _drop_layout_notes(doc):
    """Không có sơ đồ vị trí -> bỏ dòng ghi chú '(2) Vị trí đo xem sơ đồ minh họa tại trang …' và ký hiệu (2)
    ở tiêu đề cột 'Vị trí đo' (vị trí đã ghi rõ: 1, 2 theo phiếu ghi chép; lọc trái / phải)."""
    for t in doc.tables:
        for tr in list(t._tbl.findall(qn("w:tr"))):
            txt = "".join(x.text or "" for x in tr.iter(qn("w:t")))
            if "xem sơ đồ minh họa" in txt or "see the illustrated diagram" in txt.lower():
                t._tbl.remove(tr)
        for tr in t._tbl.findall(qn("w:tr"))[:2]:
            for p in tr.iter(qn("w:p")):
                txt = "".join(x.text or "" for x in p.iter(qn("w:t")))
                if txt.strip().startswith("Vị trí") and re.search(r"\(\d\)", txt):
                    m = re.search(r"\s*\(\d\)", txt)
                    b127._replace_in_paragraph(p, m.group(0), "")


def _drop_table_a1_note(doc):
    """Bỏ câu 'Số lượng vị trí đo nồng độ hạt bụi được xác định theo bảng A1, ISO 14644-1:2015' — số vị trí theo
    BBLM (01 vị trí / buồng LAF), không theo bảng A1."""
    for t in doc.tables:
        for tr in list(t._tbl.findall(qn("w:tr"))):
            txt = "".join(x.text or "" for x in tr.iter(qn("w:t")))
            if "bảng A1" not in txt:
                continue
            if "Thông số kỹ thuật" not in txt:
                t._tbl.remove(tr)
                continue
            for p in list(tr.iter(qn("w:p"))):
                if "bảng A1" in "".join(x.text or "" for x in p.iter(qn("w:t"))):
                    p.getparent().remove(p)


# Mục thêm (dựng trên khung bảng HEPA): (tiêu đề, cột vị trí, cột giá trị, nhãn dòng mức, mức, tên chỉ tiêu, căn cứ)
EXTRA = [
    ("Kết quả đo tốc độ trung bình dòng khí cấp qua lọc HEPA", "Vị trí đo", "Tốc độ trung bình, m/s",
     f"Mức quy định đối với khu vực sạch GMP Grade A theo {BASIS_VN}", "0,36 – 0,54",
     "tốc độ trung bình dòng khí cấp qua lọc HEPA", f"đối với khu vực sạch GMP Grade A theo {BASIS_VN}", 0),
    ("Kết quả đo cường độ sáng trung bình bên trong khu vực làm việc", "Vị trí đo", "Cường độ sáng trung bình, lux",
     "Mức quy định nêu tại EN 12469:2000", "≥ 750", "cường độ sáng trung bình", "tương ứng nêu tại EN 12469:2000", 1),
    ("Kết quả đo độ ồn", "Vị trí đo", "Độ ồn, dBA",
     "Mức quy định nêu tại EN 12469:2000", "≤ 65", "độ ồn", "tương ứng nêu tại EN 12469:2000", 2),
]


DIAGRAM_DIR = os.path.join(INPUT_DIR, "TBKQ", "_diagrams")
# Hình học sơ đồ (pixel ảnh nền, nhìn từ trên xuống): khoang làm việc, 2 tấm lọc HEPA, vùng nhãn lọc phía dưới
DF_W, DF_H = 1200, 960
DF_BOX = (90, 110, 1110, 770)                # x0, y0, x1, y1 khoang làm việc (vách trong ở trên, cửa ở dưới)
DF_PAD, DF_LABEL_H = 30, 64                  # lề trong khoang; chiều cao vùng nhãn "Lọc HEPA trái / phải"


def _df_filters():
    """[(x_trái, y_trên, x_phải, y_dưới, nhãn)] của 2 tấm lọc HEPA."""
    x0, y0, x1, y1 = DF_BOX
    xm = (x0 + x1) / 2
    top, bot = y0 + DF_PAD, y1 - DF_PAD - DF_LABEL_H
    return [(x0 + DF_PAD, top, xm - DF_PAD / 2, bot, "Lọc HEPA trái"),
            (xm + DF_PAD / 2, top, x1 - DF_PAD, bot, "Lọc HEPA phải")]


def df_points():
    """16 vị trí đo: mỗi tấm lọc lưới 2 cột × 4 hàng **cách đều mép tấm lọc** (tâm ô lưới: lề = nửa bước), người dùng
    yêu cầu 17/11/2026. Đánh số theo hàng từ vách trong, trái -> phải qua cả 2 tấm lọc (thứ tự ghi trên phiếu)."""
    (la, lt, lb, lbot, _), (ra, _, rb, _, _) = _df_filters()
    xs = [la + (lb - la) * (c + 0.5) / 2 for c in range(2)] + [ra + (rb - ra) * (c + 0.5) / 2 for c in range(2)]
    ys = [lt + (lbot - lt) * (r + 0.5) / 4 for r in range(4)]
    return [(str(r * 4 + c + 1), xs[c], ys[r]) for r in range(4) for c in range(4)]


def _df_plan():
    """Ảnh nền sơ đồ (lớp khóa trong draw.io): vách, cửa, 2 tấm lọc HEPA, nhãn — không có vị trí đo.
    Chữ Times New Roman (AGENTS 8C)."""
    from PIL import Image, ImageDraw, ImageFont
    fdir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    f_bold = ImageFont.truetype(os.path.join(fdir, "timesbd.ttf"), 52)
    f_it = ImageFont.truetype(os.path.join(fdir, "timesi.ttf"), 44)
    im = Image.new("RGB", (DF_W, DF_H), "white")
    d = ImageDraw.Draw(im)
    x0, y0, x1, y1 = DF_BOX

    def ctext(x, y, s, font, fill="black"):
        b = d.textbbox((0, 0), s, font=font)
        d.text((x - (b[2] - b[0]) / 2, y - (b[3] - b[1]) / 2 - b[1]), s, font=font, fill=fill)

    # vách trong (nét đậm), 2 vách bên, cửa (nét đứt)
    d.rectangle((x0 - 14, y0 - 14, x1 + 14, y0), fill=(150, 150, 150))
    d.rectangle((x0 - 14, y0, x0, y1), fill=(150, 150, 150))
    d.rectangle((x1, y0, x1 + 14, y1), fill=(150, 150, 150))
    for x in range(x0, x1, 36):
        d.line((x, y1, min(x + 20, x1), y1), fill="black", width=4)
    ctext((x0 + x1) / 2, y0 - 55, "Vách trong", f_bold)
    ctext((x0 + x1) / 2, y1 + 50, "Cửa", f_bold)
    # 2 tấm lọc HEPA (vùng tô nhạt), nhãn ngay dưới mỗi tấm (ngoài tấm lọc, không đè vị trí đo)
    for a, top, b, bot, lbl in _df_filters():
        d.rectangle((a, top, b, bot), fill=(232, 240, 250), outline=(90, 120, 170), width=3)
        ctext((a + b) / 2, bot + DF_LABEL_H / 2 + 4, lbl, f_it, fill=(40, 70, 130))
    ctext(DF_W / 2, DF_H - 50, "Mặt phẳng đo cách lưới phân phối khí 30 cm", f_bold)
    return im


_DF_REBUILD = False          # --tao-lai-so-do: dựng lại .drawio từ tọa độ code (bỏ chỉnh tay), chỉ lần đầu


def draw_downflow_diagram() -> str:
    """Sơ đồ các vị trí đo tốc độ dòng khí cấp qua lọc HEPA dạng draw.io (người dùng yêu cầu 17/11/2026):
    TBKQ/_diagrams/so_do_vi_tri.drawio — lớp nền khóa (_df_plan), lớp "Vị trí đo" 16 vòng tròn kéo thả được.
    File đã có thì giữ chỉnh tay của người dùng (sửa tab _vn, Ctrl+S, chạy lại script) — Core/Tools/drawio_layout.py.
    Trả PNG tab _vn (chỉ có TBKQ tiếng Việt)."""
    global _DF_REBUILD
    from Core.Tools.drawio_layout import ensure_layout
    markers = [{"preset": "point", "text": t, "x": x, "y": y, "color": "#C81E1E", "stroke": "#C81E1E",
                "fill": "#FFFFFF", "stroke_width": 5, "w": 76, "h": 76, "font_size": 44}
               for t, x, y in df_points()]
    pngs = ensure_layout(DIAGRAM_DIR, _df_plan, markers, overwrite=_DF_REBUILD, frame=False)
    _DF_REBUILD = False
    return pngs["vn"]


def _set_cell_text(tc, text: str, bold: bool = False):
    """Ghi chữ vào ô (giữ định dạng run đầu, bỏ các run / đoạn thừa)."""
    ps = tc.findall(qn("w:p"))
    for p in ps[1:]:
        tc.remove(p)
    p = ps[0]
    runs = p.findall(qn("w:r"))
    for r in runs[1:]:
        p.remove(r)
    if not runs:
        from copy import deepcopy
        r = p.makeelement(qn("w:r"), {})
        p.append(r)
    else:
        r = runs[0]
    for c in list(r):
        if c.tag != qn("w:rPr"):
            r.remove(c)
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = r.makeelement(qn("w:rPr"), {})
        r.insert(0, rpr)
    for b in rpr.findall(qn("w:b")) + rpr.findall(qn("w:bCs")):
        rpr.remove(b)
    if bold:
        rpr.append(rpr.makeelement(qn("w:b"), {}))
    t = r.makeelement(qn("w:t"), {})
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)


def _rebuild_downflow_table(tbl):
    """Bảng 9.3 = bảng số liệu dẫn chứng theo Sơ đồ 1: 4 cặp cột Vị trí đo | Tốc độ, mỗi hàng = 1 hàng vị trí trên
    sơ đồ, rồi tốc độ trung bình và mức quy định."""
    from copy import deepcopy
    rows = tbl.findall(qn("w:tr"))
    hdr, data = rows[0], rows[1]
    for tr in rows:
        tbl.remove(tr)
    grid = tbl.find(qn("w:tblGrid"))
    gcols = grid.findall(qn("w:gridCol"))
    total = sum(int(c.get(qn("w:w"))) for c in gcols)
    for c in gcols:
        grid.remove(c)
    fr = (0.11, 0.14) * 4
    widths = [int(total * f) for f in fr]
    widths[-1] += total - sum(widths)
    for w in widths:
        grid.append(grid.makeelement(qn("w:gridCol"), {qn("w:w"): str(w)}))
    hdr_tc = hdr.findall(qn("w:tc"))[-1]
    data_tc = data.findall(qn("w:tc"))[-1]

    def make_row(tpl_tr, tpl_tc, cells, header=False):
        """cells: [(chữ, số cột gộp, in đậm, vMerge None|'restart'|'continue')]."""
        tr = deepcopy(tpl_tr)
        for tc in tr.findall(qn("w:tc")):
            tr.remove(tc)
        c0 = 0
        for text, span, bold, vm in cells:
            tc = deepcopy(tpl_tc)
            tcpr = tc.find(qn("w:tcPr"))
            for el in tcpr.findall(qn("w:gridSpan")) + tcpr.findall(qn("w:vMerge")):
                tcpr.remove(el)
            if span > 1:
                tcpr.insert(1, tcpr.makeelement(qn("w:gridSpan"), {qn("w:val"): str(span)}))
            if vm:
                tcpr.append(tcpr.makeelement(qn("w:vMerge"), {qn("w:val"): vm} if vm == "restart" else {}))
            tcw = tcpr.find(qn("w:tcW"))
            if tcw is not None:
                tcw.set(qn("w:type"), "pct")
                tcw.set(qn("w:w"), str(round(5000 * sum(widths[c0:c0 + span]) / total)))
            _set_cell_text(tc, text, bold=bold)
            tr.append(tc)
            c0 += span
        if header:
            trpr = tr.find(qn("w:trPr"))
            if trpr is None:
                trpr = tr.makeelement(qn("w:trPr"), {})
                tr.insert(0, trpr)
            if trpr.find(qn("w:tblHeader")) is None:
                trpr.append(trpr.makeelement(qn("w:tblHeader"), {}))
        return tr

    n = AREA[0]
    out =[make_row(hdr, hdr_tc, [x for _ in range(4) for x in (("Vị trí đo", 1, False, None),
                                                               ("Tốc độ, m/s", 1, False, None))], header=True)]
    for r in range(4):
        cells = []
        for c in range(4):
            j = r * 4 + c
            cells += [(str(j + 1), 1, False, None), (_vn(JUNE[n]["df"][j], 2), 1, False, None)]
        out.append(make_row(data, data_tc, cells))
    ok = DF_LO <= df_avg(n) <= DF_HI
    out.append(make_row(data, data_tc, [("Tốc độ trung bình", 4, False, None),
                                        (_vn(df_avg(n), 2), 4, not ok, None)]))   # không đạt in đậm (mục 3)
    out.append(make_row(data, data_tc, [("Mức quy định ⁽²⁾", 4, False, None), ("0,36 – 0,54", 4, False, None)]))
    for tr in out:
        tbl.append(tr)
    # bảng nằm trọn một trang cùng sơ đồ (cantSplit + keepNext mọi hàng trừ hàng cuối — AGENTS mục 3, 15)
    for i, tr in enumerate(out):
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = tr.makeelement(qn("w:trPr"), {})
            tr.insert(0, trpr)
        if trpr.find(qn("w:cantSplit")) is None:
            trpr.append(trpr.makeelement(qn("w:cantSplit"), {}))
        for p in tr.iter(qn("w:p")):
            ppr = p.find(qn("w:pPr"))
            if ppr is None:
                ppr = p.makeelement(qn("w:pPr"), {})
                p.insert(0, ppr)
            for x in ppr.findall(qn("w:keepNext")):
                ppr.remove(x)
            if i < len(out) - 1:
                ppr.insert(0, ppr.makeelement(qn("w:keepNext"), {}))


def _insert_downflow_diagram(doc, heading_el):
    """Ngay sau tiêu đề 9.3: hình + 'Sơ đồ 1: …' (13pt nghiêng), đi liền bảng. Không đoạn dẫn (trùng chú thích) —
    để mục 9.3 + Đánh giá nằm trọn 1 trang."""
    from docx.shared import Cm, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    path = draw_downflow_diagram()
    pic = doc.add_paragraph()
    pic.add_run().add_picture(path, width=Cm(8.0))
    cap = doc.add_paragraph()
    rc = cap.add_run("Sơ đồ 1: Các vị trí đo tốc độ dòng khí cấp qua lọc HEPA")
    rc.italic = True
    for p, align in ((pic, WD_ALIGN_PARAGRAPH.CENTER), (cap, WD_ALIGN_PARAGRAPH.CENTER)):
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before, pf.space_after, pf.keep_with_next = Pt(0), Pt(3), True
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
            rpr = run._r.get_or_add_rPr()
            rf = rpr.find(qn("w:rFonts"))
            rf.set(qn("w:eastAsia"), "Times New Roman")
            rf.set(qn("w:cs"), "Times New Roman")
    hp = heading_el.find(qn("w:pPr"))
    if hp is not None and hp.find(qn("w:keepNext")) is None:
        hp.insert(0, hp.makeelement(qn("w:keepNext"), {}))     # tiêu đề 9.3 đi liền sơ đồ
    heading_el.addnext(pic._p)
    pic._p.addnext(cap._p)


def _text(el) -> str:
    return "".join(x.text or "" for x in el.iter(qn("w:t")))


def _fix_extra_sections(doc):
    from copy import deepcopy
    from docx.text.paragraph import Paragraph
    from Core.Styling import set_assessment_paragraph
    body = doc.element.body
    name = AREA[0]
    tbls = [t for t in body.iterchildren(qn("w:tbl")) if "Độ rò rỉ lọc HEPA, %" in _text(t.findall(qn("w:tr"))[0])]
    if len(tbls) != 4:
        print(f"[Warning] build_067460: thấy {len(tbls)} bảng khung HEPA (cần 4)")
        return
    head0 = tbls[0].getprevious()
    while head0 is not None and "Kết quả đo độ rò rỉ lọc HEPA" not in _text(head0):
        head0 = head0.getprevious()
    head_tpl = deepcopy(head0)
    # HEPA (bảng thứ 4) cần tiêu đề riêng; 3 bảng đầu là các mục thêm
    new_h = deepcopy(head_tpl)
    tbls[3].addprevious(new_h)
    for i, (title, c2, c3, lim_lbl, lim, what, basis, k) in enumerate(EXTRA):
        t = tbls[i]
        if i == 0:
            h = head0
        else:
            h = deepcopy(head_tpl)
            t.addprevious(h)
        b127._replace_in_paragraph(h, "Kết quả đo độ rò rỉ lọc HEPA", title)
        ok = june_checks(name)[k][2]          # None = không đánh giá
        note = t.getnext()
        while note is not None and note.tag != qn("w:tbl"):
            note = note.getnext()
        if i == 0:
            # bảng số liệu dẫn chứng 16 vị trí + sơ đồ các vị trí đo (người dùng yêu cầu 11/11/2026)
            _rebuild_downflow_table(t)
            _insert_downflow_diagram(doc, h)
            nr = deepcopy(note.findall(qn("w:tr"))[-1])
            ntcs = nr.findall(qn("w:tc"))
            _set_cell_text(ntcs[0], "⁽²⁾")
            _set_cell_text(ntcs[-1], f"Mức quy định đối với khu vực sạch GMP Grade A theo {BASIS_VN}.")
            note.append(nr)
        rows = t.findall(qn("w:tr"))
        if i > 0 and ok is None:
            # không đánh giá: bỏ dòng mức + đoạn "Đánh giá:" và câu kết luận (AGENTS mục 3)
            t.remove(rows[-1])
            el = note.getnext()
            if el is not None and _text(el).strip() == "Đánh giá:":
                sent = el.getnext()
                el.getparent().remove(el)
                sent.getparent().remove(sent)
            for p in rows[0].iter(qn("w:p")):
                b127._replace_in_paragraph(p, "Độ rò rỉ lọc HEPA, %", c3)
                if _text(p).strip() == "Vị trí đo":
                    b127._replace_in_paragraph(p, "Vị trí đo", c2)
            continue
        if i == 0:
            rows = []          # bảng đã dựng lại
        for p in (rows[0].iter(qn("w:p")) if rows else ()):
            b127._replace_in_paragraph(p, "Độ rò rỉ lọc HEPA, %", c3)
            if _text(p).strip() == "Vị trí đo":
                b127._replace_in_paragraph(p, "Vị trí đo", c2)
        for p in (rows[-1].iter(qn("w:p")) if rows else ()):
            b127._replace_in_paragraph(p, "Mức quy định nêu tại ISO 14644-3:2019", lim_lbl)
            b127._replace_in_paragraph(p, f"≤ {901 + i}", lim)
        # giá trị không đạt -> in đậm (AGENTS mục 3)
        for tr in (rows[1:-1] if ok is False else ()):
            for rpr_r in tr.findall(qn("w:tc"))[-1].iter(qn("w:r")):
                rpr = rpr_r.find(qn("w:rPr"))
                if rpr is None:
                    rpr = rpr_r.makeelement(qn("w:rPr"), {})
                    rpr_r.insert(0, rpr)
                if rpr.find(qn("w:b")) is None:
                    rpr.append(rpr.makeelement(qn("w:b"), {}))
        # câu đánh giá
        el = t.getnext()
        while el is not None and _text(el).strip() != "Đánh giá:":
            el = el.getnext()
        sent = el.getnext()
        body_txt = f"tại “{name}” {'phù hợp' if ok else 'không phù hợp'}"
        set_assessment_paragraph(Paragraph(sent, doc._body), f"Kết quả đo {what} {body_txt} với mức quy định {basis}.",
                                 font_size=13, keep_next=False)


def finish_tbkq(src: str) -> str:
    doc = Document(src)
    b127._fix_cover_lists(doc)
    b127._fix_misc(doc, False)
    n = b127._fix_dust_tables(doc, False)
    if n != 1:
        print(f"[Warning] build_067460: sửa {n} bảng bụi (cần 1)")
    _drop_layout_notes(doc)
    _drop_table_a1_note(doc)
    _fix_extra_sections(doc)
    # 5 thiết bị + 3 căn cứ: khoảng cách đoạn trong ô Table 0 (đã thu 3 -> 1,5pt khi dựng) về 0 để chữ ký nằm trang bìa
    # (như shrink_cover_break compact=2, AGENTS mục 14)
    for tr in doc.tables[0]._tbl.findall(qn("w:tr"))[:-2]:
        for sp in tr.iter(qn("w:spacing")):
            for k in ("w:before", "w:after"):
                if sp.get(qn(k)) in ("30", "60"):
                    sp.set(qn(k), "0")
    b127._keep_headings_with_tables(doc)
    b127._keep_result_blocks(doc)
    style_note_tables(doc)      # bảng Ghi chú 11pt in nghiêng (AGENTS mục 3)
    b127._highlight_placeholders(doc)
    dst = os.path.join(os.path.dirname(src), f"TBKQ - viet - {CUR_N} - phong sach - {CODE}.docx")
    doc.save(dst)
    if os.path.abspath(src) != os.path.abspath(dst):
        os.remove(src)   # tên handler đặt ("TBKQ - vn - 1 - ...") khác quy ước AGENTS mục 1 (E41)
    return dst


# ---------------------------------------------------------------------------
# BCKQ
# ---------------------------------------------------------------------------
def _set_so(n: int):
    """Số phiếu tính toán theo chứng thư: <mã>/<n>/<chỉ tiêu> (AGENTS mục 7)."""
    global SO_BUI, SO_SCAN, SO_VT, SO_LUX, SO_ON
    SO_BUI, SO_SCAN, SO_VT, SO_LUX, SO_ON = (f"{CODE}/{n}/{k}" for k in
                                             ("Bụi", "Scan lọc", "Vận tốc", "Độ sáng", "Độ ồn"))


_set_so(1)


def build_bckq(d: Dossier) -> str:
    b = BCKQBuilder(d)
    ctx = b.build_context()
    ctx["tham_chieu"] = [
        f"- Kết quả đo nồng độ hạt bụi: xem tại phiếu Tính toán kết quả nồng độ hạt bụi số {SO_BUI} ngày {SAMPLE_DATE}.",
        f"- Kết quả quét rò rỉ lọc HEPA: xem tại phiếu Tính toán kết quả rò rỉ lọc HEPA số {SO_SCAN} ngày {SAMPLE_DATE}.",
        f"- Kết quả đo tốc độ dòng khí cấp qua lọc HEPA: xem tại phiếu Tính toán vận tốc gió qua lọc HEPA số "
        f"{SO_VT} ngày {JUNE_DATE}.",
        f"- Kết quả đo cường độ sáng: xem tại phiếu Tính toán cường độ sáng trung bình số {SO_LUX} ngày {JUNE_DATE}.",
        f"- Kết quả đo độ ồn: xem tại phiếu Tính toán kết quả độ ồn số {SO_ON} ngày {JUNE_DATE}.",
    ]
    vn = AREA[0]
    checks = june_checks(vn) + [("Nồng độ hạt bụi", "", dust_ok()), ("Độ rò rỉ lọc HEPA", "", hepa_ok())]
    dat = [n for n, _, ok in checks if ok]
    khong = [f"{n} ({v})" if v else n for n, v, ok in checks if ok is False]   # None = không đánh giá
    ctx["danh_gia"] = [{
        "khu_vuc": f"“{vn}”",
        "dat": ";\n".join(dat) or "Không",
        "khong_dat": ";\n".join(khong) or "Không",
        "can_cu": (f"{BASIS_VN} (khu vực sạch {b127.GMP(GRADE)}: tốc độ dòng khí 0,36 – 0,54 m/s; trạng thái "
                   f"nghỉ d ≥ 0,5 µm ≤ 3 520 hạt/m³, d ≥ 5 µm ≤ 20 hạt/m³); ISO 14644-3:2019 (rò rỉ lọc HEPA "
                   f"{HEPA_LIM} %)" + ("; EN 12469:2000 (cường độ sáng ≥ 750 lux; độ ồn ≤ 65 dBA)"
                                       if DANH_GIA_SANG_ON else "")),
    }]
    out = b.build(ctx=ctx, out_name=f"BCKQ - {CUR_N} - phong sach - {CODE}.docx")
    doc = Document(out)
    b127._fix_misc(doc, False)
    _drop_table_a1_note(doc)
    b127._keep_headings_with_tables(doc)
    for p in doc.paragraphs:
        if p.text.strip().startswith(("Hình ảnh", "Sơ đồ", "Báo cáo của GĐV")):
            p.paragraph_format.keep_with_next = True
    b127._highlight_placeholders(doc)
    doc.save(out)
    return out


# ---------------------------------------------------------------------------
# Excel phiếu tính toán (form phòng sạch chung, AGENTS mục 8C) — "Bụi 0,3" + "Scan lọc"
# ---------------------------------------------------------------------------
def build_excel() -> str:
    book = CleanroomCalcBook()
    from Core.Tools.cleanroom_calc import _limit_cond, _eval_formula, LEFT
    from Core.Tools.noise_corr import noise_correction_formula
    from openpyxl.utils import get_column_letter as L
    room = AREA[0]
    groups = [{"room": room, "cap": b127.GMP(GRADE), "basis": BASIS_VN,
               "limits": {"Ch2": LIM05, "Ch6": LIM5}, "zero": ZERO,
               "points": [AREA[4]], "labels": [AREA[3]]}]
    ws = book.dust_aerotrak(groups, so=SO_BUI, date=SAMPLE_DATE, labels=AEROTRAK_LABELS)
    for r in (2, 3):                         # ô tên phòng gộp 2 hàng (zero + vị trí 1) -> đủ cao cho 3 dòng chữ
        ws.row_dimensions[r].height = 24
    ws.row_dimensions[5].height = 36
    for c in range(1, ws.max_column + 1):    # phiếu in nhiệt không có Temp / Humidity / Velocity
        if str(ws.cell(1, c).value or "") in ("Temp", "Humidity", "Velocity"):
            ws.column_dimensions[L(c)].hidden = True
    book.scan([{"name": f"{room} - {lbl}", "media": _pct(media), "gasket": _pct(gasket), "lim": HEPA_LIM}
               for lbl, gasket, media in AREA[5]], so=SO_SCAN, date=SAMPLE_DATE)

    # --- Đợt 18/06: vận tốc gió (downflow), độ sáng — số đọc + công thức trung bình, mức, đánh giá ---
    for key, so, title, n, nd, lim, unit in (
            ("df", SO_VT, "TÍNH TOÁN VẬN TỐC GIÓ QUA LỌC HEPA", 16, 2, DF_LIM, "m/s"),
            ("lux", SO_LUX, "TÍNH TOÁN CƯỜNG ĐỘ SÁNG TRUNG BÌNH", 9, 1, LUX_LIM, "lux")):
        sheet = "Vận tốc" if key == "df" else "Độ sáng"
        w = book.extra_sheet(sheet, so=so, date=JUNE_DATE, widths=[5, 24] + [7.5] * n + [11, 11, 10], title=title,
                             title_rows="1:2")
        c_av, c_lim, c_ev = 3 + n, 4 + n, 5 + n
        book.merge(w, 1, 1, 2, 1, "STT", bold=True)
        book.merge(w, 1, 2, 2, 2, "Khu vực đo", bold=True)
        book.merge(w, 1, 3, 1, 2 + n, f"Giá trị đọc tại các vị trí đo, {unit}", bold=True)
        for i in range(n):
            book.put(w, 2, 3 + i, i + 1, bold=True)
        for c, lbl in ((c_av, f"Trung bình, {unit}"), (c_lim, "Mức quy định"), (c_ev, "Đánh giá")):
            book.merge(w, 1, c, 2, c, lbl, bold=True)
        w.row_dimensions[1].height = 30
        fmt = "0.00" if nd == 2 else "0.0"
        r = 3
        book.put(w, r, 1, 1)
        book.put(w, r, 2, room, align=LEFT)
        for i, v in enumerate(JUNE[room][key]):
            book.put(w, r, 3 + i, v, fmt=fmt)
        book.put(w, r, c_av, f"=ROUND(AVERAGE({L(3)}{r}:{L(2 + n)}{r}),{nd})", fmt=fmt, bold=True)
        if key == "lux" and not DANH_GIA_SANG_ON:
            lim = None                    # tạm thời không đánh giá -> mức / đánh giá "-"
        book.put(w, r, c_lim, lim.replace(" - ", " – ") if lim else "-")
        book.put(w, r, c_ev, _eval_formula([_limit_cond(f"{L(c_av)}{r}", lim)]))
        w.row_dimensions[r].height = 30

    # --- Độ ồn: ồn nền, ồn vận hành, hiệu chỉnh ồn nền (công thức theo noise_corr), kết quả ---
    w = book.extra_sheet("Độ ồn", so=SO_ON, date=JUNE_DATE, widths=[5, 34, 18, 12, 13, 12, 12, 11, 10],
                         title="TÍNH TOÁN KẾT QUẢ ĐỘ ỒN")
    for c, lbl in enumerate(("STT", "Khu vực đo", "Vị trí đo", "Độ ồn nền, dBA", "Độ ồn khi vận hành, dBA",
                             "Hiệu chỉnh ồn nền, dBA", "Kết quả, dBA", "Mức quy định", "Đánh giá"), 1):
        book.put(w, 1, c, lbl, bold=True)
    w.row_dimensions[1].height = 32
    r, j = 2, JUNE[room]
    book.put(w, r, 1, 1)
    book.put(w, r, 2, room, align=LEFT)
    book.put(w, r, 3, "Bên trong buồng")
    book.put(w, r, 4, j["bg"] if j["bg"] is not None else "-", fmt="0.0")
    book.put(w, r, 5, j["op"], fmt="0.0")
    book.put(w, r, 6, noise_correction_formula(f"D{r}", f"E{r}") if j["bg"] is not None else "-", fmt="0")
    book.put(w, r, 7, f"=E{r}-F{r}" if j["bg"] is not None else f"=E{r}", fmt="0.0", bold=True)
    nlim = NOISE_LIM if DANH_GIA_SANG_ON else None
    book.put(w, r, 8, nlim or "-")
    book.put(w, r, 9, _eval_formula([_limit_cond(f"G{r}", nlim)]))
    out = os.path.join(INPUT_DIR, "ThucHien", f"{CODE} - {CUR_N} - TinhToanPhongSach.xlsx")
    book.save(out)
    print(f"[build_067460] Excel -> {out}")
    return out


def split_edoc() -> list:
    """BBLM: 11/11 (A tr.11–12) + 18/06 (A tr.17–18); Ghi chép: 11/11 (A tr.13–15), thông số tủ 18/06 (A tr.19–21,
    23–25), phiếu in AeroTrak + ATI (B tr.1–3); Khác: PKSCV, 2 bản PYC, báo giá (A tr.1, 3, 5, 7–9).
    Trang trắng bỏ, trang xoay đúng hướng (finish_edoc — AGENTS mục 8A)."""
    import fitz
    from Core.Tools.scan_splitter import finish_edoc
    edoc = os.path.join(INPUT_DIR, "TBKQ", "Edoc")
    os.makedirs(edoc, exist_ok=True)
    a, b = fitz.open(SCAN_A), fitz.open(SCAN_B)
    plan = {
        f"{CODE6} - BBLM.pdf": [(a, [11, 12, 17, 18])],
        f"{CODE6} - Ghi Chep.pdf": [(a, [13, 14, 15, 19, 20, 21, 23, 24, 25]), (b, [1, 2, 3])],
        f"{CODE6} - Khac.pdf": [(a, [1, 3, 5, 7, 8, 9])],
    }
    outs = []
    for name, parts in plan.items():
        dst = fitz.open()
        for src, pages in parts:
            for pg in pages:
                dst.insert_pdf(src, from_page=pg - 1, to_page=pg - 1)
        finish_edoc(dst, INPUT_DIR, name)
        p = os.path.join(edoc, name)
        dst.save(p)
        outs.append(p)
        print(f"[build_067460] Edoc -> {p} ({len(dst)} trang)")
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (Word/Excel COM) + chèn giấy tiêu đề UpChungThu")
    ap.add_argument("--only", default="", help="Chỉ dựng chứng thư, vd 1 hoặc 1,2")
    ap.add_argument("--tao-lai-so-do", action="store_true",
                    help="Dựng lại TBKQ/_diagrams/so_do_vi_tri.drawio từ tọa độ code (bỏ chỉnh tay)")
    args = ap.parse_args()
    global AREA, CUR_N, _DF_REBUILD
    _DF_REBUILD = args.tao_lai_so_do or not os.path.exists(os.path.join(DIAGRAM_DIR, "so_do_vi_tri.drawio"))
    only = {int(x) for x in args.only.split(",") if x.strip()} or set(range(1, len(LAFS) + 1))
    edoc_done = False
    results = []
    for n, area in enumerate(LAFS, 1):       # mỗi buồng LAF 1 chứng thư /n
        if n not in only:
            continue
        AREA, CUR_N = area, n
        _set_so(n)
        d = build_dossier()
        d.ensure_output_directories()
        if not edoc_done:
            split_edoc()
            edoc_done = True
        h = CleanroomHandler(d)
        vn = finish_tbkq(h.build_tbkq(is_en=False))
        bckq = build_bckq(d)
        xlsx = build_excel()
        docs = [bckq, vn]
        if args.pdf:
            from Core.Exporters.pdf_exporter import PDFExporter
            ex = PDFExporter(d)
            ex.export_files(docs)
            ex.export_excel_files([xlsx])
        results += [(p, d.code) for p in docs]
    print("\n===== QA =====")
    for p, code in results:
        print(f"  {check_single_docx(p, code=code).get('status')}\t{os.path.basename(p)}")


if __name__ == "__main__":
    main()
