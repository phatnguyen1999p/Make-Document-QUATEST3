# -*- coding: utf-8 -*-
"""
Dossier & Inspection Data Models for QUATEST 3 (Room N2).
Provides structured representations for Dossiers, Clients, Equipment, Locations, and Results.
"""

import os
import re
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any


# Tiền tố cũ của mục 3 "Nội dung yêu cầu giám định" (form tủ) -> thay bằng "Giám định ..." / "Inspection of ..."
_OLD_SCOPE_VN = re.compile(r"^\s*(?:Đo\s*/\s*kiểm\s+tra|Đo\s*/\s*Kiểm\s+tra|Đo\s+kiểm\s+tra|Kiểm\s+tra)\s+", re.IGNORECASE)
_OLD_SCOPE_EN = re.compile(r"^\s*(?:Measuring\s*/\s*validating|Measurement\s*/\s*validation|Inspection\s+of)\s+", re.IGNORECASE)


def inspection_scope(text: str, is_en: bool = False) -> str:
    """Nội dung yêu cầu giám định của form tủ (BSC, tủ hút, tủ cấy/clean bench) — người dùng chốt 01/11/2026:
    VN 'Giám định tủ hút khí độc' (không 'Đo/ kiểm tra "Tủ hút khí độc"'), EN 'Inspection of Fume hood'.
    Bỏ ngoặc kép quanh tên tủ; VN viết thường chữ đầu tên tủ ('Tủ cấy' -> 'tủ cấy'). Câu khác giữ nguyên."""
    s = (text or "").strip()
    if not s:
        return s
    pat = _OLD_SCOPE_EN if is_en else _OLD_SCOPE_VN
    m = pat.match(s)
    if not m:
        if is_en or not s.lower().startswith("giám định"):
            return s
        name = s[len("giám định"):]
    else:
        name = s[m.end():]
    name = name.strip().strip(".").strip().replace('"', "").replace("“", "").replace("”", "").strip()
    if not name:
        return s
    if is_en:
        return f"Inspection of {name}"
    if name[:1].isupper() and not name[1:2].isupper():  # 'Tủ cấy' -> 'tủ cấy', giữ 'LAF', 'BSC'
        name = name[:1].lower() + name[1:]
    return f"Giám định {name}"


_PROC_QT = re.compile(r"QUATEST\s*3\s*3030\s*:\s*(\d{4})", re.IGNORECASE)
# Nhóm chữ tùy chọn: 'N2.TC.BSC.01' và cả mã chỉ có số 'N2.TC.09' (vi sinh, 137058 — trước 05/11/2026 bị rơi mất)
_PROC_N2 = re.compile(r"N2\s*\.\s*TC\s*\.\s*(?:([A-Z]+)\s*\.?\s*)?(\d{2})", re.IGNORECASE)


def dates_phrase(dates, is_en: bool = False) -> str:
    """Liệt kê ngày giám định (TBKQ Table 0 mục 4, người dùng chốt 07/11/2026):
    2 ngày 'd1 và d2' / 'd1 and d2'; từ 3 ngày 'd1, d2, và d3' / 'd1, d2, and d3' (giữ dấu phẩy trước 'và' như mẫu
    người dùng). Ngày trùng bỏ, giữ thứ tự."""
    ds = [str(x).strip() for x in dict.fromkeys(dates or []) if str(x).strip()]
    if not ds:
        return ""
    if len(ds) == 1:
        return ds[0]
    conj = "and" if is_en else "và"
    if len(ds) == 2:
        return f"{ds[0]} {conj} {ds[1]}"
    return ", ".join(ds[:-1]) + f", {conj} {ds[-1]}"


def bsc_class_of(d) -> int:
    """Cấp tủ ATSH: 1 khi results_data["cabinet"]["bsc_class"] là 1/'I' hoặc tủ không có dòng khí cấp
    (downflow.skip — tủ cấp I, 134804/1); còn lại 2."""
    rd = getattr(d, "results_data", None) or {}
    cls = str((rd.get("cabinet") or {}).get("bsc_class", "")).strip().upper()
    if cls in ("1", "I"):
        return 1
    if cls in ("2", "II"):
        return 2
    return 1 if (rd.get("downflow") or {}).get("skip") else 2


def inspection_procedure(text: str, bsc_class: Optional[int] = None) -> str:
    """Mục "Quy trình thực hiện" (người dùng chốt 02/11/2026, mọi hồ sơ): chỉ ghi quy trình
    'QUATEST 3 3030:2025' và các mã N2.TC.* (vd 'QUATEST 3 3030:2025, N2.TC.BSC.02'),
    ngăn cách bằng dấu phẩy; tiêu chuẩn (EN 12469, ISO…) BBLM có ghi thì bỏ (đã nằm ở Căn cứ).
    Sửa luôn 'QUATEST3 3030' -> 'QUATEST 3 3030', 'N2.TC.BSC01' -> 'N2.TC.BSC.01'.
    Tủ ATSH chỉ ghi **một** mã (người dùng chốt 06/11/2026): có cả BSC.01 và BSC.02 thì giữ
    BSC.01 cho tủ cấp I, BSC.02 cho tủ cấp II (`bsc_class`, mặc định cấp II); `bsc_class=0` giữ cả hai
    (BCKQ gộp nhiều tủ khác cấp — mỗi tủ đã một mã).
    Chuỗi không có mã nào nhận ra thì giữ nguyên."""
    s = (text or "").strip()
    items = []
    for part in re.split(r"[;,]", s):
        items += [f"QUATEST 3 3030:{m.group(1)}" for m in _PROC_QT.finditer(part)]
        items += [f"N2.TC.{m.group(1).upper() + '.' if m.group(1) else ''}{m.group(2)}" for m in _PROC_N2.finditer(part)]
    items = list(dict.fromkeys(items))
    if bsc_class != 0 and "N2.TC.BSC.01" in items and "N2.TC.BSC.02" in items:
        items.remove("N2.TC.BSC.02" if bsc_class == 1 else "N2.TC.BSC.01")
    return ", ".join(items) if items else s


@dataclass
class ClientInfo:
    """Client information in both Vietnamese and English."""
    name_vn: str = ""
    name_en: str = ""
    address_vn: str = ""
    address_en: str = ""
    contact_name: str = ""
    contact_phone: str = ""


@dataclass
class Equipment:
    """Testing & inspection equipment with VN/EN descriptions."""
    name_vn: str = ""
    name_en: str = ""
    model: str = ""
    serial: str = ""
    origin: str = ""

    def __post_init__(self):
        # Một máy một tên (AGENTS mục 14): quy về tên chuẩn trong Core/Tools/equipment_names.py
        from Core.Tools.equipment_names import normalize_equipment
        self.name_vn, self.name_en = normalize_equipment(self.name_vn, self.name_en)


@dataclass
class PersonnelInfo:
    """Inspection personnel, GDV signer, leader signer, and internal reporter."""
    executor: str = ""  # e.g. "Phát", "Toàn", "Nghĩa"
    sign_gdv_vn: str = ""
    sign_gdv_en: str = ""
    sign_leader_vn: str = "KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC\n\n\n\nLê Đăng Huyền"
    sign_leader_en: str = "FOR DIRECTOR\nVICE DIRECTOR\n\n\n\nLe Dang Huyen"
    reporter: str = ""

    def __post_init__(self):
        # Auto-infer GDV Signer per AGENTS.md Section 2
        exec_lower = self.executor.lower().strip()
        if not self.sign_gdv_vn:
            if "toàn" in exec_lower or "toan" in exec_lower:
                self.sign_gdv_vn = "GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Khánh Toàn"
                self.sign_gdv_en = "INSPECTOR\n\n\n\nNguyen Khanh Toan"
            elif "nghĩa" in exec_lower or "nghia" in exec_lower:
                self.sign_gdv_vn = "GIÁM ĐỊNH VIÊN\n\n\n\nPhạm Quốc Nghĩa"
                self.sign_gdv_en = "INSPECTOR\n\n\n\nPham Quoc Nghia"
            else:
                self.sign_gdv_vn = "GIÁM ĐỊNH VIÊN\n\n\n\nNguyễn Xuân Nam"
                self.sign_gdv_en = "INSPECTOR\n\n\n\nNguyen Xuan Nam"

        # Auto-infer BCKQ Reporter per AGENTS.md Section 7
        if not self.reporter:
            name_cap = self.executor.strip().capitalize() if self.executor else "Phát"
            if "toàn" in exec_lower or "toan" in exec_lower:
                self.reporter = f"{name_cap} + Minh"
            elif "nghĩa" in exec_lower or "nghia" in exec_lower:
                self.reporter = f"{name_cap} + Minh"
            else:
                self.reporter = f"{name_cap} + Minh + Nam"


@dataclass
class MonitoringLocation:
    """Monitoring location / area specification."""
    name_vn: str = ""
    name_en: str = ""
    area_m2: str = ""  # e.g. "0,72 m2"
    clean_class: str = ""  # e.g. "ISO 5", "ISO 7", "Class 100"
    state_vn: str = "Trạng thái nghỉ"
    state_en: str = "At rest"
    points_summary_vn: List[str] = field(default_factory=list)
    points_summary_en: List[str] = field(default_factory=list)

    def __post_init__(self):
        # Cách gọi thống nhất (AGENTS mục 6, chốt 11/11/2026): "vị trí đo" thay "điểm đo";
        # độ sáng ghi "Cường độ sáng trung bình", không ghi số lượng — Core/Tools/wording.py
        from Core.Tools.wording import normalize_points_vn, normalize_points_en
        self.points_summary_vn = [normalize_points_vn(s) if isinstance(s, str) else s for s in self.points_summary_vn]
        self.points_summary_en = [normalize_points_en(s) if isinstance(s, str) else s for s in self.points_summary_en]


@dataclass
class Dossier:
    """
    Complete inspection profile / dossier model.
    Encapsulates all metadata and measurements needed to generate TBKQ (VN/EN) and BCKQ.
    """
    code: str  # Bare dossier code, e.g. "Q326D126588"
    cert_date: str  # Issue date, e.g. "21/10/2026"
    sample_date: str = ""  # Date of inspection, e.g. "16/10/2026"
    reg_date: str = ""  # Application registration date, e.g. "20/10/2026"
    sample_time_vn: str = ""  # Detailed time, e.g. "Ngày 16/10/2026 từ 09:00 đến 10:30"
    sample_time_en: str = ""  # e.g. "From 16/10/2026"
    # Nhiều ngày đi hiện trường (vd 132809: ["29/10/2026", "04/11/2026"]) -> TBKQ Table 0 mục 4 liệt kê từng ngày
    # "Từ ngày 29/10/2026 và 04/11/2026" (người dùng chốt 07/11/2026, AGENTS mục 2); 1 ngày thì để trống.
    inspection_dates: List[str] = field(default_factory=list)
    sub_no: Optional[int] = 1  # 1 for /1, 2 for /2... or None if single bare report
    is_multi_cert: bool = True  # True if dossier has /1, /2... False if bare code only
    
    dossier_type: str = "laf"  # 'laf', 'bsc', 'fumehood', 'cleanroom', 'compressed_air'
    has_assessment: bool = True  # True: Có đánh giá (Pass/Fail), False: Không đánh giá (Mức khuyến nghị)
    both_forms: bool = False  # True: Sinh đồng thời cả 2 form (Có đánh giá & Không đánh giá)
    scope_vn: str = ""  # Measurement requirement / scope description VN
    scope_en: str = ""  # Scope description EN
    location_desc_vn: str = ""  # Detailed testing address VN
    location_desc_en: str = ""  # Detailed testing address EN

    standards_vn: str = ""
    standards_en: str = ""
    procedure_vn: str = ""
    procedure_en: str = ""
    basis_vn: str = ""
    basis_en: str = ""

    client: ClientInfo = field(default_factory=ClientInfo)
    personnel: PersonnelInfo = field(default_factory=PersonnelInfo)
    equipments: List[Equipment] = field(default_factory=list)
    locations: List[MonitoringLocation] = field(default_factory=list)

    photos: List[Dict[str, str]] = field(default_factory=list)  # [{'path': ..., 'caption_vn': ..., 'caption_en': ...}]
    layouts: List[Dict[str, str]] = field(default_factory=list)  # [{'path': ..., 'caption_vn': ..., 'caption_en': ...}]

    results_data: Dict[str, Any] = field(default_factory=dict)
    assessment: Dict[str, Any] = field(default_factory=dict)

    input_dir: str = ""
    thuchien_dir: str = ""
    tbkq_dir: str = ""
    edoc_dir: str = ""
    pdf_dir: str = ""
    output_dir: str = ""
    network_dir: str = ""

    def __post_init__(self):
        # Enforce bare code in self.code (Bẫy 5 & AGENTS Section 2)
        if "/" in self.code:
            self.code = self.code.split("/")[0]

        # Số phụ (/1, /2, ...) chỉ có ý nghĩa khi hồ sơ có nhiều báo cáo kết quả (AGENTS mục 2).
        # Hồ sơ 1 báo cáo -> sub_no = None (mã trần); hồ sơ nhiều báo cáo mà chưa có số phụ -> 1.
        self._normalize_sub_no()

        # Không ghi gọn (AGENTS mục 15): mở rộng viết tắt trong tên KH, địa chỉ, địa điểm tiếng Việt
        from Core.Tools.no_abbrev import expand_abbrev
        self.location_desc_vn = expand_abbrev(self.location_desc_vn)
        if self.client is not None:
            self.client.name_vn = expand_abbrev(self.client.name_vn)
            self.client.address_vn = expand_abbrev(self.client.address_vn)

        if not self.input_dir:
            self.input_dir = os.path.join(os.getcwd(), "Input")
        if not self.thuchien_dir:
            self.thuchien_dir = os.path.join(self.input_dir, "ThucHien")
        if not self.tbkq_dir:
            self.tbkq_dir = os.path.join(self.input_dir, "TBKQ")
        if not self.edoc_dir:
            self.edoc_dir = os.path.join(self.tbkq_dir, "Edoc")
        if not self.pdf_dir:
            self.pdf_dir = os.path.join(self.input_dir, "pdf")
        # Note: output_dir is no longer populated by default per user rule
        # Files are only kept directly in Input/<dossier>/... and network drive.

    def _normalize_sub_no(self):
        sub = self.sub_no
        if isinstance(sub, str):
            sub = sub.strip().strip("/") or None
            if sub is not None and sub.isdigit():
                sub = int(sub)
        if not self.is_multi_cert:
            sub = None
        elif sub in (None, 0):
            sub = 1
        self.sub_no = sub

    # Số chứng thư / số báo cáo luôn tính theo sub_no hiện tại (đổi sub_no là tự cập nhật)
    @property
    def cert_no_vn(self) -> str:
        """Số chứng thư VN (header TBKQ), cũng là số BCKQ / phiếu tính toán: 'Q326D...' hoặc 'Q326D.../2'."""
        return f"{self.code}/{self.sub_no}" if (self.is_multi_cert and self.sub_no) else self.code

    @property
    def cert_no_en(self) -> str:
        return f"{self.cert_no_vn}/E"

    def ensure_output_directories(self):
        """Ensures all target directories exist per QUATEST 3 structure: ThucHien, TBKQ, TBKQ/Edoc, pdf."""
        os.makedirs(self.input_dir, exist_ok=True)
        os.makedirs(self.thuchien_dir, exist_ok=True)
        os.makedirs(self.tbkq_dir, exist_ok=True)
        os.makedirs(self.edoc_dir, exist_ok=True)
        os.makedirs(self.pdf_dir, exist_ok=True)
        if self.output_dir:
            os.makedirs(self.output_dir, exist_ok=True)
        if self.network_dir:
            try:
                os.makedirs(self.network_dir, exist_ok=True)
                os.makedirs(os.path.join(self.network_dir, "ThucHien"), exist_ok=True)
                os.makedirs(os.path.join(self.network_dir, "TBKQ"), exist_ok=True)
                os.makedirs(os.path.join(self.network_dir, "TBKQ", "Edoc"), exist_ok=True)
                os.makedirs(os.path.join(self.network_dir, "pdf"), exist_ok=True)
            except Exception as _err:
                print(f"[Warning] dossier.py:172 {type(_err).__name__}: {_err}")


def assign_cert_numbers(dossiers: List["Dossier"]) -> List["Dossier"]:
    """
    Tự động đánh số phụ cho các báo cáo kết quả của CÙNG một hồ sơ (AGENTS mục 2):
    - Hồ sơ chỉ có 1 báo cáo  -> mã trần (Q326D132918), không có /1.
    - Hồ sơ có từ 2 báo cáo   -> /1, /2, /3 ... theo thứ tự trong danh sách
      (giữ số phụ đã ghi sẵn nếu hợp lệ và không trùng; báo cáo chưa có số nhận số trống nhỏ nhất).
    Nhóm theo mã hồ sơ, nên truyền chung nhiều hồ sơ khác nhau cũng được.
    """
    groups: Dict[str, List["Dossier"]] = {}
    for d in dossiers:
        groups.setdefault(d.code, []).append(d)
    for items in groups.values():
        if len(items) == 1:
            items[0].is_multi_cert = False
            items[0].sub_no = None
            continue
        used = set()
        pending = []
        for d in items:
            s = d.sub_no
            if isinstance(s, str) and s.strip().isdigit():
                s = int(s.strip())
            if isinstance(s, int) and s > 0 and s not in used:
                used.add(s)
                d.sub_no = s
            else:
                pending.append(d)
            d.is_multi_cert = True
        nxt = 1
        for d in pending:
            while nxt in used:
                nxt += 1
            d.sub_no = nxt
            used.add(nxt)
    return dossiers
