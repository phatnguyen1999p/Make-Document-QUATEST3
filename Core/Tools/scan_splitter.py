# -*- coding: utf-8 -*-
"""
Automated Scan PDF Splitter for QUATEST 3 Inspection Dossiers.

Inspects raw scanned PDFs in a dossier folder, classifies each page by document header
(using direct PDF text or Windows.Media.Ocr), and splits them into:
1. <Mã_6_số> - BBLM.pdf     -> For "Biên bản giám định hiện trường" (BM3030.02)
2. <Mã_6_số> - Ghi Chep.pdf -> For "Ghi chép tại hiện trường" (BM3030.01, thông số tủ/phòng,
                               và tất cả các trang không có header rõ ràng như phiếu in nhiệt).
3. <Mã_6_số> - Khac.pdf     -> Trang giấy tờ khác: phiếu kiểm soát công việc, PYC, báo giá, PO/đơn hàng,
                               điều kiện chung, đề cương khách hàng, hóa đơn, GCN ĐKDN…
Trang trắng (không có điểm ảnh nào đậm) bị bỏ, không vào file nào.
(Người dùng chốt 06/11/2026 — AGENTS mục 8A.)
Trang scan ngược / nằm ngang được xoay về đúng chiều đọc trước khi OCR header và khi ghi file
(`Core/Tools/pdf_orient.py`, người dùng yêu cầu 11/11/2026).

Deliverables are strictly saved into:
    <dossier_folder>/TBKQ/Edoc/
"""

import os
import re
import sys
import json
import shutil
import tempfile
import unicodedata
import subprocess
from typing import Dict, List, Optional, Any

import fitz  # PyMuPDF

try:
    from Core.Tools.pdf_orient import orient_doc, detect_rotations
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
    from Core.Tools.pdf_orient import orient_doc, detect_rotations


def normalize_text(text: str) -> str:
    """Normalize text: remove accents, lowercase, retain only alphanumeric and single spaces."""
    if not text:
        return ""
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.lower()
    text = re.sub(r"[^a-z0-9]", " ", text)
    return " ".join(text.split())


# Header giấy tờ không phải phiếu hiện trường -> '<6 số> - Khac.pdf' (đã bỏ dấu, viết thường)
KHAC_KEYWORDS = (
    "phieu kiem soat cong viec", "phieu yeu cau", "yeu cau nhan tham dinh", "yeu cau va nhan",
    "bao gia", "quotation", "purchase order", "don dat hang", "don hang", "dieu kien chung",
    "de cuong", "hoa don", "invoice", "giay chung nhan dang ky", "dang ky doanh nghiep",
    "kiem soat",
    # Giấy tiêu đề QUATEST 3 / Ủy ban TCĐLCL (PYC, báo giá in trên giấy tiêu đề; BBLM, Ghi chép đã nhận ở trên)
    "do luong chat luong", "commission for", "standards metrology", "head office",
    # cùng giấy tiêu đề khi OCR đọc méo ("do luong chivi luong", "tieu chuan do i ij o ng")
    "quality assurance", "tieu chuan do", "ky thuat tieu chuan", "ky thu t tieu chuan",
)
# Tiêu đề + mã form ở đầu trang (đã bỏ dấu). Chỉ xét HEAD_CHARS ký tự đầu của văn bản OCR header.
# Chịu được OCR đọc méo: "bion ban giam", "bien ban giam tri rong", "gi ii chep", "tai hi n tru ong".
HEAD_CHARS = 120
BBLM_HEAD_RE = re.compile(r"\bbi[eo]n ban\b|\bban giam\b|\bbm ?3030 ?02\b")
GHI_CHEP_HEAD_RE = re.compile(r"\bghi chep\b|\bgi i* ?i* ?chep\b|\btai hi ?[eê]?n ?tru|\bbm ?3030 ?01\b")
GHI_CHEP_KEYWORDS = ("ghi chep", "bm3030 01", "bm303001", "3030 01", "thong so tu", "ghi cac thong so")


def is_blank_page(page, dark_level: int = 180, max_dark_ratio: float = 5e-5) -> bool:
    """Trang trắng: bỏ lề 6 %, render xám 40 dpi, gần như không có điểm nào tối hơn `dark_level`.
    Đo 06/11/2026 trên scan 132809, 138143: trang trắng điểm tối nhất 222–249; trang có chữ (kể cả
    giấy tính tay viết nhạt, phiếu in nhiệt) đều có điểm <= 115 và tỉ lệ >= 2,5e-4."""
    try:
        r = page.rect
        clip = fitz.Rect(r.x0 + r.width * .06, r.y0 + r.height * .06,
                         r.x1 - r.width * .06, r.y1 - r.height * .06)
        pix = page.get_pixmap(clip=clip, dpi=40, colorspace=fitz.csGRAY)
        s = pix.samples
        if not s:
            return False
        dark = sum(1 for b in s if b < dark_level)
        return dark / len(s) < max_dark_ratio
    except Exception as e:
        print(f"[Warning] scan_splitter is_blank_page: {type(e).__name__}: {e}")
        return False


def finish_edoc(doc, dossier_dir: Optional[str] = None, name: str = "") -> List[int]:
    """Hoàn thiện một file Edoc trước khi `save` (script hồ sơ tự ghép bằng `insert_pdf` — AGENTS mục 8A):
    bỏ trang trắng (`is_blank_page`) rồi xoay trang về đúng chiều đọc (`orient_doc`).
    Trả danh sách trang trắng đã bỏ (1-based, theo thứ tự trước khi bỏ)."""
    blanks = [i for i in range(doc.page_count) if is_blank_page(doc[i])]
    if blanks and len(blanks) < doc.page_count:
        for i in reversed(blanks):
            doc.delete_page(i)
        print(f"[scan_splitter] {name}: bỏ {len(blanks)} trang trắng (tr. {', '.join(str(i + 1) for i in blanks)})")
    orient_doc(doc, dossier_dir, name)
    return [i + 1 for i in blanks]


def write_edoc(dossier_dir: str, code6: str, parts: Dict[str, List[tuple]]) -> Dict[str, str]:
    """Edoc tách tay theo trang (script hồ sơ đã biết từng trang scan là gì — AGENTS mục 8A):
    parts = {"BBLM": [(file, [trang 1-based] | None = mọi trang), ...], "Ghi Chep": [...], "Khac": [...]}.
    Ghi `<hồ sơ>\\TBKQ\\Edoc\\<6 số> - <tên>.pdf`; bỏ trang trắng (`is_blank_page`); tự xoay trang về đúng chiều
    đọc (`pdf_orient.orient_doc`); nhóm rỗng thì không tạo file.
    Trang lạ không thuộc hồ sơ (vd hóa đơn của khách khác) -> đừng liệt kê."""
    out_dir = os.path.join(dossier_dir, "TBKQ", "Edoc")
    os.makedirs(out_dir, exist_ok=True)
    written = {}
    for name, items in parts.items():
        doc = fitz.open()
        for f, pages in items:
            src = fitz.open(f if os.path.isabs(f) else os.path.join(dossier_dir, f))
            for i in (range(src.page_count) if pages is None else [p - 1 for p in pages]):
                if is_blank_page(src[i]):
                    print(f"[scan_splitter] bỏ trang trắng {os.path.basename(f)} tr. {i + 1}")
                    continue
                doc.insert_pdf(src, from_page=i, to_page=i)
            src.close()
        if doc.page_count:
            path = os.path.join(out_dir, f"{code6} - {name}.pdf")
            orient_doc(doc, dossier_dir, os.path.basename(path))   # trang scan ngược / nằm ngang -> đứng thẳng
            doc.save(path)
            written[name] = path
            print(f"[scan_splitter] {name}: {doc.page_count} trang -> {path}")
        doc.close()
    return written


def classify_page_text(norm_text: str) -> str:
    """
    Classify a normalized page text string into 'BBLM', 'GHI_CHEP' or 'KHAC'.
    BBLM trước; trang có dấu hiệu Ghi chép -> GHI_CHEP; header giấy tờ khác (PYC, báo giá, PO…) -> KHAC;
    trang không có header rõ (phiếu in nhiệt…) -> GHI_CHEP.
    """
    if not norm_text:
        return "GHI_CHEP"

    # Tiêu đề / mã form nằm ở đầu trang -> cái nào xuất hiện TRƯỚC thì quyết định (11/11/2026, 139329:
    # trang 2 BBLM có cột "ghi chép" trong bảng nên trước đây bị xếp nhầm vào Ghi Chep).
    head = norm_text[:HEAD_CHARS]
    m_bblm, m_gc = BBLM_HEAD_RE.search(head), GHI_CHEP_HEAD_RE.search(head)
    pos_bblm = m_bblm.start() if m_bblm else -1
    pos_gc = m_gc.start() if m_gc else -1
    if pos_bblm >= 0 and (pos_gc < 0 or pos_bblm < pos_gc):
        return "BBLM"
    if pos_gc >= 0:
        return "GHI_CHEP"

    is_bblm = False
    # Explicit form codes for BBLM
    if "bm3030 02" in norm_text or "bm303002" in norm_text or "3030 02" in norm_text:
        is_bblm = True
    # Header keywords for BBLM
    elif (
        "bien ban giam dinh" in norm_text
        or "bien ban lay mau" in norm_text
        or "bien ban hien truong" in norm_text
    ):
        # Exclude if it explicitly mentions Ghi Chép
        if (
            "ghi chep" not in norm_text
            and "bm3030 01" not in norm_text
            and "bm303001" not in norm_text
        ):
            is_bblm = True
    # Footer / closing clause on page 2 of BBLM
    elif "bien ban nay duoc lap thanh" in norm_text or (
        "dai dien don vi yeu cau" in norm_text
        and "nguoi thuc hien" in norm_text
        and "ghi chep" not in norm_text
    ):
        is_bblm = True

    if is_bblm:
        return "BBLM"
    # khớp nguyên từ (132115: "bao gia" từng khớp vào giữa ghi chú tay "...bao gia..." -> xếp nhầm Khac)
    if any(_has_phrase(norm_text, k) for k in GHI_CHEP_KEYWORDS):
        return "GHI_CHEP"
    if any(_has_phrase(norm_text, k) for k in KHAC_KEYWORDS):
        return "KHAC"
    return "GHI_CHEP"


def _has_phrase(text: str, phrase: str) -> bool:
    return re.search(r"(?<![a-z0-9])" + re.escape(phrase) + r"(?![a-z0-9])", text) is not None


def extract_dossier_code_6(identifier: str) -> str:
    """Extract standard 6-digit dossier code (e.g. '132918' from 'Q326D132918' or 'Input/132918')."""
    if not identifier:
        return "EDOC"
    # Match standard QUATEST 3 code QxxDxxxxxx or QxxxDxxxxxx (e.g. Q326D131638, Q324D122266)
    m = re.search(r"Q\d{2,3}[A-Z](\d{6})", identifier, re.IGNORECASE)
    if m:
        return m.group(1)
    # Find any 6 consecutive digits anywhere in identifier
    m = re.search(r"(\d{6})", identifier)
    if m:
        return m.group(1)
    # Fallback to sanitized folder name
    clean = re.sub(r"[^\w\-]", "", os.path.basename(identifier))
    return clean if clean else "EDOC"


def find_candidate_scan_pdfs(dossier_dir: str) -> List[str]:
    """
    Find all raw scan PDF files in the dossier directory.
    Ignores output reports (TBKQ, BCKQ, tem tu, and previously generated edoc files).
    """
    if not os.path.isdir(dossier_dir):
        return []

    candidates = []
    ignore_prefixes = ("tbkq", "2. bckq", "bckq", "tem tu")
    ignore_keywords = ("bblm", "ghi chep", "ghi_chep", "tem_tu", "edoc", " - khac.pdf")

    for f in os.listdir(dossier_dir):
        if not f.lower().endswith(".pdf"):
            continue
        fl = f.lower()
        if any(fl.startswith(p) for p in ignore_prefixes):
            continue
        if any(k in fl for k in ignore_keywords):
            continue
        full_p = os.path.join(dossier_dir, f)
        if os.path.isfile(full_p):
            candidates.append(full_p)

    # Sort deterministically (chronological scan order or alphabetical)
    candidates.sort()
    return candidates


def ocr_images_batch(image_paths: List[str]) -> Dict[str, str]:
    """Run Windows.Media.Ocr on a batch of image paths via Core/Tools/ocr_helper.ps1."""
    if not image_paths:
        return {}

    helper_ps1 = os.path.join(os.path.dirname(__file__), "ocr_helper.ps1")
    if not os.path.isfile(helper_ps1):
        raise FileNotFoundError(f"OCR helper script not found: {helper_ps1}")

    temp_list_file = os.path.join(tempfile.gettempdir(), f"ocr_list_{os.getpid()}.txt")
    try:
        with open(temp_list_file, "w", encoding="utf-8") as f:
            for p in image_paths:
                f.write(os.path.abspath(p) + "\n")

        cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            os.path.abspath(helper_ps1),
            "-ListFile",
            os.path.abspath(temp_list_file),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
        out_json = temp_list_file.replace(".txt", "_out.json")
        if os.path.isfile(out_json):
            with open(out_json, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            try:
                os.remove(out_json)
            except OSError:
                pass
            return data
        else:
            print(f"[Warning] OCR output JSON missing. Stderr: {res.stderr}")
            return {}
    finally:
        if os.path.isfile(temp_list_file):
            try:
                os.remove(temp_list_file)
            except OSError:
                pass


def _make_checkpoint_store(dossier_dir: str, force_ocr: bool) -> Any:
    """Tạo CheckpointStore; lỗi thì trả None (khi đó chạy như cũ, không cache)."""
    try:
        try:
            from Core.Services.checkpoint import CheckpointStore
        except ImportError:
            ws = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            if ws not in sys.path:
                sys.path.insert(0, ws)
            from Core.Services.checkpoint import CheckpointStore
        return CheckpointStore(dossier_dir=dossier_dir, force=["ocr"] if force_ocr else [])
    except Exception as e:
        print(f"[Warning] Không khởi tạo được checkpoint ({type(e).__name__}: {e}); OCR không dùng cache.")
        return None


HEADER_TOP_FRACTION = 0.45  # tỉ lệ chiều cao trang cắt ra để OCR header


def split_dossier_scans(
    dossier_dir: str,
    output_dir: Optional[str] = None,
    code_override: Optional[str] = None,
    dpi: int = 150,
    store: Any = None,
    force_ocr: bool = False,
) -> Dict[str, Any]:
    """
    Scans a dossier folder for raw scan PDFs, classifies each page, and outputs:
      <output_dir>/<code_6> - BBLM.pdf
      <output_dir>/<code_6> - Ghi Chep.pdf
      <output_dir>/<code_6> - Khac.pdf      (PYC, báo giá, PO, PKSCV, đề cương…; trang trắng bị bỏ)
    Default output_dir: <dossier_dir>/TBKQ/Edoc/

    OCR header từng trang được lưu ở ThucHien/_checkpoint/header_ocr.json (khóa theo mã băm PDF + số trang
    + dpi), lần sau cùng file thì không OCR lại. `force_ocr=True` để OCR lại.
    """
    dossier_dir = os.path.abspath(dossier_dir)
    if store is None:
        store = _make_checkpoint_store(dossier_dir, force_ocr)
    elif force_ocr:
        store.force.add("ocr")
    if not os.path.isdir(dossier_dir):
        raise ValueError(f"Dossier directory does not exist: {dossier_dir}")

    # Determine 6-digit dossier code
    if code_override:
        code_6 = extract_dossier_code_6(code_override)
    else:
        code_6 = extract_dossier_code_6(os.path.basename(dossier_dir))

    # Output directory per user rule: /TBKQ/Edoc
    if not output_dir:
        output_dir = os.path.join(dossier_dir, "TBKQ", "Edoc")
    os.makedirs(output_dir, exist_ok=True)

    scan_pdfs = find_candidate_scan_pdfs(dossier_dir)
    if not scan_pdfs:
        print(f"[Info] No raw scan PDFs found in: {dossier_dir}")
        return {
            "dossier_dir": dossier_dir,
            "code_6": code_6,
            "bblm_path": None,
            "ghi_chep_path": None,
            "khac_path": None,
            "bblm_pages": 0,
            "ghi_chep_pages": 0,
            "khac_pages": 0,
            "blank_pages": 0,
            "total_pages": 0,
        }

    print(f"\n{'='*70}")
    print(f"SPLITTING SCAN PDFS FOR DOSSIER: {code_6} ({os.path.basename(dossier_dir)})")
    print(f"Target Output Directory: {output_dir}")
    print(f"Found {len(scan_pdfs)} scan PDF(s): {[os.path.basename(p) for p in scan_pdfs]}")
    print(f"{'='*70}")

    tmp_render_dir = os.path.join(tempfile.gettempdir(), f"scan_split_{code_6}_{os.getpid()}")
    os.makedirs(tmp_render_dir, exist_ok=True)

    rendered_images = []
    page_registry = []  # List of dicts with page metadata

    try:
        # Hướng đúng của từng trang (AGENTS 8A, 11/11/2026) — xác định TRƯỚC khi OCR header:
        # trang scan ngược thì header nằm ở đáy, OCR 45 % trên sẽ không thấy -> phân loại sai.
        upright: Dict[tuple, int] = {}
        _odocs = [fitz.open(p) for p in scan_pdfs]
        try:
            _pages, _keys = [], []
            for pdf_p, od in zip(scan_pdfs, _odocs):
                for pno in range(od.page_count):
                    if not is_blank_page(od[pno]):
                        _pages.append(od[pno])
                        _keys.append((pdf_p, pno))
            _rots = detect_rotations(_pages, store=store,
                                     labels=[f"{os.path.basename(a)} tr. {b + 1}" for a, b in _keys])
            upright = dict(zip(_keys, _rots))
        finally:
            for od in _odocs:
                od.close()

        global_page_idx = 1
        for pdf_p in scan_pdfs:
            doc = fitz.open(pdf_p)
            base_name = os.path.splitext(os.path.basename(pdf_p))[0]
            pdf_sha = store.file_hash(pdf_p) if store is not None else None
            for pno in range(len(doc)):
                page = doc[pno]
                if is_blank_page(page):
                    page_registry.append({
                        "global_idx": global_page_idx,
                        "pdf_path": pdf_p,
                        "page_no": pno,
                        "img_path": None,
                        "norm_text": "",
                        "category": "BLANK",
                    })
                    global_page_idx += 1
                    continue
                native_txt = page.get_text().strip()
                norm_native = normalize_text(native_txt)

                # Check if native text has clear indicator
                has_clear_native = len(norm_native) > 30 and (
                    "bm3030" in norm_native
                    or "bien ban" in norm_native
                    or "ghi chep" in norm_native
                )

                if has_clear_native:
                    page_registry.append({
                        "global_idx": global_page_idx,
                        "pdf_path": pdf_p,
                        "page_no": pno,
                        "img_path": None,
                        "norm_text": norm_native,
                        "category": classify_page_text(norm_native),
                    })
                else:
                    # Header OCR đã lưu từ lần chạy trước -> dùng lại, không render/OCR nữa
                    want_rot = upright.get((pdf_p, pno), page.rotation)
                    cache_key = store.header_key(pdf_sha, pno, dpi, HEADER_TOP_FRACTION) if store is not None else None
                    if want_rot != page.rotation:
                        page.set_rotation(want_rot)   # chỉ trên bản mở để đọc, không ghi lại scan gốc
                        if cache_key:
                            cache_key += f":rot{want_rot}"
                    cached_hdr = store.get_header_ocr(cache_key) if store is not None else None
                    if cached_hdr is not None:
                        norm_cached = normalize_text(cached_hdr)
                        page_registry.append({
                            "global_idx": global_page_idx,
                            "pdf_path": pdf_p,
                            "page_no": pno,
                            "img_path": None,
                            "norm_text": norm_cached,
                            "category": classify_page_text(norm_cached),
                            "cache_key": cache_key,
                        })
                        global_page_idx += 1
                        continue

                    # Render top 45% for header OCR
                    rect = page.rect
                    clip_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * HEADER_TOP_FRACTION)
                    pix = page.get_pixmap(clip=clip_rect, dpi=dpi)
                    img_file = os.path.join(tmp_render_dir, f"{base_name}_p{pno+1}.png")
                    pix.save(img_file)
                    rendered_images.append(img_file)

                    page_registry.append({
                        "global_idx": global_page_idx,
                        "pdf_path": pdf_p,
                        "page_no": pno,
                        "img_path": os.path.abspath(img_file),
                        "norm_text": "",
                        "category": None,
                        "cache_key": cache_key,
                    })

                global_page_idx += 1
            doc.close()

        # Perform OCR on rendered headers
        if rendered_images:
            ocr_map = ocr_images_batch(rendered_images)
            to_cache: Dict[str, str] = {}
            for item in page_registry:
                if item["category"] is None:
                    ocr_raw = ocr_map.get(item["img_path"], "")
                    norm = normalize_text(ocr_raw)
                    item["norm_text"] = norm
                    item["category"] = classify_page_text(norm)
                    if item.get("cache_key") and ocr_raw:
                        to_cache[item["cache_key"]] = ocr_raw
            if store is not None and to_cache:
                store.put_header_ocr(to_cache)

        # Print classification summary
        print("\n--- Page Classification Summary ---")
        for item in page_registry:
            short_preview = item["norm_text"][:65] if item["norm_text"] else "(no text/slip)"
            print(
                f"  Page {item['global_idx']:02d} | Source: {os.path.basename(item['pdf_path'])}:p{item['page_no']+1} "
                f"| Category: [{item['category']}] | Text: '{short_preview}'"
            )

        # Build output PDF documents
        bblm_doc = fitz.open()
        ghichep_doc = fitz.open()
        khac_doc = fitz.open()

        # Cache open source docs to avoid repeated disk reads
        src_docs = {}
        for pdf_p in scan_pdfs:
            src_docs[pdf_p] = fitz.open(pdf_p)

        bblm_count = 0
        ghichep_count = 0
        khac_count = 0
        blank_count = 0

        rotated = []
        for item in page_registry:
            s_doc = src_docs[item["pdf_path"]]
            pno = item["page_no"]
            if item["category"] == "BLANK":
                blank_count += 1
                continue
            if item["category"] == "BBLM":
                dst = bblm_doc
                bblm_count += 1
            elif item["category"] == "KHAC":
                dst = khac_doc
                khac_count += 1
            else:
                dst = ghichep_doc
                ghichep_count += 1
            dst.insert_pdf(s_doc, from_page=pno, to_page=pno)
            want = upright.get((item["pdf_path"], pno))
            if want is not None and want != dst[-1].rotation:
                rotated.append(f"{os.path.basename(item['pdf_path'])} tr. {pno + 1}: "
                               f"{(want - dst[-1].rotation) % 360}°")
                dst[-1].set_rotation(want)
        for r in rotated:
            print(f"[Orient] xoay {r}")

        for s_doc in src_docs.values():
            s_doc.close()

        # Save output files
        bblm_path = os.path.join(output_dir, f"{code_6} - BBLM.pdf")
        ghichep_path = os.path.join(output_dir, f"{code_6} - Ghi Chep.pdf")
        khac_path = os.path.join(output_dir, f"{code_6} - Khac.pdf")

        if bblm_count > 0:
            bblm_doc.save(bblm_path)
            print(f"[Success] Created BBLM PDF ({bblm_count} pages): {bblm_path}")
        else:
            print("[Warning] No pages matched BBLM category.")
            bblm_path = None
        bblm_doc.close()

        if ghichep_count > 0:
            ghichep_doc.save(ghichep_path)
            print(f"[Success] Created Ghi Chep PDF ({ghichep_count} pages): {ghichep_path}")
        else:
            print("[Warning] No pages matched Ghi Chep category.")
            ghichep_path = None
        ghichep_doc.close()

        if khac_count > 0:
            khac_doc.save(khac_path)
            print(f"[Success] Created Khac PDF ({khac_count} pages): {khac_path}")
        else:
            khac_path = None
        khac_doc.close()
        if blank_count:
            print(f"[Info] Bỏ {blank_count} trang trắng.")

        # Clean up legacy / redundant misnamed files in output_dir (e.g. Q326D<code_6> - BBLM.pdf)
        valid_names = {f"{code_6} - BBLM.pdf", f"{code_6} - Ghi Chep.pdf", f"{code_6} - Khac.pdf"}
        for legacy_f in os.listdir(output_dir):
            if legacy_f in valid_names:
                continue
            if code_6 in legacy_f and (legacy_f.endswith(" - BBLM.pdf") or legacy_f.endswith(" - Ghi Chep.pdf")):
                legacy_p = os.path.join(output_dir, legacy_f)
                try:
                    os.remove(legacy_p)
                    print(f"[Clean] Removed redundant legacy scan file: {legacy_f}")
                except OSError as e:
                    print(f"[Warning] Could not remove legacy file {legacy_f}: {e}")

        print(f"\nScan PDF splitting finished successfully for dossier: {code_6}\n")

        return {
            "dossier_dir": dossier_dir,
            "code_6": code_6,
            "bblm_path": bblm_path,
            "ghi_chep_path": ghichep_path,
            "khac_path": khac_path,
            "bblm_pages": bblm_count,
            "ghi_chep_pages": ghichep_count,
            "khac_pages": khac_count,
            "blank_pages": blank_count,
            "total_pages": len(page_registry),
        }

    finally:
        # Clean up temporary rendered images
        shutil.rmtree(tmp_render_dir, ignore_errors=True)


def split_all_dossiers_in_input(input_root: str = "Input") -> List[Dict[str, Any]]:
    """Scan and split all dossiers under the input directory."""
    input_root = os.path.abspath(input_root)
    if not os.path.isdir(input_root):
        print(f"Input root does not exist: {input_root}")
        return []

    results = []
    for item in os.listdir(input_root):
        full_p = os.path.join(input_root, item)
        # Process subdirectories that represent dossiers (ignore TBKQ folder itself)
        if os.path.isdir(full_p) and item.lower() != "tbkq":
            res = split_dossier_scans(full_p)
            if res.get("total_pages", 0) > 0:
                results.append(res)

    return results


def main():
    import argparse
    parser = argparse.ArgumentParser(description="QUATEST 3 Automated Scan PDF Splitter")
    parser.add_argument("target", nargs="?", help="Path to dossier folder (e.g. Input/132918)")
    parser.add_argument("--code", help="Override 6-digit dossier code (e.g. 132918)")
    parser.add_argument("-o", "--output-dir", help="Explicit output directory (default: <dossier>/TBKQ/Edoc)")
    parser.add_argument("--all-input", action="store_true", help="Split scan PDFs across all dossier folders in Input/")
    parser.add_argument("--force-ocr", action="store_true", help="Bỏ qua checkpoint header OCR, OCR lại")

    args = parser.parse_args()

    if args.all_input:
        split_all_dossiers_in_input("Input")
    elif args.target:
        split_dossier_scans(args.target, output_dir=args.output_dir, code_override=args.code, force_ocr=args.force_ocr)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
