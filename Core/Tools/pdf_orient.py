# -*- coding: utf-8 -*-
"""
Tự xoay trang PDF scan về đúng chiều đọc (AGENTS mục 8A — người dùng yêu cầu 11/11/2026).

Áp cho mọi PDF tách ra từ scan gốc (`TBKQ\\Edoc\\<6 số> - BBLM / Ghi Chep / Khac.pdf`):
trang bị scan ngược (180°) hoặc nằm ngang (90°/270°) được đặt lại `/Rotate` để chữ đứng thẳng.
Chỉ đổi thuộc tính xoay hiển thị của trang (không render lại, không giảm chất lượng ảnh scan).

Cách nhận hướng:
  - Trang có chữ gốc (PDF số): theo hướng dòng chữ (`get_text("dict")` -> `dir`).
  - Trang scan (ảnh): render cả trang ở 4 hướng 0/90/180/270, OCR bằng Windows.Media.Ocr,
    chọn hướng đọc được nhiều ký tự nhất. Đo 11/11/2026 trên 10 trang (BBLM, ghi chép, phiếu in
    nhiệt AeroTrak/Lasair, bảng 6ChannelFixed, PKSCV): chữ nằm dọc -> OCR gần như rỗng; lộn ngược
    -> chữ rác chỉ bằng 30–65 % số ký tự của hướng đúng.
  - Chỉ xoay khi hướng tốt nhất hơn hẳn hướng hiện tại (>= 60 ký tự và >= 1,3 lần);
    trang quá ít chữ (phiếu in nhiệt gần trắng, chữ ký…) giữ nguyên hướng, in `[Orient] chưa chắc`.

Kết quả lưu ở `ThucHien\\_checkpoint\\page_orient.json` (khóa = băm nội dung trang chưa xoay),
lần chạy sau không OCR lại.

Dùng:
    from Core.Tools.pdf_orient import auto_orient_pdf, orient_doc
    auto_orient_pdf(r"Input\\136247\\TBKQ\\Edoc\\136247 - Ghi Chep.pdf")   # sửa file đã ghi
    orient_doc(doc, dossier_dir)                                         # sửa fitz.Document trước khi save
    python Core/Tools/pdf_orient.py <file.pdf | thư mục Edoc> [--ghi]     # không --ghi = chỉ báo
"""

import hashlib
import os
import shutil
import sys
import tempfile
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF

ORIENT_DPI = 100            # độ phân giải render cả trang để OCR nhận hướng
MIN_CHARS = 60              # hướng tốt nhất phải đọc được ít nhất chừng này ký tự
MIN_RATIO = 1.3             # và hơn hướng hiện tại ít nhất 1,3 lần
HI_DPI = 300                # lượt 2 cho trang đọc được < MIN_CHARS ở ORIENT_DPI
LOW_MIN_CHARS = 25          # trang ít chữ: chấp nhận từ 25 ký tự nếu hơn hướng hiện tại >= 2 lần
LOW_MIN_RATIO = 2.0
NATIVE_MIN_CHARS = 50       # trang có ít nhất chừng này chữ gốc -> dùng hướng dòng chữ, không OCR
CACHE_NAME = "page_orient.json"
ROTS = (0, 90, 180, 270)


# ------------------------------------------------------------------ tiện ích

def _alnum_count(text: str) -> int:
    if not text or text.startswith("ERROR:") or text == "FILE_NOT_FOUND":
        return 0
    return sum(1 for c in text if c.isalnum())


def page_key(page: "fitz.Page") -> str:
    """Băm nội dung trang ở hướng gốc (bỏ /Rotate) -> cùng trang ở file nào cũng ra cùng khóa."""
    rot = page.rotation
    try:
        if rot:
            page.set_rotation(0)
        pix = page.get_pixmap(dpi=24, colorspace=fitz.csGRAY)
        h = hashlib.sha1(pix.samples)
        h.update(f"{pix.width}x{pix.height}".encode())
        return h.hexdigest()[:24]
    finally:
        if page.rotation != rot:
            page.set_rotation(rot)


def _native_rotation(page: "fitz.Page") -> Optional[int]:
    """Trang có chữ gốc: /Rotate làm hướng dòng chữ chiếm đa số thành trái -> phải. Không đủ chữ -> None."""
    try:
        if len(page.get_text().strip()) < NATIVE_MIN_CHARS:
            return None
        votes: Dict[Tuple[int, int], int] = {}
        for b in page.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                n = sum(len(s.get("text", "")) for s in ln.get("spans", []))
                k = (round(ln["dir"][0]), round(ln["dir"][1]))
                votes[k] = votes.get(k, 0) + n
        if not votes:
            return None
        dx, dy = max(votes, key=votes.get)
        # `dir` tính theo trang chưa xoay (y hướng xuống). /Rotate quay trang theo chiều kim đồng hồ.
        return {(1, 0): 0, (0, 1): 270, (-1, 0): 180, (0, -1): 90}.get((dx, dy))
    except Exception:
        return None


def _workspace_on_path() -> None:
    ws = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    if ws not in sys.path:
        sys.path.insert(0, ws)


def _store_for(dossier_dir: Optional[str]):
    if not dossier_dir:
        return None
    try:
        _workspace_on_path()
        from Core.Services.checkpoint import CheckpointStore
        return CheckpointStore(dossier_dir=dossier_dir)
    except Exception as e:
        print(f"[Orient] không mở được checkpoint ({type(e).__name__}: {e}) — không cache.")
        return None


def dossier_dir_of(pdf_path: str) -> Optional[str]:
    """`…\\Input\\<hồ sơ>\\TBKQ\\Edoc\\x.pdf` -> `…\\Input\\<hồ sơ>`; file nằm ở gốc hồ sơ / thư mục <n> cũng nhận."""
    p = os.path.dirname(os.path.abspath(pdf_path))
    for _ in range(4):
        if os.path.isdir(os.path.join(p, "ThucHien")) or os.path.isdir(os.path.join(p, "TBKQ")):
            return p
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    return None


# ------------------------------------------------------------------ nhận hướng

def _ocr_scores(pages: List["fitz.Page"], idx: List[int], dpi: int) -> Dict[int, Dict[int, int]]:
    """OCR trang `idx` ở 4 hướng -> {i: {rotate tuyệt đối: số ký tự chữ/số đọc được}}."""
    _workspace_on_path()
    from Core.Tools.scan_splitter import ocr_images_batch
    tmp = tempfile.mkdtemp(prefix="orient_")
    try:
        img_map: Dict[Tuple[int, int], str] = {}
        for i in idx:
            pg = pages[i]
            orig = pg.rotation
            try:
                for r in ROTS:
                    pg.set_rotation(r)
                    p = os.path.join(tmp, f"p{i}_r{r}.png")
                    pg.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY).save(p)
                    img_map[(i, r)] = os.path.abspath(p)
            finally:
                pg.set_rotation(orig)
        ocr = ocr_images_batch(list(img_map.values()))
        return {i: {r: _alnum_count(ocr.get(img_map[(i, r)], "")) for r in ROTS} for i in idx}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def detect_rotations(pages: List["fitz.Page"], store=None, dpi: int = ORIENT_DPI,
                     labels: Optional[List[str]] = None) -> List[int]:
    """Trả /Rotate tuyệt đối làm từng trang đứng thẳng (giữ nguyên /Rotate hiện tại khi chưa chắc).
    Không đổi trang (khôi phục /Rotate sau khi thử)."""
    cache: Dict[str, int] = (store.read_json(CACHE_NAME, {}) or {}) if store is not None else {}
    result: List[Optional[int]] = [None] * len(pages)
    keys: List[Optional[str]] = [None] * len(pages)
    todo: List[int] = []
    for i, pg in enumerate(pages):
        try:
            keys[i] = page_key(pg)
        except Exception:
            keys[i] = None
        if keys[i] and keys[i] in cache:
            result[i] = int(cache[keys[i]])
            continue
        nat = _native_rotation(pg)
        if nat is not None:
            result[i] = nat
        else:
            todo.append(i)

    if todo:
        # Lượt 1 ở ORIENT_DPI; trang đọc được quá ít chữ (bảng chữ nhỏ, phiếu in nhiệt) -> lượt 2 ở HI_DPI.
        scores_all = _ocr_scores(pages, todo, dpi)
        retry = [i for i in todo if max(scores_all[i].values()) < MIN_CHARS]
        if retry:
            hi = _ocr_scores(pages, retry, HI_DPI)
            for i in retry:
                if max(hi[i].values()) > max(scores_all[i].values()):
                    scores_all[i] = hi[i]
        for i in todo:
            cur = pages[i].rotation
            scores = scores_all[i]
            best = max(ROTS, key=lambda r: (scores[r], r == cur))
            sure = (scores[best] >= MIN_CHARS and scores[best] >= MIN_RATIO * scores[cur]) or \
                   (scores[best] >= LOW_MIN_CHARS and scores[best] >= LOW_MIN_RATIO * scores[cur])
            if best != cur and sure:
                result[i] = best
            else:
                if best != cur and scores[best] > 0:
                    lab = labels[i] if labels else f"trang {i + 1}"
                    print(f"[Orient] chưa chắc hướng {lab} (điểm {scores}) — giữ nguyên, soát bằng mắt.")
                result[i] = cur if scores[cur] >= MIN_CHARS and scores[cur] >= MIN_RATIO * max(
                    v for r, v in scores.items() if r != cur) else None

    if store is not None:
        new = {keys[i]: result[i] for i in range(len(pages)) if keys[i] and result[i] is not None
               and keys[i] not in cache}
        if new:
            cache.update(new)
            store.write_json(CACHE_NAME, cache)
    return [pages[i].rotation if r is None else r for i, r in enumerate(result)]


def orient_doc(doc: "fitz.Document", dossier_dir: Optional[str] = None, name: str = "",
               store=None) -> List[Tuple[int, int, int]]:
    """Đặt /Rotate từng trang của `doc` cho chữ đứng thẳng. Trả [(trang 1-based, cũ, mới)] các trang đã xoay."""
    if store is None:
        store = _store_for(dossier_dir)
    pages = [doc[i] for i in range(doc.page_count)]
    labels = [f"{name} tr. {i + 1}" for i in range(doc.page_count)]
    rots = detect_rotations(pages, store=store, labels=labels)
    changed = []
    for i, (pg, r) in enumerate(zip(pages, rots)):
        if r != pg.rotation:
            changed.append((i + 1, pg.rotation, r))
            pg.set_rotation(r)
    for pno, old, new in changed:
        print(f"[Orient] {name} tr. {pno}: xoay {(new - old) % 360}° (/Rotate {old} -> {new})")
    return changed


def auto_orient_pdf(path: str, dossier_dir: Optional[str] = None, write: bool = True) -> List[Tuple[int, int, int]]:
    """Xoay đúng hướng các trang của file PDF (file do AI tạo, vd Edoc) rồi ghi đè; không đổi gì thì không ghi."""
    if dossier_dir is None:
        dossier_dir = dossier_dir_of(path)
    doc = fitz.open(path)
    try:
        changed = orient_doc(doc, dossier_dir, name=os.path.basename(path))
        if changed and write:
            tmp = path + ".orient.tmp"
            doc.save(tmp, garbage=1, deflate=True)
            doc.close()
            os.replace(tmp, path)
            return changed
    finally:
        if not doc.is_closed:
            doc.close()
    return changed


def main():
    import argparse
    ap = argparse.ArgumentParser(description="Tự xoay trang PDF scan về đúng chiều đọc (AGENTS mục 8A)")
    ap.add_argument("target", help="file PDF hoặc thư mục (vd Input\\136247\\TBKQ\\Edoc)")
    ap.add_argument("--ghi", action="store_true", help="ghi đè file (mặc định chỉ báo)")
    a = ap.parse_args()
    files = [a.target] if os.path.isfile(a.target) else sorted(
        os.path.join(a.target, f) for f in os.listdir(a.target) if f.lower().endswith(".pdf"))
    total = 0
    for f in files:
        ch = auto_orient_pdf(f, write=a.ghi)
        total += len(ch)
        if not ch:
            print(f"[Orient] {os.path.basename(f)}: mọi trang đã đúng hướng")
    if total and not a.ghi:
        print(f"[Orient] {total} trang cần xoay — chạy lại với --ghi để ghi đè.")


if __name__ == "__main__":
    main()
