# -*- coding: utf-8 -*-
"""
Smart MarkItDown & Local OCR Tool for Documents and PDFs.
Combines Microsoft MarkItDown with built-in Windows Local OCR (0 credit, 100% offline).

Capabilities:
1. Digital files (.docx, .xlsx, .pptx, .csv, digital .pdf):
   - Uses MarkItDown (offline, preserves markdown tables, headings, 0 credit).
2. Scanned files (image-only .pdf, .jpg, .png):
   - Automatically detects empty text layers.
   - Falls back to Windows Media OCR (built-in, free, fast, 0 credit).
   - Optional support for Local Vision LLM (e.g. Ollama) or Cloud LLM if configured.
"""

import os
import sys
import argparse
import tempfile
from typing import Optional, List, Dict, Any

from markitdown import MarkItDown
import fitz  # PyMuPDF

# Ensure stdout handles UTF-8 on Windows console
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception as _err:
        print(f"[Warning] md_ocr.py:29 {type(_err).__name__}: {_err}")


def get_ocr_helper_path() -> str:
    """Get absolute path to ocr_helper.ps1."""
    cur_dir = os.path.dirname(os.path.abspath(__file__))
    ps1_path = os.path.join(cur_dir, "ocr_helper.ps1")
    if not os.path.isfile(ps1_path):
        raise FileNotFoundError(f"ocr_helper.ps1 not found at: {ps1_path}")
    return ps1_path


def run_windows_ocr_on_images(image_paths: List[str]) -> Dict[str, str]:
    """Run local Windows.Media.Ocr (0 credit, 100% offline) on a list of images."""
    if not image_paths:
        return {}
    
    import subprocess
    import json
    
    helper_ps1 = get_ocr_helper_path()
    temp_list_file = os.path.join(tempfile.gettempdir(), f"ocr_batch_{os.getpid()}.txt")
    
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
            helper_ps1,
            "-ListFile",
            temp_list_file,
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
            return {}
    finally:
        if os.path.isfile(temp_list_file):
            try:
                os.remove(temp_list_file)
            except OSError:
                pass


def ocr_scanned_pdf_pages(pdf_path: str, dpi: int = 200, store: Any = None, force: bool = False) -> List[str]:
    """
    OCR một PDF scan, trả về danh sách văn bản theo trang.

    Nếu có `store` (Core.Services.checkpoint.CheckpointStore): dùng lại kết quả đã lưu khi file
    chưa đổi (so mã SHA-256) và cùng dpi -> KHÔNG OCR lại. `force=True` bắt buộc OCR lại.
    Kết quả mới được ghi vào checkpoint (ocr.json + ocr.xlsx).
    """
    if store is not None:
        cached = store.get_ocr(pdf_path, dpi, force=force)
        if cached is not None:
            print(f"[OCR] Dùng lại kết quả OCR đã lưu ({len(cached)} trang), bỏ qua OCR: {os.path.basename(pdf_path)}")
            return cached
    pages = _ocr_pdf_pages_uncached(pdf_path, dpi)
    if store is not None:
        store.put_ocr(pdf_path, pages, dpi)
    return pages


def pages_to_markdown(pdf_path: str, pages: List[str]) -> str:
    """Ghép văn bản theo trang thành Markdown (định dạng giữ nguyên như bản cũ)."""
    total_pages = len(pages)
    md_lines = [f"# Tài liệu: {os.path.basename(pdf_path)}\n"]
    for page_idx, text in enumerate(pages):
        md_lines.append(f"## Trang {page_idx + 1} / {total_pages}\n")
        if text:
            md_lines.append(text)
        else:
            md_lines.append("*(Không nhận diện được nội dung chữ trên trang này)*")
        md_lines.append("\n---\n")
    return "\n".join(md_lines)


def ocr_scanned_pdf_to_markdown(pdf_path: str, dpi: int = 200, store: Any = None, force: bool = False) -> str:
    """
    Extract images from a scanned PDF and run local Windows OCR (có cache checkpoint nếu truyền `store`).
    Formats the output into clean structured Markdown by page.
    """
    return pages_to_markdown(pdf_path, ocr_scanned_pdf_pages(pdf_path, dpi=dpi, store=store, force=force))


def _ocr_pdf_pages_uncached(pdf_path: str, dpi: int = 200) -> List[str]:
    """Render từng trang ra ảnh rồi chạy Windows OCR (luôn OCR, không dùng cache)."""
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    
    temp_dir = tempfile.mkdtemp(prefix="pdf_ocr_")
    image_paths = []
    
    try:
        for page_idx in range(total_pages):
            page = doc[page_idx]
            pix = page.get_pixmap(dpi=dpi)
            img_file = os.path.join(temp_dir, f"page_{page_idx + 1:03d}.png")
            pix.save(img_file)
            image_paths.append(img_file)
            
        doc.close()
        
        # Run batch OCR
        ocr_results = run_windows_ocr_on_images(image_paths)
        
        if not ocr_results:
            print("[Warning] md_ocr: Windows OCR không trả kết quả nào (kiểm tra ocr_helper.ps1 / ngôn ngữ OCR).")
        return [ocr_results.get(os.path.abspath(p), "").strip() for p in image_paths]
    finally:
        # Cleanup temporary images
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)


def convert_document(
    file_path: str,
    output_path: Optional[str] = None,
    force_ocr: bool = False,
    llm_client: Any = None,
    llm_model: Optional[str] = None,
    dpi: int = 200,
    store: Any = None,
) -> str:
    """
    Universal document-to-markdown converter.
    
    - If digital file: runs MarkItDown (0 credit).
    - If scanned PDF (empty text): runs local Windows OCR (0 credit).
    - If llm_client provided: supports multimodal LLM vision.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    ext = os.path.splitext(file_path)[1].lower()
    
    # Non-PDF documents: let MarkItDown handle directly
    if ext in [".docx", ".xlsx", ".pptx", ".csv", ".html", ".txt"]:
        md = MarkItDown(llm_client=llm_client, llm_model=llm_model)
        result = md.convert(file_path)
        content = result.text_content
        if output_path:
            _save_output(content, output_path)
        return content
    
    # PDF documents
    if ext == ".pdf":
        if not force_ocr:
            try:
                md = MarkItDown(llm_client=llm_client, llm_model=llm_model)
                res = md.convert(file_path)
                text = res.text_content.strip()
                # Check if text is substantial
                if len(text) >= 60:
                    if output_path:
                        _save_output(text, output_path)
                    return text
            except Exception as _err:
                print(f"[Warning] md_ocr.py:172 {type(_err).__name__}: {_err}")
        
        # If text is empty or force_ocr is True, run local Windows OCR (0 credit)
        print(f"[OCR] Phát hiện PDF dạng ảnh scan. Kích hoạt Local Windows OCR (0 credit)...")
        content = ocr_scanned_pdf_to_markdown(file_path, dpi=dpi, store=store, force=force_ocr)
        if output_path:
            _save_output(content, output_path)
        return content
        
    # Images (.jpg, .png, etc.)
    if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff"]:
        if llm_client is not None:
            md = MarkItDown(llm_client=llm_client, llm_model=llm_model)
            res = md.convert(file_path)
            content = res.text_content
        else:
            cached = store.get_ocr(file_path, 0, force=force_ocr) if store is not None else None
            if cached is not None:
                print(f"[OCR] Dùng lại kết quả OCR đã lưu, bỏ qua OCR: {os.path.basename(file_path)}")
                content = cached[0]
            else:
                ocr_dict = run_windows_ocr_on_images([file_path])
                content = ocr_dict.get(os.path.abspath(file_path), "")
                if store is not None:
                    store.put_ocr(file_path, [content], 0)
        if output_path:
            _save_output(content, output_path)
        return content

    # Default fallback to MarkItDown
    md = MarkItDown(llm_client=llm_client, llm_model=llm_model)
    res = md.convert(file_path)
    content = res.text_content
    if output_path:
        _save_output(content, output_path)
    return content


def _save_output(content: str, output_path: str):
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[OK] Đã lưu kết quả Markdown ({len(content)} ký tự) -> {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Chuyển đổi tài liệu & PDF Scan sang Markdown siêu tiết kiệm credit (0 credit offline)."
    )
    parser.add_argument("file_path", help="Đường dẫn file cần đọc (.pdf, .docx, .xlsx, .png, .jpg...)")
    parser.add_argument("-o", "--output", help="Đường dẫn file .md xuất ra (mặc định in ra màn hình hoặc tạo cạnh file)")
    parser.add_argument("--force-ocr", action="store_true", help="Bắt buộc chạy Local OCR kể cả khi PDF có text layer")
    parser.add_argument("--dpi", type=int, default=200, help="Độ phân giải DPI khi render ảnh để OCR (mặc định 200)")
    parser.add_argument("--dossier", help="Thư mục hồ sơ: lưu/dùng lại kết quả OCR ở ThucHien/_checkpoint (không OCR lại file chưa đổi)")

    args = parser.parse_args()

    out_file = args.output
    if not out_file and not sys.stdout.isatty():
        # If output redirected or not specified, we can set default
        pass

    store = None
    if args.dossier:
        workspace = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        if workspace not in sys.path:
            sys.path.insert(0, workspace)
        from Core.Services.checkpoint import CheckpointStore
        store = CheckpointStore(dossier_dir=args.dossier)

    try:
        result = convert_document(
            file_path=args.file_path,
            output_path=out_file,
            force_ocr=args.force_ocr,
            dpi=args.dpi,
            store=store,
        )
        if not out_file:
            print(result)
    except Exception as e:
        print(f"[Lỗi] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
