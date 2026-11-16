"""
PDF & Image utility tool for QUATEST 3 (Room N2) profile processing.
Replaces pdf-img.js, pdf-render.ps1, img-crop.ps1, and montage.ps1 using PyMuPDF (fitz) and Pillow (PIL).
No external executables or WinRT dependencies required.
"""
import os
import sys
import argparse
import fitz  # PyMuPDF
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8")


def pdf_info(pdf_path: str):
    doc = fitz.open(pdf_path)
    print(f"PDF: {pdf_path}")
    print(f"Total pages: {len(doc)}")
    for i, page in enumerate(doc):
        txt = page.get_text().strip()
        imgs = page.get_images(full=True)
        rect = page.rect
        print(
            f"  Page {i+1:02d}: size={rect.width:.0f}x{rect.height:.0f}pt | "
            f"text_chars={len(txt)} | embedded_imgs={len(imgs)} | rotation={page.rotation}"
        )
    doc.close()


def pdf_extract_text(pdf_path: str, out_path: str = None):
    doc = fitz.open(pdf_path)
    lines = []
    for i, page in enumerate(doc):
        lines.append(f"=== PAGE {i+1}/{len(doc)} ===")
        lines.append(page.get_text("text"))
    doc.close()
    full_text = "\n".join(lines)
    if out_path:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(full_text)
        print(f"Saved text ({len(full_text)} chars) -> {out_path}")
    else:
        print(full_text)
    return full_text


def pdf_extract_images(pdf_path: str, out_dir: str):
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    saved = []
    img_idx = 1
    for p_i, page in enumerate(doc):
        image_list = page.get_images(full=True)
        for img_info in image_list:
            xref = img_info[0]
            base_img = doc.extract_image(xref)
            image_bytes = base_img["image"]
            ext = base_img["ext"]
            out_name = f"p{p_i+1:02d}_img{img_idx:02d}.{ext}"
            out_file = os.path.join(out_dir, out_name)
            with open(out_file, "wb") as f:
                f.write(image_bytes)
            saved.append(out_file)
            print(
                f"Extracted Page {p_i+1} xref={xref}: {out_name} "
                f"({base_img['width']}x{base_img['height']}, {len(image_bytes):,} bytes)"
            )
            img_idx += 1
    doc.close()
    print(f"Total embedded images extracted: {len(saved)}")
    return saved


def pdf_render_pages(pdf_path: str, out_dir: str, dpi: int = 200, pages_spec: str = None):
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    total = len(doc)
    if pages_spec:
        page_nums = []
        for part in pages_spec.split(","):
            part = part.strip()
            if "-" in part:
                s, e = part.split("-", 1)
                page_nums.extend(range(int(s), int(e) + 1))
            else:
                page_nums.append(int(part))
    else:
        page_nums = list(range(1, total + 1))

    saved = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for pno in page_nums:
        if 1 <= pno <= total:
            page = doc[pno - 1]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out_file = os.path.join(out_dir, f"p{pno:02d}.png")
            pix.save(out_file)
            saved.append(out_file)
            print(f"Rendered Page {pno:02d}/{total} -> {out_file} ({pix.width}x{pix.height})")
    doc.close()
    return saved


def img_crop_rotate(
    in_path: str,
    out_path: str,
    box: tuple = None,
    rotate_cw: int = 0,
    scale: float = 1.0,
):
    """
    Rotate, crop, and scale an image.
    Note: rotate_cw is clockwise degrees (0, 90, 180, 270).
    If box=(x, y, w, h) is provided, crop is applied BEFORE rotation unless rotate_first is needed.
    To match img-crop.ps1 behavior: crop (x, y, w, h) on input image, then rotate clockwise, then scale.
    """
    img = Image.open(in_path)
    orig_w, orig_h = img.size
    if box:
        x, y, w, h = box
        x2 = min(orig_w, x + w)
        y2 = min(orig_h, y + h)
        img = img.crop((max(0, x), max(0, y), x2, y2))
    if rotate_cw:
        # PIL rotate is counter-clockwise, so clockwise = -rotate_cw
        img = img.rotate(-rotate_cw, expand=True)
    if scale and abs(scale - 1.0) > 1e-3:
        nw = max(1, int(round(img.width * scale)))
        nh = max(1, int(round(img.height * scale)))
        img = img.resize((nw, nh), Image.Resampling.LANCZOS)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    img.save(out_path)
    print(f"Saved cropped/rotated image -> {out_path} ({img.width}x{img.height})")
    return out_path


def main():
    parser = argparse.ArgumentParser(description="PDF & Image Tool for QUATEST 3")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_info = sub.add_parser("info", help="Show PDF pages, text chars, embedded image counts")
    p_info.add_argument("pdf", help="Path to PDF file")

    p_txt = sub.add_parser("text", help="Extract text from PDF")
    p_txt.add_argument("pdf", help="Path to PDF file")
    p_txt.add_argument("-o", "--out", help="Output text file path")

    p_ext = sub.add_parser("extract-images", help="Extract embedded images from PDF")
    p_ext.add_argument("pdf", help="Path to PDF file")
    p_ext.add_argument("-o", "--out-dir", required=True, help="Output directory")

    p_ren = sub.add_parser("render", help="Render PDF pages to PNG")
    p_ren.add_argument("pdf", help="Path to PDF file")
    p_ren.add_argument("-o", "--out-dir", required=True, help="Output directory")
    p_ren.add_argument("--dpi", type=int, default=200, help="Render DPI (default 200)")
    p_ren.add_argument("--pages", help="Page range e.g. '1-3,5'")

    p_crop = sub.add_parser("crop", help="Crop, rotate, and scale an image")
    p_crop.add_argument("image", help="Input image path")
    p_crop.add_argument("-o", "--out", required=True, help="Output image path")
    p_crop.add_argument("--box", help="Crop box x,y,w,h in pixels")
    p_crop.add_argument("--rotate", type=int, default=0, choices=[0, 90, 180, 270], help="Clockwise rotation")
    p_crop.add_argument("--scale", type=float, default=1.0, help="Scale factor")

    args = parser.parse_args()
    if args.cmd == "info":
        pdf_info(args.pdf)
    elif args.cmd == "text":
        pdf_extract_text(args.pdf, args.out)
    elif args.cmd == "extract-images":
        pdf_extract_images(args.pdf, args.out_dir)
    elif args.cmd == "render":
        pdf_render_pages(args.pdf, args.out_dir, args.dpi, args.pages)
    elif args.cmd == "crop":
        box = tuple(map(int, args.box.split(","))) if args.box else None
        img_crop_rotate(args.image, args.out, box, args.rotate, args.scale)


if __name__ == "__main__":
    main()
