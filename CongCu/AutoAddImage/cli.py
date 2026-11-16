import os
import sys
import argparse
from typing import List
from drawio_builder import DrawIOBuilder
from drawio_exporter import DrawIOExporter


SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".svg"}


def collect_image_files(inputs: List[str], directory: str = None) -> List[str]:
    """Collect valid image file paths from direct inputs and directories."""
    files = []

    if directory and os.path.isdir(directory):
        for root, _, filenames in os.walk(directory):
            for filename in filenames:
                ext = os.path.splitext(filename)[1].lower()
                if ext in SUPPORTED_EXTENSIONS:
                    files.append(os.path.join(root, filename))

    for item in inputs:
        if os.path.isfile(item):
            ext = os.path.splitext(item)[1].lower()
            if ext in SUPPORTED_EXTENSIONS:
                files.append(item)
            else:
                print(f"[Warning] Skipping unsupported file type: {item}")
        elif os.path.isdir(item):
            for root, _, filenames in os.walk(item):
                for filename in filenames:
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in SUPPORTED_EXTENSIONS:
                        files.append(os.path.join(root, filename))

    # Remove duplicates preserving order
    seen = set()
    unique_files = []
    for f in files:
        norm = os.path.abspath(f)
        if norm not in seen:
            seen.add(norm)
            unique_files.append(f)

    return unique_files


def run_cli(args=None) -> int:
    parser = argparse.ArgumentParser(
        description="Draw.io Multi-Tool: Generate .drawio files from images or export all tabs to images."
    )

    parser.add_argument(
        "images",
        nargs="*",
        help="Image file paths or directory paths containing images."
    )
    parser.add_argument(
        "-o", "--output",
        help="Output .drawio file path (for generator) or destination folder (for exporter)."
    )
    parser.add_argument(
        "-d", "--dir",
        help="Directory containing images to process."
    )
    parser.add_argument(
        "-a", "--append",
        help="Path to an existing .drawio file to append image tabs into."
    )
    parser.add_argument(
        "-e", "--export",
        dest="export_file",
        help="Path to a .drawio file whose tabs will all be exported as images."
    )
    parser.add_argument(
        "--format",
        default="png",
        choices=["png", "jpg", "jpeg", "svg", "pdf"],
        help="Export image format (default: png)."
    )
    parser.add_argument(
        "--scale",
        type=float,
        default=1.0,
        help="Export scale factor (default: 1.0)."
    )
    parser.add_argument(
        "--no-embed",
        action="store_true",
        help="Link images using file:// URIs instead of embedding Base64."
    )
    parser.add_argument(
        "--margin",
        type=int,
        default=20,
        help="Margin around image in pixels (default: 20)."
    )
    parser.add_argument(
        "--no-auto-fit",
        action="store_true",
        help="Use standard canvas size instead of auto-fitting canvas to image size."
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch graphical user interface."
    )

    parsed_args = parser.parse_args(args)

    if parsed_args.gui:
        from gui import launch_gui
        launch_gui()
        return 0

    # --- Mode 1: Export Tabs from .drawio File ---
    if parsed_args.export_file:
        input_drawio = os.path.abspath(parsed_args.export_file)
        if not os.path.isfile(input_drawio):
            print(f"[Error] Draw.io file not found: {input_drawio}")
            return 1

        file_stem = os.path.splitext(os.path.basename(input_drawio))[0]
        output_dir = parsed_args.output or os.path.join(os.path.dirname(input_drawio), f"{file_stem}_exported")
        print(f"[*] Reading tabs from: {input_drawio}")

        try:
            tabs = DrawIOExporter.get_diagram_tabs(input_drawio)
            print(f"[*] Found {len(tabs)} page(s)/tab(s).")
            print(f"[*] Export destination: {os.path.abspath(output_dir)}\n")

            def on_progress(current, total, name):
                print(f"  [{current}/{total}] Exporting tab: '{name}'...")

            exported = DrawIOExporter.export_all_tabs(
                drawio_path=input_drawio,
                output_dir=output_dir,
                format=parsed_args.format,
                scale=parsed_args.scale,
                progress_callback=on_progress
            )

            print(f"\n[+] Success! All {len(exported)} tab(s) exported to: {os.path.abspath(output_dir)}")
            return 0
        except Exception as e:
            print(f"[Error] Failed to export tabs: {e}")
            return 1

    # --- Mode 2: Build / Append .drawio File from Images ---
    image_paths = collect_image_files(parsed_args.images, parsed_args.dir)
    if not image_paths:
        print("[Error] No valid image files provided. Use -e <file.drawio> to export tabs, or provide image files. Use --help for usage details.")
        return 1

    target_drawio = parsed_args.append or parsed_args.output
    if not target_drawio:
        unique_dirs = list(dict.fromkeys(os.path.dirname(os.path.abspath(p)) for p in image_paths))
        if len(unique_dirs) == 1:
            folder = unique_dirs[0]
            folder_name = os.path.basename(folder) or "diagrams"
            if len(image_paths) == 1:
                base_name = os.path.splitext(os.path.basename(image_paths[0]))[0]
            else:
                base_name = folder_name
            target_drawio = os.path.join(folder, f"{base_name}.drawio")
            print(f"[*] Auto-detected output path in image folder: {target_drawio}")
        else:
            print("[Error] Images are from multiple directories. Please specify output file path with -o <path>.")
            return 1

    print(f"[*] Processing {len(image_paths)} image(s)...")

    try:
        builder = DrawIOBuilder(drawio_path=parsed_args.append)
        embed_mode = not parsed_args.no_embed
        auto_fit = not parsed_args.no_auto_fit

        for idx, img_path in enumerate(image_paths, 1):
            file_name = os.path.basename(img_path)
            tab_name = builder.add_image_tab(
                image_path=img_path,
                embed=embed_mode,
                margin=parsed_args.margin,
                auto_fit=auto_fit
            )
            print(f"  [{idx}/{len(image_paths)}] Added '{file_name}' -> Tab: '{tab_name}'")

        final_path = builder.save(target_drawio)
        print(f"\n[+] Successfully saved Draw.io file to: {os.path.abspath(final_path)}")
        return 0

    except Exception as e:
        print(f"[Error] Failed to generate Draw.io file: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(run_cli())
