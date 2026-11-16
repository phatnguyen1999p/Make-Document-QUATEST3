import os
import sys
import shutil
import base64
import subprocess
import xml.etree.ElementTree as ET
from urllib.parse import unquote
from typing import List, Dict, Optional, Callable, Any


def clean_filename(filename: str) -> str:
    """
    Sanitize tab name to create a valid, clean filesystem filename.
    Decodes URL encoding (e.g. '%20' -> ' ') and strips invalid characters.
    """
    if not filename:
        return "Untitled_Tab"

    # Decode web encoding first
    filename = unquote(filename)

    # Remove Windows/Unix illegal filename characters
    invalid_chars = '<>:"/\\|?*\x00\n\r\t'
    for char in invalid_chars:
        filename = filename.replace(char, '')

    # Strip leading/trailing dots and spaces
    filename = filename.strip('. ')
    return filename if filename else "Untitled_Tab"


class DrawIOExporter:
    """
    Handles parsing and exporting all pages/tabs from Draw.io (.drawio) files
    to images in the same folder as the original file or a custom destination.
    """

    DEFAULT_WINDOWS_PATHS = [
        r"C:\Program Files\draw.io\draw.io.exe",
        r"C:\Program Files (x86)\draw.io\draw.io.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\draw.io\draw.io.exe"),
        os.path.expandvars(r"%APPDATA%\draw.io\draw.io.exe"),
    ]

    DEFAULT_MAC_PATHS = [
        "/Applications/draw.io.app/Contents/MacOS/draw.io",
    ]

    DEFAULT_LINUX_PATHS = [
        "/usr/bin/drawio",
        "/usr/local/bin/drawio",
        "/snap/bin/drawio",
    ]

    @classmethod
    def find_drawio_executable(cls, custom_path: Optional[str] = None) -> Optional[str]:
        """
        Locate the draw.io desktop executable on the system.
        """
        if custom_path and os.path.isfile(custom_path):
            return os.path.abspath(custom_path)

        # Check system PATH
        path_binary = shutil.which("draw.io") or shutil.which("drawio")
        if path_binary:
            return os.path.abspath(path_binary)

        # Platform-specific lookup
        if sys.platform == "win32":
            for p in cls.DEFAULT_WINDOWS_PATHS:
                if os.path.isfile(p):
                    return os.path.abspath(p)
        elif sys.platform == "darwin":
            for p in cls.DEFAULT_MAC_PATHS:
                if os.path.isfile(p):
                    return os.path.abspath(p)
        else:
            for p in cls.DEFAULT_LINUX_PATHS:
                if os.path.isfile(p):
                    return os.path.abspath(p)

        return None

    @staticmethod
    def get_diagram_tabs(drawio_path: str) -> List[Dict[str, Any]]:
        """
        Parse the .drawio file and return metadata for all diagram tabs/pages.
        """
        if not os.path.isfile(drawio_path):
            raise FileNotFoundError(f"Draw.io file not found: {drawio_path}")

        try:
            tree = ET.parse(drawio_path)
            root = tree.getroot()
        except Exception as e:
            raise ValueError(f"Failed to read or parse Draw.io XML: {e}")

        # Diagrams can be directly under root <mxfile> or anywhere
        diagram_nodes = root.findall('.//diagram')
        
        # If no <diagram> tag exists, the root might be an <mxGraphModel> (single-page legacy)
        if not diagram_nodes:
            if root.tag == "mxGraphModel" or root.find('.//mxGraphModel') is not None:
                return [{
                    "index": 0,
                    "page_number": 1,
                    "id": "page-1",
                    "raw_name": "Page-1",
                    "clean_name": "Page-1"
                }]
            return []

        tabs = []
        seen_names: Dict[str, int] = {}

        for idx, diagram in enumerate(diagram_nodes):
            raw_name = diagram.get('name') or f"Page_{idx + 1}"
            cleaned = clean_filename(raw_name)

            # Avoid duplicate target filenames
            if cleaned in seen_names:
                seen_names[cleaned] += 1
                unique_name = f"{cleaned} ({seen_names[cleaned]})"
            else:
                seen_names[cleaned] = 0
                unique_name = cleaned

            tabs.append({
                "index": idx,
                "page_number": idx + 1,
                "id": diagram.get('id', f"page_{idx + 1}"),
                "raw_name": raw_name,
                "clean_name": unique_name
            })

        return tabs

    @classmethod
    def export_all_tabs(
        cls,
        drawio_path: str,
        output_dir: Optional[str] = None,
        format: str = "png",
        scale: float = 1.0,
        drawio_exe: Optional[str] = None,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> List[str]:
        """
        Export all tabs from a Draw.io file to image files.
        If output_dir is None, exports into the same folder containing drawio_path.

        :param drawio_path: Path to the source .drawio file.
        :param output_dir: Destination folder. Defaults to the folder of drawio_path.
        :param format: Export format (png, jpg, svg, pdf).
        :param scale: Export scale factor (e.g. 1.0, 2.0).
        :param drawio_exe: Optional custom path to draw.io executable.
        :param progress_callback: Optional callback(current_idx, total_count, tab_name).
        :return: List of exported file paths.
        """
        abs_drawio_path = os.path.abspath(drawio_path)
        if not os.path.isfile(abs_drawio_path):
            raise FileNotFoundError(f"Draw.io file not found: {abs_drawio_path}")

        # Default output directory: dedicated subfolder in same path as original drawio file
        if not output_dir:
            file_stem = os.path.splitext(os.path.basename(abs_drawio_path))[0]
            output_dir = os.path.join(os.path.dirname(abs_drawio_path), f"{file_stem}_exported")
        
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        tabs = cls.get_diagram_tabs(abs_drawio_path)
        if not tabs:
            raise ValueError(f"No pages/diagrams found in Draw.io file: {abs_drawio_path}")

        exe_path = cls.find_drawio_executable(drawio_exe)
        fmt = format.lower().strip('.')

        exported_files = []
        total_pages = len(tabs)

        if exe_path:
            # Export via Draw.io Desktop CLI
            for idx, tab in enumerate(tabs):
                page_num = tab["page_number"]
                clean_name = tab["clean_name"]
                out_file = os.path.join(output_dir, f"{clean_name}.{fmt}")

                if progress_callback:
                    progress_callback(idx + 1, total_pages, clean_name)

                command = [
                    exe_path,
                    "--export",
                    "--page-index", str(page_num),
                    "--format", fmt,
                    "--output", out_file
                ]
                if scale and scale != 1.0:
                    command.extend(["--scale", str(scale)])

                command.append(abs_drawio_path)

                result = subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                if result.returncode == 0 and os.path.isfile(out_file):
                    exported_files.append(out_file)
                else:
                    # Fallback to direct extraction if available
                    fallback_file = cls._extract_embedded_tab(abs_drawio_path, idx, out_file)
                    if fallback_file:
                        exported_files.append(fallback_file)
        else:
            # Fallback when draw.io executable is not installed: extract embedded Base64 images
            for idx, tab in enumerate(tabs):
                clean_name = tab["clean_name"]
                out_file = os.path.join(output_dir, f"{clean_name}.{fmt}")

                if progress_callback:
                    progress_callback(idx + 1, total_pages, clean_name)

                fallback_file = cls._extract_embedded_tab(abs_drawio_path, idx, out_file)
                if fallback_file:
                    exported_files.append(fallback_file)

            if not exported_files:
                raise RuntimeError(
                    "Draw.io desktop application (draw.io.exe) was not found. "
                    "Please install Draw.io Desktop from https://www.drawio.com/ or specify its path."
                )

        return exported_files

    @staticmethod
    def _extract_embedded_tab(drawio_path: str, page_index: int, target_file: str) -> Optional[str]:
        """
        Directly extract embedded Base64 image from the tab's XML if present.
        """
        try:
            tree = ET.parse(drawio_path)
            root = tree.getroot()
            diagrams = root.findall('.//diagram')
            if page_index >= len(diagrams):
                return None

            diagram = diagrams[page_index]
            for cell in diagram.findall('.//mxCell'):
                style = cell.get('style', '')
                if 'image=data:image/' in style:
                    # Extract base64 payload
                    part = style.split('image=data:image/', 1)[1]
                    if ';base64,' in part:
                        img_type, b64_data = part.split(';base64,', 1)
                        # Strip trailing style parameters
                        b64_data = b64_data.split(';')[0]
                        img_bytes = base64.b64decode(b64_data)
                        with open(target_file, 'wb') as f:
                            f.write(img_bytes)
                        return target_file
        except Exception:
            pass
        return None
