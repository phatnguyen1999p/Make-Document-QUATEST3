import os
import uuid
import base64
import mimetypes
import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import List, Optional, Tuple, Dict, Any
from PIL import Image


class DrawIOBuilder:
    """
    Builder for creating and updating Draw.io (.drawio) XML documents with image tabs.
    """

    MIME_TYPES = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".gif": "image/gif",
        ".bmp": "image/bmp",
        ".svg": "image/svg+xml",
    }

    def __init__(self, drawio_path: Optional[str] = None):
        """
        Initialize builder. If drawio_path is provided and exists, load existing file.
        Otherwise, initialize a new mxfile tree.
        """
        self.drawio_path = drawio_path
        self.existing_diagram_ids = set()
        self.existing_tab_names = set()

        if drawio_path and os.path.isfile(drawio_path):
            self.root, self.tree = self._load_existing_file(drawio_path)
        else:
            self.root = ET.Element("mxfile", {
                "host": "Electron",
                "modified": "2026-01-01T00:00:00.000Z",
                "agent": "DrawIO-ImageTabGenerator",
                "version": "21.0.0",
                "type": "device"
            })
            self.tree = ET.ElementTree(self.root)

        self._scan_existing_nodes()

    def _load_existing_file(self, file_path: str) -> Tuple[ET.Element, ET.ElementTree]:
        """Load and parse an existing Draw.io file."""
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            if root.tag != "mxfile":
                raise ValueError(f"Invalid Draw.io file structure: Root element must be <mxfile>, got <{root.tag}>")
            return root, tree
        except Exception as e:
            raise ValueError(f"Failed to parse Draw.io file at '{file_path}': {e}")

    def _scan_existing_nodes(self):
        """Record existing diagram IDs and tab names to avoid collisions."""
        for diagram in self.root.findall("diagram"):
            d_id = diagram.get("id")
            name = diagram.get("name")
            if d_id:
                self.existing_diagram_ids.add(d_id)
            if name:
                self.existing_tab_names.add(name)

    @staticmethod
    def _generate_unique_id(prefix: str = "") -> str:
        """Generate a random unique identifier for diagrams and cells."""
        uid = uuid.uuid4().hex[:12]
        return f"{prefix}{uid}" if prefix else uid

    def _get_image_metadata(self, image_path: str) -> Dict[str, Any]:
        """Extract image dimensions and MIME type."""
        if not os.path.isfile(image_path):
            raise FileNotFoundError(f"Image file not found: {image_path}")

        ext = os.path.splitext(image_path)[1].lower()
        mime_type = self.MIME_TYPES.get(ext)
        if not mime_type:
            mime_type = mimetypes.guess_type(image_path)[0] or "image/png"

        try:
            with Image.open(image_path) as img:
                width, height = img.size
        except Exception as e:
            raise ValueError(f"Could not read image file '{image_path}': {e}")

        return {
            "width": width,
            "height": height,
            "mime_type": mime_type,
            "ext": ext
        }

    def _encode_image_base64(self, image_path: str, mime_type: str) -> str:
        """Read image and return data URI string formatted for Draw.io."""
        with open(image_path, "rb") as f:
            encoded_bytes = base64.b64encode(f.read()).decode("utf-8")
        return f"data:{mime_type},{encoded_bytes}"

    def get_existing_tabs(self) -> List[str]:
        """Return list of existing tab names in the file."""
        return [diagram.get("name", "Untitled") for diagram in self.root.findall("diagram")]

    def add_image_tab(
        self,
        image_path: str,
        tab_name: Optional[str] = None,
        embed: bool = True,
        margin: int = 20,
        auto_fit: bool = True,
        custom_page_width: int = 827,
        custom_page_height: int = 1169
    ) -> str:
        """
        Add a new image tab (page) to the Draw.io file.

        :param image_path: Path to the image file.
        :param tab_name: Name of the tab/page. If None, derives from image filename.
        :param embed: If True, embed image as Base64. If False, link image path.
        :param margin: Margin (padding) in pixels around the image if auto_fit is True.
        :param auto_fit: If True, set tab canvas size equal to image size + margins.
        :param custom_page_width: Page width if auto_fit is False.
        :param custom_page_height: Page height if auto_fit is False.
        :return: Allocated tab name.
        """
        meta = self._get_image_metadata(image_path)
        img_w, img_h = meta["width"], meta["height"]

        # Determine tab name
        if not tab_name:
            tab_name = os.path.splitext(os.path.basename(image_path))[0]

        # Handle duplicate tab names
        original_name = tab_name
        counter = 1
        while tab_name in self.existing_tab_names:
            tab_name = f"{original_name} ({counter})"
            counter += 1
        self.existing_tab_names.add(tab_name)

        # Diagram ID
        diagram_id = self._generate_unique_id("diag_")
        while diagram_id in self.existing_diagram_ids:
            diagram_id = self._generate_unique_id("diag_")
        self.existing_diagram_ids.add(diagram_id)

        # Image source URL/URI
        if embed:
            image_src = self._encode_image_base64(image_path, meta["mime_type"])
        else:
            # Convert file path to absolute normalized path
            abs_path = os.path.abspath(image_path).replace("\\", "/")
            image_src = f"file:///{abs_path}"

        # Page geometry & placement
        if auto_fit:
            page_w = img_w + (margin * 2)
            page_h = img_h + (margin * 2)
            img_x = margin
            img_y = margin
        else:
            page_w = custom_page_width
            page_h = custom_page_height
            # Center image on standard page, scaling down if larger than page
            scale = min((page_w - margin * 2) / img_w, (page_h - margin * 2) / img_h, 1.0)
            target_w = int(img_w * scale)
            target_h = int(img_h * scale)
            img_x = int((page_w - target_w) / 2)
            img_y = int((page_h - target_h) / 2)
            img_w, img_h = target_w, target_h

        # Create XML nodes
        diagram_elem = ET.SubElement(self.root, "diagram", {
            "id": diagram_id,
            "name": tab_name
        })

        graph_model = ET.SubElement(diagram_elem, "mxGraphModel", {
            "dx": "1000",
            "dy": "1000",
            "grid": "1",
            "gridSize": "10",
            "guides": "1",
            "tooltips": "1",
            "connect": "1",
            "arrows": "1",
            "fold": "1",
            "page": "1",
            "pageScale": "1",
            "pageWidth": str(page_w),
            "pageHeight": str(page_h),
            "math": "0",
            "shadow": "0"
        })

        root_elem = ET.SubElement(graph_model, "root")

        # Base mxCell elements (required by draw.io)
        ET.SubElement(root_elem, "mxCell", {"id": "0"})
        ET.SubElement(root_elem, "mxCell", {"id": "1", "parent": "0"})

        # Image Cell
        cell_id = self._generate_unique_id("cell_")
        style_str = f"shape=image;verticalLabelPosition=bottom;verticalAlign=top;imageAspect=0;aspect=fixed;image={image_src};"

        image_cell = ET.SubElement(root_elem, "mxCell", {
            "id": cell_id,
            "value": "",
            "style": style_str,
            "vertex": "1",
            "parent": "1"
        })

        ET.SubElement(image_cell, "mxGeometry", {
            "x": str(img_x),
            "y": str(img_y),
            "width": str(img_w),
            "height": str(img_h),
            "as": "geometry"
        })

        return tab_name

    def save(self, output_path: Optional[str] = None) -> str:
        """
        Save the XML structure to file. If output_path is None, saves to self.drawio_path.
        """
        target_path = output_path or self.drawio_path
        if not target_path:
            raise ValueError("No output path specified for saving Draw.io file.")

        # Ensure target directory exists
        os.makedirs(os.path.dirname(os.path.abspath(target_path)), exist_ok=True)

        # Convert to pretty XML string
        rough_string = ET.tostring(self.root, encoding="utf-8")
        reparsed = minidom.parseString(rough_string)
        pretty_xml = reparsed.toprettyxml(indent="  ", encoding="utf-8")

        # Strip extra blank lines that minidom sometimes adds
        lines = [line for line in pretty_xml.decode("utf-8").splitlines() if line.strip()]
        clean_xml = "\n".join(lines)

        with open(target_path, "w", encoding="utf-8") as f:
            f.write(clean_xml)

        return target_path
