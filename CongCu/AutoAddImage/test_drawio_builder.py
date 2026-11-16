import os
import tempfile
import unittest
import xml.etree.ElementTree as ET
from PIL import Image
from drawio_builder import DrawIOBuilder


class TestDrawIOBuilder(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        
        # Create test images
        self.img1_path = os.path.join(self.temp_dir.name, "test_diagram_1.png")
        img1 = Image.new("RGB", (300, 150), color="red")
        img1.save(self.img1_path)

        self.img2_path = os.path.join(self.temp_dir.name, "test_diagram_2.jpg")
        img2 = Image.new("RGB", (500, 250), color="blue")
        img2.save(self.img2_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_create_new_drawio_file(self):
        output_file = os.path.join(self.temp_dir.name, "new_output.drawio")
        builder = DrawIOBuilder()

        tab1 = builder.add_image_tab(self.img1_path, tab_name="Red Diagram", embed=True, margin=20)
        tab2 = builder.add_image_tab(self.img2_path, tab_name="Blue Diagram", embed=True, margin=30)

        self.assertEqual(tab1, "Red Diagram")
        self.assertEqual(tab2, "Blue Diagram")

        saved_path = builder.save(output_file)
        self.assertTrue(os.path.isfile(saved_path))

        # Parse saved file
        tree = ET.parse(saved_path)
        root = tree.getroot()
        self.assertEqual(root.tag, "mxfile")

        diagrams = root.findall("diagram")
        self.assertEqual(len(diagrams), 2)
        self.assertEqual(diagrams[0].get("name"), "Red Diagram")
        self.assertEqual(diagrams[1].get("name"), "Blue Diagram")

        # Check page sizes and image cells
        graph1 = diagrams[0].find("mxGraphModel")
        self.assertEqual(graph1.get("pageWidth"), str(300 + 40))  # 300 + margin*2
        self.assertEqual(graph1.get("pageHeight"), str(150 + 40))

        # Check image cell style
        cells = graph1.findall(".//mxCell")
        image_cells = [c for c in cells if "shape=image" in c.get("style", "")]
        self.assertEqual(len(image_cells), 1)
        self.assertIn("data:image/png,", image_cells[0].get("style"))

    def test_append_to_existing_drawio_file(self):
        output_file = os.path.join(self.temp_dir.name, "existing.drawio")

        # Create first version
        builder1 = DrawIOBuilder()
        builder1.add_image_tab(self.img1_path, tab_name="Original Page", embed=True)
        builder1.save(output_file)

        # Load existing file and append new tab
        builder2 = DrawIOBuilder(output_file)
        existing_tabs = builder2.get_existing_tabs()
        self.assertEqual(existing_tabs, ["Original Page"])

        builder2.add_image_tab(self.img2_path, tab_name="Appended Page", embed=True)
        builder2.save()

        # Re-verify
        tree = ET.parse(output_file)
        root = tree.getroot()
        diagrams = root.findall("diagram")
        self.assertEqual(len(diagrams), 2)
        self.assertEqual(diagrams[0].get("name"), "Original Page")
        self.assertEqual(diagrams[1].get("name"), "Appended Page")

    def test_tab_name_collision(self):
        builder = DrawIOBuilder()
        name1 = builder.add_image_tab(self.img1_path, tab_name="Duplicate")
        name2 = builder.add_image_tab(self.img2_path, tab_name="Duplicate")

        self.assertEqual(name1, "Duplicate")
        self.assertEqual(name2, "Duplicate (1)")

    def test_linked_image_path(self):
        output_file = os.path.join(self.temp_dir.name, "linked.drawio")
        builder = DrawIOBuilder()
        builder.add_image_tab(self.img1_path, tab_name="Linked Page", embed=False)
        builder.save(output_file)

        tree = ET.parse(output_file)
        cell = tree.find(".//mxCell[@vertex='1']")
        style = cell.get("style")
        self.assertIn("image=file:///", style)


if __name__ == "__main__":
    unittest.main()
