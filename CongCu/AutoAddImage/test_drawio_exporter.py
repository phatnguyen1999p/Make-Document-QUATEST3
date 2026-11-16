import os
import tempfile
import unittest
from PIL import Image
from drawio_builder import DrawIOBuilder
from drawio_exporter import DrawIOExporter, clean_filename


class TestDrawIOExporter(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

        self.img1_path = os.path.join(self.temp_dir.name, "tab1.png")
        Image.new("RGB", (120, 80), color="purple").save(self.img1_path)

        self.img2_path = os.path.join(self.temp_dir.name, "tab2.jpg")
        Image.new("RGB", (140, 90), color="orange").save(self.img2_path)

        # Create a sample drawio file with special tab names
        self.drawio_path = os.path.join(self.temp_dir.name, "sample_diagram.drawio")
        builder = DrawIOBuilder()
        builder.add_image_tab(self.img1_path, tab_name="Overview%20Flow: Phase 1?")
        builder.add_image_tab(self.img2_path, tab_name="Data Model")
        builder.save(self.drawio_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_clean_filename(self):
        self.assertEqual(clean_filename("My%20Architecture%20Diagram"), "My Architecture Diagram")
        self.assertEqual(clean_filename("Diagram: Phase 1? <Draft>"), "Diagram Phase 1 Draft")
        self.assertEqual(clean_filename("   ...valid_name...  "), "valid_name")
        self.assertEqual(clean_filename(""), "Untitled_Tab")

    def test_get_diagram_tabs(self):
        tabs = DrawIOExporter.get_diagram_tabs(self.drawio_path)
        self.assertEqual(len(tabs), 2)
        self.assertEqual(tabs[0]["raw_name"], "Overview%20Flow: Phase 1?")
        self.assertEqual(tabs[0]["clean_name"], "Overview Flow Phase 1")
        self.assertEqual(tabs[1]["clean_name"], "Data Model")

    def test_export_all_tabs_same_directory(self):
        """Export without output_dir should save into a dedicated subfolder <stem>_exported"""
        exported = DrawIOExporter.export_all_tabs(self.drawio_path, format="png")
        self.assertEqual(len(exported), 2)

        expected_folder = os.path.join(self.temp_dir.name, "sample_diagram_exported")
        for path in exported:
            self.assertTrue(os.path.isfile(path))
            self.assertEqual(os.path.dirname(os.path.abspath(path)), os.path.abspath(expected_folder))

        expected_file1 = os.path.join(expected_folder, "Overview Flow Phase 1.png")
        expected_file2 = os.path.join(expected_folder, "Data Model.png")
        self.assertTrue(os.path.isfile(expected_file1))
        self.assertTrue(os.path.isfile(expected_file2))

    def test_export_all_tabs_custom_directory(self):
        """Export to custom output directory"""
        sub_dir = os.path.join(self.temp_dir.name, "exported_images")
        exported = DrawIOExporter.export_all_tabs(self.drawio_path, output_dir=sub_dir, format="png")
        self.assertEqual(len(exported), 2)

        for path in exported:
            self.assertTrue(os.path.isfile(path))
            self.assertEqual(os.path.dirname(os.path.abspath(path)), os.path.abspath(sub_dir))


if __name__ == "__main__":
    unittest.main()
