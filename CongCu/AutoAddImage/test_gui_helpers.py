import os
import tempfile
import unittest
from unittest.mock import patch
from PIL import Image
import tkinter as tk

from gui import DrawIOGUI
from cli import run_cli
from drawio_builder import DrawIOBuilder


class TestGUIAutoPath(unittest.TestCase):

    def setUp(self):
        self.temp_dir1 = tempfile.TemporaryDirectory()
        self.temp_dir2 = tempfile.TemporaryDirectory()

        self.img1_path = os.path.join(self.temp_dir1.name, "diagram1.png")
        Image.new("RGB", (100, 100), color="red").save(self.img1_path)

        self.img2_path = os.path.join(self.temp_dir1.name, "diagram2.jpg")
        Image.new("RGB", (100, 100), color="blue").save(self.img2_path)

        self.img3_path = os.path.join(self.temp_dir2.name, "other_diagram.png")
        Image.new("RGB", (100, 100), color="green").save(self.img3_path)

        # Create a sample drawio file for export tests
        self.sample_drawio = os.path.join(self.temp_dir1.name, "multi_tab.drawio")
        builder = DrawIOBuilder()
        builder.add_image_tab(self.img1_path, tab_name="Main Arch")
        builder.add_image_tab(self.img2_path, tab_name="Sub Flow")
        builder.save(self.sample_drawio)

        # Initialize GUI instance with hidden window
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = DrawIOGUI()
        self.app.withdraw()

    def tearDown(self):
        self.app.destroy()
        self.root.destroy()
        self.temp_dir1.cleanup()
        self.temp_dir2.cleanup()

    def test_single_image_auto_path(self):
        """When 1 image is added, output path should be <folder>/<image_name>.drawio"""
        self.app.mode_var.set("new")
        self.app._append_image_file(self.img1_path)
        self.app._handle_auto_output_path(prompt_if_multipath=False)

        expected_path = os.path.join(self.temp_dir1.name, "diagram1.drawio")
        self.assertEqual(os.path.abspath(self.app.txt_path.get()), os.path.abspath(expected_path))
        self.assertTrue(self.app._auto_path_set)

    def test_multiple_images_single_folder_auto_path(self):
        """When multiple images from same folder are added, path is <folder>/<folder_name>.drawio"""
        self.app.mode_var.set("new")
        self.app._append_image_file(self.img1_path)
        self.app._append_image_file(self.img2_path)
        self.app._handle_auto_output_path(prompt_if_multipath=False)

        folder_name = os.path.basename(self.temp_dir1.name)
        expected_path = os.path.join(self.temp_dir1.name, f"{folder_name}.drawio")
        self.assertEqual(os.path.abspath(self.app.txt_path.get()), os.path.abspath(expected_path))
        self.assertTrue(self.app._auto_path_set)

    def test_multipath_prompts_user(self):
        """When images from multiple folders exist, prompt user for output file path"""
        self.app.mode_var.set("new")
        self.app._append_image_file(self.img1_path)
        self.app._append_image_file(self.img3_path)

        chosen_path = os.path.join(self.temp_dir1.name, "custom_choice.drawio")
        with patch("gui.filedialog.asksaveasfilename", return_value=chosen_path) as mock_dialog:
            self.app._handle_auto_output_path(prompt_if_multipath=True)
            mock_dialog.assert_called_once()
            self.assertEqual(os.path.abspath(self.app.txt_path.get()), os.path.abspath(chosen_path))
            self.assertFalse(self.app._auto_path_set)

    def test_mode_change_to_new_auto_sets_path(self):
        """When switching mode to 'new', auto-set path or ask on multipath"""
        self.app._append_image_file(self.img1_path)
        self.app.mode_var.set("append")
        self.app._on_mode_change()
        self.assertEqual(self.app.lbl_path.cget("text"), "Existing File:")

        self.app.mode_var.set("new")
        self.app._on_mode_change()
        self.assertEqual(self.app.lbl_path.cget("text"), "Output File:")
        expected_path = os.path.join(self.temp_dir1.name, "diagram1.drawio")
        self.assertEqual(os.path.abspath(self.app.txt_path.get()), os.path.abspath(expected_path))

    def test_gui_exporter_auto_destination_folder(self):
        """Selecting a drawio file in the Exporter tab auto-populates destination to dedicated subfolder"""
        self.app.set_export_source_file(self.sample_drawio)
        expected_export_dir = os.path.join(self.temp_dir1.name, "multi_tab_exported")
        self.assertEqual(
            os.path.abspath(self.app.txt_export_dst.get()),
            os.path.abspath(expected_export_dir)
        )
        self.assertEqual(len(self.app.export_tab_items), 2)
        self.assertEqual(self.app.export_tab_items[0]["clean_name"], "Main Arch")
        self.assertEqual(self.app.export_tab_items[1]["clean_name"], "Sub Flow")

    def test_cli_export_command(self):
        """CLI with -e should export all tabs into the drawio file's dedicated subfolder"""
        exit_code = run_cli(["-e", self.sample_drawio, "--format", "png"])
        self.assertEqual(exit_code, 0)

        export_dir = os.path.join(self.temp_dir1.name, "multi_tab_exported")
        out1 = os.path.join(export_dir, "Main Arch.png")
        out2 = os.path.join(export_dir, "Sub Flow.png")
        self.assertTrue(os.path.isfile(out1))
        self.assertTrue(os.path.isfile(out2))

    def test_cli_auto_path_single_directory(self):
        """CLI should auto-detect output path when -o is omitted for single directory images"""
        exit_code = run_cli([self.img1_path, self.img2_path])
        self.assertEqual(exit_code, 0)

        folder_name = os.path.basename(self.temp_dir1.name)
        expected_file = os.path.join(self.temp_dir1.name, f"{folder_name}.drawio")
        self.assertTrue(os.path.isfile(expected_file))

    def test_cli_error_multipath_without_output(self):
        """CLI should return error when images come from multiple folders and -o is not provided"""
        exit_code = run_cli([self.img1_path, self.img3_path])
        self.assertEqual(exit_code, 1)


if __name__ == "__main__":
    unittest.main()
