# Draw.io Image Tab Generator Tool

A powerful Python utility (GUI & CLI) to create Draw.io (`.drawio`) diagram files or append image tabs to existing `.drawio` files. Each image is placed on its own dedicated tab (page).

![Draw.io Compatibility](https://img.shields.io/badge/Draw.io-Compatible-brightgreen)
![Python 3.8+](https://img.shields.io/badge/Python-3.8+-blue)

---

## ✨ Features

- **Multi-Tab Generation**: Converts multiple images into separate tabs (pages) within a single `.drawio` file.
- **Batch Tab Exporter**: Automatically exports all pages/tabs from any `.drawio` file to separate image files in the same folder as the original `.drawio` file (or a custom destination).
- **Create New or Append**: Generate brand new `.drawio` files or append new tabs to an existing `.drawio` file.
- **Embedded Base64 Support**: By default, images are embedded directly into the `.drawio` XML as Base64 data URIs so that the file is completely portable and works offline/anywhere.
- **Linked Path Option**: Optional support for linking images via `file://` URIs to keep `.drawio` file sizes minimal.
- **Smart Auto-Fit Canvas**: Automatically calculates image dimensions (PNG, JPG, WEBP, GIF, BMP, SVG) and sets tab canvas bounds accordingly.
- **GUI & CLI Modes**: Includes a modern desktop GUI (Tkinter with multi-tab workspace) and a scriptable command-line interface.

---

## 🛠️ Installation

1. Clone or download this repository.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## 🖥️ Graphical User Interface (GUI)

Launch the interactive desktop interface:

```bash
python main.py
# or
python main.py --gui
```

### GUI Highlights:
- **Tab 1: Images to .drawio (Importer / Generator)**:
  - Add individual image files or entire folders.
  - Reorder tabs and edit custom tab names.
  - Automatically creates output `.drawio` file inside the image folder (or prompts on multipath).
  - Portability toggle (Base64 embed vs linked).
- **Tab 2: Export .drawio Tabs (Batch Exporter)**:
  - Select any `.drawio` file to instantly preview all tabs/pages.
  - Automatically sets the export destination to the same folder as the `.drawio` file.
  - Choose image format (PNG, JPG, SVG, PDF) and scale.
  - One-click batch export with progress tracking.

---

## 💻 Command-Line Interface (CLI)

### 1. Create a New `.drawio` File from Images
```bash
python main.py image1.png image2.jpg image3.webp -o output.drawio
```

### 2. Process an Entire Folder of Images
```bash
python main.py -d ./diagram_images -o my_architecture.drawio
```

### 3. Append Image Tabs to an Existing `.drawio` File
```bash
python main.py new_screen.png -a existing_workflow.drawio
```

### 4. Export All Tabs from a `.drawio` File to Images
```bash
# Exports all tabs to PNGs inside the same directory as diagrams.drawio
python main.py -e diagrams.drawio

# Export to a custom folder with custom format
python main.py -e diagrams.drawio -o ./exported_tabs --format png --scale 2.0
```

### 5. CLI Options Overview

| Flag | Description |
|---|---|
| `-e`, `--export` | Path to `.drawio` file to export all tabs from |
| `--format` | Export format: `png`, `jpg`, `svg`, `pdf` (default: `png`) |
| `--scale` | Export scale multiplier (default: `1.0`) |
| `-o`, `--output` | Output `.drawio` file path or export directory |
| `-a`, `--append` | Append tabs to an existing `.drawio` file |
| `-d`, `--dir` | Directory containing images |
| `--no-embed` | Link images by file path instead of embedding Base64 |
| `--margin` | Margin/padding around image in pixels (default: `20`) |
| `--no-auto-fit` | Use standard page size instead of fitting canvas to image |
| `--gui` | Force launch desktop GUI mode |

---

## 🧪 Running Tests

To run automated unit tests:

```bash
python -m unittest discover
```

---

## 📄 Supported Formats
- PNG (`.png`)
- JPEG (`.jpg`, `.jpeg`)
- WebP (`.webp`)
- GIF (`.gif`)
- BMP (`.bmp`)
- SVG (`.svg`)
