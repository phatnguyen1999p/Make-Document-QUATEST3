import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any
from PIL import Image

from drawio_builder import DrawIOBuilder
from drawio_exporter import DrawIOExporter, clean_filename


SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".svg"}


class DrawIOGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Draw.io Diagram Suite - Import & Export Tabs")
        
        # --- 1. Dynamic Half-Screen Window Geometry ---
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        # Compute half-screen width and comfortable height
        win_w = max(min(int(screen_w * 0.48), 920), 620)
        win_h = max(min(int(screen_h * 0.88), 980), 600)
        pos_x = 10
        pos_y = 10

        self.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.minsize(580, 480)

        # Style configuration
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        
        # Color definitions
        self.bg_color = "#f4f6f9"
        self.accent_color = "#0d6efd"
        self.accent_hover = "#0b5ed7"

        self.configure(bg=self.bg_color)
        self._set_styles()

        # State data for Tab 1 (Builder / Generator)
        self.image_items: List[Dict[str, Any]] = []
        self._auto_path_set: bool = False
        self._is_building: bool = False

        # State data for Tab 2 (Exporter)
        self.export_tab_items: List[Dict[str, Any]] = []
        self._export_auto_dir_set: bool = False
        self._is_exporting: bool = False

        # Build UI layout
        self._create_widgets()

    def _set_styles(self):
        self.style.configure(".", background=self.bg_color, font=("Segoe UI", 9))
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure("TLabelframe", background=self.bg_color, font=("Segoe UI", 10, "bold"))
        self.style.configure("TLabelframe.Label", background=self.bg_color, foreground="#212529")
        
        self.style.configure("TNotebook", background=self.bg_color)
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=[12, 5])
        
        self.style.configure("TButton", font=("Segoe UI", 9), padding=4)
        self.style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), foreground="#ffffff", background=self.accent_color)
        self.style.map("Primary.TButton", background=[("active", self.accent_hover), ("pressed", "#0a58ca")])
        
        self.style.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self.style.configure("TProgressbar", thickness=10, background=self.accent_color)

    def _create_widgets(self):
        # Header banner (optimized height for half-screen)
        header_frame = tk.Frame(self, bg="#2b3035", height=50)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)

        header_label = tk.Label(
            header_frame,
            text="Draw.io Diagram Suite",
            font=("Segoe UI", 13, "bold"),
            fg="#ffffff",
            bg="#2b3035",
            anchor="w",
            padx=12
        )
        header_label.pack(side=tk.LEFT, fill=tk.Y)

        sub_label = tk.Label(
            header_frame,
            text="Image Tab Generator & Batch Exporter",
            font=("Segoe UI", 8),
            fg="#adb5bd",
            bg="#2b3035",
            anchor="e",
            padx=12
        )
        sub_label.pack(side=tk.RIGHT, fill=tk.Y)

        # Notebook Container for Tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(6, 8))

        # --- Tab 1: Images to .drawio Generator ---
        self.tab_generator = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_generator, text="  🖼️ Images to .drawio  ")
        self._build_generator_tab(self.tab_generator)

        # --- Tab 2: Export .drawio Tabs ---
        self.tab_exporter = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.tab_exporter, text="  📤 Export .drawio Tabs  ")
        self._build_exporter_tab(self.tab_exporter)

    # =========================================================================
    # TAB 1: GENERATOR (Images -> Draw.io)
    # =========================================================================
    def _build_generator_tab(self, parent):
        # --- Top Section: File Actions & Treeview ---
        items_frame = ttk.LabelFrame(parent, text=" Images Queue ", padding=6)
        items_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # Toolbar
        toolbar = ttk.Frame(items_frame)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        self.btn_add_files = ttk.Button(toolbar, text="+ Add Images", command=self._add_files)
        self.btn_add_files.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_add_folder = ttk.Button(toolbar, text="+ Add Folder", command=self._add_folder)
        self.btn_add_folder.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_clear = ttk.Button(toolbar, text="Clear All", command=self._clear_items)
        self.btn_clear.pack(side=tk.LEFT, padx=(0, 4))

        self.lbl_queue_count = ttk.Label(toolbar, text="0 image(s)", foreground="#6c757d")
        self.lbl_queue_count.pack(side=tk.RIGHT, padx=4)

        # Treeview list & Scrollbar
        tree_container = ttk.Frame(items_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        columns = ("name", "dimensions", "tab_name", "path")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("name", text="File Name")
        self.tree.heading("dimensions", text="Size")
        self.tree.heading("tab_name", text="Tab Name")
        self.tree.heading("path", text="Path")

        # Responsive column widths optimized for half-screen
        self.tree.column("name", width=140, minwidth=90, stretch=True)
        self.tree.column("dimensions", width=75, minwidth=60, stretch=False, anchor="center")
        self.tree.column("tab_name", width=130, minwidth=80, stretch=True)
        self.tree.column("path", width=180, minwidth=100, stretch=True)

        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Side controls for list item reordering/editing
        side_controls = ttk.Frame(items_frame)
        side_controls.pack(fill=tk.Y, side=tk.RIGHT, padx=(6, 0))

        self.btn_up = ttk.Button(side_controls, text="▲ Up", width=7, command=self._move_up)
        self.btn_up.pack(fill=tk.X, pady=(0, 3))

        self.btn_down = ttk.Button(side_controls, text="▼ Down", width=7, command=self._move_down)
        self.btn_down.pack(fill=tk.X, pady=(0, 3))

        self.btn_rename = ttk.Button(side_controls, text="Rename", width=7, command=self._edit_tab_name)
        self.btn_rename.pack(fill=tk.X, pady=(0, 3))

        self.btn_remove = ttk.Button(side_controls, text="Remove", width=7, command=self._remove_selected)
        self.btn_remove.pack(fill=tk.X, pady=(0, 3))

        self.tree.bind("<Double-1>", lambda event: self._edit_tab_name())

        # --- Middle Section: Output Configuration ---
        config_frame = ttk.LabelFrame(parent, text=" Output Settings ", padding=8)
        config_frame.pack(fill=tk.X, pady=(0, 6))

        # Mode Selection
        self.mode_var = tk.StringVar(value="new")
        mode_frame = ttk.Frame(config_frame)
        mode_frame.pack(fill=tk.X, pady=(0, 4))

        self.rb_new = ttk.Radiobutton(
            mode_frame, text="Create New .drawio", value="new",
            variable=self.mode_var, command=self._on_mode_change
        )
        self.rb_new.pack(side=tk.LEFT, padx=(0, 12))

        self.rb_append = ttk.Radiobutton(
            mode_frame, text="Append to Existing .drawio", value="append",
            variable=self.mode_var, command=self._on_mode_change
        )
        self.rb_append.pack(side=tk.LEFT)

        # Target Path Picker
        path_frame = ttk.Frame(config_frame)
        path_frame.pack(fill=tk.X, pady=(0, 6))

        self.lbl_path = ttk.Label(path_frame, text="Output File:", width=11)
        self.lbl_path.pack(side=tk.LEFT)

        self.txt_path = ttk.Entry(path_frame)
        self.txt_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.btn_browse_out = ttk.Button(path_frame, text="Browse...", command=self._browse_output)
        self.btn_browse_out.pack(side=tk.RIGHT)

        # Advanced Toggles
        toggles_frame = ttk.Frame(config_frame)
        toggles_frame.pack(fill=tk.X)

        self.var_embed = tk.BooleanVar(value=True)
        chk_embed = ttk.Checkbutton(
            toggles_frame, text="Embed Base64", variable=self.var_embed
        )
        chk_embed.pack(side=tk.LEFT, padx=(0, 15))

        self.var_autofit = tk.BooleanVar(value=True)
        chk_autofit = ttk.Checkbutton(
            toggles_frame, text="Auto-fit Canvas", variable=self.var_autofit
        )
        chk_autofit.pack(side=tk.LEFT, padx=(0, 15))

        lbl_margin = ttk.Label(toggles_frame, text="Margin (px):")
        lbl_margin.pack(side=tk.LEFT, padx=(0, 4))

        self.spn_margin = ttk.Spinbox(toggles_frame, from_=0, to=200, width=4)
        self.spn_margin.set("20")
        self.spn_margin.pack(side=tk.LEFT)

        # --- Bottom Section: Progress Bar & Build Button ---
        bottom_frame = ttk.Frame(parent)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(2, 0))

        # Progress bar for adding & building
        self.prog_generator = ttk.Progressbar(bottom_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.prog_generator.pack(fill=tk.X, pady=(0, 4))

        status_action_frame = ttk.Frame(bottom_frame)
        status_action_frame.pack(fill=tk.X)

        self.lbl_status = ttk.Label(status_action_frame, text="Ready. Add images to get started.", foreground="#6c757d")
        self.lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_build = ttk.Button(
            status_action_frame,
            text="🚀 Generate Draw.io File",
            style="Primary.TButton",
            command=self._build_drawio
        )
        self.btn_build.pack(side=tk.RIGHT)

    def _get_image_directories(self) -> List[str]:
        """Return unique list of directory paths containing queued images."""
        seen = set()
        dirs = []
        for item in self.image_items:
            folder = os.path.dirname(os.path.abspath(item["path"]))
            if folder not in seen:
                seen.add(folder)
                dirs.append(folder)
        return dirs

    def _get_default_new_filename(self, folder: str) -> str:
        """Calculate default .drawio filename for single or multiple images."""
        if len(self.image_items) == 1:
            base_name = self.image_items[0]["tab_name"]
        else:
            base_name = os.path.basename(folder) or "diagrams"
        return f"{base_name}.drawio"

    def _handle_auto_output_path(self, prompt_if_multipath: bool = True):
        """
        Auto-set output path to image directory if single folder, or prompt user if multiple folders.
        """
        if self.mode_var.get() != "new":
            return
        if not self.image_items:
            return

        dirs = self._get_image_directories()
        if len(dirs) == 1:
            folder = dirs[0]
            filename = self._get_default_new_filename(folder)
            auto_path = os.path.join(folder, filename)
            self.txt_path.delete(0, tk.END)
            self.txt_path.insert(0, auto_path)
            self._auto_path_set = True
            self.lbl_status.config(text=f"Auto output: {auto_path}", foreground="#212529")
        elif len(dirs) > 1:
            if prompt_if_multipath:
                initial_dir = dirs[0] if dirs else None
                selected = filedialog.asksaveasfilename(
                    title="Select Output Draw.io File (Images from multiple folders)",
                    defaultextension=".drawio",
                    initialdir=initial_dir,
                    initialfile="output.drawio",
                    filetypes=[("Draw.io Diagrams", "*.drawio"), ("XML Files", "*.xml"), ("All Files", "*.*")]
                )
                if selected:
                    self.txt_path.delete(0, tk.END)
                    self.txt_path.insert(0, selected)
                    self._auto_path_set = False
                    self.lbl_status.config(text=f"Output file: {os.path.basename(selected)}", foreground="#212529")

    def _on_mode_change(self):
        mode = self.mode_var.get()
        if mode == "new":
            self.lbl_path.config(text="Output File:")
            self._handle_auto_output_path(prompt_if_multipath=True)
        else:
            self.lbl_path.config(text="Existing File:")
            if self._auto_path_set:
                self.txt_path.delete(0, tk.END)
                self._auto_path_set = False

    def _browse_output(self):
        mode = self.mode_var.get()
        dirs = self._get_image_directories()
        initial_dir = dirs[0] if dirs else None

        if mode == "new":
            initial_file = self._get_default_new_filename(initial_dir) if initial_dir else "output.drawio"
            filename = filedialog.asksaveasfilename(
                title="Save Draw.io File",
                defaultextension=".drawio",
                initialdir=initial_dir,
                initialfile=initial_file,
                filetypes=[("Draw.io Diagrams", "*.drawio"), ("XML Files", "*.xml"), ("All Files", "*.*")]
            )
        else:
            filename = filedialog.askopenfilename(
                title="Select Existing Draw.io File",
                initialdir=initial_dir,
                filetypes=[("Draw.io Diagrams", "*.drawio"), ("XML Files", "*.xml"), ("All Files", "*.*")]
            )

        if filename:
            self.txt_path.delete(0, tk.END)
            self.txt_path.insert(0, filename)
            self._auto_path_set = False

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Select Image Files",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.webp;*.gif;*.bmp;*.svg"), ("All Files", "*.*")]
        )
        if files:
            self._process_adding_files(list(files))

    def _add_folder(self):
        folder = filedialog.askdirectory(title="Select Folder with Images")
        if folder:
            found_files = []
            for root, _, filenames in os.walk(folder):
                for filename in filenames:
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in SUPPORTED_EXTENSIONS:
                        found_files.append(os.path.join(root, filename))
            if found_files:
                self._process_adding_files(found_files)
            else:
                messagebox.showinfo("No Images Found", f"No supported images found in:\n{folder}")

    def _process_adding_files(self, file_paths: List[str]):
        """Load image files with live progress bar feedback."""
        total = len(file_paths)
        if total == 0:
            return

        self.prog_generator["value"] = 0
        self.prog_generator["maximum"] = total
        self.lbl_status.config(text=f"Loading {total} image(s)...", foreground=self.accent_color)
        self.update_idletasks()

        for idx, f in enumerate(file_paths, 1):
            self._append_image_file(f)
            if total > 5 and idx % 2 == 0:
                self.prog_generator["value"] = idx
                self.update_idletasks()

        self.prog_generator["value"] = total
        self._refresh_tree()

        if self.mode_var.get() == "new":
            current_path = self.txt_path.get().strip()
            if not current_path or self._auto_path_set:
                self._handle_auto_output_path(prompt_if_multipath=True)

        self.after(500, lambda: self.prog_generator.configure(value=0))

    def _append_image_file(self, file_path: str):
        norm_path = os.path.abspath(file_path)
        if any(item["path"] == norm_path for item in self.image_items):
            return

        file_name = os.path.basename(file_path)
        default_tab_name = os.path.splitext(file_name)[0]

        dimensions_str = "Unknown"
        try:
            with Image.open(file_path) as img:
                dimensions_str = f"{img.width}x{img.height}"
        except Exception:
            pass

        self.image_items.append({
            "name": file_name,
            "path": norm_path,
            "dimensions": dimensions_str,
            "tab_name": default_tab_name
        })

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, item in enumerate(self.image_items):
            self.tree.insert("", tk.END, iid=str(idx), values=(
                item["name"],
                item["dimensions"],
                item["tab_name"],
                item["path"]
            ))

        count = len(self.image_items)
        self.lbl_queue_count.config(text=f"{count} image(s)")
        self.lbl_status.config(text=f"{count} image(s) in queue.", foreground="#212529")

    def _clear_items(self):
        self.image_items.clear()
        self._refresh_tree()
        self.prog_generator["value"] = 0
        if self._auto_path_set:
            self.txt_path.delete(0, tk.END)
            self._auto_path_set = False

    def _get_selected_index(self) -> int:
        selected = self.tree.selection()
        if not selected:
            return -1
        return int(selected[0])

    def _remove_selected(self):
        idx = self._get_selected_index()
        if idx >= 0:
            del self.image_items[idx]
            self._refresh_tree()
            if not self.image_items:
                if self._auto_path_set:
                    self.txt_path.delete(0, tk.END)
                    self._auto_path_set = False
            elif self._auto_path_set and self.mode_var.get() == "new":
                self._handle_auto_output_path(prompt_if_multipath=False)

    def _move_up(self):
        idx = self._get_selected_index()
        if idx > 0:
            self.image_items[idx], self.image_items[idx - 1] = self.image_items[idx - 1], self.image_items[idx]
            self._refresh_tree()
            self.tree.selection_set(str(idx - 1))

    def _move_down(self):
        idx = self._get_selected_index()
        if 0 <= idx < len(self.image_items) - 1:
            self.image_items[idx], self.image_items[idx + 1] = self.image_items[idx + 1], self.image_items[idx]
            self._refresh_tree()
            self.tree.selection_set(str(idx + 1))

    def _edit_tab_name(self):
        idx = self._get_selected_index()
        if idx < 0:
            return

        current_item = self.image_items[idx]
        
        dialog = tk.Toplevel(self)
        dialog.title("Edit Tab Name")
        dialog.geometry("380x130")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        lbl = ttk.Label(dialog, text=f"Enter tab name for '{current_item['name']}':")
        lbl.pack(anchor="w", padx=12, pady=(12, 4))

        entry = ttk.Entry(dialog)
        entry.pack(fill=tk.X, padx=12, pady=(0, 12))
        entry.insert(0, current_item["tab_name"])
        entry.select_range(0, tk.END)
        entry.focus_set()

        def save_name():
            new_name = entry.get().strip()
            if new_name:
                self.image_items[idx]["tab_name"] = new_name
                self._refresh_tree()
                if self._auto_path_set and len(self.image_items) == 1 and self.mode_var.get() == "new":
                    self._handle_auto_output_path(prompt_if_multipath=False)
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=12)

        btn_save = ttk.Button(btn_frame, text="Save", command=save_name)
        btn_save.pack(side=tk.RIGHT, padx=(4, 0))

        btn_cancel = ttk.Button(btn_frame, text="Cancel", command=dialog.destroy)
        btn_cancel.pack(side=tk.RIGHT)

        dialog.bind("<Return>", lambda e: save_name())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    def _build_drawio(self):
        if self._is_building:
            return

        if not self.image_items:
            messagebox.showwarning("No Images", "Please add at least one image to the queue.")
            return

        mode = self.mode_var.get()
        target_file = self.txt_path.get().strip()

        if not target_file and mode == "new":
            self._handle_auto_output_path(prompt_if_multipath=True)
            target_file = self.txt_path.get().strip()

        if not target_file:
            messagebox.showwarning("Target File Missing", "Please select an output file path.")
            return

        if not target_file.lower().endswith((".drawio", ".xml")):
            target_file += ".drawio"
            self.txt_path.delete(0, tk.END)
            self.txt_path.insert(0, target_file)

        is_append = (mode == "append")
        if is_append and not os.path.isfile(target_file):
            messagebox.showerror("File Not Found", f"Cannot append. Existing file not found:\n{target_file}")
            return

        try:
            margin = int(self.spn_margin.get())
        except ValueError:
            margin = 20

        embed = self.var_embed.get()
        autofit = self.var_autofit.get()
        total_items = len(self.image_items)

        self._is_building = True
        self.btn_build.config(state=tk.DISABLED)
        self.prog_generator["value"] = 0
        self.prog_generator["maximum"] = total_items
        self.lbl_status.config(text="Building Draw.io file...", foreground=self.accent_color)
        self.update_idletasks()

        def run_build_thread():
            try:
                builder = DrawIOBuilder(drawio_path=target_file if is_append else None)

                for idx, item in enumerate(self.image_items, 1):
                    self.lbl_status.config(
                        text=f"Adding tab {idx} of {total_items}: '{item['tab_name']}'...",
                        foreground=self.accent_color
                    )
                    self.prog_generator["value"] = idx
                    self.update_idletasks()

                    builder.add_image_tab(
                        image_path=item["path"],
                        tab_name=item["tab_name"],
                        embed=embed,
                        margin=margin,
                        auto_fit=autofit
                    )

                self.lbl_status.config(text="Saving file to disk...", foreground=self.accent_color)
                self.update_idletasks()
                saved_path = builder.save(target_file)

                self.prog_generator["value"] = total_items
                self.lbl_status.config(text=f"Done! Saved to {os.path.basename(saved_path)}", foreground="#198754")

                resp = messagebox.askyesno(
                    "Success!",
                    f"Draw.io file created successfully with {total_items} tab(s):\n\n"
                    f"{saved_path}\n\n"
                    f"Would you like to open the output folder?"
                )
                if resp:
                    folder = os.path.dirname(os.path.abspath(saved_path))
                    if sys.platform == "win32":
                        os.startfile(folder)
                    elif sys.platform == "darwin":
                        subprocess.run(["open", folder])
                    else:
                        subprocess.run(["xdg-open", folder])

            except Exception as e:
                self.lbl_status.config(text="Error occurred during build.", foreground="#dc3545")
                messagebox.showerror("Build Error", f"An error occurred while generating the file:\n{e}")
            finally:
                self._is_building = False
                self.btn_build.config(state=tk.NORMAL)
                self.after(800, lambda: self.prog_generator.configure(value=0))

        threading.Thread(target=run_build_thread, daemon=True).start()

    # =========================================================================
    # TAB 2: EXPORTER (Draw.io -> Images)
    # =========================================================================
    def _build_exporter_tab(self, parent):
        # Top Container: Source & Destination Setup
        setup_frame = ttk.LabelFrame(parent, text=" Source & Destination ", padding=8)
        setup_frame.pack(fill=tk.X, pady=(0, 6))

        # 1. Source Draw.io File Picker
        src_frame = ttk.Frame(setup_frame)
        src_frame.pack(fill=tk.X, pady=(0, 4))

        lbl_src = ttk.Label(src_frame, text="Draw.io File:", width=12)
        lbl_src.pack(side=tk.LEFT)

        self.txt_export_src = ttk.Entry(src_frame)
        self.txt_export_src.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        btn_browse_src = ttk.Button(src_frame, text="Select File...", command=self._browse_export_source)
        btn_browse_src.pack(side=tk.RIGHT)

        # 2. Output Destination Folder Picker (Created folder)
        dst_frame = ttk.Frame(setup_frame)
        dst_frame.pack(fill=tk.X, pady=(0, 6))

        lbl_dst = ttk.Label(dst_frame, text="Export Folder:", width=12)
        lbl_dst.pack(side=tk.LEFT)

        self.txt_export_dst = ttk.Entry(dst_frame)
        self.txt_export_dst.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        btn_browse_dst = ttk.Button(dst_frame, text="Browse...", command=self._browse_export_dest)
        btn_browse_dst.pack(side=tk.RIGHT)

        # 3. Format & Scale Options
        opts_frame = ttk.Frame(setup_frame)
        opts_frame.pack(fill=tk.X)

        lbl_fmt = ttk.Label(opts_frame, text="Format:")
        lbl_fmt.pack(side=tk.LEFT, padx=(0, 4))

        self.cbo_export_fmt = ttk.Combobox(opts_frame, values=["png", "jpg", "svg", "pdf"], width=5, state="readonly")
        self.cbo_export_fmt.set("png")
        self.cbo_export_fmt.pack(side=tk.LEFT, padx=(0, 15))
        self.cbo_export_fmt.bind("<<ComboboxSelected>>", lambda e: self._refresh_export_table_filenames())

        lbl_scale = ttk.Label(opts_frame, text="Scale:")
        lbl_scale.pack(side=tk.LEFT, padx=(0, 4))

        self.spn_export_scale = ttk.Spinbox(opts_frame, from_=0.5, to=5.0, increment=0.5, width=4)
        self.spn_export_scale.set("1.0")
        self.spn_export_scale.pack(side=tk.LEFT, padx=(0, 15))

        # Check for draw.io app status badge
        drawio_binary = DrawIOExporter.find_drawio_executable()
        if drawio_binary:
            status_text = "✓ Draw.io Desktop Ready"
            status_fg = "#198754"
        else:
            status_text = "⚠ App not found (Embedded mode)"
            status_fg = "#fd7e14"

        lbl_engine = tk.Label(opts_frame, text=status_text, fg=status_fg, bg=self.bg_color, font=("Segoe UI", 8, "bold"))
        lbl_engine.pack(side=tk.RIGHT)

        # Middle Container: Tabs Preview Table
        tabs_frame = ttk.LabelFrame(parent, text=" Detected Tabs in Draw.io File ", padding=6)
        tabs_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        tabs_toolbar = ttk.Frame(tabs_frame)
        tabs_toolbar.pack(fill=tk.X, pady=(0, 4))

        btn_reload_tabs = ttk.Button(tabs_toolbar, text="🔄 Reload Tabs", command=self._reload_export_tabs)
        btn_reload_tabs.pack(side=tk.LEFT, padx=(0, 4))

        self.lbl_tabs_summary = ttk.Label(tabs_toolbar, text="No file selected.", foreground="#6c757d")
        self.lbl_tabs_summary.pack(side=tk.LEFT, padx=(6, 0))

        # Treeview for tabs optimized for half-screen
        export_tree_container = ttk.Frame(tabs_frame)
        export_tree_container.pack(fill=tk.BOTH, expand=True)

        exp_columns = ("page_num", "tab_name", "target_filename", "status")
        self.export_tree = ttk.Treeview(export_tree_container, columns=exp_columns, show="headings", selectmode="browse")

        self.export_tree.heading("page_num", text="Page #")
        self.export_tree.heading("tab_name", text="Original Tab Name")
        self.export_tree.heading("target_filename", text="Output Image File")
        self.export_tree.heading("status", text="Status")

        self.export_tree.column("page_num", width=55, minwidth=40, stretch=False, anchor="center")
        self.export_tree.column("tab_name", width=150, minwidth=90, stretch=True)
        self.export_tree.column("target_filename", width=170, minwidth=100, stretch=True)
        self.export_tree.column("status", width=90, minwidth=70, stretch=False, anchor="center")

        exp_scrollbar = ttk.Scrollbar(export_tree_container, orient=tk.VERTICAL, command=self.export_tree.yview)
        self.export_tree.configure(yscroll=exp_scrollbar.set)

        self.export_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        exp_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom Container: Progress Bar & Export Action
        bottom_exp_frame = ttk.Frame(parent)
        bottom_exp_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(2, 0))

        # Progress bar for exporting
        self.prog_exporter = ttk.Progressbar(bottom_exp_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.prog_exporter.pack(fill=tk.X, pady=(0, 4))

        exp_status_action_frame = ttk.Frame(bottom_exp_frame)
        exp_status_action_frame.pack(fill=tk.X)

        self.lbl_export_status = ttk.Label(exp_status_action_frame, text="Select a .drawio file to export tabs.", foreground="#6c757d")
        self.lbl_export_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_export = ttk.Button(
            exp_status_action_frame,
            text="🚀 Export All Tabs",
            style="Primary.TButton",
            command=self._execute_export_all
        )
        self.btn_export.pack(side=tk.RIGHT)

    def _browse_export_source(self):
        filename = filedialog.askopenfilename(
            title="Select Draw.io File to Export Tabs",
            filetypes=[("Draw.io Diagrams", "*.drawio"), ("XML Files", "*.xml"), ("All Files", "*.*")]
        )
        if filename:
            self.set_export_source_file(filename)

    def set_export_source_file(self, filename: str):
        norm_file = os.path.abspath(filename)
        self.txt_export_src.delete(0, tk.END)
        self.txt_export_src.insert(0, norm_file)

        # Automatically create dedicated subfolder <drawio_stem>_exported in same path
        file_stem = os.path.splitext(os.path.basename(norm_file))[0]
        export_folder = os.path.join(os.path.dirname(norm_file), f"{file_stem}_exported")
        self.txt_export_dst.delete(0, tk.END)
        self.txt_export_dst.insert(0, export_folder)
        self._export_auto_dir_set = True

        self._reload_export_tabs()

    def _browse_export_dest(self):
        initial = self.txt_export_dst.get().strip() or None
        folder = filedialog.askdirectory(title="Select Output Folder for Exported Images", initialdir=initial)
        if folder:
            self.txt_export_dst.delete(0, tk.END)
            self.txt_export_dst.insert(0, os.path.abspath(folder))
            self._export_auto_dir_set = False

    def _reload_export_tabs(self):
        source_file = self.txt_export_src.get().strip()
        if not source_file or not os.path.isfile(source_file):
            for item in self.export_tree.get_children():
                self.export_tree.delete(item)
            self.lbl_tabs_summary.config(text="No valid .drawio file selected.", foreground="#dc3545")
            return

        try:
            self.export_tab_items = DrawIOExporter.get_diagram_tabs(source_file)
            self._refresh_export_table_filenames()
            count = len(self.export_tab_items)
            self.lbl_tabs_summary.config(
                text=f"Found {count} tab(s).",
                foreground="#198754" if count > 0 else "#6c757d"
            )
            dst_name = os.path.basename(self.txt_export_dst.get().strip())
            self.lbl_export_status.config(
                text=f"Ready to export {count} tab(s) into '{dst_name}'.",
                foreground="#212529"
            )
        except Exception as e:
            messagebox.showerror("Error Reading Draw.io File", f"Could not parse tabs from file:\n{e}")
            self.lbl_tabs_summary.config(text="Error parsing tabs.", foreground="#dc3545")

    def _refresh_export_table_filenames(self):
        for item in self.export_tree.get_children():
            self.export_tree.delete(item)

        fmt = self.cbo_export_fmt.get().strip() or "png"
        for tab in self.export_tab_items:
            target_name = f"{tab['clean_name']}.{fmt}"
            self.export_tree.insert("", tk.END, values=(
                str(tab["page_number"]),
                tab["raw_name"],
                target_name,
                "Ready"
            ))

    def _execute_export_all(self):
        if self._is_exporting:
            return

        source_file = self.txt_export_src.get().strip()
        if not source_file or not os.path.isfile(source_file):
            messagebox.showwarning("File Missing", "Please select a valid .drawio file to export.")
            return

        output_dir = self.txt_export_dst.get().strip()
        if not output_dir:
            file_stem = os.path.splitext(os.path.basename(source_file))[0]
            output_dir = os.path.join(os.path.dirname(os.path.abspath(source_file)), f"{file_stem}_exported")
            self.txt_export_dst.delete(0, tk.END)
            self.txt_export_dst.insert(0, output_dir)

        if not self.export_tab_items:
            self._reload_export_tabs()
            if not self.export_tab_items:
                messagebox.showwarning("No Tabs", "No diagram tabs were found in the selected Draw.io file.")
                return

        fmt = self.cbo_export_fmt.get().strip() or "png"
        try:
            scale = float(self.spn_export_scale.get())
        except ValueError:
            scale = 1.0

        total_pages = len(self.export_tab_items)
        self._is_exporting = True
        self.btn_export.config(state=tk.DISABLED)
        self.prog_exporter["value"] = 0
        self.prog_exporter["maximum"] = total_pages
        self.lbl_export_status.config(text=f"Exporting 0 of {total_pages} tabs...", foreground=self.accent_color)
        self.update_idletasks()

        def run_export_thread():
            try:
                def on_progress(current, total, name):
                    percent = int((current / total) * 100)
                    self.lbl_export_status.config(
                        text=f"Exporting tab {current} of {total}: '{name}.{fmt}' ({percent}%)...",
                        foreground=self.accent_color
                    )
                    self.prog_exporter["value"] = current
                    
                    # Update status in treeview
                    children = self.export_tree.get_children()
                    if current - 1 < len(children):
                        item_id = children[current - 1]
                        vals = list(self.export_tree.item(item_id, "values"))
                        vals[3] = "✓ Exported"
                        self.export_tree.item(item_id, values=vals)
                    
                    self.update_idletasks()

                exported = DrawIOExporter.export_all_tabs(
                    drawio_path=source_file,
                    output_dir=output_dir,
                    format=fmt,
                    scale=scale,
                    progress_callback=on_progress
                )

                self.prog_exporter["value"] = total_pages
                self.lbl_export_status.config(
                    text=f"✓ Successfully exported {len(exported)} tab(s) to '{os.path.basename(output_dir)}'!",
                    foreground="#198754"
                )

                resp = messagebox.askyesno(
                    "Export Complete!",
                    f"Successfully exported {len(exported)} tab(s) to folder:\n\n"
                    f"{output_dir}\n\n"
                    f"Would you like to open the export folder?"
                )
                if resp:
                    if sys.platform == "win32":
                        os.startfile(output_dir)
                    elif sys.platform == "darwin":
                        subprocess.run(["open", output_dir])
                    else:
                        subprocess.run(["xdg-open", output_dir])

            except Exception as e:
                self.lbl_export_status.config(text="Export failed.", foreground="#dc3545")
                messagebox.showerror("Export Failed", f"An error occurred while exporting tabs:\n{e}")
            finally:
                self._is_exporting = False
                self.btn_export.config(state=tk.NORMAL)
                self.after(1000, lambda: self.prog_exporter.configure(value=0))

        threading.Thread(target=run_export_thread, daemon=True).start()


def launch_gui():
    app = DrawIOGUI()
    app.mainloop()


if __name__ == "__main__":
    launch_gui()
