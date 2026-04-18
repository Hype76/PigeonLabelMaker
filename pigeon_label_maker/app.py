from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from PIL import Image, ImageGrab, ImageTk

from .config import load_settings, load_user_presets, save_settings, save_user_presets, setup_logging
from .models import ALIGNMENTS, LAYER_MODES, LINE_MODES, OUTPUT_MODES, AppSettings
from .presets import get_builtin_presets, serialize_preset
from .printing import (
    PROFILE_TEMPLATES,
    apply_print_processing,
    build_print_command,
    list_serial_ports,
    profile_by_name,
    profile_names,
    save_command_file,
    send_to_printer,
    validate_settings,
)
from .rendering import list_font_names, render_label

try:
    import windnd
except ImportError:  # pragma: no cover
    windnd = None


IMAGE_FILE_TYPES = [("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff")]


class LabelMakerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.logger = setup_logging()
        self.settings = load_settings()
        self.current_image: Image.Image | None = None
        self.preview_after_id: str | None = None
        self.preview_refs: dict[str, ImageTk.PhotoImage] = {}
        self.builtin_presets = get_builtin_presets()
        self.user_presets = load_user_presets()

        self.root.title("Pigeon Label Maker")
        self.root.geometry("1320x820")
        self.root.minsize(1120, 720)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.font_var = tk.StringVar(value=self.settings.font_name)
        self.line_mode_var = tk.StringVar(value=self.settings.line_mode)
        self.layer1_text_var = tk.StringVar(value=self.settings.layer1.text)
        self.layer1_mode_var = tk.StringVar(value=self.settings.layer1.mode)
        self.layer1_align_var = tk.StringVar(value=self.settings.layer1.align)
        self.layer2_text_var = tk.StringVar(value=self.settings.layer2.text)
        self.layer2_mode_var = tk.StringVar(value=self.settings.layer2.mode)
        self.layer2_align_var = tk.StringVar(value=self.settings.layer2.align)
        self.profile_var = tk.StringVar(value=self.profile_name_for_id(self.settings.profile_id))
        self.width_var = tk.DoubleVar(value=self.settings.label_width_mm)
        self.height_var = tk.DoubleVar(value=self.settings.label_height_mm)
        self.gap_var = tk.DoubleVar(value=self.settings.gap_mm)
        self.render_dpi_var = tk.IntVar(value=self.settings.render_dpi)
        self.print_dpi_var = tk.IntVar(value=self.settings.print_dpi)
        self.baud_rate_var = tk.IntVar(value=self.settings.baud_rate)
        self.port_var = tk.StringVar(value=self.settings.port)
        self.copies_var = tk.IntVar(value=self.settings.copies)
        self.density_var = tk.IntVar(value=self.settings.density)
        self.contrast_var = tk.DoubleVar(value=self.settings.contrast)
        self.threshold_var = tk.IntVar(value=self.settings.threshold)
        self.invert_var = tk.BooleanVar(value=self.settings.invert)
        self.output_mode_var = tk.StringVar(value=self.settings.output_mode)
        self.preset_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready.")
        self.image_path_var = tk.StringVar(value="No image loaded")

        self.build_ui()
        self.bind_shortcuts()
        self.attach_variable_traces()
        self.populate_preset_choices()
        self.refresh_ports(preserve_current=True)
        self.restore_last_image()
        self.configure_drag_and_drop()
        self.update_control_states()
        self.schedule_preview()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def build_ui(self) -> None:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

        main = ttk.Frame(self.root, padding=12)
        main.grid(sticky="nsew")
        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        controls = ttk.Frame(main)
        controls.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        previews = ttk.Frame(main)
        previews.grid(row=0, column=1, sticky="nsew")
        previews.columnconfigure(0, weight=1)
        previews.columnconfigure(1, weight=1)
        previews.rowconfigure(0, weight=1)

        self.build_preset_frame(controls)
        self.build_layout_frame(controls)
        self.build_layer_frame(controls, "Layer 1", 2)
        self.build_layer_frame(controls, "Layer 2", 3)
        self.build_image_frame(controls)
        self.build_print_frame(controls)
        self.build_action_frame(controls)

        design_frame = ttk.Labelframe(previews, text="Design Preview", padding=10)
        design_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        design_frame.columnconfigure(0, weight=1)
        design_frame.rowconfigure(0, weight=1)
        self.design_preview = ttk.Label(design_frame, anchor="center")
        self.design_preview.grid(sticky="nsew")

        print_frame = ttk.Labelframe(previews, text="Print Preview", padding=10)
        print_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        print_frame.columnconfigure(0, weight=1)
        print_frame.rowconfigure(0, weight=1)
        self.print_preview = ttk.Label(print_frame, anchor="center")
        self.print_preview.grid(sticky="nsew")

        status = ttk.Label(self.root, textvariable=self.status_var, anchor="w", padding=(12, 4))
        status.grid(row=1, column=0, sticky="ew")

    def build_preset_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="Presets", padding=10)
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(0, weight=1)
        self.preset_combo = ttk.Combobox(frame, textvariable=self.preset_var, state="readonly")
        self.preset_combo.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(frame, text="Apply", command=self.apply_selected_preset).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(frame, text="Save Current", command=self.save_current_preset).grid(row=0, column=2)

    def build_layout_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="Layout", padding=10)
        frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Font").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Combobox(frame, textvariable=self.font_var, values=list_font_names(), state="readonly").grid(
            row=0,
            column=1,
            sticky="ew",
            pady=2,
        )

        ttk.Label(frame, text="Text Layout").grid(row=1, column=0, sticky="w", pady=2)
        self.line_mode_combo = ttk.Combobox(frame, textvariable=self.line_mode_var, values=LINE_MODES, state="readonly")
        self.line_mode_combo.grid(row=1, column=1, sticky="ew", pady=2)

    def build_layer_frame(self, parent: ttk.Frame, title: str, row: int) -> None:
        frame = ttk.Labelframe(parent, text=title, padding=10)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(1, weight=1)

        is_first = title.endswith("1")
        text_var = self.layer1_text_var if is_first else self.layer2_text_var
        mode_var = self.layer1_mode_var if is_first else self.layer2_mode_var
        align_var = self.layer1_align_var if is_first else self.layer2_align_var

        ttk.Label(frame, text="Content").grid(row=0, column=0, sticky="w", pady=2)
        entry = ttk.Entry(frame, textvariable=text_var)
        entry.grid(row=0, column=1, columnspan=3, sticky="ew", pady=2)

        upper_button = ttk.Button(frame, text="UPPER", command=lambda target=text_var: self.transform_text(target, "upper"))
        upper_button.grid(row=1, column=1, sticky="w", pady=2)
        title_button = ttk.Button(frame, text="Title Case", command=lambda target=text_var: self.transform_text(target, "title"))
        title_button.grid(row=1, column=2, sticky="w", pady=2)

        ttk.Label(frame, text="Type").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Combobox(frame, textvariable=mode_var, values=LAYER_MODES, state="readonly").grid(
            row=2,
            column=1,
            columnspan=3,
            sticky="ew",
            pady=2,
        )

        ttk.Label(frame, text="Align").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Combobox(frame, textvariable=align_var, values=ALIGNMENTS, state="readonly").grid(
            row=3,
            column=1,
            columnspan=3,
            sticky="ew",
            pady=2,
        )

        if is_first:
            self.layer1_entry = entry
            self.layer1_upper_button = upper_button
            self.layer1_title_button = title_button
        else:
            self.layer2_entry = entry
            self.layer2_upper_button = upper_button
            self.layer2_title_button = title_button

    def build_image_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="Image", padding=10)
        frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)

        ttk.Label(frame, textvariable=self.image_path_var, wraplength=360).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Button(frame, text="Load Image", command=self.load_image).grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(frame, text="Paste Image", command=self.paste_image).grid(row=1, column=1, sticky="ew", padx=6, pady=(8, 0))
        ttk.Button(frame, text="Clear Image", command=self.clear_image).grid(row=1, column=2, sticky="ew", pady=(8, 0))
        ttk.Label(frame, text="Tip: drag an image file onto the window.").grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def build_print_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="Print Settings", padding=10)
        frame.grid(row=5, column=0, sticky="ew", pady=(0, 10))
        for index in range(4):
            frame.columnconfigure(index, weight=1)

        ttk.Label(frame, text="Profile").grid(row=0, column=0, sticky="w", pady=2)
        self.profile_combo = ttk.Combobox(frame, textvariable=self.profile_var, values=profile_names(), state="readonly")
        self.profile_combo.grid(row=0, column=1, columnspan=3, sticky="ew", pady=2)
        self.profile_combo.bind("<<ComboboxSelected>>", self.on_profile_selected)

        ttk.Label(frame, text="Width mm").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Spinbox(frame, from_=10, to=100, increment=1, textvariable=self.width_var).grid(row=1, column=1, sticky="ew", pady=2)
        ttk.Label(frame, text="Height mm").grid(row=1, column=2, sticky="w", pady=2)
        ttk.Spinbox(frame, from_=10, to=100, increment=1, textvariable=self.height_var).grid(row=1, column=3, sticky="ew", pady=2)

        ttk.Label(frame, text="Gap mm").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Spinbox(frame, from_=0, to=20, increment=0.5, textvariable=self.gap_var).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Label(frame, text="Render DPI").grid(row=2, column=2, sticky="w", pady=2)
        ttk.Spinbox(frame, from_=150, to=600, increment=50, textvariable=self.render_dpi_var).grid(row=2, column=3, sticky="ew", pady=2)

        ttk.Label(frame, text="Print DPI").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Spinbox(frame, from_=150, to=600, increment=1, textvariable=self.print_dpi_var).grid(row=3, column=1, sticky="ew", pady=2)
        ttk.Label(frame, text="Output").grid(row=3, column=2, sticky="w", pady=2)
        ttk.Combobox(frame, textvariable=self.output_mode_var, values=OUTPUT_MODES, state="readonly").grid(
            row=3,
            column=3,
            sticky="ew",
            pady=2,
        )

        ttk.Label(frame, text="Port").grid(row=4, column=0, sticky="w", pady=2)
        self.port_combo = ttk.Combobox(frame, textvariable=self.port_var)
        self.port_combo.grid(row=4, column=1, sticky="ew", pady=2)
        ttk.Label(frame, text="Baud").grid(row=4, column=2, sticky="w", pady=2)
        ttk.Spinbox(frame, from_=9600, to=230400, increment=9600, textvariable=self.baud_rate_var).grid(
            row=4,
            column=3,
            sticky="ew",
            pady=2,
        )

        self.refresh_ports_button = ttk.Button(frame, text="Refresh Ports", command=self.refresh_ports)
        self.refresh_ports_button.grid(row=5, column=0, sticky="ew", pady=2)
        ttk.Label(frame, text="Copies").grid(row=5, column=1, sticky="w", pady=2)
        ttk.Spinbox(frame, from_=1, to=50, textvariable=self.copies_var).grid(row=5, column=2, sticky="ew", pady=2)
        ttk.Label(frame, text="Density").grid(row=5, column=3, sticky="w", pady=2)
        ttk.Scale(frame, from_=1, to=15, variable=self.density_var, orient="horizontal").grid(
            row=6,
            column=0,
            columnspan=4,
            sticky="ew",
            pady=2,
        )

        ttk.Label(frame, text="Contrast").grid(row=7, column=0, sticky="w", pady=2)
        ttk.Scale(frame, from_=1.0, to=4.0, variable=self.contrast_var, orient="horizontal").grid(
            row=7,
            column=1,
            columnspan=3,
            sticky="ew",
            pady=2,
        )

        ttk.Label(frame, text="Threshold").grid(row=8, column=0, sticky="w", pady=2)
        ttk.Scale(frame, from_=0, to=255, variable=self.threshold_var, orient="horizontal").grid(
            row=8,
            column=1,
            columnspan=3,
            sticky="ew",
            pady=2,
        )

        ttk.Checkbutton(frame, text="Invert", variable=self.invert_var).grid(row=9, column=0, sticky="w", pady=(4, 0))

    def build_action_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.grid(row=6, column=0, sticky="ew")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)
        ttk.Button(frame, text="Refresh Preview", command=self.update_preview).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(frame, text="Export PNG", command=self.export_png).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(frame, text="Print / Save Command", command=self.handle_print).grid(row=0, column=2, sticky="ew", padx=(4, 0))

    def bind_shortcuts(self) -> None:
        self.root.bind("<Control-p>", lambda event: self.handle_print())
        self.root.bind("<Control-s>", lambda event: self.export_png())
        self.root.bind("<Control-o>", lambda event: self.load_image())
        self.root.bind("<Control-Shift-V>", lambda event: self.paste_image())
        self.root.bind("<F5>", lambda event: self.update_preview())
        self.root.bind("<Control-1>", lambda event: self.transform_text(self.layer1_text_var, "upper"))
        self.root.bind("<Control-2>", lambda event: self.transform_text(self.layer2_text_var, "upper"))
        self.root.bind("<Control-Shift-KeyPress-1>", lambda event: self.transform_text(self.layer1_text_var, "title"))
        self.root.bind("<Control-Shift-KeyPress-2>", lambda event: self.transform_text(self.layer2_text_var, "title"))
        self.root.bind("<Delete>", lambda event: self.clear_image())

    def attach_variable_traces(self) -> None:
        variables = [
            self.font_var,
            self.line_mode_var,
            self.layer1_text_var,
            self.layer1_mode_var,
            self.layer1_align_var,
            self.layer2_text_var,
            self.layer2_mode_var,
            self.layer2_align_var,
            self.width_var,
            self.height_var,
            self.gap_var,
            self.render_dpi_var,
            self.print_dpi_var,
            self.baud_rate_var,
            self.port_var,
            self.copies_var,
            self.density_var,
            self.contrast_var,
            self.threshold_var,
            self.invert_var,
            self.output_mode_var,
        ]
        for variable in variables:
            variable.trace_add("write", self.on_setting_changed)

    def configure_drag_and_drop(self) -> None:
        if windnd is None:
            self.status_var.set("Ready. Drag and drop is unavailable because the optional dependency is missing.")
            return
        windnd.hook_dropfiles(self.root, func=self.on_files_dropped)

    def on_files_dropped(self, file_list: list[bytes | str]) -> None:
        if not file_list:
            return
        first = file_list[0]
        path = first.decode("utf-8") if isinstance(first, bytes) else first
        self.load_image(path)

    def profile_name_for_id(self, profile_id: str) -> str:
        profile = PROFILE_TEMPLATES.get(profile_id)
        if profile:
            return profile.name
        return profile_names()[0]

    def populate_preset_choices(self) -> None:
        values = list(self.builtin_presets.keys()) + list(self.user_presets.keys())
        self.preset_combo["values"] = values
        if values and not self.preset_var.get():
            self.preset_var.set(values[0])

    def apply_selected_preset(self) -> None:
        preset_name = self.preset_var.get()
        if not preset_name:
            return
        preset = self.builtin_presets.get(preset_name) or self.user_presets.get(preset_name)
        if not preset:
            return
        payload = self.current_settings().to_dict()
        payload.update(preset)
        self.apply_settings_to_vars(AppSettings.from_dict(payload))
        self.status_var.set(f"Preset loaded: {preset_name}")
        self.schedule_preview()

    def save_current_preset(self) -> None:
        name = simpledialog.askstring("Save Preset", "Preset name:", parent=self.root)
        if not name:
            return
        clean_name = name.strip()
        if not clean_name:
            return
        self.user_presets[clean_name] = serialize_preset(self.current_settings())
        save_user_presets(self.user_presets)
        self.populate_preset_choices()
        self.preset_var.set(clean_name)
        self.status_var.set(f"Preset saved: {clean_name}")

    def apply_settings_to_vars(self, settings: AppSettings) -> None:
        self.font_var.set(settings.font_name)
        self.line_mode_var.set(settings.line_mode)
        self.layer1_text_var.set(settings.layer1.text)
        self.layer1_mode_var.set(settings.layer1.mode)
        self.layer1_align_var.set(settings.layer1.align)
        self.layer2_text_var.set(settings.layer2.text)
        self.layer2_mode_var.set(settings.layer2.mode)
        self.layer2_align_var.set(settings.layer2.align)
        self.profile_var.set(self.profile_name_for_id(settings.profile_id))
        self.width_var.set(settings.label_width_mm)
        self.height_var.set(settings.label_height_mm)
        self.gap_var.set(settings.gap_mm)
        self.render_dpi_var.set(settings.render_dpi)
        self.print_dpi_var.set(settings.print_dpi)
        self.baud_rate_var.set(settings.baud_rate)
        self.port_var.set(settings.port)
        self.copies_var.set(settings.copies)
        self.density_var.set(settings.density)
        self.contrast_var.set(settings.contrast)
        self.threshold_var.set(settings.threshold)
        self.invert_var.set(settings.invert)
        self.output_mode_var.set(settings.output_mode)

    def current_settings(self) -> AppSettings:
        profile = profile_by_name(self.profile_var.get())
        return AppSettings.from_dict(
            {
                "font_name": self.font_var.get(),
                "line_mode": self.line_mode_var.get(),
                "layer1": {
                    "text": self.layer1_text_var.get(),
                    "mode": self.layer1_mode_var.get(),
                    "align": self.layer1_align_var.get(),
                },
                "layer2": {
                    "text": self.layer2_text_var.get(),
                    "mode": self.layer2_mode_var.get(),
                    "align": self.layer2_align_var.get(),
                },
                "profile_id": profile.identifier if profile else "custom",
                "label_width_mm": self.width_var.get(),
                "label_height_mm": self.height_var.get(),
                "gap_mm": self.gap_var.get(),
                "render_dpi": self.render_dpi_var.get(),
                "print_dpi": self.print_dpi_var.get(),
                "baud_rate": self.baud_rate_var.get(),
                "port": self.port_var.get(),
                "copies": self.copies_var.get(),
                "density": int(round(self.density_var.get())),
                "contrast": round(float(self.contrast_var.get()), 2),
                "threshold": int(round(self.threshold_var.get())),
                "invert": bool(self.invert_var.get()),
                "output_mode": self.output_mode_var.get(),
                "last_image_path": self.settings.last_image_path,
                "export_dir": self.settings.export_dir,
                "command_output_dir": self.settings.command_output_dir,
            }
        )

    def on_profile_selected(self, _event: object | None = None) -> None:
        profile = profile_by_name(self.profile_var.get())
        if not profile:
            return
        self.width_var.set(profile.width_mm)
        self.height_var.set(profile.height_mm)
        self.gap_var.set(profile.gap_mm)
        self.render_dpi_var.set(profile.render_dpi)
        self.print_dpi_var.set(profile.print_dpi)
        self.baud_rate_var.set(profile.baud_rate)
        self.schedule_preview()

    def on_setting_changed(self, *_args: object) -> None:
        self.update_control_states()
        self.schedule_preview()

    def update_control_states(self) -> None:
        if self.layer1_mode_var.get() == "Text":
            self.layer1_upper_button.state(["!disabled"])
            self.layer1_title_button.state(["!disabled"])
        else:
            self.layer1_upper_button.state(["disabled"])
            self.layer1_title_button.state(["disabled"])

        if self.layer2_mode_var.get() == "Text":
            self.layer2_upper_button.state(["!disabled"])
            self.layer2_title_button.state(["!disabled"])
        else:
            self.layer2_upper_button.state(["disabled"])
            self.layer2_title_button.state(["disabled"])

        has_text_layer = self.layer1_mode_var.get() == "Text" or self.layer2_mode_var.get() == "Text"
        if has_text_layer:
            self.line_mode_combo.state(["!disabled"])
        else:
            self.line_mode_combo.state(["disabled"])

        printer_mode = self.output_mode_var.get() == "Printer"
        if printer_mode:
            self.port_combo.state(["!disabled"])
            self.refresh_ports_button.state(["!disabled"])
        else:
            self.port_combo.state(["disabled"])
            self.refresh_ports_button.state(["disabled"])

    def refresh_ports(self, preserve_current: bool = False) -> None:
        current = self.port_var.get()
        ports = list_serial_ports()
        self.port_combo["values"] = ports
        if preserve_current and current in ports:
            self.port_var.set(current)
        elif ports and not current:
            self.port_var.set(ports[0])
        self.status_var.set(f"Detected {len(ports)} serial port(s).")

    def restore_last_image(self) -> None:
        if not self.settings.last_image_path:
            return
        image_path = Path(self.settings.last_image_path)
        if image_path.exists():
            self.load_image(str(image_path), announce=False)

    def load_image(self, path: str | None = None, announce: bool = True) -> None:
        if not path:
            path = filedialog.askopenfilename(filetypes=IMAGE_FILE_TYPES)
        if not path:
            return
        try:
            with Image.open(path) as loaded:
                self.current_image = loaded.copy()
            self.settings.last_image_path = path
            self.image_path_var.set(f"Image: {path}")
            self.logger.info("Loaded image %s", path)
            if announce:
                self.status_var.set(f"Loaded image: {Path(path).name}")
            self.schedule_preview()
        except Exception as exc:
            self.logger.exception("Failed to load image")
            messagebox.showerror("Image Error", f"Could not load image.\n\n{exc}")

    def paste_image(self) -> None:
        try:
            clipboard = ImageGrab.grabclipboard()
            if isinstance(clipboard, Image.Image):
                self.current_image = clipboard
                self.settings.last_image_path = ""
                self.image_path_var.set("Image: clipboard")
                self.status_var.set("Pasted image from clipboard.")
                self.schedule_preview()
                return
            if isinstance(clipboard, list) and clipboard:
                self.load_image(str(clipboard[0]))
                return
            messagebox.showinfo("Paste Image", "Clipboard does not contain an image.")
        except Exception as exc:
            self.logger.exception("Paste image failed")
            messagebox.showerror("Paste Error", f"Could not paste image.\n\n{exc}")

    def clear_image(self) -> None:
        self.current_image = None
        self.settings.last_image_path = ""
        self.image_path_var.set("No image loaded")
        self.status_var.set("Image cleared.")
        self.schedule_preview()

    def transform_text(self, target: tk.StringVar, transform: str) -> None:
        if transform == "upper":
            target.set(target.get().upper())
        elif transform == "title":
            target.set(target.get().title())
        self.schedule_preview()

    def schedule_preview(self) -> None:
        if self.preview_after_id is not None:
            self.root.after_cancel(self.preview_after_id)
        self.preview_after_id = self.root.after(120, self.update_preview)

    def fit_preview(self, image: Image.Image, max_size: tuple[int, int]) -> Image.Image:
        preview = image.copy()
        preview.thumbnail(max_size)
        return preview

    def update_preview(self) -> None:
        self.preview_after_id = None
        try:
            settings = self.current_settings()
            rendered = render_label(settings, self.current_image)
            print_ready = apply_print_processing(rendered, settings)
            design_image = self.fit_preview(rendered, (520, 220))
            print_image = self.fit_preview(print_ready.convert("L"), (320, 520))
            self.preview_refs["design"] = ImageTk.PhotoImage(design_image)
            self.preview_refs["print"] = ImageTk.PhotoImage(print_image)
            self.design_preview.configure(image=self.preview_refs["design"])
            self.print_preview.configure(image=self.preview_refs["print"])
            self.status_var.set("Preview updated.")
        except Exception:
            self.logger.exception("Preview update failed")
            self.status_var.set("Preview failed. Check the log file for details.")

    def export_png(self) -> None:
        try:
            settings = self.current_settings()
            rendered = render_label(settings, self.current_image)
            initial_dir = self.settings.export_dir or str(Path.home())
            path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG Image", "*.png")],
                initialdir=initial_dir,
                initialfile="label.png",
            )
            if not path:
                return
            rendered.save(path, format="PNG")
            self.settings.export_dir = str(Path(path).parent)
            self.status_var.set(f"Exported PNG: {Path(path).name}")
            self.logger.info("Exported PNG %s", path)
        except Exception as exc:
            self.logger.exception("Export failed")
            messagebox.showerror("Export Error", f"Could not export PNG.\n\n{exc}")

    def handle_print(self) -> None:
        try:
            settings = self.current_settings()
            errors = validate_settings(settings)
            if errors:
                messagebox.showerror("Validation Error", "\n".join(errors))
                return

            rendered = render_label(settings, self.current_image)
            print_ready = apply_print_processing(rendered, settings)
            command = build_print_command(print_ready, settings)

            if settings.output_mode == "Mock File":
                initial_dir = self.settings.command_output_dir or str(Path.home())
                path = filedialog.asksaveasfilename(
                    defaultextension=".bin",
                    filetypes=[("Printer Command", "*.bin"), ("All Files", "*.*")],
                    initialdir=initial_dir,
                    initialfile="label-command.bin",
                )
                if not path:
                    return
                save_command_file(Path(path), command)
                self.settings.command_output_dir = str(Path(path).parent)
                self.status_var.set(f"Saved print command: {Path(path).name}")
                self.logger.info("Saved mock print command to %s", path)
                return

            send_to_printer(command, settings)
            self.status_var.set(f"Sent {settings.copies} label(s) to {settings.port}.")
            self.logger.info("Printed %s copies to %s", settings.copies, settings.port)
            messagebox.showinfo("Print Complete", f"Sent {settings.copies} label(s) to {settings.port}.")
        except Exception as exc:
            self.logger.exception("Print failed")
            messagebox.showerror("Print Error", f"Could not complete the print job.\n\n{exc}")

    def on_close(self) -> None:
        self.settings = self.current_settings()
        save_settings(self.settings)
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    LabelMakerApp(root)
    root.mainloop()
