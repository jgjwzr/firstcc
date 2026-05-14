"""View layer — all Tkinter + matplotlib UI widgets."""

import os
import tkinter as tk
from tkinter import ttk, colorchooser

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from zahner_plotter.model import Model, COLOR_PALETTE

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["path.simplify"] = False


# ══════════════════════════════════════════════════════════════════
#  Left panel — matplotlib figure
# ══════════════════════════════════════════════════════════════════

class LeftPanel(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_propagate(False)

        self.fig, self.ax = plt.subplots(figsize=(6.5, 4.5), dpi=200)
        self.fig.subplots_adjust(left=0.12, right=0.95, top=0.93, bottom=0.14)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().configure(width=1, height=1)
        self.canvas.get_tk_widget().grid(row=1, column=0, sticky="nsew")

        toolbar_frame = ttk.Frame(self)
        toolbar_frame.grid(row=0, column=0, sticky="ew")
        self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        self.toolbar.update()

    def set_figsize(self, w, h):
        self.fig.set_size_inches(w, h)


# ══════════════════════════════════════════════════════════════════
#  Right panel sections
# ══════════════════════════════════════════════════════════════════

class FileSection(ttk.LabelFrame):
    """HER / OER file selection."""

    def __init__(self, parent, model: Model):
        super().__init__(parent, text="数据文件", padding=6)
        self.model = model
        self._on_files_added = None  # callback(paths, category)

        self._build()

    def _build(self):
        for cat, color in [("HER", "#e74c3c"), ("OER", "#3498db")]:
            row = ttk.Frame(self)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"{cat}:", width=5, anchor="e").pack(side="left")
            btn = tk.Button(row, text="选择文件", width=8, bg=color, fg="white",
                           font=("Microsoft YaHei", 9))
            btn.pack(side="left", padx=(4, 0))
            btn.configure(command=lambda c=cat: self._select(c))
            lbl = ttk.Label(row, text="未选择", foreground="gray")
            lbl.pack(side="left", padx=6)
            setattr(self, f"btn_{cat.lower()}", btn)
            setattr(self, f"lbl_{cat.lower()}", lbl)

    def _select(self, category):
        from tkinter import filedialog
        paths = filedialog.askopenfilenames(
            title=f"选择 {category} 数据文件",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if paths and self._on_files_added:
            self._on_files_added(list(paths), category)

    def set_callback(self, cb):
        self._on_files_added = cb

    def update_label(self, category: str, count: int):
        lbl: ttk.Label = getattr(self, f"lbl_{category.lower()}")
        lbl.config(text=f"{count} 个文件", foreground="black" if count else "gray")


class AxisSection(ttk.LabelFrame):
    """Axis range + unit controls."""

    def __init__(self, parent, model: Model):
        super().__init__(parent, text="坐标轴", padding=6)
        self.model = model
        self._on_preset = None
        self._build()

    def _build(self):
        # Preset row
        preset_row = ttk.Frame(self)
        preset_row.pack(fill="x")
        ttk.Label(preset_row, text="预设:", width=5).pack(side="left")
        self.preset_var = tk.StringVar(value="her")
        for val, text in [("her", "HER"), ("oer", "OER"), ("manual", "手动")]:
            ttk.Radiobutton(preset_row, text=text, variable=self.preset_var,
                           value=val, command=self._on_preset_changed).pack(side="left")

        # X axis
        x_row = ttk.Frame(self)
        x_row.pack(fill="x", pady=2)
        ttk.Label(x_row, text="X 轴:", width=5).pack(side="left")
        self.x_min = tk.StringVar(value=self.model.x_min)
        self.x_max = tk.StringVar(value=self.model.x_max)
        ttk.Entry(x_row, textvariable=self.x_min, width=8).pack(side="left", padx=2)
        ttk.Label(x_row, text="~").pack(side="left")
        ttk.Entry(x_row, textvariable=self.x_max, width=8).pack(side="left", padx=2)
        ttk.Label(x_row, text="V").pack(side="left", padx=4)

        # Y axis
        y_row = ttk.Frame(self)
        y_row.pack(fill="x", pady=2)
        ttk.Label(y_row, text="Y 轴:", width=5).pack(side="left")
        self.y_min = tk.StringVar(value=self.model.y_min)
        self.y_max = tk.StringVar(value=self.model.y_max)
        self.auto_y = tk.BooleanVar(value=self.model.auto_y)
        ttk.Entry(y_row, textvariable=self.y_min, width=8).pack(side="left", padx=2)
        ttk.Label(y_row, text="~").pack(side="left")
        ttk.Entry(y_row, textvariable=self.y_max, width=8).pack(side="left", padx=2)
        ttk.Checkbutton(y_row, text="自动", variable=self.auto_y).pack(side="left", padx=4)

        # Unit
        unit_row = ttk.Frame(self)
        unit_row.pack(fill="x", pady=2)
        self.use_density = tk.BooleanVar(value=self.model.use_density)
        ttk.Checkbutton(unit_row, text="转换为电流密度 (mA/cm²)",
                       variable=self.use_density).pack(side="left")
        ttk.Label(unit_row, text="面积:").pack(side="left", padx=(6, 2))
        self.electrode_area = tk.StringVar(value=self.model.electrode_area)
        ttk.Entry(unit_row, textvariable=self.electrode_area, width=6).pack(side="left")
        ttk.Label(unit_row, text="cm²").pack(side="left")

    def _on_preset_changed(self):
        if self._on_preset:
            self._on_preset(self.preset_var.get())

    def set_preset_callback(self, cb):
        self._on_preset = cb

    def sync_to_model(self):
        """Push widget values → model (called before replot)."""
        self.model.x_min = self.x_min.get()
        self.model.x_max = self.x_max.get()
        self.model.y_min = self.y_min.get()
        self.model.y_max = self.y_max.get()
        self.model.auto_y = self.auto_y.get()
        self.model.use_density = self.use_density.get()
        self.model.electrode_area = self.electrode_area.get()
        self.model.preset_mode = self.preset_var.get()

    def sync_from_model(self):
        """Pull model values → widgets."""
        self.x_min.set(self.model.x_min)
        self.x_max.set(self.model.x_max)
        self.preset_var.set(self.model.preset_mode)


class StyleSection(ttk.LabelFrame):
    """Figure size, grid, fonts."""

    def __init__(self, parent, model: Model):
        super().__init__(parent, text="图形设置", padding=6)
        self.model = model
        self._on_apply_size = None
        self._build()

    def _build(self):
        # Figure size row
        size_row = ttk.Frame(self)
        size_row.pack(fill="x")
        ttk.Label(size_row, text="尺寸:").pack(side="left")
        self.fig_w = tk.StringVar(value=self.model.fig_width)
        ttk.Spinbox(size_row, textvariable=self.fig_w, from_=2, to=20,
                    increment=0.5, width=5).pack(side="left", padx=2)
        ttk.Label(size_row, text="×").pack(side="left")
        self.fig_h = tk.StringVar(value=self.model.fig_height)
        ttk.Spinbox(size_row, textvariable=self.fig_h, from_=1.5, to=15,
                    increment=0.5, width=5).pack(side="left", padx=2)
        ttk.Label(size_row, text="英寸").pack(side="left", padx=(2, 6))
        self.btn_size = ttk.Button(size_row, text="应用尺寸")
        self.btn_size.pack(side="left")

        # Grid
        self.show_grid = tk.BooleanVar(value=self.model.show_grid)
        ttk.Checkbutton(size_row, text="网格", variable=self.show_grid).pack(side="left", padx=10)

        # Legend font
        font_row = ttk.Frame(self)
        font_row.pack(fill="x", pady=(4, 2))
        ttk.Label(font_row, text="图例字体:").pack(side="left")
        self.legend_font_size = tk.StringVar(value=self.model.legend_font_size)
        ttk.Spinbox(font_row, textvariable=self.legend_font_size,
                    from_=5, to=20, increment=1, width=4).pack(side="left", padx=2)
        self.legend_frame_on = tk.BooleanVar(value=self.model.legend_frame_on)
        ttk.Checkbutton(font_row, text="图例边框",
                       variable=self.legend_frame_on).pack(side="left", padx=8)

        # Tick font
        font_row2 = ttk.Frame(self)
        font_row2.pack(fill="x", pady=(0, 2))
        ttk.Label(font_row2, text="刻度字体:").pack(side="left")
        self.tick_font_size = tk.StringVar(value=self.model.tick_font_size)
        ttk.Spinbox(font_row2, textvariable=self.tick_font_size,
                    from_=6, to=20, increment=1, width=4).pack(side="left", padx=2)
        ttk.Label(font_row2, text="(坐标轴数字)", foreground="gray").pack(side="left", padx=4)

    def set_size_callback(self, cb):
        self.btn_size.configure(command=cb)

    def sync_to_model(self):
        self.model.fig_width = self.fig_w.get()
        self.model.fig_height = self.fig_h.get()
        self.model.show_grid = self.show_grid.get()
        self.model.legend_font_size = self.legend_font_size.get()
        self.model.legend_frame_on = self.legend_frame_on.get()
        self.model.tick_font_size = self.tick_font_size.get()


class CurveSettings(ttk.Frame):
    """Dynamic per-file line-settings list. Rebuilt from model data."""

    def __init__(self, parent, model: Model):
        super().__init__(parent)
        self.model = model
        self._on_changed = None  # callback()
        self._rows: list[dict] = []  # [{path, enabled_var, width_var, label_var}]

    def set_callback(self, cb):
        self._on_changed = cb

    def rebuild(self):
        """Destroy all rows and rebuild from model.files."""
        for w in self.winfo_children():
            w.destroy()
        self._rows.clear()

        files = self.model.all_files
        if not files:
            ttk.Label(self, text="加载文件后此处显示曲线设置", foreground="gray").pack(pady=8)
            return

        # Header
        hdr = ttk.Frame(self)
        hdr.pack(fill="x", pady=(0, 2))
        ttk.Label(hdr, text="✓", width=3).pack(side="left")
        ttk.Label(hdr, text="文件", width=16, anchor="w").pack(side="left")
        ttk.Label(hdr, text="颜色", width=5).pack(side="left", padx=1)
        ttk.Label(hdr, text="线宽", width=5).pack(side="left", padx=1)
        ttk.Label(hdr, text="标签").pack(side="left", padx=2)

        for path in files:
            self._build_row(path)

    def _build_row(self, path: str):
        s = self.model.files[path]
        row = ttk.Frame(self)
        row.pack(fill="x", pady=1)

        # Enabled checkbox
        enabled_var = tk.BooleanVar(value=s["enabled"])
        cb = ttk.Checkbutton(row, variable=enabled_var, width=2,
                            command=lambda p=path, v=enabled_var: self._on_toggle(p, v))
        cb.pack(side="left")

        # Filename
        import os
        fname = os.path.basename(path)
        display = fname if len(fname) <= 18 else fname[:16] + "…"
        ttk.Label(row, text=display, width=18, anchor="w", font=("Consolas", 9)).pack(side="left")

        # Color button
        color_btn = tk.Button(row, text="  ", bg=s["color"], activebackground=s["color"],
                             width=3, relief="ridge",
                             command=lambda p=path: self._pick_color(p))
        color_btn.pack(side="left", padx=3)

        # Width
        width_var = tk.StringVar(value=str(s["width"]))
        w_spin = ttk.Spinbox(row, textvariable=width_var, from_=0.5, to=8.0,
                            increment=0.5, width=4,
                            command=lambda p=path, v=width_var: self._on_width(p, v))
        w_spin.pack(side="left", padx=3)
        # Also bind on focus-out so manual edits are caught
        w_spin.bind("<FocusOut>", lambda e, p=path, v=width_var: self._on_width(p, v))

        # Label
        label_var = tk.StringVar(value=self.model.file_label(path))
        label_entry = ttk.Entry(row, textvariable=label_var, width=12,
                               font=("Microsoft YaHei", 9))
        label_entry.pack(side="left", padx=3)
        label_entry.bind("<FocusOut>", lambda e, p=path, v=label_var: self._on_label(p, v))

        self._rows.append({
            "path": path,
            "enabled_var": enabled_var,
            "width_var": width_var,
            "label_var": label_var,
            "color_btn": color_btn,
        })

    def _on_toggle(self, path, var):
        self.model.update_file(path, "enabled", var.get())
        if self._on_changed:
            self._on_changed()

    def _on_width(self, path, var):
        try:
            self.model.update_file(path, "width", float(var.get()))
        except ValueError:
            pass
        if self._on_changed:
            self._on_changed()

    def _on_label(self, path, var):
        self.model.update_file(path, "label", var.get())
        if self._on_changed:
            self._on_changed()

    def _pick_color(self, path):
        current = self.model.files[path]["color"]
        new_color = colorchooser.askcolor(color=current, title="选择线条颜色")
        if new_color and new_color[1]:
            self.model.update_file(path, "color", new_color[1])
            self.rebuild()
            if self._on_changed:
                self._on_changed()

    def sync_to_model(self):
        """Push all current widget values → model (called before parent rebuilds)."""
        for row_info in self._rows:
            path = row_info["path"]
            self.model.update_file(path, "enabled", row_info["enabled_var"].get())
            try:
                self.model.update_file(path, "width", float(row_info["width_var"].get()))
            except ValueError:
                pass
            self.model.update_file(path, "label", row_info["label_var"].get())


class ActionSection(ttk.Frame):
    """Plot / Save / Export buttons + status label."""

    def __init__(self, parent):
        super().__init__(parent)
        self._build()

    def _build(self):
        self.btn_plot = tk.Button(self, text="绘图", width=10, bg="#27ae60", fg="white",
                                 font=("Microsoft YaHei", 11, "bold"))
        self.btn_plot.pack(side="left", padx=2)

        self.btn_save = tk.Button(self, text="保存图片", width=10, bg="#2c3e50", fg="white",
                                 font=("Microsoft YaHei", 10))
        self.btn_save.pack(side="left", padx=2)

        self.btn_export = tk.Button(self, text="导出 Excel", width=10, bg="#8e44ad", fg="white",
                                   font=("Microsoft YaHei", 10))
        self.btn_export.pack(side="left", padx=2)

        self.lbl_status = ttk.Label(self, text="", foreground="gray")
        self.lbl_status.pack(side="left", padx=12)


# ══════════════════════════════════════════════════════════════════
#  Right panel — scrollable container
# ══════════════════════════════════════════════════════════════════

class RightPanel(ttk.Frame):
    def __init__(self, parent, model: Model):
        super().__init__(parent, width=400)
        self.pack_propagate(False)

        canvas = tk.Canvas(self, highlightthickness=0, width=380)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)

        self.inner = ttk.Frame(canvas)
        self.inner.bind("<Configure>",
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mousewheel scrolling
        def _on_enter(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _on_leave(event):
            canvas.unbind_all("<MouseWheel>")
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", _on_enter)
        canvas.bind("<Leave>", _on_leave)

        # Build sections inside self.inner
        self.file_section = FileSection(self.inner, model)
        self.file_section.pack(fill="x", padx=4, pady=(4, 2))

        self.axis_section = AxisSection(self.inner, model)
        self.axis_section.pack(fill="x", padx=4, pady=2)

        self.style_section = StyleSection(self.inner, model)
        self.style_section.pack(fill="x", padx=4, pady=2)

        curve_frame = ttk.LabelFrame(self.inner, text="曲线设置", padding=6)
        curve_frame.pack(fill="both", expand=True, padx=4, pady=2)
        self.curve_settings = CurveSettings(curve_frame, model)
        self.curve_settings.pack(fill="both", expand=True)

        self.action_section = ActionSection(self.inner)
        self.action_section.pack(fill="x", padx=4, pady=(4, 8))


# ══════════════════════════════════════════════════════════════════
#  Main View — assembles left + right
# ══════════════════════════════════════════════════════════════════

class View:
    """Top-level view. Owns root window and both panels."""

    def __init__(self, root: tk.Tk, model: Model):
        self.root = root
        self.model = model

        root.title("Zahner 数据绘图工具 v2")
        root.geometry("1280x720")
        root.minsize(960, 540)

        self.left = LeftPanel(root)
        self.left.pack(side="left", fill="both", expand=True)

        self.right = RightPanel(root, model)
        self.right.pack(side="right", fill="y", padx=(0, 4), pady=4)

        self.ax = self.left.ax
        self.fig = self.left.fig
        self.canvas = self.left.canvas

    def update_curve_ui(self):
        self.right.curve_settings.rebuild()

    def set_status(self, text: str, color: str = "green"):
        self.right.action_section.lbl_status.config(text=text, foreground=color)
