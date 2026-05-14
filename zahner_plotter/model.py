"""Data model for Zahner Plotter — single source of truth for all state."""

COLOR_PALETTE = [
    "#e74c3c", "#3498db", "#2ecc71", "#9b59b6",
    "#f39c12", "#1abc9c", "#e67e22", "#2980b9",
    "#c0392b", "#8e44ad", "#16a085", "#d35400",
]

HER_PRESET = {"x_min": "-1.623", "x_max": "-1.023", "x_label": "Potential (V vs. RHE)", "title": "HER LSV"}
OER_PRESET = {"x_min": "0.0", "x_max": "0.7", "x_label": "Potential (V vs. RHE)", "title": "OER LSV"}


class Model:
    """Holds all application state. View reads from it; Controller writes to it."""

    def __init__(self):
        # {path: {color, width, label, enabled, category}}
        self.files: dict[str, dict] = {}
        # {path: (voltages, currents)}
        self.parsed: dict[str, tuple] = {}

        # Axis
        self.x_min = "-1.623"
        self.x_max = "-1.023"
        self.y_min = "0"
        self.y_max = "0"
        self.auto_y = True

        # Unit
        self.use_density = False
        self.electrode_area = "1"

        # Style
        self.show_grid = True
        self.tick_font_size = "10"
        self.legend_font_size = "9"
        self.legend_frame_on = False
        self.fig_width = "6.5"
        self.fig_height = "4.5"

        # Current preset mode
        self.preset_mode = "her"  # "her" | "oer" | "manual"

    # ── file management ──────────────────────────────────────────

    def add_files(self, paths: list[str], category: str):
        """Register files; skip duplicates. Auto-assign color. Category: 'HER' or 'OER'."""
        for p in paths:
            if p not in self.files:
                idx = len(self.files)
                self.files[p] = {
                    "color": COLOR_PALETTE[idx % len(COLOR_PALETTE)],
                    "width": 1.5,
                    "label": "",
                    "enabled": True,
                    "category": category,
                }

    def remove_file(self, path: str):
        self.files.pop(path, None)
        self.parsed.pop(path, None)

    def toggle_file(self, path: str):
        if path in self.files:
            self.files[path]["enabled"] = not self.files[path]["enabled"]

    def update_file(self, path: str, key: str, value):
        if path in self.files:
            self.files[path][key] = value

    # ── queries ──────────────────────────────────────────────────

    def files_by_category(self, category: str) -> list[str]:
        return [p for p, s in self.files.items() if s.get("category") == category]

    @property
    def all_files(self) -> list[str]:
        """HER first, then OER, stable order within each category."""
        return self.files_by_category("HER") + self.files_by_category("OER")

    @property
    def active_files(self) -> list[str]:
        return [p for p in self.all_files if self.files[p]["enabled"]]

    def file_label(self, path: str) -> str:
        import os
        s = self.files.get(path, {})
        lbl = s.get("label", "")
        return lbl if lbl else os.path.basename(path).rsplit(".", 1)[0]

    # ── axis presets ─────────────────────────────────────────────

    def apply_preset(self, mode: str):
        self.preset_mode = mode
        if mode == "her":
            self.x_min = HER_PRESET["x_min"]
            self.x_max = HER_PRESET["x_max"]
        elif mode == "oer":
            self.x_min = OER_PRESET["x_min"]
            self.x_max = OER_PRESET["x_max"]

    # ── parsing (cached) ─────────────────────────────────────────

    def parse(self, path: str) -> tuple[list[float], list[float]]:
        if path in self.parsed:
            return self.parsed[path]
        voltages, currents = [], []
        with open(path, "r", encoding="utf-8") as fh:
            fh.readline()  # skip header
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                parts = line.split("\t")
                if len(parts) < 3:
                    continue
                voltages.append(float(parts[1]))
                currents.append(float(parts[2]))
        self.parsed[path] = (voltages, currents)
        return self.parsed[path]
