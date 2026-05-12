import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from openpyxl import Workbook

class ZahnerToExcel:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Zahner 数据 → Excel")
        self.root.geometry("520x520")
        self.root.resizable(True, True)
        self.her_files = []
        self.oer_files = []

        # --- Title ---
        tk.Label(self.root, text="Zahner 数据导入工具",
                 font=("Microsoft YaHei", 14, "bold")).pack(pady=(16, 12))

        # --- HER section ---
        her_frame = tk.LabelFrame(self.root, text="HER 文件", padx=8, pady=6)
        her_frame.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        btn_row = tk.Frame(her_frame)
        btn_row.pack(fill="x")
        self.btn_her = tk.Button(btn_row, text="选择 HER 文件", width=14,
                                 command=self.select_her, bg="#e74c3c", fg="white",
                                 font=("Microsoft YaHei", 10))
        self.btn_her.pack(side="left")
        self.lbl_her_count = tk.Label(btn_row, text="尚未选择文件",
                                      fg="gray", font=("Microsoft YaHei", 9))
        self.lbl_her_count.pack(side="left", padx=12)

        self.her_listbox = tk.Listbox(her_frame, height=6, font=("Consolas", 9))
        self.her_listbox.pack(fill="both", expand=True, pady=(4, 0))
        her_scroll = tk.Scrollbar(self.her_listbox, orient="vertical",
                                  command=self.her_listbox.yview)
        self.her_listbox.configure(yscrollcommand=her_scroll.set)
        her_scroll.pack(side="right", fill="y")

        # --- OER section ---
        oer_frame = tk.LabelFrame(self.root, text="OER 文件", padx=8, pady=6)
        oer_frame.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        btn_row2 = tk.Frame(oer_frame)
        btn_row2.pack(fill="x")
        self.btn_oer = tk.Button(btn_row2, text="选择 OER 文件", width=14,
                                 command=self.select_oer, bg="#3498db", fg="white",
                                 font=("Microsoft YaHei", 10))
        self.btn_oer.pack(side="left")
        self.lbl_oer_count = tk.Label(btn_row2, text="尚未选择文件",
                                      fg="gray", font=("Microsoft YaHei", 9))
        self.lbl_oer_count.pack(side="left", padx=12)

        self.oer_listbox = tk.Listbox(oer_frame, height=6, font=("Consolas", 9))
        self.oer_listbox.pack(fill="both", expand=True, pady=(4, 0))
        oer_scroll = tk.Scrollbar(self.oer_listbox, orient="vertical",
                                  command=self.oer_listbox.yview)
        self.oer_listbox.configure(yscrollcommand=oer_scroll.set)
        oer_scroll.pack(side="right", fill="y")

        # --- Generate button ---
        self.btn_generate = tk.Button(self.root, text="生成 Excel", width=20,
                                      command=self.generate_excel, bg="#27ae60",
                                      fg="white", font=("Microsoft YaHei", 12, "bold"),
                                      height=2)
        self.btn_generate.pack(pady=(10, 6))

        # --- Status ---
        self.lbl_status = tk.Label(self.root, text="", fg="green",
                                   font=("Microsoft YaHei", 9))
        self.lbl_status.pack(pady=(0, 12))

    # ---------- File selection ----------
    def select_her(self):
        paths = filedialog.askopenfilenames(
            title="选择 HER 数据文件",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if paths:
            self.her_files = list(paths)
            self._refresh_listbox(self.her_listbox, self.her_files)
            self.lbl_her_count.config(text=f"已选 {len(self.her_files)} 个文件")
        else:
            self.lbl_her_count.config(text="已取消" if not self.her_files else
                                      f"已选 {len(self.her_files)} 个文件")

    def select_oer(self):
        paths = filedialog.askopenfilenames(
            title="选择 OER 数据文件",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if paths:
            self.oer_files = list(paths)
            self._refresh_listbox(self.oer_listbox, self.oer_files)
            self.lbl_oer_count.config(text=f"已选 {len(self.oer_files)} 个文件")
        else:
            self.lbl_oer_count.config(text="已取消" if not self.oer_files else
                                      f"已选 {len(self.oer_files)} 个文件")

    def _refresh_listbox(self, lb, files):
        lb.delete(0, tk.END)
        for f in files:
            lb.insert(tk.END, os.path.basename(f))

    # ---------- Parse ----------
    def parse_file(self, path):
        voltages = []
        currents = []
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
        return voltages, currents

    # ---------- Generate Excel ----------
    def generate_excel(self):
        if not self.her_files and not self.oer_files:
            messagebox.showwarning("提示", "请先选择 HER 或 OER 文件。")
            return

        out_path = filedialog.asksaveasfilename(
            title="保存 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not out_path:
            self.lbl_status.config(text="已取消", fg="gray")
            return

        wb = Workbook()
        # Remove default sheet; we'll create our own
        wb.remove(wb.active)

        if self.her_files:
            self._write_sheet(wb, "HER", self.her_files)
        if self.oer_files:
            self._write_sheet(wb, "OER", self.oer_files)

        wb.save(out_path)
        self.lbl_status.config(
            text=f"已生成：{out_path}", fg="green")

    def _write_sheet(self, wb, sheet_name, files):
        ws = wb.create_sheet(title=sheet_name)

        # Parse all files first to get max row count
        parsed = []
        max_rows = 0
        for f in files:
            v, c = self.parse_file(f)
            parsed.append((os.path.basename(f), v, c))
            if len(v) > max_rows:
                max_rows = len(v)

        # Row 1: filenames (every 2 columns, col offset 0-based)
        for i, (fname, _, _) in enumerate(parsed):
            col = i * 2 + 1  # 1-based column index (A=1, C=3, E=5, ...)
            ws.cell(row=1, column=col, value=fname)

        # Row 2: headers
        for i in range(len(parsed)):
            col_v = i * 2 + 1
            col_c = i * 2 + 2
            ws.cell(row=2, column=col_v, value="Voltage (V)")
            ws.cell(row=2, column=col_c, value="Current (A)")

        # Row 3+: data
        for row_idx in range(max_rows):
            excel_row = row_idx + 3
            for i, (_, voltages, currents) in enumerate(parsed):
                col_v = i * 2 + 1
                col_c = i * 2 + 2
                if row_idx < len(voltages):
                    ws.cell(row=excel_row, column=col_v, value=voltages[row_idx])
                    ws.cell(row=excel_row, column=col_c, value=currents[row_idx])

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ZahnerToExcel()
    app.run()
