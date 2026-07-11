import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
import os

# ====================== 预定义颜色 ======================
PREDEFINED_COLORS = [
    ("白色", "#FFFFFF"), ("黑色", "#000000"), ("红色", "#FF0000"),
    ("綠色", "#00AA00"), ("藍色", "#0000FF"), ("黃色", "#FFFF00"),
    ("青色", "#00FFFF"), ("洋紅", "#FF00FF"), ("灰色", "#AAAAAA"),
    ("橙色", "#FF8000"), ("紫色", "#800080"), ("棕色", "#8B4513"),
    ("粉色", "#FFC0CB"), ("淺綠", "#90EE90"), ("淺藍", "#ADD8E6"),
]

FONT = ("微軟雅黑", 9)
FONT_BOLD = ("微軟雅黑", 9, "bold")
FONT_MATRIX = ("微軟雅黑", 8)   # 矩阵内文字

class WaferMapApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Wafer Map 轉換工具")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)

        self.xml_path = ""
        self.wafer_info_text = ""
        self.matrix = []
        self.rows = 0
        self.cols = 0
        self.transformed_data = []
        self.bin_colors = {}
        self.unique_bins = []

        # 缩放控制
        self.zoom_scale = 1.0        # 当前缩放倍数
        self.base_cell_size = 12     # 基础单元格像素（缩放系数=1时）
        self.max_zoom = 5.0
        self.min_zoom = 0.3

        # 页面容器
        self.page1 = tk.Frame(self.root)
        self.page2 = tk.Frame(self.root)

        self.build_page1()
        self.build_page2()
        self.page1.pack(fill="both", expand=True)

    # ==================== 第一页构建 ====================
    def build_page1(self):
        # 顶部按钮栏
        top_bar = tk.Frame(self.page1)
        top_bar.pack(fill="x", padx=10, pady=5)
        tk.Button(top_bar, text="選擇 XML 文件", command=self.load_xml,
                  font=FONT_BOLD, width=14).pack(side="left")
        self.file_label = tk.Label(top_bar, text="尚未選擇文件", anchor="w",
                                   fg="gray", font=FONT)
        self.file_label.pack(side="left", padx=10)

        # 主区域：左右分栏
        main_pw = tk.PanedWindow(self.page1, orient="horizontal",
                                 sashrelief="raised", sashwidth=4)
        main_pw.pack(fill="both", expand=True, padx=10, pady=(0, 5))

        # ----- 左侧面板 (表头 + 规则) -----
        left_frame = tk.Frame(main_pw)
        left_frame.config(width=360)
        left_frame.pack_propagate(False)

        # 表头信息
        info_lf = tk.LabelFrame(left_frame, text="表頭完整信息", font=FONT_BOLD)
        info_lf.pack(fill="both", expand=True, padx=2, pady=(0, 3))
        self.info_text = tk.Text(info_lf, wrap="word", state="disabled",
                                 font=FONT, relief="flat")
        info_scroll = ttk.Scrollbar(info_lf, orient="vertical",
                                    command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=info_scroll.set)
        self.info_text.pack(side="left", fill="both", expand=True)
        info_scroll.pack(side="right", fill="y")

        # 颜色规则
        rule_lf = tk.LabelFrame(left_frame, text="字元顏色規則", font=FONT_BOLD)
        rule_lf.pack(fill="x", padx=2, pady=(0, 2))
        self.rules_container = tk.Frame(rule_lf)
        self.rules_container.pack(fill="x", pady=3)

        btn_frame = tk.Frame(rule_lf)
        btn_frame.pack(fill="x", pady=3)
        tk.Button(btn_frame, text="新增規則", command=self.add_color_rule,
                  font=FONT, width=8).pack(side="left", padx=2)
        tk.Button(btn_frame, text="刪除選中", command=self.remove_color_rule,
                  font=FONT, width=8).pack(side="left", padx=2)
        tk.Button(btn_frame, text="應用顏色", command=self.apply_colors,
                  font=FONT, bg="#4CAF50", fg="white", width=8).pack(side="left", padx=2)

        self.rule_rows = []

        main_pw.add(left_frame, width=360)

        # ----- 右侧面板 (矩阵预览 + 缩放) -----
        right_frame = tk.Frame(main_pw)
        matrix_lf = tk.LabelFrame(right_frame, text="Wafer 矩陣預覽", font=FONT_BOLD)
        matrix_lf.pack(fill="both", expand=True)

        # 缩放控制栏
        zoom_bar = tk.Frame(matrix_lf)
        zoom_bar.pack(fill="x", pady=2)
        tk.Button(zoom_bar, text="−", command=self.zoom_out,
                  font=("微軟雅黑", 10, "bold"), width=3).pack(side="left", padx=2)
        self.zoom_scale_var = tk.DoubleVar(value=1.0)
        self.zoom_slider = ttk.Scale(zoom_bar, from_=self.min_zoom, to=self.max_zoom,
                                     orient="horizontal", variable=self.zoom_scale_var,
                                     command=self.on_zoom_slider)
        self.zoom_slider.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(zoom_bar, text="+", command=self.zoom_in,
                  font=("微軟雅黑", 10, "bold"), width=3).pack(side="left", padx=2)
        self.zoom_label = tk.Label(zoom_bar, text="100%", font=FONT, width=5)
        self.zoom_label.pack(side="left", padx=5)
        tk.Button(zoom_bar, text="複製矩陣文字", command=self.copy_matrix_text,
                  font=FONT, width=12).pack(side="right", padx=5)

        # Canvas
        canvas_frame = tk.Frame(matrix_lf)
        canvas_frame.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(canvas_frame, bg="white", cursor="arrow")
        h_scroll = ttk.Scrollbar(canvas_frame, orient="horizontal",
                                 command=self.canvas.xview)
        v_scroll = ttk.Scrollbar(canvas_frame, orient="vertical",
                                 command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=h_scroll.set,
                              yscrollcommand=v_scroll.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        h_scroll.grid(row=1, column=0, sticky="ew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

        main_pw.add(right_frame)

        # 底部转换按钮
        tk.Button(self.page1, text="轉換 → 生成坐標列表", command=self.go_to_page2,
                  bg="#2196F3", fg="white", font=("微軟雅黑", 11, "bold"),
                  width=25, height=2).pack(pady=10)

    # ==================== 第二页构建 ====================
    def build_page2(self):
        top_frame = tk.Frame(self.page2)
        top_frame.pack(fill="x", padx=10, pady=5)
        tk.Button(top_frame, text="← 返回預覽", command=self.go_to_page1,
                  font=FONT, width=12).pack(side="left")
        self.status_label = tk.Label(top_frame, text="", fg="blue", font=FONT)
        self.status_label.pack(side="left", padx=20)

        table_frame = tk.Frame(self.page2)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("RW_XY", "Bin", "X", "Y")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                 height=25)
        style = ttk.Style()
        style.configure("Treeview", font=FONT)
        style.configure("Treeview.Heading", font=FONT_BOLD)
        self.tree.heading("RW_XY", text="RW X_Y")
        self.tree.heading("Bin", text="Bin")
        self.tree.heading("X", text="X")
        self.tree.heading("Y", text="Y")
        self.tree.column("RW_XY", width=100, anchor="center")
        self.tree.column("Bin", width=60, anchor="center")
        self.tree.column("X", width=80, anchor="center")
        self.tree.column("Y", width=80, anchor="center")

        scroll_y = ttk.Scrollbar(table_frame, orient="vertical",
                                 command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal",
                                 command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set,
                            xscrollcommand=scroll_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # 绑定 Ctrl+C
        self.tree.bind("<Control-c>", self.copy_tree_selection)
        self.tree.bind("<Control-C>", self.copy_tree_selection)

        btn_frame = tk.Frame(self.page2)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="複製全部數據到剪貼板", command=self.copy_all,
                  bg="#2196F3", fg="white", font=("微軟雅黑", 11),
                  width=25, height=2).pack()

    # ==================== 页面切换 ====================
    def go_to_page1(self):
        self.page2.pack_forget()
        self.page1.pack(fill="both", expand=True)

    def go_to_page2(self):
        if not self.matrix:
            messagebox.showwarning("警告", "請先載入有效的 XML 文件。")
            return
        self.transform()
        self.page1.pack_forget()
        self.page2.pack(fill="both", expand=True)

    # ==================== XML 加载 ====================
    def load_xml(self):
        path = filedialog.askopenfilename(
            title="選擇 Wafer Map XML 文件",
            filetypes=[("XML 文件", "*.xml"), ("所有文件", "*.*")]
        )
        if not path:
            return

        try:
            tree = ET.parse(path)
            root = tree.getroot()
        except Exception as e:
            messagebox.showerror("錯誤", f"無法解析 XML 文件:\n{e}")
            return

        device = root.find(".//{*}Device")
        if device is None:
            messagebox.showerror("錯誤", "找不到 <Device> 節點")
            return

        self.wafer_info_text = self.build_header_text(root, device)

        cols_str = device.get("Columns", "0")
        rows_str = device.get("Rows", "0")
        try:
            self.cols = int(cols_str)
            self.rows = int(rows_str)
        except:
            messagebox.showerror("錯誤", "Columns 或 Rows 屬性無效")
            return

        null_bin = device.get("NullBin", "F")
        data = device.find(".//{*}Data")
        if data is None:
            messagebox.showerror("錯誤", "找不到 <Data> 節點")
            return

        rows_elem = data.findall(".//{*}Row")
        raw_lines = []
        for row in rows_elem:
            text = row.text
            if text:
                text = text.strip()
                if text:
                    raw_lines.append(text)

        self.matrix = []
        for line in raw_lines:
            if len(line) < self.cols:
                line = line.ljust(self.cols, null_bin)
            else:
                line = line[:self.cols]
            self.matrix.append(list(line))

        while len(self.matrix) < self.rows:
            self.matrix.append([null_bin] * self.cols)

        # 唯一字符
        all_chars = set()
        for row in self.matrix:
            all_chars.update(row)
        self.unique_bins = sorted(list(all_chars))

        self.auto_color_map(device, data)
        self.update_header_display()
        self.zoom_scale = 1.0
        self.zoom_scale_var.set(1.0)
        self.draw_matrix()
        self.file_label.config(text=os.path.basename(path), fg="black")

    def build_header_text(self, root, device):
        lines = []
        map_elem = root.find(".//{*}Map")
        if map_elem is None:
            map_elem = root
        lines.append("[Map]")
        for k, v in map_elem.attrib.items():
            lines.append(f"  {k} = {v}")

        lines.append("\n[Device]")
        for k, v in device.attrib.items():
            lines.append(f"  {k} = {v}")

        ref = device.find(".//{*}ReferenceDevice")
        if ref is not None:
            lines.append("\n[ReferenceDevice]")
            for k, v in ref.attrib.items():
                lines.append(f"  {k} = {v}")

        data = device.find(".//{*}Data")
        if data is not None:
            lines.append("\n[Data]")
            for k, v in data.attrib.items():
                lines.append(f"  {k} = {v}")
            bins = data.findall(".//{*}Bin")
            if bins:
                lines.append("\n[Bin Definitions]")
                for i, bin_elem in enumerate(bins, 1):
                    attrs = ", ".join([f"{k}={v}" for k, v in bin_elem.attrib.items()])
                    lines.append(f"  Bin{i}: {attrs}")
        return "\n".join(lines)

    def auto_color_map(self, device, data):
        self.bin_colors.clear()
        bins = data.findall(".//{*}Bin") if data else []
        if bins:
            for bin_elem in bins:
                code = bin_elem.get("BinCode", "")
                quality = bin_elem.get("BinQuality", "")
                if code:
                    if "pass" in quality.lower():
                        color = "#90EE90"
                    elif "fail" in quality.lower() or "reject" in quality.lower():
                        color = "#FF9999"
                    else:
                        color = "#FFFFFF"
                    self.bin_colors[code] = color
        else:
            self.bin_colors = {
                'F': '#CCCCCC', 'y': '#90EE90',
                'x': '#FF9999', 'X': '#8B4513'
            }

        for code in self.unique_bins:
            if code not in self.bin_colors:
                self.bin_colors[code] = "#FFFFFF"

    # ==================== 界面更新 ====================
    def update_header_display(self):
        self.info_text.config(state="normal")
        self.info_text.delete("1.0", "end")
        self.info_text.insert("1.0", self.wafer_info_text)
        self.info_text.config(state="disabled")

    # ==================== 矩阵绘制 (支持缩放) ====================
    def get_cell_size(self):
        """返回当前缩放后的单元格大小 (至少为2)"""
        size = int(self.base_cell_size * self.zoom_scale)
        return max(size, 2)

    def draw_matrix(self):
        self.canvas.delete("all")
        if not self.matrix:
            return

        cell = self.get_cell_size()
        rows = self.rows
        cols = self.cols
        w = cols * cell
        h = rows * cell
        self.canvas.config(scrollregion=(0, 0, w, h))

        font_size = max(6, int(cell * 0.65))
        # 根据单元格大小微调字体
        for r in range(rows):
            for c in range(cols):
                x1 = c * cell
                y1 = r * cell
                x2 = x1 + cell
                y2 = y1 + cell
                char = self.matrix[r][c]
                color = self.bin_colors.get(char, "#FFFFFF")
                self.canvas.create_rectangle(x1, y1, x2, y2,
                                             fill=color, outline="#D0D0D0", width=1)
                self.canvas.create_text(x1 + cell/2, y1 + cell/2,
                                        text=char, font=("微軟雅黑", font_size),
                                        fill="black")

        self.update_zoom_label()

    def update_zoom_label(self):
        percent = int(self.zoom_scale * 100)
        self.zoom_label.config(text=f"{percent}%")
        self.zoom_scale_var.set(self.zoom_scale)

    def zoom_in(self):
        if self.zoom_scale < self.max_zoom:
            self.zoom_scale = min(self.max_zoom, round(self.zoom_scale + 0.2, 1))
            self.draw_matrix()

    def zoom_out(self):
        if self.zoom_scale > self.min_zoom:
            self.zoom_scale = max(self.min_zoom, round(self.zoom_scale - 0.2, 1))
            self.draw_matrix()

    def on_zoom_slider(self, event=None):
        val = self.zoom_scale_var.get()
        self.zoom_scale = round(val, 1)
        self.draw_matrix()

    # ==================== 颜色规则管理 ====================
    def add_color_rule(self):
        if not self.unique_bins:
            messagebox.showinfo("提示", "請先載入矩陣。")
            return

        row_frame = tk.Frame(self.rules_container)
        row_frame.pack(fill="x", pady=1)

        bin_var = tk.StringVar()
        bin_combo = ttk.Combobox(row_frame, textvariable=bin_var,
                                 values=self.unique_bins, width=4,
                                 state="readonly", font=FONT)
        bin_combo.pack(side="left", padx=2)
        if self.unique_bins:
            bin_combo.current(0)

        color_var = tk.StringVar()
        color_names = [name for name, _ in PREDEFINED_COLORS]
        color_combo = ttk.Combobox(row_frame, textvariable=color_var,
                                   values=color_names, width=8,
                                   state="readonly", font=FONT)
        color_combo.pack(side="left", padx=2)
        color_combo.current(0)

        preview = tk.Label(row_frame, text="   ", bg="white",
                           relief="ridge", width=3)
        preview.pack(side="left", padx=5)

        def update_preview(*args):
            name = color_var.get()
            hex_color = next((h for n, h in PREDEFINED_COLORS if n == name), "#FFFFFF")
            preview.config(bg=hex_color)
        color_var.trace("w", update_preview)
        update_preview()

        del_btn = tk.Button(row_frame, text="✕", width=2,
                            font=FONT,
                            command=lambda f=row_frame: self.delete_rule_row(f))
        del_btn.pack(side="right", padx=2)

        self.rule_rows.append((row_frame, bin_var, color_var, preview))

    def delete_rule_row(self, row_frame):
        for i, (f, _, _, _) in enumerate(self.rule_rows):
            if f == row_frame:
                f.destroy()
                del self.rule_rows[i]
                break

    def remove_color_rule(self):
        if self.rule_rows:
            f, _, _, _ = self.rule_rows[-1]
            f.destroy()
            del self.rule_rows[-1]

    def apply_colors(self):
        if not self.matrix:
            return
        for _, bin_var, color_var, _ in self.rule_rows:
            code = bin_var.get()
            if code:
                hex_color = next((h for n, h in PREDEFINED_COLORS if n == color_var.get()), "#FFFFFF")
                self.bin_colors[code] = hex_color
        self.draw_matrix()
        messagebox.showinfo("成功", "顏色規則已應用。")

    def copy_matrix_text(self):
        if not self.matrix:
            return
        text = "\n".join("".join(row) for row in self.matrix)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("成功", "矩陣文字已複製到剪貼板。")

    # ==================== 转换结果 ====================
    def transform(self):
        self.transformed_data.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

        for y in range(1, self.rows + 1):
            for x in range(1, self.cols + 1):
                bin_char = self.matrix[y-1][x-1]
                rw_xy = f"{x}_{y}"
                self.transformed_data.append((rw_xy, bin_char, x, y))
                self.tree.insert("", "end", values=(rw_xy, bin_char, x, y))

        self.status_label.config(text=f"共 {len(self.transformed_data)} 條記錄")

    def copy_tree_selection(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        lines = []
        for item in selected:
            values = self.tree.item(item, "values")
            lines.append("\t".join(str(v) for v in values))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))

    def copy_all(self):
        if not self.transformed_data:
            return
        lines = ["RW X_Y\tBin\tX\tY"]
        for rw_xy, bin_char, x, y in self.transformed_data:
            lines.append(f"{rw_xy}\t{bin_char}\t{x}\t{y}")
        text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("成功", "數據已複製到剪貼板。")


if __name__ == "__main__":
    root = tk.Tk()
    # 全局默认字体（未使用 ttk 的部分）
    root.option_add("*Font", FONT)
    app = WaferMapApp(root)
    root.mainloop()
