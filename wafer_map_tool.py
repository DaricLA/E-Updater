import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import xml.etree.ElementTree as ET
import os
from collections import OrderedDict

# ------------------ 预定义颜色列表 ------------------
PREDEFINED_COLORS = [
    ("白色", "#FFFFFF"), ("黑色", "#000000"), ("红色", "#FF0000"),
    ("绿色", "#00AA00"), ("蓝色", "#0000FF"), ("黄色", "#FFFF00"),
    ("青色", "#00FFFF"), ("洋红", "#FF00FF"), ("灰色", "#AAAAAA"),
    ("橙色", "#FF8000"), ("紫色", "#800080"), ("棕色", "#8B4513"),
    ("粉色", "#FFC0CB"), ("浅绿", "#90EE90"), ("浅蓝", "#ADD8E6"),
]

class WaferMapApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Wafer Map 轉換工具")
        self.root.geometry("1100x800")
        self.root.resizable(True, True)

        # 数据存储
        self.xml_path = ""
        self.wafer_info_text = ""     # 完整表头文本
        self.matrix = []              # 二维列表 [row][col]
        self.rows = 0
        self.cols = 0
        self.transformed_data = []

        # 颜色映射 { bin_code : color_hex }
        self.bin_colors = {}
        self.unique_bins = []         # 矩阵中出现的所有 bin code

        # 页面容器
        self.page1 = tk.Frame(self.root)
        self.page2 = tk.Frame(self.root)

        self.build_page1()
        self.build_page2()
        self.page1.pack(fill="both", expand=True)

    # ==================== 第一页 ====================
    def build_page1(self):
        # 顶部按钮
        top_frame = tk.Frame(self.page1)
        top_frame.pack(fill="x", padx=10, pady=5)
        tk.Button(top_frame, text="選擇 XML 文件", command=self.load_xml, width=15).pack(side="left")
        self.file_label = tk.Label(top_frame, text="尚未選擇文件", anchor="w", fg="gray")
        self.file_label.pack(side="left", padx=10)

        # 左侧主显示区 (表头+矩阵)
        paned = tk.PanedWindow(self.page1, orient="vertical", sashrelief="raised", sashwidth=5)
        paned.pack(fill="both", expand=True, padx=10, pady=5)

        # --- 表头信息 (Text，可复制) ---
        info_frame = tk.LabelFrame(paned, text="表頭完整信息", padx=5, pady=5)
        self.info_text = tk.Text(info_frame, height=8, wrap="word", state="disabled",
                                 font=("TkDefaultFont", 9))
        info_scroll = ttk.Scrollbar(info_frame, orient="vertical", command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=info_scroll.set)
        self.info_text.pack(side="left", fill="both", expand=True)
        info_scroll.pack(side="right", fill="y")
        paned.add(info_frame, height=200)

        # --- 矩阵预览 (Canvas 带滚动) ---
        matrix_frame = tk.LabelFrame(paned, text="Wafer 矩陣預覽 (方框模式)", padx=5, pady=5)
        self.canvas = tk.Canvas(matrix_frame, bg="white", cursor="arrow")
        h_scroll = ttk.Scrollbar(matrix_frame, orient="horizontal", command=self.canvas.xview)
        v_scroll = ttk.Scrollbar(matrix_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        h_scroll.grid(row=1, column=0, sticky="ew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        matrix_frame.grid_rowconfigure(0, weight=1)
        matrix_frame.grid_columnconfigure(0, weight=1)
        paned.add(matrix_frame)

        # --- 颜色规则设置区 ---
        color_frame = tk.LabelFrame(self.page1, text="字元顏色規則 (影響矩陣方框背景)", padx=10, pady=10)
        color_frame.pack(fill="x", padx=10, pady=5)

        self.rules_container = tk.Frame(color_frame)
        self.rules_container.pack(fill="x", pady=5)

        btn_frame = tk.Frame(color_frame)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="新增規則", command=self.add_color_rule).pack(side="left", padx=5)
        tk.Button(btn_frame, text="刪除選中規則", command=self.remove_color_rule).pack(side="left", padx=5)
        tk.Button(btn_frame, text="應用顏色", command=self.apply_colors).pack(side="left", padx=5)
        tk.Button(btn_frame, text="複製矩陣文字", command=self.copy_matrix_text).pack(side="left", padx=5)

        # 存储规则行组件列表
        self.rule_rows = []   # 每行: (frame, bin_var, color_var, preview_label)

        # 转换按钮
        tk.Button(self.page1, text="轉換 → 生成坐標列表", command=self.go_to_page2,
                  bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                  width=25, height=2).pack(pady=10)

    # ==================== 第二页 ====================
    def build_page2(self):
        top_frame = tk.Frame(self.page2)
        top_frame.pack(fill="x", padx=10, pady=5)
        tk.Button(top_frame, text="← 返回預覽", command=self.go_to_page1, width=12).pack(side="left")
        self.status_label = tk.Label(top_frame, text="", fg="blue")
        self.status_label.pack(side="left", padx=20)

        table_frame = tk.Frame(self.page2)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("RW_XY", "Bin", "X", "Y")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=25)
        self.tree.heading("RW_XY", text="RW X_Y")
        self.tree.heading("Bin", text="Bin")
        self.tree.heading("X", text="X")
        self.tree.heading("Y", text="Y")
        self.tree.column("RW_XY", width=100, anchor="center")
        self.tree.column("Bin", width=60, anchor="center")
        self.tree.column("X", width=80, anchor="center")
        self.tree.column("Y", width=80, anchor="center")

        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # 绑定 Ctrl+C 复制选中行
        self.tree.bind("<Control-c>", self.copy_tree_selection)
        self.tree.bind("<Control-C>", self.copy_tree_selection)

        btn_frame = tk.Frame(self.page2)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="複製全部數據到剪貼板", command=self.copy_all,
                  bg="#2196F3", fg="white", font=("Arial", 11),
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

        # 提取所有命名空间 (用于格式化输出)
        namespaces = dict([node for _, node in ET.iterparse(path, events=['start-ns'])])
        # 重新解析以获得完整命名空间信息
        tree = ET.parse(path)
        root = tree.getroot()

        # 查找 Device
        device = root.find(".//{*}Device")
        if device is None:
            messagebox.showerror("錯誤", "找不到 <Device> 節點")
            return

        # 构建完整表头文本
        self.wafer_info_text = self.build_header_text(root, device)

        # 提取基本属性
        cols_str = device.get("Columns", "0")
        rows_str = device.get("Rows", "0")
        try:
            self.cols = int(cols_str)
            self.rows = int(rows_str)
        except:
            messagebox.showerror("錯誤", "Columns 或 Rows 屬性無效")
            return

        null_bin = device.get("NullBin", "F")

        # 查找 Data 节点
        data = device.find(".//{*}Data")
        if data is None:
            messagebox.showerror("錯誤", "找不到 <Data> 節點")
            return

        # 提取所有 Row 文本
        rows_elem = data.findall(".//{*}Row")
        raw_lines = []
        for row in rows_elem:
            text = row.text
            if text:
                text = text.strip()
                if text:
                    raw_lines.append(text)

        # 构建矩阵
        self.matrix = []
        for line in raw_lines:
            if len(line) < self.cols:
                line = line.ljust(self.cols, null_bin)
            else:
                line = line[:self.cols]
            self.matrix.append(list(line))

        while len(self.matrix) < self.rows:
            self.matrix.append([null_bin] * self.cols)

        # 获取所有唯一 bin code
        all_chars = set()
        for row in self.matrix:
            all_chars.update(row)
        self.unique_bins = sorted(list(all_chars))

        # 自动生成初始颜色映射 (基于 Bin 元素或默认)
        self.auto_color_map(device, data)

        # 更新界面
        self.update_header_display()
        self.draw_matrix()
        self.file_label.config(text=os.path.basename(path), fg="black")

    def build_header_text(self, root, device):
        """将 Map, Device, Data(不含Row) 的属性格式化为可读文本"""
        lines = []

        # Map 元素
        map_elem = root.find(".//{*}Map")
        if map_elem is None:
            map_elem = root  # 如果根就是 Map
        lines.append("[Map]")
        for k, v in map_elem.attrib.items():
            lines.append(f"  {k} = {v}")

        # Device
        lines.append("\n[Device]")
        for k, v in device.attrib.items():
            lines.append(f"  {k} = {v}")

        # ReferenceDevice
        ref = device.find(".//{*}ReferenceDevice")
        if ref is not None:
            lines.append("\n[ReferenceDevice]")
            for k, v in ref.attrib.items():
                lines.append(f"  {k} = {v}")

        # Data (属性及 Bin 定义)
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
        """根据 XML 中的 Bin 元素自动建立颜色映射，若无则使用默认"""
        self.bin_colors.clear()
        bins = data.findall(".//{*}Bin") if data is not None else []
        if bins:
            for bin_elem in bins:
                code = bin_elem.get("BinCode", "")
                quality = bin_elem.get("BinQuality", "")
                if code:
                    color = self.quality_to_color(quality)
                    self.bin_colors[code] = color
        else:
            # 默认映射
            default_map = {
                'F': '#CCCCCC',   # 灰色
                'y': '#90EE90',   # 浅绿 (Pass)
                'x': '#FF9999',   # 浅红 (Fail)
                'X': '#8B4513',   # 棕色
            }
            for code, color in default_map.items():
                self.bin_colors[code] = color

        # 为未映射的 bin 分配白色
        for code in self.unique_bins:
            if code not in self.bin_colors:
                self.bin_colors[code] = "#FFFFFF"

    def quality_to_color(self, quality):
        quality_lower = quality.lower()
        if "pass" in quality_lower:
            return "#90EE90"  # 浅绿
        elif "fail" in quality_lower:
            return "#FF9999"  # 浅红
        elif "reject" in quality_lower:
            return "#FFB6C1"  # 粉红
        else:
            return "#FFFFFF"

    # ==================== 界面更新 ====================
    def update_header_display(self):
        self.info_text.config(state="normal")
        self.info_text.delete("1.0", "end")
        self.info_text.insert("1.0", self.wafer_info_text)
        self.info_text.config(state="disabled")

    def draw_matrix(self):
        """在 Canvas 上绘制彩色矩阵方格"""
        self.canvas.delete("all")
        if not self.matrix:
            return

        cell_size = 22   # 像素
        rows = self.rows
        cols = self.cols

        # 设置滚动区域
        self.canvas.config(scrollregion=(0, 0, cols * cell_size, rows * cell_size))

        # 绘制格子
        for r in range(rows):
            for c in range(cols):
                x1 = c * cell_size
                y1 = r * cell_size
                x2 = x1 + cell_size
                y2 = y1 + cell_size
                char = self.matrix[r][c]
                color = self.bin_colors.get(char, "#FFFFFF")
                self.canvas.create_rectangle(x1, y1, x2, y2, fill=color, outline="#CCCCCC", width=1)
                self.canvas.create_text(x1 + cell_size/2, y1 + cell_size/2,
                                        text=char, font=("Courier", 10), fill="black")

    # ==================== 颜色规则管理 ====================
    def add_color_rule(self):
        if not self.unique_bins:
            messagebox.showinfo("提示", "請先載入矩陣。")
            return

        row_frame = tk.Frame(self.rules_container)
        row_frame.pack(fill="x", pady=2)

        # 字符选择
        bin_var = tk.StringVar()
        bin_combo = ttk.Combobox(row_frame, textvariable=bin_var, values=self.unique_bins,
                                 width=4, state="readonly")
        bin_combo.pack(side="left", padx=2)
        if self.unique_bins:
            bin_combo.current(0)

        # 颜色选择 (下拉)
        color_var = tk.StringVar()
        color_names = [name for name, _ in PREDEFINED_COLORS]
        color_combo = ttk.Combobox(row_frame, textvariable=color_var, values=color_names,
                                   width=10, state="readonly")
        color_combo.pack(side="left", padx=2)
        color_combo.current(0)  # 默认白色

        # 颜色预览方块
        preview = tk.Label(row_frame, text="   ", bg="white", relief="ridge", width=3)
        preview.pack(side="left", padx=5)

        # 更新预览
        def update_preview(*args):
            name = color_var.get()
            hex_color = self.get_color_hex(name)
            preview.config(bg=hex_color)
        color_var.trace("w", update_preview)
        update_preview()

        # 删除按钮 (带小×)
        del_btn = tk.Button(row_frame, text="✕", width=2,
                            command=lambda: self.delete_rule_row(row_frame))
        del_btn.pack(side="right", padx=2)

        self.rule_rows.append((row_frame, bin_var, color_var, preview))

    def delete_rule_row(self, row_frame):
        for i, (f, _, _, _) in enumerate(self.rule_rows):
            if f == row_frame:
                f.destroy()
                del self.rule_rows[i]
                break

    def remove_color_rule(self):
        if not self.rule_rows:
            return
        # 删除最后一个规则
        f, _, _, _ = self.rule_rows[-1]
        f.destroy()
        del self.rule_rows[-1]

    def get_color_hex(self, color_name):
        for name, hex_val in PREDEFINED_COLORS:
            if name == color_name:
                return hex_val
        return "#FFFFFF"

    def apply_colors(self):
        """根据规则列表更新 bin_colors 并重绘矩阵"""
        if not self.matrix:
            return
        # 保留自动映射为基础，用规则覆盖
        for _, bin_var, color_var, _ in self.rule_rows:
            code = bin_var.get()
            if code:
                hex_color = self.get_color_hex(color_var.get())
                self.bin_colors[code] = hex_color
        self.draw_matrix()
        messagebox.showinfo("成功", "顏色規則已應用。")

    def copy_matrix_text(self):
        """将当前矩阵文本复制到剪贴板"""
        if not self.matrix:
            return
        text_lines = [''.join(row) for row in self.matrix]
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(text_lines))
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
        """复制 Treeview 中选中的行"""
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
    app = WaferMapApp(root)
    root.mainloop()
