import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import xml.etree.ElementTree as ET
import os

class WaferMapApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Wafer Map 轉換工具")
        self.root.geometry("1000x750")
        self.root.resizable(True, True)

        # 数据存储
        self.xml_path = ""
        self.wafer_info = {}          # 表头信息
        self.matrix = []              # 矩阵 (二维列表, 行 × 列)
        self.rows = 0
        self.cols = 0
        self.transformed_data = []    # 转换结果列表

        # 页面容器 (使用 Frame 切换)
        self.page1 = tk.Frame(self.root)
        self.page2 = tk.Frame(self.root)

        self.build_page1()
        self.build_page2()

        # 默认显示第一页
        self.page1.pack(fill="both", expand=True)

    # ---------- 第一页：文件选择与预览 ----------
    def build_page1(self):
        # 顶部按钮
        top_frame = tk.Frame(self.page1)
        top_frame.pack(fill="x", padx=10, pady=5)
        tk.Button(top_frame, text="選擇 XML 文件", command=self.load_xml, width=15).pack(side="left")
        self.file_label = tk.Label(top_frame, text="尚未選擇文件", anchor="w", fg="gray")
        self.file_label.pack(side="left", padx=10)

        # 表头信息
        info_frame = tk.LabelFrame(self.page1, text="表頭信息", padx=10, pady=10)
        info_frame.pack(fill="x", padx=10, pady=5)
        self.info_text = tk.Text(info_frame, height=5, wrap="word", state="disabled")
        self.info_text.pack(fill="x")

        # 矩阵预览
        matrix_frame = tk.LabelFrame(self.page1, text="Wafer 矩陣預覽 (等寬字體)", padx=10, pady=10)
        matrix_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.matrix_display = scrolledtext.ScrolledText(
            matrix_frame, wrap="none", font=("Courier", 10),
            state="disabled", width=120, height=20
        )
        self.matrix_display.pack(fill="both", expand=True)

        # 转换按钮
        tk.Button(self.page1, text="轉換 → 生成坐標列表", command=self.go_to_page2,
                  bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                  width=25, height=2).pack(pady=15)

    # ---------- 第二页：转换结果与复制 ----------
    def build_page2(self):
        # 顶部导航
        top_frame = tk.Frame(self.page2)
        top_frame.pack(fill="x", padx=10, pady=5)
        tk.Button(top_frame, text="← 返回預覽", command=self.go_to_page1, width=12).pack(side="left")
        self.status_label = tk.Label(top_frame, text="", fg="blue")
        self.status_label.pack(side="left", padx=20)

        # 表格框架
        table_frame = tk.Frame(self.page2)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Treeview 带滚动条
        columns = ("XY", "Bin", "Row", "Col")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=25)
        self.tree.heading("XY", text="Col_Row")
        self.tree.heading("Bin", text="Bin")
        self.tree.heading("Row", text="Row")
        self.tree.heading("Col", text="Col")
        self.tree.column("XY", width=100, anchor="center")
        self.tree.column("Bin", width=60, anchor="center")
        self.tree.column("Row", width=80, anchor="center")
        self.tree.column("Col", width=80, anchor="center")

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # 复制按钮
        btn_frame = tk.Frame(self.page2)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="複製全部數據到剪貼板", command=self.copy_all,
                  bg="#2196F3", fg="white", font=("Arial", 11),
                  width=25, height=2).pack()

    # ---------- 页面切换 ----------
    def go_to_page1(self):
        self.page2.pack_forget()
        self.page1.pack(fill="both", expand=True)

    def go_to_page2(self):
        if not self.matrix:
            messagebox.showwarning("警告", "請先載入有效的 XML 文件並生成矩陣。")
            return
        # 执行转换
        self.transform()
        self.page1.pack_forget()
        self.page2.pack(fill="both", expand=True)

    # ---------- XML 加载与矩阵构建 ----------
    def load_xml(self):
        path = filedialog.askopenfilename(
            title="選擇 Wafer Map XML 文件",
            filetypes=[("XML 文件", "*.xml"), ("所有文件", "*.*")]
        )
        if not path:
            return

        try:
            tree = ET.parse(path)
            root_elem = tree.getroot()
        except Exception as e:
            messagebox.showerror("錯誤", f"無法解析 XML 文件:\n{e}")
            return

        # 命名空间
        ns = {"ns": "http://www.1234.com"}
        # 查找 Device 元素
        device = root_elem.find("ns:Device", ns) if ns else root_elem.find("Device")
        if device is None:
            # 尝试无命名空间
            device = root_elem.find("Device")
            ns = {}
        if device is None:
            messagebox.showerror("錯誤", "找不到 <Device> 節點")
            return

        # 提取表头属性
        self.wafer_info = {
            "WaferId": device.get("WaferId", ""),
            "LotId": device.get("LotId", ""),
            "ProductId": device.get("ProductId", ""),
            "SupplierName": device.get("SupplierName", ""),
            "WaferSize": device.get("WaferSize", ""),
            "Columns": device.get("Columns", "0"),
            "Rows": device.get("Rows", "0"),
            "Orientation": device.get("Orientation", ""),
            "NullBin": device.get("NullBin", "F")
        }
        self.cols = int(self.wafer_info["Columns"])
        self.rows = int(self.wafer_info["Rows"])
        if self.cols == 0 or self.rows == 0:
            messagebox.showerror("錯誤", "Columns 或 Rows 屬性無效")
            return

        # 提取 Data 下的 Row 文本
        data = device.find("ns:Data", ns) if ns else device.find("Data")
        if data is None:
            messagebox.showerror("錯誤", "找不到 <Data> 節點")
            return

        rows_elem = data.findall("ns:Row", ns) if ns else data.findall("Row")
        raw_lines = []
        for row in rows_elem:
            text = row.text
            if text:
                # 去除首尾空白（含空格、换行）
                text = text.strip()
                if text:
                    raw_lines.append(text)
        if len(raw_lines) != self.rows:
            # 警告但继续，以实际读取行为准
            print(f"警告: XML 中有 {len(raw_lines)} 行，但 Rows 属性为 {self.rows}")

        # 构建规整矩阵：补齐或截断到 Columns 长度
        null_bin = self.wafer_info["NullBin"]
        self.matrix = []
        for line in raw_lines:
            # 如果长度不足，用 NullBin 填充右侧；超出则截断
            if len(line) < self.cols:
                line = line.ljust(self.cols, null_bin)
            else:
                line = line[:self.cols]
            self.matrix.append(list(line))

        # 如果行数不足 Rows，用全 NullBin 行补齐
        while len(self.matrix) < self.rows:
            self.matrix.append([null_bin] * self.cols)

        # 更新预览
        self.update_preview()
        self.file_label.config(text=os.path.basename(path), fg="black")
        self.xml_path = path

    def update_preview(self):
        # 表头信息
        info = self.wafer_info
        info_str = (
            f"WaferId: {info['WaferId']}   LotId: {info['LotId']}   Product: {info['ProductId']}\n"
            f"Supplier: {info['SupplierName']}   Size: {info['WaferSize']}   "
            f"Orientation: {info['Orientation']}   NullBin: {info['NullBin']}\n"
            f"Matrix Size: {self.cols} columns × {self.rows} rows"
        )
        self.info_text.config(state="normal")
        self.info_text.delete("1.0", "end")
        self.info_text.insert("1.0", info_str)
        self.info_text.config(state="disabled")

        # 矩阵文本预览
        self.matrix_display.config(state="normal")
        self.matrix_display.delete("1.0", "end")
        for row_chars in self.matrix:
            self.matrix_display.insert("end", "".join(row_chars) + "\n")
        self.matrix_display.config(state="disabled")

    # ---------- 转换逻辑 ----------
    def transform(self):
        self.transformed_data.clear()
        # 清空 tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 遍历矩阵，生成记录 (row 从1开始，col 从1开始)
        for r_idx, row_chars in enumerate(self.matrix, start=1):
            for c_idx, bin_char in enumerate(row_chars, start=1):
                xy = f"{c_idx}_{r_idx}"   # Col_Row 组合
                self.transformed_data.append((xy, bin_char, r_idx, c_idx))
                # 插入 Treeview
                self.tree.insert("", "end", values=(xy, bin_char, r_idx, c_idx))

        self.status_label.config(text=f"共 {len(self.transformed_data)} 條記錄")

    # ---------- 复制全部到剪贴板 ----------
    def copy_all(self):
        if not self.transformed_data:
            messagebox.showinfo("提示", "沒有數據可複製。")
            return
        # 生成 TSV 文本，包含列标题
        lines = ["Col_Row\tBin\tRow\tCol"]
        for xy, bin_char, r, c in self.transformed_data:
            lines.append(f"{xy}\t{bin_char}\t{r}\t{c}")
        text = "\n".join(lines)

        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("成功", "數據已複製到剪貼板（可用 Ctrl+V 粘貼到 Excel 等）。")


if __name__ == "__main__":
    root = tk.Tk()
    app = WaferMapApp(root)
    root.mainloop()
