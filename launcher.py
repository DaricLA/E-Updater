import json
import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk

TOOLS_CONFIG = "tools.json"
TOOLS_DIR = "Tools"

# 默认按钮颜色和字体大小
DEFAULT_BUTTON_COLOR = "#0078D7"
DEFAULT_FONT_SIZE = 12

def load_tools():
    """从配置文件加载工具列表，过滤不存在的exe"""
    if not os.path.exists(TOOLS_CONFIG):
        messagebox.showerror("錯誤", f"找不到設定檔 {TOOLS_CONFIG}，請確認檔案是否存在。")
        return []
    try:
        with open(TOOLS_CONFIG, "r", encoding="utf-8") as f:
            tools = json.load(f)
    except Exception as e:
        messagebox.showerror("錯誤", f"設定檔格式錯誤：{e}")
        return []
    valid_tools = []
    for tool in tools:
        exe_path = os.path.join(TOOLS_DIR, tool.get("exe", ""))
        if os.path.exists(exe_path):
            # 补充默认值
            tool.setdefault("button_color", DEFAULT_BUTTON_COLOR)
            tool.setdefault("font_size", DEFAULT_FONT_SIZE)
            valid_tools.append(tool)
        else:
            print(f"警告：工具 '{tool.get('name', '未知')}' 的程式檔 {exe_path} 不存在，將跳過。")
    if not valid_tools:
        messagebox.showinfo("提示", "沒有可用的工具，請檢查 Tools 資料夾和設定檔。")
    return valid_tools

class Launcher:
    def __init__(self, root):
        self.root = root
        self.root.title("VSAPE工具箱")
        self.root.geometry("900x650")          # 固定窗口大小，适合2列布局
        self.root.resizable(False, False)      # 禁止拉伸
        self.tools = load_tools()

        # 字体统一为微软雅黑
        default_font = ("微软雅黑", 10)
        self.root.option_add("*Font", default_font)

        # ---------- 顶部标题 ----------
        title_frame = ttk.Frame(root)
        title_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        ttk.Label(title_frame, text="選擇要啟動的工具", font=("微软雅黑", 14, "bold")).pack()

        # ---------- 进度条（隐藏） ----------
        self.progress = ttk.Progressbar(root, mode='indeterminate', length=400)

        # ---------- 滚动区域（Canvas + Scrollbar） ----------
        canvas_frame = ttk.Frame(root)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 15))

        self.canvas = tk.Canvas(canvas_frame, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # 鼠标滚轮支持
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

        if self.tools:
            self.create_tool_grid()
        else:
            ttk.Label(self.scrollable_frame, text="無可用工具", font=("微软雅黑", 12)).pack(pady=50)

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def create_tool_grid(self):
        """在滚动区域创建 2 列的网格布局，每张卡片显示一个大按钮和描述"""
        # 设置滚动区域的列权重（两列均分）
        self.scrollable_frame.columnconfigure(0, weight=1, uniform="col")
        self.scrollable_frame.columnconfigure(1, weight=1, uniform="col")

        row = 0
        col = 0
        for tool in self.tools:
            # 卡片容器（灰色背景，模拟卡片）
            card = tk.Frame(self.scrollable_frame, bg="#f5f5f5", bd=1, relief=tk.RIDGE,
                            padx=8, pady=8)
            card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")

            # 按钮：工具名称，可配置颜色和字体大小
            btn_font_size = tool.get("font_size", DEFAULT_FONT_SIZE)
            btn_color = tool.get("button_color", DEFAULT_BUTTON_COLOR)
            btn = tk.Button(card, text=tool["name"],
                            font=("微软雅黑", btn_font_size, "bold"),
                            bg=btn_color, fg="white",
                            activebackground=btn_color, activeforeground="white",
                            relief=tk.RAISED, bd=2,
                            wraplength=300,              # 按钮文字自动换行
                            command=lambda exe=tool["exe"]: self.run_tool(exe))
            btn.pack(fill=tk.X, pady=(5, 5))

            # 描述标签（自动换行）
            desc = tool.get("description", "")
            if desc:
                desc_label = tk.Label(card, text=desc,
                                      font=("微软雅黑", 9), fg="gray",
                                      bg="#f5f5f5", wraplength=300,
                                      justify="left")
                desc_label.pack(fill=tk.X, pady=(0, 5))

            # 切换到下一列，满2列换行
            col += 1
            if col > 1:
                col = 0
                row += 1

    def run_tool(self, exe_name):
        exe_path = os.path.join(TOOLS_DIR, exe_name)
        if not os.path.exists(exe_path):
            messagebox.showerror("錯誤", f"找不到程式：{exe_path}")
            return

        self.progress.pack(pady=(0, 5))
        self.progress.start()
        self.root.update()

        try:
            subprocess.Popen([exe_path], shell=True)
        except Exception as e:
            messagebox.showerror("啟動失敗", str(e))
        finally:
            self.progress.stop()
            self.progress.pack_forget()
            self.root.update()

if __name__ == "__main__":
    root = tk.Tk()
    app = Launcher(root)
    root.mainloop()
