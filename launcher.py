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

def parse_color(value):
    """将颜色配置转换为 tkinter 可接受的格式"""
    if value is None:
        return DEFAULT_BUTTON_COLOR
    if isinstance(value, str) and value.startswith("#"):
        return value
    if isinstance(value, (list, tuple)) and len(value) == 3:
        r, g, b = value
        return f'#{r:02x}{g:02x}{b:02x}'
    if isinstance(value, str) and ',' in value:
        try:
            parts = [int(x.strip()) for x in value.split(',')]
            if len(parts) == 3:
                r, g, b = parts
                return f'#{r:02x}{g:02x}{b:02x}'
        except:
            pass
    return DEFAULT_BUTTON_COLOR

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
        self.root.configure(bg="#f5f5f5")

        # 统一字体
        default_font = ("微软雅黑", 10)
        self.root.option_add("*Font", default_font)

        self.tools = load_tools()
        self.processes = {}          # 存储已启动进程，键为 exe_name，值为 Popen 对象

        # ---------- 顶部标题 ----------
        title_frame = ttk.Frame(root)
        title_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        ttk.Label(title_frame, text="選擇要啟動的工具",
                  font=("微软雅黑", 14, "bold")).pack()

        # ---------- 进度条（隐藏） ----------
        self.progress = ttk.Progressbar(root, mode='indeterminate', length=400)

        # ---------- 滚动区域 ----------
        canvas_frame = ttk.Frame(root)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 15))

        self.canvas = tk.Canvas(canvas_frame, borderwidth=0, highlightthickness=0, bg="#f5f5f5")
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

        # 创建工具卡片
        if self.tools:
            self.create_tool_grid()
        else:
            ttk.Label(self.scrollable_frame, text="無可用工具",
                      font=("微软雅黑", 12)).pack(pady=50)

        # 自适应窗口宽度，锁定水平调整
        self.root.update_idletasks()
        req_width = self.scrollable_frame.winfo_reqwidth() + 50
        self.root.geometry(f"{max(req_width, 700)}x620")
        self.root.resizable(False, True)   # 禁止水平拉伸，允许垂直滚动

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def create_tool_grid(self):
        """创建2列的网格布局，按钮统一尺寸"""
        self.scrollable_frame.columnconfigure(0, weight=1, uniform="col")
        self.scrollable_frame.columnconfigure(1, weight=1, uniform="col")

        row = 0
        col = 0
        for tool in self.tools:
            # 卡片容器
            card = tk.Frame(self.scrollable_frame, bg="#f0f0f0", bd=1, relief=tk.RIDGE,
                            padx=10, pady=10)
            card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")

            # 按钮颜色和字体大小
            btn_color = parse_color(tool.get("button_color", DEFAULT_BUTTON_COLOR))
            btn_font_size = tool.get("font_size", DEFAULT_FONT_SIZE)

            # 按钮：高度设为2行，文字自动换行，填充卡片宽度
            btn = tk.Button(card, text=tool["name"],
                            font=("微软雅黑", btn_font_size, "bold"),
                            bg=btn_color, fg="white",
                            activebackground=btn_color, activeforeground="white",
                            relief=tk.RAISED, bd=2,
                            wraplength=250, height=2,        # 统一高度
                            command=lambda exe=tool["exe"]: self.run_tool(exe))
            btn.pack(fill=tk.X, pady=(5, 5))

            # 描述标签
            desc = tool.get("description", "")
            if desc:
                desc_label = tk.Label(card, text=desc,
                                      font=("微软雅黑", 9), fg="gray",
                                      bg="#f0f0f0", wraplength=250,
                                      justify="left")
                desc_label.pack(fill=tk.X, pady=(0, 5))

            col += 1
            if col > 1:
                col = 0
                row += 1

    def run_tool(self, exe_name):
        """启动工具，防止重复运行"""
        exe_path = os.path.join(TOOLS_DIR, exe_name)
        if not os.path.exists(exe_path):
            messagebox.showerror("錯誤", f"找不到程式：{exe_path}")
            return

        # 检查是否已经在运行
        existing = self.processes.get(exe_name)
        if existing is not None and existing.poll() is None:
            messagebox.showinfo("提示", f"工具 '{exe_name}' 已經在執行中。")
            return

        # 显示进度条，并开始模拟加载
        self.progress.pack(pady=(0, 5))
        self.progress.start()
        self.root.update()

        # 启动工具
        try:
            proc = subprocess.Popen([exe_path], shell=True)
            self.processes[exe_name] = proc
        except Exception as e:
            messagebox.showerror("啟動失敗", str(e))
            self.progress.stop()
            self.progress.pack_forget()
            return

        # 监控进程，确认启动后隐藏进度条
        self._check_process_started(exe_name, proc)

    def _check_process_started(self, exe_name, proc):
        """轮询进程是否已启动，启动后延迟隐藏进度条"""
        if proc.poll() is not None:
            # 进程已退出，可能启动失败
            self.progress.stop()
            self.progress.pack_forget()
            if proc.returncode != 0:
                messagebox.showerror("錯誤", f"工具 '{exe_name}' 啟動失敗，返回碼：{proc.returncode}")
            # 清理记录
            self.processes.pop(exe_name, None)
        else:
            # 进程仍在运行，延迟0.8秒后认为启动成功
            self.root.after(800, lambda: self._hide_progress(exe_name))

    def _hide_progress(self, exe_name):
        """隐藏进度条，并确认进程还在运行"""
        proc = self.processes.get(exe_name)
        if proc and proc.poll() is None:
            # 进程正常运行，进度条消失
            self.progress.stop()
            self.progress.pack_forget()
        else:
            # 进程已退出，可能出错
            self.progress.stop()
            self.progress.pack_forget()
            if proc:
                messagebox.showerror("錯誤", f"工具 '{exe_name}' 啟動失敗，返回碼：{proc.returncode}")
            self.processes.pop(exe_name, None)

if __name__ == "__main__":
    root = tk.Tk()
    app = Launcher(root)
    root.mainloop()
