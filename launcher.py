import json
import os
import sys
import time
import subprocess
import ctypes
import ctypes.wintypes
import tkinter as tk
from tkinter import messagebox, ttk

# ---------- 单实例控制（命名互斥体 + 窗口查找） ----------
MUTEX_NAME = "Global\\VSAPE_Toolbox_Launcher_Mutex"
WM_ACTIVATE = 0x8000 + 1  # 自定义消息，用于激活窗口

def activate_previous_instance():
    """查找已有窗口并激活"""
    # 首先尝试通过互斥体获取的窗口句柄（更可靠的方法见下，但我们先用查找窗口标题）
    # 我们的窗口标题是 "VSAPE工具箱" 可能附带版本号，因此用 FindWindowW 通配符匹配？FindWindow不支持通配符。
    # 简单方法：枚举所有顶层窗口，找到标题以 "VSAPE工具箱" 开头的第一个窗口，然后激活它。
    hwnd = None
    # 枚举窗口回调
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    found_windows = []

    def enum_callback(hwnd, lParam):
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buff = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
        title = buff.value
        if title.startswith("VSAPE工具箱"):
            # 检查是否可见（不是托盘区域等）
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                found_windows.append(hwnd)
                return False  # 停止枚举
        return True

    cb = WNDENUMPROC(enum_callback)
    ctypes.windll.user32.EnumWindows(cb, 0)
    if found_windows:
        hwnd = found_windows[0]
        # 显示窗口（如果最小化）
        ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        return True
    return False

def check_instance():
    """检查是否已有实例，如有则激活并返回 False，否则返回 True 并持有互斥体"""
    # 尝试创建互斥体
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if mutex == 0:
        # 创建失败，致命错误，但忽略
        return True
    last_error = ctypes.windll.kernel32.GetLastError()
    if last_error == 183:  # ERROR_ALREADY_EXISTS
        # 已有实例
        ctypes.windll.kernel32.CloseHandle(mutex)
        # 激活旧窗口
        activate_previous_instance()
        return False
    # 新实例，持有互斥体（当进程退出时系统会自动释放）
    # 我们保存互斥体句柄以便稍后显式关闭（但非必须）
    return True

# 注意：mutex 句柄需保持有效，所以我们将它设为全局或传递给主窗口类（简单起见，在 __main__ 中保存引用）

# ---------- 原有启动器代码 ----------
TOOLS_CONFIG = "tools.json"
TOOLS_DIR = "Tools"

DEFAULT_BUTTON_COLOR = "#0078D7"
DEFAULT_FONT_SIZE = 12
DEFAULT_TEXT_COLOR = "#FFFFFF"

def parse_color(value, default):
    if value is None:
        return default
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
    return default

def load_tools():
    if not os.path.exists(TOOLS_CONFIG):
        messagebox.showerror("錯誤", f"找不到設定檔 {TOOLS_CONFIG}，請確認檔案是否存在。")
        return [], ""
    try:
        with open(TOOLS_CONFIG, "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception as e:
        messagebox.showerror("錯誤", f"設定檔格式錯誤：{e}")
        return [], ""

    if isinstance(config, list):
        tools = config
        version = ""
    else:
        tools = config.get("tools", [])
        version = config.get("version", "")

    valid_tools = []
    for tool in tools:
        exe_path = os.path.join(TOOLS_DIR, tool.get("exe", ""))
        if os.path.exists(exe_path):
            tool.setdefault("start_method", "subprocess")
            tool.setdefault("set_cwd", True)
            tool.setdefault("shell", False)
            tool.setdefault("clean_env", False)
            valid_tools.append(tool)
        else:
            print(f"警告：工具 '{tool.get('name', '未知')}' 的程式檔 {exe_path} 不存在，將跳過。")
    if not valid_tools:
        messagebox.showinfo("提示", "沒有可用的工具，請檢查 Tools 資料夾和設定檔。")
    return valid_tools, version

class Launcher:
    def __init__(self, root):
        self.root = root
        self.tools, self.version = load_tools()

        title = "VSAPE工具箱"
        if self.version:
            title += f" {self.version}"
        self.root.title(title)
        self.root.configure(bg="#f5f5f5")

        default_font = ("微软雅黑", 10)
        self.root.option_add("*Font", default_font)

        self.processes = {}
        self.launch_records = {}

        title_frame = ttk.Frame(root)
        title_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        ttk.Label(title_frame, text="選擇要啟動的工具",
                  font=("微软雅黑", 14, "bold")).pack()

        self.progress = ttk.Progressbar(root, mode='determinate', length=400, maximum=100)

        canvas_frame = ttk.Frame(root)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(5, 10))

        self.canvas = tk.Canvas(canvas_frame, borderwidth=0, highlightthickness=0, bg="#f5f5f5")
        self.scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

        if self.tools:
            self.create_tool_grid()
        else:
            ttk.Label(self.scrollable_frame, text="無可用工具",
                      font=("微软雅黑", 12)).pack(pady=50)

        self.root.update_idletasks()
        self._adjust_window_size()

    def _on_canvas_configure(self, event):
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)

    def _adjust_window_size(self):
        self.scrollable_frame.update_idletasks()
        req_width = self.scrollable_frame.winfo_reqwidth()
        req_height = self.scrollable_frame.winfo_reqheight()

        title_height = 70
        progress_height = 25
        padding = 30

        card_height = 0
        children = self.scrollable_frame.winfo_children()
        if children:
            card_height = children[0].winfo_reqheight()
        visible_height = card_height * 3 if card_height else req_height

        win_height = title_height + progress_height + visible_height + padding
        scrollbar_width = 20
        win_width = req_width + scrollbar_width + 40

        win_width = max(win_width, 800)
        win_height = max(win_height, 500)

        self.root.geometry(f"{win_width}x{win_height}")
        self.root.resizable(False, False)

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def create_tool_grid(self):
        self.scrollable_frame.columnconfigure(0, weight=1, uniform="col")
        self.scrollable_frame.columnconfigure(1, weight=1, uniform="col")

        row = 0
        col = 0
        for tool in self.tools:
            card = tk.Frame(self.scrollable_frame, bg="#f0f0f0", bd=1, relief=tk.RIDGE,
                            padx=10, pady=10)
            card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")

            btn_color = parse_color(tool.get("button_color"), DEFAULT_BUTTON_COLOR)
            text_color = parse_color(tool.get("text_color"), DEFAULT_TEXT_COLOR)
            btn_font_size = tool.get("font_size", DEFAULT_FONT_SIZE)

            btn = tk.Button(card, text=tool["name"],
                            font=("微软雅黑", btn_font_size, "bold"),
                            bg=btn_color, fg=text_color,
                            activebackground=btn_color, activeforeground=text_color,
                            relief=tk.RAISED, bd=2,
                            wraplength=250, height=2,
                            command=lambda t=tool: self.run_tool(t))
            btn.pack(fill=tk.X, pady=(5, 5))

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

    def run_tool(self, tool_config):
        exe_name = tool_config["exe"]
        rel_path = os.path.join(TOOLS_DIR, exe_name)
        target_path = os.path.abspath(rel_path)
        if not os.path.exists(target_path):
            messagebox.showerror("錯誤", f"找不到程式：{target_path}")
            return

        # explorer 启动
        if tool_config.get("start_method") == "explorer":
            last_time = self.launch_records.get(exe_name, 0)
            now = time.time()
            if now - last_time < 5:
                messagebox.showinfo("提示", f"工具 '{exe_name}' 可能正在啟動，請稍後再試。")
                return
            self.launch_records[exe_name] = now

            self.progress.pack(pady=(0, 5))
            self.progress['value'] = 0
            self.root.update()
            self._animate_simple_progress()

            try:
                subprocess.Popen(['explorer.exe', target_path])
            except Exception as e:
                messagebox.showerror("啟動失敗", str(e))
                self.progress.pack_forget()
            return

        # startfile 启动
        if tool_config.get("start_method") == "startfile":
            last_time = self.launch_records.get(exe_name, 0)
            now = time.time()
            if now - last_time < 5:
                messagebox.showinfo("提示", f"工具 '{exe_name}' 可能正在啟動，請稍後再試。")
                return
            self.launch_records[exe_name] = now

            self.progress.pack(pady=(0, 5))
            self.progress['value'] = 0
            self.root.update()
            self._animate_simple_progress()

            try:
                os.startfile(target_path)
            except Exception as e:
                messagebox.showerror("啟動失敗", str(e))
                self.progress.pack_forget()
            return

        # subprocess 启动
        set_cwd = tool_config.get("set_cwd", True)
        use_shell = tool_config.get("shell", False)
        clean_env = tool_config.get("clean_env", False)
        work_dir = os.path.dirname(target_path) if set_cwd else None

        env = os.environ.copy()
        if clean_env:
            for bad_key in ["PYTHONHOME", "PYTHONPATH", "PYTHONSTARTUP", "VIRTUAL_ENV"]:
                env.pop(bad_key, None)

        existing = self.processes.get(exe_name)
        if existing is not None and existing.poll() is None:
            messagebox.showinfo("提示", f"工具 '{exe_name}' 已經在執行中。")
            return

        self.progress.pack(pady=(0, 5))
        self.progress['value'] = 0
        self.root.update()

        try:
            if use_shell:
                proc = subprocess.Popen(f'"{target_path}"', shell=True, cwd=work_dir, env=env)
            else:
                proc = subprocess.Popen([target_path], cwd=work_dir, env=env)
            self.processes[exe_name] = proc
        except Exception as e:
            messagebox.showerror("啟動失敗", str(e))
            self.progress.pack_forget()
            return

        self.root.after(300, lambda: self._check_launch_error(exe_name, proc))

    def _animate_simple_progress(self):
        current = self.progress['value']
        if current < 100:
            new_val = min(current + 20, 100)
            self.progress['value'] = new_val
            self.root.update()
            self.root.after(50, self._animate_simple_progress)
        else:
            self.root.after(200, self.progress.pack_forget)

    def _check_launch_error(self, exe_name, proc):
        if proc.poll() is not None:
            self.progress.pack_forget()
            messagebox.showerror("啟動失敗", f"工具 '{exe_name}' 返回碼 {proc.returncode}，可能缺少必要檔案或程式損壞。")
            self.processes.pop(exe_name, None)
        else:
            self._simulate_progress(exe_name, proc)

    def _simulate_progress(self, exe_name, proc):
        current = self.progress['value']
        if current < 70:
            new_val = min(current + 5, 70)
            self.progress['value'] = new_val
            self.root.update()
            if new_val < 70:
                self.root.after(50, lambda: self._simulate_progress(exe_name, proc))
            else:
                self.root.after(200, lambda: self._finalize_progress(exe_name, proc))
        else:
            self.root.after(200, lambda: self._finalize_progress(exe_name, proc))

    def _finalize_progress(self, exe_name, proc):
        if proc.poll() is None:
            self.progress['value'] = 100
            self.root.update()
            self.root.after(300, lambda: self._hide_progress(exe_name))
        else:
            self.progress.pack_forget()
            messagebox.showerror("啟動失敗", f"工具 '{exe_name}' 返回碼 {proc.returncode}，可能因為缺少檔案或配置錯誤。")
            self.processes.pop(exe_name, None)

    def _hide_progress(self, exe_name):
        self.progress.pack_forget()
        proc = self.processes.get(exe_name)
        if proc and proc.poll() is not None:
            self.processes.pop(exe_name, None)

if __name__ == "__main__":
    # 单实例检查
    if not check_instance():
        sys.exit(0)

    root = tk.Tk()
    app = Launcher(root)
    root.mainloop()
