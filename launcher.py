import json
import os
import sys
import time
import subprocess
import ctypes
import ctypes.wintypes
import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox

# ========== 单实例控制 ==========
MUTEX_NAME = "Global\\VSAPE_Toolbox_Launcher_Mutex"

def activate_previous_instance():
    def enum_callback(hwnd, lParam):
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buff = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
        title = buff.value
        if title.startswith("VSAPE工具箱"):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                ctypes.windll.user32.ShowWindow(hwnd, 9)
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                return False
        return True
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_callback), 0)

def check_instance():
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if mutex == 0:
        return True
    if ctypes.windll.kernel32.GetLastError() == 183:
        ctypes.windll.kernel32.CloseHandle(mutex)
        activate_previous_instance()
        return False
    return True

# ========== 启动器配置 ==========
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
        self.desc_labels = []

        # 标题栏
        title_frame = tb.Frame(root, padding=(15, 15, 15, 5))
        title_frame.pack(fill=X)
        tb.Label(title_frame, text="選擇要啟動的工具",
                 font=("微软雅黑", 14, "bold")).pack()

        # 进度条
        self.progress = tb.Progressbar(root, mode='determinate', length=400, maximum=100, bootstyle="info")

        # 外层容器（无 padx，用于容纳左侧间隔、canvas 和滚动条）
        outer_frame = tb.Frame(root)
        outer_frame.pack(fill=BOTH, expand=YES, pady=(5, 10))

        # 左侧 20px 间隔（模拟左侧空白）
        left_spacer = tb.Frame(outer_frame, width=20, height=1)
        left_spacer.pack(side=LEFT, fill=Y)

        # canvas 和滚动条的容器
        canvas_frame = tb.Frame(outer_frame)
        canvas_frame.pack(side=LEFT, fill=BOTH, expand=YES)

        self.canvas = tk.Canvas(canvas_frame, borderwidth=0, highlightthickness=0, bg="#f5f5f5")
        self.scrollbar = tb.Scrollbar(canvas_frame, orient=VERTICAL, command=self.canvas.yview, bootstyle="round")
        self.scrollable_frame = tb.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=YES)

        # 全局滚轮绑定
        self.root.bind("<MouseWheel>", self._on_root_mousewheel)
        self.canvas.bind("<MouseWheel>", self._on_root_mousewheel)

        if self.tools:
            self.create_tool_grid()
        else:
            tb.Label(self.scrollable_frame, text="無可用工具",
                     font=("微软雅黑", 12)).pack(pady=50)

        self.root.after(100, self._delayed_layout_update)

    def _delayed_layout_update(self):
        self._update_descriptions_wrap()
        self.root.update_idletasks()
        self._adjust_window_size()

    def _on_root_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _adjust_window_size(self):
        self.scrollable_frame.update_idletasks()
        self.scrollable_frame.update()
        children = self.scrollable_frame.winfo_children()
        if not children:
            return

        card_height = children[0].winfo_height()
        if card_height <= 0:
            card_height = children[0].winfo_reqheight()

        visible_height = card_height * 3 if card_height > 0 else 600
        title_height = 70
        progress_height = 25
        padding = 30
        win_height = title_height + progress_height + visible_height + padding

        req_width = self.scrollable_frame.winfo_reqwidth()
        scrollbar_width = 20
        win_width = req_width + scrollbar_width + 20  # 左侧间隔20 + 滚动条20

        win_width = max(win_width, 800)
        win_height = max(win_height, 400)

        self.root.geometry(f"{win_width}x{win_height}")
        self.root.resizable(False, False)

    def create_tool_grid(self):
        style = self.root.style
        self.scrollable_frame.columnconfigure(0, weight=1, uniform="col")
        self.scrollable_frame.columnconfigure(1, weight=1, uniform="col")

        row = 0
        col = 0
        for idx, tool in enumerate(self.tools):
            card = tb.LabelFrame(self.scrollable_frame, text="", padding=10, bootstyle="info")
            card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")

            style_name = f"Tool{idx}.TButton"
            bg = parse_color(tool.get("button_color"), DEFAULT_BUTTON_COLOR)
            fg = parse_color(tool.get("text_color"), DEFAULT_TEXT_COLOR)
            font_size = tool.get("font_size", DEFAULT_FONT_SIZE)

            style.configure(style_name,
                            background=bg,
                            foreground=fg,
                            font=("微软雅黑", font_size, "bold"),
                            borderwidth=2,
                            focusthickness=2,
                            focuscolor=style.colors.get('primary'))

            hover_bg = tool.get("hover_color")
            if not hover_bg:
                r = int(bg[1:3], 16); g = int(bg[3:5], 16); b = int(bg[5:7], 16)
                r = max(0, r - 25); g = max(0, g - 25); b = max(0, b - 25)
                hover_bg = f"#{r:02x}{g:02x}{b:02x}"
            style.map(style_name,
                      background=[("active", hover_bg), ("!active", bg)])

            btn = tb.Button(card, text=tool["name"], style=style_name,
                            command=lambda t=tool: self.run_tool(t))
            btn.pack(fill=X, pady=(5, 5))

            desc = tool.get("description", "")
            if desc:
                desc_label = tb.Label(card, text=desc, font=("微软雅黑", 9),
                                      foreground="gray", wraplength=1, justify="left",
                                      anchor="w")
                desc_label.pack(fill=X, pady=(0, 5))
                self.desc_labels.append(desc_label)

            col += 1
            if col > 1:
                col = 0
                row += 1

    def _update_descriptions_wrap(self):
        for label in self.desc_labels:
            card = label.master
            card.update_idletasks()
            card_width = card.winfo_width()
            if card_width > 30:
                label.configure(wraplength=card_width - 15)

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

# ========== 程序入口 ==========
if __name__ == "__main__":
    if not check_instance():
        sys.exit(0)

    root = tb.Window(themename="flatly")
    root.withdraw()
    app = Launcher(root)

    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass

    root.deiconify()
    root.mainloop()
