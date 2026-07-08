import json
import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk

TOOLS_CONFIG = "tools.json"
TOOLS_DIR = "Tools"

def load_tools():
    """從設定檔載入工具列表，並過濾掉不存在的 exe"""
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
        self.root.geometry("550x500")
        self.tools = load_tools()

        # ---------- 頂部標題 ----------
        title_frame = ttk.Frame(root)
        title_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        ttk.Label(title_frame, text="選擇要啟動的工具", font=("微軟正黑體", 14)).pack()

        # ---------- 進度條（隱藏） ----------
        self.progress = ttk.Progressbar(root, mode='indeterminate', length=400)
        # 不立即 pack，需要時再顯示

        # ---------- 可滾動區域 ----------
        # 創建 Canvas 與 Scrollbar
        self.canvas = tk.Canvas(root, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True, padx=(15, 0), pady=(5, 15))
        self.scrollbar.pack(side="right", fill="y", padx=(0, 15), pady=(5, 15))

        # 綁定滑鼠滾輪
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

        # 填充工具卡片
        if self.tools:
            self.create_tool_cards()
        else:
            ttk.Label(self.scrollable_frame, text="無可用工具", font=("微軟正黑體", 12)).pack(pady=50)

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def create_tool_cards(self):
        """在滾動區域內動態產生工具卡片"""
        for tool in self.tools:
            # 卡片外框（淺灰背景，模擬卡片效果）
            card = tk.Frame(self.scrollable_frame, bg="#f0f0f0", bd=0, highlightthickness=0, padx=10, pady=8)
            card.pack(fill=tk.X, padx=5, pady=5)

            # 工具名稱（粗體）
            name_label = tk.Label(card, text=tool["name"], font=("微軟正黑體", 11, "bold"),
                                  bg="#f0f0f0", anchor="w")
            name_label.pack(fill=tk.X, pady=(2, 0))

            # 工具描述（限制最大寬度為 400，自動換行，灰色小字）
            desc = tool.get("description", "")
            if desc:
                desc_label = tk.Label(card, text=desc, font=("微軟正黑體", 9),
                                      fg="gray", bg="#f0f0f0", anchor="w",
                                      wraplength=400, justify="left")
                desc_label.pack(fill=tk.X, pady=(0, 5))

            # 啟動按鈕（藍色，與一鍵更新報告按鈕風格一致）
            btn = tk.Button(card, text="啟動", font=("微軟正黑體", 10),
                            bg="#0078D7", fg="white",
                            activebackground="#005A9E", activeforeground="white",
                            relief=tk.RAISED, bd=2,
                            width=8, height=1,
                            command=lambda exe=tool["exe"]: self.run_tool(exe))
            btn.pack(anchor=tk.E, pady=(5, 2))

    def run_tool(self, exe_name):
        exe_path = os.path.join(TOOLS_DIR, exe_name)
        if not os.path.exists(exe_path):
            messagebox.showerror("錯誤", f"找不到程式：{exe_path}")
            return

        # 顯示進度條
        self.progress.pack(pady=(0, 5))
        self.progress.start()
        self.root.update()

        try:
            # 啟動工具（非同步，不阻塞主執行緒）
            subprocess.Popen([exe_path], shell=True)
        except Exception as e:
            messagebox.showerror("啟動失敗", str(e))
        finally:
            # 隱藏進度條
            self.progress.stop()
            self.progress.pack_forget()
            self.root.update()

if __name__ == "__main__":
    root = tk.Tk()
    app = Launcher(root)
    root.mainloop()
