import json
import os
import subprocess
import tkinter as tk
from tkinter import messagebox

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
        self.root.geometry("550x450")
        self.tools = load_tools()
        if self.tools:
            self.create_widgets()
        else:
            tk.Label(root, text="無可用工具", font=("微軟正黑體", 14)).pack(pady=50)

    def create_widgets(self):
        tk.Label(self.root, text="選擇要啟動的工具", font=("微軟正黑體", 14)).pack(pady=20)

        for tool in self.tools:
            frame = tk.Frame(self.root, relief=tk.RIDGE, bd=2)
            frame.pack(fill=tk.X, padx=20, pady=5)

            tk.Label(frame, text=tool["name"], font=("微軟正黑體", 11, "bold")).pack(anchor=tk.W, padx=10, pady=(5,0))
            tk.Label(frame, text=tool.get("description", ""), fg="gray").pack(anchor=tk.W, padx=10, pady=(0,5))

            btn = tk.Button(frame, text="啟動", command=lambda exe=tool["exe"]: self.run_tool(exe))
            btn.pack(side=tk.RIGHT, padx=10, pady=5)

    def run_tool(self, exe_name):
        exe_path = os.path.join(TOOLS_DIR, exe_name)
        if not os.path.exists(exe_path):
            messagebox.showerror("錯誤", f"找不到程式：{exe_path}")
            return
        try:
            subprocess.Popen([exe_path], shell=True)
        except Exception as e:
            messagebox.showerror("啟動失敗", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = Launcher(root)
    root.mainloop()
