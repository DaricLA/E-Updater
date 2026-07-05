import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

try:
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor
except ImportError as e:
    import traceback
    err_msg = traceback.format_exc()
    messagebox.showerror("库导入失败", f"缺失必要组件，请反馈以下信息:\n{err_msg}")
    sys.exit(1)

# ---------- 单位转换 ----------
def cm_to_px(cm_val):
    """厘米转像素（openpyxl 按 96 DPI 处理宽高）"""
    return int(cm_val * 37.795)

def cm_to_emu(cm_val):
    """厘米转 EMU（精确公式，1 cm = 360000 EMU）"""
    return int(cm_val * 360000)

CONFIG_FILE = "update_config.json"

def save_config(data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return get_default_config()
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    # 向后兼容旧格式
    if "data_source_path" in cfg and "data_sources" not in cfg:
        old_path = cfg.pop("data_source_path")
        cfg["data_sources"] = [{"alias": "默认数据源", "path": old_path}]
        for m in cfg.get("data_mappings", []):
            if "source_alias" not in m:
                m["source_alias"] = "默认数据源"
    for m in cfg.get("image_mappings", []):
        if "position" not in m:
            m["position"] = "top-left"
            m["offset_x_cm"] = 0
            m["offset_y_cm"] = 0
    if "output_dir" not in cfg:
        cfg["output_dir"] = ""
    return cfg

def get_default_config():
    return {
        "data_sources": [],
        "template_path": "",
        "output_dir": "",
        "output_suffix": "_已更新",
        "data_mappings": [],
        "image_mappings": []
    }

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("VSA_MBO_PBO 一键生成器")
        self.root.geometry("1050x800")
        self.config = load_config()
        self.current_config_name = CONFIG_FILE

        # ---------- 提前创建所有路径变量（避免后续状态栏调用时未定义） ----------
        self.tpl_path_var = tk.StringVar(value=self.config.get("template_path", ""))
        self.out_dir_var = tk.StringVar(value=self.config.get("output_dir", ""))
        self.suffix_var = tk.StringVar(value=self.config.get("output_suffix", "_已更新"))

        # ---------- 顶部状态栏 ----------
        self.status_frame = ttk.Frame(root, relief=tk.RAISED, borderwidth=2)
        self.status_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
        self.status_var = tk.StringVar()
        ttk.Label(self.status_frame, textvariable=self.status_var,
                  background="#D9EAF7", font=("微软雅黑", 10, "bold")).pack(fill=tk.X, padx=10, pady=5)
        self.update_status_bar()

        # ---------- 模板与输出设置 ----------
        frm_tpl = ttk.LabelFrame(root, text="模板与输出设置")
        frm_tpl.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(frm_tpl, text="模板文件:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Entry(frm_tpl, textvariable=self.tpl_path_var, width=70).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(frm_tpl, text="浏览", command=self.browse_tpl).grid(row=0, column=2, padx=5)
        ttk.Label(frm_tpl, text="输出文件夹:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Entry(frm_tpl, textvariable=self.out_dir_var, width=70).grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(frm_tpl, text="浏览", command=self.browse_out_dir).grid(row=1, column=2, padx=5)
        ttk.Label(frm_tpl, text="输出文件后缀:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Entry(frm_tpl, textvariable=self.suffix_var, width=15).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Label(frm_tpl, text="输出文件夹留空则保存到模板所在目录", foreground="gray").grid(row=3, column=1, sticky=tk.W, padx=5)

        # ---------- 数据源管理 ----------
        frm_ds = ttk.LabelFrame(root, text="数据源管理")
        frm_ds.pack(fill=tk.X, padx=10, pady=5)
        self.ds_tree = ttk.Treeview(frm_ds, columns=("alias", "path"), show="headings", height=3)
        self.ds_tree.heading("alias", text="别名")
        self.ds_tree.heading("path", text="路径")
        self.ds_tree.column("alias", width=150)
        self.ds_tree.column("path", width=600)
        self.ds_tree.pack(fill=tk.X, padx=5, pady=5)
        btn_frm_ds = ttk.Frame(frm_ds)
        btn_frm_ds.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(btn_frm_ds, text="添加数据源", command=self.add_datasource).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_ds, text="编辑选中", command=self.edit_datasource).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_ds, text="删除选中", command=self.delete_datasource).pack(side=tk.LEFT, padx=5)

        # ---------- 映射管理 ----------
        nb = ttk.Notebook(root)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ---- 数据映射页 ----
        frm_data = ttk.Frame(nb)
        nb.add(frm_data, text="数据映射")
        self.data_tree = ttk.Treeview(frm_data, columns=("source_alias", "source_cell", "target_cell"), show="headings", height=6)
        self.data_tree.heading("source_alias", text="数据源")
        self.data_tree.heading("source_cell", text="源单元格")
        self.data_tree.heading("target_cell", text="目标单元格")
        self.data_tree.column("source_alias", width=120)
        self.data_tree.column("source_cell", width=220)
        self.data_tree.column("target_cell", width=220)
        self.data_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        btn_frm_data = ttk.Frame(frm_data)
        btn_frm_data.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(btn_frm_data, text="添加", command=self.add_data_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_data, text="编辑选中", command=self.edit_data_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_data, text="复制选中", command=self.copy_data_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_data, text="删除选中", command=lambda: self.delete_selected(self.data_tree, "data")).pack(side=tk.LEFT, padx=5)

        # ---- 图片映射页 ----
        frm_img = ttk.Frame(nb)
        nb.add(frm_img, text="图片映射")
        self.img_tree = ttk.Treeview(frm_img, columns=("number", "folder", "target", "width", "height", "position"), show="headings", height=6)
        self.img_tree.heading("number", text="图片编号")
        self.img_tree.heading("folder", text="图片文件夹")
        self.img_tree.heading("target", text="目标单元格")
        self.img_tree.heading("width", text="宽度(cm)")
        self.img_tree.heading("height", text="高度(cm)")
        self.img_tree.heading("position", text="位置")
        self.img_tree.column("number", width=80)
        self.img_tree.column("folder", width=200)
        self.img_tree.column("target", width=120)
        self.img_tree.column("width", width=70)
        self.img_tree.column("height", width=70)
        self.img_tree.column("position", width=100)
        self.img_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        btn_frm_img = ttk.Frame(frm_img)
        btn_frm_img.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(btn_frm_img, text="添加", command=self.add_image_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_img, text="编辑选中", command=self.edit_image_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_img, text="复制选中", command=self.copy_image_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_img, text="删除选中", command=lambda: self.delete_selected(self.img_tree, "image")).pack(side=tk.LEFT, padx=5)

        # ---------- 控制按钮 ----------
        btn_frm_ctrl = ttk.Frame(root)
        btn_frm_ctrl.pack(pady=10)
        ttk.Button(btn_frm_ctrl, text="一键更新报告", command=self.run_update).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frm_ctrl, text="导出配置", command=self.export_config).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frm_ctrl, text="导入配置", command=self.import_config).pack(side=tk.LEFT, padx=10)

        self.refresh_datasource_tree()
        self.refresh_data_tree()
        self.refresh_image_tree()

    # ---------- 状态栏 ----------
    def update_status_bar(self):
        tpl_name = os.path.basename(self.tpl_path_var.get()) or "未选择"
        cfg_name = os.path.basename(self.current_config_name)
        self.status_var.set(f"📄 当前模板：{tpl_name}    |    ⚙️ 配置文件：{cfg_name}")

    # ---------- 浏览 ----------
    def browse_tpl(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if path:
            self.tpl_path_var.set(path)
            self.update_status_bar()

    def browse_out_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.out_dir_var.set(path)

    # ---------- 数据源管理 ----------
    def refresh_datasource_tree(self):
        for i in self.ds_tree.get_children():
            self.ds_tree.delete(i)
        for ds in self.config.get("data_sources", []):
            self.ds_tree.insert("", tk.END, values=(ds["alias"], ds["path"]))

    def add_datasource(self):
        self._datasource_dialog(None)

    def edit_datasource(self):
        selected = self.ds_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一个数据源")
            return
        idx = self.ds_tree.index(selected[0])
        item = self.config["data_sources"][idx]
        self._datasource_dialog(idx, item)

    def _datasource_dialog(self, edit_idx, item=None):
        popup = tk.Toplevel(self.root)
        popup.title("编辑数据源" if item else "添加数据源")
        ttk.Label(popup, text="别名:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        alias_entry = ttk.Entry(popup, width=30)
        alias_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(popup, text="路径:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        path_var = tk.StringVar()
        path_entry = ttk.Entry(popup, textvariable=path_var, width=30)
        path_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(popup, text="浏览", command=lambda: path_var.set(filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")]))).grid(row=1, column=2, padx=5)
        if item:
            alias_entry.insert(0, item["alias"])
            path_var.set(item["path"])

        def save():
            alias = alias_entry.get().strip()
            path = path_var.get().strip()
            if not alias or not path:
                messagebox.showwarning("输入不完整", "别名和路径不能为空")
                return
            for i, ds in enumerate(self.config["data_sources"]):
                if (edit_idx is None or i != edit_idx) and ds["alias"] == alias:
                    messagebox.showwarning("别名重复", "已存在相同别名的数据源")
                    return
            new_ds = {"alias": alias, "path": path}
            if edit_idx is not None:
                old_alias = self.config["data_sources"][edit_idx]["alias"]
                self.config["data_sources"][edit_idx] = new_ds
                for m in self.config.get("data_mappings", []):
                    if m.get("source_alias") == old_alias:
                        m["source_alias"] = alias
            else:
                self.config["data_sources"].append(new_ds)
            self.refresh_datasource_tree()
            popup.destroy()
        ttk.Button(popup, text="确定", command=save).grid(row=2, column=0, columnspan=3, pady=10)

    def delete_datasource(self):
        selected = self.ds_tree.selection()
        if not selected:
            return
        idx = self.ds_tree.index(selected[0])
        alias = self.config["data_sources"][idx]["alias"]
        refs = [m for m in self.config.get("data_mappings", []) if m.get("source_alias") == alias]
        if refs and not messagebox.askyesno("确认删除", f"数据源 '{alias}' 被 {len(refs)} 条映射引用，继续？"):
            return
        del self.config["data_sources"][idx]
        self.refresh_datasource_tree()

    # ---------- 数据映射操作 ----------
    def refresh_data_tree(self):
        for i in self.data_tree.get_children():
            self.data_tree.delete(i)
        for m in self.config.get("data_mappings", []):
            self.data_tree.insert("", tk.END, values=(m.get("source_alias", ""), m["source_cell"], m["target_cell"]))

    def add_data_mapping(self, prefill=None):
        self._data_dialog(None, prefill)

    def edit_data_mapping(self):
        selected = self.data_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一条映射")
            return
        idx = self.data_tree.index(selected[0])
        item = self.config["data_mappings"][idx]
        self._data_dialog(idx, item)

    def copy_data_mapping(self):
        selected = self.data_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一条要复制的映射")
            return
        idx = self.data_tree.index(selected[0])
        item = self.config["data_mappings"][idx].copy()
        self._data_dialog(None, item)

    def _data_dialog(self, edit_idx, item=None):
        popup = tk.Toplevel(self.root)
        popup.title("编辑数据映射" if edit_idx is not None else "添加数据映射")
        ttk.Label(popup, text="数据源:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        alias_var = tk.StringVar()
        aliases = [ds["alias"] for ds in self.config.get("data_sources", [])]
        if not aliases:
            messagebox.showwarning("无数据源", "请先添加数据源")
            popup.destroy()
            return
        combo = ttk.Combobox(popup, textvariable=alias_var, values=aliases, state="readonly", width=28)
        combo.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(popup, text="源单元格:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        src_entry = ttk.Entry(popup, width=30)
        src_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(popup, text="目标单元格:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        tgt_entry = ttk.Entry(popup, width=30)
        tgt_entry.grid(row=2, column=1, padx=5, pady=5)

        if item:
            alias_var.set(item.get("source_alias", aliases[0]))
            src_entry.insert(0, item["source_cell"])
            tgt_entry.insert(0, item["target_cell"])
        else:
            combo.current(0)

        def save():
            alias = alias_var.get()
            src = src_entry.get().strip()
            tgt = tgt_entry.get().strip()
            if not alias or not src or not tgt:
                messagebox.showwarning("输入不完整", "所有字段不能为空")
                return
            new_map = {"source_alias": alias, "source_cell": src, "target_cell": tgt}
            if edit_idx is not None:
                self.config["data_mappings"][edit_idx] = new_map
            else:
                self.config["data_mappings"].append(new_map)
            self.refresh_data_tree()
            popup.destroy()
        ttk.Button(popup, text="确定", command=save).grid(row=3, column=0, columnspan=2, pady=10)

    # ---------- 图片映射操作 ----------
    def refresh_image_tree(self):
        for i in self.img_tree.get_children():
            self.img_tree.delete(i)
        for m in self.config.get("image_mappings", []):
            pos_text = "默认" if m.get("position") == "top-left" else f"偏移({m.get('offset_x_cm',0)},{m.get('offset_y_cm',0)})cm"
            self.img_tree.insert("", tk.END, values=(
                m["image_number"],
                m["image_folder"],
                m["target_cell"],
                m["width_cm"],
                m["height_cm"],
                pos_text
            ))

    def add_image_mapping(self, prefill=None):
        self._image_dialog(None, prefill)

    def edit_image_mapping(self):
        selected = self.img_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一条映射")
            return
        idx = self.img_tree.index(selected[0])
        item = self.config["image_mappings"][idx]
        self._image_dialog(idx, item)

    def copy_image_mapping(self):
        selected = self.img_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一条要复制的映射")
            return
        idx = self.img_tree.index(selected[0])
        item = self.config["image_mappings"][idx].copy()
        self._image_dialog(None, item)

    def _image_dialog(self, edit_idx, item=None):
        popup = tk.Toplevel(self.root)
        popup.title("编辑图片映射" if edit_idx is not None else "添加图片映射")
        row_idx = 0
        ttk.Label(popup, text="图片编号:").grid(row=row_idx, column=0, padx=5, pady=5, sticky=tk.W)
        num_entry = ttk.Entry(popup, width=30)
        num_entry.grid(row=row_idx, column=1, padx=5, pady=5)
        row_idx += 1
        ttk.Label(popup, text="图片文件夹:").grid(row=row_idx, column=0, padx=5, pady=5, sticky=tk.W)
        folder_var = tk.StringVar()
        folder_entry = ttk.Entry(popup, textvariable=folder_var, width=30)
        folder_entry.grid(row=row_idx, column=1, padx=5, pady=5)
        ttk.Button(popup, text="浏览", command=lambda: folder_var.set(filedialog.askdirectory())).grid(row=row_idx, column=2, padx=5)
        row_idx += 1
        ttk.Label(popup, text="目标单元格:").grid(row=row_idx, column=0, padx=5, pady=5, sticky=tk.W)
        tgt_entry = ttk.Entry(popup, width=30)
        tgt_entry.grid(row=row_idx, column=1, padx=5, pady=5)
        row_idx += 1
        ttk.Label(popup, text="宽度 (cm):").grid(row=row_idx, column=0, padx=5, pady=5, sticky=tk.W)
        w_entry = ttk.Entry(popup, width=10)
        w_entry.grid(row=row_idx, column=1, sticky=tk.W, padx=5)
        row_idx += 1
        ttk.Label(popup, text="高度 (cm):").grid(row=row_idx, column=0, padx=5, pady=5, sticky=tk.W)
        h_entry = ttk.Entry(popup, width=10)
        h_entry.grid(row=row_idx, column=1, sticky=tk.W, padx=5)
        row_idx += 1
        # 位置选择
        ttk.Label(popup, text="位置:").grid(row=row_idx, column=0, padx=5, pady=5, sticky=tk.W)
        pos_var = tk.StringVar(value="top-left")
        pos_frame = ttk.Frame(popup)
        pos_frame.grid(row=row_idx, column=1, columnspan=2, sticky=tk.W)
        ttk.Radiobutton(pos_frame, text="默认（左上角）", variable=pos_var, value="top-left").pack(side=tk.LEFT)
        ttk.Radiobutton(pos_frame, text="自定义偏移", variable=pos_var, value="custom").pack(side=tk.LEFT, padx=10)
        row_idx += 1
        offset_frame = ttk.Frame(popup)
        offset_frame.grid(row=row_idx, column=1, columnspan=2, sticky=tk.W, pady=5)
        ttk.Label(offset_frame, text="X偏移(cm):").pack(side=tk.LEFT)
        x_entry = ttk.Entry(offset_frame, width=8)
        x_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(offset_frame, text="Y偏移(cm):").pack(side=tk.LEFT, padx=(15,0))
        y_entry = ttk.Entry(offset_frame, width=8)
        y_entry.pack(side=tk.LEFT, padx=5)

        def toggle_offset(*args):
            if pos_var.get() == "custom":
                x_entry.config(state="normal")
                y_entry.config(state="normal")
            else:
                x_entry.config(state="disabled")
                y_entry.config(state="disabled")
                x_entry.delete(0, tk.END)
                x_entry.insert(0, "0")
                y_entry.delete(0, tk.END)
                y_entry.insert(0, "0")
        pos_var.trace("w", toggle_offset)

        if item:
            num_entry.insert(0, item["image_number"])
            folder_var.set(item["image_folder"])
            tgt_entry.insert(0, item["target_cell"])
            w_entry.insert(0, str(item["width_cm"]))
            h_entry.insert(0, str(item["height_cm"]))
            pos_var.set(item.get("position", "top-left"))
            x_entry.insert(0, str(item.get("offset_x_cm", 0)))
            y_entry.insert(0, str(item.get("offset_y_cm", 0)))
        else:
            w_entry.insert(0, "3.5")
            h_entry.insert(0, "2.8")
            x_entry.insert(0, "0")
            y_entry.insert(0, "0")
        toggle_offset()

        def save():
            num = num_entry.get().strip()
            folder = folder_var.get().strip()
            tgt = tgt_entry.get().strip()
            try:
                w = float(w_entry.get().strip())
                h = float(h_entry.get().strip())
            except ValueError:
                messagebox.showwarning("输入错误", "宽度和高度必须为数字")
                return
            if pos_var.get() == "custom":
                try:
                    off_x = float(x_entry.get().strip())
                    off_y = float(y_entry.get().strip())
                except ValueError:
                    messagebox.showwarning("输入错误", "偏移值必须为数字")
                    return
            else:
                off_x = 0.0
                off_y = 0.0
            if not num or not folder or not tgt:
                messagebox.showwarning("输入不完整", "所有字段必填")
                return
            new_map = {
                "image_number": num,
                "image_folder": folder,
                "target_cell": tgt,
                "width_cm": w,
                "height_cm": h,
                "position": pos_var.get(),
                "offset_x_cm": off_x,
                "offset_y_cm": off_y
            }
            if edit_idx is not None:
                self.config["image_mappings"][edit_idx] = new_map
            else:
                self.config["image_mappings"].append(new_map)
            self.refresh_image_tree()
            popup.destroy()
        ttk.Button(popup, text="确定", command=save).grid(row=row_idx+1, column=0, columnspan=3, pady=10)

    def delete_selected(self, tree, map_type):
        selected = tree.selection()
        if not selected:
            return
        idx = tree.index(selected[0])
        if map_type == "data":
            del self.config["data_mappings"][idx]
            self.refresh_data_tree()
        else:
            del self.config["image_mappings"][idx]
            self.refresh_image_tree()

    # ---------- 配置导入导出 ----------
    def export_config(self):
        self.save_current_config()
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json")],
            initialfile="config_backup.json"
        )
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(self.config, f, indent=2, ensure_ascii=False)
                messagebox.showinfo("导出成功", f"配置已保存到:\n{path}")
            except Exception as e:
                messagebox.showerror("导出失败", str(e))

    def import_config(self):
        path = filedialog.askopenfilename(filetypes=[("JSON文件", "*.json")])
        if not path:
            return
        if not messagebox.askyesno("确认导入", "导入配置将覆盖当前所有设置，确定继续吗？"):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                new_cfg = json.load(f)
            required_keys = ["data_sources", "template_path", "data_mappings", "image_mappings"]
            for k in required_keys:
                if k not in new_cfg:
                    raise ValueError(f"配置文件缺少必要字段: {k}")
            self.config = new_cfg
            self.current_config_name = path
            self.tpl_path_var.set(self.config.get("template_path", ""))
            self.out_dir_var.set(self.config.get("output_dir", ""))
            self.suffix_var.set(self.config.get("output_suffix", "_已更新"))
            self.refresh_datasource_tree()
            self.refresh_data_tree()
            self.refresh_image_tree()
            self.update_status_bar()
            messagebox.showinfo("导入成功", "配置已导入并更新界面")
        except Exception as e:
            messagebox.showerror("导入失败", f"文件格式错误:\n{str(e)}")

    # ---------- 保存当前配置 ----------
    def save_current_config(self):
        self.config["template_path"] = self.tpl_path_var.get()
        self.config["output_dir"] = self.out_dir_var.get()
        self.config["output_suffix"] = self.suffix_var.get()
        save_config(self.config)

    # ---------- 核心执行 ----------
    def run_update(self):
        self.save_current_config()
        cfg = self.config

        if not cfg.get("template_path") or not os.path.exists(cfg["template_path"]):
            messagebox.showerror("错误", "模板文件不存在")
            return

        data_wbs = {}
        for ds in cfg.get("data_sources", []):
            if not os.path.exists(ds["path"]):
                messagebox.showerror("错误", f"数据源文件不存在: {ds['alias']} ({ds['path']})")
                return
            try:
                data_wbs[ds["alias"]] = load_workbook(ds["path"], data_only=True)
            except Exception as e:
                messagebox.showerror("打开数据源失败", f"{ds['alias']}: {e}")
                return
        try:
            wb = load_workbook(cfg["template_path"])
        except Exception as e:
            for w in data_wbs.values():
                w.close()
            messagebox.showerror("打开模板失败", str(e))
            return

        # ----- 数据写入 -----
        for i, m in enumerate(cfg.get("data_mappings", [])):
            try:
                alias = m.get("source_alias")
                if alias not in data_wbs:
                    raise ValueError(f"数据源别名 '{alias}' 不存在或未加载")
                wb_src = data_wbs[alias]
                if "!" not in m["source_cell"]:
                    raise ValueError("缺少 '!' 分隔符")
                if "!" not in m["target_cell"]:
                    raise ValueError("缺少 '!' 分隔符")
                src_sh, src_cell = m["source_cell"].split("!", 1)
                tgt_sh, tgt_cell = m["target_cell"].split("!", 1)

                if src_sh not in wb_src.sheetnames:
                    raise KeyError(f"数据源[{alias}]中不存在工作表：'{src_sh}'\n可用工作表：{wb_src.sheetnames}")
                if tgt_sh not in wb.sheetnames:
                    raise KeyError(f"模板中不存在工作表：'{tgt_sh}'\n可用工作表：{wb.sheetnames}")

                ws_src = wb_src[src_sh]
                ws_tgt = wb[tgt_sh]
                ws_tgt[tgt_cell].value = ws_src[src_cell].value
            except Exception as e:
                for w in data_wbs.values():
                    w.close()
                wb.close()
                messagebox.showerror("数据写入出错",
                    f"映射 {i+1}:\n源 {m['source_cell']} → 目标 {m['target_cell']}\n错误：{e}")
                return

        # ----- 图片插入（修正偏移） -----
        inserted_count = 0
        skipped_details = []
        for i, m in enumerate(cfg.get("image_mappings", [])):
            try:
                number = m["image_number"]
                folder = Path(m["image_folder"])
                if not folder.exists() or not folder.is_dir():
                    skipped_details.append(f"映射{i+1}: 文件夹不存在或不是文件夹 {folder}")
                    continue
                folder_files = {}
                for f in folder.iterdir():
                    if f.is_file():
                        folder_files[f.name.lower()] = f.name
                if not folder_files:
                    skipped_details.append(f"映射{i+1}: 文件夹为空 {folder}")
                    continue
                img_path = None
                for ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif"]:
                    candidate = f"{number}{ext}".lower()
                    if candidate in folder_files:
                        img_path = str(folder / folder_files[candidate])
                        break
                if not img_path:
                    file_list = "\n".join(sorted(folder_files.values())[:15])
                    skipped_details.append(f"映射{i+1}: 未找到编号 '{number}' 的图片\n文件夹内容(前15):\n{file_list}")
                    continue

                if "!" not in m["target_cell"]:
                    skipped_details.append(f"映射{i+1}: 目标单元格格式错误")
                    continue
                tgt_sh, tgt_cell = m["target_cell"].split("!", 1)
                if tgt_sh not in wb.sheetnames:
                    skipped_details.append(f"映射{i+1}: 模板中不存在工作表 '{tgt_sh}'")
                    continue

                ws_tgt = wb[tgt_sh]
                img = XLImage(img_path)
                img.width = cm_to_px(m["width_cm"])
                img.height = cm_to_px(m["height_cm"])
                ws_tgt.add_image(img, tgt_cell)

                # 只有明确设置为 "custom" 才应用偏移
                if m.get("position") == "custom":
                    off_x = m.get("offset_x_cm", 0)
                    off_y = m.get("offset_y_cm", 0)
                    anchor = img.anchor
                    if isinstance(anchor, (OneCellAnchor, TwoCellAnchor)):
                        from_marker = anchor._from
                        from_marker.colOff = cm_to_emu(off_x)
                        from_marker.rowOff = cm_to_emu(off_y)
                    else:
                        skipped_details.append(f"映射{i+1}: 无法设置偏移，锚点类型为 {type(anchor).__name__}")

                inserted_count += 1

            except Exception as e:
                for w in data_wbs.values():
                    w.close()
                wb.close()
                messagebox.showerror("图片插入出错",
                    f"图片映射 {i+1}:\n编号 {m['image_number']}，目标 {m['target_cell']}\n错误：{e}")
                return

        for w in data_wbs.values():
            w.close()

        # 输出路径
        tpl_path = Path(cfg["template_path"])
        suffix = cfg.get("output_suffix", "_已更新")
        out_dir = cfg.get("output_dir", "")
        if out_dir and os.path.isdir(out_dir):
            out_path = Path(out_dir) / f"{tpl_path.stem}{suffix}.xlsx"
        else:
            out_path = tpl_path.parent / f"{tpl_path.stem}{suffix}.xlsx"

        try:
            wb.save(str(out_path))
        except Exception as e:
            wb.close()
            messagebox.showerror("保存失败", str(e))
            return
        wb.close()

        summary = f"数据源: {len(data_wbs)} 个\n"
        summary += f"数据映射: {len(cfg.get('data_mappings', []))} 条\n"
        summary += f"图片映射: {len(cfg.get('image_mappings', []))} 条\n"
        summary += f"成功插入图片: {inserted_count} 张\n"
        if skipped_details:
            summary += "\n未插入图片原因:\n" + "\n\n".join(skipped_details)
        messagebox.showinfo("执行结果", summary)

        try:
            os.startfile(out_path)
        except Exception:
            pass

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
