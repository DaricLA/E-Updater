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
except ImportError as e:
    import traceback
    err_msg = traceback.format_exc()
    messagebox.showerror("库导入失败", f"缺失必要组件，请反馈以下信息:\n{err_msg}")
    sys.exit(1)

# 单位转换
def cm_to_px(cm_val):
    """厘米转像素（96 DPI）"""
    return int(cm_val * 37.795)

def col_width_to_px(col_width):
    """Excel列宽（字符数）近似转像素"""
    return int(col_width * 7.0)

def row_height_to_px(row_height):
    """行高（磅）转像素（1磅≈1.333像素）"""
    return int(row_height * 4 / 3)

def px_to_emu(px):
    """像素转EMU（1像素=9525 EMU）"""
    return int(px * 9525)

CONFIG_FILE = "update_config.json"

def save_config(data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {
            "data_sources": [],
            "template_path": "",
            "output_suffix": "_已更新",
            "data_mappings": [],
            "image_mappings": []
        }
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel报告一键更新器")
        self.root.geometry("1000x750")
        self.config = load_config()

        # ---------- 基础路径 ----------
        frm_path = ttk.LabelFrame(root, text="基础设置")
        frm_path.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(frm_path, text="模板路径:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.tpl_path_var = tk.StringVar(value=self.config.get("template_path", ""))
        ttk.Entry(frm_path, textvariable=self.tpl_path_var, width=80).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(frm_path, text="浏览", command=self.browse_tpl).grid(row=0, column=2, padx=5)

        ttk.Label(frm_path, text="输出后缀:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.suffix_var = tk.StringVar(value=self.config.get("output_suffix", "_已更新"))
        ttk.Entry(frm_path, textvariable=self.suffix_var, width=20).grid(row=1, column=1, sticky=tk.W, padx=5)

        # ---------- 映射管理 ----------
        nb = ttk.Notebook(root)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ---- 数据源管理页 ----
        frm_src = ttk.Frame(nb)
        nb.add(frm_src, text="数据源管理")
        self.src_tree = ttk.Treeview(frm_src, columns=("alias", "path"), show="headings", height=5)
        self.src_tree.heading("alias", text="别名")
        self.src_tree.heading("path", text="文件路径")
        self.src_tree.column("alias", width=150)
        self.src_tree.column("path", width=400)
        self.src_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        btn_frm_src = ttk.Frame(frm_src)
        btn_frm_src.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(btn_frm_src, text="添加数据源", command=self.add_data_source).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_src, text="删除选中", command=self.delete_data_source).pack(side=tk.LEFT, padx=5)

        # ---- 数据映射页 ----
        frm_data = ttk.Frame(nb)
        nb.add(frm_data, text="数据映射")
        self.data_tree = ttk.Treeview(frm_data, columns=("source_alias", "source_cell", "target_cell"), show="headings", height=8)
        self.data_tree.heading("source_alias", text="数据源")
        self.data_tree.heading("source_cell", text="源单元格")
        self.data_tree.heading("target_cell", text="目标单元格")
        self.data_tree.column("source_alias", width=120)
        self.data_tree.column("source_cell", width=150)
        self.data_tree.column("target_cell", width=150)
        self.data_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        btn_frm_data = ttk.Frame(frm_data)
        btn_frm_data.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(btn_frm_data, text="添加", command=self.add_data_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_data, text="编辑选中", command=self.edit_data_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_data, text="删除选中", command=lambda: self.delete_selected(self.data_tree, "data")).pack(side=tk.LEFT, padx=5)

        # ---- 图片映射页 ----
        frm_img = ttk.Frame(nb)
        nb.add(frm_img, text="图片映射")
        self.img_tree = ttk.Treeview(frm_img, columns=("number", "folder", "target", "width", "height", "pos"), show="headings", height=8)
        self.img_tree.heading("number", text="编号")
        self.img_tree.heading("folder", text="文件夹")
        self.img_tree.heading("target", text="目标单元格")
        self.img_tree.heading("width", text="宽(cm)")
        self.img_tree.heading("height", text="高(cm)")
        self.img_tree.heading("pos", text="位置")
        self.img_tree.column("number", width=80)
        self.img_tree.column("folder", width=200)
        self.img_tree.column("target", width=120)
        self.img_tree.column("width", width=60)
        self.img_tree.column("height", width=60)
        self.img_tree.column("pos", width=80)
        self.img_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        btn_frm_img = ttk.Frame(frm_img)
        btn_frm_img.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(btn_frm_img, text="添加", command=self.add_image_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_img, text="编辑选中", command=self.edit_image_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm_img, text="删除选中", command=lambda: self.delete_selected(self.img_tree, "image")).pack(side=tk.LEFT, padx=5)

        ttk.Button(root, text="一键更新报告", command=self.run_update).pack(pady=10)

        self.refresh_all()

    # ---------- 数据源操作 ----------
    def add_data_source(self):
        popup = tk.Toplevel(self.root)
        popup.title("添加数据源")
        ttk.Label(popup, text="别名 (如 销售数据):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        alias_entry = ttk.Entry(popup, width=30)
        alias_entry.grid(row=0, column=1, padx=5)
        ttk.Label(popup, text="文件路径:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        path_var = tk.StringVar()
        ttk.Entry(popup, textvariable=path_var, width=30).grid(row=1, column=1, padx=5)
        ttk.Button(popup, text="浏览", command=lambda: path_var.set(filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")]))).grid(row=1, column=2, padx=5)

        def save():
            alias = alias_entry.get().strip()
            path = path_var.get().strip()
            if not alias or not path:
                messagebox.showwarning("输入不完整", "别名和路径不能为空")
                return
            self.config.setdefault("data_sources", []).append({"alias": alias, "path": path})
            self.refresh_src_tree()
            popup.destroy()
        ttk.Button(popup, text="确定", command=save).grid(row=2, column=0, columnspan=3, pady=10)

    def delete_data_source(self):
        selected = self.src_tree.selection()
        if not selected:
            return
        idx = self.src_tree.index(selected[0])
        alias = self.config["data_sources"][idx]["alias"]
        # 检查是否有映射引用了该数据源
        for m in self.config.get("data_mappings", []):
            if m.get("source_alias") == alias:
                messagebox.showwarning("无法删除", f"数据源 '{alias}' 正被数据映射使用，请先删除相关映射。")
                return
        del self.config["data_sources"][idx]
        self.refresh_src_tree()

    def refresh_src_tree(self):
        for i in self.src_tree.get_children():
            self.src_tree.delete(i)
        for ds in self.config.get("data_sources", []):
            self.src_tree.insert("", tk.END, values=(ds["alias"], ds["path"]))

    # ---------- 数据映射操作 ----------
    def get_data_source_aliases(self):
        return [ds["alias"] for ds in self.config.get("data_sources", [])]

    def add_data_mapping(self):
        if not self.config.get("data_sources"):
            messagebox.showinfo("提示", "请先添加数据源")
            return
        self._data_dialog(None)

    def edit_data_mapping(self):
        selected = self.data_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一条映射")
            return
        idx = self.data_tree.index(selected[0])
        item = self.config["data_mappings"][idx]
        self._data_dialog(idx, item)

    def _data_dialog(self, edit_idx, item=None):
        popup = tk.Toplevel(self.root)
        popup.title("编辑数据映射" if item else "添加数据映射")
        aliases = self.get_data_source_aliases()
        if not aliases:
            messagebox.showinfo("提示", "没有可用数据源", parent=popup)
            popup.destroy()
            return
        ttk.Label(popup, text="数据源:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        alias_var = tk.StringVar()
        alias_combo = ttk.Combobox(popup, textvariable=alias_var, values=aliases, state="readonly", width=27)
        alias_combo.grid(row=0, column=1, padx=5)
        ttk.Label(popup, text="源单元格 (如 Sheet1!B2):").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        src_entry = ttk.Entry(popup, width=30)
        src_entry.grid(row=1, column=1, padx=5)
        ttk.Label(popup, text="目标单元格 (如 Sheet1!D5):").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        tgt_entry = ttk.Entry(popup, width=30)
        tgt_entry.grid(row=2, column=1, padx=5)
        if item:
            alias_var.set(item["source_alias"])
            src_entry.insert(0, item["source_cell"])
            tgt_entry.insert(0, item["target_cell"])
        else:
            alias_var.set(aliases[0])

        def save():
            alias = alias_var.get().strip()
            src = src_entry.get().strip()
            tgt = tgt_entry.get().strip()
            if not alias or not src or not tgt:
                messagebox.showwarning("输入不完整", "所有字段必填")
                return
            new_map = {"source_alias": alias, "source_cell": src, "target_cell": tgt}
            if edit_idx is not None:
                self.config["data_mappings"][edit_idx] = new_map
            else:
                self.config.setdefault("data_mappings", []).append(new_map)
            self.refresh_data_tree()
            popup.destroy()
        ttk.Button(popup, text="确定", command=save).grid(row=3, column=0, columnspan=2, pady=10)

    def refresh_data_tree(self):
        for i in self.data_tree.get_children():
            self.data_tree.delete(i)
        for m in self.config.get("data_mappings", []):
            self.data_tree.insert("", tk.END, values=(m.get("source_alias", ""), m["source_cell"], m["target_cell"]))

    # ---------- 图片映射操作 ----------
    def add_image_mapping(self):
        self._image_dialog(None)

    def edit_image_mapping(self):
        selected = self.img_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一条映射")
            return
        idx = self.img_tree.index(selected[0])
        item = self.config["image_mappings"][idx]
        self._image_dialog(idx, item)

    def _image_dialog(self, edit_idx, item=None):
        popup = tk.Toplevel(self.root)
        popup.title("编辑图片映射" if item else "添加图片映射")
        ttk.Label(popup, text="图片编号:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        num_entry = ttk.Entry(popup, width=30)
        num_entry.grid(row=0, column=1, padx=5)
        ttk.Label(popup, text="图片文件夹:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        folder_var = tk.StringVar()
        folder_entry = ttk.Entry(popup, textvariable=folder_var, width=30)
        folder_entry.grid(row=1, column=1, padx=5)
        ttk.Button(popup, text="浏览", command=lambda: folder_var.set(filedialog.askdirectory())).grid(row=1, column=2, padx=5)
        ttk.Label(popup, text="目标单元格:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        tgt_entry = ttk.Entry(popup, width=30)
        tgt_entry.grid(row=2, column=1, padx=5)
        ttk.Label(popup, text="宽度 (cm):").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        w_entry = ttk.Entry(popup, width=10)
        w_entry.grid(row=3, column=1, sticky=tk.W, padx=5)
        ttk.Label(popup, text="高度 (cm):").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        h_entry = ttk.Entry(popup, width=10)
        h_entry.grid(row=4, column=1, sticky=tk.W, padx=5)
        ttk.Label(popup, text="位置:").grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
        pos_var = tk.StringVar(value="top-left")
        ttk.Combobox(popup, textvariable=pos_var, values=["top-left", "center"], state="readonly", width=10).grid(row=5, column=1, sticky=tk.W, padx=5)

        if item:
            num_entry.insert(0, item["image_number"])
            folder_var.set(item["image_folder"])
            tgt_entry.insert(0, item["target_cell"])
            w_entry.insert(0, str(item["width_cm"]))
            h_entry.insert(0, str(item["height_cm"]))
            pos_var.set(item.get("position", "top-left"))
        else:
            w_entry.insert(0, "3.5")
            h_entry.insert(0, "2.8")

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
            if not num or not folder or not tgt:
                messagebox.showwarning("输入不完整", "所有字段必填")
                return
            new_map = {
                "image_number": num,
                "image_folder": folder,
                "target_cell": tgt,
                "width_cm": w,
                "height_cm": h,
                "position": pos_var.get()
            }
            if edit_idx is not None:
                self.config["image_mappings"][edit_idx] = new_map
            else:
                self.config.setdefault("image_mappings", []).append(new_map)
            self.refresh_image_tree()
            popup.destroy()
        ttk.Button(popup, text="确定", command=save).grid(row=6, column=0, columnspan=3, pady=10)

    def refresh_image_tree(self):
        for i in self.img_tree.get_children():
            self.img_tree.delete(i)
        for m in self.config.get("image_mappings", []):
            self.img_tree.insert("", tk.END, values=(
                m["image_number"],
                m["image_folder"],
                m["target_cell"],
                m["width_cm"],
                m["height_cm"],
                m.get("position", "top-left")
            ))

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

    def refresh_all(self):
        self.refresh_src_tree()
        self.refresh_data_tree()
        self.refresh_image_tree()

    def save_current_config(self):
        self.config["template_path"] = self.tpl_path_var.get()
        self.config["output_suffix"] = self.suffix_var.get()
        save_config(self.config)

    # ---------- 核心执行 ----------
    def run_update(self):
        self.save_current_config()
        cfg = self.config

        if not cfg.get("template_path") or not os.path.exists(cfg["template_path"]):
            messagebox.showerror("错误", "模板文件不存在")
            return
        # 验证数据源
        ds_paths = {}
        for ds in cfg.get("data_sources", []):
            if not os.path.exists(ds["path"]):
                messagebox.showerror("错误", f"数据源 '{ds['alias']}' 文件不存在:\n{ds['path']}")
                return
            ds_paths[ds["alias"]] = ds["path"]
        if not ds_paths:
            messagebox.showerror("错误", "未配置任何数据源")
            return

        # 打开所有数据源工作簿（缓存）
        wb_src_cache = {}
        try:
            for alias, path in ds_paths.items():
                wb_src_cache[alias] = load_workbook(path, data_only=True)
        except Exception as e:
            for wb in wb_src_cache.values():
                wb.close()
            messagebox.showerror("打开数据源失败", str(e))
            return

        try:
            wb_tpl = load_workbook(cfg["template_path"])
        except Exception as e:
            for wb in wb_src_cache.values():
                wb.close()
            messagebox.showerror("打开模板失败", str(e))
            return

        try:
            # ----- 数据写入 -----
            for i, m in enumerate(cfg.get("data_mappings", [])):
                alias = m.get("source_alias")
                if alias not in wb_src_cache:
                    raise KeyError(f"数据源 '{alias}' 未加载")
                wb_src = wb_src_cache[alias]
                try:
                    if "!" not in m["source_cell"] or "!" not in m["target_cell"]:
                        raise ValueError("缺少 '!' 分隔符")
                    src_sh, src_cell = m["source_cell"].split("!", 1)
                    tgt_sh, tgt_cell = m["target_cell"].split("!", 1)

                    if src_sh not in wb_src.sheetnames:
                        raise KeyError(f"数据源 '{alias}' 中不存在工作表 '{src_sh}'")
                    if tgt_sh not in wb_tpl.sheetnames:
                        raise KeyError(f"模板中不存在工作表 '{tgt_sh}'")

                    ws_src = wb_src[src_sh]
                    ws_tgt = wb_tpl[tgt_sh]
                    ws_tgt[tgt_cell].value = ws_src[src_cell].value
                except Exception as e:
                    raise Exception(f"映射 {i+1} (源 {m['source_cell']} → 目标 {m['target_cell']}) 出错: {e}")

            # ----- 图片插入 -----
            inserted = 0
            skipped = []
            for i, m in enumerate(cfg.get("image_mappings", [])):
                try:
                    number = m["image_number"]
                    folder = Path(m["image_folder"])
                    if not folder.is_dir():
                        skipped.append(f"映射{i+1}: 文件夹不存在 {folder}")
                        continue

                    folder_files = {f.name.lower(): f.name for f in folder.iterdir() if f.is_file()}
                    if not folder_files:
                        skipped.append(f"映射{i+1}: 文件夹为空")
                        continue

                    img_path = None
                    for ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif"]:
                        cand = f"{number}{ext}".lower()
                        if cand in folder_files:
                            img_path = str(folder / folder_files[cand])
                            break
                    if not img_path:
                        file_list = "\n".join(sorted(folder_files.values())[:15])
                        skipped.append(f"映射{i+1}: 未找到编号 '{number}'\n文件夹内容:\n{file_list}")
                        continue

                    if "!" not in m["target_cell"]:
                        skipped.append(f"映射{i+1}: 目标单元格格式错误")
                        continue
                    tgt_sh, tgt_cell = m["target_cell"].split("!", 1)
                    if tgt_sh not in wb_tpl.sheetnames:
                        skipped.append(f"映射{i+1}: 模板中不存在工作表 '{tgt_sh}'")
                        continue

                    ws_tgt = wb_tpl[tgt_sh]
                    img = XLImage(img_path)
                    img.width = cm_to_px(m["width_cm"])
                    img.height = cm_to_px(m["height_cm"])
                    img.anchor = tgt_cell  # 默认锚点

                    # 居中处理
                    if m.get("position") == "center":
                        # 获取单元格尺寸（近似像素）
                        col_letter = ''.join(filter(str.isalpha, tgt_cell))
                        if col_letter:
                            col_idx = ws_tgt.column_dimensions[col_letter].width or 8.43
                            col_px = col_width_to_px(col_idx)
                        else:
                            col_px = 60
                        try:
                            row_num = int(''.join(filter(str.isdigit, tgt_cell)))
                            row_height = ws_tgt.row_dimensions[row_num].height or 15
                            row_px = row_height_to_px(row_height)
                        except:
                            row_px = 20

                        img_w_px = img.width
                        img_h_px = img.height
                        offset_x = max(0, (col_px - img_w_px) // 2)
                        offset_y = max(0, (row_px - img_h_px) // 2)
                        img.anchor = openpyxl.drawing.spreadsheet_drawing.AnchorMarker(
                            col=ws_tgt[tgt_cell].column - 1,  # 0-based
                            colOff=px_to_emu(offset_x),
                            row=ws_tgt[tgt_cell].row - 1,
                            rowOff=px_to_emu(offset_y)
                        )

                    ws_tgt.add_image(img, tgt_cell if m.get("position") != "center" else None)
                    inserted += 1

                except Exception as e:
                    raise Exception(f"图片映射 {i+1} (编号 {m['image_number']}) 出错: {e}")

            # 保存
            tpl_path = Path(cfg["template_path"])
            suffix = cfg.get("output_suffix", "_已更新")
            out_path = tpl_path.parent / f"{tpl_path.stem}{suffix}.xlsx"
            wb_tpl.save(str(out_path))

            summary = f"数据映射: {len(cfg.get('data_mappings', []))} 条\n图片: 成功 {inserted} 张"
            if skipped:
                summary += "\n\n跳过详情:\n" + "\n\n".join(skipped)
            messagebox.showinfo("完成", summary)
            os.startfile(out_path)

        finally:
            for wb in wb_src_cache.values():
                wb.close()
            wb_tpl.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
