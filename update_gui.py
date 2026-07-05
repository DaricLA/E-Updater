def __init__(self, root):
    self.root = root
    self.root.title("Excel报告一键更新器")
    self.root.geometry("1050x800")
    self.config = load_config()
    self.current_config_name = CONFIG_FILE

    # ---------- 提前创建所有路径变量（避免后续状态栏调用时未定义） ----------
    self.tpl_path_var = tk.StringVar(value=self.config.get("template_path", ""))
    self.out_dir_var = tk.StringVar(value=self.config.get("output_dir", ""))
    self.suffix_var = tk.StringVar(value=self.config.get("output_suffix", "_已更新"))

    # ---------- 顶部状态栏（先创建界面，变量已存在） ----------
    self.status_frame = ttk.Frame(root, relief=tk.RAISED, borderwidth=2)
    self.status_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
    self.status_var = tk.StringVar()
    ttk.Label(self.status_frame, textvariable=self.status_var,
              background="#D9EAF7", font=("微软雅黑", 10, "bold")).pack(fill=tk.X, padx=10, pady=5)
    self.update_status_bar()

    # ---------- 模板与输出设置（界面位于状态栏下方） ----------
    frm_tpl = ttk.LabelFrame(root, text="模板与输出设置")
    frm_tpl.pack(fill=tk.X, padx=10, pady=5)
    # 模板文件
    ttk.Label(frm_tpl, text="模板文件:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
    ttk.Entry(frm_tpl, textvariable=self.tpl_path_var, width=70).grid(row=0, column=1, padx=5, pady=2)
    ttk.Button(frm_tpl, text="浏览", command=self.browse_tpl).grid(row=0, column=2, padx=5)
    # 输出文件夹
    ttk.Label(frm_tpl, text="输出文件夹:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
    ttk.Entry(frm_tpl, textvariable=self.out_dir_var, width=70).grid(row=1, column=1, padx=5, pady=2)
    ttk.Button(frm_tpl, text="浏览", command=self.browse_out_dir).grid(row=1, column=2, padx=5)
    # 输出后缀
    ttk.Label(frm_tpl, text="输出文件后缀:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
    ttk.Entry(frm_tpl, textvariable=self.suffix_var, width=15).grid(row=2, column=1, sticky=tk.W, padx=5)
    # 提示
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
