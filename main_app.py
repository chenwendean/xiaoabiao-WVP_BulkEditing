#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主界面 - 通道查询、编辑、批量修改、Excel导入导出
v7.0 - 新增设备列表选择，登录后自动查询全部国标设备

Copyright (C) 2025 Xiaoabiao

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import hashlib
import requests
import threading
from datetime import datetime
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

# 尝试导入 openpyxl（非必需，但导出/导入需要）
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

from progress_dialog import ProgressDialog


class MainApplication:
    def __init__(self, root, token, user_info, host):
        self.root = root
        self.access_token = token
        self.login_user = user_info
        self.server_host = tk.StringVar(value=host)

        self.root.title("WVP通道管理工具 v7.0 by Xiaoabiao")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 650)
        self.root.resizable(True, True)

        # 设备列表相关
        self.all_devices = []           # 全部设备列表
        self.selected_device_id = None  # 当前选中的设备ID
        self.selected_device_name = ""
        self.selected_device_status = ""

        self.status_text = tk.StringVar(value=f"已登录: {user_info.get('username', '')}")
        self.page_num = 1
        self.page_size = 50000
        self.total_channels = 0
        self.all_channels = []
        self.item_to_channel = {}

        self.edit_entry = None
        self.edit_item = None
        self.edit_column = None

        self.check_vars = {}
        self.select_all_var = tk.BooleanVar(value=False)

        # 设备列表复选框
        self.dev_check_vars = {}

        self.progress = None          # 进度弹窗实例
        self.max_workers = 8          # 并发线程数（可自定义）

        self.setup_styles()
        self.build_ui()
        # 登录后自动查询设备列表
        self.root.after(500, self.do_query_devices)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=26, font=("Microsoft YaHei", 9))
        style.configure("Treeview.Heading", font=("Microsoft YaHei", 10, "bold"))
        style.configure("TButton", font=("Microsoft YaHei", 9))
        style.configure("TLabel", font=("Microsoft YaHei", 9))
        style.configure("TEntry", font=("Microsoft YaHei", 9))

    def build_ui(self):
        # 顶栏
        top_bar = ttk.Frame(self.root)
        top_bar.pack(fill=tk.X, padx=10, pady=(5, 0))
        ttk.Label(top_bar, text=f"  {self.login_user.get('username', '')}",
                  font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT)
        self.thread_btn = ttk.Button(top_bar, text=f"并发: {self.max_workers}",
                                      command=self.set_max_workers, width=10)
        self.thread_btn.pack(side=tk.RIGHT, padx=5)
        ttk.Button(top_bar, text="退出登录", command=self.logout).pack(side=tk.RIGHT, padx=5)
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X, padx=10, pady=(2, 0))

        # ========== 主内容区域：左侧设备列表 + 右侧通道列表 ==========
        main_pw = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pw.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ----- 左侧：设备列表 -----
        left_frame = ttk.LabelFrame(main_pw, text=" 国标设备列表 ", padding=5)
        main_pw.add(left_frame, weight=1)

        # 设备列表顶栏（按钮横排）
        dev_toolbar = ttk.Frame(left_frame)
        dev_toolbar.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(dev_toolbar, text=" 刷新设备", command=self.do_query_devices, width=12).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(dev_toolbar, text=" 导出勾选通道", command=self.export_selected_devices, width=14).pack(side=tk.LEFT, padx=3)
        ttk.Button(dev_toolbar, text=" 导入通道Excel", command=self.import_device_excel, width=14).pack(side=tk.LEFT, padx=3)
        self.dev_count_var = tk.StringVar(value="设备数: -")
        ttk.Label(dev_toolbar, textvariable=self.dev_count_var, foreground="gray").pack(side=tk.RIGHT, padx=5)

        # 设备列表（带复选框 + 状态/型号/厂家）
        dev_tree_frame = ttk.Frame(left_frame)
        dev_tree_frame.pack(fill=tk.BOTH, expand=True)
        dev_tree_frame.grid_rowconfigure(0, weight=1)
        dev_tree_frame.grid_columnconfigure(0, weight=1)

        dev_columns = ("☐", "设备ID", "名称", "在线", "型号", "厂家")
        self.dev_tree = ttk.Treeview(dev_tree_frame, columns=dev_columns, show="headings",
                                      selectmode="browse", height=15)
        dev_col_widths = {"☐": 30, "设备ID": 160, "名称": 120, "在线": 45, "型号": 80, "厂家": 80}
        for col in dev_columns:
            w = dev_col_widths.get(col, 80)
            self.dev_tree.heading(col, text=col)
            self.dev_tree.column(col, width=w, anchor=tk.CENTER, minwidth=w)
        # 点击 ☐ 表头 = 全选/取消全选
        self.dev_tree.heading("#1", command=self.toggle_dev_select_all)

        dev_vsb = ttk.Scrollbar(dev_tree_frame, orient=tk.VERTICAL, command=self.dev_tree.yview)
        self.dev_tree.configure(yscrollcommand=dev_vsb.set)
        self.dev_tree.grid(row=0, column=0, sticky="nsew")
        dev_vsb.grid(row=0, column=1, sticky="ns")

        self.dev_tree.bind("<Double-1>", self.on_device_double_click)
        self.dev_tree.bind("<Button-1>", self.on_dev_checkbox_click)

        # 底部状态栏
        self.statusbar_var = tk.StringVar(value="就绪")
        statusbar = ttk.Label(self.root, textvariable=self.statusbar_var, relief=tk.SUNKEN,
                              anchor=tk.W, font=("Microsoft YaHei", 8), padding=(5, 2))
        statusbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.update_ui_state()