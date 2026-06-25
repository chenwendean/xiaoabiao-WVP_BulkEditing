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

    # ---------- UI 状态 ----------
    def update_ui_state(self):
        pass

    def set_statusbar(self, msg):
        self.statusbar_var.set(f"{msg}  [{datetime.now().strftime('%H:%M:%S')}]")

    def _show_progress(self, title, total):
        if self.progress:
            try:
                self.progress.close()
            except:
                pass
        self.progress = ProgressDialog(self.root, title, total)

    def _close_progress(self):
        if self.progress:
            try:
                self.progress.close()
            except:
                pass
            self.progress = None

    def set_max_workers(self):
        val = simpledialog.askinteger("并发设置",
            f"当前并发线程数: {self.max_workers}\n请输入新的数值（1-32）:",
            initialvalue=self.max_workers, minvalue=1, maxvalue=32, parent=self.root)
        if val:
            self.max_workers = val
            self.thread_btn.configure(text=f"并发: {self.max_workers}")
            self.set_statusbar(f"并发线程数已设为 {self.max_workers}")

    # ---------- 并发分页查询通道 ----------
    def _concurrent_channel_query(self, device_id, page_size=100, progress=None):
        host = self.server_host.get().strip().rstrip('/')
        token = self.access_token
        headers = {"Accept": "*/*", "access-token": token}
        url = f"{host}/api/device/query/devices/{device_id}/channels"
        base_params = {"query": "", "cameraQuery": "", "nvrQuery": ""}

        if progress:
            progress(0, 1, "正在获取总页数...")

        try:
            r = requests.get(url, headers=headers, params={
                **base_params, "page": 1, "count": page_size}, timeout=30)
            if r.status_code != 200:
                return [], 0
            data = r.json()
            if data.get("code") != 0:
                return [], 0
            inner = data.get("data", {})
            first_batch = inner.get("list", [])
            total = inner.get("total", 0) or len(first_batch)
        except Exception:
            return [], 0

        if total <= page_size:
            if progress:
                progress(1, 1, f"查询完成，共 {total} 个通道")
            return first_batch, total

        total_pages = (total + page_size - 1) // page_size
        pages_list = {1: first_batch}
        page_range = range(2, total_pages + 1)

        if progress:
            progress(0, total_pages, f"共 {total_pages} 页，正在并发查询...")

        def fetch_page(pg):
            try:
                resp = requests.get(url, headers=headers, params={
                    **base_params, "page": pg, "count": page_size}, timeout=30)
                if resp.status_code == 200:
                    d = resp.json()
                    if d.get("code") == 0:
                        return pg, d.get("data", {}).get("list", [])
            except Exception:
                pass
            return pg, []

        fetched = [len(first_batch)]
        done = [1]
        with ThreadPoolExecutor(max_workers=self.max_workers) as ex:
            fut_map = {ex.submit(fetch_page, p): p for p in page_range}
            for fut in as_completed(fut_map):
                pg, lst = fut.result()
                pages_list[pg] = lst
                fetched[0] += len(lst)
                done[0] += 1
                if progress:
                    progress(fetched[0], total,
                        f"正在查询 ({fetched[0]}/{total} 个通道)...")

        all_channels = []
        for pg in sorted(pages_list):
            all_channels.extend(pages_list[pg])

        if progress:
            progress(total, total, f"查询完成，共 {total} 个通道")
        return all_channels, total

    # ========== 设备列表相关 ==========

    def do_query_devices(self):
        if not self.access_token:
            messagebox.showwarning("警告", "请先登录")
            return
        self.set_statusbar("正在查询全部国标设备...")
        self.dev_count_var.set("查询中...")
        host = self.server_host.get().strip().rstrip('/')
        token = self.access_token

        def task():
            nonlocal host, token
            try:
                url = f"{host}/api/device/query/devices"
                headers = {
                    "Accept": "*/*",
                    "access-token": token,
                    "Content-Type": "application/x-www-form-urlencoded"
                }
                all_devices = []
                page = 1
                count = 100
                total = 0

                while True:
                    params = {
                        "page": page,
                        "count": count,
                        "query": "",
                        "status": ""
                    }
                    resp = requests.get(url, headers=headers, params=params, timeout=30)

                    if resp.status_code == 200:
                        data = resp.json()
                        if data.get("code") != 0:
                            self.root.after(0, lambda m=data.get("msg", "未知错误"): self._devices_fail(m))
                            return

                        inner = data.get("data", {})
                        devices = inner.get("list", [])
                        total = inner.get("total", 0)
                        all_devices.extend(devices)

                        self.root.after(0, lambda p=page: self.set_statusbar(f"正在加载设备第 {p} 页..."))

                        if len(devices) < count or len(all_devices) >= total:
                            break
                        page += 1
                    else:
                        msg = resp.json().get("msg", f"HTTP {resp.status_code}")
                        self.root.after(0, lambda m=msg: self._devices_fail(m))
                        return

                self.all_devices = all_devices
                self.root.after(0, lambda: self._devices_success(all_devices, total))

            except Exception as e:
                self.root.after(0, lambda e=e: self._devices_fail(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _devices_success(self, devices, total):
        self.dev_tree.delete(*self.dev_tree.get_children())
        self.dev_check_vars.clear()
        self.dev_count_var.set(f"设备数: {total}")

        for dev in devices:
            device_id = dev.get("deviceId", "")
            name = dev.get("name", "")
            online = "ON" if dev.get("onLine") else "OFF"
            model = dev.get("model", dev.get("gbModel", ""))
            mfr = dev.get("manufacturer", dev.get("gbManufacturer", ""))
            var = tk.BooleanVar(value=False)
            values = ("☐", device_id, name, online, model, mfr)
            item = self.dev_tree.insert("", tk.END, values=values)
            self.dev_check_vars[item] = var

        self.set_statusbar(f"查询成功 - 共 {total} 个国标设备")
        self.update_ui_state()

    def _update_dev_header(self):
        if not self.dev_check_vars:
            return
        all_checked = all(var.get() for var in self.dev_check_vars.values())
        any_checked = any(var.get() for var in self.dev_check_vars.values())
        if all_checked:
            self.dev_tree.heading("#1", text="☑")
        elif any_checked:
            self.dev_tree.heading("#1", text="☐")
        else:
            self.dev_tree.heading("#1", text="☐")

    def toggle_dev_select_all(self):
        if not self.dev_check_vars:
            return
        all_checked = all(var.get() for var in self.dev_check_vars.values())
        new_state = not all_checked
        for item, var in self.dev_check_vars.items():
            var.set(new_state)
            self.dev_tree.set(item, "#1", "☑" if new_state else "☐")
        self._update_dev_header()