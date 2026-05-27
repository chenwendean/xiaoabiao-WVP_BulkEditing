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
        main_pw.add(left_frame, weight=1)  # 较小的权重

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
        # UI已精简，暂无动态控件需要更新
        pass

    def set_statusbar(self, msg):
        self.statusbar_var.set(f"{msg}  [{datetime.now().strftime('%H:%M:%S')}]")

    def _show_progress(self, title, total):
        """显示进度弹窗"""
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
        """弹窗修改并发线程数"""
        val = simpledialog.askinteger("并发设置",
            f"当前并发线程数: {self.max_workers}\n请输入新的数值（1-32）:",
            initialvalue=self.max_workers, minvalue=1, maxvalue=32, parent=self.root)
        if val:
            self.max_workers = val
            self.thread_btn.configure(text=f"并发: {self.max_workers}")
            self.set_statusbar(f"并发线程数已设为 {self.max_workers}")

    # ---------- 并发分页查询通道 ----------
    def _concurrent_channel_query(self, device_id, page_size=100, progress=None):
        """并发分页查询设备通道，返回 (channels, total)
           progress: 可选回调 (current, total, text)
        """
        host = self.server_host.get().strip().rstrip('/')
        token = self.access_token
        headers = {"Accept": "*/*", "access-token": token}
        url = f"{host}/api/device/query/devices/{device_id}/channels"
        base_params = {"query": "", "cameraQuery": "", "nvrQuery": ""}

        if progress:
            progress(0, 1, "正在获取总页数...")

        # 先查第一页获取总数
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
        """查询全部国标设备"""
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
        """更新设备列表 ☐ 表头状态"""
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
        """点击 ☐ 表头 = 全选/取消全选"""
        if not self.dev_check_vars:
            return
        # 如果全部已勾选则取消全选，否则全选
        all_checked = all(var.get() for var in self.dev_check_vars.values())
        new_state = not all_checked
        for item, var in self.dev_check_vars.items():
            var.set(new_state)
            self.dev_tree.set(item, "#1", "☑" if new_state else "☐")
        self._update_dev_header()

    def on_dev_checkbox_click(self, event):
        """设备列表复选框点击（仅处理数据行的点击，表头点击由 toggle_dev_select_all 处理）"""
        region = self.dev_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.dev_tree.identify_column(event.x)
        if col != "#1":
            return
        item = self.dev_tree.identify_row(event.y)
        if not item:
            return
        var = self.dev_check_vars.get(item)
        if var:
            var.set(not var.get())
            self.dev_tree.set(item, "#1", "☑" if var.get() else "☐")
            self._update_dev_header()
        return "break"

    def _devices_fail(self, msg):
        self.dev_count_var.set("查询失败")
        self.set_statusbar(f"设备列表查询失败: {msg}")
        messagebox.showerror("查询失败", msg)

    def on_device_double_click(self, event):
        """设备列表双击 - 弹出新窗口显示通道"""
        item = self.dev_tree.identify_row(event.y)
        if not item:
            return
        values = self.dev_tree.item(item, "values")
        if not values:
            return
        device_id = values[1]
        device_name = values[2]
        self.set_statusbar(f"正在加载设备 {device_name} 的通道...")
        threading.Thread(target=self._open_device_window,
                         args=(device_id, device_name), daemon=True).start()

    def _open_device_window(self, device_id, device_name):
        """后台查询通道并在新窗口展示（并发分页+进度条）"""
        host = self.server_host.get().strip().rstrip('/')
        token = self.access_token
        try:
            headers = {"Accept": "*/*", "access-token": token}
            self.root.after(0, lambda: self._show_progress(
                f"正在查询 {device_name}", 1))

            def on_progress(c, t, text):
                self.root.after(0, lambda: self.progress.update(c, t, text))

            channels, total = self._concurrent_channel_query(
                device_id, progress=on_progress)
            # 进度条继续显示合并全局通道数据
            merge_done = [0]
            merge_total = len(channels)

            def merge_one(dc):
                dc_id = dc.get("id") or dc.get("gbId")
                if not dc_id:
                    return dc
                try:
                    one_resp = requests.get(f"{host}/api/common/channel/one",
                                            headers=headers,
                                            params={"id": dc_id}, timeout=15)
                    if one_resp.status_code == 200:
                        od = one_resp.json()
                        if od.get("code") == 0:
                            gc = od.get("data")
                            if gc and isinstance(gc, dict):
                                for fld in ("gbManufacturer", "gbLongitude", "gbLatitude",
                                            "gbName", "gbCivilCode"):
                                    if fld in gc and gc[fld] is not None:
                                        dc[fld] = gc[fld]
                except Exception:
                    pass
                return dc

            def merge_with_progress(dc):
                result = merge_one(dc)
                merge_done[0] += 1
                if merge_done[0] % 50 == 0:
                    on_progress(merge_done[0], merge_total,
                        f"合并通道数据 ({merge_done[0]}/{merge_total})...")
                return result

            with ThreadPoolExecutor(max_workers=self.max_workers) as ex:
                list(ex.map(merge_with_progress, channels))
            self.root.after(0, lambda: self._close_progress())
        except Exception as e:
            self.root.after(0, lambda e=e:
                messagebox.showerror("查询失败", str(e)))
            return
        # 在新窗口中显示
        self.root.after(0, lambda: self._show_device_window(
            device_id, device_name, channels))

    def _show_device_window(self, device_id, device_name, channels):
        """创建新窗口展示设备通道列表（含编辑、导入导出）"""
        win = tk.Toplevel(self.root)
        win.title(f"通道列表 - {device_name} ({device_id})")
        win.geometry("1000x680")
        win.minsize(850, 500)

        # ---- 顶部工具栏 ----
        toolbar = ttk.Frame(win)
        toolbar.pack(fill=tk.X, padx=10, pady=(8, 0))
        ttk.Label(toolbar, text=f"设备: {device_name}    ID: {device_id}",
                  font=("Microsoft YaHei", 11, "bold")).pack(side=tk.LEFT)
        ttk.Label(toolbar, text=f"共 {len(channels)} 个通道",
                  foreground="gray").pack(side=tk.RIGHT)

        # 按钮行
        btn_row = ttk.Frame(win)
        btn_row.pack(fill=tk.X, padx=10, pady=(5, 0))
        ttk.Button(btn_row, text=" 刷新",
                   command=lambda: self._win_refresh(win, device_id, device_name,
                                                     tree, item_to_ch, check_vars, set_st)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_row, text="导出Excel",
                   command=lambda: self._win_export(win, channels, device_name)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_row, text="导入Excel",
                   command=lambda: self._win_import(win, channels, device_id, device_name)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="批量修改区域编码",
                   command=lambda: self._win_batch_region(win, tree, check_vars, item_to_ch)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="批量修改厂家",
                   command=lambda: self._win_batch_manufacturer(win, tree, check_vars, item_to_ch)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="批量修改经纬度",
                   command=lambda: self._win_batch_lonlat(win, tree, check_vars, item_to_ch)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="批量修改在线状态",
                   command=lambda: self._win_batch_status(win, tree, check_vars, item_to_ch)).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_row, text="双击单元格修改", foreground="gray",
                  font=("Microsoft YaHei", 9)).pack(side=tk.RIGHT)

        # ---- 通道表格（带复选框）----
        frame = ttk.Frame(win)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        columns = ("☐", "序号", "通道名称", "通道类型", "在线状态", "区域编码", "经度", "纬度", "厂家", "数据库ID")
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=20)
        col_widths = [30, 40, 200, 80, 60, 100, 80, 80, 120, 80]
        for col, w in zip(columns, col_widths):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor=tk.CENTER, minwidth=w)

        # 表头 ☐ 点击 = 全选/取消全选
        def update_ch_header():
            if not check_vars:
                return
            all_checked = all(v.get() for v in check_vars.values())
            tree.heading("#1", text="☑" if all_checked else "☐")

        def toggle_all_ch():
            if not check_vars:
                return
            all_checked = all(v.get() for v in check_vars.values())
            new_state = not all_checked
            for item, v in check_vars.items():
                v.set(new_state)
                tree.set(item, "#1", "☑" if new_state else "☐")
            update_ch_header()
        tree.heading("#1", command=toggle_all_ch)

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # 存储数据
        item_to_ch = {}
        check_vars = {}
        
        for i, ch in enumerate(channels, 1):
            var = tk.BooleanVar(value=False)
            item = tree.insert("", tk.END, values=(
                "☐", i,
                ch.get("name", ""),
                "子目录" if ch.get("channelType") else "设备通道",
                ({"ON": "在线", "OFF": "离线"}.get(ch.get("status")) or ch.get("status") or ""),
                ch.get("civilCode", ""),
                ch.get("gbLongitude", 0),
                ch.get("gbLatitude", 0),
                ch.get("gbManufacturer", ""),
                ch.get("id", ""),
            ))
            item_to_ch[item] = ch
            check_vars[item] = var

        # 复选框点击
        def on_chk_click(event):
            region = tree.identify_region(event.x, event.y)
            if region != "cell":
                return
            col = tree.identify_column(event.x)
            if col != "#1":
                return
            it = tree.identify_row(event.y)
            if it and it in check_vars:
                v = check_vars[it]
                v.set(not v.get())
                tree.set(it, "#1", "☑" if v.get() else "☐")
                update_ch_header()
            return "break"

        tree.bind("<Button-1>", on_chk_click)

        # 双击编辑
        edit_entry_ref = [None]
        edit_item_ref = [None]
        edit_col_ref = [None]

        def cancel_edit():
            if edit_entry_ref[0]:
                edit_entry_ref[0].destroy()
                edit_entry_ref[0] = None
                edit_item_ref[0] = None
                edit_col_ref[0] = None

        def save_edit():
            if not edit_entry_ref[0]:
                return
            new_val = edit_entry_ref[0].get().strip()
            it, col = edit_item_ref[0], edit_col_ref[0]
            cancel_edit()
            if not it or not new_val:
                return
            ch = item_to_ch.get(it)
            if not ch:
                return
            # 判断是哪一列
            if col == "#3":  # 通道名称
                if new_val == ch.get("name", ""):
                    return
                upd = {"gbName": new_val}
                fld = "名称"
            elif col == "#6":  # 区域编码
                if new_val == ch.get("civilCode", ""):
                    return
                upd = {"gbCivilCode": new_val}
                fld = "区域编码"
            elif col == "#7":  # 经度
                try:
                    nv = float(new_val)
                    if not (-180 <= nv <= 180):
                        raise ValueError("out_of_range")
                except ValueError:
                    messagebox.showerror("错误", "经度必须是数字（-180 ~ 180）", parent=win)
                    return
                old = float(ch.get("gbLongitude", 0) or 0)
                if abs(nv - old) < 0.000001:
                    return
                upd = {"gbLongitude": nv}
                fld = "经度"
            elif col == "#8":  # 纬度
                try:
                    nv = float(new_val)
                    if not (-90 <= nv <= 90):
                        raise ValueError("out_of_range")
                except ValueError:
                    messagebox.showerror("错误", "纬度必须是数字（-90 ~ 90）", parent=win)
                    return
                old = float(ch.get("gbLatitude", 0) or 0)
                if abs(nv - old) < 0.000001:
                    return
                upd = {"gbLatitude": nv}
                fld = "纬度"
            elif col == "#9":  # 厂家
                if new_val == ch.get("gbManufacturer", ""):
                    return
                upd = {"gbManufacturer": new_val}
                fld = "厂家"
            else:
                return
            if not messagebox.askyesno("确认",
                    f"将 {fld} 改为 '{new_val}'？", parent=win):
                return
            # 发请求
            host = self.server_host.get().strip().rstrip('/')
            headers = {"Accept": "*/*", "access-token": self.access_token,
                       "Content-Type": "application/json"}
            body = self.build_channel_body(ch, upd)
            try:
                resp = requests.post(f"{host}/api/common/channel/update",
                                     headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    tree.set(it, col, new_val)
                    for k, v in upd.items():
                        ch[k] = v
                    self.set_statusbar(f"修改成功: {fld}")
                else:
                    msg = resp.json().get("msg", "修改失败")
                    messagebox.showerror("失败", msg, parent=win)
            except Exception as e:
                messagebox.showerror("异常", str(e), parent=win)

        def on_dblclick(event):
            col = tree.identify_column(event.x)
            if col in ("#1", "#2", "#4", "#5", "#10"):
                return
            cancel_edit()
            it = tree.identify_row(event.y)
            if not it:
                return
            val = tree.set(it, col)
            bbox = tree.bbox(it, col)
            if not bbox:
                return
            x, y, w, h = bbox
            entry = tk.Entry(tree, font=("Microsoft YaHei", 9))
            entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, val)
            entry.select_range(0, tk.END)
            entry.focus_set()
            entry.bind("<Return>", lambda e: save_edit())
            entry.bind("<FocusOut>", lambda e: win.after(200, save_edit))
            edit_entry_ref[0] = entry
            edit_item_ref[0] = it
            edit_col_ref[0] = col

        tree.bind("<Double-1>", on_dblclick)

        # 底部状态
        status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(win, textvariable=status_var, relief=tk.SUNKEN,
                               anchor=tk.W, font=("Microsoft YaHei", 8), padding=(5, 2))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        def set_st(msg):
            status_var.set(msg)

        # ---- 窗口内的导入导出 ----
        def do_win_export():
            if not channels:
                messagebox.showwarning("提示", "通道列表为空", parent=win)
                return
            f = filedialog.asksaveasfilename(defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")], title="导出通道",
                initialfile=f"{device_name}.xlsx")
            if not f:
                return
            wb = Workbook()
            ws = wb.active
            ws.title = "通道列表"
            hdrs = ["设备名称", "设备ID", "通道名称", "通道类型", "在线状态", "区域编码",
                    "经度", "纬度", "厂家", "数据库ID"]
            for c, h in enumerate(hdrs, 1):
                cell = ws.cell(row=1, column=c, value=h)
                cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            for i, ch in enumerate(channels, 1):
                row = [device_name, device_id,
                       ch.get("name", ""),
                       "子目录" if ch.get("channelType") else "设备通道",
                       ({"ON": "在线", "OFF": "离线"}.get(ch.get("status")) or ch.get("status") or ""),
                       ch.get("civilCode", ""),
                       ch.get("gbLongitude", 0), ch.get("gbLatitude", 0),
                       ch.get("gbManufacturer", ""), ch.get("id", "")]
                for j, v in enumerate(row, 1):
                    ws.cell(row=i + 1, column=j, value=v)
            for i, w in enumerate([18, 24, 25, 10, 10, 15, 10, 10, 15, 10], 1):
                ws.column_dimensions[ws.cell(1, i).column_letter].width = w
            wb.save(f)
            set_st(f"导出成功: {os.path.basename(f)}")

        def do_win_import():
            f = filedialog.askopenfilename(title="选择修改后的Excel",
                filetypes=[("Excel", "*.xlsx")])
            if not f:
                return
            try:
                wb = load_workbook(f)
                ws = wb.active
            except Exception as e:
                messagebox.showerror("读取失败", str(e), parent=win)
                return
            excel_data = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 10:
                    continue
                did = row[9]
                if not did:
                    continue
                try:
                    did = int(did)
                except:
                    continue
                excel_data[did] = {
                    "name": str(row[2] or "").strip(),
                    "civilCode": str(row[5] or "").strip(),
                    "gbLongitude": str(row[6] or "").strip(),
                    "gbLatitude": str(row[7] or "").strip(),
                    "gbManufacturer": str(row[8] or "").strip(),
                }
            if not excel_data:
                messagebox.showwarning("无数据", "Excel无有效数据", parent=win)
                return
            tasks = []
            for ch in channels:
                cid = ch.get("id")
                if cid in excel_data:
                    d = excel_data[cid]
                    n = d["name"]
                    c = d["civilCode"]
                    lo = d["gbLongitude"]
                    la = d["gbLatitude"]
                    mf = d["gbManufacturer"]
                    if (n != ch.get("name", "") or c != ch.get("civilCode", "") or
                        (lo and str(lo) != str(ch.get("gbLongitude", 0))) or
                        (la and str(la) != str(ch.get("gbLatitude", 0))) or
                        mf != ch.get("gbManufacturer", "")):
                        tasks.append((ch, n, c, lo, la, mf))
            if not tasks:
                messagebox.showinfo("提示", "没有需要修改的数据", parent=win)
                return
            if not messagebox.askyesno("确认",
                    f"检测到 {len(tasks)} 条修改，是否提交？", parent=win):
                return
            set_st("正在导入...")
            threading.Thread(target=self._win_batch_update,
                args=(win, tasks, tree, item_to_ch, set_st), daemon=True).start()

        # 绑定按钮
        for child in btn_row.winfo_children():
            if isinstance(child, ttk.Button):
                txt = child.cget("text")
                if "导出" in txt:
                    child.configure(command=do_win_export)
                elif "导入" in txt:
                    child.configure(command=do_win_import)

        # 关闭按钮
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=(0, 8))

    def _win_batch_update(self, win, tasks, tree, item_to_ch, set_st):
        """窗口内批量提交导入的修改（并发8线程）"""
        host = self.server_host.get().strip().rstrip('/')
        headers = {"Accept": "*/*", "access-token": self.access_token,
                   "Content-Type": "application/json"}
        total = len(tasks)
        done_count = [0]
        success = [0]
        fail = [0]
        updated_channels = []

        def do_update(task):
            ch, nn, nc, lo, la, mf = task
            upd = {}
            if nn != ch.get("name", ""):
                upd["gbName"] = nn
            if nc != ch.get("civilCode", ""):
                upd["gbCivilCode"] = nc
            if lo:
                try:
                    val = float(lo)
                    if abs(val - float(ch.get("gbLongitude", 0) or 0)) > 0.000001:
                        upd["gbLongitude"] = val
                except ValueError:
                    pass
            if la:
                try:
                    val = float(la)
                    if abs(val - float(ch.get("gbLatitude", 0) or 0)) > 0.000001:
                        upd["gbLatitude"] = val
                except ValueError:
                    pass
            if mf and mf != ch.get("gbManufacturer", ""):
                upd["gbManufacturer"] = mf
            if not upd:
                return True, ch, None
            try:
                body = self.build_channel_body(ch, upd)
                resp = requests.post(f"{host}/api/common/channel/update",
                                     headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    for k, v in upd.items():
                        ch[k] = v
                    return True, ch, None
            except Exception:
                pass
            return False, None, upd

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(do_update, task): task for task in tasks}
            for future in as_completed(futures):
                ok, ch, upd = future.result()
                if ok:
                    success[0] += 1
                    if ch:
                        updated_channels.append(ch)
                else:
                    fail[0] += 1
                done_count[0] += 1
                self.root.after(0, lambda c=done_count[0], t=total:
                    set_st(f"导入中 ({c}/{t})..."))
        self.root.after(0, lambda: set_st(f"导入完成: 成功 {success[0]}, 失败 {fail[0]}"))
        self.root.after(0, lambda: messagebox.showinfo("结果",
            f"成功: {success[0]} 条\n失败: {fail[0]} 条", parent=win))
        if success[0] > 0:
            self.root.after(0, lambda: self._refresh_win_tree(tree, [t[0] for t in tasks], item_to_ch))

    def _win_batch_region(self, win, tree, check_vars, item_to_ch):
        """弹窗内批量修改区域编码"""
        checked = [it for it, var in check_vars.items() if var.get()]
        if not checked:
            messagebox.showwarning("提示", "请先勾选要修改的通道", parent=win)
            return
        new_region = simpledialog.askstring("批量修改区域编码",
                                            f"已选中 {len(checked)} 个通道\n请输入新的区域编码:",
                                            parent=win)
        if not new_region or not new_region.strip():
            return
        new_region = new_region.strip()
        if not messagebox.askyesno("确认",
                f"将为选中的 {len(checked)} 个通道设置区域编码为:\n'{new_region}'?\n\n确认修改?",
                parent=win):
            return
        threading.Thread(target=self._win_do_batch_region,
                         args=(win, checked, tree, new_region, item_to_ch), daemon=True).start()

    def _win_do_batch_region(self, win, checked_items, tree, new_region, item_to_ch):
        """实际执行批量修改（并发8线程）"""
        host = self.server_host.get().strip().rstrip('/')
        headers = {"Accept": "*/*", "access-token": self.access_token,
                   "Content-Type": "application/json"}
        total = len(checked_items)
        done_count = [0]
        success = [0]
        fail = [0]

        def do_update(it):
            ch = item_to_ch.get(it)
            if not ch:
                return False, None, None
            updates = {"gbCivilCode": new_region, "civilCode": new_region}
            try:
                body = self.build_channel_body(ch, updates)
                resp = requests.post(f"{host}/api/common/channel/update",
                                     headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    return True, ch, it
            except Exception:
                pass
            return False, None, None

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(do_update, it): it for it in checked_items}
            for future in as_completed(futures):
                ok, ch, it = future.result()
                if ok:
                    success[0] += 1
                    if ch and it:
                        ch["civilCode"] = new_region
                        self.root.after(0, lambda i=it: tree.set(i, "#6", new_region))
                else:
                    fail[0] += 1
                done_count[0] += 1
                self.root.after(0, lambda c=done_count[0], t=total:
                    self.set_statusbar(f"批量修改中 ({c}/{t})..."))
        self.root.after(0, lambda: self.set_statusbar(
            f"批量修改完成: 成功 {success[0]}, 失败 {fail[0]}"))
        self.root.after(0, lambda: messagebox.showinfo("结果",
            f"成功: {success[0]} 条\n失败: {fail[0]} 条", parent=win))

    def _ask_manufacturer(self, parent, count):
        """弹出自定义对话框，下拉列表选择或输入厂家"""
        dialog = tk.Toplevel(parent)
        dialog.title("批量修改厂家")
        dialog.geometry("420x220")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()

        dialog.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 420) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 220) // 2
        dialog.geometry(f"+{x}+{y}")

        result = [None]

        ttk.Label(dialog, text=f"已选中 {count} 个通道\n请选择或输入厂家名称:",
                  font=("Microsoft YaHei", 10)).pack(pady=(15, 5))

        var = tk.StringVar()
        combo = ttk.Combobox(dialog, textvariable=var, font=("Microsoft YaHei", 10),
                             values=["HIKVISION", "Dahua", "Uniview", "Tiandy", "自定义..."],
                             state="normal", width=30)
        combo.pack(pady=5)
        combo.current(0)

        custom_frame = ttk.Frame(dialog)
        custom_frame.pack(pady=5)
        custom_label = ttk.Label(custom_frame, text="自定义:",
                                 font=("Microsoft YaHei", 9))
        custom_var = tk.StringVar()
        custom_entry = ttk.Entry(custom_frame, textvariable=custom_var,
                                 font=("Microsoft YaHei", 10), width=25)

        def on_combo_select(event):
            if var.get() == "自定义...":
                custom_label.pack(side=tk.LEFT, padx=(0, 5))
                custom_entry.pack(side=tk.LEFT)
                custom_entry.focus_set()
            else:
                custom_label.pack_forget()
                custom_entry.pack_forget()

        combo.bind("<<ComboboxSelected>>", on_combo_select)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=(15, 10))

        def on_ok():
            val = var.get()
            if val == "自定义...":
                val = custom_var.get().strip()
            if not val:
                messagebox.showwarning("提示", "请输入厂家名称", parent=dialog)
                return
            result[0] = val
            dialog.destroy()

        ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=8)

        dialog.wait_window()
        return result[0]

    def _win_batch_manufacturer(self, win, tree, check_vars, item_to_ch):
        """弹窗内批量修改厂家"""
        checked = [it for it, var in check_vars.items() if var.get()]
        if not checked:
            messagebox.showwarning("提示", "请先勾选要修改的通道", parent=win)
            return
        mfr = self._ask_manufacturer(win, len(checked))
        if not mfr:
            return
        if not messagebox.askyesno("确认",
                f"将为选中的 {len(checked)} 个通道设置厂家为:\n'{mfr}'?\n\n确认修改?",
                parent=win):
            return
        threading.Thread(target=self._win_do_batch_manufacturer,
                         args=(win, checked, tree, mfr, item_to_ch), daemon=True).start()

    def _win_do_batch_manufacturer(self, win, checked_items, tree, mfr, item_to_ch):
        """实际执行批量修改厂家（并发）"""
        host = self.server_host.get().strip().rstrip('/')
        headers = {"Accept": "*/*", "access-token": self.access_token,
                   "Content-Type": "application/json"}
        total = len(checked_items)
        done_count = [0]
        success = [0]
        fail = [0]

        def do_update(it):
            ch = item_to_ch.get(it)
            if not ch:
                return False, None, None
            updates = {"gbManufacturer": mfr}
            try:
                body = self.build_channel_body(ch, updates)
                resp = requests.post(f"{host}/api/common/channel/update",
                                     headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    return True, ch, it
            except Exception:
                pass
            return False, None, None

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(do_update, it): it for it in checked_items}
            for future in as_completed(futures):
                ok, ch, it = future.result()
                if ok:
                    success[0] += 1
                    if ch and it:
                        ch["gbManufacturer"] = mfr
                        self.root.after(0, lambda i=it: tree.set(i, "#9", mfr))
                else:
                    fail[0] += 1
                done_count[0] += 1
                self.root.after(0, lambda c=done_count[0], t=total:
                    self.set_statusbar(f"批量修改中 ({c}/{t})..."))
        self.root.after(0, lambda: self.set_statusbar(
            f"批量修改完成: 成功 {success[0]}, 失败 {fail[0]}"))
        self.root.after(0, lambda: messagebox.showinfo("结果",
            f"成功: {success[0]} 条\n失败: {fail[0]} 条", parent=win))

    def _win_do_batch_lonlat(self, win, checked_items, tree, lon, lat, item_to_ch):
        """实际执行批量修改经纬度（并发）"""
        host = self.server_host.get().strip().rstrip('/')
        headers = {"Accept": "*/*", "access-token": self.access_token,
                   "Content-Type": "application/json"}
        total = len(checked_items)
        done_count = [0]
        success = [0]
        fail = [0]

        def do_update(it):
            ch = item_to_ch.get(it)
            if not ch:
                return False, None, None
            updates = {}
            if lon is not None:
                old = float(ch.get("gbLongitude", 0) or 0)
                if abs(lon - old) >= 0.000001:
                    updates["gbLongitude"] = lon
            if lat is not None:
                old = float(ch.get("gbLatitude", 0) or 0)
                if abs(lat - old) >= 0.000001:
                    updates["gbLatitude"] = lat
            if not updates:
                return True, ch, it
            try:
                body = self.build_channel_body(ch, updates)
                resp = requests.post(f"{host}/api/common/channel/update",
                                     headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    return True, ch, it
            except Exception:
                pass
            return False, None, None

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(do_update, it): it for it in checked_items}
            for future in as_completed(futures):
                ok, ch, it = future.result()
                if ok:
                    success[0] += 1
                    if ch and it:
                        if lon is not None:
                            ch["gbLongitude"] = lon
                            self.root.after(0, lambda i=it: tree.set(i, "#7", lon))
                        if lat is not None:
                            ch["gbLatitude"] = lat
                            self.root.after(0, lambda i=it: tree.set(i, "#8", lat))
                else:
                    fail[0] += 1
                done_count[0] += 1
                self.root.after(0, lambda c=done_count[0], t=total:
                    self.set_statusbar(f"批量修改中 ({c}/{t})..."))
        self.root.after(0, lambda: self.set_statusbar(
            f"批量修改完成: 成功 {success[0]}, 失败 {fail[0]}"))
        self.root.after(0, lambda: messagebox.showinfo("结果",
            f"成功: {success[0]} 条\n失败: {fail[0]} 条", parent=win))

    def _ask_status(self, parent, count):
        """弹出自定义对话框，选择在线状态"""
        dialog = tk.Toplevel(parent)
        dialog.title("批量修改在线状态")
        dialog.geometry("360x160")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()

        dialog.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 360) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 160) // 2
        dialog.geometry(f"+{x}+{y}")

        result = [None]

        ttk.Label(dialog, text=f"已选中 {count} 个通道\n请选择要设置的在线状态:",
                  font=("Microsoft YaHei", 10)).pack(pady=(15, 10))

        var = tk.StringVar()
        combo = ttk.Combobox(dialog, textvariable=var, font=("Microsoft YaHei", 10),
                             values=["ON（在线）", "OFF（离线）"],
                             state="readonly", width=20)
        combo.pack(pady=5)
        combo.current(0)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=(15, 10))

        def on_ok():
            val = var.get()
            if val.startswith("ON"):
                result[0] = "ON"
            else:
                result[0] = "OFF"
            dialog.destroy()

        ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=8)

        combo.bind("<<ComboboxSelected>>", lambda e: on_ok())

        dialog.wait_window()
        return result[0]

    def _win_batch_status(self, win, tree, check_vars, item_to_ch):
        """弹窗内批量修改通道在线状态"""
        checked = [it for it, var in check_vars.items() if var.get()]
        if not checked:
            messagebox.showwarning("提示", "请先勾选要修改的通道", parent=win)
            return
        status = self._ask_status(win, len(checked))
        if not status:
            return
        label = "在线" if status == "ON" else "离线"
        if not messagebox.askyesno("确认",
                f"将为选中的 {len(checked)} 个通道设置状态为:\n{label}?\n\n确认修改?",
                parent=win):
            return
        threading.Thread(target=self._win_do_batch_status,
                         args=(win, checked, tree, status, item_to_ch), daemon=True).start()

    def _win_do_batch_status(self, win, checked_items, tree, status, item_to_ch):
        """实际执行批量修改在线状态（并发）"""
        host = self.server_host.get().strip().rstrip('/')
        headers = {"Accept": "*/*", "access-token": self.access_token,
                   "Content-Type": "application/json"}
        total = len(checked_items)
        done_count = [0]
        success = [0]
        fail = [0]

        def do_update(it):
            ch = item_to_ch.get(it)
            if not ch:
                return False, None, None
            # 跳过已经是目标状态的通道
            old = ch.get("status", "")
            if old == status:
                return True, ch, it
            updates = {"gbStatus": status}
            try:
                body = self.build_channel_body(ch, updates)
                resp = requests.post(f"{host}/api/common/channel/update",
                                     headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    return True, ch, it
            except Exception:
                pass
            return False, None, None

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(do_update, it): it for it in checked_items}
            for future in as_completed(futures):
                ok, ch, it = future.result()
                if ok:
                    success[0] += 1
                    if ch:
                        ch["status"] = status
                else:
                    fail[0] += 1
                done_count[0] += 1
                self.root.after(0, lambda c=done_count[0], t=total:
                    self.set_statusbar(f"批量修改中 ({c}/{t})..."))
        self.root.after(0, lambda: self.set_statusbar(
            f"批量修改完成: 成功 {success[0]}, 失败 {fail[0]}"))
        self.root.after(0, lambda: messagebox.showinfo("结果",
            f"成功: {success[0]} 条\n失败: {fail[0]} 条", parent=win))

    def _ask_lonlat(self, parent, count):
        """弹出自定义对话框，输入经度和纬度"""
        dialog = tk.Toplevel(parent)
        dialog.title("批量修改经纬度")
        dialog.geometry("360x210")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()

        dialog.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 360) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 210) // 2
        dialog.geometry(f"+{x}+{y}")

        result = [None]

        ttk.Label(dialog, text=f"已选中 {count} 个通道\n请输入经纬度（留空表示不修改该字段）:",
                  font=("Microsoft YaHei", 10)).pack(pady=(15, 10))

        # 经度
        lon_frame = ttk.Frame(dialog)
        lon_frame.pack(pady=3)
        ttk.Label(lon_frame, text="经度 (-180 ~ 180):", font=("Microsoft YaHei", 9),
                  width=16, anchor=tk.E).pack(side=tk.LEFT)
        lon_var = tk.StringVar()
        lon_entry = ttk.Entry(lon_frame, textvariable=lon_var, font=("Microsoft YaHei", 10), width=18)
        lon_entry.pack(side=tk.LEFT, padx=(5, 0))

        # 纬度
        lat_frame = ttk.Frame(dialog)
        lat_frame.pack(pady=3)
        ttk.Label(lat_frame, text="纬度 (-90 ~ 90):", font=("Microsoft YaHei", 9),
                  width=16, anchor=tk.E).pack(side=tk.LEFT)
        lat_var = tk.StringVar()
        lat_entry = ttk.Entry(lat_frame, textvariable=lat_var, font=("Microsoft YaHei", 10), width=18)
        lat_entry.pack(side=tk.LEFT, padx=(5, 0))

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=(15, 10))

        def on_ok():
            lon_str = lon_var.get().strip()
            lat_str = lat_var.get().strip()
            if not lon_str and not lat_str:
                messagebox.showwarning("提示", "请至少输入经度或纬度", parent=dialog)
                return
            lon_val = lat_val = None
            if lon_str:
                try:
                    lon_val = float(lon_str)
                    if not (-180 <= lon_val <= 180):
                        raise ValueError
                except ValueError:
                    messagebox.showerror("错误", "经度必须是 -180 ~ 180 之间的数字", parent=dialog)
                    return
            if lat_str:
                try:
                    lat_val = float(lat_str)
                    if not (-90 <= lat_val <= 90):
                        raise ValueError
                except ValueError:
                    messagebox.showerror("错误", "纬度必须是 -90 ~ 90 之间的数字", parent=dialog)
                    return
            result[0] = (lon_val, lat_val)
            dialog.destroy()

        ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=8)

        lon_entry.bind("<Return>", lambda e: lat_entry.focus_set())
        lat_entry.bind("<Return>", lambda e: on_ok())

        dialog.wait_window()
        return result[0]

    def _win_batch_lonlat(self, win, tree, check_vars, item_to_ch):
        """弹窗内批量修改经纬度"""
        checked = [it for it, var in check_vars.items() if var.get()]
        if not checked:
            messagebox.showwarning("提示", "请先勾选要修改的通道", parent=win)
            return
        vals = self._ask_lonlat(win, len(checked))
        if not vals:
            return
        lon, lat = vals
        parts = []
        if lon is not None:
            parts.append(f"经度={lon}")
        if lat is not None:
            parts.append(f"纬度={lat}")
        msg = "、".join(parts)
        if not messagebox.askyesno("确认",
                f"将为选中的 {len(checked)} 个通道设置\n{msg}?\n\n确认修改?",
                parent=win):
            return
        threading.Thread(target=self._win_do_batch_lonlat,
                         args=(win, checked, tree, lon, lat, item_to_ch), daemon=True).start()

    def _win_refresh(self, win, device_id, device_name,
                     tree, item_to_ch, check_vars, set_st):
        """重新查询设备通道并刷新窗口表格"""
        set_st("刷新中...")
        threading.Thread(target=self._do_win_refresh,
            args=(win, device_id, device_name,
                  tree, item_to_ch, check_vars, set_st), daemon=True).start()

    def _do_win_refresh(self, win, device_id, device_name,
                        tree, item_to_ch, check_vars, set_st):
        """后台重新查询并刷新表格（并发分页）"""
        host = self.server_host.get().strip().rstrip('/')
        token = self.access_token
        try:
            headers = {"Accept": "*/*", "access-token": token}
            new_channels, total = self._concurrent_channel_query(device_id)
            # 并发合并全局通道数据
            merge_done = [0]
            merge_total = len(new_channels)

            def merge_one(dc):
                dc_id = dc.get("id") or dc.get("gbId")
                if not dc_id:
                    return dc
                try:
                    one_resp = requests.get(f"{host}/api/common/channel/one",
                                            headers=headers,
                                            params={"id": dc_id}, timeout=15)
                    if one_resp.status_code == 200:
                        od = one_resp.json()
                        if od.get("code") == 0:
                            gc = od.get("data")
                            if gc and isinstance(gc, dict):
                                for fld in ("gbManufacturer", "gbLongitude", "gbLatitude",
                                            "gbName", "gbCivilCode"):
                                    if fld in gc and gc[fld] is not None:
                                        dc[fld] = gc[fld]
                except Exception:
                    pass
                return dc

            def merge_with_progress(dc):
                result = merge_one(dc)
                merge_done[0] += 1
                if merge_done[0] % 100 == 0:
                    self.root.after(0, lambda c=merge_done[0], t=merge_total:
                        set_st(f"合并通道数据 ({c}/{t})..."))
                return result

            with ThreadPoolExecutor(max_workers=self.max_workers) as ex:
                list(ex.map(merge_with_progress, new_channels))
        except Exception as e:
            self.root.after(0, lambda e=e: set_st(f"刷新失败: {e}"))
            return
        # 在主线程更新表格
        self.root.after(0, lambda: self._rebuild_win_tree(
            win, tree, new_channels, item_to_ch, check_vars, set_st))

    def _rebuild_win_tree(self, win, tree, new_channels, item_to_ch, check_vars, set_st):
        """重建窗口表格数据"""
        # 清除旧数据
        for item in tree.get_children():
            tree.delete(item)
        item_to_ch.clear()
        check_vars.clear()
        # 填入新数据
        for i, ch in enumerate(new_channels, 1):
            var = tk.BooleanVar(value=False)
            item = tree.insert("", tk.END, values=(
                "☐", i,
                ch.get("name", ""),
                "子目录" if ch.get("channelType") else "设备通道",
                ({"ON": "在线", "OFF": "离线"}.get(ch.get("status")) or ch.get("status") or ""),
                ch.get("civilCode", ""),
                ch.get("gbLongitude", 0),
                ch.get("gbLatitude", 0),
                ch.get("gbManufacturer", ""),
                ch.get("id", ""),
            ))
            item_to_ch[item] = ch
            check_vars[item] = var
        # 更新标题栏的通道数
        for child in win.winfo_children():
            if isinstance(child, ttk.Frame):
                for sub in child.winfo_children():
                    if isinstance(sub, ttk.Label) and "共" in sub.cget("text"):
                        sub.configure(text=f"共 {len(new_channels)} 个通道")
                        break
        set_st(f"已刷新，共 {len(new_channels)} 个通道")

    @staticmethod
    def _refresh_win_tree(tree, channels, item_to_ch):
        """刷新窗口内表格显示"""
        for item in tree.get_children():
            ch = item_to_ch.get(item)
            if ch:
                tree.set(item, "#3", ch.get("name", ""))
                tree.set(item, "#5", ({"ON": "在线", "OFF": "离线"}.get(ch.get("status")) or ch.get("status") or ""))
                tree.set(item, "#6", ch.get("civilCode", ""))
                tree.set(item, "#7", ch.get("gbLongitude", 0))
                tree.set(item, "#8", ch.get("gbLatitude", 0))
                tree.set(item, "#9", ch.get("gbManufacturer", ""))



    def build_channel_body(self, channel, updates):
        """构建 body：gbId + 有值的已有字段 + 要更新的字段
        注意：设备通道查询返回短名称(id, deviceId, name, civilCode)，
        但全局通道 API 需要 gb 前缀(gbId, gbDeviceId, gbName, gbCivilCode)，
        这里做字段映射。
        """
        # gbId：兼容设备通道(id)和全局通道(gbId)两种字段名
        gb_id_val = channel.get("id") or channel.get("gbId") or 0
        body = {"gbId": gb_id_val}
        # 字段映射：(body里的key, channel里的key, 默认值)
        field_map = [
            ("gbDeviceId", "deviceId", ""),
            ("gatewayDeviceId", "gatewayDeviceId", ""),
            ("gbName", "name", ""),
            ("gbManufacturer", "gbManufacturer", ""),
            ("gbModel", "gbModel", ""),
            ("gbOwner", "gbOwner", ""),
            ("gbCivilCode", "civilCode", ""),
            ("gbBlock", "gbBlock", ""),
            ("gbAddress", "gbAddress", ""),
            ("gbParental", "gbParental", 0),
            ("gbParentId", "gbParentId", ""),
            ("gbSafetyWay", "gbSafetyWay", 0),
            ("gbRegisterWay", "gbRegisterWay", 0),
            ("gbCertNum", "gbCertNum", ""),
            ("gbCertifiable", "gbCertifiable", 0),
            ("gbErrCode", "gbErrCode", 0),
            ("gbEndTime", "gbEndTime", ""),
            ("gbSecrecy", "gbSecrecy", 0),
            ("gbIpAddress", "gbIpAddress", ""),
            ("gbPort", "gbPort", 0),
            ("gbPassword", "gbPassword", ""),
            ("gbStatus", "gbStatus", ""),
            ("gbLongitude", "gbLongitude", 0),
            ("gbLatitude", "gbLatitude", 0),
            ("gpsAltitude", "gpsAltitude", 0),
            ("gpsSpeed", "gpsSpeed", 0),
            ("gpsDirection", "gpsDirection", 0),
            ("gpsTime", "gpsTime", ""),
            ("gbBusinessGroupId", "gbBusinessGroupId", ""),
            ("gbPtzType", "gbPtzType", 0),
            ("gbPositionType", "gbPositionType", 0),
            ("gbRoomType", "gbRoomType", 0),
            ("gbUseType", "gbUseType", 0),
            ("gbSupplyLightType", "gbSupplyLightType", 0),
            ("gbDirectionType", "gbDirectionType", 0),
            ("gbResolution", "gbResolution", ""),
            ("gbDownloadSpeed", "gbDownloadSpeed", ""),
            ("gbSvcSpaceSupportMod", "gbSvcSpaceSupportMod", 0),
            ("gbSvcTimeSupportMode", "gbSvcTimeSupportMode", 0),
            ("recordPLan", "recordPLan", 0),
            ("dataType", "dataType", 0),
            ("dataDeviceId", "dataDeviceId", 0),
            ("createTime", "createTime", ""),
            ("updateTime", "updateTime", ""),
        ]
        for body_key, ch_key, dflt in field_map:
            # 兼容短名称(id/name)和gb前缀(gbId/gbName)
            val = channel.get(ch_key) or channel.get(body_key) or dflt
            if val not in (None, "", 0, 0.0):
                body[body_key] = val
        body.update(updates)
        return body


    # ---------- 导出选中设备的所有通道 ----------
    def export_selected_devices(self):
        """导出左侧勾选的设备下的全部通道"""
        if not self.access_token:
            messagebox.showwarning("警告", "请先登录")
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("缺少库", "请安装 openpyxl: pip install openpyxl")
            return
        # 获取勾选的设备
        checked_items = [it for it, var in self.dev_check_vars.items() if var.get()]
        if not checked_items:
            messagebox.showwarning("提示", "请先在左侧设备列表中勾选要导出的设备（点击 ☐ 框）")
            return
        selected_devices = []
        for item in checked_items:
            values = self.dev_tree.item(item, "values")
            if values:
                dev_id = values[1]  # 设备ID在第2列（索引1）
                for dev in self.all_devices:
                    if dev.get("deviceId", "") == dev_id:
                        selected_devices.append(dev)
                        break
        if not selected_devices:
            messagebox.showwarning("提示", "未找到选中的设备数据")
            return
        if not messagebox.askyesno("确认", f"将导出 {len(selected_devices)} 个选中设备的全部通道，\n是否继续?"):
            return
        # 默认文件名 = 设备名称
        name_list = [d.get("name", f"设备{i+1}") for i, d in enumerate(selected_devices)]
        default_name = "、".join(name_list[:3])
        if len(name_list) > 3:
            default_name += f"等{len(name_list)}个设备"
        default_name += ".xlsx"
        file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel", "*.xlsx")],
                                             title="导出选中通道",
                                             initialfile=default_name)
        if not file:
            return
        self.set_statusbar("正在导出选中设备通道...")
        threading.Thread(target=self._do_export_selected,
                         args=(file, selected_devices), daemon=True).start()

    def _do_export_selected(self, file, devices):
        try:
            host = self.server_host.get().strip().rstrip('/')
            token = self.access_token
            headers = {"Accept": "*/*", "access-token": token}
            all_rows = []
            total = len(devices)
            for idx, dev in enumerate(devices, 1):
                device_id = dev.get("deviceId", "")
                device_name = dev.get("name", "")
                self.root.after(0, lambda c=idx, d=device_name:
                    self.set_statusbar(f"导出中 ({c}/{total}): {d}"))
                try:
                    channels, _ = self._concurrent_channel_query(device_id)
                    # 并发合并全局通道数据
                    def merge_one(dc):
                        dc_id = dc.get("id") or dc.get("gbId")
                        if not dc_id:
                            return dc
                        try:
                            one = requests.get(f"{host}/api/common/channel/one",
                                               headers=headers,
                                               params={"id": dc_id}, timeout=15)
                            if one.status_code == 200:
                                od = one.json()
                                if od.get("code") == 0:
                                    gc = od.get("data")
                                    if gc and isinstance(gc, dict):
                                        for f in ("gbManufacturer", "gbLongitude", "gbLatitude",
                                                  "gbName", "gbCivilCode"):
                                            if f in gc and gc[f] is not None:
                                                dc[f] = gc[f]
                        except Exception:
                            pass
                        return dc
                    with ThreadPoolExecutor(max_workers=self.max_workers) as ex:
                        list(ex.map(merge_one, channels))
                    for ch in channels:
                        all_rows.append((
                            device_name,
                            device_id,
                            ch.get("name", ""),
                            "子目录" if ch.get("channelType") else "设备通道",
                            ch.get("civilCode", ""),
                            ch.get("gbLongitude", 0),
                            ch.get("gbLatitude", 0),
                            ch.get("gbManufacturer", ""),
                            ch.get("id", ""),
                        ))
                except Exception:
                    pass
            # 写入 Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "通道列表"
            headers_row = ["设备名称", "设备ID", "通道名称", "通道类型", "区域编码",
                           "经度", "纬度", "厂家", "数据库ID"]
            for c, h in enumerate(headers_row, 1):
                cell = ws.cell(row=1, column=c, value=h)
                cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            for i, row_data in enumerate(all_rows, 1):
                for j, v in enumerate(row_data, 1):
                    ws.cell(row=i + 1, column=j, value=v)
            col_widths = [18, 24, 25, 10, 15, 10, 10, 15, 10]
            for i, w in enumerate(col_widths, 1):
                ws.column_dimensions[ws.cell(1, i).column_letter].width = w
            wb.save(file)
            self.root.after(0, lambda: self.set_statusbar(
                f"导出完成: {len(all_rows)} 条通道 → {os.path.basename(file)}"))
            self.root.after(0, lambda: messagebox.showinfo(
                "成功", f"共导出 {len(all_rows)} 条通道\n保存到: {file}"))
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda e=e: messagebox.showerror("导出失败", str(e)))

    # ---------- 导入（从"导出勾选通道"生成的Excel）----------
    def import_device_excel(self):
        """导入修改后的通道Excel（支持多设备）"""
        if not self.access_token:
            messagebox.showwarning("警告", "请先登录")
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("缺少库", "请安装 openpyxl: pip install openpyxl")
            return
        file = filedialog.askopenfilename(title="选择修改后的Excel",
                                           filetypes=[("Excel", "*.xlsx")])
        if not file:
            return
        try:
            wb = load_workbook(file)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        # 读取每一行
        updates_by_id = {}  # db_id → {field: new_value}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 9:
                continue
            db_id = row[8]  # 数据库ID
            if not db_id:
                continue
            try:
                db_id = int(db_id)
            except (ValueError, TypeError):
                continue
            new_name = str(row[2] or "").strip()
            new_civil = str(row[4] or "").strip()
            new_lon = str(row[5] or "").strip()
            new_lat = str(row[6] or "").strip()
            new_mfr = str(row[7] or "").strip()
            updates_by_id[db_id] = (new_name, new_civil, new_lon, new_lat, new_mfr)
        if not updates_by_id:
            messagebox.showwarning("无数据", "Excel中未找到有效数据")
            return
        # 确认
        if not messagebox.askyesno("确认", f"检测到 {len(updates_by_id)} 条修改，是否提交?"):
            return
        self.set_statusbar("正在导入修改...")
        threading.Thread(target=self._do_import_device,
                         args=(file, updates_by_id), daemon=True).start()

    def _do_import_device(self, file, updates_by_id):
        """批量提交导入的修改（并发8线程）"""
        host = self.server_host.get().strip().rstrip('/')
        token = self.access_token
        headers = {"Accept": "*/*", "access-token": token,
                   "Content-Type": "application/json"}
        total = len(updates_by_id)
        done_count = [0]
        success = [0]
        fail = [0]

        def process_one(args):
            db_id, (new_name, new_civil, new_lon, new_lat, new_mfr) = args
            try:
                one = requests.get(f"{host}/api/common/channel/one",
                                   headers=headers,
                                   params={"id": db_id}, timeout=15)
                if one.status_code != 200:
                    return False
                od = one.json()
                if od.get("code") != 0:
                    return False
                ch = od.get("data")
                if not ch or not isinstance(ch, dict):
                    return False
            except Exception:
                return False
            updates = {}
            if new_name and new_name != ch.get("gbName", ch.get("name", "")):
                updates["gbName"] = new_name
            if new_civil and new_civil != ch.get("gbCivilCode", ch.get("civilCode", "")):
                updates["gbCivilCode"] = new_civil
            if new_lon:
                try:
                    val = float(new_lon)
                    old = float(ch.get("gbLongitude", 0) or 0)
                    if abs(val - old) > 0.000001:
                        updates["gbLongitude"] = val
                except ValueError:
                    pass
            if new_lat:
                try:
                    val = float(new_lat)
                    old = float(ch.get("gbLatitude", 0) or 0)
                    if abs(val - old) > 0.000001:
                        updates["gbLatitude"] = val
                except ValueError:
                    pass
            if new_mfr and new_mfr != ch.get("gbManufacturer", ""):
                updates["gbManufacturer"] = new_mfr
            if not updates:
                return True
            try:
                body = self.build_channel_body(ch, updates)
                resp = requests.post(f"{host}/api/common/channel/update",
                                     headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    return True
            except Exception:
                pass
            return False

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {executor.submit(process_one, item): item
                       for item in updates_by_id.items()}
            for future in as_completed(futures):
                if future.result():
                    success[0] += 1
                else:
                    fail[0] += 1
                done_count[0] += 1
                self.root.after(0, lambda c=done_count[0], t=total:
                    self.set_statusbar(f"导入中 ({c}/{t})..."))

        self.root.after(0, lambda: messagebox.showinfo(
            "结果", f"成功: {success[0]} 条\n失败: {fail[0]} 条"))
        self.root.after(0, lambda: self.set_statusbar(
            f"导入完成: 成功 {success[0]}, 失败 {fail[0]}"))



    # ---------- 退出 ----------
    def logout(self):
        if messagebox.askyesno("退出", "确定要退出登录吗？"):
            self.root.destroy()

    def on_close(self):
        self.root.destroy()
