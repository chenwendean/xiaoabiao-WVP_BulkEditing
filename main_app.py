#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主界面 - 通道查询、编辑、批量修改、Excel导入导出
v7.0 - 新增设备列表选择，登录后自动查询全部国标设备
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import hashlib
import requests
import threading
from datetime import datetime
import os

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

        # 设备列表相关
        self.all_devices = []           # 全部设备列表
        self.selected_device_id = None  # 当前选中的设备ID
        self.selected_device_name = ""
        self.selected_device_status = ""

        self.status_text = tk.StringVar(value=f"已登录: {user_info.get('username', '')}")
        self.page_num = 1
        self.page_size = 100
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

    def on_dev_checkbox_click(self, event):
        """设备列表复选框点击"""
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
        """后台查询通道并在新窗口展示"""
        channels = []
        # 主线程取值，避免线程中访问 tk 变量
        host = self.server_host.get().strip().rstrip('/')
        token = self.access_token
        try:
            headers = {"Accept": "*/*", "access-token": token}
            page = 1
            while True:
                url = f"{host}/api/device/query/devices/{device_id}/channels"
                resp = requests.get(url, headers=headers, params={
                    "page": page, "count": 100,
                    "query": "", "cameraQuery": "", "nvrQuery": ""
                }, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    if data.get("code") == 0:
                        lst = data.get("data", {}).get("list", [])
                        channels.extend(lst)
                        if len(lst) < 100:
                            break
                        page += 1
                    else:
                        break
                else:
                    break
            # 合并全局通道数据（厂家/经纬度）
            for dc in channels:
                dc_id = dc.get("id") or dc.get("gbId")
                if not dc_id:
                    continue
                try:
                    one_url = f"{host}/api/common/channel/one"
                    one_resp = requests.get(one_url, headers=headers,
                                            params={"id": dc_id}, timeout=15)
                    if one_resp.status_code == 200:
                        one_data = one_resp.json()
                        if one_data.get("code") == 0:
                            gc = one_data.get("data")
                            if gc and isinstance(gc, dict):
                                for fld in ("gbManufacturer", "gbLongitude", "gbLatitude",
                                            "gbName", "gbCivilCode"):
                                    if fld in gc and gc[fld] is not None:
                                        dc[fld] = gc[fld]
                except Exception:
                    pass
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
                   command=lambda: self._win_batch_region(win, channels, tree)).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_row, text="双击单元格修改", foreground="gray",
                  font=("Microsoft YaHei", 9)).pack(side=tk.RIGHT)

        # ---- 通道表格（带复选框）----
        frame = ttk.Frame(win)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        columns = ("☐", "序号", "通道名称", "通道类型", "区域编码", "经度", "纬度", "厂家", "数据库ID")
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=20)
        col_widths = [30, 40, 200, 80, 100, 80, 80, 120, 80]
        for col, w in zip(columns, col_widths):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor=tk.CENTER, minwidth=w)

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
            col = tree.identify_column(event.x)
            if col != "#1":
                return
            it = tree.identify_row(event.y)
            if it and it in check_vars:
                v = check_vars[it]
                v.set(not v.get())
                tree.set(it, "#1", "☑" if v.get() else "☐")

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
            elif col == "#5":  # 区域编码
                if new_val == ch.get("civilCode", ""):
                    return
                upd = {"gbCivilCode": new_val}
                fld = "区域编码"
            elif col == "#6":  # 经度
                try:
                    nv = float(new_val)
                except ValueError:
                    messagebox.showerror("错误", "经度必须是数字", parent=win)
                    return
                old = float(ch.get("gbLongitude", 0) or 0)
                if abs(nv - old) < 0.000001:
                    return
                upd = {"gbLongitude": nv}
                fld = "经度"
            elif col == "#7":  # 纬度
                try:
                    nv = float(new_val)
                except ValueError:
                    messagebox.showerror("错误", "纬度必须是数字", parent=win)
                    return
                old = float(ch.get("gbLatitude", 0) or 0)
                if abs(nv - old) < 0.000001:
                    return
                upd = {"gbLatitude": nv}
                fld = "纬度"
            elif col == "#8":  # 厂家
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
            if col in ("#1", "#2", "#4", "#9"):
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
            hdrs = ["设备名称", "设备ID", "通道名称", "通道类型", "区域编码",
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
                       ch.get("civilCode", ""),
                       ch.get("gbLongitude", 0), ch.get("gbLatitude", 0),
                       ch.get("gbManufacturer", ""), ch.get("id", "")]
                for j, v in enumerate(row, 1):
                    ws.cell(row=i + 1, column=j, value=v)
            for i, w in enumerate([18, 24, 25, 10, 15, 10, 10, 15, 10], 1):
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
                if not row or len(row) < 9:
                    continue
                did = row[8]
                if not did:
                    continue
                try:
                    did = int(did)
                except:
                    continue
                excel_data[did] = {
                    "name": str(row[2] or "").strip(),
                    "civilCode": str(row[4] or "").strip(),
                    "gbLongitude": str(row[5] or "").strip(),
                    "gbLatitude": str(row[6] or "").strip(),
                    "gbManufacturer": str(row[7] or "").strip(),
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
        """窗口内批量提交导入的修改"""
        host = self.server_host.get().strip().rstrip('/')
        headers = {"Accept": "*/*", "access-token": self.access_token,
                   "Content-Type": "application/json"}
        success = fail = 0
        total = len(tasks)
        for idx, (ch, nn, nc, lo, la, mf) in enumerate(tasks, 1):
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
                success += 1
                continue
            try:
                body = self.build_channel_body(ch, upd)
                resp = requests.post(f"{host}/api/common/channel/update",
                                     headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    success += 1
                    for k, v in upd.items():
                        ch[k] = v
                else:
                    fail += 1
            except Exception:
                fail += 1
            self.root.after(0, lambda c=idx, t=total:
                set_st(f"导入中 ({c}/{t})..."))
        self.root.after(0, lambda: set_st(f"导入完成: 成功 {success}, 失败 {fail}"))
        self.root.after(0, lambda: messagebox.showinfo("结果",
            f"成功: {success} 条\n失败: {fail} 条", parent=win))
        if success > 0:
            self.root.after(0, lambda: self._refresh_win_tree(tree, channels, item_to_ch))

    def _win_refresh(self, win, device_id, device_name,
                     tree, item_to_ch, check_vars, set_st):
        """重新查询设备通道并刷新窗口表格"""
        set_st("刷新中...")
        threading.Thread(target=self._do_win_refresh,
            args=(win, device_id, device_name,
                  tree, item_to_ch, check_vars, set_st), daemon=True).start()

    def _do_win_refresh(self, win, device_id, device_name,
                        tree, item_to_ch, check_vars, set_st):
        """后台重新查询并刷新表格"""
        host = self.server_host.get().strip().rstrip('/')
        token = self.access_token
        new_channels = []
        try:
            headers = {"Accept": "*/*", "access-token": token}
            page = 1
            while True:
                url = f"{host}/api/device/query/devices/{device_id}/channels"
                resp = requests.get(url, headers=headers, params={
                    "page": page, "count": 100,
                    "query": "", "cameraQuery": "", "nvrQuery": ""
                }, timeout=30)
                if resp.status_code == 200:
                    data = resp.json()
                    if data.get("code") == 0:
                        lst = data.get("data", {}).get("list", [])
                        new_channels.extend(lst)
                        if len(lst) < 100:
                            break
                        page += 1
                    else:
                        break
                else:
                    break
            # 合并全局通道数据
            for dc in new_channels:
                dc_id = dc.get("id") or dc.get("gbId")
                if not dc_id:
                    continue
                try:
                    one_resp = requests.get(
                        f"{host}/api/common/channel/one",
                        headers=headers, params={"id": dc_id}, timeout=15)
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
                tree.set(item, "#5", ch.get("civilCode", ""))
                tree.set(item, "#6", ch.get("gbLongitude", 0))
                tree.set(item, "#7", ch.get("gbLatitude", 0))
                tree.set(item, "#8", ch.get("gbManufacturer", ""))

    # ========== 通道查询 ==========

    def do_query_channels(self):
        if not self.access_token:
            messagebox.showwarning("警告", "请先登录")
            return
        if not self.selected_device_id:
            messagebox.showwarning("提示", "请先从左侧设备列表中选择一个设备")
            return
        device_id = self.selected_device_id
        self.set_statusbar(f"正在查询设备 {device_id} 的全部通道...")
        self.query_btn.configure(state=tk.DISABLED, text="查询中...")
        self.cancel_edit()

        def task():
            try:
                host = self.server_host.get().strip().rstrip('/')
                url = f"{host}/api/device/query/devices/{device_id}/channels"
                headers = {
                    "Accept": "*/*",
                    "access-token": self.access_token,
                    "Content-Type": "application/x-www-form-urlencoded"
                }
                all_channels = []
                page = 1
                count = self.page_size       # 每页请求条数，默认100
                total = 0

                while True:
                    params = {
                        "page": page,
                        "count": count,
                        "query": "",
                        "cameraQuery": "",
                        "nvrQuery": ""
                    }
                    resp = requests.get(url, headers=headers, params=params, timeout=30)

                    if resp.status_code == 200:
                        data = resp.json()
                        if data.get("code") != 0:
                            self.root.after(0, lambda m=data.get("msg", "未知错误"): self._query_fail(m))
                            return

                        inner = data.get("data", {})
                        channels = inner.get("list", [])
                        total = inner.get("total", 0)
                        all_channels.extend(channels)

                        # 更新界面状态提示
                        self.root.after(0, lambda p=page: self.set_statusbar(f"正在加载通道第 {p} 页..."))

                        # 判断是否已取完所有数据
                        if len(channels) < count or len(all_channels) >= total:
                            break
                        page += 1
                    else:
                        msg = resp.json().get("msg", f"HTTP {resp.status_code}")
                        self.root.after(0, lambda m=msg: self._query_fail(m))
                        return

                # 逐个查询全局通道，合并厂家/经纬度等字段
                try:
                    merged_count = 0
                    for dc in all_channels:
                        dc_id = dc.get("id") or dc.get("gbId")
                        if not dc_id:
                            continue
                        one_url = f"{host}/api/common/channel/one"
                        one_resp = requests.get(one_url, headers=headers,
                                                params={"id": dc_id}, timeout=15)
                        if one_resp.status_code == 200:
                            one_data = one_resp.json()
                            if one_data.get("code") == 0:
                                gc = one_data.get("data")
                                if gc and isinstance(gc, dict):
                                    for fld in ("gbManufacturer", "gbLongitude", "gbLatitude",
                                                "gbName", "gbCivilCode"):
                                        if fld in gc and gc[fld] is not None:
                                            dc[fld] = gc[fld]
                                    merged_count += 1
                    print(f"[DEBUG] 合并了 {merged_count}/{len(all_channels)} 条全局通道数据")
                except Exception as e:
                    import traceback
                    print(f"[DEBUG] 查询全局通道异常: {e}")
                    traceback.print_exc()

                self.all_channels = all_channels
                self.total_channels = total
                self.root.after(0, lambda: self._query_success(all_channels, total, device_id))

            except Exception as e:
                self.root.after(0, lambda e=e: self._query_fail(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _query_success(self, channels, total, device_id):
        self.query_btn.configure(state=tk.NORMAL, text=" 查询通道")
        self.tree.delete(*self.tree.get_children())
        self.item_to_channel.clear()
        self.check_vars.clear()
        self.select_all_var.set(False)

        for i, ch in enumerate(channels, 1):
            db_id = ch.get("id")
            ch_type = ch.get("channelType", 0)
            type_text = "子目录" if ch_type else "设备通道"
            var = tk.BooleanVar(value=False)
            values = ("☐", i, ch.get("deviceId", ""), ch.get("name", ""), type_text, ch.get("civilCode", ""), ch.get("gbLongitude", 0), ch.get("gbLatitude", 0), ch.get("gbManufacturer", ""), ch.get("id", ""))
            item = self.tree.insert("", tk.END, values=values)
            self.item_to_channel[item] = ch
            self.check_vars[item] = var
            var.trace_add("write", lambda *args, it=item: self.update_check_display(it))

        self.page_info_var.set(f"共 {total} 条")
        self.set_statusbar(f"查询成功 - 设备 {device_id} 下共 {total} 个通道")
        self.update_ui_state()

    def _query_fail(self, msg):
        self.query_btn.configure(state=tk.NORMAL, text=" 查询通道")
        self.set_statusbar(f"查询失败: {msg}")
        messagebox.showerror("查询失败", msg)

    # ---------- 复选框 ----------
    def update_check_display(self, item):
        var = self.check_vars.get(item)
        if var:
            self.tree.set(item, "#1", "☑" if var.get() else "☐")

    def on_checkbox_click(self, event):
        if self.tree.identify_column(event.x) != "#1":
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        var = self.check_vars.get(item)
        if var:
            var.set(not var.get())
            self.update_select_all_state()

    def toggle_select_all(self):
        state = self.select_all_var.get()
        for var in self.check_vars.values():
            var.set(state)

    def update_select_all_state(self):
        if not self.check_vars:
            self.select_all_var.set(False)
            return
        all_checked = all(v.get() for v in self.check_vars.values())
        self.select_all_var.set(all_checked)

    def get_selected_channels(self):
        return [self.item_to_channel[it] for it, var in self.check_vars.items()
                if var.get() and it in self.item_to_channel]

    # ---------- 批量修改区域编码 ----------
    def batch_modify_region(self):
        selected = self.get_selected_channels()
        if not selected:
            messagebox.showwarning("提示", "请至少勾选一个通道")
            return
        new_region = simpledialog.askstring("批量修改区域编码",
                                            f"已选中 {len(selected)} 个通道\n请输入新的区域编码:",
                                            parent=self.root)
        if not new_region or not new_region.strip():
            return
        new_region = new_region.strip()
        if not messagebox.askyesno("确认", f"将为选中的 {len(selected)} 个通道设置区域编码为:\n'{new_region}'?\n\n确认修改?"):
            return
        self.batch_region_btn.configure(state=tk.DISABLED)
        self.set_statusbar("正在批量修改区域编码...")
        threading.Thread(target=self._batch_update_region, args=(selected, new_region), daemon=True).start()

    def _batch_update_region(self, channels, new_region):
        success = fail = 0
        host = self.server_host.get().strip().rstrip('/')
        headers = {"Accept": "*/*", "access-token": self.access_token, "Content-Type": "application/json"}
        for idx, ch in enumerate(channels, 1):
            updates = {"gbCivilCode": new_region, "civilCode": new_region}
            try:
                body = self.build_channel_body(ch, updates)
                resp = requests.post(f"{host}/api/common/channel/update", headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    success += 1
                else:
                    fail += 1
            except:
                fail += 1
            self.root.after(0, lambda c=idx: self.set_statusbar(f"批量修改中: {c}/{len(channels)}"))
        self.root.after(0, lambda: self._batch_finished(success, fail))

    @staticmethod
    def _val(v, default=""):
        """取非None的值，None转默认值"""
        return v if v is not None else default

    @staticmethod
    def _num(v, default=0):
        """取非None的数值"""
        if v is None:
            return default
        try:
            return float(v)
        except (ValueError, TypeError):
            return default

    def build_channel_body(self, channel, updates):
        """构建 body：gbId + 有值的已有字段 + 要更新的字段
        注意：设备通道查询返回短名称(id, deviceId, name, civilCode)，
        但全局通道 API 需要 gb 前缀(gbId, gbDeviceId, gbName, gbCivilCode)，
        这里做字段映射。
        """
        # gbId：兼容设备通道(id)和全局通道(gbId)两种字段名
        gb_id_val = channel.get("id") or channel.get("gbId") or 0
        print(f"[DEBUG] build_channel_body: gbId={gb_id_val!r} (type={type(gb_id_val).__name__})")
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
        print(f"[DEBUG] build_channel_body 完成({len(body)}个字段): body={body}")
        return body

    def do_update(self, channel, updates, new_value, item, column):
        try:
            host = self.server_host.get().strip().rstrip('/')
            headers = {"Accept": "*/*", "access-token": self.access_token, "Content-Type": "application/json"}
            body = self.build_channel_body(channel, updates)
            url = f"{host}/api/common/channel/update"
            print(f"[DEBUG] ===== 开始修改通道 =====")
            print(f"[DEBUG] 请求URL: {url}")
            print(f"[DEBUG] 更新的字段: {updates}")
            print(f"[DEBUG] 完整body: {body}")
            resp = requests.post(url, headers=headers, json=body, timeout=15)
            print(f"[DEBUG] 响应状态码: {resp.status_code}")
            print(f"[DEBUG] 响应内容: {resp.text}")
            
            if resp.status_code == 200:
                resp_data = resp.json()
                if resp_data.get("code") == 0:
                    # 更新树显示
                    self.root.after(0, lambda: self.tree.set(item, column, new_value))
                    # 更新本地 channel dict
                    dc = self.item_to_channel.get(item)
                    if dc:
                        for k, v in updates.items():
                            dc[k] = v
                    self.root.after(0, lambda: self.set_statusbar("修改成功"))
                else:
                    msg = resp_data.get("msg", f"code={resp_data.get('code')}")
                    self.root.after(0, lambda m=msg: messagebox.showerror("修改失败", m))
            else:
                msg = f"HTTP {resp.status_code}"
                try:
                    msg = resp.json().get("msg", msg)
                except:
                    pass
                self.root.after(0, lambda m=msg: messagebox.showerror("修改失败", m))
        except Exception as e:
            import traceback
            print(f"[DEBUG] 修改异常: {e}")
            traceback.print_exc()
            self.root.after(0, lambda e=e: messagebox.showerror("修改异常", str(e)))

    def _batch_finished(self, success, fail):
        self.batch_region_btn.configure(state=tk.NORMAL)
        self.set_statusbar(f"批量修改完成: 成功 {success}, 失败 {fail}")
        messagebox.showinfo("结果", f"成功: {success}\n失败: {fail}")
        if success > 0:
            self.do_query_channels()

    # ---------- 双击编辑 ----------
    def on_double_click(self, event):
        col = self.tree.identify_column(event.x)
        if col == "#1":
            return
        if self.edit_entry:
            self.save_edit()
        item = self.tree.identify_row(event.y)
        if not item or col not in ("#4", "#6", "#7", "#8", "#9"):
            return
        self.edit_item = item
        self.edit_column = col
        value = self.tree.set(item, col)
        bbox = self.tree.bbox(item, col)
        if not bbox:
            return
        x, y, w, h = bbox
        self.edit_entry = tk.Entry(self.tree, font=("Microsoft YaHei", 9))
        self.edit_entry.place(x=x, y=y, width=w, height=h)
        self.edit_entry.insert(0, value)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus_set()
        self.edit_entry.bind("<Return>", lambda e: self.save_edit())
        self.edit_entry.bind("<FocusOut>", self.on_focus_out)

    def on_focus_out(self, event):
        if self.edit_entry:
            self.root.after(100, self.save_edit)

    def save_edit(self):
        if not self.edit_entry:
            return
        new_value = self.edit_entry.get().strip()
        item, col = self.edit_item, self.edit_column
        self.cancel_edit()
        if not item or not new_value:
            return
        channel = self.item_to_channel.get(item)
        if not channel:
            return
        if col == "#4":
            old = channel.get("name", "")
            if new_value == old:
                return
            updates = {"gbName": new_value}
            field = "名称"
        elif col == "#6":
            old = channel.get("civilCode", "")
            if new_value == old:
                return
            updates = {"gbCivilCode": new_value}
            field = "区域编码"
        elif col == "#7":
            try:
                new_val = float(new_value)
            except ValueError:
                self.root.after(0, lambda: messagebox.showerror("错误", "经度必须是数字"))
                return
            old = channel.get("gbLongitude", 0)
            if old == "" or old is None:
                old = 0
            try:
                old = float(old)
            except (ValueError, TypeError):
                old = 0.0
            if abs(new_val - old) < 0.000001:
                return
            updates = {"gbLongitude": new_val}
            field = "经度"
        elif col == "#8":
            try:
                new_val = float(new_value)
            except ValueError:
                self.root.after(0, lambda: messagebox.showerror("错误", "纬度必须是数字"))
                return
            old = channel.get("gbLatitude", 0)
            if old == "" or old is None:
                old = 0
            try:
                old = float(old)
            except (ValueError, TypeError):
                old = 0.0
            if abs(new_val - old) < 0.000001:
                return
            updates = {"gbLatitude": new_val}
            field = "纬度"
        else:
            old = channel.get("gbManufacturer", "")
            if old is None:
                old = ""
            if new_value == old:
                return
            updates = {"gbManufacturer": new_value}
            field = "厂家"
        if not messagebox.askyesno("确认", f"将 {field} 从 '{old}' 改为 '{new_value}'?"):
            return
        threading.Thread(target=self.do_update, args=(channel, updates, new_value, item, col), daemon=True).start()

    def cancel_edit(self):
        if self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
            self.edit_item = None
            self.edit_column = None



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
        file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel", "*.xlsx")],
                                             title="导出选中通道",
                                             initialfile="选中通道.xlsx")
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
                page = 1
                while True:
                    try:
                        url = f"{host}/api/device/query/devices/{device_id}/channels"
                        resp = requests.get(url, headers=headers, params={
                            "page": page, "count": 100,
                            "query": "", "cameraQuery": "", "nvrQuery": ""
                        }, timeout=30)
                        if resp.status_code == 200:
                            data = resp.json()
                            if data.get("code") == 0:
                                channels = data.get("data", {}).get("list", [])
                                # 合并全局通道数据
                                for dc in channels:
                                    dc_id = dc.get("id") or dc.get("gbId")
                                    if dc_id:
                                        try:
                                            one = requests.get(
                                                f"{host}/api/common/channel/one",
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
                                if len(channels) < 100:
                                    break
                                page += 1
                            else:
                                break
                        else:
                            break
                    except Exception:
                        break
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
        """批量提交导入的修改（直接查全局通道 API，不依赖本地缓存）"""
        try:
            host = self.server_host.get().strip().rstrip('/')
            token = self.access_token
            headers = {"Accept": "*/*", "access-token": token,
                       "Content-Type": "application/json"}
            success = fail = 0
            total = len(updates_by_id)
            for idx, (db_id, (new_name, new_civil, new_lon, new_lat, new_mfr)) in enumerate(
                    updates_by_id.items(), 1):
                self.root.after(0, lambda c=idx, t=total:
                    self.set_statusbar(f"导入中 ({c}/{t})..."))
                # 从全局通道查询当前数据
                try:
                    one = requests.get(f"{host}/api/common/channel/one",
                                       headers=headers,
                                       params={"id": db_id}, timeout=15)
                    if one.status_code != 200:
                        fail += 1
                        continue
                    od = one.json()
                    if od.get("code") != 0:
                        fail += 1
                        continue
                    ch = od.get("data")
                    if not ch or not isinstance(ch, dict):
                        fail += 1
                        continue
                except Exception:
                    fail += 1
                    continue
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
                    success += 1
                    continue
                try:
                    body = self.build_channel_body(ch, updates)
                    resp = requests.post(f"{host}/api/common/channel/update",
                                         headers=headers, json=body, timeout=15)
                    if resp.status_code == 200 and resp.json().get("code") == 0:
                        success += 1
                    else:
                        fail += 1
                except Exception:
                    fail += 1
            self.root.after(0, lambda: messagebox.showinfo(
                "结果", f"成功: {success} 条\n失败: {fail} 条"))
            self.root.after(0, lambda: self.set_statusbar(
                f"导入完成: 成功 {success}, 失败 {fail}"))
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda e=e: messagebox.showerror("导入异常", str(e)))

    # ---------- 导出Excel ----------
    def export_excel(self):
        if not self.all_channels:
            messagebox.showwarning("提示", "通道列表为空")
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("缺少库", "请安装 openpyxl: pip install openpyxl")
            return
        default_name = f"{self.selected_device_name or '通道列表'}.xlsx"
        file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel", "*.xlsx")],
                                             title="导出通道列表",
                                             initialfile=default_name)
        if not file:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "通道列表"
            headers = ["序号", "设备ID", "名称", "通道类型", "区域编码", "经度", "纬度", "厂家", "数据库ID"]
            for c, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=c, value=h)
                cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            for i, ch in enumerate(self.all_channels, 1):
                row = [i, ch.get("deviceId", ""), ch.get("name", ""),
                       "子目录" if ch.get("channelType") else "设备通道",
                       ch.get("civilCode", ""), ch.get("gbLongitude", 0), ch.get("gbLatitude", 0), ch.get("gbManufacturer", ""), ch.get("id", "")]
                for j, v in enumerate(row, 1):
                    ws.cell(row=i+1, column=j, value=v)
            for i, w in enumerate([6, 22, 25, 10, 15, 10, 10, 15, 10], 1):
                ws.column_dimensions[ws.cell(1, i).column_letter].width = w
            wb.save(file)
            self.set_statusbar(f"导出成功: {os.path.basename(file)}")
            messagebox.showinfo("成功", f"已导出到:\n{file}")
        except Exception as e:
            messagebox.showerror("失败", str(e))

    # ---------- 导入Excel（带进度条）----------
    def import_excel(self):
        if not self.access_token:
            messagebox.showwarning("警告", "请先登录")
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("缺少库", "请安装 openpyxl: pip install openpyxl")
            return
        if not self.all_channels:
            messagebox.showwarning("提示", "请先查询通道再导入")
            return
        file = filedialog.askopenfilename(title="选择修改后的Excel", filetypes=[("Excel", "*.xlsx")])
        if not file:
            return
        try:
            wb = load_workbook(file)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        excel = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 9:
                continue
            did = row[8]
            if not did:
                continue
            try:
                did = int(did)
            except:
                continue
            excel[did] = {"name": str(row[2] or "").strip(), "civilCode": str(row[4] or "").strip(), "gbLongitude": str(row[5] or "").strip(), "gbLatitude": str(row[6] or "").strip(), "gbManufacturer": str(row[7] or "").strip()}
        if not excel:
            messagebox.showwarning("无数据", "Excel无有效数据")
            return
        tasks = []
        for ch in self.all_channels:
            if ch["id"] in excel:
                n = excel[ch["id"]]["name"]
                c = excel[ch["id"]]["civilCode"]
                lo = excel[ch["id"]]["gbLongitude"]
                la = excel[ch["id"]]["gbLatitude"]
                mf = excel[ch["id"]]["gbManufacturer"]
                if (n != ch.get("name", "") or c != ch.get("civilCode", "") or
                    (lo and str(lo) != str(ch.get("gbLongitude", 0))) or
                    (la and str(la) != str(ch.get("gbLatitude", 0))) or
                    mf != ch.get("gbManufacturer", "")):
                    tasks.append((ch, n, c, lo, la, mf))
        if not tasks:
            messagebox.showinfo("提示", "没有需要修改的数据")
            return
        if not messagebox.askyesno("确认", f"检测到 {len(tasks)} 条修改，是否继续?"):
            return
        self.progress = ProgressDialog(self.root, "正在导入修改", len(tasks))
        self.import_btn.configure(state=tk.DISABLED)
        threading.Thread(target=self.batch_excel_update, args=(tasks,), daemon=True).start()

    def batch_excel_update(self, tasks):
        success = fail = 0
        host = self.server_host.get().strip().rstrip('/')
        headers = {"Accept": "*/*", "access-token": self.access_token, "Content-Type": "application/json"}
        total = len(tasks)
        for i, (ch, nn, nc, lo, la, mf) in enumerate(tasks, 1):
            if self.progress and self.progress.is_cancelled():
                break
            upd = {}
            if nn != ch.get("name", ""):
                upd["gbName"] = nn
            if nc != ch.get("civilCode", ""):
                upd["gbCivilCode"] = nc
            if lo and str(lo) != str(ch.get("gbLongitude", 0)):
                try:
                    val = float(lo)
                    upd["gbLongitude"] = val
                except ValueError:
                    pass
            if la and str(la) != str(ch.get("gbLatitude", 0)):
                try:
                    val = float(la)
                    upd["gbLatitude"] = val
                except ValueError:
                    pass
            if mf != ch.get("gbManufacturer", ""):
                upd["gbManufacturer"] = mf
            try:
                body = self.build_channel_body(ch, upd)
                resp = requests.post(f"{host}/api/common/channel/update", headers=headers, json=body, timeout=15)
                if resp.status_code == 200 and resp.json().get("code") == 0:
                    success += 1
                else:
                    fail += 1
            except:
                fail += 1
            self.root.after(0, lambda c=i, t=total: self.progress.update(c, t, f"正在处理 {c}/{t}"))
        self.root.after(0, self.progress.close)
        self.root.after(0, lambda: self._batch_finished(success, fail))
        self.root.after(0, lambda: self.import_btn.configure(state=tk.NORMAL))

    # ---------- 退出 ----------
    def logout(self):
        if messagebox.askyesno("退出", "确定要退出登录吗？"):
            self.root.destroy()

    def on_close(self):
        self.cancel_edit()
        self.root.destroy()
