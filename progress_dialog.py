#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
进度弹窗 - 用于导入Excel时显示处理进度
"""

import tkinter as tk
from tkinter import ttk


class ProgressDialog:
    def __init__(self, parent, title="进度", total=0):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("420x160")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()

        self.top.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 420) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 160) // 2
        self.top.geometry(f"+{x}+{y}")

        style = ttk.Style()
        style.configure("green.Horizontal.TProgressbar", troughcolor='#E0E0E0',
                        background='#4CAF50', thickness=18)

        main_frame = ttk.Frame(self.top, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.label = ttk.Label(main_frame, text="正在准备导入...", font=("Microsoft YaHei", 10))
        self.label.pack(pady=(0, 10))

        self.progress = ttk.Progressbar(main_frame, style="green.Horizontal.TProgressbar",
                                        length=360, mode='determinate', maximum=total)
        self.progress.pack(pady=(0, 5))

        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X)

        self.percent_var = tk.StringVar(value="0%")
        ttk.Label(info_frame, textvariable=self.percent_var, font=("Microsoft YaHei", 9, "bold"),
                  foreground="#4CAF50").pack(side=tk.LEFT)

        self.detail = tk.StringVar(value=f"0 / {total}")
        ttk.Label(info_frame, textvariable=self.detail, font=("Microsoft YaHei", 9)).pack(side=tk.RIGHT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(10, 0))
        ttk.Button(btn_frame, text="取消", command=self.on_cancel).pack()

        self._cancelled = False
        self.top.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def update(self, current, total, text=None):
        if self._cancelled:
            return
        self.progress['value'] = current
        percent = int((current / total) * 100) if total > 0 else 0
        self.percent_var.set(f"{percent}%")
        self.detail.set(f"{current} / {total}")
        if text:
            self.label.config(text=text)
        self.top.update_idletasks()

    def is_cancelled(self):
        return self._cancelled

    def on_cancel(self):
        self._cancelled = True
        try:
            self.top.destroy()
        except:
            pass

    def close(self):
        if not self._cancelled:
            self.top.destroy()
