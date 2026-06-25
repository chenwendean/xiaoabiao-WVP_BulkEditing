#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
登录窗口（简洁版）
"""

import tkinter as tk
from tkinter import ttk, messagebox
import hashlib
import requests
import threading
import json
import os
import sys
import base64

def _config_dir():
    """exe 运行时存在 exe 同目录，源码运行时存在脚本同目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(_config_dir(), ".login_config.json")
DEFAULT_HOST = "http://127.0.0.1:18080"
DEFAULT_USERNAME = "admin"


class LoginWindow:
    def __init__(self, on_success_callback):
        self.on_success = on_success_callback
        self.root = tk.Tk()
        self.root.title("WVP通道管理工具 - 登录")
        self.root.geometry("400x280")
        self.root.resizable(False, False)

        # 居中显示
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 400) // 2
        y = (self.root.winfo_screenheight() - 280) // 2
        self.root.geometry(f"+{x}+{y}")

        # 变量
        self.server_host = tk.StringVar(value=DEFAULT_HOST)
        self.username_var = tk.StringVar(value=DEFAULT_USERNAME)
        self.password_var = tk.StringVar(value="")
        self.remember_var = tk.BooleanVar(value=False)

        self.create_widgets()
        self.load_config()

    def create_widgets(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack(padx=25, pady=(20, 10), fill=tk.BOTH, expand=True)

        # 标题
        tk.Label(main_frame, text="WVP 通道管理工具",
                 font=("Microsoft YaHei", 14, "bold"),
                 fg="#1a73e8").grid(row=0, column=0, columnspan=2, pady=(0, 15))

        # 服务器地址
        tk.Label(main_frame, text="服务器地址:", font=("Microsoft YaHei", 9)).grid(
            row=1, column=0, sticky="w", pady=3)
        self.server_entry = tk.Entry(main_frame, textvariable=self.server_host,
                                      font=("Microsoft YaHei", 10))
        self.server_entry.grid(row=1, column=1, pady=3, padx=(5, 0), sticky="ew")

        # 用户名
        tk.Label(main_frame, text="用户名:", font=("Microsoft YaHei", 9)).grid(
            row=2, column=0, sticky="w", pady=3)
        self.user_entry = tk.Entry(main_frame, textvariable=self.username_var,
                                    font=("Microsoft YaHei", 10))
        self.user_entry.grid(row=2, column=1, pady=3, padx=(5, 0), sticky="ew")

        # 密码
        tk.Label(main_frame, text="密码:", font=("Microsoft YaHei", 9)).grid(
            row=3, column=0, sticky="w", pady=3)
        self.pass_entry = tk.Entry(main_frame, textvariable=self.password_var,
                                    show="*", font=("Microsoft YaHei", 10))
        self.pass_entry.grid(row=3, column=1, pady=3, padx=(5, 0), sticky="ew")
        self.pass_entry.bind("<Return>", lambda e: self.do_login())

        # 记住密码
        self.remember_cb = tk.Checkbutton(main_frame, text="记住密码",
                                           font=("Microsoft YaHei", 9),
                                           variable=self.remember_var)
        self.remember_cb.grid(row=4, column=0, columnspan=2, pady=(5, 0), sticky="w")

        # 状态标签
        self.status_label = tk.Label(main_frame, text="", fg="red",
                                      font=("Microsoft YaHei", 9))
        self.status_label.grid(row=5, column=0, columnspan=2, pady=(3, 0), sticky="w")

        # 登录按钮
        self.login_btn = tk.Button(main_frame, text="登录",
                                    font=("Microsoft YaHei", 10, "bold"),
                                    bg="#1a73e8", fg="white",
                                    activebackground="#1557b0", activeforeground="white",
                                    relief="flat", bd=0,
                                    cursor="hand2",
                                    command=self.do_login)
        self.login_btn.grid(row=6, column=0, columnspan=2, pady=(12, 0), ipady=4, sticky="ew")

        # 列权重
        main_frame.columnconfigure(1, weight=1)

    # ---------- 配置文件读写 ----------
    def config_path(self):
        return CONFIG_FILE

    def load_config(self):
        try:
            if os.path.exists(self.config_path()):
                with open(self.config_path(), "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                if cfg.get("host"):
                    self.server_host.set(cfg["host"])
                if cfg.get("username"):
                    self.username_var.set(cfg["username"])
                if cfg.get("password"):
                    try:
                        pwd = base64.b64decode(cfg["password"]).decode("utf-8")
                        self.password_var.set(pwd)
                        self.remember_var.set(True)
                    except:
                        pass
        except Exception:
            pass

    def save_config(self):
        try:
            cfg = {
                "host": self.server_host.get().strip().rstrip('/'),
                "username": self.username_var.get().strip(),
                "password": "",
            }
            if self.remember_var.get():
                cfg["password"] = base64.b64encode(
                    self.password_var.get().strip().encode("utf-8")
                ).decode("ascii")
            with open(self.config_path(), "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ---------- 登录逻辑 ----------
    @staticmethod
    def md5(text):
        return hashlib.md5(text.encode('utf-8')).hexdigest()

    def do_login(self):
        host = self.server_host.get().strip().rstrip('/')
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()

        if not all([host, username, password]):
            self.status_label.config(text="请填写完整信息")
            return

        if not host.startswith(('http://', 'https://')):
            host = 'http://' + host

        self.login_btn.config(state=tk.DISABLED, text="登录中...")
        self.status_label.config(text="正在连接...", fg="blue")

        def task():
            try:
                url = f"{host}/api/user/login"
                resp = requests.get(url, params={
                    "username": username,
                    "password": self.md5(password)
                }, timeout=15)

                if resp.status_code == 200:
                    data = resp.json()
                    token = data.get("accessToken") or resp.headers.get("access-token")
                    if token:
                        self.save_config()
                        self.root.after(0, lambda: self.on_success(token, data, host))
                    else:
                        self.root.after(0, lambda: self.fail("未获取到token"))
                else:
                    msg = resp.json().get("msg", f"HTTP {resp.status_code}")
                    self.root.after(0, lambda m=msg: self.fail(m))
            except Exception as e:
                self.root.after(0, lambda e=e: self.fail(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def fail(self, msg):
        self.login_btn.config(state=tk.NORMAL, text="登录")
        self.status_label.config(text=f"错误: {msg}", fg="red")


if __name__ == "__main__":
    def test_callback(token, data, host):
        print(f"登录成功！Token: {token[:20]}...")
        messagebox.showinfo("成功", "登录成功")

    win = LoginWindow(test_callback)
    win.root.mainloop()
