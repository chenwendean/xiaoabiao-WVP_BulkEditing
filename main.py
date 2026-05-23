#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WVP国标平台 - 通道管理工具 v1.0
程序入口，负责显示登录窗口，登录成功后启动主界面

Copyright (C) 2025 Xiaoabiao

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
"""

import tkinter as tk
from login_window import LoginWindow
from main_app import MainApplication


def main():
    def on_login_success(token, user, host):
        # 清除登录窗口内容，复用同一个 root
        for w in login_window.root.winfo_children():
            w.destroy()
        login_window.root.title("WVP通道管理工具 v7.0 by Xiaoabiao")
        MainApplication(login_window.root, token, user, host)

    login_window = LoginWindow(on_success_callback=on_login_success)
    login_window.root.mainloop()


if __name__ == "__main__":
    main()
