#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WVP国标平台 - 通道管理工具 v1.0
程序入口，负责显示登录窗口，登录成功后启动主界面
"""

import tkinter as tk
from login_window import LoginWindow
from main_app import MainApplication


def main():
    def on_login_success(token, user, host):
        login_window.root.destroy()          # 销毁登录窗口
        root = tk.Tk()
        MainApplication(root, token, user, host)
        root.mainloop()

    login_window = LoginWindow(on_success_callback=on_login_success)
    login_window.root.mainloop()


if __name__ == "__main__":
    main()