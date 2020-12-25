#!/usr/bin/env python
# coding:utf-8
"""
# 遍历所有的窗口
# Author: Karonheaven
"""
# +---------------+需要的包+---------------+
# 标准库导入
from typing import *
import os

# 第三方库导入
import win32con
import win32ui
import win32gui
import win32api
import win32process

import psutil


# 本地库/自定义库导入


# +---------------+全局变量&预定义+---------------+
class Ergodic():
    def __init__(self):
        self.hwnd_title = dict()
    
    def get_all_hwnd(self, hwnd, mouse):
        # if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            self.hwnd_title.update(
                {hwnd: win32gui.GetWindowText(hwnd) + ';' + win32gui.GetClassName(hwnd)})
    
    """
    def get_all_hwnd1(self, hwnd, mouse):
        # if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            self.windows.update({hwnd: win32gui.GetWindowText(hwnd) + ';' + win32gui.GetClassName(hwnd)})

    def get_child_window(self,hwnd):
        self.windows = dict()
        win32gui.EnumChildWindows(hwnd, self.get_all_hwnd1, self.windows)
        print(self.windows)
        return self.windows
    """
    
    def __call__(self, *args, **kwargs):
        
        win32gui.EnumWindows(self.get_all_hwnd, 0)
        
        for h, t in self.hwnd_title.items():
            if t != "":
                print(h, t)
        return self.hwnd_title


# +---------------+主程序+---------------+
if __name__ == '__main__':
    ergo = Ergodic()
    ergo()
