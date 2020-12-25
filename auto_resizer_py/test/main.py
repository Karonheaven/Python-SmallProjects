#!/usr/bin/env python
# coding:utf-8
"""
# Auto Resize Program Windows
# Author: Karonheaven
"""
# +---------------+需要的包+---------------+
# 标准库导入
from typing import *
import win32con
import win32ui
import win32gui
import win32api
import win32process
import os
import sys
import time

# 第三方库导入
import psutil
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import *


# 本地库/自定义库导入


# +---------------+全局变量&预定义+---------------+
def get_windows(windowsname, filename):
    # 获取窗口句柄，如果你的输入参数是handle的话直接跳过这一步
    handle = win32gui.FindWindow(None, windowsname)
    # 获取窗口进程
    threadpid, procpid = win32process.GetWindowThreadProcessId(handle)
    pp = psutil.Process(procpid)
    # 可以参考psutil获取进程详细信息
    print(pp.exe(), pp.name(), pp.status())
    # 将窗口放在前台，并激活该窗口（窗口不能最小化）
    # win32gui.SetForegroundWindow(handle)
    # 获取窗口DC
    hdDC = win32gui.GetWindowDC(handle)
    # 根据句柄创建一个DC
    newhdDC = win32ui.CreateDCFromHandle(hdDC)
    # 创建一个兼容设备内存的DC
    saveDC = newhdDC.CreateCompatibleDC()
    # 创建bitmap保存图片
    saveBitmap = win32ui.CreateBitmap()
    
    # 获取窗口的位置信息
    left, top, right, bottom = win32gui.GetWindowRect(handle)
    # 窗口长宽
    width = right - left
    height = bottom - top
    # bitmap初始化
    saveBitmap.CreateCompatibleBitmap(newhdDC, width, height)
    saveDC.SelectObject(saveBitmap)
    saveDC.BitBlt((0, 0), (width, height), newhdDC, (0, 0), win32con.SRCCOPY)
    saveBitmap.SaveBitmapFile(saveDC, filename)
    # 释放内存
    win32gui.DeleteObject(saveBitmap.GetHandle())
    saveDC.DeleteDC()
    # 杀死该进程及关闭窗口
    os.popen('taskkill.exe /F /pid:' + str(procpid))


def screenshot(handle, filename):
    app = QApplication(sys.argv)
    screen = QApplication.primaryScreen()
    img = screen.grabWindow(handle).toImage()
    img.save(filename)


def get_window_from_cursor():
    prevWindow = None
    curX, curY = win32gui.GetCursorPos()
    handle = win32gui.WindowFromPoint((curX, curY))
    prevWindow = handle
    window_name = win32gui.GetWindowText(handle)
    r = get_windowsinfo(window_name)
    print(r)
    hightligt_window(handle)
    screenshot(handle, 'pp.png')
    refreshWindow(handle)


def kill_pro(handle):
    threadpid, procpid = win32process.GetWindowThreadProcessId(handle)
    os.popen('taskkill.exe /F /pid:' + str(procpid))
    # 杀死该进程及关闭窗口


def hightligt_window(handle):
    left, top, right, bottom = win32gui.GetWindowRect(handle)
    rectanglePen = win32gui.CreatePen(win32con.PS_SOLID, 3, win32api.RGB(255, 0, 0))
    windowDc = win32gui.GetWindowDC(handle)
    if windowDc:
        prevPen = win32gui.SelectObject(windowDc, rectanglePen)
        prevBrush = win32gui.SelectObject(windowDc, win32gui.GetStockObject(win32con.HOLLOW_BRUSH))
        
        win32gui.Rectangle(windowDc, 0, 0, right - left, bottom - top)
        win32gui.SelectObject(windowDc, prevPen)
        win32gui.SelectObject(windowDc, prevBrush)
        win32gui.ReleaseDC(handle, windowDc)


# 刷新窗口的函数：

def refreshWindow(handle):
    win32gui.InvalidateRect(handle, None, True)
    win32gui.UpdateWindow(handle)
    win32gui.RedrawWindow(handle,
                          None,
                          None,
                          win32con.RDW_FRAME |
                          win32con.RDW_INVALIDATE |
                          win32con.RDW_UPDATENOW |
                          win32con.RDW_ALLCHILDREN)


# +---------------+主程序+---------------+
pass
