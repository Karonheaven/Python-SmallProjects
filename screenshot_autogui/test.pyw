#!/usr/bin/env python
# coding:utf-8
"""
# Test
# Author: Karonheaven
"""
# +---------------+需要的包+---------------+
# 标准库导入
from typing import *

# 第三方库导入
import pyautogui
import screenshot_autogui
import pyscreenshot

# 本地库/自定义库导入


# +---------------+全局变量&预定义+---------------+
pyautogui.FAILSAFE = True

# +---------------+主程序+---------------+
if __name__ == "__main__":
    scr_width, scr_height = pyautogui.size()
    print(scr_width, scr_height)
    
    while True:
        current_mouse_x, current_mouse_y = pyautogui.position()
        
        print(current_mouse_x, current_mouse_y)
    
    pass
