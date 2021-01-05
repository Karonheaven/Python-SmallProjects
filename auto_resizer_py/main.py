#!/usr/bin/env python
# coding:utf-8
"""
# 捕获指定进程并重新为应用程序设置位置和大小
# Author: Karonheaven
"""
# +---------------+需要的包+---------------+
# 标准库导入
from typing import *
from time import sleep
from datetime import datetime

import win32con
import win32ui
import win32gui
import win32api
import win32process
import psutil


# 第三方库导入


# 本地库/自定义库导入


# +---------------+TO DO+---------------+
# TODO(karonheaven@gmail.com): 获取屏幕尺寸
# TODO(karonheaven@gmail.com): 根据Class Name List筛选指定进程
# TODO(karonheaven@gmail.com): 根据进程ID调整窗口大小

# +---------------+全局变量&预定义+---------------+
# 调用底层API函数
# user32 = windll.user32
# kernel32 = windll.kernel32
# psapi = windll.psapi

# 获取当前屏幕的有效尺寸



# 实时查看当前窗口
def getProcessInfo():
    """
    获取当前焦点窗口的信息
    
    :return: ？
    """
    # 获取当前焦点窗口的句柄
    hwnd = win32gui.GetForegroundWindow()
    # 获取线程ID和进程ID
    thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
    # print("线程ID：{}".format(thread_id))
    # print("进程ID：{}".format(process_id))
    
    # 获取Process名称和窗体名称
    process = psutil.Process(pid=process_id)
    process_name = process.name()
    # print("进程名：{}".format(process_name))
    
    win_text: str = win32gui.GetWindowText(hwnd)
    # print("窗体名：{}".format(win_text))
    class_name: str = win32gui.GetClassName(hwnd)
    # print("Class Name: {}".format(class_name))
    
    return (process_id, process_name, win_text, class_name)


def monitor(target_classname_list:List) -> None:
    """
    监控焦点进程的Class Name，根据Target ClassName List判断调整的大小和位置
    
    :param target_classname_list:
    :return:
    """
    return None


def main():
    previous_process_info: Tuple = ("Process ID", "Process Name", "Windows Name", "Class Name")
    front_process_info: Tuple = ("Process ID", "Process Name", "Windows Name", "Class Name")
    
    getProcessInfo()
    
    while True:
        # 如果切换焦点则输出一份
        front_process_info = getProcessInfo()
        
        if front_process_info[0] != previous_process_info[0]:
            print("+---------------+{}+---------------+".format(datetime.now()))
            
            print("Transfer from: ")
            print("    {}".format(previous_process_info))
            print("To: ")
            print("    {}".format(front_process_info))
            
            previous_process_info = front_process_info
        
        sleep(0.1)
        
        pass
    
    # global win_ids
    # win_ids = [None, None]
    #
    # while True:
    #     getProcessInfo()
    #
    #     # 如果用户切换窗口则进行提示
    #     if win_ids[0] != win_ids[1]:
    #         print('=' * 30)
    #         print(str(datetime.now())[:19], win_ids[0], '==>', win_ids[1])
    #         sleep(0.2)


# +---------------+主程序+---------------+
if __name__ == "__main__":
    main()

# hwnd = win32gui.GetForegroundWindow()
# print(hwnd)
# pid = c_ulong(0)
# print(pid, type(pid))
