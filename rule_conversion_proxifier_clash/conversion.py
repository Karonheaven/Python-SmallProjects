#!/usr/bin/env python
# coding:utf-8
"""
# 转换规则：proxifier→clash
# Author: Karonheaven

转换样例
# 对于domain
domain_name.yaml
payload:
  -



"""
# +---------------+需要的包+---------------+
# 标准库导入
from typing import List, Dict, Tuple, Set
import os
import re

# +---------------+主程序+---------------+
# 全局变量
rule_providers: List[str] = ["rule-providers:\n"]
rules: List[str] = ["rules:\n"]


def fix_illegal_chars(raw_str: str) -> str:
    """
    将Windows文件管理器不支持的非法字符替换为下划线
    
    :param raw_str: 原字符串
    :return: 替换后的字符串
    """
    rstr = r"[\/\\\:\*\?\"\<\>\|]"
    return re.sub(rstr, "_", raw_str)


# def split_domain(raw_str: str) -> Tuple[bool, str]:
#     """
#     区别IP和网址，并返回(是否为domain, domain字符串)
#
#     "www.baidu.com"→(true, "baidu.com")
#
#     "127.0.0.1"→(false, "127.0.0.1")
#
#     :param raw_str: 原始字符串
#     :return: (是否为domain, domain字符串)
#     """
#     is_domain:bool =
#
#     return None


def split_multi_targets(target_str: str) -> List[str]:
    """
    将以分号分隔开的多个目标拆分为列表
    
    :param target_str: 多个目标
    :return: 列表形式的独立目标
    """
    list_raw = target_str.split(";")
    list_target: List[str] = []
    for target in list_raw:
        target = target.strip()
        if target != "":
            list_target.append(target)
    
    return list_target


def get_label_and_data(line_str: str) -> Tuple[str, List[str]]:
    """
    分离标签和内容
    
    :param line_str: 行内容
    :return: (标签, [应用或地址a, 应用或地址b, ...])
    """
    line_str = line_str.strip()
    label_name = line_str[line_str.find("<") + 1: line_str.find(">")].strip()
    label_data = line_str[line_str.find(">") + 1: line_str.find("<", line_str.find(">") + 1)]
    list_data = split_multi_targets(target_str=label_data)
    
    return (label_name, list_data)


def process_and_save(line_enable: str, line_name: str, line_targets: str, line_action: str) -> None:
    """
    传入行数据后进行处理
    
    :param line_enable: <Rule enabled=行
    :param line_name: <Name>行
    :param line_targets: <Applications> or <Targets>行
    :param line_action: <Action type=行
    :return: None，写入文件
    """
    # 提取Name
    tg_name = get_label_and_data(line_str=line_name)[1][0]
    # 判断规则是否启用
    if line_enable.find("true") != -1:
        tg_enabled = 1
    elif line_enable.find("false") != -1:
        tg_enabled = 0
    else:
        raise ValueError("错误的enabled: {}".format(tg_name))
    # 获取targets
    tg_type, tg_list = get_label_and_data(line_str=line_targets)
    # 获取action
    if line_action.find("Direct") != -1:
        tg_action = "DIRECT"
    elif line_action.find("Proxy") != -1:
        tg_action = "Proxy"
    elif line_action.find("Block") != -1:
        tg_action = "REJECT"
    else:
        raise ValueError("错误的action: {}".format(tg_name))
    
    # 判断规则为domain还是process-name
    # 若为domain，则使用domain的behavior
    if tg_type == "Targets":
        
        
        
        pass
    
    if tg_type == "Applications":
        rule_single_app: List[str] = []
        rule_single_app.append("  # {}\n".format(tg_name))
        for target in tg_list:
            rule_single_app.append("  - PROCESS-NAME,{},{}\n".format(target, tg_action))
        # 如果规则未启用，则注释掉该条规则(在行头加#)
        if not tg_enabled:
            for _ in rule_single_app:
                rules.append("#{}".format(_))
        else:
            for _ in rule_single_app:
                rules.append(_)
    
    # 若规则为domain，则创建为一个独立的rule-provider
    # 开始创建文件
    if tg_enabled == 1:
        file_write = open(file="./rules/{}.yaml".format(fix_illegal_chars(raw_str=tg_name)), mode="w+",
                          encoding="utf-8")
    else:
        file_write = open(file="./rules/{}.yaml".format(fix_illegal_chars(raw_str=tg_name)), mode="w+",
                          encoding="utf-8")
    
    # 写入首行
    file_write.write("payload:\n")
    
    return None


# 创建文件夹
if not os.path.exists("./rules-set-enabled"):
    os.makedirs("./rules-set-enabled")
if not os.path.exists("./rules-set-disabled"):
    os.makedirs("./rules-set-disabled")

# 读取.ppx文件
with open(file="./4Share - 10808.ppx", encoding="utf-8") as file_read:
    ppx_data_all = file_read.readlines()

# 通过开头判断是否为.ppx文件
if ppx_data_all[0] != "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n":
    raise ValueError

# 最后将rule-providers和rules写入同一个yaml中
file_write = open(file="./rules-set-enabled/allRules.yaml", mode="w+", encoding="utf-8")
file_write.write("此文件仅为clash最终设置的部分内容，不能单独作为clash的设置文件\n")
file_write.writelines(rule_providers)
file_write.write("\n\n")
file_write.writelines(rules)
file_write.close()
