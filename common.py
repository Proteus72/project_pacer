# -*- coding: utf-8 -*-
"""公共工具：配置加载与路径解析（相对路径一律相对项目根目录）"""
import json
import os
import sys

# 打包成 exe（PyInstaller）后 __file__ 指向临时解压目录，必须改用 exe 所在目录，
# 这样 config.json / output/ / data/ 仍然相对 pm_workbench 文件夹解析，
# exe 与源码脚本（run_daily.py 等）始终共享同一份数据。
if getattr(sys, "frozen", False):
    ROOT = os.path.dirname(os.path.abspath(sys.executable))
else:
    ROOT = os.path.dirname(os.path.abspath(__file__))


def load_config():
    """读取 config.json（UTF-8），返回 dict。
    优先读 exe/脚本同目录的 config.json（用户可编辑）；
    找不到时（如 exe 被单独挪走）回退到打包内置的默认配置。"""
    p = os.path.join(ROOT, "config.json")
    if not os.path.exists(p) and getattr(sys, "frozen", False):
        p = os.path.join(sys._MEIPASS, "config.json")
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def resolve(p):
    """绝对路径原样返回；相对路径相对项目根目录解析"""
    if os.path.isabs(p):
        return os.path.normpath(p)
    return os.path.normpath(os.path.join(ROOT, p))


def task_key(t):
    """任务稳定标识：文件名|任务|子任务|计划开始。
    用于桌面小组件勾选状态匹配；修改计划开始日期后 key 变化，需重新勾选。
    文件名取 base_file（**原表**文件名）——读的是回写新版本时文件名带了版本后缀，
    用原表名才能让勾选状态跨版本继续命中。"""
    name = t.get("base_file") or os.path.basename(t.get("source_file", ""))
    return (f"{name}|{t.get('task', '')}|{t.get('subtask') or ''}|{t.get('plan_start', '')}")
