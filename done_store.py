# -*- coding: utf-8 -*-
"""手动勾选完成状态存储。
Excel 计划表是只读数据源，勾选状态绝不写回 Excel，保存在 data/done_flags.json。
key 口径（common.task_key）：文件名|任务|子任务|计划开始——
修改计划开始日期后 key 变化，需重新勾选。"""
import datetime as dt
import json
import os

from common import resolve, task_key

FLAGS_PATH = "data/done_flags.json"


def flags_path():
    return resolve(FLAGS_PATH)


def load_flags():
    p = flags_path()
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_flags(flags):
    p = flags_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(flags, f, ensure_ascii=False, indent=1)


def mark_done(flags, task):
    """把一条任务标记为已完成（存任务快照，Excel 删行后仍留在已完成清单）"""
    flags[task_key(task)] = {
        "done_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "snapshot": task,
    }


def unmark(flags, key):
    flags.pop(key, None)


def apply_flags(tasks, flags):
    """把勾选记录应用到解析结果：命中 key 的任务强制置为 done。
    同时给每条任务写 done_manual 字段（仅手动勾选为 True，Excel 原生完成为 False）。
    返回命中数量。"""
    n = 0
    for t in tasks:
        k = task_key(t)
        if k in flags:
            t["status"] = "done"
            t["status_raw"] = "已完成(勾选)"
            t["done_manual"] = True
            n += 1
        else:
            t["done_manual"] = False
    return n
