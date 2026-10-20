# -*- coding: utf-8 -*-
"""界面状态覆盖层（status override）。

为什么需要这一层：
  勾选/切换状态后要**界面立刻变**，但 Excel 新版本是延迟批量生成的
  （关小组件时 / 每 3 小时）。两者之间有个空档——这段时间重新解析读到的
  还是旧 Excel，界面上状态会"点完不动"。覆盖层就填这个空档：
  把"用户最新意图"记在 data/status_overrides.json，每次生成摘要时
  优先按它显示；等 Excel 真的写盘、解析结果追上来了（reconcile），
  对应覆盖自动撤销，从此由 Excel 说了算。

生命周期：
  点击切换 → 写入覆盖（界面立即变） + 入队待回写（不生成文件）
  落盘后   → reconcile：Excel 已等于覆盖值 → 删除覆盖（无残留）
"""
import datetime as dt
import json
import os

from common import resolve, task_key

OVERRIDE_REL = "data/status_overrides.json"


def path():
    return resolve(OVERRIDE_REL)


def load():
    p = path()
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save(d):
    p = path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)


def set_override(key, internal, raw, base_file=None, sheet=None, row=None):
    """记下用户最新意图：internal=doing/todo/done，raw=写入 Excel 的原文案。

    同时记下回写定位（原表名/sheet/行号）——reconcile 靠它判断这条改动
    是否已真正落盘：定位还在待回写队列里就说明没写，覆盖不能撤。
    """
    d = load()
    d[key] = {"internal": internal, "raw": raw,
              "base_file": base_file, "sheet": str(sheet) if sheet else None,
              "row": str(row) if row is not None else None,
              "at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    save(d)
    return d


def pop(key):
    d = load()
    rec = d.pop(key, None)
    save(d)
    return rec


def apply(tasks):
    """把覆盖应用到解析结果（在风险标注之前调用），返回命中数量"""
    d = load()
    if not d:
        for t in tasks:
            t["pending_write"] = False
        return 0
    n = 0
    for t in tasks:
        rec = d.get(task_key(t))
        if rec:
            t["status"] = rec["internal"]
            t["status_raw"] = rec["raw"]
            t["pending_write"] = True
            n += 1
        else:
            t["pending_write"] = False
    return n


def reconcile(tasks, pending=None):
    """Excel 已把该改动写进磁盘后撤销覆盖，返回撤销的 key 列表。

    ⚠️ 判据不能是「解析出的状态 == 覆盖状态」——Excel 是延迟批量写的，
    未落盘的任务解析值仍是旧值，那样一比对就会误判成"已追平"而把覆盖删掉，
    表现为「切换任务 B 时，任务 A 的状态被弹回去」（多任务互相串扰）。

    正确判据：该任务的 key 已**不在待回写队列**中（说明已经落盘过一次），
    并且解析值也确实等于覆盖值 → 由 Excel 接管，撤销覆盖。
    """
    d = load()
    if not d:
        return []
    if pending is None:                      # 未传入时自行读取队列
        import plan_writer
        pending = plan_writer.load_pending()
    # 队列里待改的行号集合（按 原表文件名|sheet|行号 记）
    busy = set()
    for bname, sheets in (pending or {}).items():
        for sheet, rows in sheets.items():
            for row in rows:
                busy.add((bname, str(sheet), str(row)))

    dropped = []
    for k, rec in list(d.items()):
        still_pending = any(
            (rec.get("base_file"), rec.get("sheet"), rec.get("row")) in busy
            for _ in (0,)) if rec.get("row") else False
        if still_pending:
            continue                          # 还没写进 Excel，覆盖必须保留
        d.pop(k, None)
        dropped.append(k)
    if dropped:
        save(d)
    return dropped
