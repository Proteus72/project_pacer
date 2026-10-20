# -*- coding: utf-8 -*-
"""计划表状态回写（新版本落盘，绝不覆盖原表）。

用户在小组件里勾选「已完成」或点击状态按钮切换 Ongoing / To be start 时，
把对应任务行的 Status 单元格写入**一个全新版本的 xlsx**：

  原表：      real_plans/AR882 Project Plan_xxx.xlsx          （只读，永不改动）
  新版本：    data/updated_plans/AR882 Project Plan_xxx__v1_20261015_104500.xlsx
  下次再改：  ...__v2_20261015_110200.xlsx（始终基于当前最新版本继续累加）

哪个版本是「最新」记录在 data/plan_versions.json；解析器（parser.parse_all）
每次先查这张表，命中就解析最新版本，从而「下次读数据从最新的 Excel 读」。

【为什么不用 openpyxl 读+另存】
真实计划表里 Task ID 等列是公式（如 ="AR-"&"0"&ROW()-22），Excel 在文件里
保存了公式的**缓存值**；openpyxl 另存时不会写缓存值，新版本再被解析就会得到
空 ID，整张表直接崩掉。所以这里改成"外科手术式"改写：
  复制原 zip 的所有条目 → 只替换目标 sheet XML 里那一个 <c r="J99"> 单元格 →
  其余字节原样保留（公式、缓存值、条件格式、图表、图片全部无损）。
目标单元格写成 inlineStr（内联字符串），Excel / WPS / openpyxl 都能读。
"""
import datetime as dt
import json
import os
import re
import zipfile
from xml.sax.saxutils import escape, unescape

from common import resolve

INDEX_REL = "data/plan_versions.json"
PENDING_REL = "data/pending_writes.json"
DEFAULT_DIR = "data/updated_plans"


def index_path():
    return resolve(INDEX_REL)


def pending_path():
    return resolve(PENDING_REL)


def versions_dir(cfg):
    return resolve(cfg.get("updated_dir", DEFAULT_DIR))


def load_index():
    p = index_path()
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_index(idx):
    p = index_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=1)


def base_name(path):
    """原表文件名（不含路径）——作为版本链的稳定主键"""
    return os.path.basename(path)


def latest_path(base_path, cfg=None):
    """返回该计划表当前应读取的文件：(路径, 版本号)。没回写过就返回原表 (0)。"""
    idx = load_index()
    rec = idx.get(base_name(base_path))
    if rec and os.path.exists(rec.get("path", "")):
        return rec["path"], int(rec.get("version", 1))
    return base_path, 0


def status_label(cfg, internal, fallback=None):
    """把内部状态 done/doing/todo 转成写回 Excel 的文案。
    优先用 config.status_write_labels（用户可改），其次用表内真实出现过的原文。"""
    wl = cfg.get("status_write_labels") or {}
    if wl.get(internal):
        return wl[internal]
    smap = cfg.get("status_map", {}) or {}
    for raw, val in smap.items():
        if val == internal:
            return raw
    return fallback or {"doing": "Ongoing", "todo": "To be start",
                        "done": "Completed"}.get(internal, internal)


# ---------------------------------------------------------------- 底层改写


def col_letter(idx):
    """1 基列号 -> Excel 列字母（1->A, 27->AA）"""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _sheet_part(zf, sheet_name):
    """由工作表名定位它在 zip 内的 XML 路径"""
    wbxml = zf.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r"<sheet[^>]*name=\"" + re.escape(sheet_name) + r"\"[^>]*>", wbxml)
    if not m:
        names = re.findall(r'<sheet[^>]*name="([^"]+)"', wbxml)
        raise KeyError(f"未找到工作表 '{sheet_name}'，可用：{names}")
    rid = re.search(r'r:id="([^"]+)"', m.group(0))
    if not rid:
        raise KeyError(f"工作表 '{sheet_name}' 缺少 r:id")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rm = re.search(r'<Relationship[^>]*Id="' + re.escape(rid.group(1)) + r'"[^>]*>', rels)
    if not rm:
        raise KeyError(f"未找到关系 {rid.group(1)}")
    tgt = re.search(r'Target="([^"]+)"', rm.group(0)).group(1)
    tgt = tgt.lstrip("/")
    if not tgt.startswith("xl/"):
        tgt = "xl/" + tgt
    return tgt


def _make_cell(ref, text, style_attr=""):
    return (f'<c r="{ref}"{style_attr} t="inlineStr">'
            f"<is><t>{escape(str(text))}</t></is></c>")


def _shared_strings(zf):
    """读出共享字符串表，用于把旧值还原成可读文本（变更日志用）"""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    xml = zf.read("xl/sharedStrings.xml").decode("utf-8")
    out = []
    for si in re.findall(r"<si>(.*?)</si>", xml, re.S):
        out.append(unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
    return out


def _replace_cell(xml, ref, text, sst=None):
    """把 sheet XML 里 ref 单元格的内容换成 text；返回 (新xml, 旧值文本)"""
    sst = sst or []
    m = re.search(r'<c r="' + re.escape(ref) + r'"([^>]*?)(/>|>)', xml)
    if m:
        attrs = m.group(1)
        st = re.search(r'\bs="(\d+)"', attrs)
        style = f' s="{st.group(1)}"' if st else ""
        start = m.start()
        if m.group(2) == "/>":
            end = m.end()
            old = ""
        else:
            close = xml.find("</c>", m.end())
            if close < 0:
                raise ValueError(f"单元格 {ref} 未闭合")
            end = close + len("</c>")
            inner = xml[m.end():close]
            old = ""
            v = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
            if v:
                old = unescape(v.group(1))
            elif re.search(r'\bt="s"', attrs):       # 共享字符串：按索引还原
                iv = re.search(r"<v>(\d+)</v>", inner)
                if iv and int(iv.group(1)) < len(sst):
                    old = sst[int(iv.group(1))]
        return xml[:start] + _make_cell(ref, text, style) + xml[end:], old

    # 单元格不存在（稀疏行）：插进对应 row 里
    row = re.search(r'[^0-9]', ref)
    row_num = ref[row.start() + 1:] if row else ref.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    rm = re.search(r'<row[^>]*\br="' + re.escape(row_num) + r'"[^>]*/>', xml)
    if rm:
        new_row = rm.group(0)[:-2] + ">" + _make_cell(ref, text) + "</row>"
        return xml[:rm.start()] + new_row + xml[rm.end():], ""
    rm = re.search(r'(<row[^>]*\br="' + re.escape(row_num) + r'"[^>]*>)(.*?)</row>', xml, re.S)
    if rm:
        return (xml[:rm.end(2)] + _make_cell(ref, text) + xml[rm.end(2):], "")
    raise KeyError(f"单元格 {ref} 及其所在行均不存在")


def write_status(base_path, sheet_name, edits, cfg):
    """把 edits=[(行号, 列号1基, 新值), ...] 写入当前最新版本，另存为新版本文件。

    其余字节原样复制，公式与缓存值不丢。返回 dict：{path, version, base, from, changed}
    """
    src, ver = latest_path(base_path, cfg)
    dst_dir = versions_dir(cfg)
    os.makedirs(dst_dir, exist_ok=True)
    stem = os.path.splitext(base_name(base_path))[0]
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(dst_dir, f"{stem}__v{ver + 1}_{stamp}.xlsx")

    with zipfile.ZipFile(src, "r") as zin:
        part = _sheet_part(zin, sheet_name)
        xml = zin.read(part).decode("utf-8")
        sst = _shared_strings(zin)
        changed = []
        for row, col, val in edits:
            ref = f"{col_letter(col)}{row}"
            xml, old = _replace_cell(xml, ref, val, sst)
            old_v = old
            if old_v != str(val):
                changed.append({"cell": ref, "from": old_v, "to": str(val)})
        names = zin.namelist()
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                data = xml.encode("utf-8") if n == part else zin.read(n)
                zout.writestr(n, data)

    idx = load_index()
    idx[base_name(base_path)] = {
        "version": ver + 1,
        "path": dst,
        "updated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "from": src,
        "sheet": sheet_name,
        "changes": changed,
    }
    save_index(idx)
    return {"path": dst, "version": ver + 1, "base": base_name(base_path),
            "from": src, "changed": changed}


def load_pending():
    """读取待回写队列：{原表文件名: {sheet: {行号: {'col','value','note'}}}}"""
    p = pending_path()
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_pending(q):
    p = pending_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(q, f, ensure_ascii=False, indent=1)


def queue_write(base_path, sheet_name, row, col, value, note=""):
    """把一次状态改动放进待回写队列——**不立即生成新版本文件**。

    同一单元格重复改只保留最后一次（后来的覆盖先前的），避免版本文件堆积。
    真正的落盘由 flush_pending() 在退出 / 每 3 小时 / 手动触发时一次性完成。
    """
    q = load_pending()
    sheet = str(sheet_name)
    entry = q.setdefault(base_name(base_path), {}).setdefault(sheet, {})
    entry[str(row)] = {"col": int(col), "value": str(value),
                       "note": note, "at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    save_pending(q)
    return q


def pending_count():
    return sum(len(s) for b in load_pending().values() for s in b.values())


def flush_pending(cfg=None):
    """把队列里所有改动一次性写入各自的新版本文件（每张表只生成 **1 个** 版本）。

    返回 (结果列表, 剩余队列)。队列为空时什么都不做，不产生任何文件。
    """
    q = load_pending()
    cfg = cfg or {}
    results = []
    if not q:
        return results, {}
    for bname, sheets in q.items():
        base_path = None
        from common import ROOT  # 由调用方保证 plans_dirs 配置正确
        for d in (cfg.get("plans_dirs") or []):
            cand = resolve(os.path.join(d, bname))
            if os.path.exists(cand):
                base_path = cand
                break
        if base_path is None:
            results.append({"base": bname, "error": "未找到原表（可能在 plans_dirs 之外）"})
            continue
        for sheet, rows in sheets.items():
            edits = [(int(r), v["col"], v["value"]) for r, v in sorted(
                rows.items(), key=lambda x: int(x[0]))]
            try:
                res = write_status(base_path, sheet, edits, cfg)
                results.append({"base": bname, "sheet": sheet,
                                "path": res["path"], "version": res["version"],
                                "cells": len(edits), "changed": res["changed"]})
            except Exception as e:
                results.append({"base": bname, "sheet": sheet, "error": str(e)})
    save_pending({})            # 无论成败都清空：避免坏条目反复阻塞后续落盘
    return results, load_pending()


def reset(base_name_, cfg=None):
    """丢弃某张表的回写版本链（下次重新从原表读）。不删除已生成的版本文件。"""
    idx = load_index()
    rec = idx.pop(base_name_, None)
    save_index(idx)
    return rec
