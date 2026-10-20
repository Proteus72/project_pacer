# -*- coding: utf-8 -*-
"""每日摘要生成：本周/下周任务视图 + 延期预警 + 已完成清单。
产出 digest.json（桌面小组件数据源）与 digest.md（手机推送/邮件正文）。"""
import datetime as dt

from common import task_key

RISK_MARK = {"已延期": "🔴已延期", "延期风险": "🔴延期风险"}
COMPLETED_MAX = 200  # 已完成清单渲染上限，防拖慢小组件


def week_bounds(d):
    """返回 (周一, 周日) 两个 date"""
    monday = d - dt.timedelta(days=d.weekday())
    return monday, monday + dt.timedelta(days=6)


def _label(t):
    if t["subtask"] and t["task"] and t["subtask"] != t["task"]:
        return f"{t['task']} / {t['subtask']}"
    return t["subtask"] or t["task"]


def _to_date(iso):
    return dt.date.fromisoformat(iso) if iso else None


def build(tasks, flags=None):
    flags = flags or {}
    today = dt.date.today()
    w1s, w1e = week_bounds(today)
    w2s, w2e = week_bounds(today + dt.timedelta(days=7))

    def active_this_week(t):
        """本周进行中：未完成，且计划窗口覆盖本周。
        含跨周任务（上周已启动、本周仍未结束）——归「本周」而非「延期预警」，
        风险以卡片红字标注，不重复进预警类目。"""
        if t["status"] == "done":
            return False
        ps, pe = _to_date(t["plan_start"]), _to_date(t["plan_end"])
        if not ps or ps > w1e:
            return False
        return pe is None or pe >= w1s

    def starts_in(t, ws, we):
        ps = _to_date(t["plan_start"])
        return bool(ps and ws <= ps <= we)

    this_week = sorted([t for t in tasks if active_this_week(t)],
                       key=lambda t: (t["project"], t["plan_start"]))
    # 下周只收「下周新启动」的任务；跨周任务已在本周视图，不重复收录
    next_week = sorted(
        [t for t in tasks
         if t["status"] != "done" and starts_in(t, w2s, w2e)],
        key=lambda t: (t["project"], t["plan_start"]))
    # 延期预警只收「计划窗口整体早于本周」的已延期任务（真逾期）；
    # 本周进行中（含跨周）的任务即使带风险标记也归本周，红字标注在卡片上
    alerts = sorted(
        [t for t in tasks if t.get("risk") == "已延期" and not active_this_week(t)],
        key=lambda t: (t["project"], t["plan_start"]))

    # 已完成清单：Excel 标记完成 + 手动勾选完成；再补勾选快照中已不在表内的任务
    completed = [t for t in tasks if t["status"] == "done"]
    completed.sort(key=lambda t: ((t["plan_end"] or ""), (t["plan_start"] or "")), reverse=True)
    matched = {task_key(t) for t in tasks}
    for k, rec in flags.items():
        snap = rec.get("snapshot")
        if snap and k not in matched:
            s = dict(snap)
            s["status"] = "done"
            s["status_raw"] = "已完成(勾选)"
            s["done_manual"] = True
            s["risk"] = None
            completed.append(s)
    done_manual_count = sum(1 for t in tasks if t.get("done_manual"))
    completed_total = len(completed)
    if len(completed) > COMPLETED_MAX:
        completed = completed[:COMPLETED_MAX]

    return {
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "today": today.isoformat(),
        "this_week_label": f"{w1s.month}/{w1s.day} - {w1e.month}/{w1e.day}",
        "next_week_label": f"{w2s.month}/{w2s.day} - {w2e.month}/{w2e.day}",
        "this_week": this_week,
        "next_week": next_week,
        "alerts": alerts,
        "completed": completed,
        "completed_total": completed_total,
        "done_manual_count": done_manual_count,
    }


def _row(t):
    risk = RISK_MARK.get(t.get("risk"), "")
    return (f"| {t['project']} | {_label(t)} | {t['plan_start']} "
            f"| {t['plan_end'] or '-'} | {t['owner'] or '-'} "
            f"| {t['status_raw'] or t['status']} | {risk} |")


def _table(lines, tasks):
    lines.append("| 项目 | 任务/子任务 | 计划开始 | 计划结束 | 负责人 | 状态 | 风险 |")
    lines.append("|---|---|---|---|---|---|---|")
    for t in tasks:
        lines.append(_row(t))


def render_md(view):
    p = []
    p.append("# 项目管理工作台 · 每日任务摘要")
    p.append("")
    p.append(f"生成时间：{view['generated_at']}"
             f"（已完成 {view['completed_total']} 项，其中手动勾选 {view['done_manual_count']} 项）")
    p.append("")
    p.append(f"## 一、已延期（{len(view['alerts'])} 项，红字标注）")
    p.append("")
    if view["alerts"]:
        _table(p, view["alerts"])
    else:
        p.append("无已延期项。")
    p.append("")
    p.append(f"## 二、本周任务 {view['this_week_label']}（{len(view['this_week'])} 项）")
    p.append("")
    if view["this_week"]:
        _table(p, view["this_week"])
    else:
        p.append("本周无任务。")
    p.append("")
    p.append(f"## 三、下周任务 {view['next_week_label']}（{len(view['next_week'])} 项）")
    p.append("")
    if view["next_week"]:
        _table(p, view["next_week"])
    else:
        p.append("下周无任务。")
    p.append("")
    p.append("> 口径（2026-10-15 起）：每个任务只归一个类目，不重复计算。")
    p.append("> 🔴已延期 = 计划开始日与计划结束日**都早于本周一**（窗口整体在上周及更早）且仍未完成；"
             "「开始在上周、结束在本周、仍 Ongoing」属正常进行中，不计已延期。")
    p.append("> 🔴延期风险（本周类目内浅红标注）= ①开始日与结束日都在本周却仍 To be start（该启动未启动）；"
             "②结束日在本周一/周二、今天已过周二仍 Ongoing（含逾期 ≥2 天的一般化口径）。"
             "阈值见 config.json -> risk.overdue_grace_days。")
    return "\n".join(p)
