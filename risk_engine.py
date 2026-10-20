# -*- coding: utf-8 -*-
"""延期风险规则引擎（2026-10-15 用户定口径）

判定口径（"本周" = 今天所在自然周的周一~周日）：

1. 已延期（强红，单独类目）
   - 计划开始日与计划结束日**都早于本周一**（计划窗口整体在上周及更早），且仍未完成。
   - 「开始在上周、结束在本周、状态 Ongoing」属于**正常进行中**，不算已延期。

2. 延期风险（本周类目内浅红标注）
   a. 计划开始日与计划结束日都在本周，状态仍是 To be start（该启动却没启动）；
      或计划开始日已经过去却仍未启动。
   b. 计划结束日在本周前半周（周一/周二），而今天已经到周三及以后，状态仍 Ongoing
      ——即"本周初就该结束、周中已经过了还没完"。
      另含宽限口径：结束日已过且逾期天数 >= grace+1 天时同样标记。

3. 可选的临期提醒：距计划结束日不足 N 天（config risk.deadline_warn_days，0=关闭）。

完成(done)状态一律不计风险。
"""
import datetime as dt


def _d(iso):
    return dt.date.fromisoformat(iso) if iso else None


def week_bounds(d):
    """返回 (本周一, 本周日)"""
    monday = d - dt.timedelta(days=d.weekday())
    return monday, monday + dt.timedelta(days=6)


def assess(task, risk_cfg, today=None):
    """对单条任务评估风险，返回 None / '已延期' / '延期风险'"""
    if task["status"] == "done":
        return None
    today = today or dt.date.today()
    w_start, w_end = week_bounds(today)
    grace = int(risk_cfg.get("overdue_grace_days", 1))
    warn = int(risk_cfg.get("deadline_warn_days", 0) or 0)

    ps = _d(task.get("plan_start"))
    pe = _d(task.get("plan_end"))

    # ---- 1) 已延期：计划窗口整体早于本周，且仍未完成 ----
    if pe and pe < w_start:
        return "已延期"
    if not pe and ps and ps < w_start:
        return "已延期"

    # ---- 2a) 该启动却没启动 ----
    #   开始日与结束日都在本周却仍是 To be start；或开始日已经过去仍是 To be start
    if task["status"] == "todo" and ps and ps <= w_end:
        if ps < today or (pe and pe <= w_end):
            return "延期风险"

    # ---- 2b) 本周结束日已过、仍在进行中 ----
    if task["status"] == "doing" and pe and w_start <= pe < today:
        overdue = (today - pe).days
        # 结束日在本周一/周二且今天已过周二；或已逾期 >= grace+1 天
        if overdue >= grace + 1 or (pe.weekday() <= 1 and today.weekday() >= 2):
            return "延期风险"

    # ---- 3) 可选：临期提醒 ----
    if warn > 0 and pe and 0 <= (pe - today).days <= warn:
        return "延期风险"
    return None


def annotate(tasks, risk_cfg, today=None):
    """就地给每条任务写 risk 字段；返回 (tasks, 已延期数, 延期风险数)"""
    today = today or dt.date.today()
    n_delay, n_risk = 0, 0
    for t in tasks:
        r = assess(t, risk_cfg, today=today)
        t["risk"] = r
        if r == "已延期":
            n_delay += 1
        elif r == "延期风险":
            n_risk += 1
    return tasks, n_delay, n_risk
