# -*- coding: utf-8 -*-
"""每日主入口：解析计划表 -> 入库 -> 风险标注 -> 生成本周/下周摘要。
产出 output/digest.json（小组件数据源）与 output/digest.md（推送正文）。
用法：python run_daily.py   （显式输出，非静默）"""
import json
import os
import sys
import threading

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from common import load_config, resolve
import parser as pparser
import done_store
import risk_engine
import status_store
import plan_writer
import db as dbm
import digest


_RUN_LOCK = threading.Lock()   # 小组件后台刷新线程与 UI 操作并发调用时串行化


def run(quiet=False):
    def out(msg):
        if not quiet:
            print(msg)

    with _RUN_LOCK:
        return _run(quiet, out)


def _run(quiet, out):
    cfg = load_config()
    tasks, files, skipped = pparser.parse_all(cfg, quiet=quiet)
    if not tasks:
        out("[警告] 未解析到任何任务，请检查 config.json 的 plans_dirs 与列映射")

    flags = done_store.load_flags()
    n_flagged = done_store.apply_flags(tasks, flags)
    if n_flagged and not quiet:
        out(f"[勾选] {n_flagged} 项手动勾选已完成，本期按完成处理")

    # 界面状态覆盖层：勾选/切换状态后界面立即生效，Excel 新版本延迟批量生成。
    # 必须在风险标注之前应用，否则风险还是按旧状态算。
    # 每条任务是独立的：覆盖按 task_key 逐条存，互不影响。
    n_override = status_store.apply(tasks)
    if n_override and not quiet:
        out(f"[覆盖] {n_override} 项按用户最新操作显示（Excel 待批量回写）")
    # 只撤销「已真正写进 Excel」的覆盖；还在队列里的必须留着，
    # 否则会出现「切换任务 B 时任务 A 的状态被弹回去」的串扰
    status_store.reconcile(tasks, pending=plan_writer.load_pending())

    tasks, n_delay, n_risk = risk_engine.annotate(tasks, cfg.get("risk", {}))

    conn = dbm.connect(cfg)
    try:
        dbm.refresh_tasks(conn, tasks)
        dbm.log_run(conn, files, len(tasks), n_delay, n_risk)
    finally:
        conn.close()

    view = digest.build(tasks, flags)
    if not tasks:
        # 显式告警：避免「解析失败」和「真的没任务」在界面上长得一样
        view["warn"] = ("未解析到任何任务：请检查 config.json 的 plans_dirs 与列映射；"
                        "若刚修改过 pm_workbench 下的 .py 代码，需重新打包 exe 才会生效")
    out_dir = resolve(cfg.get("output_dir", "output"))
    os.makedirs(out_dir, exist_ok=True)
    jpath = os.path.join(out_dir, "digest.json")
    mpath = os.path.join(out_dir, "digest.md")
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(view, f, ensure_ascii=False, indent=1)
    with open(mpath, "w", encoding="utf-8") as f:
        f.write(digest.render_md(view))

    out(f"[完成] 文件 {files} 个 / 任务 {len(tasks)} 条 / 跳过 {skipped} 行")
    out(f"[完成] 本周 {len(view['this_week'])} 项, 下周 {len(view['next_week'])} 项, "
        f"已延期 {n_delay} 项, 延期风险 {n_risk} 项")
    for t in view["alerts"]:
        out(f"[预警] {t['risk']} | {t['project']} | {t['task']}"
            f"{(' / ' + t['subtask']) if t['subtask'] and t['subtask'] != t['task'] else ''}"
            f" | 计划 {t['plan_start']}~{t['plan_end'] or '?'} "
            f"| {t['status_raw'] or t['status']}")
    out(f"[输出] {mpath}")
    return view


if __name__ == "__main__":
    run()
