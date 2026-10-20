# -*- coding: utf-8 -*-
"""桌面小组件：常驻小窗显示本周/下周任务 + 延期预警（红字标注）。
双 sheet（左下角切换）：
  - 主 sheet「进行中」：已延期 + 本周 + 下周任务（每项带勾选框与状态按钮）
  - 次 sheet「已完成」：已完成任务（含 Excel 标记完成与手动勾选）
视图规则（2026-10-15 起，按用户口径）：
  - 不重复计算：每个任务只归一个类目
  - 「已延期」= 计划开始日与结束日都早于本周一（窗口整体在上周及更早）且未完成，强红色
  - 「开始在上周、结束在本周、仍 Ongoing」= 正常进行中，归「本周」不打标
  - 「本周」内浅红标注两类延期风险：该启动却仍 To be start；
    结束日在本周一/周二而今天已过周二仍 Ongoing（含逾期≥2 天一般化口径）
  - 「下周」只收下周新启动的任务
  - 每个类目下按项目分组，默认全部展开；点项目名折叠/展开该组
交互：
  - 点 □ 标记完成 / 点 ☑ 取消：把该任务 Status 单元格写入 **Excel 新版本**
    （data/updated_plans/<原表名>__vN_时间戳.xlsx，原表不覆盖），下次读数据读最新版
  - 点状态按钮（Ongoing / To be start）在两者间来回切换，同样排入回写队列
  - 勾选/切换状态同时保留 data/done_flags.json 记录（取消勾选时用它还原原状态）
  - **回写是延迟批量的**：操作时只入队 data/pending_writes.json，到点（退出时 /
    每 flush_interval_minutes 分钟）才一次性合并写盘，每张表只生成 1 个新版本，
    避免「每点一次就多一个 Excel」。界面顶部提示条可点，立即触发落盘
- 启动时先刷新一次，之后每 10 分钟后台刷新；窗口每 60 秒重渲染
- 错误写入 data/widget.log
开机自启动由计划任务 PM_Workbench_Widget（登录触发）拉起本文件。"""
import atexit
import json
import os
import sys
import threading
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import webview

from common import load_config, resolve, task_key
import done_store
import mica
import parser as pparser
import plan_writer
import status_store
import run_daily

CFG = load_config()
DIGEST = resolve(os.path.join(CFG.get("output_dir", "output"), "digest.json"))
LOG = resolve("data/widget.log")

# 应用名：出现在两处，必须一致 ———
#   1) 原生窗口标题（任务栏悬停、Alt+Tab、窗口管理器标题）
#   2) HTML 自绘标题栏 #titlebar .tt
# 改这里一处即可，__APP_TITLE__ 占位符会自动替换 HTML 里那份。
APP_TITLE = "Project Pacer"


def log_err(msg):
    try:
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + msg.rstrip() + "\n")
    except Exception:
        pass


def refresh_once():
    try:
        run_daily.run(quiet=True)
    except Exception:
        log_err(traceback.format_exc())


def flush_writer(reason=""):
    """把待回写队列一次性落盘成新版本 Excel（每张表只生成 1 个版本）。

    落盘后再跑一次 run_daily —— 此时 Excel 已承载这些状态，覆盖层会自动
    reconcile（撤销追平了的覆盖），界面改由 Excel 说了算。"""
    try:
        res, _left = plan_writer.flush_pending(CFG)
        if res:
            for r in res:
                tag = (f"{os.path.basename(r['path'])}（v{r['version']}，"
                       f"{r['cells']} 个单元格）") if r.get("path") else f"失败：{r.get('error')}"
                log_err(f"[回写{reason}] {r['base']} -> {tag}")
            refresh_once()          # 版本链变了 + 覆盖层 reconcile，重新读一次
    except Exception:
        log_err(traceback.format_exc())


def _exit_flush():
    """进程退出（含被强杀的非硬杀场景）时的兜底落盘；atexit 保证一定被调用一次"""
    try:
        if plan_writer.pending_count():
            flush_writer("退出")
    except Exception:
        pass


def write_back_daemon():
    """每 flush_interval_minutes 分钟（默认 180）批量落盘一次"""
    minutes = int(CFG.get("flush_interval_minutes", 180) or 180)
    while True:
        time.sleep(max(60, minutes * 60) if minutes >= 1 else 60)
        flush_writer("定时")


def refresher():
    while True:
        time.sleep(600)
        refresh_once()


def _read_digest():
    try:
        with open(DIGEST, encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        return json.dumps({"error": "读取摘要失败: " + str(e)})


def _ok(digest_text, msg=""):
    """操作成功：返回 {ok, msg, digest}，前端重渲染并弹提示"""
    try:
        d = json.loads(digest_text)
    except Exception:
        d = {"error": "摘要解析失败"}
    return json.dumps({"ok": True, "msg": msg, "digest": d}, ensure_ascii=False)


def _fail(msg):
    """操作失败：只弹提示，不动当前视图"""
    return json.dumps({"ok": False, "msg": msg}, ensure_ascii=False)


_FLUSH_LOCK = threading.Lock()


def write_back(task, internal_status):
    """把任务状态改动放进待回写队列（**不立即生成新版本文件**）。

    真正写盘由后台线程每 3 小时、或小组件退出时批量执行一次
    ——每张表每次落盘只生成 1 个新版本，不会每点一次就多一个文件。
    返回 (True, 提示文本) / (False, 错误文本)。原表永不被覆盖。"""
    if not CFG.get("write_back_excel", True):
        return False, "未启用回写（config.json -> write_back_excel=false）"
    orig = task.get("orig_file") or task.get("source_file")
    row, col = task.get("source_row"), task.get("status_col")
    if not orig or not row or not col:
        return False, "该任务缺少回写定位信息（行号或 Status 列）"
    label = plan_writer.status_label(CFG, internal_status)
    with _FLUSH_LOCK:
        plan_writer.queue_write(orig, task.get("sheet") or CFG.get("sheet_name"),
                                row, col, label,
                                note=f"{task.get('project')} {task.get('task')}")
    n = plan_writer.pending_count()
    return True, f"已排队待回写（当前 {n} 处改动）"


class Api:
    def get_data(self):
        return _read_digest()

    def toggle_done(self, key, value):
        """勾选/取消勾选。界面立即切换；Excel 新版本延迟批量生成。"""
        try:
            want_done = str(value).lower() in ("true", "1")
            flags = done_store.load_flags()
            tasks, _, _ = pparser.parse_all(CFG, quiet=True)
            match = next((t for t in tasks if task_key(t) == key), None)
            if want_done:
                if match is None:
                    return _fail("未找到该任务（可能已改名或改了计划开始日期，需重新勾选）")
                if match["status"] == "done" and key not in flags:
                    run_daily.run(quiet=True)      # Excel 已标记完成，仅刷新
                    return _ok(_read_digest())
                prev_raw = match.get("status_raw") or ""
                ok, info = write_back(match, "done")
                if not ok:
                    return _fail(info)
                if key not in flags:
                    done_store.mark_done(flags, match)
                flags[key]["prev_status_raw"] = prev_raw   # 取消勾选时还原用
                done_store.save_flags(flags)
                run_daily.run(quiet=True)
                return _ok(_read_digest(), f"已标记完成 · {info}")
            rec = flags.get(key) or {}
            prev_raw = rec.get("prev_status_raw")
            if match is not None:
                internal = "doing"
                if prev_raw:
                    internal = pparser._norm_status(prev_raw, CFG.get("status_map", {}))[0]
                ok, info = write_back(match, internal)
                if not ok:
                    return _fail(info)
                # 界面立即还原为该状态（覆盖层），落盘后自动撤销
                status_store.set_override(
                    key, internal, plan_writer.status_label(CFG, internal),
                    base_file=match.get("base_file"), sheet=match.get("sheet"),
                    row=match.get("source_row"))
            done_store.unmark(flags, key)
            done_store.save_flags(flags)
            run_daily.run(quiet=True)
            msg = "已取消勾选 · 状态已还原" if match is not None else "已取消勾选"
            return _ok(_read_digest(), msg)
        except Exception:
            log_err(traceback.format_exc())
            return _fail("操作失败，详见 data/widget.log")

    def toggle_status(self, key):
        """在 Ongoing / To be start 之间来回切换，并排入回写队列。"""
        try:
            tasks, _, _ = pparser.parse_all(CFG, quiet=True)
            match = next((t for t in tasks if task_key(t) == key), None)
            if match is None:
                return _fail("未找到该任务（可能已改名或改了计划开始日期）")
            if match["status"] == "done":
                return _fail("已完成的任务不参与状态切换；如需修改请先取消勾选")
            # 取反必须基于「界面当前显示的状态」而不是 Excel 原始值：
            # Excel 是延迟写的，用原始值取反会导致连点两次切不动（每次都算出同一个结果）
            shown = status_store.load().get(key, {}).get("internal") or match["status"]
            new_internal = "todo" if shown == "doing" else "doing"
            ok, info = write_back(match, new_internal)
            if not ok:
                return _fail(info)
            label = plan_writer.status_label(CFG, new_internal)
            # 界面立即生效（覆盖层）；Excel 仍延迟批量落盘。
            # 带上回写定位，reconcile 才知道这条改动到底落盘了没有
            status_store.set_override(
                key, new_internal, label,
                base_file=match.get("base_file"), sheet=match.get("sheet"),
                row=match.get("source_row"))
            run_daily.run(quiet=True)
            return _ok(_read_digest(), f"状态已改为 {label} · 界面已更新")
        except Exception:
            log_err(traceback.format_exc())
            return _fail("切换失败，详见 data/widget.log")

    def flush_now(self):
        """手动触发：立即把排队的改动落盘成新版本 Excel。"""
        try:
            n0 = plan_writer.pending_count()
            if not n0:
                return json.dumps({"ok": True, "msg": "没有待回写的改动"},
                                  ensure_ascii=False)
            results = plan_writer.flush_pending(CFG)[0]
            refresh_once()
            done = [r for r in results if r.get("path")]
            fail = [r for r in results if not r.get("path")]
            if fail:
                return _ok(_read_digest(),
                           f"已落盘 {len(done)} 张表，失败 {len(fail)} 张：{fail[0].get('error')}")
            return _ok(_read_digest(),
                       "已生成 " + "、".join(f"v{r['version']}（{r['cells']} 处）"
                                            for r in done))
        except Exception:
            log_err(traceback.format_exc())
            return _fail("落盘失败，详见 data/widget.log")

    def minimize(self):
        """最小化到任务栏"""
        try:
            webview.windows[0].minimize()
        except Exception:
            log_err(traceback.format_exc())

    def _hwnd(self):
        """取原生窗口句柄（复用 mica 的实现，已处理 ToInt32 等细节）"""
        try:
            return mica._hwnd_of(webview.windows[0])
        except Exception:
            return None

    def _win_state(self):
        """返回窗口状态：1=normal 2=minimized 3=maximized（-1=取不到）

        ⚠️ 用 GetWindowPlacement 而不是 IsZoomed：
        本窗口是无边框+透明样式，实测 IsZoomed 读数前后矛盾（同一状态下
        一会儿 0 一会儿 1），据此判断会走错分支 → 表现为「最大化后还原不了」。
        WINDOWPLACEMENT.showCmd 是权威的窗口显示命令状态。
        """
        hwnd = self._hwnd()
        if not hwnd:
            return -1
        try:
            import ctypes
            from ctypes import wintypes

            class POINT(ctypes.Structure):
                _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

            class RECT(ctypes.Structure):
                _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                            ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

            class WINDOWPLACEMENT(ctypes.Structure):
                _fields_ = [("length", wintypes.UINT), ("flags", wintypes.UINT),
                            ("showCmd", wintypes.UINT),
                            ("ptMinPosition", POINT), ("ptMaxPosition", POINT),
                            ("rcNormalPosition", RECT)]

            wp = WINDOWPLACEMENT()
            wp.length = ctypes.sizeof(WINDOWPLACEMENT)
            if not ctypes.windll.user32.GetWindowPlacement(
                    wintypes.HWND(hwnd), ctypes.byref(wp)):
                return -1
            return int(wp.showCmd)
        except Exception:
            return -1

    def _set_maximized(self, want: bool):
        """设置最大化/还原。

        ⚠️ 必须用 maximize()/restore()，不能用 toggle_fullscreen()：
        后者是「全屏」——把 FormBorderStyle 设为 None 再铺满屏幕，
        showCmd 不会变成 maximized，按钮图标永远不切。
        maximize() 设 FormWindowState.Maximized，才是真正的最大化状态。
        """
        try:
            (webview.windows[0].maximize() if want
             else webview.windows[0].restore())
            return True
        except Exception:
            pass
        # Win32 兜底：ShowWindow(SW_MAXIMIZE / SW_RESTORE)
        hwnd = self._hwnd()
        if not hwnd:
            return False
        try:
            import ctypes
            from ctypes import wintypes
            SW_MAXIMIZE, SW_RESTORE = 3, 9
            ctypes.windll.user32.ShowWindow(wintypes.HWND(hwnd),
                                            SW_MAXIMIZE if want else SW_RESTORE)
            return True
        except Exception:
            log_err(traceback.format_exc())
            return False

    def is_maximized(self):
        """当前是否最大化（showCmd == SW_SHOWMAXIMIZED(3)）"""
        try:
            return self._win_state() == 3
        except Exception:
            return False

    def toggle_maximize(self):
        """最大化 / 还原。返回切换后的状态，供前端同步按钮图标"""
        try:
            before = (self._win_state() == 3)
            ok = self._set_maximized(not before)
            # pywebview 的 Invoke 是异步的，切换未必立刻生效；
            # 短轮询等状态真正翻转再回报，避免图标与窗口不一致
            after = before
            for _ in range(20):
                st = self._win_state()
                cur = (st == 3)
                if cur != before:
                    after = cur
                    break
                time.sleep(0.05)
            return json.dumps({"ok": ok, "maximized": after},
                              ensure_ascii=False)
        except Exception:
            log_err(traceback.format_exc())
            return json.dumps({"ok": False, "maximized": False}, ensure_ascii=False)

    def close_window(self):
        """关闭窗口（触发退出落盘）"""
        try:
            webview.windows[0].destroy()
        except Exception:
            log_err(traceback.format_exc())

    def pending_info(self):
        """待回写改动数量，用于在界面上提示"""
        try:
            return json.dumps({"count": plan_writer.pending_count()},
                              ensure_ascii=False)
        except Exception:
            return json.dumps({"count": 0})


_HTML_TPL = r"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
html,body{margin:0;height:100%;background:transparent}
/* 自绘标题栏：深灰半透明（可拖动窗口）。**浮在内容层之上**（absolute 定位），
   内容从它下方滚过，半透明透出的是任务卡片的红/绿——这是原先的观感。
   观感 = 背后内容先经 blur+brightness 处理，再与 rgba 深灰层混合：
   __BAR_BLUR__ 模糊半径、__BAR_DIM__ 压暗程度(<1 变暗)、__BAR_ALPHA__ 深灰层不透明度。 */
#titlebar{position:absolute;top:0;left:0;right:0;height:34px;z-index:30;
 display:flex;align-items:center;gap:8px;padding:0 12px 0 10px;
 background:rgba(23,24,28,__BAR_ALPHA__);-webkit-app-region:drag;user-select:none;
 border-bottom:1px solid rgba(255,255,255,.07);
 backdrop-filter:blur(__BAR_BLUR__) brightness(__BAR_DIM__);
 -webkit-backdrop-filter:blur(__BAR_BLUR__) brightness(__BAR_DIM__);
 border-top-left-radius:__RADIUS__;
 border-top-right-radius:__RADIUS__}
#titlebar .tt{font-size:12px;color:#c8cdd6;font-weight:600;flex:1;overflow:hidden;
 white-space:nowrap;text-overflow:ellipsis}
#titlebar .tbtn{width:30px;height:24px;line-height:22px;text-align:center;border-radius:4px;
 font-size:12px;color:#c8cdd6;cursor:pointer;-webkit-app-region:no-drag;user-select:none}
#titlebar .tbtn:hover{background:rgba(255,255,255,.12);color:#fff}
#titlebar .tbtn.close:hover{background:#e04b4b;color:#fff}
/* 最大化按钮：☐(U+2610) 未最大化 / 🗗(U+1F5D7) 已最大化（可还原）。
   emoji 字形渲染更满，故单独降一档字号与其余按钮对齐（行高保持 22px 不偏位） */
#titlebar .tbtn#tb-max{font-size:13px;line-height:22px}
#titlebar .tbtn#tb-max.maxed{font-size:11px}
/* body 透明（否则不透明背景垫在标题栏背后会让半透明失效），
   #scroll 铺满整个窗口高度（不是"标题栏以下"）——这样内容滚动时会从
   标题栏下方经过，标题栏透出的就是任务卡片的红/绿，而不是桌面。 */
body{margin:0;height:100%;font-family:"Microsoft YaHei",sans-serif;color:#e8e8e8;font-size:13px;
 background:transparent;position:relative;
 border-radius:__RADIUS__;overflow:hidden}
/* 滚动层：铺满整窗，标题栏浮在其上；顶部留 34px 让首条内容不被标题栏压住。
   滚动条出现在本层右边缘，从标题栏下沿开始（因为标题栏是浮层不再占布局位）。 */
#scroll{position:absolute;top:0;left:0;right:0;bottom:0;
 overflow-y:auto;overflow-x:hidden;
 padding-top:34px;box-sizing:border-box;background:#17181c}
#scroll::-webkit-scrollbar{width:8px}
#scroll::-webkit-scrollbar-thumb{background:rgba(255,255,255,.16);border-radius:4px}
#scroll::-webkit-scrollbar-thumb:hover{background:rgba(255,255,255,.28)}
#scroll::-webkit-scrollbar-track{background:transparent}
#wrap{padding:10px 10px 52px}
h2{font-size:14px;margin:14px 0 6px;color:#9ecbff;border-left:3px solid #4a7bd4;padding-left:8px}
#alerts h2{color:#ff8a7a;border-left-color:#ff5f5f}
.meta{color:#7a7f8a;font-size:11px;margin-bottom:8px}
.card{background:rgba(33,34,41,.62);border-radius:8px;padding:8px 10px;margin:4px 0}
.card .name{font-weight:600;margin:2px 0}
.card .info{color:#9aa0ab;font-size:11px}
/* 已延期类目：强红色 */
.card.delay{background:#3a1414;border:1px solid #c04040}
.card.delay .name{color:#ff4d4d;font-weight:700}
.card.delay .risk{color:#ff4d4d}
/* 本周内带风险（跨周进行中/临期）：浅红色 */
.card.softred{background:#45262c;border:1px solid #7c424a}
.card.softred .name{color:#f2b5b5;font-weight:600}
.card.softred .risk{color:#f2b5b5}
.card.donecard{background:#1b2a1e;border:1px solid #2f5d3a}
.card.donecard .name{color:#9fd3a8;text-decoration:line-through}
.row1{display:flex;align-items:center;gap:6px}
.cb{width:18px;text-align:center;cursor:pointer;color:#c8cdd6;font-size:15px;flex:none}
.cb:hover{color:#ffffff}
.cb.locked{cursor:default;color:#6b7080}
.risk{font-weight:700}
/* 状态按钮：Ongoing / To be start 可点击切换 */
.sbtn{display:inline-block;padding:1px 7px;margin-left:5px;border-radius:3px;font-size:11px;
 cursor:pointer;user-select:none;background:#2a2f39;border:1px solid #4a5160;color:#cfd6e4}
.sbtn:hover{background:#3a4250;color:#ffffff}
.sbtn.doing{color:#8fd0ff;border-color:#3d6a91}
.sbtn.todo{color:#e8c07a;border-color:#8a6a35}
#pending{display:none;background:#3a2f18;border:1px solid #7a6231;color:#e8c07a;
 padding:6px 10px;border-radius:8px;margin:0 0 8px;font-size:12px;line-height:1.5;cursor:pointer}
#pending:hover{background:#4a3d20}
#pending b{color:#ffd89a}
#toast{position:fixed;left:10px;right:10px;bottom:44px;display:none;z-index:20;
 background:#243244;border:1px solid #4a7bd4;color:#dbe6f7;padding:7px 10px;
 border-radius:6px;font-size:12px;line-height:1.4}
.empty{color:#7a7f8a;padding:6px 2px}
#banner{display:none;background:#5a1a1a;border:1px solid #c04040;color:#ffb3b3;padding:8px 10px;border-radius:8px;margin:0 0 8px;font-size:12px;line-height:1.5}
#err{color:#ff5f5f;padding:10px}
#tabs{position:fixed;left:10px;bottom:10px;display:flex;gap:6px;z-index:10}
.tab{background:rgba(33,34,41,.66);border:1px solid #33363f;color:#9aa0ab;padding:4px 14px;border-radius:14px;cursor:pointer;font-size:12px;user-select:none}
.tab:hover{color:#e8e8e8}
.tab.active{background:#2c4a7c;color:#ffffff;border-color:#4a7bd4}
/* 项目分组 */
.pgroup{margin:6px 0}
.phead{display:flex;align-items:center;gap:6px;cursor:pointer;padding:5px 8px;background:rgba(38,40,50,.55);border:1px solid #33363f;border-radius:6px;user-select:none}
.phead:hover{border-color:#4a7bd4}
.phead .arrow{color:#8fa3c8;font-size:10px;flex:none;width:12px;text-align:center}
.phead .pname{font-weight:700;color:#c8d6f0;flex:1}
.phead .pcount{color:#7a7f8a;font-size:11px}
.pbody{padding:2px 0 2px 10px;border-left:1px solid #33363f;margin:2px 0 4px 5px}
</style></head><body>__TITLEBAR__<div id="scroll"><div id="wrap">
<div class="meta" id="meta">加载中…</div>
<div id="pending"></div>
<div id="banner"></div>
<div id="view-main">
  <div id="alerts"></div>
  <div id="thisweek"></div>
  <div id="nextweek"></div>
</div>
<div id="view-done" style="display:none"></div>
</div></div>
<div id="toast"></div>
<div id="tabs">
  <div class="tab active" id="tab-main">进行中</div>
  <div class="tab" id="tab-done">已完成</div>
</div>
<script>
window._tasks = {};
window._flushMin = __FLUSH_MIN__;
var collapsed = {};   // '类目|项目' -> true 表示折叠（会话内保持，默认全展开）
function esc(s){return (s==null?"":String(s)).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}
function key(t){return (t.base_file||(t.source_file||'').split(/[\\\/]/).pop())+'|'+(t.task||'')+'|'+(t.subtask||'')+'|'+(t.plan_start||'')}
function label(t){var n=esc(t.subtask||t.task);if(t.subtask&&t.task&&t.subtask!==t.task){n=esc(t.task)+" / "+esc(t.subtask)}return n}
var _tt=null;
function toast(msg){
  if(!msg) return;
  var el=document.getElementById("toast");
  el.textContent=msg; el.style.display="block";
  if(_tt) clearTimeout(_tt);
  _tt=setTimeout(function(){el.style.display="none"},4500);
}
/* 状态扁平按钮：点击在 Ongoing / To be start 间切换（已完成项不显示） */
function statusBtn(t){
  if(t.status==='done') return '';
  var doing = t.status==='doing';
  var txt = esc(t.status_raw || (doing?'Ongoing':'To be start'));
  return '<span class="sbtn '+(doing?'doing':'todo')+'" data-k="'+esc(key(t))+'" '
    +'title="点击切换 Ongoing / To be start（并写入 Excel 新版本）">'+txt+'</span>';
}
function boxHtml(t){
  var done = t.status==='done';
  var manual = !!t.done_manual;
  var cls = 'cb' + (done && !manual ? ' locked' : '');
  var tip = done ? (manual ? '点击取消勾选' : 'Excel 中已标记完成') : '点击标记完成';
  return '<span class="'+cls+'" data-k="'+esc(key(t))+'" title="'+tip+'">'+(done?'☑':'□')+'</span>';
}
function card(t, soft){
  window._tasks[key(t)] = t;
  var done = t.status==='done';
  var cls = 'card';
  if (done) { cls += ' donecard'; }
  else if (soft && t.risk) { cls += ' softred'; }
  else if (t.risk==='已延期') { cls += ' delay'; }
  var risk = (!done && t.risk) ? (' <span class="risk">【'+esc(t.risk)+'】</span>') : '';
  return '<div class="'+cls+'">'+
    '<div class="row1">'+boxHtml(t)+'<span class="name">'+label(t)+risk+'</span>'+statusBtn(t)+'</div>'+
    '<div class="info">启动 '+esc(t.plan_start)+(t.plan_end?(' ~ '+esc(t.plan_end)):'')+
    (t.owner?(' · '+esc(t.owner)):'')+'</div></div>';
}
function groupByProject(list){
  var groups={}, order=[];
  for(var i=0;i<list.length;i++){var t=list[i];var p=t.project||'(未命名项目)';
    if(!groups[p]){groups[p]=[];order.push(p);}
    groups[p].push(t);}
  return order.map(function(p){return {project:p,tasks:groups[p]};});
}
function sectionHtml(sec, title, count, list, emptyText, soft){
  var h='<h2>'+esc(title)+'（'+count+'）</h2>';
  if(!list||!list.length){return h+'<div class="empty">'+emptyText+'</div>';}
  var html='';
  groupByProject(list).forEach(function(g){
    var k=sec+'|'+g.project, col=!!collapsed[k];
    html+='<div class="pgroup">'
      +'<div class="phead'+(col?' col':'')+'" data-g="'+esc(k)+'">'
      +'<span class="arrow">'+(col?'▶':'▼')+'</span>'
      +'<span class="pname">'+esc(g.project)+'</span>'
      +'<span class="pcount">'+g.tasks.length+' 项</span></div>'
      +'<div class="pbody"'+(col?' style="display:none"':'')+'>';
    g.tasks.forEach(function(t){html+=card(t, soft);});
    html+='</div></div>';
  });
  return h+html;
}
function render(v){
  window._tasks = {};
  if(v.error){document.getElementById("banner").style.display="none";document.getElementById("meta").textContent="";document.getElementById("view-done").innerHTML="";document.getElementById("alerts").innerHTML='<div id="err">'+esc(v.error)+'</div>';return}
  var bn=document.getElementById("banner");
  if(v.warn){bn.textContent="⚠ "+v.warn;bn.style.display="block"}else{bn.style.display="none"}
  document.getElementById("meta").textContent='更新于 '+v.generated_at+' · 左下角切换 sheet · 点 □ 标记完成 · 点状态按钮切换 Ongoing/To be start · 点项目名折叠/展开';
  document.getElementById("alerts").innerHTML=sectionHtml('alerts','已延期',v.alerts.length,v.alerts,'无已延期项',false);
  document.getElementById("thisweek").innerHTML=sectionHtml('thisweek','本周 '+v.this_week_label,v.this_week.length,v.this_week,'本周无任务',true);
  document.getElementById("nextweek").innerHTML=sectionHtml('nextweek','下周 '+v.next_week_label,v.next_week.length,v.next_week,'下周无新启动任务',false);
  var cl=v.completed||[];
  document.getElementById("view-done").innerHTML=sectionHtml('done','已完成',(v.completed_total||cl.length),cl,'暂无已完成任务',false);
}
function switchTab(name){
  document.getElementById("tab-main").classList.toggle("active", name==="main");
  document.getElementById("tab-done").classList.toggle("active", name==="done");
  document.getElementById("view-main").style.display = name==="main" ? "block" : "none";
  document.getElementById("view-done").style.display = name==="done" ? "block" : "none";
}
document.getElementById("pending").addEventListener("click", function(){
  window.pywebview.api.flush_now().then(handle);
});
/* 自绘标题栏按钮：最小化 / 最大化 / 关闭 */
function bindTitleBtn(id, fn){
  var el = document.getElementById(id);
  if (el) { el.addEventListener("click", fn); }
}
bindTitleBtn("tb-min", function(){ window.pywebview.api.minimize() });
/* 最大化 / 还原：图标随实际状态切换。
   ☐ U+2610 = 未最大化（可最大化）  🗗 U+1F5D7 = 已最大化（可还原）
   用后端查真实状态（GetWindowPlacement.showCmd）而不是在前端记点击次数
   ——这样双击标题栏还原等其它路径也能同步图标 */
function setMaxIcon(maximized){
  var el = document.getElementById("tb-max");
  if (!el) return;
  el.textContent = maximized ? "\uD83D\uDDD7" : "\u2610";
  el.title = maximized ? "还原" : "最大化";
  /* emoji 字形更满，切换时同步字号档位 */
  if (maximized) { el.classList.add("maxed"); } else { el.classList.remove("maxed"); }
}
function syncMaxIcon(){
  window.pywebview.api.is_maximized().then(function(v){
    setMaxIcon(v === true || v === "True" || v === "true" || v == 1);
  }).catch(function(){});
}
bindTitleBtn("tb-max", function(){
  window.pywebview.api.toggle_maximize().then(function(s){
    var o; try{o=JSON.parse(s)}catch(e){o={maximized:false}}
    setMaxIcon(!!o.maximized);
  }).catch(function(){ setMaxIcon(false); });
});
/* 双击标题栏也能最大化/还原，故定时同步 + 窗口获得焦点时同步 */
setInterval(syncMaxIcon, 1000);
window.addEventListener("focus", syncMaxIcon);
/* 兜底：双击标题栏切最大化/还原（万一按钮失灵也能恢复） */
(function(){
  var tb = document.getElementById("titlebar");
  if (!tb) return;
  tb.addEventListener("dblclick", function(e){
    if (e.target && e.target.classList && e.target.classList.contains("tbtn")) return;
    window.pywebview.api.toggle_maximize().then(function(s){
      var o; try{o=JSON.parse(s)}catch(err){o={maximized:false}}
      setMaxIcon(!!o.maximized);
    }).catch(function(){});
  });
})();
bindTitleBtn("tb-close", function(){ window.pywebview.api.close_window() });
document.getElementById("tab-main").addEventListener("click", function(){switchTab("main")});
document.getElementById("tab-done").addEventListener("click", function(){switchTab("done")});
/* 操作返回统一处理：{ok,msg,digest} —— 失败只弹提示，成功重渲染+提示 */
function refreshPending(){
  window.pywebview.api.pending_info().then(function(s){
    var o; try{o=JSON.parse(s)}catch(e){return}
    var el=document.getElementById("pending");
    if(o && o.count>0){el.style.display="block";
      el.innerHTML="⏳ <b>"+o.count+"</b> 处状态改动待回写 Excel（退出小组件或每 "+window._flushMin+" 分钟自动合并写入，点此立即生成新版本）";}
    else{el.style.display="none";}
  });
}
function handle(s){
  var o;
  try { o = JSON.parse(s); } catch(e){ return; }
  if (o && o.ok === false) { toast('⚠ '+o.msg); return; }
  render(o.digest || o);
  toast(o.msg);
  refreshPending();
}
function onCheck(t){
  var done = t.status==='done';
  if (done && !t.done_manual) { return; }
  window.pywebview.api.toggle_done(key(t), done?"false":"true").then(handle);
}
function onToggleStatus(t){
  window.pywebview.api.toggle_status(key(t)).then(handle);
}
document.body.addEventListener("click", function(e){
  var el = e.target;
  while (el && el !== document.body) {
    if (el.classList && el.classList.contains("phead")) {
      var k = el.getAttribute("data-g");
      collapsed[k] = !collapsed[k];
      var body = el.nextElementSibling;
      el.classList.toggle("col", collapsed[k]);
      var arrow = el.querySelector(".arrow");
      if (arrow) { arrow.textContent = collapsed[k] ? "▶" : "▼"; }
      if (body && body.classList.contains("pbody")) { body.style.display = collapsed[k] ? "none" : "block"; }
      return;
    }
    if (el.classList && el.classList.contains("cb")) {
      var t = window._tasks[el.getAttribute("data-k")];
      if (t) { onCheck(t); }
      return;
    }
    if (el.classList && el.classList.contains("sbtn")) {
      var t2 = window._tasks[el.getAttribute("data-k")];
      if (t2) { onToggleStatus(t2); }
      return;
    }
    el = el.parentNode;
  }
});
function load(){
  window.pywebview.api.get_data().then(function(s){try{render(JSON.parse(s))}catch(e){document.getElementById("meta").textContent="渲染失败: "+e}});
}
window.addEventListener('pywebviewready',function(){load();refreshPending();setInterval(load,60000)});
</script></body></html>"""
# 标题栏：自绘，提供最小化 / 最大化 / 关闭（easy_drag 已开启拖动）
_TITLEBAR = """
<div id="titlebar"><span class="tt">__APP_TITLE__</span>
  <span class="tbtn" id="tb-min">─</span>
  <span class="tbtn" id="tb-max">☐</span>
  <span class="tbtn close" id="tb-close">✕</span></div>
"""
_TITLEBAR = _TITLEBAR.replace("__APP_TITLE__", APP_TITLE, 1)
_HTML_TPL = _HTML_TPL.replace("__TITLEBAR__", _TITLEBAR, 1) if "__TITLEBAR__" in _HTML_TPL else _HTML_TPL

HTML = (_HTML_TPL
        .replace("__FLUSH_MIN__", str(int(CFG.get("flush_interval_minutes", 180) or 180)))
        .replace("__BODY_ALPHA__", str(CFG.get("body_alpha", 0.66)))
        .replace("__BAR_ALPHA__", str(CFG.get("titlebar_alpha", 0.55)))
        .replace("__BAR_BLUR__", str(CFG.get("titlebar_blur_px", 12)) + "px")
        .replace("__BAR_DIM__", str(CFG.get("titlebar_dim", 0.35)))
        .replace("__RADIUS__", str(CFG.get("window_radius_px", 8)) + "px"))


def apply_glass(window):
    """按 config 设置窗口外观：backdrop=none 时只设深色标题栏（透明模式用）。

    必须等窗口句柄就绪后再调 DWM —— 这里用 events.shown 事件触发，
    比 sleep 猜时间可靠。返回重申函数（窗口激活后 DWM 属性会被 pywebview
    重置，需要定期重申）。
    """
    try:
        style = str(CFG.get("window_backdrop", "mica") or "none").lower()
        if style == "none":
            # 透明窗口模式：不铺磨砂，设原生圆角 + 标题栏压成深色（对抗强调色蓝）
            def reassert():
                try:
                    hwnd = mica.wait_hwnd(window, timeout=2)
                    if hwnd:
                        mica._dwm_set(hwnd, mica.DWMWA_USE_IMMERSIVE_DARK_MODE, 1)
                        mica.set_round_corners(hwnd, 2)
                        mica.set_caption_color(hwnd, CFG.get("caption_color", "#17181c"))
                except Exception:
                    pass
            hwnd = mica.wait_hwnd(window, timeout=5)
            if hwnd:
                mica._dwm_set(hwnd, mica.DWMWA_USE_IMMERSIVE_DARK_MODE, 1)
                r_ok = mica.set_round_corners(hwnd, 2)
                c_ok = mica.set_caption_color(hwnd, CFG.get("caption_color", "#17181c"))
                log_err(f"[外观] 深色标题栏{' + 原生圆角' if r_ok else ''} -> "
                        f"{'成功' if c_ok else '部分失败'}")
            return reassert
        ok, detail, reassert = mica.apply_all(
            window, backdrop=style, dark=True,
            extend=bool(CFG.get("extend_frame", True)),
            caption=CFG.get("caption_color", "#17181c"))
        log_err(f"[磨砂] backdrop={style} -> {'成功' if ok else '未生效'}: {detail}")
        return reassert
    except Exception:
        log_err(traceback.format_exc())
        return lambda: None


def caption_guard(reassert):
    """每 5 秒重申一次 DWM 外观属性——窗口每次激活时 pywebview 会重置它们，
    表现为「切出去再切回来，窗口变回系统蓝」。后台线程兜住。"""
    def loop():
        while True:
            time.sleep(5)
            reassert()
    threading.Thread(target=loop, daemon=True).start()


def main():
    refresh_once()                       # 启动即刷新，保证数据新鲜
    log_err("[info] widget started, data=" + DIGEST)
    # 上次被强杀（关机/任务结束）遗留的队列：启动即落盘，保证改动不丢
    try:
        if plan_writer.pending_count():
            flush_writer("启动补写")
    except Exception:
        log_err(traceback.format_exc())
    atexit.register(_exit_flush)
    threading.Thread(target=refresher, daemon=True).start()
    threading.Thread(target=write_back_daemon, daemon=True).start()
    try:
        want_transparent = bool(CFG.get("transparent_window", False))
        # 关掉磨砂后用「无边框 + 自绘标题栏」：标题栏深灰半透明，内容区保持不透
        win = webview.create_window(
            APP_TITLE, html=HTML, js_api=Api(),
            width=430, height=820, min_size=(380, 500),
            background_color="#000000",      # 只能是 6/3 位 HEX（pywebview 正则限制）
            frameless=want_transparent,      # 去掉系统标题栏（那层是系统主题蓝）
            easy_drag=want_transparent,      # pywebview 自带的无边框拖动
            transparent=want_transparent,    # WebView2 背景透明 → 网页 alpha 透到桌面
        )
        # 窗口真正显示后再套磨砂（若启用）：此刻 native 句柄一定已就绪。
        # reassert 交给后台线程每 5 秒重申（对抗 pywebview 激活时重置 DWM 属性）
        _re = []
        def _on_shown():
            _re.append(apply_glass(win))
            caption_guard(_re[0])
        win.events.shown += _on_shown
        webview.start(debug=False)
    except Exception:
        log_err(traceback.format_exc())
        raise
    finally:
        _exit_flush()
        log_err("[info] widget exited")


if __name__ == "__main__":
    main()
