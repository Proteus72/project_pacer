# -*- coding: utf-8 -*-
"""SQLite 中心库：任务快照 + 运行记录。
数据源（Excel）是唯一事实，每次刷新全量重载 tasks 表。"""
import datetime as dt
import os
import sqlite3

from common import resolve

SCHEMA = """
CREATE TABLE IF NOT EXISTS tasks(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project TEXT NOT NULL,
    task TEXT, subtask TEXT, owner TEXT,
    plan_start TEXT, plan_end TEXT,
    status TEXT, status_raw TEXT, risk TEXT,
    predecessors TEXT,
    source_file TEXT, source_row INTEGER,
    loaded_at TEXT
);
CREATE TABLE IF NOT EXISTS runs(
    run_at TEXT, files INTEGER, tasks INTEGER,
    delayed INTEGER, risk INTEGER
);
"""

INSERT = """INSERT INTO tasks(project,task,subtask,owner,plan_start,plan_end,
status,status_raw,risk,predecessors,source_file,source_row,loaded_at)
VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)"""


def connect(cfg):
    path = resolve(cfg["db_path"])
    os.makedirs(os.path.dirname(path), exist_ok=True)
    conn = sqlite3.connect(path)
    conn.executescript(SCHEMA)
    return conn


def refresh_tasks(conn, tasks):
    """全量重载任务表，返回删除的旧记录数"""
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur = conn.execute("DELETE FROM tasks")
    deleted = cur.rowcount
    conn.executemany(INSERT, [
        (t["project"], t["task"], t["subtask"], t["owner"],
         t["plan_start"], t["plan_end"], t["status"], t["status_raw"],
         t["risk"], t.get("predecessors"), t["source_file"], t["source_row"], now)
        for t in tasks
    ])
    conn.commit()
    return deleted


def log_run(conn, files, n_tasks, delayed, risk):
    conn.execute("INSERT INTO runs VALUES(?,?,?,?,?)",
                 (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                  files, n_tasks, delayed, risk))
    conn.commit()
