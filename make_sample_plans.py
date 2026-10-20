# -*- coding: utf-8 -*-
"""生成示例计划表（日期相对当前周动态计算），兼作新计划表的结构模板。
用法：python make_sample_plans.py [输出目录]（默认 sample_plans/）"""
import datetime as dt
import os
import sys

from openpyxl import Workbook

HEADERS = ["项目名称", "任务名称", "子任务", "负责人", "计划开始", "计划结束", "状态", "前置任务"]


def monday_of_week():
    t = dt.date.today()
    return t - dt.timedelta(days=t.weekday())


def sample_rows(m):
    A = "项目A-原料药工艺验证"
    B = "项目B-仿制药开发"
    # (项目, 任务, 子任务, 负责人, 开始偏移, 结束偏移, 状态, 前置任务)  偏移相对本周一
    return [
        (A, "工艺验证", "批记录审核", "张三", 0, 2, "进行中", "-"),
        (A, "工艺验证", "设备清洁确认", "李四", 3, 4, "未开始", "-"),
        (A, "工艺验证", "偏差调查关闭", "王五", -3, -1, "未开始", "-"),
        (A, "工艺验证", "方法转移报告", "李四", -6, -1, "已完成", "-"),
        (A, "项目管理", "阶段评审会", "Sam", -2, -1, "进行中", "-"),
        (A, "注册申报", "撰写CTD模块三", "赵六", 3, 12, "未开始", "-"),
        (A, "注册申报", "稳定性数据汇总", "赵六", 8, 20, "未开始", "撰写CTD模块三"),
        (A, "注册申报", "eCTD格式转换", "赵六", 9, 14, "未开始", "稳定性数据汇总"),
        (B, "处方开发", "参比制剂逆向分析", "钱七", 2, 9, "进行中", "-"),
        (B, "处方开发", "初步处方筛选", "钱七", 7, 13, "未开始", "参比制剂逆向分析"),
        (B, "处方开发", "原辅料采购", "孙八", -4, 1, "已完成", "-"),
        (B, "分析方法", "含量方法开发", "周九", 10, 20, "未开始", "-"),
        (B, "注册策略", "预BE会议资料准备", "Sam", 4, 8, "未开始", "-"),
    ]


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "sample_plans")
    os.makedirs(out_dir, exist_ok=True)
    m = monday_of_week()
    projects = {}
    for r in sample_rows(m):
        projects.setdefault(r[0], []).append(r)
    for proj, rows in projects.items():
        wb = Workbook()
        ws = wb.active
        ws.title = "计划"
        ws.append(HEADERS)
        for (p, task, sub, owner, s, e, status, pre) in rows:
            ws.append([p, task, sub, owner, m + dt.timedelta(days=s),
                       m + dt.timedelta(days=e), status, pre])
        path = os.path.join(out_dir, f"{proj}_计划表.xlsx")
        wb.save(path)
        print(f"[生成] {path}（{len(rows)} 行）")


if __name__ == "__main__":
    main()
