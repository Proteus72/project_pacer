# -*- coding: utf-8 -*-
"""圆角失效诊断：在【目标电脑】上运行此脚本，定位组件窗口变尖角的原因。

用法：在目标电脑双击运行，或
    python 圆角诊断.py
"""
import ctypes
import sys
from ctypes import wintypes

u = ctypes.windll.user32
d = ctypes.windll.dwmapi
CB = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

TITLE = "项目管理工作台"


def find_windows():
    out = []

    def cb(h, l):
        n = u.GetWindowTextLengthW(h)
        if n:
            b = ctypes.create_unicode_buffer(n + 1)
            u.GetWindowTextW(h, b, n + 1)
            if TITLE in b.value:
                out.append((h, b.value))
        return True

    u.EnumWindows(CB(cb), 0)
    return out


def os_build():
    try:
        import platform
        return platform.version()
    except Exception:
        return "unknown"


def main():
    print("=" * 60)
    print(" 圆角诊断 —— 组件窗口变尖角的原因定位")
    print("=" * 60)

    # 1. 系统版本
    ver = os_build()
    major = 0
    try:
        major = int(str(ver).split(".")[2])
    except Exception:
        pass
    print("\n[1] Windows 版本号(build)：%s" % ver)
    if major < 22000:
        print("    ⚠️ 低于 22000 = Windows 10。")
        print("       Windows 10 没有 DWMWA_WINDOW_CORNER_PREFERENCE(33) 属性，")
        print("       所以**系统层面无法产生原生 R 角**——这是原因。")
    elif major < 22000:
        pass
    else:
        print("    系统支持圆角属性(DWMWA_WINDOW_CORNER_PREFERENCE)。")

    # 2. 找窗口
    wins = find_windows()
    print("\n[2] 找到 %d 个相关窗口：" % len(wins))
    if not wins:
        print("    ⚠️ 没找到组件窗口。请先双击 exe 把组件打开，")
        print("       然后重新运行本脚本。")
        return
    for h, t in wins:
        vis = u.IsWindowVisible(h)
        print("    hwnd=%d 标题=%r 可见=%s" % (h, t, vis))

    # 3. 回读 DWM 圆角属性
    print("\n[3] 回读每个窗口的圆角属性：")
    readable = False
    for h, t in wins:
        if not u.IsWindowVisible(h):
            continue
        v = ctypes.c_int(0)
        r = d.DwmGetWindowAttribute(h, 33, ctypes.byref(v), ctypes.sizeof(v))
        hr = r & 0xFFFFFFFF
        if hr == 0:
            readable = True
            val = v.value
            state = {0: "default", 1: "强制不圆", 2: "圆角(round)", 3: "小圆角"}
            print("    hwnd=%d 圆角属性=%d (%s)" % (h, val, state.get(val, "未知")))
        else:
            print("    hwnd=%d 读取失败 hr=0x%08X" % (h, hr))

    print("\n" + "=" * 60)
    if major < 22000:
        print(" 结论：系统为 Windows 10，不支持原生窗口圆角。")
        print(" 方案见 README「窗口圆角」一节的降级处理。")
    elif not readable:
        print(" 结论：系统支持圆角，但读取失败，可能是窗口刚启动尚未就绪。")
        print(" 请关掉组件重新打开，10 秒后再运行本脚本。")
    else:
        print(" 结论：系统支持圆角且属性可读。")
        print(" 若仍显示尖角，请在组件目录运行本脚本并连同输出一起反馈，")
        print(" 说明问题出在 DWM 设置环节而非系统限制。")
    print("=" * 60)


if __name__ == "__main__":
    main()