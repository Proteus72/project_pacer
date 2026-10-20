"""最大化/还原往返测试 —— 自动适应起始状态"""
import ctypes, time, pyautogui
from ctypes import wintypes

u = ctypes.windll.user32
CB = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
class WPL(ctypes.Structure):
    _fields_ = [("length", wintypes.UINT), ("flags", wintypes.UINT),
                ("showCmd", wintypes.UINT), ("ptMinPosition", POINT),
                ("ptMaxPosition", POINT), ("rcNormalPosition", _RECT)]

def find():
    out = []
    def cb(h, l):
        n = u.GetWindowTextLengthW(h)
        if n:
            b = ctypes.create_unicode_buffer(n + 1)
            u.GetWindowTextW(h, b, n + 1)
            if b.value == '项目管理工作台' and u.IsWindowVisible(h):
                out.append(h)
        return True
    u.EnumWindows(CB(cb), 0)
    return out

def showcmd(h):
    w = WPL(); w.length = ctypes.sizeof(WPL)
    return int(w.showCmd) if u.GetWindowPlacement(wintypes.HWND(h), ctypes.byref(w)) else -1

def rect(h):
    r = _RECT(); u.GetWindowRect(wintypes.HWND(h), ctypes.byref(r))
    return (r.left, r.top, r.right - r.left, r.bottom - r.top)

h = find()[0]
# 确保从最小化回到可见
if showcmd(h) == 2:
    u.ShowWindow(wintypes.HWND(h), 9); time.sleep(1.0)

def btn():
    L, T, W, _H = rect(h)
    return (L + W - 65, T + 17)      # 最大化按钮（关闭按钮 left-27，gap8+30 -> -65）

sw, sh = u.GetSystemMetrics(0), u.GetSystemMetrics(1)
start_max = (showcmd(h) == 3)
print("screen=%dx%d 起始 showCmd=%s rect=%s" % (sw, sh, showcmd(h), rect(h)))

results = []
for i in range(4):
    expect = 1 if (start_max and i % 2 == 0) or ((not start_max) and i % 2 == 1) else 3
    b = showcmd(h)
    x, y = btn()
    pyautogui.click(x, y)
    time.sleep(1.6)
    a = showcmd(h); rc = rect(h)
    ok = (a == expect)
    # 辅助校验：最大化时窗口应铺满整屏
    fills = (rc[2] >= sw - 2 and rc[3] >= sh - 2)
    if expect == 3 and ok and not fills:
        ok = False; note = "(尺寸未铺满)"
    else:
        note = ""
    results.append(ok)
    print("[%d] %s -> showCmd=%s rect=%s expect=%s %s%s"
          % (i, "max" if b == 3 else "norm", a, rc,
             "max" if expect == 3 else "norm", "OK" if ok else "*** FAIL ***", note))

print("\n==== 往返测试：%d/%d 通过 ====" % (sum(results), len(results)))
