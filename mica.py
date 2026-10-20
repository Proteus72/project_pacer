# -*- coding: utf-8 -*-
"""Windows 11 磨砂玻璃背景（Mica / Acrylic）。

两条腿走路，叠加使用效果最好：
  1) DWM 系统背景材质：DwmSetWindowAttribute(DWMWA_SYSTEMBACKDROP_TYPE / 38)
     —— 由系统在窗口“后面”铺一层带壁纸色调的磨砂层。
  2) CSS backdrop-filter：WebView2 是 Chromium，网页可以做毛玻璃模糊
     （见 widget.pyw 的 .glass 样式），负责把背景真正“磨”开。

backdrop 取值（38 号属性）：
  1 = None  普通     2 = Mica  磨砂
  3 = Acrylic 亚克力  4 = Mica Alt（更强调壁纸色）

踩过的坑，都在这里规避了：
  - **句柄时机**：pywebview 是异步建窗的，start() 后 native 窗口未必就绪。
    必须轮询等 hwnd（带超时），否则静默无效（不报错但没效果）。
  - **background_color 只接受 6/3 位 HEX**（pywebview 正则 `^#(?:[0-9a-fA-F]{3}){1,2}$`），
    写 8 位带 alpha 的会直接 ValueError 崩掉。
  - **网页不透明会盖住 Mica**：body 背景必须半透明，否则看起来毫无变化。
"""
import ctypes
import sys
import time
from ctypes import wintypes

IS_WINDOWS = sys.platform.startswith("win")

DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_WINDOW_CORNER_PREFERENCE = 33  # Win11: 0=default 1=dontround 2=round 3=roundsmall
DWMWA_CAPTION_COLOR = 35             # 标题栏颜色（COLORREF 0x00BBGGRR）
DWMWA_TEXT_COLOR = 36                # 标题栏文字颜色
DWMWA_SYSTEMBACKDROP_TYPE = 38       # Win11 build 22523+
DWMWA_MICA_EFFECT = 1029             # Win11 build 22000~22523（旧接口）

BACKDROP_NONE = 1
BACKDROP_MICA = 2
BACKDROP_ACRYLIC = 3
BACKDROP_MICA_ALT = 4


class _MARGINS(ctypes.Structure):
    _fields_ = [("cxLeftWidth", ctypes.c_int), ("cxRightWidth", ctypes.c_int),
                ("cyTopHeight", ctypes.c_int), ("cyBottomHeight", ctypes.c_int)]


def _hwnd_of(window):
    """从 pywebview window 取原生句柄 hwnd（int）；取不到返回 None"""
    try:
        native = getattr(window, "native", None)
        if native is None:
            return None
        handle = getattr(native, "Handle", None)
        if handle is None:
            return None
        return int(handle.ToInt32()) if hasattr(handle, "ToInt32") else int(handle)
    except Exception:
        return None


def wait_hwnd(window, timeout=15.0, interval=0.05):
    """轮询等窗口句柄就绪（start() 之后窗口是异步创建的）"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        hwnd = _hwnd_of(window)
        if hwnd:
            return hwnd
        time.sleep(interval)
    return None


def _dwm_set(hwnd, attr, value):
    try:
        v = ctypes.c_int(int(value))
        r = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            wintypes.HWND(hwnd), wintypes.DWORD(attr),
            ctypes.byref(v), ctypes.sizeof(v))
        return r == 0
    except Exception:
        return False


def _dwm_get(hwnd, attr):
    try:
        v = ctypes.c_int(0)
        r = ctypes.windll.dwmapi.DwmGetWindowAttribute(
            wintypes.HWND(hwnd), wintypes.DWORD(attr),
            ctypes.byref(v), ctypes.sizeof(v))
        return (r == 0), v.value
    except Exception:
        return False, None


def extend_frame(hwnd):
    """边框延伸到客户区（margins=-1）→ 背景材质铺满整窗含标题栏"""
    try:
        m = _MARGINS(-1, -1, -1, -1)
        return ctypes.windll.dwmapi.DwmExtendFrameIntoClientArea(
            wintypes.HWND(hwnd), ctypes.byref(m)) == 0
    except Exception:
        return False


def get_build():
    """Windows build 号（失败返回 0）"""
    if not IS_WINDOWS:
        return 0
    try:
        import winreg
        k = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                           r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        return int(winreg.QueryValueEx(k, "CurrentBuildNumber")[0])
    except Exception:
        return 0


# ---------------- 未公开接口：SetWindowCompositionAttribute（亚克力兜底）----------------

class _ACCENT_POLICY(ctypes.Structure):
    _fields_ = [("AccentState", ctypes.c_int), ("AccentFlags", ctypes.c_int),
                ("GradientColor", ctypes.c_int), ("AnimationId", ctypes.c_int)]


class _WINCOMPATTRDATA(ctypes.Structure):
    _fields_ = [("Attribute", ctypes.c_int), ("Data", ctypes.c_void_p),
                ("SizeOfData", ctypes.c_size_t)]


def enable_acrylic_legacy(hwnd, dark=True):
    """用未公开 API 做亚克力/模糊（Win10 与部分 Win11 场景兜底）。

    注意：未公开接口，不同版本可能失效，失败即静默返回 False，不影响启动。
    """
    if not IS_WINDOWS or not hwnd:
        return False
    try:
        accent = _ACCENT_POLICY()
        accent.AccentState = 4          # ACCENT_ENABLE_ACRYLICBLURBEHIND
        accent.AccentFlags = 2
        # ABGR：A=cc(80%) 深灰
        accent.GradientColor = (0xCC << 24) | (0x1F << 16) | (0x1F << 8) | 0x22 if dark \
            else (0x99 << 24) | (0xF3 << 16) | (0xF3 << 8) | 0xF3
        data = _WINCOMPATTRDATA()
        data.Attribute = 19             # WCA_ACCENT_POLICY
        data.Data = ctypes.cast(ctypes.byref(accent), ctypes.c_void_p)
        data.SizeOfData = ctypes.sizeof(accent)
        r = ctypes.windll.user32.SetWindowCompositionAttribute(
            wintypes.HWND(hwnd), ctypes.byref(data))
        return bool(r)
    except Exception:
        return False


def hex_to_colorref(hexstr):
    """#RRGGBB -> COLORREF (0x00BBGGRR)"""
    h = str(hexstr).lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r | (g << 8) | (b << 16)


def set_round_corners(hwnd, preference=2):
    """Win11 原生窗口圆角（2=round，约 8px；3=roundsmall 约 4px）。

    对 frameless 窗口同样有效——系统会直接把窗口形状裁圆，
    与 HTML 内的 border-radius 双保险（若系统不支持则网页圆角兜底）。
    """
    return _dwm_set(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, preference)


def set_caption_color(hwnd, hexcolor="#17181c", text_color="#e8e8e8"):
    """显式设置标题栏底色/文字色——压掉系统强调色（否则标题栏是用户主题蓝）。"""
    ok1 = _dwm_set(hwnd, DWMWA_CAPTION_COLOR, hex_to_colorref(hexcolor))
    ok2 = _dwm_set(hwnd, DWMWA_TEXT_COLOR, hex_to_colorref(text_color))
    return ok1 and ok2


def apply_all(window, backdrop="mica", dark=True, extend=True,
              caption="#17181c", timeout=15.0):
    """一次性设置全部外观属性；返回 (是否成功, 说明)。

    ⚠️ pywebview 会在窗口激活/系统主题变化时重设 attr 20/38（见
    winforms.py 的 update_title_bar_theme：浅色主题会把 38 重置为 1=无背景），
    把我们的磨砂覆盖掉。所以这里返回一个 reassert 函数，调用方需要定期
    （或收到 activated 事件时）重新执行。
    """
    if not IS_WINDOWS:
        return False, "非 Windows，跳过", (lambda: None)
    hwnd = wait_hwnd(window, timeout=timeout)
    if not hwnd:
        return False, "取得窗口句柄超时", (lambda: None)

    mapping = {"none": BACKDROP_NONE, "mica": BACKDROP_MICA,
               "mica_alt": BACKDROP_MICA_ALT, "acrylic": BACKDROP_ACRYLIC}
    btype = mapping.get(str(backdrop or "mica").lower(), BACKDROP_MICA)

    def _once():
        steps = [("dark", _dwm_set(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, 1 if dark else 0)),
                 ("round", set_round_corners(hwnd, 2))]
        ok_new = _dwm_set(hwnd, DWMWA_SYSTEMBACKDROP_TYPE, btype)
        steps.append(("backdrop38", ok_new))
        if not ok_new and btype in (BACKDROP_MICA, BACKDROP_MICA_ALT):
            steps.append(("mica1029", _dwm_set(hwnd, DWMWA_MICA_EFFECT, 1)))
        if extend:
            steps.append(("extend_frame", extend_frame(hwnd)))
        steps.append(("caption", set_caption_color(hwnd, caption)))
        return steps

    steps = _once()
    gok, gv = _dwm_get(hwnd, DWMWA_SYSTEMBACKDROP_TYPE)
    steps.append(("verify", gok))
    ok = gok and gv == btype
    detail = ", ".join(f"{n}={'OK' if v else 'FAIL'}" for n, v in steps)
    if gok:
        detail += f" (回读 backdrop={gv})"

    def reassert():
        try:
            _once()
        except Exception:
            pass

    return ok, f"hwnd={hwnd} | {detail}", reassert
