# -*- coding: utf-8 -*-
"""干什么：动态壁纸（GIF + 视频）——把会动的画面贴在桌面图标底下，当壁纸用。
        GIF 交给 Qt 的 QMovie；视频（mp4/webm/mkv/mov…）交给随包的 ffmpeg.exe
        解成 rawvideo 从管道吐字节，我们只挑最新一帧贴上去，静音、循环、实时节奏。
        它只是"贴在桌面底层的一张贴纸"，从不改系统壁纸设置，出事就摘掉，桌面不会崩。
怎么用：由 收纳盒.py 调用——
        from 动态壁纸 import 壁纸管理器, 全网找GIF, 是壁纸素材
        管理器 = 壁纸管理器([素材目录, 内置目录], 状态变化=回调)
        管理器.换上(路径) / 管理器.停下() / 管理器.列素材()
        全网找GIF("关键词", 目标目录) → [(存好的路径, 来源名)]
        （百度图片 / 360 图片搜不到时自动退到 GitHub 仓库翻）
        是壁纸素材(路径) → 能不能当壁纸（GIF 或视频后缀）
注意：写日志函数由主程序注入（写日志 = ...），默认不输出。

★ 这台 Win11 24H2 上实测出来的关键结论（别再走回头路）：
  1) Progman 的扩展样式带 WS_EX_NOREDIRECTIONBITMAP(0x00200000)，
     普通子窗（Qt 也好、原生 GDI 也好）塞进去一律不合成 → 截图里啥也看不见；
  2) 0x052C 那套"唤 WorkerW"的老办法在这台机器上无效：发完消息 DefView 还在 Progman 里，
     也没有任何顶层 WorkerW 接管图标层；
  3) 把窗口当"最底层的顶层窗"会连图标一起盖住（图标画在 Progman 那棵树里）；
  4) 唯一可行：窗口必须是 Progman 的 **WS_EX_LAYERED 子窗**，再插到 SHELLDLL_DefView 之下。
     实测（可见性定位12.py）：满屏红块盖住壁纸、桌面图标和图标名字全部照旧可见。
  所以这里不用 Qt 的窗口，用原生窗口 + GDI 贴图，GIF 解码仍然交给 Qt 的 QMovie。"""
import ctypes
import hashlib
import http.cookiejar
import json
import os
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import threading
import urllib.error
import urllib.request
from ctypes import wintypes
from urllib.parse import quote, urljoin, urlsplit, urlunsplit

from PyQt5.QtCore import QObject, QTimer, Qt
from PyQt5.QtGui import QMovie, QImage

写日志 = lambda 文本: None  # 主程序会覆盖成真的日志函数

# ---------------- Win32 声明 ----------------

用户32 = ctypes.windll.user32
壳32 = ctypes.windll.shell32
图形32 = ctypes.windll.gdi32
核心32 = ctypes.windll.kernel32
核心32.GetModuleHandleW.restype = ctypes.c_void_p
核心32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
取实例 = 核心32.GetModuleHandleW

for _名, _类型 in (
    ("FindWindowW", ctypes.c_void_p), ("FindWindowExW", ctypes.c_void_p),
    ("GetWindow", ctypes.c_void_p), ("GetAncestor", ctypes.c_void_p),
    ("SetParent", ctypes.c_void_p), ("CreateWindowExW", ctypes.c_void_p),
    ("GetWindowLongW", ctypes.c_uint32), ("SetWindowLongW", ctypes.c_long),
    ("SetWindowPos", ctypes.c_bool), ("DefWindowProcW", ctypes.c_ssize_t),
    ("GetClientRect", ctypes.c_bool), ("IsWindow", ctypes.c_bool),
    ("GetSystemMetrics", ctypes.c_int), ("GetDC", ctypes.c_void_p),
    ("ReleaseDC", ctypes.c_int), ("BeginPaint", ctypes.c_void_p),
    ("InvalidateRect", ctypes.c_bool), ("UpdateWindow", ctypes.c_bool),
    ("SendMessageW", ctypes.c_long), ("DestroyWindow", ctypes.c_bool),
    ("ShowWindow", ctypes.c_bool), ("EndPaint", ctypes.c_bool),
    ("RegisterClassW", ctypes.c_ushort),
    ("SetLayeredWindowAttributes", ctypes.c_bool),
    ("GetClassNameW", ctypes.c_int),
):
    getattr(用户32, _名).restype = _类型
for _名, _类型 in (
    ("CreateCompatibleDC", ctypes.c_void_p), ("CreateDIBSection", ctypes.c_void_p),
    ("SelectObject", ctypes.c_void_p), ("DeleteDC", ctypes.c_bool),
    ("DeleteObject", ctypes.c_bool), ("StretchBlt", ctypes.c_bool),
    ("SetStretchBltMode", ctypes.c_int), ("CreateSolidBrush", ctypes.c_void_p),
    ("GetDeviceCaps", ctypes.c_int), ("CreateCompatibleBitmap", ctypes.c_void_p),
):
    getattr(图形32, _名).restype = _类型
图形32.SetStretchBltMode.argtypes = [ctypes.c_void_p, ctypes.c_int]
图形32.StretchBlt.argtypes = [ctypes.c_void_p,                        # 目标 dc
                              ctypes.c_int, ctypes.c_int,
                              ctypes.c_int, ctypes.c_int,             # 目标 x y 宽 高
                              ctypes.c_void_p,                        # 源 dc
                              ctypes.c_int, ctypes.c_int,
                              ctypes.c_int, ctypes.c_int,             # 源 x y 宽 高
                              ctypes.c_uint32]                        # SRCCOPY
图形32.CreateDIBSection.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint,
                                    ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]
图形32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
图形32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
图形32.DeleteDC.argtypes = [ctypes.c_void_p]
图形32.DeleteObject.argtypes = [ctypes.c_void_p]
图形32.CreateSolidBrush.argtypes = [ctypes.c_uint]
用户32.GetDC.argtypes = [ctypes.c_void_p]
用户32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
用户32.SetLayeredWindowAttributes.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                              ctypes.c_uint, ctypes.c_uint]
用户32.GetClientRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
用户32.BeginPaint.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
用户32.EndPaint.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
用户32.InvalidateRect.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_bool]
用户32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
用户32.IsWindow.argtypes = [ctypes.c_void_p]
用户32.DestroyWindow.argtypes = [ctypes.c_void_p]
用户32.UpdateWindow.argtypes = [ctypes.c_void_p]
用户32.RegisterClassW.argtypes = [ctypes.c_void_p]
用户32.SetWindowPos.argtypes = [ctypes.c_void_p, ctypes.c_void_p] + [ctypes.c_int] * 4 + [ctypes.c_uint]
用户32.CreateWindowExW.argtypes = [ctypes.c_ulong, wintypes.LPCWSTR, wintypes.LPCWSTR,
                                   ctypes.c_ulong, ctypes.c_int, ctypes.c_int,
                                   ctypes.c_int, ctypes.c_int, ctypes.c_void_p,
                                   ctypes.c_void_p, wintypes.HINSTANCE, ctypes.c_void_p]
用户32.DefWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                  ctypes.c_void_p, ctypes.c_void_p]
用户32.FindWindowExW.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                 wintypes.LPCWSTR, wintypes.LPCWSTR]
用户32.SetParent.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
用户32.GetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int]
用户32.SetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_long]

GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_POPUP = 0x80000000
WS_CHILD = 0x40000000
WS_VISIBLE = 0x10000000
WS_CLIPCHILDREN = 0x02000000
WS_CLIPSIBLINGS = 0x04000000
WS_EX_LAYERED = 0x00080000
WS_EX_TRANSPARENT = 0x00000020      # 鼠标穿透：桌面图标照常点、照常拖框
WS_EX_TOOLWINDOW = 0x00000080       # 不进 Alt+Tab
WS_EX_NOACTIVATE = 0x08000000       # 永远不抢焦点
LWA_ALPHA = 0x00000002
GW_CHILD = 5
GW_HWNDNEXT = 2
GA_PARENT = 1
HWND_BOTTOM = 1
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_NOSIZE = 0x0001
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
SWP_FRAMECHANGED = 0x0020
SRCCOPY = 0x00CC0020
HALFTONE = 4
WM_PAINT = 0x000F
WM_ERASEBKGND = 0x0014
WM_SIZE = 0x0005
BI_RGB = 0
DIB_RGB_COLORS = 0
SM_CXSCREEN, SM_CYSCREEN = 0, 1

# SHQueryUserNotificationState 的状态码
跑着D3D全屏 = 3
正在放映 = 4


class 矩形类(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class 位图信息头(ctypes.Structure):
    _fields_ = [("biSize", ctypes.c_uint), ("biWidth", ctypes.c_int),
                ("biHeight", ctypes.c_int), ("biPlanes", ctypes.c_ushort),
                ("biBitCount", ctypes.c_ushort), ("biCompression", ctypes.c_uint),
                ("biSizeImage", ctypes.c_uint), ("biXPelsPerMeter", ctypes.c_int),
                ("biYPelsPerMeter", ctypes.c_int),
                ("biClrUsed", ctypes.c_uint), ("biClrImportant", ctypes.c_uint)]


class 绘制结构类(ctypes.Structure):
    """PAINTSTRUCT，一个字节都不能少。
    ★ 2026-11-05 踩过的坑：这里原来写的是
        hdc/fErase/rcPaint/restore/left/top/right/bottom/hrgn/flags/reserved
      一共只有 64 字节，而 Windows 的 PAINTSTRUCT 是 72 字节
      （hdc8 + fErase4 + RECT16 + fRestore4 + fIncUpdate4 + rgbReserved 32，按 8 对齐）。
      BeginPaint 每次都往里写满 72 字节，多出的 8 字节踩在堆上，
      当时看不出来，过一会儿随便一次垃圾回收释放堆块时就 0xc0000005 崩掉
      （表现就是"一点设置⚙️就崩溃"，因为开设置框时 GC 刚好在跑）。
      所以必须把 fIncUpdate 和 rgbReserved[32] 照原样补全。"""
    _fields_ = [("hdc", ctypes.c_void_p), ("fErase", ctypes.c_int),
                ("rcPaint", 矩形类), ("restore", ctypes.c_int),
                ("fIncUpdate", ctypes.c_int),
                ("rgbReserved", ctypes.c_ubyte * 32)]


def 窗口类名(窗):
    缓冲区 = ctypes.create_unicode_buffer(256)
    用户32.GetClassNameW(窗, 缓冲区, 256)
    return 缓冲区.value


def 取桌面窗():
    """Progman 句柄（桌面本体）。"""
    return 用户32.FindWindowW("Progman", None)


def 取图标层(桌面=None):
    """桌面图标那一层 SHELLDLL_DefView 的句柄；动图要插在它下面。"""
    桌面 = 桌面 or 取桌面窗()
    if not 桌面:
        return None
    子 = 用户32.FindWindowExW(桌面, None, "SHELLDLL_DefView", None)
    return 子 or None


def 前台在全屏():
    """游戏 / 放映 / 全屏视频在最前面时返回 True，这时候动图该歇着省电。"""
    状态 = ctypes.c_int(0)
    try:
        错 = 壳32.SHQueryUserNotificationState(ctypes.byref(状态))
    except Exception:
        return False
    if 错 != 0:
        return False
    return 状态.value in (跑着D3D全屏, 正在放映)


# ---------------- 原生贴纸窗口 ----------------

层表 = {}          # 句柄 → 壁纸层实例（窗口过程要靠它找回自己）
过程们 = []        # 保住 WNDPROC 回调不被回收
已注册类 = [False]
类名 = "收纳盒动图贴纸类"


def _注册类():
    if 已注册类[0]:
        return
    wndproc类型 = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_void_p, ctypes.c_uint,
                                     ctypes.c_void_p, ctypes.c_void_p)

    def _过程(窗, 消, 参1, 参2):
        try:
            if 消 == WM_ERASEBKGND:
                return 1          # 背景由我们自己画，不让系统擦，避免闪烁
            if 消 == WM_PAINT:
                自 = 层表.get(窗)
                if 自 is not None:
                    画 = 绘制结构类()
                    图片底 = 用户32.BeginPaint(窗, ctypes.byref(画))
                    区 = 矩形类()
                    用户32.GetClientRect(窗, ctypes.byref(区))
                    自._刷到(图片底, 区)
                    用户32.EndPaint(窗, ctypes.byref(画))
                    return 0
        except Exception as e:
            写日志(f"动图贴纸绘制出错：{type(e).__name__}: {e}")
        return 用户32.DefWindowProcW(窗, 消, 参1, 参2)

    回调 = wndproc类型(_过程)
    过程们.append(回调)

    class 窗口类(ctypes.Structure):
        _fields_ = [("style", ctypes.c_uint), ("lpfnWndProc", wndproc类型),
                    ("cbClsExtra", ctypes.c_int), ("cbWndExtra", ctypes.c_int),
                    ("hInstance", ctypes.c_void_p), ("hIcon", ctypes.c_void_p),
                    ("hCursor", ctypes.c_void_p), ("hbrBackground", ctypes.c_void_p),
                    ("lpszMenuName", wintypes.LPCWSTR), ("lpszClassName", wintypes.LPCWSTR)]

    类 = 窗口类()
    类.style = 0
    类.lpfnWndProc = 回调
    类.hInstance = 取实例(None)
    类.hbrBackground = 图形32.CreateSolidBrush(0x00000000)
    类.lpszClassName = 类名
    if not 用户32.RegisterClassW(ctypes.byref(类)):
        raise ValueError("注册动图贴纸窗口类失败")
    已注册类[0] = True


class 贴纸层(QObject):
    """Progman 底下的分层子窗：建窗、建 DIB、GDI 贴图、歇/继续/换父窗都在这儿。
    具体"帧从哪来"交给子类：GIF 层用 QMovie，视频层用 ffmpeg 管道。"""

    def __init__(self):
        super().__init__()
        self.句柄 = None
        self.父窗 = None
        self.位图 = None
        self.像素地址 = 0        # CreateDIBSection 给的那块显存/内存的起始地址
        self.像素长 = 0
        self.帧宽 = 0
        self.帧高 = 0

    def _开始播放(self):
        """子类覆写：开始出帧。"""

    def _暂停播放(self):
        """子类覆写：停止出帧、释放解码器。"""

    # ---- 帧搬运 ----

    def _新帧(self):
        """QMovie 每出一帧，就把像素拷进我们自己的 DIB，再叫窗口重画一次。
        图是局部变量，函数返回前一直活着，所以拷进去是安全的（不用留引用）。"""
        if not self.句柄:
            return
        图 = self.影片.currentImage().convertToFormat(QImage.Format_ARGB32)
        if 图.isNull():
            return
        宽, 高 = 图.width(), 图.height()
        if 宽 < 1 or 高 < 1:
            return
        if 宽 != self.帧宽 or 高 != self.帧高 or not self.位图:
            try:
                self._建位图(宽, 高)
            except Exception as e:
                写日志(f"动图建位图失败：{type(e).__name__}: {e}")
                return
        try:
            ctypes.memmove(self.像素地址, int(图.constBits()),
                           min(图.sizeInBytes(), self.像素长))
        except Exception as e:
            写日志(f"动图帧拷进位图失败：{type(e).__name__}: {e}")
            return
        用户32.InvalidateRect(self.句柄, None, False)
        用户32.UpdateWindow(self.句柄)

    def _建位图(self, 宽, 高):
        信息 = 位图信息头()
        信息.biSize = ctypes.sizeof(位图信息头)
        信息.biWidth = 宽
        信息.biHeight = -高          # 负数 = 自上而下，跟 QImage 的内存顺序一致
        信息.biPlanes = 1
        信息.biBitCount = 32
        信息.biCompression = BI_RGB
        像素指针 = ctypes.c_void_p()
        内存dc = 用户32.GetDC(None)
        新图 = 图形32.CreateDIBSection(内存dc, ctypes.byref(信息), DIB_RGB_COLORS,
                                    ctypes.byref(像素指针), None, 0)
        用户32.ReleaseDC(None, 内存dc)
        if not 新图 or not 像素指针.value:
            raise ValueError("CreateDIBSection 失败，动图放不下")
        if self.位图:
            图形32.DeleteObject(self.位图)
        self.位图 = 新图
        self.像素地址 = 像素指针.value
        self.像素长 = 宽 * 高 * 4
        self.帧宽, self.帧高 = 宽, 高

    def _刷到(self, 目标dc, 区):
        if not self.位图:
            return
        内存dc = 图形32.CreateCompatibleDC(目标dc)
        旧 = 图形32.SelectObject(内存dc, self.位图)
        图形32.SetStretchBltMode(目标dc, HALFTONE)
        图形32.StretchBlt(目标dc, 区.left, 区.top, 区.right - 区.left, 区.bottom - 区.top,
                         内存dc, 0, 0, self.帧宽, self.帧高, SRCCOPY)
        图形32.SelectObject(内存dc, 旧)
        图形32.DeleteDC(内存dc)

    # ---- 上墙 / 下墙 ----

    def 贴上去(self):
        _注册类()
        父 = 取桌面窗()
        if not 父:
            raise ValueError("找不到桌面窗口 Progman")
        宽 = 用户32.GetSystemMetrics(SM_CXSCREEN)
        高 = 用户32.GetSystemMetrics(SM_CYSCREEN)
        样式 = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
        扩展 = WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
        句柄 = 用户32.CreateWindowExW(扩展, 类名, "收纳盒动图", 样式,
                                   0, 0, 宽, 高, 父, None,
                                   取实例(None), None)
        if not 句柄:
            raise ValueError("动图贴纸窗口创建失败")
        self.句柄 = 句柄
        self.父窗 = 父
        层表[句柄] = self
        # 全不透明的分层窗：内容靠 GDI 画，DWM 负责合成
        用户32.SetLayeredWindowAttributes(句柄, 0, 255, LWA_ALPHA)
        self._插到图标层底下()
        self._开始播放()
        用户32.InvalidateRect(句柄, None, True)
        用户32.UpdateWindow(句柄)

    def _插到图标层底下(self):
        """插在 SHELLDLL_DefView「之后」= 排在它下面 = 动图在图标底下。"""
        图标 = 取图标层(self.父窗)
        用户32.SetWindowPos(self.句柄, 图标 or HWND_BOTTOM, 0, 0, 0, 0,
                           SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW)

    def 铺满(self):
        if not self.句柄:
            return
        区 = 矩形类()
        if self.父窗 and 用户32.GetClientRect(self.父窗, ctypes.byref(区)):
            宽 = max(1, int(区.right - 区.left))
            高 = max(1, int(区.bottom - 区.top))
        else:
            宽 = 用户32.GetSystemMetrics(SM_CXSCREEN)
            高 = 用户32.GetSystemMetrics(SM_CYSCREEN)
        用户32.SetWindowPos(self.句柄, 0, 0, 0, 宽, 高,
                           SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW)
        用户32.InvalidateRect(self.句柄, None, True)

    def 歇着(self):
        self._暂停播放()
        if self.句柄:
            用户32.ShowWindow(self.句柄, 0)   # SW_HIDE

    def 继续(self):
        if not self.句柄:
            return
        用户32.ShowWindow(self.句柄, 5)   # SW_SHOW
        self._插到图标层底下()
        self._开始播放()
        用户32.InvalidateRect(self.句柄, None, True)

    def 换父窗(self):
        """桌面结构变了（explorer 重启）就重新认领一次，返回是否真换了。"""
        现父 = 取桌面窗()
        if not 现父 or 现父 == self.父窗:
            return False
        if self.句柄:
            层表.pop(self.句柄, None)
            用户32.DestroyWindow(self.句柄)
            self.句柄 = None
        self.父窗 = 现父
        self.贴上去()
        return True

    def 摘掉(self):
        try:
            self._暂停播放()
        except Exception:
            pass
        if self.句柄:
            层表.pop(self.句柄, None)
            用户32.DestroyWindow(self.句柄)
            self.句柄 = None
        if self.位图:
            图形32.DeleteObject(self.位图)
            self.位图 = None


# 能当壁纸的文件后缀：GIF 走 QMovie，视频走 ffmpeg
视频后缀集 = (".mp4", ".webm", ".mkv", ".mov", ".m4v", ".avi", ".wmv")


def 是壁纸素材(路径):
    小 = str(路径).lower()
    return 小.endswith(".gif") or 小.endswith(视频后缀集)


def 找ffmpeg():
    """找随包发的 ffmpeg.exe：程序目录 → _internal → 源码树 工具 文件夹 → PATH。
    打包时它放在 _internal\\ffmpeg.exe；源码运行时在 动态壁纸.py 旁边的 工具 里。"""
    候选 = []
    try:
        程序目录 = os.path.dirname(os.path.abspath(sys.executable))
        候选.append(os.path.join(程序目录, "ffmpeg.exe"))
        候选.append(os.path.join(程序目录, "_internal", "ffmpeg.exe"))
    except Exception:
        pass
    try:
        候选.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "工具", "ffmpeg.exe"))
    except NameError:
        pass
    if getattr(sys, "_MEIPASS", ""):
        候选.append(os.path.join(sys._MEIPASS, "ffmpeg.exe"))
    for 路径 in 候选:
        try:
            if os.path.isfile(路径):
                return 路径
        except Exception:
            continue
    return shutil.which("ffmpeg") or ""


def 探视频(路径, 超时=20):
    """让 ffmpeg 只解第一帧，确认这文件真能放；不能放就把它的错误原话带回来。
    有这一步，坏文件/编码不支持时会保留原来那张壁纸，不会先变黑再报错。"""
    程序 = 找ffmpeg()
    if not 程序:
        raise ValueError("没找到随包的 ffmpeg.exe，视频壁纸用不了")
    try:
        结果 = subprocess.run(
            [程序, "-hide_banner", "-v", "error", "-i", 路径,
             "-frames:v", "1", "-f", "null", "-"],
            stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
            timeout=超时, creationflags=0x08000000)
    except subprocess.TimeoutExpired:
        raise ValueError("这个视频 20 秒都没解出第一帧，可能文件坏了")
    错 = (结果.stderr or b"").decode("utf-8", "ignore").strip().splitlines()
    if 结果.returncode != 0:
        raise ValueError("这个视频放不了：" + (错[-1] if 错 else "未知错误"))
    return True


class 壁纸层(贴纸层):
    """GIF 版：QMovie 解码，画面用 GDI 贴上去。"""

    def __init__(self, gif路径):
        super().__init__()
        self.gif路径 = gif路径
        self.影片 = QMovie(gif路径)
        if not self.影片.isValid():
            raise ValueError("这文件不是会动的 GIF（帧读不出来）")
        self.影片.frameChanged.connect(lambda _序: self._新帧())  # 槽名必须 ASCII
        self.影片.setCacheMode(QMovie.CacheMode.CacheNone)

    def _开始播放(self):
        self.影片.start()

    def _暂停播放(self):
        self.影片.stop()


class 视频层(贴纸层):
    """视频版：ffmpeg 按原速（-re）循环解码，rawvideo(BGRA) 从管道吐出来；
    后台线程只留最新一帧，主线程按 30fps 拷进 DIB 贴上去。
    这样解码慢一点、界面偶尔卡一下，最多掉几帧，不会把主线程堵死。"""

    def __init__(self, 视频路径):
        super().__init__()
        self.视频路径 = 视频路径
        if not os.path.isfile(视频路径):
            raise ValueError("视频文件不在了")
        探视频(视频路径)          # 坏文件在这一步就拦住，原来那张继续放着
        self.进程 = None
        self.读线程 = None
        self.锁 = threading.Lock()
        self.最新帧 = None
        self.帧序号 = 0
        self.已画序号 = 0
        self.停 = False
        self.画心跳 = QTimer()
        self.画心跳.setInterval(30)   # 约 33fps，够用；视频再快也只看最新一张
        self.画心跳.timeout.connect(lambda: self._画最新帧())  # 槽名必须 ASCII，中文一律 lambda 包

    def _目标尺寸(self):
        """解码尺寸 = 桌面客户区，最大 1920×1080：4K 屏也按 1080p 解再拉伸铺满，
        省一半以上 CPU 和内存带宽，肉眼看不出差别。"""
        区 = 矩形类()
        if self.父窗 and 用户32.GetClientRect(self.父窗, ctypes.byref(区)):
            宽 = int(区.right - 区.left)
            高 = int(区.bottom - 区.top)
        else:
            宽 = 用户32.GetSystemMetrics(SM_CXSCREEN)
            高 = 用户32.GetSystemMetrics(SM_CYSCREEN)
        return max(320, min(宽, 1920)), max(180, min(高, 1080))

    def _开始播放(self):
        if self.进程 is not None:
            return
        宽, 高 = self._目标尺寸()
        if 宽 != self.帧宽 or 高 != self.帧高 or not self.位图:
            self._建位图(宽, 高)
        程序 = 找ffmpeg()
        if not 程序:
            raise ValueError("没找到随包的 ffmpeg.exe，视频壁纸用不了")
        尺寸 = "%d:%d" % (宽, 高)
        命令 = [程序, "-hide_banner", "-loglevel", "error",
                "-re", "-stream_loop", "-1", "-i", self.视频路径,
                "-an", "-sn", "-dn",
                "-vf", "scale=%s:force_original_aspect_ratio=increase,crop=%s" % (尺寸, 尺寸),
                "-f", "rawvideo", "-pix_fmt", "bgra", "-"]
        self.进程 = subprocess.Popen(命令, stdin=subprocess.DEVNULL,
                                     stdout=subprocess.PIPE,
                                     stderr=subprocess.DEVNULL,
                                     bufsize=宽 * 高 * 4,
                                     creationflags=0x08000000)  # 不弹黑框
        self.停 = False
        with self.锁:
            self.最新帧 = None
            self.帧序号 = 0
            self.已画序号 = 0
        self.读线程 = threading.Thread(target=self._读帧,
                                       args=(self.进程, 宽 * 高 * 4), daemon=True)
        self.读线程.start()
        self.画心跳.start()

    def _读帧(self, 进程, 帧长):
        """后台线程：只负责把管道读空、留下最新一帧；旧帧直接丢，不排队。"""
        try:
            while not self.停:
                块 = 进程.stdout.read(帧长)
                if not 块 or len(块) != 帧长:
                    break
                with self.锁:
                    self.最新帧 = 块
                    self.帧序号 += 1
        except Exception as e:
            写日志(f"视频解码线程退出：{type(e).__name__}: {e}")

    def _画最新帧(self):
        if not self.句柄 or not self.位图:
            return
        with self.锁:
            if self.最新帧 is None or self.帧序号 == self.已画序号:
                return
            块 = self.最新帧
            self.已画序号 = self.帧序号
        try:
            ctypes.memmove(self.像素地址, 块, min(len(块), self.像素长))
        except Exception as e:
            写日志(f"视频帧拷进位图失败：{type(e).__name__}: {e}")
            return
        用户32.InvalidateRect(self.句柄, None, False)
        用户32.UpdateWindow(self.句柄)

    def _暂停播放(self):
        self.停 = True
        self.画心跳.stop()
        进程 = self.进程
        self.进程 = None
        if 进程 is not None:
            try:
                进程.kill()
            except Exception:
                pass
        if self.读线程 is not None:
            self.读线程.join(timeout=1.0)
            self.读线程 = None
        if 进程 is not None:
            try:
                进程.stdout.close()
            except Exception:
                pass
        with self.锁:
            self.最新帧 = None
            self.帧序号 = 0
            self.已画序号 = 0


# ---------------- 对外管事的 ----------------

class 壁纸管理器:
    """开关、换图、全屏自动暂停、桌面被重建了重新贴回去。"""

    def __init__(self, 素材目录们, 状态变化=None, 允许全屏暂停=True):
        self.素材目录们 = list(素材目录们)
        self.状态变化 = 状态变化
        self.允许全屏暂停 = 允许全屏暂停
        self.层 = None
        self.当前 = ""
        self.暂停中 = False
        self.心跳 = QTimer()
        self.心跳.setInterval(2500)
        self.心跳.timeout.connect(self._tick)  # 槽名必须 ASCII
        self.心跳.start()

    def 列素材(self):
        """内置 + 他自己下载/存进来的 GIF 和视频，按类型筛掉明显不能用的。
        同名只留一个：首次启动会把内置那几张复制进他的素材目录，
        不去重的话设置下拉里每张都会出现两遍。素材目录在前，优先用它那份。"""
        出 = []
        见过的 = set()
        for 目录 in self.素材目录们:
            if not os.path.isdir(目录):
                continue
            for 名 in sorted(os.listdir(目录)):
                小 = 名.lower()
                if 小.endswith(".gif"):
                    下限, 上限 = 50_000, 40_000_000
                elif 小.endswith(视频后缀集):
                    下限, 上限 = 100_000, 600_000_000
                else:
                    continue
                全 = os.path.join(目录, 名)
                try:
                    大小 = os.path.getsize(全)
                except OSError:
                    continue
                if 名 not in 见过的 and 下限 <= 大小 <= 上限:
                    见过的.add(名)
                    出.append(全)
        return 出

    def 换上(self, 路径):
        """换一张动图 / 一个视频贴上。顺序很重要：先把新的贴好，再摘旧的——
        万一新文件有问题（坏文件、编码放不了），原来那张继续放着，桌面不会突然空一块。"""
        if self.层 is not None and self.当前 == 路径:
            return True  # 已经是它了，别重贴（重贴会跳回第一帧）
        try:
            类 = 视频层 if str(路径).lower().endswith(视频后缀集) else 壁纸层
            新层 = 类(路径)
            新层.贴上去()
        except Exception as e:
            写日志(f"动态壁纸没能开：{type(e).__name__}: {e}")
            self._通知("这个素材贴不上，桌面保持原样")
            return False
        旧层 = self.层
        self.层 = 新层
        self.当前 = 路径
        self.暂停中 = False
        if 旧层 is not None:
            try:
                旧层.摘掉()
            except Exception:
                pass
        写日志(f"动态壁纸已贴上：{路径}")
        self._通知("动态壁纸已开启")
        return True

    def 停下(self):
        if self.层 is not None:
            try:
                self.层.摘掉()
            except Exception:
                pass
        self.层 = None
        self.当前 = ""
        self.暂停中 = False

    def 在运行(self):
        return self.层 is not None

    def 设全屏暂停(self, 允许):
        """设置里那个勾选：关掉就立刻恢复播放，别再歇着。"""
        self.允许全屏暂停 = bool(允许)
        if not self.允许全屏暂停 and self.层 is not None and self.暂停中:
            self.层.继续()
            self.暂停中 = False

    def _tick(self):
        if self.层 is None:
            return
        try:
            if not self.层.句柄 or not 用户32.IsWindow(self.层.句柄):
                # 窗口没了（explorer 重启把子窗带走了）→ 重新贴一张
                路径 = self.当前
                self.停下()
                if 路径 and os.path.exists(路径):
                    self.换上(路径)
                return
            要暂停 = self.允许全屏暂停 and 前台在全屏()
            if 要暂停 and not self.暂停中:
                self.层.歇着()
                self.暂停中 = True
            elif not 要暂停 and self.暂停中:
                self.层.继续()
                self.暂停中 = False
            self.层.换父窗()
            self.层._插到图标层底下()   # 图标层句柄可能变，顺手校正一次
        except Exception as e:
            写日志(f"动态壁纸心跳出错，自己停掉保平安：{type(e).__name__}: {e}")
            try:
                self.停下()
            except Exception:
                pass
            self._通知("动态壁纸自己停了，桌面不受影响")

    def _通知(self, 文本):
        if self.状态变化:
            try:
                self.状态变化(文本)
            except Exception:
                pass


# ---------------- 从网上找动图（GitHub 上带 gif 的仓库） ----------------

搜索地址 = ("https://api.github.com/search/repositories"
            "?q={}&sort=stars&order=desc&per_page=8")
包地址 = "https://codeload.github.com/{}/legacy.tar.gz/refs/heads/{}"

# 他习惯打中文，但 GitHub 的仓库搜索基本只认英文，先替他翻一层
关键词词典 = {
    "极光": "aurora", "星空": "space stars", "宇宙": "space", "星云": "nebula",
    "银河": "galaxy", "黑洞": "black hole", "地球": "earth", "月亮": "moon", "太阳": "sun",
    "海": "ocean", "浪": "waves", "瀑布": "waterfall", "山": "mountain", "雪": "snow",
    "雨": "rain", "云": "clouds", "森林": "forest", "花": "flower", "城市": "city",
    "夜景": "night city", "赛博": "cyberpunk", "抽象": "abstract loop", "循环": "loop",
    "猫": "cat", "狗": "dog", "壁纸": "wallpaper", "动图": "gif", "动画": "animation",
}

推荐词 = "aurora / nebula / ocean / rain / city / abstract loop"

# 一次搜索只能替换一层词，按长词优先匹配，免得「海浪」被拆成 ocean+waves 粘一起
词典正则 = re.compile("|".join(sorted(关键词词典, key=len, reverse=True)))


def 改关键词(词):
    """返回一组该试的搜索词，按"最可能出货"排队：
    翻译过的整句 → 整句砍到两个词 → 原句 → 第一个词。
    砍词有用是因为 GitHub 仓库搜索要求每个词都命中，「nasa aurora gif」这种三个词的
    多半是 0 条，而「nasa aurora」有货。"""
    翻 = " ".join(词典正则.sub(lambda m: " " + 关键词词典[m.group(0)] + " ", 词).split())
    出 = []
    for 候 in (翻,
               " ".join(翻.split(" ")[:2]) if len(翻.split(" ")) >= 3 else "",
               " ".join(词.split()),
               翻.split(" ")[0] if 翻 else ""):
        if 候 and 候 not in 出:
            出.append(候)
    return 出[:4]


def _搜仓库(词):
    """一次仓库搜索。返回 (仓库列表, 服务器给的提示语)。"""
    数据 = json.loads(_取(搜索地址.format(quote(词))).decode("utf-8", "ignore"))
    return 数据.get("items", []) or [], 数据.get("message", "")


def _取(网址, 超时=25, 大小上限=None):
    请求 = urllib.request.Request(网址, headers={
        "User-Agent": "Mozilla/5.0 (desktop-storage-box)",
        "Accept": "application/vnd.github+json",
    })
    with urllib.request.urlopen(请求, timeout=超时) as 响应:
        if not 大小上限:
            return 响应.read()
        块 = []
        已读 = 0
        while True:
            数据 = 响应.read(65536)
            if not 数据:
                break
            已读 += len(数据)
            if 已读 > 大小上限:
                raise ValueError("包太大，跳过这个仓库")
            块.append(数据)
        return b"".join(块)


def 清理名(名字):
    名字 = "".join(c for c in 名字 if c not in '\\/:*?"<>|').strip()
    return 名字 or "网上动图.gif"


def _规范网址(网址):
    """把链接里不能直接上战场的字符补上转义：
    中文文件名不转 %XX 的话，http.client 会当场 UnicodeEncodeError；
    域名是中文（很少见）走 idna。转坏了就退回原样。"""
    try:
        部 = urlsplit(网址)
        域 = 部.netloc
        if any(ord(c) > 127 for c in 域):
            域 = 域.encode("idna").decode("ascii")
        return urlunsplit((部.scheme, 域, quote(部.path, safe="/%~"),
                           quote(部.query, safe="=&%"), 部.fragment))
    except Exception:
        return 网址


def _取图(网址, 大小上限=12 * 1024 * 1024, 超时=25):
    """直接下一个网址的字节（收外部动图链接用，不带 GitHub 那套头）。"""
    if not 网址.lower().startswith(("http://", "https://")):
        raise ValueError("只认 http/https 链接")
    请求 = urllib.request.Request(_规范网址(网址), headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Accept": "image/gif,image/*;q=0.8,*/*;q=0.5",
    })
    try:
        with urllib.request.urlopen(请求, timeout=超时) as 响应:
            块 = []
            已读 = 0
            while True:
                数据 = 响应.read(65536)
                if not 数据:
                    break
                已读 += len(数据)
                if 已读 > 大小上限:
                    raise ValueError("这张动图超过 12MB，贴上桌面太吃内存，换一张")
                块.append(数据)
    except ValueError:
        raise
    except urllib.error.HTTPError as e:
        raise ValueError(_下不动的话(e.code)) from None
    except (urllib.error.URLError, OSError) as e:
        raise ValueError("连不上这个地址：%s" % (getattr(e, "reason", None) or e)) from None
    return b"".join(块)


def _下不动的话(码):
    """把 HTTP 错误码翻成人话——他看不懂 403 是什么，但他看得懂"要粘图片本身的地址"。"""
    return {401: "这张图要登录才看得到，程序没有你的账号，下不了",
            403: "对方不让直接下（403）：要粘的是图片本身的地址——"
                 "网页上右键图片 →「复制图片地址」再粘进来",
            404: "对方那儿没有这张图（404），换一个链接",
            429: "对方嫌你下太快了（429），过一会儿再试",
            }.get(码, "对方回了个 %d，下不动，换一个链接" % 码)


def 存网上动图(网址, 目标目录, 进度=None):
    """干什么：把他粘进来的链接收成一张桌面动图，返回 (本地路径, 来源网址)。
    怎么用：链接可以是 .gif 直链，也可以是网页——是网页就在里面挑第一张够大的 gif 收下；
            失败抛 ValueError，消息本身就是给人看的话，界面直接显示。
    为什么要有这条：GitHub 仓库搜索翻出来的多半是软件演示图，真想要好看的动图，
            从抖音/微博/壁纸站右键「复制图片地址」粘过来才是最快最稳的路。"""
    网址 = (网址 or "").strip().strip('"').strip("'").strip("《》")
    if not 网址:
        raise ValueError("先粘一个链接进来")
    os.makedirs(目标目录, exist_ok=True)
    if 进度:
        进度("正在打开链接…")
    数据 = _取图(网址)
    GIF头 = (b"GIF87a", b"GIF89a")
    if not 数据.startswith(GIF头):
        # 粘的是网页：把页面里的 gif 地址都挑出来（绝对的、相对的都算），
        # 逐个试，收下第一张"够大、真能下"的——第一张常常是站标或加载圈。
        文本 = 数据.decode("utf-8", "ignore")
        候选 = []
        for u in (re.findall(r"""https?://[^\s"'<>\\]+\.gif""", 文本, re.I)
                  + [urljoin(网址, u) for u in
                     re.findall(r"""["'()]([^"'()<>\s]+\.gif)""", 文本, re.I)]):
            if u.lower().startswith(("http://", "https://")) and u not in 候选:
                候选.append(u)
        if not 候选:
            raise ValueError("这个链接里没有动图：直接粘 .gif 的图地址，"
                             "或者粘一个里面有 gif 的网页")
        中 = None
        for 候 in 候选[:8]:
            try:
                d = _取图(候)
            except Exception as e:
                写日志("网页里的 %s 下不动：%s" % (候[:60], type(e).__name__))
                continue
            if d.startswith(GIF头) and len(d) >= 30_000:
                中 = (候, d)
                break
        if not 中:
            raise ValueError("网页里那 %d 张 gif 都不合适（下不下来、或不到 30KB 太小）"
                             % min(len(候选), 8))
        网址, 数据 = 中
        if 进度:
            进度("这是个网页，收的是里面这张：%s" % 网址[:70])
    if len(数据) < 30_000:
        raise ValueError("这张动图不到 30KB，多半是个小图标，当壁纸会糊")
    名字 = 清理名(os.path.basename(网址.split("?")[0].split("#")[0]) or "网上动图.gif")
    if not 名字.lower().endswith(".gif"):
        名字 += ".gif"
    存到 = os.path.join(目标目录, 名字)
    序 = 2
    while os.path.exists(存到):
        存到 = os.path.join(目标目录, "%s_%d.gif" % (os.path.splitext(名字)[0], 序))
        序 += 1
    with open(存到, "wb") as f:
        f.write(数据)
    if 进度:
        进度("收好了：%s（%d KB）" % (os.path.basename(存到), len(数据) // 1024))
    return 存到, 网址


def 网络找GIF(关键词, 目标目录, 最多=8, 进度=None):
    """搜 GitHub 仓库 → 下它的代码包 → 把包里合格的 gif 抽到 目标目录。
    返回 [(本地路径, 来源仓库全名)]。任何一步失败都只记日志、返回已有结果，绝不炸界面。
    搜索词会自动翻成英文多试一句：GitHub 仓库搜索是"每个词都要命中"，
    直接打「极光壁纸 gif」这种多半是 0 条。未登录限流是每分钟 10 次，所以翻到
    第一条有货的词就停手，不连着打。"""
    os.makedirs(目标目录, exist_ok=True)
    结果 = []
    仓库们 = []
    限流了 = False
    试过的 = []
    for 词 in 改关键词(关键词):
        if 词 in 试过的:
            continue
        试过的.append(词)
        if 进度:
            进度(f"正在搜「{词}」…")
        try:
            找到, 提示 = _搜仓库(词)
        except urllib.error.HTTPError as e:
            写日志(f"GitHub 搜索「{词}」失败：HTTP {e.code}")
            if e.code in (403, 422) and "rate" in str(e.reason or "").lower():
                限流了 = True
            elif e.code == 403:
                限流了 = True
            continue
        except Exception as e:
            写日志(f"GitHub 搜索「{词}」失败：{type(e).__name__}: {e}")
            continue
        if 提示:
            写日志(f"GitHub 搜索「{词}」提示：{提示}")
            if "rate limit" in 提示.lower():
                限流了 = True
        if 找到:
            仓库们 = 找到
            break

    if not 仓库们:
        if 进度:
            if 限流了:
                进度("GitHub 限流了（未登录每分钟只给 10 次搜索），等一分钟再搜")
            else:
                进度(f"没搜到相关仓库。换个词试试，比如 {推荐词}")
        return 结果

    翻过 = 0
    for 项 in 仓库们[:6]:
        if len(结果) >= 最多:
            break
        全名 = 项.get("full_name") or ""
        if not 全名 or int(项.get("size", 10 ** 9)) > 30000:  # 仓库 >30MB 不碰
            continue
        翻过 += 1
        分支 = 项.get("default_branch") or "master"
        if 进度:
            进度(f"正在翻 {全名} …")
        临时 = tempfile.mkdtemp(prefix="壁纸包_")
        包路径 = os.path.join(临时, "包.tar")
        try:
            with open(包路径, "wb") as f:
                f.write(_取(包地址.format(全名, 分支), 大小上限=30 * 1024 * 1024))
            with tarfile.open(包路径) as 塔:
                候选 = [m for m in 塔.getmembers()
                        if m.isfile() and m.name.lower().endswith(".gif")
                        and 200_000 <= m.size <= 8_000_000]
                for 成员 in 候选[:最多 - len(结果)]:
                    try:
                        文件 = 塔.extractfile(成员)
                        数据 = 文件.read() if 文件 else b""
                    except Exception:
                        continue
                    if not 数据.startswith((b"GIF87a", b"GIF89a")):
                        continue
                    存到 = os.path.join(目标目录, 清理名(os.path.basename(成员.name)))
                    if os.path.exists(存到):
                        continue
                    with open(存到, "wb") as f:
                        f.write(数据)
                    结果.append((存到, 全名))
        except Exception as e:
            写日志(f"{全名} 取包失败：{type(e).__name__}: {e}")
        finally:
            try:
                if os.path.exists(包路径):
                    os.remove(包路径)
                os.rmdir(临时)
            except Exception:
                pass
    if 进度:
        if 结果:
            进度(f"翻了 {翻过} 个仓库，找到 {len(结果)} 张")
        else:
            进度(f"翻了 {翻过} 个仓库，里头没有合适的动图"
                 f"（要 200KB~8MB 的 gif），换个词再试：{推荐词}")
    return 结果


# ---------------- 全网搜动图（百度图片 / 360 图片，GitHub 兜底） ----------------
# 他原话："你要全网搜啊，你就给我弄一个搜索引擎在那里一样，或者连百度。"
# 所以这里不再只翻 GitHub 仓库，而是直接问图片搜索引擎要 GIF 原图：
#   1) 百度图片：结果里的 is_gif 字段自己标了是不是动图，replaceUrl 里是原图地址；
#      原图放防盗链时退到它自己的 CDN（thumbURL 带 f=GIF，也是真 GIF）。
#   2) 360 图片：list 里的 img 常常就是 .gif 原图，thumb 是 360 缓存过的 GIF，做备胎。
#   3) 两边都没货才去 GitHub 翻仓库（原来那套保留，慢但稳）。
# ★ 实测坑：百度不先访问一次首页拿 BAIDUID 这个 cookie，会隔几次就回一个空壳；
#   带 cookie 会话后连搜两次都是 30 条、30 条 is_gif=1。
# ★ 360 的 /j 接口要带 Referer=https://image.so.com/，不带有时回空。
# ★ 搜狗图片的 napi 接口直接回 {"info":"forbid"}，别在那上面浪费时间。
# ★ Tenor / Giphy 接口在本机连不上（超时），不做。

_网页会话 = [None]


def _会话():
    """带 cookie 的浏览器会话：百度要先拿 BAIDUID 才肯稳定给结果。"""
    if _网页会话[0] is None:
        罐 = http.cookiejar.CookieJar()
        _网页会话[0] = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(罐))
    return _网页会话[0]


def _网页取(网址, 来源页="", 超时=20, 大小上限=8 * 1024 * 1024):
    """搜图片用的一次请求。来源页（Referer）得带，百度/360 都看这个。"""
    头 = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Accept": "text/html,application/json,application/javascript,*/*;q=0.5",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    }
    if 来源页:
        头["Referer"] = 来源页
    请求 = urllib.request.Request(_规范网址(网址), headers=头)
    with _会话().open(请求, timeout=超时) as 响应:
        块 = []
        已读 = 0
        while True:
            数据 = 响应.read(65536)
            if not 数据:
                break
            已读 += len(数据)
            if 已读 > 大小上限:
                raise ValueError("这个搜索页太大，放弃")
            块.append(数据)
        return b"".join(块)


def _网页搜索词(关键词):
    """搜索引擎吃中文就够，不用翻英文。但补上「动图 / gif」能把它引到动图堆里。
    已经是动图类词就不硬加了，免得"极光动图 动图"。"""
    词 = " ".join((关键词 or "").split())
    if not 词:
        return []
    出 = []
    低 = 词.lower()
    if not any(k in 低 for k in ("gif", "动图", "动态图")):
        出.append(词 + " 动图")
        出.append(词 + " gif")
    出.append(词)
    去重 = []
    for 候 in 出:
        if 候 not in 去重:
            去重.append(候)
    return 去重[:2]


def _项里尺寸(项):
    """百度/360 都带宽高（有时是字符串），拿来做一道"太小不要"的门槛。"""
    try:
        宽 = int(float(str(项.get("width") or 0)))
        高 = int(float(str(项.get("height") or 0)))
        return 宽, 高
    except (TypeError, ValueError):
        return 0, 0


def _搜百度(词, 进度=None):
    """百度图片：is_gif=1 的结果里挑原图；原图可能防盗链，所以同时给 CDN 备胎。"""
    q = quote(词)
    if 进度:
        进度(f"正在百度图片里搜「{词}」…")
    try:
        _网页取("https://image.baidu.com/", 超时=10, 大小上限=2 * 1024 * 1024)
    except Exception:
        pass  # 拿不到 cookie 也照样试，只是成功率低一点
    数据 = _网页取(
        "https://image.baidu.com/search/acjson?tn=resultjson_com&ipn=rj"
        "&ct=201326592&queryWord=%s&word=%s&ie=utf-8&oe=utf-8&pn=0&rn=30&gsm=1e"
        % (q, q), 来源页="https://image.baidu.com/")
    try:
        内容 = json.loads(数据.decode("utf-8", "ignore"))
    except ValueError:
        return []
    出 = []
    for 项 in 内容.get("data") or []:
        if not isinstance(项, dict) or not 项.get("thumbURL"):
            continue
        if not 项.get("is_gif"):
            continue
        宽, 高 = _项里尺寸(项)
        if 宽 and 高 and (宽 < 320 or 高 < 180):
            continue
        原图 = ""
        for 换 in 项.get("replaceUrl") or []:
            if not isinstance(换, dict):
                continue
            网址 = 换.get("ObjURL") or 换.get("ObjUrl") or ""
            if ".gif" in 网址.lower():
                原图 = 网址
                break
        CDN = 项.get("thumbURL") or ""
        if "f=gif" not in CDN.lower():
            CDN = ""
        if 原图:
            出.append((原图, "百度图片", 项))
        if CDN:
            出.append((CDN, "百度图片", 项))
    return 出


def _搜360(词, 进度=None, 页=0):
    """360 图片：img 是原图（常是 .gif），thumb 是它缓存过的 GIF，做备胎。"""
    if 进度:
        进度(f"正在 360 图片里搜「{词}」…")
    数据 = _网页取(
        "https://image.so.com/j?q=%s&src=srp&sn=%d&pn=30"
        % (quote(词), 页 * 30), 来源页="https://image.so.com/")
    try:
        内容 = json.loads(数据.decode("utf-8", "ignore"))
    except ValueError:
        return []
    出 = []
    for 项 in 内容.get("list") or []:
        if not isinstance(项, dict):
            continue
        宽, 高 = _项里尺寸(项)
        if 宽 and 高 and (宽 < 320 or 高 < 180):
            continue
        原图 = str(项.get("img") or "")
        备胎 = str(项.get("thumb") or "")
        if ".gif" in 原图.lower():
            出.append((原图, "360图片", 项))
        if ".gif" in 备胎.lower():
            出.append((备胎, "360图片", 项))
    return 出


def _存一张(数据, 目标目录, 关键词, 序号):
    """把下好的 GIF 字节落盘，名字给人看：关键词_01.gif。"""
    名字 = 清理名("%s_%02d.gif" % ((关键词 or "网上动图").strip()[:20], 序号))
    存到 = os.path.join(目标目录, 名字)
    序 = 2
    while os.path.exists(存到):
        存到 = os.path.join(目标目录, "%s_%d.gif" % (os.path.splitext(名字)[0], 序))
        序 += 1
    with open(存到, "wb") as f:
        f.write(数据)
    return 存到


def _GIF信息(数据):
    """返回 (宽, 高, 帧数)。只读 GIF 的头和块结构，不依赖 Qt，也不解码像素。
    帧数 <2 就是一张静图（很多"gif"其实是静图改了扩展名），当动图壁纸没意义。
    ★ GIF 块结构：0x21=扩展块、0x2C=图像描述块（一帧）、0x3B=结束；
      每块后面都跟一串"长度字节+数据"的子块，遇到 0 长度就结束。"""
    if len(数据) < 13 or not 数据.startswith((b"GIF87a", b"GIF89a")):
        return 0, 0, 0
    宽 = 数据[6] | (数据[7] << 8)
    高 = 数据[8] | (数据[9] << 8)
    标志 = 数据[10]
    位置 = 13
    if 标志 & 0x80:                       # 有全局调色板就跳过
        位置 += 3 * (2 ** ((标志 & 0x07) + 1))
    帧数 = 0

    def 跳子块(位):
        while 位 < len(数据):
            长 = 数据[位]
            位 += 1
            if 长 == 0:
                return 位
            位 += 长
        return 位

    while 位置 < len(数据):
        块 = 数据[位置]
        if 块 == 0x3B:
            break
        if 块 == 0x21:                     # 扩展块（含图形控制块）
            位置 += 2                      # 0x21 + 扩展类型
            位置 = 跳子块(位置)
        elif 块 == 0x2C:                   # 一帧图像
            帧数 += 1
            位置 += 10                     # 描述块 10 字节（含打包位）
            打包 = 数据[位置 - 1]
            if 打包 & 0x80:                # 局部调色板在 LZW 码长之前
                位置 += 3 * (2 ** ((打包 & 0x07) + 1))
            位置 += 1                      # LZW 最小码长
            位置 = 跳子块(位置)
        else:
            break
    return 宽, 高, 帧数


def 全网找GIF(关键词, 目标目录, 最多=10, 进度=None, 要GitHub兜底=True):
    """干什么：像搜索引擎一样全网找 GIF——百度图片 → 360 图片 → GitHub 兜底，
            把合格（真 GIF、30KB~12MB）的图存进 目标目录。
    返回 [(本地路径, 来源名)]，来源名就是"百度图片/360图片/GitHub"。
    任何一步失败只记日志、继续下一家，绝不把界面炸掉；
    他关了对话框时，进度回调里会抛 收图中止，这里原样往外传。"""
    os.makedirs(目标目录, exist_ok=True)
    候选 = []
    见过网址 = set()
    搜到了 = False
    for 引擎 in (_搜百度, _搜360):
        for 词 in _网页搜索词(关键词):
            try:
                找到 = 引擎(词, 进度)
            except Exception as e:
                if type(e).__name__ == "收图中止":
                    raise
                写日志(f"{引擎.__name__}「{词}」失败：{type(e).__name__}: {e}")
                continue
            if 找到:
                搜到了 = True
                for 网址, 来源, 项 in 找到:
                    if 网址 not in 见过网址:
                        见过网址.add(网址)
                        候选.append((网址, 来源, 项))
                break  # 这家有货就不再换词，免得请求太密被限流
        if len(候选) >= 最多 * 6:
            break

    if 进度:
        进度(f"搜到 {len(候选)} 个候选，正在挑能下的 GIF…")

    结果 = []
    摘要们 = set()
    试过 = 0
    次数上限 = max(10, 最多 + 6)   # 别为了几张图无休止地试，界面会等太久
    for 网址, 来源, _项 in 候选:
        if len(结果) >= 最多 or 试过 >= 次数上限:
            break
        试过 += 1
        try:
            数据 = _取图(网址, 大小上限=12 * 1024 * 1024, 超时=20)
        except Exception as e:
            if type(e).__name__ == "收图中止":
                raise
            写日志(f"候选下不动（{来源}）：{type(e).__name__}: {e}")
            continue
        if not 数据.startswith((b"GIF87a", b"GIF89a")):
            continue
        if len(数据) < 30_000:
            continue
        宽, 高, 帧数 = _GIF信息(数据)
        if 帧数 < 2:
            写日志(f"跳过一张不会动的 GIF（{来源}，{宽}x{高}）")
            continue
        if 宽 < 320 or 高 < 180:
            写日志(f"跳过一张太小的 GIF（{来源}，{宽}x{高}）")
            continue
        摘要 = hashlib.md5(数据).hexdigest()
        if 摘要 in 摘要们:
            continue  # 同一张图在两家引擎/两个尺寸里重复出现
        摘要们.add(摘要)
        存到 = _存一张(数据, 目标目录, 关键词, len(结果) + 1)
        结果.append((存到, 来源))
        if 进度:
            进度(f"已下 {len(结果)} 张（{来源}）…")

    if not 结果 and 要GitHub兜底:
        if 进度:
            进度("百度/360 没找到能下的，再去 GitHub 翻翻…")
        结果 = 网络找GIF(关键词, 目标目录, 最多=最多, 进度=进度)

    if 进度:
        if 结果:
            进度(f"拿到 {len(结果)} 张，双击一张就能贴上桌面。")
        elif 搜到了:
            进度("搜到了一些图，但都下不下来、不是会动的 GIF 或太小，换个词再试：%s" % 推荐词)
        else:
            进度("网上没搜到，换个词试试：%s" % 推荐词)
    return 结果
