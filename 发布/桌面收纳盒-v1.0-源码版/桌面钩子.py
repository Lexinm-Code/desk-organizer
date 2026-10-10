# -*- coding: utf-8 -*-
"""全局鼠标钩子：检测"双击桌面空白处"。

在后台线程安装 Windows 低级鼠标钩子（WH_MOUSE_LL），
通过 WindowFromPoint + LVM_HITTEST 区分"点在桌面图标上"与"点在桌面空白处"，
只有双击空白处才会触发回调。

注意：回调在钩子线程中执行，主程序里应通过 Qt 信号转回主线程。
"""
import ctypes
import os
import sys
import time
import threading
from ctypes import wintypes

import win32api
import win32con
import win32gui

# 钩子模块和主程序共用同一份启动日志，方便把"双击桌面"的判定过程也记录下来
_钩子日志路径 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "启动日志.txt")


def 写钩子日志(内容):
    """往启动日志追加一行（钩子运行在独立线程，出错也不能影响鼠标正常使用）"""
    try:
        with open(_钩子日志路径, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%H:%M:%S')}] [鼠标钩子] {内容}\n")
    except Exception:
        pass

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

# --- ctypes 签名，避免 64 位下句柄截断 ---
WH_MOUSE_LL = 14
WM_LBUTTONDOWN = 0x0201
WM_QUIT = 0x0012
LLMHF_INJECTED = 0x0001
GA_ROOT = 2
LVM_HITTEST = 0x1012  # LVM_FIRST + 18

PROCESS_VM_OPERATION = 0x0008
PROCESS_VM_READ = 0x0010
PROCESS_VM_WRITE = 0x0020
MEM_COMMIT_RESERVE = 0x3000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04

user32.SetWindowsHookExW.restype = wintypes.HHOOK
user32.SetWindowsHookExW.argtypes = [ctypes.c_int, ctypes.c_void_p, wintypes.HINSTANCE, wintypes.DWORD]
user32.UnhookWindowsHookEx.argtypes = [wintypes.HHOOK]
user32.CallNextHookEx.restype = ctypes.c_longlong
user32.CallNextHookEx.argtypes = [wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.SendMessageW.restype = ctypes.c_longlong
user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
kernel32.GetCurrentThreadId.restype = wintypes.DWORD
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.VirtualAllocEx.restype = ctypes.c_void_p
kernel32.VirtualAllocEx.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD, wintypes.DWORD]
kernel32.VirtualFreeEx.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD]
kernel32.WriteProcessMemory.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
kernel32.ReadProcessMemory.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]


class 点(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class 鼠标结构(ctypes.Structure):
    """MSLLHOOKSTRUCT"""
    _fields_ = [
        ("pt", 点),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class 命中信息(ctypes.Structure):
    """LVHITTESTINFO"""
    _fields_ = [
        ("pt", 点),
        ("flags", wintypes.UINT),
        ("iItem", ctypes.c_int),
        ("iSubItem", ctypes.c_int),
    ]


钩子过程类型 = ctypes.WINFUNCTYPE(ctypes.c_longlong, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)


def _取类名(hwnd):
    缓冲 = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, 缓冲, 256)
    return 缓冲.value


def _找桌面图标列表(顶级):
    """在 Progman/WorkerW 下定位桌面图标的 SysListView32 窗口"""
    定义视图 = win32gui.FindWindowEx(顶级, None, "SHELLDLL_DefView", None)
    if 定义视图:
        return win32gui.FindWindowEx(定义视图, None, "SysListView32", None)
    # 壁纸软件场景：图标列表可能被挪到 Progman 的兄弟 WorkerW 下
    桌面根 = win32gui.FindWindow("Progman", None)
    if not 桌面根:
        return None
    枚举 = 0
    while True:
        枚举 = win32gui.FindWindowEx(None, 枚举, "WorkerW", None)
        if not 枚举:
            return None
        定义视图 = win32gui.FindWindowEx(枚举, None, "SHELLDLL_DefView", None)
        if 定义视图:
            return win32gui.FindWindowEx(定义视图, None, "SysListView32", None)


def _是桌面空白(坐标):
    """判断屏幕坐标是否位于桌面图标区域的空白处（而不是某个图标上）"""
    try:
        hwnd = win32gui.WindowFromPoint((int(坐标[0]), int(坐标[1])))
        if not hwnd:
            return False
        顶级 = win32gui.GetAncestor(hwnd, GA_ROOT)
        if _取类名(顶级) not in ("Progman", "WorkerW"):
            return False
        列表 = _找桌面图标列表(顶级)
        if not 列表:
            # 找不到图标列表时，退化为"在桌面上就算空白"
            return True
        # 向 explorer 的图标列表跨进程发送命中测试：命中图标项则不算空白
        信息 = 命中信息()
        信息.pt.x = int(坐标[0])
        信息.pt.y = int(坐标[1])
        进程id = wintypes.DWORD()
        user32.GetWindowThreadProcessId(列表, ctypes.byref(进程id))
        进程 = kernel32.OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, False, 进程id.value)
        if not 进程:
            return True
        try:
            远程 = kernel32.VirtualAllocEx(进程, None, ctypes.sizeof(信息), MEM_COMMIT_RESERVE, PAGE_READWRITE)
            if not 远程:
                return True
            try:
                实际 = ctypes.c_size_t()
                kernel32.WriteProcessMemory(进程, 远程, ctypes.byref(信息), ctypes.sizeof(信息), ctypes.byref(实际))
                命中 = user32.SendMessageW(列表, LVM_HITTEST, 0, 远程)
                kernel32.ReadProcessMemory(进程, 远程, ctypes.byref(信息), ctypes.sizeof(信息), ctypes.byref(实际))
                return 命中 < 0  # -1 表示未命中任何图标项（空白处）
            finally:
                kernel32.VirtualFreeEx(进程, 远程, 0, MEM_RELEASE)
        finally:
            kernel32.CloseHandle(进程)
    except Exception:
        return False


class 桌面双击钩子:
    """后台线程安装低级鼠标钩子，双击桌面空白处时执行回调：回调(x, y)"""

    def __init__(self, 回调):
        self._回调 = 回调
        self._钩子 = None
        self._线程id = 0
        self._上次按下 = None  # (时间, x, y)
        self._过程 = 钩子过程类型(self._钩子过程)
        self._线程 = threading.Thread(target=self._运行, name="桌面双击钩子", daemon=True)
        # 安装状态：None=还没装完，True=成功，False=失败（主程序会读这个写进启动日志）
        self.安装状态 = None

    def start(self):
        self._线程.start()

    def stop(self):
        if self._线程id:
            user32.PostThreadMessageW(self._线程id, WM_QUIT, 0, 0)
        if self._线程.is_alive():
            self._线程.join(timeout=2)

    def _运行(self):
        self._线程id = kernel32.GetCurrentThreadId()
        # 低级钩子回调在本进程内执行，hMod 传 None 即可
        # （不能传 GetModuleHandleW 的返回值：64 位下会被截断成无效句柄，导致安装失败报 126）
        self._钩子 = user32.SetWindowsHookExW(WH_MOUSE_LL, self._过程, None, 0)
        if not self._钩子:
            # 安装失败：记下系统错误码（例如 126=句柄无效，1428=钩子链异常）
            self.安装状态 = False
            写钩子日志(f"安装失败！错误码 {ctypes.get_last_error()}，双击桌面功能不可用")
            return
        self.安装状态 = True
        写钩子日志("安装成功，正在监听双击桌面空白处")
        消息 = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(消息), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(消息))
            user32.DispatchMessageW(ctypes.byref(消息))
        user32.UnhookWindowsHookEx(self._钩子)
        self._钩子 = None

    def _钩子过程(self, 代码, 消息, 参数):
        try:
            if 代码 == 0 and 消息 == WM_LBUTTONDOWN:
                信息 = 鼠标结构.from_address(参数)
                if not (信息.flags & LLMHF_INJECTED):
                    # 只处理真实鼠标点击（忽略程序模拟的点击，防止误触发）
                    self._分析按下(信息)
        except Exception as e:
            # 判定过程万一出错要记录下来，绝不能像之前一样静默吞掉
            import traceback as _tb
            写钩子日志("判定过程发生异常：" + "".join(_tb.format_exception(type(e), e, e.__traceback__)))
        return user32.CallNextHookEx(None, 代码, 消息, 参数)

    def _分析按下(self, 信息):
        本次 = (信息.time, 信息.pt.x, 信息.pt.y)
        上次 = self._上次按下
        self._上次按下 = 本次
        if not 上次:
            return
        # 系统规定的双击判定标准：最大间隔毫秒数、两次点击允许的最大移动距离（像素）
        # 注意：GetDoubleClickTime 在 win32gui 模块（不能写成 win32api，那个模块没这个函数）
        间隔 = (本次[0] - 上次[0]) & 0xFFFFFFFF
        间隔上限 = win32gui.GetDoubleClickTime()
        dx = abs(本次[1] - 上次[1])
        dy = abs(本次[2] - 上次[2])
        x上限 = win32api.GetSystemMetrics(win32con.SM_CXDOUBLECLK)
        y上限 = win32api.GetSystemMetrics(win32con.SM_CYDOUBLECLK)
        # 点得不够快、手移动超范围：属于正常的两次单击，不写日志（太常见，避免刷屏）
        if 间隔 > 间隔上限:
            return
        if dx > x上限 or dy > y上限:
            return
        # 到这里说明：时间够近、位置够近，已判定为"双击"，再确认是不是点在桌面空白处
        是空白 = _是桌面空白((本次[1], 本次[2]))
        if 是空白:
            写钩子日志(f"双击桌面空白处 坐标({本次[1]},{本次[2]}) → 通知主窗口")
            self._回调(本次[1], 本次[2])
        else:
            # 点在桌面图标上、任务栏或其他窗口上都不算
            写钩子日志(f"双击成立但不在桌面空白处 坐标({本次[1]},{本次[2]}) → 不弹出")
