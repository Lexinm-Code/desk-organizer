# -*- coding: utf-8 -*-
"""桌面收纳盒

一个悬浮的桌面图标收纳盒：
- 双击桌面空白处 / 双击托盘图标 / 全局热键 唤出，鼠标移开自动隐藏
- 右上角图钉可固定窗口（固定后不再自动隐藏，方便往里拖图标）
- 拖入的图标一律以"快捷方式"存放，绝不移动原文件
- 桌面快捷方式拖入时会从桌面移除原图标，右键盒内图标可"恢复到原来的位置"
- 分类可自定义（新建 / 重命名 / 删除），图标可在分类间拖动
"""
import json
import os
import re
import shutil
import sys
import time
import ctypes
import traceback
from ctypes import wintypes

# ---------------- 启动日志（排查"启动慢/开不出来"用） ----------------
# 日志文件和程序放一起，每次启动会清空重写，只保留最近一次启动的记录
# 打包后 sys.frozen=True，sys.executable 的父目录就是 exe 所在目录（用户看得见的地方）
if getattr(sys, 'frozen', False):
    _启动日志路径 = os.path.join(os.path.dirname(sys.executable), "启动日志.txt")
else:
    _启动日志路径 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "启动日志.txt")
# 程序刚跑起来的起点（用高精度计时器，能精确到 0.01 秒）
_启动零点 = time.perf_counter()


def 写启动日志(阶段):
    """把当前启动阶段和已用时间写进日志文件，方便对照看卡在哪一步"""
    try:
        with open(_启动日志路径, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%H:%M:%S')}] 已用 {time.perf_counter() - _启动零点:5.2f} 秒  {阶段}\n")
    except Exception:
        pass  # 日志本身绝不能影响程序启动，写失败就静默跳过


def _记录崩溃(异常类型, 异常对象, 调用栈):
    """pythonw 启动没有黑窗口，程序一旦报错直接闪退且看不到原因；
    这里把崩溃信息完整写进日志文件，排查问题全靠它"""
    try:
        with open(_启动日志路径, "a", encoding="utf-8") as f:
            f.write("！！！程序崩溃：\n")
            f.write("".join(traceback.format_exception(异常类型, 异常对象, 调用栈)))
    except Exception:
        pass


# 把上面的崩溃记录函数注册为全局异常处理器
sys.excepthook = _记录崩溃
# 每次启动先清空旧日志（用"写模式"打开一次即可清空），只保留本次启动记录
open(_启动日志路径, "w", encoding="utf-8").close()
# 第一条日志：脚本刚开始执行（此时第三方库一个都还没加载）
写启动日志("脚本开始执行")

from PyQt5.QtCore import (Qt, QRect, QSize, QTimer, QEvent, pyqtSignal, QMimeData,
                          QFileInfo, QSharedMemory, QAbstractNativeEventFilter)
写启动日志("PyQt5.QtCore 加载完成")
from PyQt5.QtGui import (QColor, QCursor, QFont, QIcon, QImage, QImageReader,
                         QKeySequence, QDrag, QPainter, QPixmap, QGuiApplication)
from PyQt5.QtWidgets import (QApplication, QAction, QCheckBox, QDialog, QFrame,
                             QGridLayout, QHBoxLayout, QInputDialog, QLabel, QMenu,
                             QMessageBox, QPushButton, QKeySequenceEdit, QScrollArea,
                             QSpinBox, QStyle, QSystemTrayIcon, QTabBar, QTabWidget,
                             QToolButton, QVBoxLayout, QWidget, QFileIconProvider,
                             QStackedWidget, QLineEdit, QComboBox)
写启动日志("PyQt5 界面库全部加载完成")

import win32api
import win32con
import win32com.client
import win32gui  # 用来做"无边框窗口拖边缘缩放"时的坐标换算
import winsound  # 内置库：播 Windows 系统音效，给盒内图标做音效提示
写启动日志("win32 系统库加载完成")

from 桌面钩子 import 桌面双击钩子
写启动日志("鼠标钩子模块加载完成，开始建界面")

# ---------------- 常量与配置 ----------------

# 程序目录：打包后（PyInstaller）sys.frozen=True，exe 在 sys.executable 的父目录；
# 开发时 sys.frozen 不存在，用 __file__ 的父目录。
# 注意：打包后 PyInstaller 会把临时解压目录设成 sys._MEIPASS（那是 _internal/），
# 我们不要用那个，要用 sys.executable 的父目录——那才是用户放 exe 的地方。
if getattr(sys, 'frozen', False):
    程序目录 = os.path.dirname(sys.executable)   # 打包后：exe 所在目录
else:
    程序目录 = os.path.dirname(os.path.abspath(__file__))  # 开发时：源码所在目录
数据目录 = os.path.join(程序目录, "收纳盒数据")
版本 = "1.0.3"  # 托盘"关于"对话框显示用（1.0.3：图片按需小尺寸解码 + 图标缓存 + 系统图标取够用档）
配置文件 = os.path.join(数据目录, "配置.json")
内部拖放格式 = "application/x-shouna-tiaomu"
非法文件字符 = re.compile(r'[\\/:*?"<>|]')


def 获取桌面路径():
    """用 Windows API 获取用户桌面路径（兼容 OneDrive 重定向）。
    SHGetKnownFolderPath(FOLDERID_Desktop) 比读注册表或硬拼
    C:\\Users\\xxx\\Desktop 可靠得多——企业域账号或 OneDrive 同步桌面
    都会把桌面重定向到别的地方。
    拿不到就 fallback 到 os.path.expanduser("~\\Desktop")。"""
    try:
        import ctypes
        from ctypes import wintypes
        # FOLDERID_Desktop 的 GUID
        FOLDERID_Desktop = ctypes.GUID(
            0xB4BFCC3A, 0xDB2C, 0x4BC0,
            (0xA9, 0x47, 0x97, 0x1C, 0xAD, 0xF9, 0xB1, 0x90)
        )
        SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
        SHGetKnownFolderPath.argtypes = [
            ctypes.POINTER(ctypes.GUID),  # rfid
            wintypes.DWORD,               # dwFlags
            wintypes.HANDLE,              # hToken
            ctypes.POINTER(wintypes.LPWSTR),  # ppszPath
        ]
        SHGetKnownFolderPath.restype = ctypes.c_long
        buf = wintypes.LPWSTR()
        hr = SHGetKnownFolderPath(ctypes.byref(FOLDERID_Desktop), 0, None, ctypes.byref(buf))
        if hr == 0 and buf.value:
            # 释放 CoTaskMemAlloc 分配的内存
            ctypes.windll.ole32.CoTaskMemFree(buf)
            return os.path.normpath(buf.value)
    except Exception:
        pass
    return os.path.normpath(os.path.expanduser("~\\Desktop"))


# 启动时就把桌面路径算好（后面拖入判断会频繁用到）
桌面路径 = 获取桌面路径()


def 是否在桌面(路径):
    """判断一个绝对路径是否在桌面目录下（包括桌面子文件夹）。
    先 normpath 统一斜杠，再大小写不敏感比较前缀。"""
    if not 桌面路径 or not 路径:
        return False
    路径 = os.path.normpath(os.path.abspath(路径)).lower()
    桌面 = 桌面路径.lower()
    # 直接就在桌面根目录：C:\Users\xxx\Desktop\xxx → True
    # 或桌面子文件夹：C:\Users\xxx\Desktop\新建文件夹\xxx → True
    return 路径 == 桌面 or 路径.startswith(桌面 + os.sep)


# ---------------- 开机自启（Windows 注册表实现）----------------

# 注册表 Run 键路径：用户级（HKCU），不需要管理员权限，每个用户独立
自启注册表键 = r"Software\Microsoft\Windows\CurrentVersion\Run"
# 注册表值名（随便取，只要唯一就行）
自启值名 = "桌面收纳盒"

def _取启动命令():
    """拼出开机自启时要执行的命令行字符串。
    注意：Run 注册表的值是系统直接执行的（相当于 CreateProcess），
    不经过 cmd！之前写 "cmd.exe /c start ... bat" 会在开机时闪一个
    黑窗口（重启后弹"鬼东西"的根源）；而且 "start" 开头在 Run 键里
    根本没人解析，exe 版那样写自启其实是失效的。
    正确姿势：开发时用 pythonw.exe 直接跑脚本（pythonw 没有控制台，
    开机不会弹任何窗口）；打包后直接写 exe 路径。"""
    if getattr(sys, 'frozen', False):
        # 打包后：注册表直接写 exe 路径，系统自己能执行
        return f'"{sys.executable}"'
    # 开发时：把 python.exe 换成同目录的 pythonw.exe，直接带脚本路径
    解释器 = sys.executable
    if 解释器.lower().endswith("python.exe"):
        pythonw = os.path.join(os.path.dirname(解释器), "pythonw.exe")
        if os.path.exists(pythonw):
            解释器 = pythonw
    脚本 = os.path.abspath(__file__)
    return f'"{解释器}" "{脚本}"'

def 设开机自启(开启):
    """开启或关闭开机自启。写入/删除 HKCU Run 键下的一个值。
    返回 True 表示成功，False 表示失败（记日志不弹框）。"""
    try:
        import winreg  # Python 内置
        # OpenKey 的 0 表示 KEY_READ，KEY_SET_VALUE 需要同时有写权限
        访问 = winreg.KEY_SET_VALUE if 开启 else winreg.KEY_SET_VALUE
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 自启注册表键, 0, winreg.KEY_SET_VALUE) as 键:
            if 开启:
                winreg.SetValueEx(键, 自启值名, 0, winreg.REG_SZ, _取启动命令())
            else:
                try:
                    winreg.DeleteValue(键, 自启值名)
                except FileNotFoundError:
                    pass  # 本来就没设过，不报错
        return True
    except Exception:
        return False

def 查开机自启():
    """查注册表当前是否有自启项。返回 True/False。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 自启注册表键, 0, winreg.KEY_READ) as 键:
            try:
                winreg.QueryValueEx(键, 自启值名)
                return True
            except FileNotFoundError:
                return False
    except Exception:
        return False

# 窗口默认尺寸（用户在"设置"里改过、或拖边缘缩放后，会记住新尺寸写进配置）
默认宽 = 320
默认高 = 460
# 窗口允许缩放到的最小尺寸（再小就没法正常看内容了）
最小宽 = 260
最小高 = 320
# 判定"鼠标贴到窗口边缘、进入缩放状态"的边带宽度（物理像素）
缩放宽 = 7
# 图标默认间距（像素）：图标之间的空隙 + 页面内边距，用户可在设置里改
默认图标间距 = 6
# 图标按钮固定大小（盒内图标按钮本身的尺寸）
图标按钮宽 = 62
图标按钮高 = 68

配置 = {}


def 加载配置():
    global 配置
    try:
        with open(配置文件, "r", encoding="utf-8") as f:
            数据 = json.load(f)
    except Exception:
        数据 = {}
    配置 = {
        "热键": 数据.get("热键", "Ctrl+Alt+S"),
        "固定": bool(数据.get("固定", False)),
        # 记住的窗口尺寸，第一次使用时用默认尺寸
        "宽": int(数据.get("宽", 默认宽)),
        "高": int(数据.get("高", 默认高)),
        # 图标之间的间距 + 页面内边距
        "图标间距": int(数据.get("图标间距", 默认图标间距)),
        # 音效提示开关（悬停"叮" / 单击"嗒" / 双击"咚"），默认开
        "音效": bool(数据.get("音效", True)),
        # 常用置顶：打开次数多的图标自动排最前（默认关，避免和用户手动排序打架）
        "常用置顶": bool(数据.get("常用置顶", False)),
        # 按类型自动归类：拖入文件时按后缀分进 文档/图片/视频/音乐/程序/压缩包/其他
        "自动分类": bool(数据.get("自动分类", True)),
        # 屏幕边缘唤出：鼠标顶住窗口停靠的那条屏幕边约 0.3 秒自动弹出（拖到哪边认哪边）
        "边缘唤出": bool(数据.get("边缘唤出", True)),
        # 配色主题：深色 / 浅色 / 玻璃
        "主题": str(数据.get("主题", "深色")),
        "分类": 数据.get("分类") or [{"名称": "常用", "条目": []}],
        # 类型分类只预置一次：用户之后删掉的分类不会被重新造出来
        "类型分类已预置": bool(数据.get("类型分类已预置", False)),
    }
    for 分类 in 配置["分类"]:
        分类.setdefault("名称", "未命名")
        分类.setdefault("条目", [])
    if not 配置["类型分类已预置"]:
        # 第一次启动：左边分类栏常驻建好"全部"总览 + 七种类型分类。
        # "全部"是虚拟分类（聚合视图，不存文件），永远排最前面
        if not any(c.get("虚拟") for c in 配置["分类"]):
            配置["分类"].insert(0, {"名称": "全部", "条目": [], "虚拟": True})
        for 类型名, _后缀组 in 类型分类表:
            if not any(c["名称"] == 类型名 for c in 配置["分类"]):
                配置["分类"].append({"名称": 类型名, "条目": []})
        if not any(c["名称"] == "其他" for c in 配置["分类"]):
            配置["分类"].append({"名称": "其他", "条目": []})
        for 分类 in 配置["分类"]:
            try:
                os.makedirs(os.path.join(数据目录, 分类["名称"]), exist_ok=True)
            except Exception:
                pass
        配置["类型分类已预置"] = True
        保存配置()
    # 双区导航归属（每次加载重算）：预置类型分类和"全部"走左边按钮列，
    # 用户自己的分类（常用/魔兽争霸/以后新建的）走顶部标签栏
    类型名集 = {名 for 名, _组 in 类型分类表} | {"其他"}
    for 分类 in 配置["分类"]:
        分类["类型"] = bool(分类.get("虚拟")) or 分类["名称"] in 类型名集


def 保存配置():
    os.makedirs(数据目录, exist_ok=True)
    with open(配置文件, "w", encoding="utf-8") as f:
        json.dump(配置, f, ensure_ascii=False, indent=2)


# ---------------- 音效提示 ----------------
# 清脆小"叮"：不用系统那几声浑厚提示音，改用纯标准库（wave+math）现场合成
# 三个短促高音拨弦 wav（指数衰减、尾巴干脆、音量偏低），winsound 从内存异步
# 播放——不卡界面、不依赖磁盘文件、不跟随系统声音方案，音量大小由合成参数控制。
# 设置里关掉"音效提示"后（配置["音效"]=False）全部静音。
_悬停音效时刻 = [0.0]  # 悬停限流：鼠标快速划过一串图标时不要"哒哒哒"连响
_悬停静音截止 = [0.0]  # 幻影悬停静音到期时刻：此刻之前的悬停音一律不播
_幻影悬停静音秒 = 0.45  # 弹出盒子 / 重建网格后的静音窗口长度（秒）


def _静音悬停音(秒=None):
    """弹出盒子、重建图标网格之后掐一小段悬停音，专治"手没动却叮一声"。

    盒子是贴着鼠标弹出的（弹出() 把窗口中心挪到指针位置），图标等于凭空出现在指针底下，
    Qt 会给这个新出现的控件补发一次 enterEvent —— 那不是用户去悬停，是幻影。
    重建网格() 整片换新控件同理。真人把鼠标移到某个图标上，必然发生在这 0.45 秒之后，
    所以真悬停照样有声音，只是不再被系统自己冒出来的那一下骗响。
    """
    _悬停静音截止[0] = time.perf_counter() + (_幻影悬停静音秒 if 秒 is None else 秒)


def _合成wav(频率, 时长毫秒, 音量=0.18):
    """合成单声道 22050Hz 16bit 的短拨弦音字节：基频+八度泛音，指数衰减包络"""
    import math
    import struct
    import wave
    import io as _io
    采样率 = 22050
    缓冲 = _io.BytesIO()
    with wave.open(缓冲, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(采样率)
        帧 = bytearray()
        总数 = int(采样率 * 时长毫秒 / 1000.0)
        # 短音衰减更快：只留"叮"的一下，没有嗡嗡拖尾
        衰减 = 40.0 if 时长毫秒 <= 80 else 26.0
        for i in range(总数):
            t = i / 采样率
            包络 = math.exp(-t * 衰减)
            采样 = (math.sin(2 * math.pi * 频率 * t) * 0.72 +
                    math.sin(2 * math.pi * 频率 * 2 * t) * 0.28)  # 八度泛音让音色更透亮
            值 = int(采样 * 包络 * 音量 * 32767)
            帧 += struct.pack("<h", max(-32767, min(32767, 值)))
        w.writeframes(bytes(帧))
    return 缓冲.getvalue()


# 注意：Windows 的 PlaySound 不支持"内存 + 异步"组合（SND_MEMORY|SND_ASYNC 直接报错），
# 所以启动时把合成音写成 wav 文件放在数据目录，播放走"文件 + 异步"这条路。
音效目录 = os.path.join(数据目录, "音效")
_音效规格 = {
    "悬停": (1318.5, 90),   # E6：轻柔清亮"叮~"
    "单击": (1568.0, 60),   # G6：更短更脆"哒"
    "双击": (1046.5, 140),  # C6：稍饱满的"叮咚"
}


def _准备音效文件():
    """合成并写盘；已存在就跳过（想调音色删掉这些 wav 重启即可重新生成）"""
    os.makedirs(音效目录, exist_ok=True)
    for 名称, (频率, 时长) in _音效规格.items():
        文件 = os.path.join(音效目录, 名称 + ".wav")
        if not os.path.exists(文件):
            with open(文件, "wb") as f:
                f.write(_合成wav(频率, 时长, 音量=0.25))


try:
    _准备音效文件()
except Exception:
    pass  # 合成/写盘失败就静音兜底，绝不弹窗打扰


def _播放音效(名称):
    if not 配置.get("音效", True):
        return
    文件 = os.path.join(音效目录, 名称 + ".wav")
    if not os.path.exists(文件):
        return
    try:
        # 文件 + 异步：立即返回不卡界面，新声音顶掉未播完的旧声音
        winsound.PlaySound(文件, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except Exception:
        pass


def 播放悬停音效():
    """鼠标移到图标上：轻柔“叮”，0.18 秒内不重复"""
    此刻 = time.perf_counter()
    if 此刻 < _悬停静音截止[0]:
        return  # 幻影悬停：刚弹出/刚重建，指针没动也不该响
    if 此刻 - _悬停音效时刻[0] < 0.18:
        return
    _悬停音效时刻[0] = 此刻
    _播放音效("悬停")


def 播放单击音效():
    """单击图标：一声清脆短音“哒”"""
    _播放音效("单击")


def 播放双击音效():
    """双击图标：稍饱满的“叮咚”一声"""
    _播放音效("双击")


# ---------------- 按类型自动归类 ----------------

类型分类表 = [
    ("文档", (".txt", ".md", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
              ".pdf", ".rtf", ".csv")),
    ("图片", (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".ico", ".tiff", ".svg")),
    ("视频", (".mp4", ".mkv", ".avi", ".mov", ".flv", ".wmv", ".webm", ".mpg")),
    ("音乐", (".mp3", ".wav", ".flac", ".m4a", ".aac", ".ogg", ".mid")),
    ("程序", (".exe", ".lnk", ".msi", ".bat", ".cmd", ".url")),
    ("压缩包", (".zip", ".rar", ".7z", ".tar", ".gz", ".bz2")),
]


def 取类型分类(路径):
    """按文件后缀判断该进哪个自动分类；文件夹和认不出的一律进"其他" """
    if os.path.isdir(路径):
        return "其他"
    后缀 = os.path.splitext(路径)[1].lower()
    for 名, 后缀组 in 类型分类表:
        if 后缀 in 后缀组:
            return 名
    return "其他"


def 确保分类(名称):
    """在配置里找同名分类，没有就现场新建（含建文件夹），返回分类 dict"""
    分类 = next((c for c in 配置["分类"] if c["名称"] == 名称), None)
    if 分类 is None:
        分类 = {"名称": 名称, "条目": []}
        配置["分类"].append(分类)
        try:
            os.makedirs(os.path.join(数据目录, 名称), exist_ok=True)
        except Exception:
            pass
    return 分类


# ---------------- 快捷方式与文件辅助 ----------------

_shell = None


def 取shell():
    global _shell
    if _shell is None:
        _shell = win32com.client.Dispatch("WScript.Shell")
    return _shell


def 建快捷方式(lnk路径, 目标路径):
    """在 lnk路径 处创建指向 目标路径 的快捷方式（原文件不动）"""
    s = 取shell().CreateShortCut(lnk路径)
    s.Targetpath = 目标路径
    if os.path.isdir(目标路径):
        s.WorkingDirectory = 目标路径
    else:
        s.WorkingDirectory = os.path.dirname(目标路径)
    s.save()


def 避让路径(路径):
    """路径已存在时追加序号，避免覆盖"""
    if not os.path.exists(路径):
        return 路径
    根, 后缀 = os.path.splitext(路径)
    序号 = 1
    while True:
        新 = f"{根}({序号}){后缀}"
        if not os.path.exists(新):
            return 新
        序号 += 1


def 唯一lnk名(文件夹, 基名):
    return os.path.basename(避让路径(os.path.join(文件夹, 基名 + ".lnk")))


def 清理名称(名字, 兜底="未命名"):
    名字 = 非法文件字符.sub("", str(名字)).strip()
    return 名字 or 兜底


# ---------------- 图标 ----------------

图标提供者 = QFileIconProvider()


def 生成程序图标(文字="收"):
    画布 = QPixmap(64, 64)
    画布.fill(Qt.transparent)
    画笔 = QPainter(画布)
    画笔.setRenderHint(QPainter.Antialiasing)
    画笔.setPen(Qt.NoPen)
    画笔.setBrush(QColor("#4f8cff"))
    画笔.drawRoundedRect(4, 4, 56, 56, 16, 16)
    画笔.setPen(QColor("#ffffff"))
    画笔.setFont(QFont("Microsoft YaHei UI", 26, QFont.Bold))
    画笔.drawText(画布.rect(), Qt.AlignCenter, 文字)
    画笔.end()
    return QIcon(画布)


def _不透明包围盒(图):
    """逐像素扫 alpha，返回不透明内容的最小外接矩形；全不透明则返回整图矩形"""
    w, h = 图.width(), 图.height()
    左, 上, 右, 下 = w, h, -1, -1
    for y in range(h):
        for x in range(w):
            if (图.pixel(x, y) >> 24) & 0xFF:
                if x < 左:
                    左 = x
                if x > 右:
                    右 = x
                if y < 上:
                    上 = y
                if y > 下:
                    下 = y
    if 右 < 0:
        return 图.rect()  # 全透明兜底当无留白
    return QRect(左, 上, 右 - 左 + 1, 下 - 上 + 1)


def _规格化(图标, 边=36):
    """系统图标原生尺寸/画布留白各不相同，直接缩放会显得一大一小。
    取最大的一张原生图，裁掉四周透明留白，再等比撑满 边x边 画布，视觉就齐了。"""
    try:
        名列表 = 图标.availableSizes()
        if not 名列表:
            return 图标
        # 显示只有 边x边（36），取"够用的最大档"就够了：以前直接拿 256x256 那张，
        # 中间位图和 逐像素扫留白（6.5 万点）都白贵一个量级。没有 ≤64 的档位时退回最大档，
        # 保证任何图标都取得到图。
        够用 = [s for s in 名列表 if s.width() <= 64 and s.height() <= 64]
        if 够用:
            最大 = max(够用, key=lambda s: s.width() * s.height())
        else:
            最大 = max(名列表, key=lambda s: s.width() * s.height())
        pix = 图标.pixmap(最大)
        if pix.isNull():
            return 图标
        if not pix.hasAlphaChannel():
            return 图标  # 无透明通道说明是满幅位图，没有留白可裁
        图 = pix.toImage().convertToFormat(QImage.Format_ARGB32)
        区 = _不透明包围盒(图)
        if 区.isEmpty() or 区 == 图.rect():
            return 图标  # 全透明或本就无留白，原样返回
        裁 = 图.copy(区).scaled(边, 边, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        # 画布用 QImage + drawImage：QPainter.drawPixmap 直接收 QImage 会原生崩溃
        画布 = QImage(边, 边, QImage.Format_ARGB32)
        画布.fill(Qt.transparent)
        笔 = QPainter(画布)
        笔.drawImage((边 - 裁.width()) // 2, (边 - 裁.height()) // 2, 裁)
        笔.end()
        return QIcon(QPixmap.fromImage(画布))
    except Exception:
        return 图标


def 取文件图标(路径):
    try:
        图 = 图标提供者.icon(QFileInfo(路径))
        if not 图.isNull():
            return _规格化(图)
    except Exception:
        pass
    return QApplication.style().standardIcon(QStyle.SP_FileIcon)


图片后缀集 = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".ico", ".tiff"}


_缩略图解码上限 = 192  # 只要 48 像素的显示效果，192 是 4 倍超采样，肉眼与全图解码无差别


def _小尺寸解码(路径):
    """按显示需要的小尺寸解码图片，别把整张原图搬进内存。

    原先 QPixmap(路径) 会把 1024x768 这类原图整个解出来（一张约 3MB）只为缩成 48 像素；
    这里先问解码器要 192 以内的尺寸 —— JPEG 会走它内置的 1/2、1/4、1/8 降采样，
    PNG 至少能少分配一大块位图。尺寸不到上限的原图则按原样解码，行为不变。
    """
    读 = QImageReader(路径)
    尺寸 = 读.size()
    try:
        if 尺寸.isValid() and max(尺寸.width(), 尺寸.height()) > _缩略图解码上限:
            比 = _缩略图解码上限 / float(max(尺寸.width(), 尺寸.height()))
            读.setScaledSize(QSize(max(1, int(尺寸.width() * 比)),
                                   max(1, int(尺寸.height() * 比))))
    except Exception:
        pass  # 万一 setScaledSize 不配合，退回原尺寸解码，绝不能因此取不到图
    return 读.read()


_图标缓存 = {}        # (路径, mtime, 大小) -> 算好的 QIcon
_图标缓存上限 = 1200  # 超过就整张清空重来，防止用久了无限涨


def _缓存键(路径):
    """路径 + 文件指纹当键：文件被换掉、改名、移走都会自动当没缓存，不会显示旧图标"""
    try:
        信 = os.stat(路径)
        return (路径, int(信.st_mtime), 信.st_size)
    except OSError:
        return None


def 取图标展示(路径):
    """图片文件优先显示真实缩略图（等比缩到 48px），其余走系统图标。
    只对盒子里存放的"原图文件"生效；.lnk 快捷方式不解析目标（省 COM 开销）。
    同一路径算过一次就进缓存：搜索/换列/拖入/改名引起的 重建网格 不再重复解码、
    不再重复向系统要图标、不再重复逐像素扫留白。"""
    键 = _缓存键(路径)
    if 键 is not None:
        存 = _图标缓存.get(键)
        if 存 is not None:
            return 存  # 命中：连系统图标那一趟都不跑了
    图 = None
    try:
        后缀 = os.path.splitext(路径)[1].lower()
        if 后缀 in 图片后缀集 and os.path.isfile(路径):
            小 = _小尺寸解码(路径)
            if not 小.isNull():
                缩 = 小.scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                图 = QIcon(QPixmap.fromImage(缩))  # 本机 PyQt5 的 QIcon 不吃 QImage，必须先转 QPixmap
    except Exception:
        图 = None
    if 图 is None:
        图 = 取文件图标(路径)
    if 键 is not None and 图 is not None:
        if len(_图标缓存) >= _图标缓存上限:
            _图标缓存.clear()  # 上限保护：整张清空，下一轮自然重算
        _图标缓存[键] = 图
    return 图


# ---------------- 盒内图标按钮 ----------------

class 盒内图标(QToolButton):
    def __init__(self, 主窗, 卡片, 条目, 所属分类=None):
        super().__init__()
        self.主窗 = 主窗
        self.卡片 = 卡片
        self.条目 = 条目
        # "全部"总览页里卡片的分类是虚拟分类，打开/移动/恢复/删除
        # 都必须用条目真正所属的分类，所以每个图标自己记一份
        self.所属 = 所属分类 or 卡片.分类
        self.setObjectName("盒内图标")
        self.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.setIconSize(QSize(36, 36))
        self.setFixedSize(QSize(图标按钮宽, 图标按钮高))
        名称 = 条目["名称"]
        self.setText(名称 if len(名称) <= 8 else 名称[:7] + "…")
        self.setToolTip(名称)
        self._详情已建 = False  # 悬停详情卡片懒生成：首次悬停才统计文件信息
        self.setIcon(取图标展示(self.绝对路径()))
        self.setCursor(Qt.PointingHandCursor)
        self.setAcceptDrops(True)  # 图标自身也接收拖放：拖到别的图标上 = 排序
        self._按下位置 = None
        self._单击音定时 = None  # 延后待放的单击音定时器，确认是双击/拖拽就撤

    def 绝对路径(self):
        return os.path.join(数据目录, self.所属["名称"], self.条目["文件"])

    def _open_item(self):
        """打开条目（双击或菜单），顺手记一次打开数（常用置顶用）"""
        try:
            os.startfile(self.绝对路径())
            self.条目["打开次数"] = int(self.条目.get("打开次数", 0)) + 1
            保存配置()
        except Exception:
            pass
        self.主窗.打开后收起()

    def enterEvent(self, 事件):
        """鼠标悬停到图标上：播提示音；首次悬停把 tooltip 升级成文件详情卡片"""
        播放悬停音效()
        if not self._详情已建:
            self._详情已建 = True
            try:
                self.setToolTip(self._构造详情())
            except Exception:
                pass
        super().enterEvent(事件)

    def _构造详情(self):
        """悬停详情：名称 / 大小 / 修改时间 / 快捷方式目标 / 原位置。
        只在第一次悬停时算一次，不拖慢建网格。"""
        行 = ["<b>%s</b>" % self.条目.get("名称", "")]
        p = self.绝对路径()
        if os.path.exists(p):
            st = os.stat(p)
            大小 = st.st_size
            if 大小 > 1048576:
                大小文 = "%.1f MB" % (大小 / 1048576)
            elif 大小 > 1024:
                大小文 = "%.1f KB" % (大小 / 1024)
            else:
                大小文 = "%d B" % 大小
            行.append("大小：%s" % 大小文)
            行.append("修改时间：" + time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime)))
            if p.lower().endswith(".lnk"):
                sc = 取shell().CreateShortcut(p)
                if sc.TargetPath:
                    行.append("目标：" + sc.TargetPath)
        if self.条目.get("原路径"):
            行.append("原位置：" + self.条目["原路径"])
        行.append("打开 %d 次" % int(self.条目.get("打开次数", 0)))
        return "<br>".join(行)

    def _排单击音(self):
        """把单击音延后 0.16 秒再放：这段时间里确认是双击就撤掉。
        注意只延后了声音，点击动作本身一点没变慢。"""
        self._取消单击音()
        定时 = QTimer(self)
        定时.setSingleShot(True)
        定时.timeout.connect(播放单击音效)
        self._单击音定时 = 定时
        定时.start(160)

    def _取消单击音(self):
        """确认这一下是双击或者已经拖起来了：撤掉还没响的那声哒"""
        定时 = self._单击音定时
        self._单击音定时 = None
        if 定时 is not None:
            定时.stop()

    def mouseDoubleClickEvent(self, 事件):
        self._取消单击音()  # 双击只响一声：撤掉第一次按下延后待放的哒
        播放双击音效()
        self._open_item()

    def mousePressEvent(self, 事件):
        if 事件.button() == Qt.LeftButton:
            self._按下位置 = 事件.pos()
            # 注意：本机 PyQt5 5.15.11 的 QMouseEvent 没有 clickCount()，一问就抛异常，
            # 所以不问第几次按下 —— 一律延后 0.16 秒出声。真双击时第二次按下会把这声哒
            # 重新排上，紧随其后的 mouseDoubleClickEvent 再撤掉，照样只响一声。
            self._排单击音()
        super().mousePressEvent(事件)

    def mouseMoveEvent(self, 事件):
        if self._按下位置 is None or not (事件.buttons() & Qt.LeftButton):
            return
        if (事件.pos() - self._按下位置).manhattanLength() < 12:
            return
        self._按下位置 = None
        self._取消单击音()  # 是拖出去排序，不是点一下，撤掉待放的哒
        拖 = QDrag(self)
        数据 = QMimeData()
        数据.setData(内部拖放格式, json.dumps(
            {"源": self.所属["名称"], "文件": self.条目["文件"]},
            ensure_ascii=False).encode("utf-8"))
        拖.setMimeData(数据)
        拖.setPixmap(self.icon().pixmap(32, 32))
        拖.exec_(Qt.MoveAction)

    def dragEnterEvent(self, 事件):
        """拖进来的若是盒内图标，本图标就成为一个合法的落点"""
        if 事件.mimeData().hasFormat(内部拖放格式):
            事件.acceptProposedAction()

    def dropEvent(self, 事件):
        """落到本图标上：插到本图标前面（子级接收后 Qt 不会再传给页面）"""
        数据 = 事件.mimeData()
        if not 数据.hasFormat(内部拖放格式):
            return
        try:
            载荷 = json.loads(bytes(数据.data(内部拖放格式)).decode("utf-8"))
        except Exception:
            return
        事件.acceptProposedAction()
        self.主窗.移动条目(载荷.get("源", ""), 载荷.get("文件", ""),
                          self.所属, 目标文件名=self.条目["文件"])

    def contextMenuEvent(self, 事件):
        """图标右键菜单：常用 Windows 文件操作都在这里"""
        菜单 = QMenu(self)
        菜单.addAction("打开", self._open_item)
        菜单.addSeparator()
        菜单.addAction("📂 打开所在位置", self._open_folder)   # 槽名 ASCII！
        菜单.addAction("📋 复制文件路径", self._copy_path)      # 槽名 ASCII！
        菜单.addAction("🗂 查看文件属性", self._show_properties)  # 槽名 ASCII！
        菜单.addSeparator()
        if self.条目.get("原路径"):
            菜单.addAction("↩ 恢复到原来的位置",
                           lambda: self.主窗.恢复条目(self.所属, self.条目))
        else:
            菜单.addAction("🗑 移除快捷方式",
                           lambda: self.主窗.恢复条目(self.所属, self.条目))
        菜单.addAction("🗑 删除图标（进回收站）", self._delete_to_trash)  # 槽名 ASCII！
        菜单.exec_(事件.globalPos())

    def _open_folder(self):
        """用资源管理器打开所在文件夹并选中文件（Windows Shell 特性）"""
        import subprocess
        p = self.绝对路径()
        if not os.path.exists(p):
            return
        try:
            # /select 让资源管理器打开文件夹并选中该文件
            subprocess.Popen(["explorer", "/select,", p])
        except Exception:
            os.startfile(os.path.dirname(p))

    def _delete_to_trash(self):
        """删除图标 = 把盒子里这个文件送进回收站（后悔可去回收站找回），再移除条目。
        走系统正规删除 SHFileOperationW + FOF_ALLOWUNDO，不是永久抹掉。"""
        答 = QMessageBox.question(
            self, "删除图标",
            "要删除「%s」吗？\n文件会进回收站，后悔了可以去回收站找回。" % self.条目.get("名称", ""))
        if 答 != QMessageBox.Yes:
            return
        p = self.绝对路径()
        try:
            if os.path.exists(p):
                import ctypes.wintypes as wt

                class _文件操作结构(ctypes.Structure):
                    _fields_ = [("hwnd", wt.HWND), ("wFunc", wt.UINT),
                                ("pFrom", ctypes.c_wchar_p), ("pTo", ctypes.c_wchar_p),
                                ("fFlags", ctypes.c_uint16),
                                ("fAnyOperationsAborted", wt.BOOL),
                                ("hNameMappings", ctypes.c_void_p),
                                ("lpszProgressTitle", ctypes.c_wchar_p)]

                op = _文件操作结构()
                op.hwnd = None
                op.wFunc = 3                      # FO_DELETE
                op.pFrom = p + "\0\0"           # 系统要求路径以双空字符结尾
                op.pTo = None
                op.fFlags = 0x40 | 0x10 | 0x4     # 进回收站 + 不再系统弹窗 + 静默
                ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
            分类 = self.所属
            if self.条目 in 分类["条目"]:
                分类["条目"].remove(self.条目)
            保存配置()
            self.主窗.刷新全部()
        except Exception as e:
            QMessageBox.warning(self, "删除失败", str(e))

    def _copy_path(self):
        """把文件路径复制到剪贴板"""
        真实路径 = self.条目.get("原路径") or self.绝对路径()
        QApplication.clipboard().setText(真实路径)

    def _show_properties(self):
        """显示文件/快捷方式的详细信息（名称、路径、大小、时间、快捷方式目标）。
        不用 ShellExecuteW verb='properties'——那个在 hwnd=0 时返回 31 静默失败。
        这里直接读文件系统元数据，弹自己的对话框，100% 可靠。"""
        import stat as _stat
        from PyQt5.QtWidgets import QDialog, QLabel, QVBoxLayout, QPushButton, QHBoxLayout
        from PyQt5.QtCore import Qt

        # 优先展示原路径（如果有），否则展示快捷方式路径
        原路径 = self.条目.get("原路径")
        盒子里的文件 = self.绝对路径()
        if 原路径 and os.path.exists(原路径):
            展示路径 = 原路径
        elif os.path.exists(盒子里的文件):
            展示路径 = 盒子里的文件
        else:
            return  # 两个都不存在就不弹

        dlg = QDialog(None)
        dlg.setWindowFlags(Qt.Dialog | Qt.Window)
        dlg.setWindowTitle("文件属性")
        dlg.setMinimumWidth(420)
        行 = QVBoxLayout(dlg)

        # 用分隔线把各信息行隔开
        def 加行(标题, 值):
            l = QLabel(f"<b>{标题}：</b> {值}")
            l.setStyleSheet("background: transparent;")
            l.setWordWrap(True)
            l.setTextInteractionFlags(Qt.TextSelectableByMouse)  # 文字可选复制
            行.addWidget(l)

        加行("图标名称", self.条目.get("名称", ""))
        行.addWidget(QLabel(""))  # 空行分隔

        # 类型判断：看盒子里存的真实文件后缀（因为原文件被搬进来了）
        盒子后缀 = os.path.splitext(self.条目.get("文件", ""))[1].lower()
        if 盒子后缀 == ".lnk":
            加行("类型", "桌面收纳盒创建的快捷方式 (.lnk)")
        else:
            加行("类型", f"原始文件（原扩展名：{盒子后缀 or '无'}）")

        加行("盒子里的位置", f"<code>{盒子里的文件}</code>")
        if 原路径:
            加行("原始位置（可恢复到这里）", f"<code>{原路径}</code>")

        if os.path.exists(展示路径):
            st = os.stat(展示路径)
            大小 = st.st_size
            if 大小 > 1024 * 1024:
                大小显示 = f"{大小 / 1024 / 1024:.1f} MB ({大小} 字节)"
            elif 大小 > 1024:
                大小显示 = f"{大小 / 1024:.1f} KB ({大小} 字节)"
            else:
                大小显示 = f"{大小} 字节"
            加行("文件大小", 大小显示)
            加行("创建时间", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_ctime)))
            加行("修改时间", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_mtime)))

        # 快捷方式 (.lnk) 的话额外解析一下目标路径
        if 展示路径.lower().endswith(".lnk") and os.path.exists(展示路径):
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                sc = shell.CreateShortcut(展示路径)
                if sc.TargetPath:
                    加行("快捷方式目标", f"<code>{sc.TargetPath}</code>")
                if sc.Arguments:
                    加行("启动参数", f"<code>{sc.Arguments}</code>")
                if sc.WorkingDirectory:
                    加行("工作目录", f"<code>{sc.WorkingDirectory}</code>")
            except Exception:
                pass

        行.addStretch()
        按钮行 = QHBoxLayout()
        按钮行.addStretch()
        关闭 = QPushButton("关闭")
        关闭.clicked.connect(dlg.accept)
        按钮行.addWidget(关闭)
        行.addLayout(按钮行)

        dlg.exec_()


# ---------------- 分类页（简化版：只放图标网格，标题和管理菜单移到顶部标签栏） ----------------

class 分类页(QWidget):
    """一个分类对应的页面，只负责展示该分类下的图标网格 + 接收拖放。
    分类名称显示在主窗口的 QTabBar 标签上，分类管理（重命名/删除）改由标签栏右键处理。
    特点：图标从左上角自动排列；可画格子背景让每个图标格子位置清楚；图标间距可设置。"""

    def __init__(self, 主窗, 分类):
        super().__init__()
        self.主窗 = 主窗
        self.分类 = 分类
        self.虚拟 = bool(分类.get("虚拟"))  # "全部"总览页：聚合展示，不存文件
        self.setAcceptDrops(True)

        # 外层布局：把网格顶在上面（图标从左上角排，排完后下面空着）
        布局 = QVBoxLayout(self)
        布局.setContentsMargins(0, 0, 0, 0)
        布局.setSpacing(0)

        # 图标网格：所有盒内图标按钮按行列排列，窗口变宽时自动多列
        self.网格 = QGridLayout()
        self.应用间距()  # 根据配置设置 spacing 和 margin
        布局.addLayout(self.网格)
        布局.addStretch(1)  # 关键！让网格顶在最上面，下面多余空间留空

        self._上次列数 = None  # 记住上次排列用的列数，宽窄没变就不重排（防闪烁）
        self.重建网格()

    def 应用间距(self):
        """根据配置里的图标间距设置网格的外边距和图标间距。
        外边距 = 图标间距 + 4（页面内容区还有一点呼吸感），
        图标之间的 spacing 直接用图标间距。"""
        s = max(0, int(配置.get("图标间距", 默认图标间距)))
        self.网格.setContentsMargins(s + 4, s + 4, s + 4, s + 4)
        self.网格.setSpacing(s)

    def _当前列数(self):
        """根据页面实际宽度算一行能放几个图标：
        每个图标占 (图标宽 + 间距) 像素，减去总边距后除以单格宽度。"""
        s = max(0, int(配置.get("图标间距", 默认图标间距)))
        可用宽 = self.width() - 2 * (s + 4)
        单格宽 = 图标按钮宽 + s
        列数 = max(3, 可用宽 // 单格宽)
        return int(列数)

    def resizeEvent(self, 事件):
        """窗口被拖宽/拖窄时触发：列数变了就重新排列图标"""
        super().resizeEvent(事件)
        列数 = self._当前列数()
        if 列数 != self._上次列数:
            self.重建网格()
        self.update()  # 重绘格子背景（让格子线始终填满新尺寸）

    def 重建网格(self):
        """清空网格里所有图标按钮，然后根据当前条目和列数重新放进去。
        关键：设 setColumnStretch(0) 让列不被拉伸，setAlignment 让网格顶左上角，
        这样每个格子的实际大小就和我们手动算的一致，画格子线能对上。
        重要修复：清空时必须先 w.setParent(None) 再 deleteLater()——
        deleteLater 只是标记"稍后删"，widget 还在 Qt 对象树里，
        如果只 deleteLater 就重建，旧按钮还占着屏幕位置，新按钮叠上去就重叠了！
        setParent(None) 立刻从父 widget 树里摘掉（不再渲染、不再被 findChildren 找到），
        然后 deleteLater 再让 Qt 事件循环最终释放内存。"""
        self.应用间距()  # 每次重建都同步最新间距
        列数 = self._当前列数()
        self._上次列数 = 列数
        # 清空旧内容（同时清掉旧的 stretch 设置）
        while self.网格.count():
            项 = self.网格.takeAt(0)
            w = 项.widget()
            if w:
                w.setParent(None)  # ← 关键！立刻从父窗口摘掉，防止新按钮和旧按钮重叠
                w.deleteLater()
        # 每列都不拉伸（stretch=0），用固定宽度
        for col in range(列数):
            self.网格.setColumnStretch(col, 0)
        # 整个网格对齐到左上角，别分散填满
        self.网格.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        # 每项是 (条目, 真正所属分类) 二元组；"全部"总览页聚合所有真实分类
        if self.虚拟:
            条目列表 = [(t, c) for c in 配置["分类"] if not c.get("虚拟")
                        for t in c["条目"]]
        else:
            条目列表 = [(t, self.分类) for t in self.分类["条目"]]
        # 常用置顶：按累计打开次数从多到少排（sort 稳定，次数相同保持原顺序）
        if 配置.get("常用置顶", False):
            条目列表.sort(key=lambda p: -int(p[0].get("打开次数", 0)))
        # 搜索过滤：只显示名称里含关键字的图标
        过滤字 = getattr(self.主窗, "搜索字", "").strip().lower()
        if 过滤字:
            条目列表 = [p for p in 条目列表
                        if 过滤字 in str(p[0].get("名称", "")).lower()]
        if 条目列表:
            for 序, (条目, 所属) in enumerate(条目列表):
                self.网格.addWidget(盒内图标(self.主窗, self, 条目, 所属),
                                    序 // 列数, 序 % 列数)
        else:
            提示 = QLabel("把桌面图标拖到这里" if not 过滤字 else "没有匹配的图标")
            提示.setObjectName("提示文字")
            提示.setAlignment(Qt.AlignCenter)
            提示.setStyleSheet("color: #6a7386; font-size: 9pt;")
            # 空提示横跨整行（占满所有列）
            self.网格.addWidget(提示, 0, 0, 1, 列数)
        _静音悬停音()  # 整片换新控件 = 幻影 enter 的来源，一并掐掉
        self.update()  # 重建完图标也重绘格子线

    def paintEvent(self, 事件):
        """在页面上画浅色格子背景：让每个图标的格子位置一目了然。
        因为 QGridLayout 已经不拉伸了（setColumnStretch=0），图标按钮的位置可以精确计算：
        widget 中心 = 边距 + i*(按钮宽+间距) + 按钮宽/2"""
        super().paintEvent(事件)
        绘制 = QPainter(self)
        绘制.setRenderHint(QPainter.Antialiasing)
        颜色 = QColor(255, 255, 255, 22)  # 很淡的白色，不抢图标风头
        绘制.setPen(颜色)

        s = max(0, int(配置.get("图标间距", 默认图标间距)))
        边距 = s + 4
        列数 = self._当前列数()
        格子步宽 = 图标按钮宽 + s   # 每列之间的中心距离
        格子步高 = 图标按钮高 + s

        if self.虚拟:
            条目数 = sum(len(c["条目"]) for c in 配置["分类"] if not c.get("虚拟"))
        else:
            条目数 = len(self.分类["条目"])
        总行数 = max(3, (条目数 + 列数 - 1) // 列数)

        # 画竖线：每条线对齐一个图标按钮的中心
        for i in range(列数):
            中心x = 边距 + i * 格子步宽 + 图标按钮宽 // 2
            if 中心x > self.width():
                break
            绘制.drawLine(int(中心x), int(边距),
                          int(中心x), int(边距 + 总行数 * 格子步高))
        # 画横线：每条线对齐一个图标按钮的中心
        for j in range(总行数):
            中心y = 边距 + j * 格子步高 + 图标按钮高 // 2
            if 中心y > self.height():
                break
            绘制.drawLine(int(边距), int(中心y),
                          int(边距 + 列数 * 格子步宽), int(中心y))

    # --- 拖放：既接收外部文件，也接收盒内图标移动 ---

    def dragEnterEvent(self, 事件):
        数据 = 事件.mimeData()
        if 数据.hasFormat(内部拖放格式):
            事件.acceptProposedAction()
        elif 数据.hasUrls() and any(u.isLocalFile() for u in 数据.urls()):
            事件.acceptProposedAction()

    def dragMoveEvent(self, 事件):
        事件.acceptProposedAction()

    def dropEvent(self, 事件):
        数据 = 事件.mimeData()
        if 数据.hasFormat(内部拖放格式):
            try:
                载荷 = json.loads(bytes(数据.data(内部拖放格式)).decode("utf-8"))
            except Exception:
                return
            事件.acceptProposedAction()
            目标 = self.分类
            if self.虚拟:
                # "全部"是总览不是真分类：落到它空白处 = 挪到该条目自己分类的末尾
                目标 = next((c for c in 配置["分类"]
                             if c["名称"] == 载荷.get("源", "")), None)
                if 目标 is None:
                    return
            self.主窗.移动条目(载荷.get("源", ""), 载荷.get("文件", ""), 目标)
        elif 数据.hasUrls():
            路径 = [u.toLocalFile() for u in 数据.urls() if u.isLocalFile()]
            if 路径:
                事件.acceptProposedAction()
                self.主窗.拖入(self.分类, 路径)


def 前台是全屏():
    """前台窗口完整盖住了它所在的整个显示器 → 认为是游戏独占全屏/全屏视频/F11 全屏。
    普通"最大化窗口"不算（它让出了任务栏，面积小于整屏），不影响日常边缘唤出。"""
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return False
        屏矩形 = win32api.GetMonitorInfo(
            win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST))[0]
        窗矩形 = win32gui.GetWindowRect(hwnd)
        return ((窗矩形[2] - 窗矩形[0]) >= (屏矩形[2] - 屏矩形[0]) and
                (窗矩形[3] - 窗矩形[1]) >= (屏矩形[3] - 屏矩形[1]))
    except Exception:
        return False


# ---------------- 主窗口 ----------------

class 收纳盒窗口(QWidget):
    # 信号名必须是 ASCII（Qt 元对象系统限制），中文类名/方法名不受影响
    show_request = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._固定 = False
        # 鼠标是否进入过窗口：刚从托盘点开时鼠标还停在托盘上，
        # 必须"进入过窗口之后再离开"才自动隐藏，否则刚点开就会被误判为"鼠标移开"而闪没
        self._鼠标进过窗口 = False
        self._拖动偏移 = None
        self.setWindowTitle("桌面收纳盒")
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        # 不再写死窗口宽度：限制最小尺寸，初始尺寸用配置里记住的大小
        self.setMinimumSize(最小宽, 最小高)
        self.resize(配置["宽"], 配置["高"])

        根 = QFrame(self)
        根.setObjectName("根窗口")
        根.installEventFilter(self)  # 整个根窗口都能拖（eventFilter 里会避开按钮等交互控件）
        外 = QVBoxLayout(self)
        外.setContentsMargins(0, 0, 0, 0)
        外.addWidget(根)

        内 = QVBoxLayout(根)
        内.setContentsMargins(14, 12, 14, 12)
        内.setSpacing(8)

        # 标题栏
        头 = QHBoxLayout()
        标题 = QLabel("桌面收纳盒")
        标题.setObjectName("标题文字")
        标题.installEventFilter(self)  # 标题栏拖动窗口
        self.固定钮 = QToolButton()
        self.固定钮.setObjectName("固定钮")
        self.固定钮.setText("📌 固定")
        self.固定钮.setCheckable(True)
        self.固定钮.setCursor(Qt.PointingHandCursor)
        self.固定钮.setToolTip("固定后窗口不再自动隐藏，方便往里拖图标")
        self.固定钮.toggled.connect(self._on_pin_toggled)
        # 设置按钮：固定和关闭之间，点了弹设置对话框
        self.设置钮 = QToolButton()
        self.设置钮.setObjectName("固定钮")  # 复用现有按钮样式
        self.设置钮.setText("⚙️")
        self.设置钮.setCursor(Qt.PointingHandCursor)
        self.设置钮.setToolTip("设置：窗口大小 / 图标间距")
        self.设置钮.clicked.connect(self._show_settings)  # 槽名必须 ASCII！

        # 赞助按钮：低调灰色小图标，悬停变彩色
        self.赞助钮 = QToolButton()
        self.赞助钮.setObjectName("赞助钮")
        self.赞助钮.setText("☕")
        self.赞助钮.setCursor(Qt.PointingHandCursor)
        self.赞助钮.setToolTip("如果本工具帮到了您，可以请作者喝杯咖啡 ❤️")
        self.赞助钮.clicked.connect(self._show_support)  # 槽名 ASCII！

        关闭钮 = QToolButton()
        关闭钮.setObjectName("关闭钮")
        关闭钮.setText("✕")
        关闭钮.setCursor(Qt.PointingHandCursor)
        关闭钮.setToolTip("隐藏（托盘里随时可以唤出）")
        关闭钮.clicked.connect(self.hide)
        头.addWidget(标题)
        头.addStretch()
        头.addWidget(self.固定钮)
        头.addWidget(self.设置钮)
        头.addWidget(self.赞助钮)  # 赞助（低调灰色，悬停变彩）
        头.addWidget(关闭钮)
        内.addLayout(头)

        提示 = QLabel("把桌面图标拖进分类 · 双击图标打开 · 右键图标可恢复/移除")
        提示.setObjectName("提示文字")
        提示.setWordWrap(True)
        内.addWidget(提示)

        # 搜索过滤框：打字实时过滤当前所有分类页的图标，Esc 或 ✕ 清空
        self.搜索字 = ""
        self.搜索框 = QLineEdit()
        self.搜索框.setObjectName("搜索框")
        self.搜索框.setPlaceholderText("搜索图标…")
        self.搜索框.setClearButtonEnabled(True)
        self.搜索框.textChanged.connect(self._on_search)  # 槽名 ASCII！
        内.addWidget(self.搜索框)

        # 分类导航分两区：左边 = 类型分类按钮列（全部/文档/图片…），
        # 右上 = 用户自己的分类标签栏（常用/魔兽争霸/新建的），右下 = 图标页
        主体行 = QHBoxLayout()
        主体行.setContentsMargins(0, 2, 0, 0)
        主体行.setSpacing(6)

        # 左列：QToolButton 横排文字从上堆到下（竖排 QTabBar 会叠字，已弃用）
        self.侧栏容器 = QWidget()
        self.侧栏容器.setObjectName("类型侧栏")
        侧栏列 = QVBoxLayout(self.侧栏容器)
        侧栏列.setContentsMargins(0, 0, 0, 0)
        侧栏列.setSpacing(4)
        侧栏列.addStretch(1)  # 末尾弹性留白：按钮始终从顶部开始堆
        self._侧栏按钮 = []   # 刷新全部 时重建
        主体行.addWidget(self.侧栏容器)

        右列 = QVBoxLayout()
        右列.setSpacing(4)

        标签栏行 = QHBoxLayout()
        标签栏行.setContentsMargins(0, 0, 0, 0)
        标签栏行.setSpacing(4)

        self.标签栏 = QTabBar()
        self.标签栏.setObjectName("分类标签栏")
        self.标签栏.setDocumentMode(True)      # 让标签和下方内容无缝衔接
        self.标签栏.setUsesScrollButtons(True)  # 标签多了自动出现左右箭头
        self.标签栏.setTabsClosable(False)      # 关掉默认的"×"关闭按钮（用右键菜单删分类）
        self.标签栏.setMovable(False)           # 不让用户拖动标签改顺序（简化逻辑）
        self.标签栏.currentChanged.connect(self._on_tab_changed)  # 切换标签时更新页面
        # 标签上右键 → 重命名/删除分类；空白处右键 → 新建分类
        self.标签栏.customContextMenuRequested.connect(self._on_tab_menu)
        self.标签栏.setContextMenuPolicy(Qt.CustomContextMenu)
        标签栏行.addWidget(self.标签栏, 1)

        # "+"新建分类按钮，固定在标签栏最右边（新建的分类都进顶部标签栏）
        self.新建钮 = QToolButton()
        self.新建钮.setObjectName("新建分类")
        self.新建钮.setText("＋")
        self.新建钮.setCursor(Qt.PointingHandCursor)
        self.新建钮.setToolTip("新建分类")
        self.新建钮.clicked.connect(self._add_category)  # 槽名必须 ASCII！
        标签栏行.addWidget(self.新建钮)
        右列.addLayout(标签栏行)

        # QStackedWidget：每次只显示一个分类的图标页面，两套导航都能切换它
        self.页面栈 = QStackedWidget()
        self.页面栈.setObjectName("分类页面栈")
        右列.addWidget(self.页面栈, 1)

        主体行.addLayout(右列)
        内.addLayout(主体行)

        # 自动隐藏：定时检查鼠标是否离开窗口
        self._隐藏计时 = QTimer(self)
        self._隐藏计时.setInterval(100)  # 100毫秒轮询：边缘唤出约0.3秒触发
        self._隐藏计时.timeout.connect(self._check_autohide)
        self._隐藏计时.start()

        # 屏幕边缘唤出用：连续贴边计数（约 0.3 秒触发），藏在隐藏定时器里顺带检测
        self._贴边次数 = 0
        # 用户把窗口拖到屏幕哪条边贴住（上/左/右），边缘唤出就只认那几条边
        self._停靠边 = set()
        # 双区分类导航：页面栈索引 = 配置["分类"] 索引，两张映射表换算导航内部序号
        self._导航同步 = False
        self._标签映射 = []
        self._侧栏映射 = []

        self.show_request.connect(self.toggle)  # 鼠标钩子线程 → 主线程
        self.刷新全部()
        if 配置["固定"]:
            self.固定钮.setChecked(True)

    # --- 刷新与自动隐藏 ---

    def 刷新全部(self):
        """重建两套分类导航 + 页面：左边按钮列 = 类型分类（全部/文档/图片…），
        顶部标签栏 = 用户自己的分类（常用/魔兽争霸/新建的）。
        页面栈顺序永远与 配置["分类"] 一致，两套导航用映射表换算索引。"""
        原来的索引 = self.页面栈.currentIndex()
        # 清空页面栈（旧页面会被 deleteLater 释放）
        while self.页面栈.count():
            w = self.页面栈.widget(0)
            self.页面栈.removeWidget(w)
            if w:
                w.setParent(None)  # 同重建网格：立刻摘掉，防止旧页面残留
                w.deleteLater()
        # 清空顶部标签栏和左栏按钮
        while self.标签栏.count():
            self.标签栏.removeTab(0)
        for 钮 in self._侧栏按钮:
            钮.setParent(None)
            钮.deleteLater()
        self._侧栏按钮 = []
        self._标签映射 = []  # 顶部标签序号 → 分类索引
        self._侧栏映射 = []  # 左栏按钮序号 → 分类索引
        for 索引, 分类 in enumerate(配置["分类"]):
            页面 = 分类页(self, 分类)
            self.页面栈.addWidget(页面)
            # 计数："全部"总览是所有真实分类加起来
            if 分类.get("虚拟"):
                数目 = sum(len(c["条目"]) for c in 配置["分类"] if not c.get("虚拟"))
            else:
                数目 = len(分类["条目"])
            if 分类.get("类型"):
                # 左栏按钮：横排文字，点它切换分类；右键可改名/删除（"全部"除外）
                钮 = QToolButton()
                钮.setObjectName("类型侧栏钮")
                钮.setText(f"{分类['名称']} ({数目})")
                钮.setCheckable(True)
                钮.setCursor(Qt.PointingHandCursor)
                钮.clicked.connect(lambda _=False, i=索引: self._跳转分类(i))  # 槽名 ASCII！
                if 分类.get("虚拟"):
                    钮.setToolTip("所有分类的总览，拖文件到这里会自动按类型归类")
                else:
                    钮.setToolTip("右键可重命名或删除")
                    钮.setContextMenuPolicy(Qt.CustomContextMenu)
                    钮.customContextMenuRequested.connect(
                        lambda _p, i=索引: self._on_side_menu(i))
                # 插到末尾 stretch 之前：按钮从顶部开始一个个堆
                self.侧栏容器.layout().insertWidget(len(self._侧栏按钮), 钮)
                self._侧栏按钮.append(钮)
                self._侧栏映射.append(索引)
            else:
                self._标签映射.append(索引)
                self.标签栏.addTab(f"{分类['名称']}  ({数目})")
        # 尽量停在原来正看着的分类上（原索引可能被删了，就切到相邻的）
        if self.页面栈.count():
            self._跳转分类(max(0, min(原来的索引, self.页面栈.count() - 1)))

    def _on_tab_changed(self, 标签序):
        """顶部标签栏切换：标签序 是"栏内序号"，查映射表换成真正的分类索引"""
        if self._导航同步 or 标签序 < 0 or 标签序 >= len(self._标签映射):
            return
        self._跳转分类(self._标签映射[标签序])

    def _跳转分类(self, 索引):
        """统一的切换分类入口：翻页面 + 同步两套导航的高亮（_导航同步 防信号回环）"""
        if 索引 < 0 or 索引 >= self.页面栈.count():
            return
        self._导航同步 = True
        try:
            self.页面栈.setCurrentIndex(索引)
            标签序 = self._标签映射.index(索引) if 索引 in self._标签映射 else -1
            if self.标签栏.currentIndex() != 标签序:
                self.标签栏.setCurrentIndex(标签序)
            for 序, 钮 in enumerate(self._侧栏按钮):
                钮.setChecked(索引 == self._侧栏映射[序])
        finally:
            self._导航同步 = False

    def _on_side_menu(self, 索引):
        """左边类型分类按钮的右键菜单：重命名 / 删除（和顶部标签一样）"""
        菜单 = QMenu(self)
        菜单.addAction("重命名分类…", lambda: self._重命名分类(索引))
        菜单.addAction("删除分类…", lambda: self._删除分类(索引))
        菜单.exec_(QCursor.pos())

    def _on_tab_menu(self, 位置):
        """标签栏右键菜单：
        - 点在某个标签上 → 重命名 / 删除 那个分类
        - 点在标签栏空白处（没点到标签）→ 新建分类"""
        全局位置 = self.标签栏.mapToGlobal(位置)
        # tabAt 返回 -1 表示点在空白处，否则是点中的标签索引
        点中的标签索引 = self.标签栏.tabAt(位置)
        菜单 = QMenu(self)
        if 0 <= 点中的标签索引 < len(self._标签映射):
            # 顶部标签只装用户自己的分类：栏内序号先换算成真正的分类索引
            # （"全部"等类型分类在左栏，右键保护在它们自己的按钮上）
            分类索引 = self._标签映射[点中的标签索引]
            菜单.addAction("重命名分类…", lambda: self._重命名分类(分类索引))
            菜单.addAction("删除分类…", lambda: self._删除分类(分类索引))
            菜单.addSeparator()
        # 无论点在标签还是空白处，都可以新建
        菜单.addAction("新建分类…", self._add_category)  # 槽名必须 ASCII！
        菜单.exec_(全局位置)

    def _重命名分类(self, 索引):
        """重命名指定索引的分类：检查重名、改文件夹名、刷新"""
        分类 = 配置["分类"][索引]
        旧名 = 分类["名称"]
        文本, 确定 = QInputDialog.getText(self, "重命名分类", "新的分类名称：", text=旧名)
        if not 确定:
            return
        新名 = 清理名称(文本, 旧名)
        if 新名 == 旧名:
            return
        if any(c["名称"] == 新名 for c in 配置["分类"]):
            QMessageBox.warning(self, "重命名失败", "已存在同名分类")
            return
        try:
            旧文件夹 = os.path.join(数据目录, 旧名)
            if os.path.exists(旧文件夹):
                os.rename(旧文件夹, os.path.join(数据目录, 新名))
        except Exception as e:
            QMessageBox.warning(self, "重命名失败", str(e))
            return
        分类["名称"] = 新名
        保存配置()
        self.刷新全部()

    def _删除分类(self, 索引):
        """删除指定索引的分类：可恢复图标移回原位，其余快捷方式删除；删完刷新"""
        分类 = 配置["分类"][索引]
        可恢复数 = sum(1 for t in 分类["条目"] if t.get("原路径"))
        提示 = f"确定删除分类「{分类['名称']}」吗？"
        if 分类["条目"]:
            提示 += f"\n\n其中 {可恢复数} 个图标将移回原来的位置，其余快捷方式将被删除。"
        if QMessageBox.question(self, "删除分类", 提示) != QMessageBox.Yes:
            return
        for 条目 in list(分类["条目"]):
            self.恢复条目(分类, 条目, 刷新=False)
        shutil.rmtree(os.path.join(数据目录, 分类["名称"]), ignore_errors=True)
        配置["分类"].pop(索引)  # 删掉这个分类（pop 比 remove 高效，且需要按索引删）
        保存配置()
        self.刷新全部()

    def _check_autohide(self):
        if not self.isVisible():
            self._检测边缘唤出()
            return
        if self._固定:
            return
        # 鼠标还没进入过窗口（例如刚从托盘点开，鼠标还停在托盘上）：先不隐藏
        if not self._鼠标进过窗口:
            return
        if QApplication.activePopupWidget() is not None or QApplication.activeModalWidget() is not None:
            return
        if QApplication.mouseButtons() != Qt.NoButton:  # 正在拖拽/按键
            return
        范围 = self.frameGeometry().adjusted(-6, -6, 6, 6)
        if not 范围.contains(QCursor.pos()):
            self.hide()

    def _更新停靠边(self):
        """用户拖动窗口松手后调用：记录窗口贴住了屏幕哪条边（上/左/右，容差12像素）。
        边缘唤出只响应贴住的边——没贴任何边就不触发（用户要求：不要哪里都显示）。"""
        self._停靠边 = set()
        屏 = QGuiApplication.screenAt(self.frameGeometry().center()) or QGuiApplication.primaryScreen()
        几何 = 屏.availableGeometry()
        窗 = self.frameGeometry()
        if 窗.top() - 几何.top() <= 12:
            self._停靠边.add("上")
        if 窗.left() - 几何.left() <= 12:
            self._停靠边.add("左")
        if 几何.right() - 窗.right() <= 12:
            self._停靠边.add("右")

    def _检测边缘唤出(self):
        """窗口隐藏时鼠标顶住窗口停靠的那条屏幕边 → 约 0.3 秒自动弹出收纳盒。
        只认用户把窗口拖过去贴住的边（_停靠边）：拖到顶上只认上边，拖到左边只认左边；
        下边缘永远留给任务栏；
        多屏时以鼠标所在屏幕的可用区为准；按住鼠标键时不触发（防拖窗口误触）；
        前台是全屏窗口（游戏/全屏视频）时自动安静，绝不抢焦点。"""
        if not 配置.get("边缘唤出", True):
            self._贴边次数 = 0
            return
        if QApplication.mouseButtons() != Qt.NoButton:
            self._贴边次数 = 0
            return
        if 前台是全屏():
            self._贴边次数 = 0
            return
        if not self._停靠边:
            self._贴边次数 = 0
            return  # 窗口没贴过任何边：边缘唤出不生效，等用户拖一次再说
        位 = QCursor.pos()
        屏 = QGuiApplication.screenAt(位) or QGuiApplication.primaryScreen()
        几何 = 屏.availableGeometry()
        # 只认停靠的那几条边；下边永远留给任务栏
        贴边 = (("上" in self._停靠边 and 位.y() <= 几何.top() + 2) or
                ("左" in self._停靠边 and 位.x() <= 几何.left() + 2) or
                ("右" in self._停靠边 and 位.x() >= 几何.right() - 2))
        if 贴边:
            self._贴边次数 += 1
            if self._贴边次数 >= 3:  # 轮询间隔100毫秒，3帧≈0.3秒弹出
                self._贴边次数 = 0
                self.弹出()
        else:
            self._贴边次数 = 0

    # --- 显示与固定 ---

    def toggle(self):
        if self.isVisible():
            if not self._固定:
                self.hide()
            else:
                self.raise_()
                self.activateWindow()
        else:
            self.弹出()

    def 弹出(self):
        # 每次弹出都重置标记：等鼠标真正进入窗口后，才启用"移开自动隐藏"
        self._鼠标进过窗口 = False
        屏 = QGuiApplication.screenAt(QCursor.pos()) or QGuiApplication.primaryScreen()
        几何 = 屏.availableGeometry()
        # 套用记住的尺寸，但不能比当前屏幕的可用区域还大（换了小屏幕时防止超出边界）
        宽 = max(最小宽, min(配置["宽"], 几何.width()))
        高 = max(最小高, min(配置["高"], 几何.height()))
        self.resize(宽, 高)
        x = QCursor.pos().x() - 宽 // 2
        y = QCursor.pos().y() - 高 // 2
        x = max(几何.left() + 4, min(x, 几何.right() - 宽 - 4))
        y = max(几何.top() + 4, min(y, 几何.bottom() - 高 - 4))
        self.move(x, y)
        _静音悬停音()  # 盒子贴着鼠标弹出，先掐掉 Qt 补发给指针底下新图标的幻影悬停音
        self.show()
        self.raise_()
        self.activateWindow()

    def _on_search(self, 文字):
        """搜索框文字变了：记住关键字，让每个分类页重建网格（重建内部做过滤）"""
        self.搜索字 = 文字 or ""
        for i in range(self.页面栈.count()):
            页 = self.页面栈.widget(i)
            if isinstance(页, 分类页):
                页._上次列数 = None
                页.重建网格()

    def _show_settings(self):
        """点标题栏 ⚙️ 按钮打开设置对话框。
        parent=self：对话框挂在主窗上（能跟随主窗位置、模态阻止主窗交互）。
        关键：创建后用 setWindowFlags 强制覆盖主窗的 透明/无边框 属性——
        只保留正常的 Dialog | Window 标志，别继承 FramelessWindowHint/WA_TranslucentBackground。"""
        对话框 = 设置对话框(self, self)  # parent=主窗
        # 强制重设 window flags：去掉 FramelessWindowHint 和 WindowStaysOnTopHint，
        # 加上正常的 Dialog 标志，让它有边框、能正常显示、不透明
        对话框.setWindowFlags(Qt.Dialog | Qt.Window)
        # 位置：弹在主窗右边，紧贴主窗；如果右边放不下就弹在左边
        主窗几何 = self.frameGeometry()
        设置宽 = 对话框.sizeHint().width()
        if 设置宽 <= 0:
            设置宽 = 420
        设置高 = 对话框.sizeHint().height()
        if 设置高 <= 0:
            设置高 = 380
        屏宽 = QGuiApplication.primaryScreen().availableGeometry().width()
        # 先假设弹右边
        x = 主窗几何.right() + 4
        y = 主窗几何.top()
        if x + 设置宽 > 屏宽:  # 右边放不下 → 弹左边
            x = 主窗几何.left() - 设置宽 - 4
        对话框.move(x, y)
        对话框.exec_()

    def _show_support(self):
        """点标题栏 ☕ 按钮打开赞助对话框。
        和设置对话框一样：parent=None + setWindowFlags 确保不继承主窗的透明/无边框。"""
        对话框 = 赞助对话框()
        对话框.setWindowFlags(Qt.Dialog | Qt.Window)
        对话框.exec_()

    def _on_pin_toggled(self, 勾选):
        self._固定 = 勾选
        self.固定钮.setText("📌 已固定" if 勾选 else "📌 固定")
        配置["固定"] = 勾选
        保存配置()

    def 打开后收起(self):
        if not self._固定:
            QTimer.singleShot(500, self._hide_delayed)

    def _hide_delayed(self):
        if not self._固定:
            self.hide()

    def enterEvent(self, 事件):
        # 鼠标进入窗口：记下"进过"，此后鼠标移开窗口就会自动隐藏
        self._鼠标进过窗口 = True
        super().enterEvent(事件)

    def hideEvent(self, 事件):
        # 窗口隐藏时（鼠标移开自动隐藏/点✕/双击收起）把当前尺寸记进配置，
        # 下次唤出就还是用户拖出来的大小
        配置["宽"] = self.width()
        配置["高"] = self.height()
        保存配置()
        super().hideEvent(事件)

    @staticmethod
    def _有符号16位(值):
        """Windows 消息里的坐标是 16 位有符号整数，自己转一下（可能为负数，多屏时会出现）"""
        return 值 - 0x10000 if 值 >= 0x8000 else 值

    def nativeEvent(self, 事件类型, 消息):
        """接管 Windows 的底层窗口消息，实现"无边框窗口拖边缘缩放"。

        原理：鼠标在窗口上移动时，Windows 会不断发 WM_NCHITTEST 消息问
        "鼠标现在在窗口的哪个部位？"。我们检测鼠标是否贴着窗口边缘，
        贴着就回答"左边/右下角/上边…"，Windows 收到后会自动把鼠标变成
        双向箭头、并接管拖动缩放的全过程（系统原生手感，跟普通窗口一样）。
        """
        WM_NCHITTEST = 0x0084  # Windows 询问"鼠标命中窗口哪个部位"的消息编号
        # 八个部位的系统编号：告诉 Windows 后它就知道该往哪个方向缩放
        HTLEFT, HTRIGHT, HTTOP, HTBOTTOM = 10, 11, 12, 15
        HTTOPLEFT, HTTOPRIGHT, HTBOTTOMLEFT, HTBOTTOMRIGHT = 13, 14, 16, 17
        if 事件类型 == b"windows_generic_MSG":
            结构 = wintypes.MSG.from_address(int(消息))
            if 结构.message == WM_NCHITTEST:
                # lParam 的低16位=鼠标屏幕X，高16位=屏幕Y
                屏幕x = self._有符号16位(结构.lParam & 0xFFFF)
                屏幕y = self._有符号16位((结构.lParam >> 16) & 0xFFFF)
                # 换算成"相对本窗口左上角"的坐标
                客户x, 客户y = win32gui.ScreenToClient(int(self.winId()), (屏幕x, 屏幕y))
                宽, 高 = self.width(), self.height()
                # 判断鼠标是否贴在四条边上（缩放宽 像素的边带内）
                贴左 = 客户x < 缩放宽
                贴右 = 客户x >= 宽 - 缩放宽
                贴上 = 客户y < 缩放宽
                贴下 = 客户y >= 高 - 缩放宽
                命中 = 0  # 0=不在边缘，交给 Qt 按普通客户区处理
                if 贴上 and 贴左:
                    命中 = HTTOPLEFT
                elif 贴上 and 贴右:
                    命中 = HTTOPRIGHT
                elif 贴下 and 贴左:
                    命中 = HTBOTTOMLEFT
                elif 贴下 and 贴右:
                    命中 = HTBOTTOMRIGHT
                elif 贴上:
                    命中 = HTTOP
                elif 贴下:
                    命中 = HTBOTTOM
                elif 贴左:
                    命中 = HTLEFT
                elif 贴右:
                    命中 = HTRIGHT
                if 命中:
                    # 返回 (True, 部位编号)：告诉 Qt"这条消息我们自己处理了，结果是这个部位"
                    return True, 命中
        return super().nativeEvent(事件类型, 消息)

    def keyPressEvent(self, 事件):
        if 事件.key() == Qt.Key_Escape:
            # Esc 先清搜索框，其次才收起窗口
            if getattr(self, "搜索字", ""):
                self.搜索框.clear()
                return
            if not self._固定:
                self.hide()
        super().keyPressEvent(事件)

    def eventFilter(self, 对象, 事件):
        """整个根窗口都能拖，但以下控件自身需要交互，不能让拖窗口逻辑拦截：
        - QToolButton（固定/关闭/+ 新建按钮）
        - QTabBar（标签栏点击切换）
        - 盒内图标（拖拽/双击打开）
        - QScrollArea / QStackedWidget（滚动页面）"""
        if 事件.type() == QEvent.MouseButtonPress and 事件.button() == Qt.LeftButton:
            # 交互控件白名单：这些控件上的左键按下不拖窗口
            if isinstance(对象, (QToolButton, QTabBar, QScrollArea, QStackedWidget,
                                QDialog, QLineEdit, QComboBox, QSpinBox)):
                return super().eventFilter(对象, 事件)
            名称 = 对象.objectName() if hasattr(对象, 'objectName') else ""
            if 名称 == "盒内图标":  # 图标按钮自身有拖拽逻辑
                return super().eventFilter(对象, 事件)
            self._拖动偏移 = 事件.globalPos() - self.frameGeometry().topLeft()
            return True
        if 事件.type() == QEvent.MouseMove and self._拖动偏移 is not None \
                and 事件.buttons() & Qt.LeftButton:
            self.move(事件.globalPos() - self._拖动偏移)
            return True
        if 事件.type() == QEvent.MouseButtonRelease:
            if self._拖动偏移 is not None:
                self._更新停靠边()  # 只跟用户手动拖窗口走：拖到哪条边就认哪条边
            self._拖动偏移 = None
        return super().eventFilter(对象, 事件)

    # --- 拖入 / 移动 / 恢复 ---

    def 拖入(self, 分类, 路径列表):
        """把外部文件收进分类。
        关键规则（按用户要求 2026-10-11 改）：
        - 来源在桌面（不管啥后缀）→ 移动原文件进盒子，记录原路径可恢复
        - 来源不在桌面（比如开始菜单、下载文件夹）→ 在盒子里建快捷方式，不动原文件
        .lnk/.url 在移动时直接搬文件本体（快捷方式本身就是"指向别处"的）。"""
        文件夹 = os.path.join(数据目录, 分类["名称"])
        os.makedirs(文件夹, exist_ok=True)
        成功 = 0
        for 路径 in 路径列表:
            路径 = os.path.normpath(路径)
            if not os.path.exists(路径):
                continue
            if 分类.get("虚拟"):
                # 拖到「全部」总览页 = 默认落点：自动按类型归进对应分类；
                # 类型对不上分类就进第一个真实分类（一般是"常用"）
                分类 = next((c for c in 配置["分类"]
                             if not c.get("虚拟")
                             and c["名称"] == 取类型分类(路径)),
                            None) or next((c for c in 配置["分类"]
                                           if not c.get("虚拟")), 分类)
                文件夹 = os.path.join(数据目录, 分类["名称"])
                os.makedirs(文件夹, exist_ok=True)
            elif 配置.get("自动分类", True):
                # 按类型自动归类：只往"已经存在的"对应分类里放；
                # 该类型分类还没建就老老实实进当前拖入的分类，绝不自作主张新建标签
                命中 = next((c for c in 配置["分类"]
                             if c["名称"] == 取类型分类(路径)), None)
                if 命中 is not None and 命中 is not 分类:
                    分类 = 命中
                    文件夹 = os.path.join(数据目录, 分类["名称"])
                    os.makedirs(文件夹, exist_ok=True)
            基名 = 清理名称(os.path.splitext(os.path.basename(路径))[0], "图标")
            try:
                lnk名 = 唯一lnk名(文件夹, 基名)
                目标 = os.path.join(文件夹, lnk名)

                if 是否在桌面(路径):
                    # ===== 从桌面拖来的 → 实际移动文件进盒子 =====
                    if 路径.lower().endswith((".lnk", ".url")):
                        # .lnk/.url：直接搬文件本体进盒子
                        shutil.move(路径, 目标)
                    else:
                        # 普通文件（.exe/.txt/.jpg 等）：**也搬原文件进盒子**
                        # 给它改名成 .lnk 后缀没用（又不是快捷方式），
                        # 所以就搬原文件本身，保留原扩展名
                        # 先构造不冲突的目标名（去掉 .lnk 后缀逻辑——我们要原名）
                        原名 = os.path.basename(路径)
                        真实目标 = os.path.join(文件夹, 避让路径(目标 + ".bak"))  # 先占位不用
                        # 直接搬，用原名（唯一lnk名返回的基名+".lnk"没用，手动构造唯一原名）
                        真实目标 = os.path.join(文件夹, 避让路径(os.path.join(文件夹, 原名)))
                        shutil.move(路径, 真实目标)
                        目标 = 真实目标  # 更新后面引用
                        # lnk名也改成原名（盒子里存的就是原文件本身，不是 .lnk）
                        lnk名 = 原名

                    原路径 = 路径  # 记录原位置，右键可恢复
                else:
                    # ===== 从别处（非桌面）拖来的 → 只建快捷方式，不动原文件 =====
                    if 路径.lower().endswith((".lnk", ".url")):
                        shutil.copy2(路径, 目标)
                    else:
                        建快捷方式(目标, 路径)
                    原路径 = ""  # 没移动过，恢复就只是删盒子里的东西

            except Exception as e:
                QMessageBox.warning(self, "收纳失败", f"{路径}\n{e}")
                continue
            分类["条目"].append({"名称": 基名, "文件": lnk名, "原路径": 原路径})
            成功 += 1
        if 成功:
            保存配置()
            self.刷新全部()

    def 移动条目(self, 源分类名, 文件名, 目标分类, 目标文件名=None):
        """把源分类里的条目移到目标分类。
        目标文件名：落在某个图标上时传它，插到该图标前面（拖拽排序）；
        不传则挪到分类末尾。同一分类内调用 = 纯排序，不动磁盘文件。"""
        源分类 = next((c for c in 配置["分类"] if c["名称"] == 源分类名), None)
        if 源分类 is None:
            return
        条目 = next((t for t in 源分类["条目"] if t["文件"] == 文件名), None)
        if 条目 is None:
            return
        if 源分类名 == 目标分类["名称"]:
            if 文件名 == 目标文件名:
                return  # 拖回自己身上，什么都不做
            列表 = 源分类["条目"]
            源序 = next((i for i, t in enumerate(列表) if t["文件"] == 文件名), None)
            目标序 = next((i for i, t in enumerate(列表)
                           if t["文件"] == 目标文件名), None) if 目标文件名 else None
            if 目标序 is None:
                # 落在空白处：挪到本分类最后
                列表.remove(条目)
                列表.append(条目)
            else:
                # 落在图标上：两个图标互换位置。
                # （之前用"插到目标前面"，拖到自己紧挨着的邻居时等于原地不动，
                #  看起来就像"第二次拖失效"；互换则任何一次拖放都必有变化）
                列表[源序], 列表[目标序] = 列表[目标序], 列表[源序]
            保存配置()
            self.刷新全部()
            return
        目标文件夹 = os.path.join(数据目录, 目标分类["名称"])
        os.makedirs(目标文件夹, exist_ok=True)
        try:
            新名 = 唯一lnk名(目标文件夹, os.path.splitext(文件名)[0])
            shutil.move(os.path.join(数据目录, 源分类名, 文件名),
                        os.path.join(目标文件夹, 新名))
        except Exception as e:
            QMessageBox.warning(self, "移动失败", str(e))
            return
        条目["文件"] = 新名
        源分类["条目"].remove(条目)
        目标序 = next((i for i, t in enumerate(目标分类["条目"])
                       if t["文件"] == 目标文件名), None) if 目标文件名 else None
        if 目标序 is None:
            目标分类["条目"].append(条目)
        else:
            目标分类["条目"].insert(目标序, 条目)
        保存配置()
        self.刷新全部()

    def 恢复条目(self, 分类, 条目, 刷新=True):
        """有原路径的移回原位置，没有的直接删除快捷方式"""
        lnk = os.path.join(数据目录, 分类["名称"], 条目["文件"])
        原 = 条目.get("原路径") or ""
        try:
            if 原:
                shutil.move(lnk, 避让路径(原))
            else:
                os.remove(lnk)
            分类["条目"].remove(条目)
            if 刷新:
                保存配置()
                self.刷新全部()
        except Exception as e:
            QMessageBox.warning(self, "操作失败", str(e))

    # --- 分类管理 ---

    def _add_category(self):  # add_group 改 ASCII（PyQt5 槽名必须 ASCII）
        已有 = {c["名称"] for c in 配置["分类"]}
        序号 = 1
        while f"新分类{序号}" in 已有:
            序号 += 1
        文本, 确定 = QInputDialog.getText(self, "新建分类", "分类名称：", text=f"新分类{序号}")
        if not 确定:
            return
        名 = 清理名称(文本, "未命名")
        if 名 in 已有:
            QMessageBox.warning(self, "创建失败", "已存在同名分类")
            return
        os.makedirs(os.path.join(数据目录, 名), exist_ok=True)
        配置["分类"].append({"名称": 名, "条目": []})
        保存配置()
        # 刷新全部后，新分类是最后一个标签，直接切过去让用户可以立即往里拖图标
        self.刷新全部()
        最后索引 = self.标签栏.count() - 1
        if 最后索引 >= 0:
            self.标签栏.setCurrentIndex(最后索引)
            self.页面栈.setCurrentIndex(最后索引)


# ---------------- 设置对话框 ----------------

class 设置对话框(QDialog):
    """收纳盒统一设置对话框：窗口大小 + 图标间距（热键已移除）。
    每改一个设置项点"应用"后立即生效，点"确定"则自动应用并关闭。"""

    def __init__(self, parent, 主窗):
        super().__init__(parent)
        self.主窗 = 主窗
        self.setWindowTitle("收纳盒设置")
        布局 = QVBoxLayout(self)
        布局.setSpacing(12)

        # ---------- 第 1 组：窗口大小 ----------
        组1 = QLabel("📐 窗口大小")
        组1.setStyleSheet("font-weight: 600;")
        布局.addWidget(组1)

        宽行 = QHBoxLayout()
        宽行.addWidget(QLabel("宽度："))
        self.宽输入 = QSpinBox()
        self.宽输入.setRange(最小宽, 3840)
        self.宽输入.setSingleStep(10)
        self.宽输入.setValue(配置["宽"])
        宽行.addWidget(self.宽输入)
        宽行.addWidget(QLabel("像素"))
        宽行.addStretch()
        布局.addLayout(宽行)

        高行 = QHBoxLayout()
        高行.addWidget(QLabel("高度："))
        self.高输入 = QSpinBox()
        self.高输入.setRange(最小高, 2160)
        self.高输入.setSingleStep(10)
        self.高输入.setValue(配置["高"])
        高行.addWidget(self.高输入)
        高行.addWidget(QLabel("像素"))
        高行.addStretch()
        布局.addLayout(高行)

        # ---------- 第 2 组：图标间距 ----------
        组2 = QLabel("🔢 图标间距")
        组2.setStyleSheet("font-weight: 600; margin-top: 8px;")
        布局.addWidget(组2)

        间距行 = QHBoxLayout()
        间距行.addWidget(QLabel("图标之间的空隙："))
        self.间距输入 = QSpinBox()
        self.间距输入.setRange(0, 30)
        self.间距输入.setSingleStep(1)
        self.间距输入.setValue(配置.get("图标间距", 默认图标间距))
        self.间距输入.setToolTip("调整后图标之间的距离和格子的密度都会变")
        间距行.addWidget(self.间距输入)
        间距行.addWidget(QLabel("像素"))
        间距行.addStretch()
        布局.addLayout(间距行)

        # ---------- 第 3 组：开机自启 ----------
        组3 = QLabel("🚀 开机自启动")
        组3.setStyleSheet("font-weight: 600; margin-top: 8px;")
        布局.addWidget(组3)

        self.自启勾选 = QCheckBox("开机时自动启动收纳盒（放在系统托盘）")
        self.自启勾选.setChecked(配置.get("开机自启", 查开机自启()))
        self.自启勾选.setToolTip(
            "开启后 Windows 登录时会自动启动本程序（藏在托盘里，不会挡事）。\n"
            "用的是系统注册表 HKCU\\Run 项，不需要管理员权限，随时可以关掉。"
        )
        布局.addWidget(self.自启勾选)

        # ---------- 第 4 组：音效提示 ----------
        组4 = QLabel("🔔 音效提示")
        组4.setStyleSheet("font-weight: 600; margin-top: 8px;")
        布局.addWidget(组4)

        self.音效勾选 = QCheckBox("盒内图标悬停 / 点击时播放声音")
        self.音效勾选.setChecked(bool(配置.get("音效", True)))
        self.音效勾选.setToolTip(
            "鼠标移到图标上会轻响一声，单击一声短音，双击一声重音。\n"
            "觉得吵就取消勾选，立即生效。"
        )
        布局.addWidget(self.音效勾选)

        # ---------- 第 5 组：排序与归类 ----------
        组5 = QLabel("🗂 排序与归类")
        组5.setStyleSheet("font-weight: 600; margin-top: 8px;")
        布局.addWidget(组5)

        self.置顶勾选 = QCheckBox("常用置顶（打开次数多的图标自动排到最前面）")
        self.置顶勾选.setChecked(bool(配置.get("常用置顶", False)))
        self.置顶勾选.setToolTip("开着时显示顺序按打开次数排；手动拖拽排序照样能拖，\n但下次刷新时顺序会被使用频率排回去")
        布局.addWidget(self.置顶勾选)

        self.自动类勾选 = QCheckBox("按类型自动归类（拖入文件时自动分进对应分类）")
        self.自动类勾选.setChecked(bool(配置.get("自动分类", True)))
        self.自动类勾选.setToolTip("拖入文件时：如果你已经建了 文档/图片/视频/音乐/程序/压缩包/其他 这些分类，\n会按后缀自动分进对应分类；没建就不动（不会自己冒出新分类标签）\n不想被自动分就取消勾选")
        布局.addWidget(self.自动类勾选)

        # ---------- 第 6 组：外观与唤出 ----------
        组6 = QLabel("🎨 外观与唤出")
        组6.setStyleSheet("font-weight: 600; margin-top: 8px;")
        布局.addWidget(组6)

        主题行 = QHBoxLayout()
        主题行.addWidget(QLabel("配色主题："))
        self.主题选择 = QComboBox()
        self.主题选择.addItems(list(主题配色.keys()))
        self.主题选择.setCurrentText(str(配置.get("主题", "深色")))
        self.主题选择.setToolTip("五套配色：深色 / 浅色 / 玻璃 / 青墨 / 暖沙，切换立即生效")
        主题行.addWidget(self.主题选择)
        主题行.addStretch()
        布局.addLayout(主题行)

        self.边缘勾选 = QCheckBox("鼠标顶住窗口停靠的屏幕边约 0.3 秒，自动唤出收纳盒")
        self.边缘勾选.setToolTip("把窗口拖到屏幕上/左/右边贴住，唤出就只响应那一条边；\n没贴边则不触发，下边缘永远留给任务栏；\n游戏或全屏应用在前台时自动安静，不抢焦点。")
        self.边缘勾选.setChecked(bool(配置.get("边缘唤出", True)))
        布局.addWidget(self.边缘勾选)

        # ---------- 按钮行 ----------
        布局.addStretch(1)
        行 = QHBoxLayout()
        恢复钮 = QPushButton("恢复默认")
        恢复钮.clicked.connect(self._reset_all)  # 槽名必须 ASCII！
        应用钮 = QPushButton("应用")
        应用钮.clicked.connect(self._apply_all)  # 槽名必须 ASCII！
        确定 = QPushButton("确定")
        取消 = QPushButton("取消")
        确定.clicked.connect(self._confirm)  # 槽名必须 ASCII！
        取消.clicked.connect(self.reject)
        行.addWidget(恢复钮)
        行.addStretch()
        行.addWidget(应用钮)
        行.addWidget(确定)
        行.addWidget(取消)
        布局.addLayout(行)

    def _apply_all(self):
        """把窗口大小、图标间距、开机自启都写进配置并立即生效。"""
        # 窗口大小
        配置["宽"] = int(self.宽输入.value())
        配置["高"] = int(self.高输入.value())
        self.主窗.resize(配置["宽"], 配置["高"])

        # 图标间距
        配置["图标间距"] = int(self.间距输入.value())
        self.主窗.刷新全部()

        # 音效开关（全局配置即时生效，不用重启）
        配置["音效"] = self.音效勾选.isChecked()

        # 排序 / 归类 / 边缘唤出（都是配置项，即时生效）
        配置["常用置顶"] = self.置顶勾选.isChecked()
        配置["自动分类"] = self.自动类勾选.isChecked()
        配置["边缘唤出"] = self.边缘勾选.isChecked()

        # 换肤：主题变了就整体重发样式表，立即生效
        新主题 = self.主题选择.currentText()
        if 新主题 != 配置.get("主题", "深色"):
            配置["主题"] = 新主题
            QApplication.instance().setStyleSheet(取样式表())
        self.主窗.刷新全部()  # 置顶顺序等变化立即反映到网格

        # 开机自启（写注册表）
        自启 = self.自启勾选.isChecked()
        配置["开机自启"] = 自启
        if not 设开机自启(自启):
            # 写注册表失败（少见），弹一下提示但不拦截保存
            QMessageBox.warning(self, "开机自启",
                                "注册表写入失败，开机自启可能没设上。\n"
                                "可以打开「任务管理器 → 启动」手动管理。")

        保存配置()
        return True

    def _confirm(self):
        """确定按钮：先应用，成功了就关闭"""
        if self._apply_all():
            self.accept()

    def _reset_all(self):
        """一键恢复默认值"""
        self.宽输入.setValue(默认宽)
        self.高输入.setValue(默认高)
        self.间距输入.setValue(默认图标间距)
        # 开机自启默认不开
        self.自启勾选.setChecked(False)
        # 音效提示默认开
        self.音效勾选.setChecked(True)
        # 排序归类/外观默认：不置顶、自动分类开、边缘唤出开、深色主题
        self.置顶勾选.setChecked(False)
        self.自动类勾选.setChecked(True)
        self.边缘勾选.setChecked(True)
        self.主题选择.setCurrentText("深色")
        self._apply_all()


# ---------------- 托盘 ----------------

class 系统托盘(QSystemTrayIcon):
    def __init__(self, 主窗):
        super().__init__(生成程序图标(), 主窗)
        self.主窗 = 主窗
        self.setToolTip("桌面收纳盒")
        菜单 = QMenu()
        动作显示 = QAction("显示 / 隐藏收纳盒", 菜单)
        动作显示.triggered.connect(主窗.toggle)
        动作设置 = QAction("设置…", 菜单)
        动作设置.triggered.connect(self.open_settings)
        动作退出 = QAction("退出", 菜单)
        动作退出.triggered.connect(QApplication.quit)
        菜单.addAction(动作显示)
        菜单.addSeparator()
        菜单.addAction(动作设置)
        菜单.addSeparator()
        动作赞助 = QAction("☕ 赞助开发…", 菜单)
        动作赞助.triggered.connect(self.show_support)  # 槽名 ASCII！
        菜单.addAction(动作赞助)
        动作关于 = QAction("ℹ️ 关于此插件…", 菜单)
        动作关于.triggered.connect(self.show_about)  # 槽名 ASCII！
        菜单.addAction(动作关于)
        菜单.addSeparator()
        菜单.addAction(动作退出)
        self.setContextMenu(菜单)
        self.activated.connect(self._on_activated)
        self.show()

    def _on_activated(self, 原因):
        if 原因 in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self.主窗.toggle()

    def open_settings(self):
        """托盘菜单点"设置"：先确保主窗可见（因为标题栏的设置按钮在主窗上），
        然后直接复用主窗的 _show_settings（那个方法正确处理了置顶/透明继承，不闪）。"""
        if not self.主窗.isVisible():
            self.主窗.弹出()
        # 直接调主窗自己的设置对话框方法，不用在这里搞一套
        self.主窗._show_settings()

    def show_about(self):
        """托盘菜单点"关于此插件…"：弹版本信息 + 开发者 + 功能列表 + 感谢话术。
        注意：about 的第一个参数必须是 QWidget 或 None——之前传托盘对象
        （QSystemTrayIcon 不是窗口）导致点击后静默失败什么都不弹。"""
        QMessageBox.about(None, "关于 桌面收纳盒",
            f"<b>桌面收纳盒</b><br>版本：{版本}<br>"
            "开发者：Lexinm-Code（GitHub）<br>"
            "联系邮箱：326495489@qq.com"
            "<hr>"
            "· 双击桌面空白处 / 托盘图标 / 全局热键唤出，鼠标移开自动隐藏<br>"
            "· 上方标签自定义分类，左侧类型分类（文档/图片/视频…），拖入自动归类<br>"
            "· 「全部」页总览所有图标；支持拖拽排序、右键打开位置 / 删除（进回收站）<br>"
            "· 边缘唤出：你把窗口贴到哪条屏幕边，就只认哪条边；游戏全屏时自动安静<br>"
            "· 搜索过滤、常用置顶、图片缩略图、悬停详情<br>"
            "· 五套皮肤、音效提示、图钉固定、开机自启"
            "<hr>"
            "感谢你选择桌面收纳盒，愿它让你的桌面从此整洁有序、图标随手可寻。<br>"
            "如有任何建议，欢迎给我发邮件，或到 GitHub 上给我留言，让它变得更好。")

    def show_support(self):
        """托盘菜单点"赞助开发…"：弹一个对话框展示微信/支付宝收款码图片。
        图片放在 程序目录/ 下（根目录直接放），文件名固定 wechat.png / alipay.png。
        如果图片不存在就友好提示，不崩溃。"""
        dlg = 赞助对话框()
        dlg.exec_()


# ---------------- 赞助对话框 ----------------

class 赞助对话框(QDialog):
    """托盘菜单"☕ 赞助开发…"弹出的对话框。
    展示两张收款码图片（微信 + 支付宝），附感谢文字。
    图片路径相对程序目录（根目录直接放）：wechat.png 和 alipay.png。"""

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(Qt.Dialog | Qt.Window)
        self.setWindowTitle("☕ 支持开发")
        self.setMinimumWidth(480)
        布局 = QVBoxLayout(self)
        布局.setSpacing(12)
        布局.setContentsMargins(20, 20, 20, 16)

        # 顶部标题
        标题 = QLabel("❤️ 感谢您的支持！")
        标题.setStyleSheet("font-size: 13pt; font-weight: 700; color: #ff7a8a; background: transparent;")
        标题.setAlignment(Qt.AlignCenter)
        布局.addWidget(标题)

        # 顶部说明
        说明 = QLabel(
            "这个小工具从构思到上线，敲了几千行代码，<br>"
            "踩了无数坑才稳定下来。<br>"
            "如果它帮到了您，<b>一杯咖啡</b>就是对我最大的鼓励~"
        )
        说明.setStyleSheet("font-size: 9.5pt; color: #c9d1e0; background: transparent;")
        说明.setWordWrap(True)
        说明.setAlignment(Qt.AlignCenter)
        布局.addWidget(说明)

        # 收款码区域
        布局.addSpacing(4)
        二维码区 = QHBoxLayout()
        二维码区.setSpacing(16)

        # 微信
        微信栏 = QVBoxLayout()
        微信栏.setSpacing(6)
        微信标签 = QLabel("💬 微信")
        微信标签.setStyleSheet("font-weight: 600; color: #07C160; background: transparent; font-size: 10pt;")
        微信标签.setAlignment(Qt.AlignCenter)
        微信栏.addWidget(微信标签)
        微信图片 = self._加载收款码("wechat.png")
        微信栏.addWidget(微信图片, alignment=Qt.AlignCenter)
        二维码区.addLayout(微信栏)

        # 支付宝
        支付宝栏 = QVBoxLayout()
        支付宝栏.setSpacing(6)
        支付宝标签 = QLabel("🔵 支付宝")
        支付宝标签.setStyleSheet("font-weight: 600; color: #1677FF; background: transparent; font-size: 10pt;")
        支付宝标签.setAlignment(Qt.AlignCenter)
        支付宝栏.addWidget(支付宝标签)
        支付宝图片 = self._加载收款码("alipay.png")
        支付宝栏.addWidget(支付宝图片, alignment=Qt.AlignCenter)
        二维码区.addLayout(支付宝栏)

        布局.addLayout(二维码区)

        # 底部标语
        布局.addSpacing(4)
        标语 = QLabel("— 本软件永久免费，赞助纯属自愿 —")
        标语.setStyleSheet("color: #6a7386; font-size: 8.5pt; background: transparent;")
        标语.setAlignment(Qt.AlignCenter)
        布局.addWidget(标语)

        # 关闭按钮
        按钮行 = QHBoxLayout()
        按钮行.addStretch()
        关闭 = QPushButton("好的，谢谢")
        关闭.setFixedWidth(120)
        关闭.clicked.connect(self.accept)
        按钮行.addWidget(关闭)
        布局.addLayout(按钮行)

    def _加载收款码(self, 文件名):
        """加载收款码图片。
        优先级：
        1. 先查 exe 同级（用户可手动替换收款码）
        2. 再查 sys._MEIPASS（onefile 打包后 PyInstaller 解压的临时目录，内嵌资源在这）
        3. 都没有 → 显示占位提示
        这样 onefile 打包默认带内嵌收款码，但用户想自己换的话也可以。"""
        图片路径 = os.path.join(程序目录, 文件名)
        if not os.path.exists(图片路径) and getattr(sys, 'frozen', False):
            # onefile 模式：资源解压到 sys._MEIPASS 临时目录
            内嵌路径 = os.path.join(sys._MEIPASS, 文件名)
            if os.path.exists(内嵌路径):
                图片路径 = 内嵌路径
        if os.path.exists(图片路径):
            pix = QPixmap(图片路径)
            if not pix.isNull():
                # 等比缩放到最大 180x180
                pix = pix.scaled(180, 180, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                lbl = QLabel()
                lbl.setPixmap(pix)
                lbl.setStyleSheet(
                    "background: #fff; border-radius: 8px; padding: 4px;"
                    "border: 1px solid #3a3f4e;"
                )
                return lbl
        # 图片不存在 → 显示占位提示
        占位 = QLabel(f"（{文件名} 未找到）")
        占位.setStyleSheet(
            "background: #22242c; color: #8d97ab; border: 1px dashed #444; "
            "border-radius: 6px; padding: 20px; font-size: 8pt;"
        )
        占位.setFixedSize(180, 180)
        占位.setAlignment(Qt.AlignCenter)
        return 占位


# ---------------- 样式 ----------------

# ---------------- 样式（三套主题换肤） ----------------
# 样式表写成"模板 + 配色变量"：@@颜色名@@ 占位，取样式表() 按 配置["主题"] 替换。
# 设置对话框里换主题 -> QApplication 重新 setStyleSheet 立即整体换肤。

主题配色 = {
    "深色": {
        "根背景": "rgba(30,32,40,244)", "根边框": "rgba(255,255,255,30)",
        "标题文字": "#eef1f7", "提示文字": "#8d97ab", "弱文字": "#9aa4b8",
        "按钮文字": "#c9d0dd", "按钮背景": "rgba(255,255,255,22)", "按钮悬停": "rgba(255,255,255,40)",
        "强调色": "#4f8cff", "强调底": "rgba(79,140,255,70)", "强调微底": "rgba(79,140,255,30)",
        "图标文字": "#e5e9f2", "图标悬停": "rgba(255,255,255,28)",
        "新建底": "rgba(255,255,255,10)", "新建边框": "rgba(255,255,255,45)",
        "标签底": "rgba(255,255,255,10)", "标签悬停": "rgba(255,255,255,18)",
        "栈底": "rgba(255,255,255,8)", "滚动条": "rgba(255,255,255,70)",
        "弹层底": "#2b2e37", "弹层字": "#e6e9f0", "弹层边框": "rgba(255,255,255,35)",
        "输入底": "#22242c", "输入边框": "rgba(255,255,255,40)",
        "对话框按钮": "#3a3f4c", "分隔线": "rgba(255,255,255,28)", "反白": "#ffffff",
    },
    "浅色": {
        "根背景": "rgba(244,246,251,244)", "根边框": "rgba(0,0,0,30)",
        "标题文字": "#1f2430", "提示文字": "#6a7386", "弱文字": "#5a6275",
        "按钮文字": "#3a4152", "按钮背景": "rgba(0,0,0,14)", "按钮悬停": "rgba(0,0,0,28)",
        "强调色": "#2f6fe4", "强调底": "rgba(47,111,228,45)", "强调微底": "rgba(47,111,228,20)",
        "图标文字": "#262b36", "图标悬停": "rgba(0,0,0,20)",
        "新建底": "rgba(0,0,0,6)", "新建边框": "rgba(0,0,0,40)",
        "标签底": "rgba(0,0,0,8)", "标签悬停": "rgba(0,0,0,16)",
        "栈底": "rgba(0,0,0,6)", "滚动条": "rgba(0,0,0,45)",
        "弹层底": "#f2f5fa", "弹层字": "#232837", "弹层边框": "rgba(0,0,0,30)",
        "输入底": "#ffffff", "输入边框": "rgba(0,0,0,30)",
        "对话框按钮": "#dde3ee", "分隔线": "rgba(0,0,0,20)", "反白": "#ffffff",
    },
    "玻璃": {
        "根背景": "rgba(24,28,40,150)", "根边框": "rgba(255,255,255,70)",
        "标题文字": "#f4f7ff", "提示文字": "#c2cbdc", "弱文字": "#c8d2e4",
        "按钮文字": "#e2e8f4", "按钮背景": "rgba(255,255,255,34)", "按钮悬停": "rgba(255,255,255,58)",
        "强调色": "#57a0ff", "强调底": "rgba(87,160,255,80)", "强调微底": "rgba(87,160,255,40)",
        "图标文字": "#eef2fa", "图标悬停": "rgba(255,255,255,40)",
        "新建底": "rgba(255,255,255,16)", "新建边框": "rgba(255,255,255,70)",
        "标签底": "rgba(255,255,255,16)", "标签悬停": "rgba(255,255,255,28)",
        "栈底": "rgba(255,255,255,14)", "滚动条": "rgba(255,255,255,90)",
        "弹层底": "#262b38", "弹层字": "#eef2fa", "弹层边框": "rgba(255,255,255,50)",
        "输入底": "rgba(255,255,255,30)", "输入边框": "rgba(255,255,255,60)",
        "对话框按钮": "#3b4150", "分隔线": "rgba(255,255,255,40)", "反白": "#ffffff",
    },
    "青墨": {
        "根背景": "rgba(20,38,36,242)", "根边框": "rgba(120,220,200,40)",
        "标题文字": "#e6f7f1", "提示文字": "#8fb4ab", "弱文字": "#9fc3b8",
        "按钮文字": "#cfe8e0", "按钮背景": "rgba(140,230,210,24)", "按钮悬停": "rgba(140,230,210,44)",
        "强调色": "#2fbf9f", "强调底": "rgba(47,191,159,80)", "强调微底": "rgba(47,191,159,36)",
        "图标文字": "#e8f5ef", "图标悬停": "rgba(140,230,210,30)",
        "新建底": "rgba(140,230,210,12)", "新建边框": "rgba(140,230,210,45)",
        "标签底": "rgba(140,230,210,12)", "标签悬停": "rgba(140,230,210,22)",
        "栈底": "rgba(140,230,210,10)", "滚动条": "rgba(160,240,220,70)",
        "弹层底": "#1d3532", "弹层字": "#e6f2ee", "弹层边框": "rgba(140,230,210,40)",
        "输入底": "#162926", "输入边框": "rgba(140,230,210,45)",
        "对话框按钮": "#2c4a45", "分隔线": "rgba(140,230,210,28)", "反白": "#ffffff",
    },
    "暖沙": {
        "根背景": "rgba(250,243,232,244)", "根边框": "rgba(160,120,60,40)",
        "标题文字": "#3d2f1e", "提示文字": "#94806a", "弱文字": "#8a7458",
        "按钮文字": "#5a4630", "按钮背景": "rgba(150,110,60,16)", "按钮悬停": "rgba(150,110,60,32)",
        "强调色": "#c47b2c", "强调底": "rgba(196,123,44,45)", "强调微底": "rgba(196,123,44,22)",
        "图标文字": "#40332a", "图标悬停": "rgba(150,110,60,22)",
        "新建底": "rgba(150,110,60,10)", "新建边框": "rgba(150,110,60,40)",
        "标签底": "rgba(150,110,60,10)", "标签悬停": "rgba(150,110,60,20)",
        "栈底": "rgba(150,110,60,8)", "滚动条": "rgba(150,110,60,50)",
        "弹层底": "#faf3e8", "弹层字": "#3d2f1e", "弹层边框": "rgba(150,110,60,40)",
        "输入底": "#ffffff", "输入边框": "rgba(150,110,60,45)",
        "对话框按钮": "#eadfce", "分隔线": "rgba(150,110,60,26)", "反白": "#ffffff",
    },
}

样式表模板 = """
QWidget { font-family: "Microsoft YaHei UI"; }
#根窗口 { background: @@根背景@@; border: 1px solid @@根边框@@; border-radius: 14px; }
#标题文字 { color: @@标题文字@@; font-size: 11pt; font-weight: 600; background: transparent; }
#提示文字 { color: @@提示文字@@; font-size: 8pt; background: transparent; }
QToolButton#固定钮, QToolButton#关闭钮 {
    color: @@按钮文字@@; background: @@按钮背景@@; border: none;
    border-radius: 6px; padding: 3px 8px; font-size: 9pt;
}
QToolButton#固定钮:hover, QToolButton#关闭钮:hover { background: @@按钮悬停@@; }
QToolButton#固定钮:checked { background: @@强调色@@; color: @@反白@@; }
/* 赞助按钮：低调，悬停变彩色 */
QToolButton#赞助钮 {
    color: @@提示文字@@; background: transparent; border: none;
    border-radius: 6px; padding: 2px 4px; font-size: 10pt;
}
QToolButton#赞助钮:hover { color: #ff7a8a; background: @@按钮背景@@; }
QToolButton#盒内图标 {
    background: transparent; border: none; border-radius: 8px;
    color: @@图标文字@@; font-size: 8pt; padding: 3px 1px;
}
QToolButton#盒内图标:hover { background: @@图标悬停@@; }
/* "+" 新建分类按钮 */
QToolButton#新建分类 {
    color: @@弱文字@@; background: @@新建底@@; border: 1px dashed @@新建边框@@;
    border-radius: 8px; font-size: 14pt; padding: 2px 10px; min-width: 28px; max-width: 36px;
}
QToolButton#新建分类:hover { color: @@反白@@; border-color: @@强调色@@; background: @@强调微底@@; }
/* 分类标签栏 */
QTabBar#分类标签栏 { background: transparent; }
QTabBar#分类标签栏::tab {
    color: @@弱文字@@; background: @@标签底@@; border: none;
    border-top-left-radius: 8px; border-top-right-radius: 8px;
    padding: 5px 12px; margin-right: 4px; font-size: 9pt;
}
QTabBar#分类标签栏::tab:hover { color: @@标题文字@@; background: @@标签悬停@@; }
QTabBar#分类标签栏::tab:selected {
    color: @@反白@@; background: @@强调底@@; border-bottom: 2px solid @@强调色@@;
}
QTabBar#分类标签栏::scroller { background: transparent; border: none; }
QTabBar#分类标签栏::left-arrow, QTabBar#分类标签栏::right-arrow {
    width: 16px; height: 16px; background: @@按钮背景@@; border-radius: 4px;
}
QWidget#类型侧栏 { background: transparent; }
QToolButton#类型侧栏钮 {
    color: @@弱文字@@; background: @@标签底@@; border: none;
    border-radius: 8px; padding: 5px 10px; font-size: 9pt;
}
QToolButton#类型侧栏钮:hover { color: @@标题文字@@; background: @@标签悬停@@; }
QToolButton#类型侧栏钮:checked {
    color: @@反白@@; background: @@强调底@@; border-left: 2px solid @@强调色@@;
}
QStackedWidget#分类页面栈 { background: @@栈底@@; border-radius: 10px; }
QScrollArea { border: none; background: transparent; }
QScrollArea > QWidget > QWidget { background: transparent; }
QScrollBar:vertical { background: transparent; width: 6px; margin: 2px 1px; }
QScrollBar::handle:vertical { background: @@滚动条@@; border-radius: 3px; min-height: 30px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; border: none; height: 0; width: 0; }
/* 搜索框 */
QLineEdit#搜索框 {
    background: @@输入底@@; color: @@图标文字@@; border: 1px solid @@输入边框@@;
    border-radius: 8px; padding: 4px 8px; font-size: 9pt;
}
QLineEdit#搜索框:focus { border-color: @@强调色@@; }
QMenu { background: @@弹层底@@; color: @@弹层字@@; border: 1px solid @@弹层边框@@; border-radius: 8px; padding: 4px; }
QMenu::item { padding: 6px 20px; border-radius: 6px; }
QMenu::item:selected { background: @@强调色@@; color: @@反白@@; }
QMenu::separator { height: 1px; background: @@分隔线@@; margin: 4px 8px; }
QMessageBox, QInputDialog, QDialog { background: @@弹层底@@; color: @@弹层字@@; }
QMessageBox QLabel, QInputDialog QLabel, QDialog QLabel { color: @@弹层字@@; background: transparent; }
QMessageBox QPushButton, QInputDialog QPushButton, QDialog QPushButton, QMessageBox QCheckBox, QDialog QCheckBox {
    background: @@对话框按钮@@; color: @@弹层字@@; border: none; border-radius: 6px; padding: 5px 16px;
}
QMessageBox QPushButton:hover, QInputDialog QPushButton:hover, QDialog QPushButton:hover { background: @@强调色@@; }
QCheckBox { color: @@弹层字@@; background: transparent; spacing: 6px; }
QKeySequenceEdit { background: @@输入底@@; color: @@弹层字@@; border: 1px solid @@输入边框@@; border-radius: 6px; padding: 4px; }
QSpinBox { background: @@输入底@@; color: @@弹层字@@; border: 1px solid @@输入边框@@; border-radius: 6px; padding: 3px 6px; }
QSpinBox::up-button, QSpinBox::down-button { width: 18px; background: @@对话框按钮@@; border: none; }
QSpinBox::up-button:hover, QSpinBox::down-button:hover { background: @@强调色@@; }
QComboBox { background: @@输入底@@; color: @@弹层字@@; border: 1px solid @@输入边框@@; border-radius: 6px; padding: 4px 8px; }
QComboBox::drop-down { border: none; width: 22px; }
QComboBox QAbstractItemView { background: @@弹层底@@; color: @@弹层字@@; selection-background-color: @@强调色@@; border: 1px solid @@弹层边框@@; }
QInputDialog QLineEdit { background: @@输入底@@; color: @@弹层字@@; border: 1px solid @@输入边框@@; border-radius: 6px; padding: 4px; }
"""


def 取样式表():
    """按配置里的主题名生成最终样式表；主题名不认识就回落深色"""
    配色 = 主题配色.get(配置.get("主题", "深色"), 主题配色["深色"])
    样式 = 样式表模板
    for 键, 值 in 配色.items():
        样式 = 样式.replace("@@" + 键 + "@@", 值)
    return 样式


# ---------------- 入口 ----------------

def main():
    写启动日志("进入 main()")
    # --- DPI 感知声明：必须放在窗口创建之前 ---
    # 让本进程使用"物理像素"坐标。否则在系统缩放（如 125%）下，
    # 鼠标钩子报上来的坐标和窗口判断用的坐标对不上，"双击桌面唤出"就会判定错位失效
    # 这里用诊断脚本里实测验证过的 SetProcessDPIAware（全机统一缩放级别）
    ctypes.windll.user32.SetProcessDPIAware()
    写启动日志("DPI 感知声明完成")

    # 让 Qt 按 DPI 自动缩放界面：高分屏下不模糊、大小正常
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)

    app = QApplication(sys.argv)
    写启动日志("Qt 主程序对象创建完成")
    app.setQuitOnLastWindowClosed(False)
    app.setWindowIcon(生成程序图标())
    app.setStyleSheet(取样式表())

    # 防止重复启动
    单实例 = QSharedMemory("桌面收纳盒-单实例")
    if not 单实例.create(1):
        # 已有一个收纳盒在运行（点窗口✕只是隐藏，进程仍在托盘，不算关闭）
        写启动日志("检测到已有实例在运行，本次启动直接退出")
        QMessageBox.information(None, "桌面收纳盒", "收纳盒已经在运行了（看右下角托盘，双击它就能唤出）")
        return
    写启动日志("单实例锁创建成功（说明之前没有残留进程）")

    首次 = not os.path.exists(配置文件)
    加载配置()

    # 开机自启项自修复：注册表里存在自启项时，每次启动都按当前真实路径重写一遍。
    # 1) 旧版命令是 "cmd.exe /c start bat"，开机必闪黑窗口——在这里顺手洗掉；
    # 2) 开发版/exe 版切换、程序目录搬家后，旧 Run 值会指向错误位置，同样被纠正。
    if 查开机自启():
        设开机自启(True)

    主窗 = 收纳盒窗口()
    托盘 = 系统托盘(主窗)
    写启动日志("主窗口和托盘图标创建完成")

    # 全局鼠标钩子：双击桌面空白处唤出（回调经信号转回主线程）
    钩子 = 桌面双击钩子(lambda x, y: 主窗.show_request.emit())
    钩子.start()
    # 钩子在独立线程里安装，最多等 1 秒拿到安装结果并记入日志
    for _ in range(20):
        if 钩子.安装状态 is not None:
            break
        time.sleep(0.05)
    if 钩子.安装状态:
        写启动日志("确认：鼠标钩子安装成功，双击桌面功能可用")
    else:
        写启动日志("警告：鼠标钩子安装失败，双击桌面将无反应（详见上面的错误码）")

    写启动日志("主窗口和托盘图标创建完成")

    # 鼠标钩子已在 桌面钩子.py 里独立安装，双击桌面空白处即可唤出

    def _on_quit():  # 退出清理
        try:
            钩子.stop()
        except Exception:
            pass

    app.aboutToQuit.connect(_on_quit)

    写启动日志("进入事件循环（托盘图标此刻应该可见）")

    # 正常启动完，5 秒后删掉启动日志（只用来排查"开不出来"的情况，
    # 正常跑起来没用，放根目录也脏）。删不掉就拉倒，别影响程序。
    QTimer.singleShot(5000, lambda: (os.remove(_启动日志路径) if os.path.exists(_启动日志路径) else None))

    if 首次:
        保存配置()
        主窗.弹出()
        托盘.showMessage("桌面收纳盒",
                         "双击桌面空白处唤出；右上角图钉可固定窗口",
                         QSystemTrayIcon.Information, 5000)

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
