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

from PyQt5.QtCore import (Qt, QSize, QTimer, QEvent, pyqtSignal, QMimeData,
                          QFileInfo, QSharedMemory, QAbstractNativeEventFilter)
写启动日志("PyQt5.QtCore 加载完成")
from PyQt5.QtGui import (QColor, QCursor, QFont, QIcon, QKeySequence, QDrag,
                         QPainter, QPixmap, QGuiApplication)
from PyQt5.QtWidgets import (QApplication, QAction, QCheckBox, QDialog, QFrame,
                             QGridLayout, QHBoxLayout, QInputDialog, QLabel, QMenu,
                             QMessageBox, QPushButton, QKeySequenceEdit, QScrollArea,
                             QSpinBox, QStyle, QSystemTrayIcon, QTabBar, QTabWidget,
                             QToolButton, QVBoxLayout, QWidget, QFileIconProvider,
                             QStackedWidget)
写启动日志("PyQt5 界面库全部加载完成")

import win32api
import win32con
import win32com.client
import win32gui  # 用来做"无边框窗口拖边缘缩放"时的坐标换算
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
    Run 注册表的值是 cmd.exe 来执行的（不是 Python！）。
    打包后（sys.frozen=True）直接启动 exe 就行，不用 bat；
    开发时用 bat 里调 pythonw.exe。"""
    if getattr(sys, 'frozen', False):
        # 打包后：注册表直接启动 exe，不用套 cmd /c start 了
        # 但加个 "" 当 start 的标题参数，防止路径被误解析
        return f'start "" "{sys.executable}"'
    else:
        # 开发时：cmd /c start 调 bat，bat 里再调 pythonw.exe
        bat路径 = os.path.join(程序目录, "启动收纳盒.bat")
        return f'cmd.exe /c start "" "{bat路径}"'

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
        "分类": 数据.get("分类") or [{"名称": "常用", "条目": []}],
    }
    for 分类 in 配置["分类"]:
        分类.setdefault("名称", "未命名")
        分类.setdefault("条目", [])


def 保存配置():
    os.makedirs(数据目录, exist_ok=True)
    with open(配置文件, "w", encoding="utf-8") as f:
        json.dump(配置, f, ensure_ascii=False, indent=2)


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


def 取文件图标(路径):
    try:
        图 = 图标提供者.icon(QFileInfo(路径))
        if not 图.isNull():
            return 图
    except Exception:
        pass
    return QApplication.style().standardIcon(QStyle.SP_FileIcon)


# ---------------- 盒内图标按钮 ----------------

class 盒内图标(QToolButton):
    def __init__(self, 主窗, 卡片, 条目):
        super().__init__()
        self.主窗 = 主窗
        self.卡片 = 卡片
        self.条目 = 条目
        self.setObjectName("盒内图标")
        self.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.setIconSize(QSize(36, 36))
        self.setFixedSize(QSize(图标按钮宽, 图标按钮高))
        名称 = 条目["名称"]
        self.setText(名称 if len(名称) <= 8 else 名称[:7] + "…")
        self.setToolTip(名称)
        self.setIcon(取文件图标(self.绝对路径()))
        self.setCursor(Qt.PointingHandCursor)
        self._按下位置 = None

    def 绝对路径(self):
        return os.path.join(数据目录, self.卡片.分类["名称"], self.条目["文件"])

    def _open_item(self):
        """打开条目（双击或菜单）"""
        try:
            os.startfile(self.绝对路径())
        except Exception:
            pass
        self.主窗.打开后收起()

    def mouseDoubleClickEvent(self, 事件):
        self._open_item()

    def mousePressEvent(self, 事件):
        if 事件.button() == Qt.LeftButton:
            self._按下位置 = 事件.pos()
        super().mousePressEvent(事件)

    def mouseMoveEvent(self, 事件):
        if self._按下位置 is None or not (事件.buttons() & Qt.LeftButton):
            return
        if (事件.pos() - self._按下位置).manhattanLength() < 12:
            return
        self._按下位置 = None
        拖 = QDrag(self)
        数据 = QMimeData()
        数据.setData(内部拖放格式, json.dumps(
            {"源": self.卡片.分类["名称"], "文件": self.条目["文件"]},
            ensure_ascii=False).encode("utf-8"))
        拖.setMimeData(数据)
        拖.setPixmap(self.icon().pixmap(32, 32))
        拖.exec_(Qt.MoveAction)

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
                           lambda: self.主窗.恢复条目(self.卡片.分类, self.条目))
        else:
            菜单.addAction("🗑 移除快捷方式",
                           lambda: self.主窗.恢复条目(self.卡片.分类, self.条目))
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
        if self.分类["条目"]:
            for 序, 条目 in enumerate(self.分类["条目"]):
                self.网格.addWidget(盒内图标(self.主窗, self, 条目), 序 // 列数, 序 % 列数)
        else:
            提示 = QLabel("把桌面图标拖到这里")
            提示.setObjectName("提示文字")
            提示.setAlignment(Qt.AlignCenter)
            提示.setStyleSheet("color: #6a7386; font-size: 9pt;")
            # 空提示横跨整行（占满所有列）
            self.网格.addWidget(提示, 0, 0, 1, 列数)
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

        总行数 = max(3, (len(self.分类["条目"]) + 列数 - 1) // 列数)

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
            self.主窗.移动条目(载荷.get("源", ""), 载荷.get("文件", ""), self.分类)
        elif 数据.hasUrls():
            路径 = [u.toLocalFile() for u in 数据.urls() if u.isLocalFile()]
            if 路径:
                事件.acceptProposedAction()
                self.主窗.拖入(self.分类, 路径)


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

        # 分类标签栏 + 页面切换区
        标签栏行 = QHBoxLayout()
        标签栏行.setContentsMargins(0, 2, 0, 0)
        标签栏行.setSpacing(4)

        # QTabBar：可点的分类标签，标签多时自动出现左右翻页箭头
        self.标签栏 = QTabBar()
        self.标签栏.setObjectName("分类标签栏")
        self.标签栏.setDocumentMode(True)      # 让标签和下方内容无缝衔接
        self.标签栏.setUsesScrollButtons(True)  # 标签多了自动出现左右箭头
        self.标签栏.setTabsClosable(False)      # 关掉默认的"×"关闭按钮（我们用右键菜单删分类）
        self.标签栏.setMovable(False)           # 不让用户拖动标签改顺序（简化逻辑）
        self.标签栏.currentChanged.connect(self._on_tab_changed)  # 切换标签时更新页面
        # 标签上右键 → 重命名/删除分类；空白处右键 → 新建分类
        self.标签栏.customContextMenuRequested.connect(self._on_tab_menu)
        self.标签栏.setContextMenuPolicy(Qt.CustomContextMenu)
        标签栏行.addWidget(self.标签栏, 1)

        # "+"新建分类按钮，固定在标签栏最右边
        self.新建钮 = QToolButton()
        self.新建钮.setObjectName("新建分类")
        self.新建钮.setText("＋")
        self.新建钮.setCursor(Qt.PointingHandCursor)
        self.新建钮.setToolTip("新建分类")
        self.新建钮.clicked.connect(self._add_category)  # 槽名必须 ASCII！
        标签栏行.addWidget(self.新建钮)

        内.addLayout(标签栏行)

        # QStackedWidget：每次只显示一个分类的图标页面，点标签时切换
        self.页面栈 = QStackedWidget()
        self.页面栈.setObjectName("分类页面栈")
        内.addWidget(self.页面栈, 1)

        # 自动隐藏：定时检查鼠标是否离开窗口
        self._隐藏计时 = QTimer(self)
        self._隐藏计时.setInterval(300)
        self._隐藏计时.timeout.connect(self._check_autohide)
        self._隐藏计时.start()

        self.show_request.connect(self.toggle)  # 鼠标钩子线程 → 主线程
        self.刷新全部()
        if 配置["固定"]:
            self.固定钮.setChecked(True)

    # --- 刷新与自动隐藏 ---

    def 刷新全部(self):
        """重建所有标签和对应页面：配置里的分类变了（新建/删/重命名）时调用。
        策略：先记住当前选中的标签索引，清空标签栏和页面栈，按配置里的分类列表重新建，
        最后尽量切回原来那个标签（原来的索引可能被删了，就切到相邻的）"""
        原来的索引 = self.标签栏.currentIndex()
        # 清空页面栈（旧页面会被 deleteLater 释放）
        while self.页面栈.count():
            w = self.页面栈.widget(0)
            self.页面栈.removeWidget(w)
            if w:
                w.setParent(None)  # 同重建网格：立刻摘掉，防止旧页面残留
                w.deleteLater()
        # 清空标签栏
        while self.标签栏.count():
            self.标签栏.removeTab(0)
        # 按配置里的分类列表逐个建标签 + 对应页面
        for 分类 in 配置["分类"]:
            页面 = 分类页(self, 分类)
            self.页面栈.addWidget(页面)
            # 标签文字：分类名 + 条目计数，让用户一眼看到每个分类装了多少
            标签文字 = f"{分类['名称']}  ({len(分类['条目'])})"
            self.标签栏.addTab(标签文字)
        # 切回原来的索引（夹在有效范围内防越界）
        总数 = self.标签栏.count()
        if 总数 > 0:
            目标 = max(0, min(原来的索引, 总数 - 1))
            self.标签栏.setCurrentIndex(目标)
            self.页面栈.setCurrentIndex(目标)

    def _on_tab_changed(self, 索引):
        """用户点击标签栏切换分类时，同步切换 StackedWidget 里的页面"""
        self.页面栈.setCurrentIndex(索引)

    def _on_tab_menu(self, 位置):
        """标签栏右键菜单：
        - 点在某个标签上 → 重命名 / 删除 那个分类
        - 点在标签栏空白处（没点到标签）→ 新建分类"""
        全局位置 = self.标签栏.mapToGlobal(位置)
        # tabAt 返回 -1 表示点在空白处，否则是点中的标签索引
        点中的标签索引 = self.标签栏.tabAt(位置)
        菜单 = QMenu(self)
        if 点中的标签索引 >= 0 and 点中的标签索引 < len(配置["分类"]):
            # 点在具体标签上 → 管理这个分类
            菜单.addAction("重命名分类…", lambda: self._重命名分类(点中的标签索引))
            菜单.addAction("删除分类…", lambda: self._删除分类(点中的标签索引))
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
        if not self.isVisible() or self._固定:
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
        self.show()
        self.raise_()
        self.activateWindow()

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
        if 事件.key() == Qt.Key_Escape and not self._固定:
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
            if isinstance(对象, (QToolButton, QTabBar, QScrollArea, QStackedWidget, QDialog)):
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

    def 移动条目(self, 源分类名, 文件名, 目标分类):
        if 源分类名 == 目标分类["名称"]:
            return
        源分类 = next((c for c in 配置["分类"] if c["名称"] == 源分类名), None)
        if 源分类 is None:
            return
        条目 = next((t for t in 源分类["条目"] if t["文件"] == 文件名), None)
        if 条目 is None:
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
        目标分类["条目"].append(条目)
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
        组1.setStyleSheet("font-weight: 600; color: #eef1f7;")
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
        组2.setStyleSheet("font-weight: 600; color: #eef1f7; margin-top: 8px;")
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
        组3.setStyleSheet("font-weight: 600; color: #eef1f7; margin-top: 8px;")
        布局.addWidget(组3)

        self.自启勾选 = QCheckBox("开机时自动启动收纳盒（放在系统托盘）")
        self.自启勾选.setChecked(配置.get("开机自启", 查开机自启()))
        self.自启勾选.setToolTip(
            "开启后 Windows 登录时会自动启动本程序（藏在托盘里，不会挡事）。\n"
            "用的是系统注册表 HKCU\\Run 项，不需要管理员权限，随时可以关掉。"
        )
        布局.addWidget(self.自启勾选)

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

样式表 = """
QWidget { font-family: "Microsoft YaHei UI"; }
#根窗口 { background: rgba(30,32,40,244); border: 1px solid rgba(255,255,255,30); border-radius: 14px; }
#标题文字 { color: #eef1f7; font-size: 11pt; font-weight: 600; background: transparent; }
#提示文字 { color: #8d97ab; font-size: 8pt; background: transparent; }
QToolButton#固定钮, QToolButton#关闭钮 {
    color: #c9d0dd; background: rgba(255,255,255,22); border: none;
    border-radius: 6px; padding: 3px 8px; font-size: 9pt;
}
QToolButton#固定钮:hover, QToolButton#关闭钮:hover { background: rgba(255,255,255,40); }
QToolButton#固定钮:checked { background: #4f8cff; color: #ffffff; }
/* 赞助按钮：低调灰色，悬停才变彩色（方案 C）*/
QToolButton#赞助钮 {
    color: #5a6275;  /* 深灰，非常低调 */
    background: transparent;
    border: none;
    border-radius: 6px;
    padding: 2px 4px;
    font-size: 10pt;
}
QToolButton#赞助钮:hover {
    color: #ff7a8a;  /* 悬停变粉红色 */
    background: rgba(255,255,255,25);
}
QToolButton#盒内图标 {
    background: transparent; border: none; border-radius: 8px;
    color: #e5e9f2; font-size: 8pt; padding: 3px 1px;
}
QToolButton#盒内图标:hover { background: rgba(255,255,255,28); }
/* "+" 新建分类按钮 */
QToolButton#新建分类 {
    color: #9aa4b8; background: rgba(255,255,255,10);
    border: 1px dashed rgba(255,255,255,45); border-radius: 8px;
    font-size: 14pt; padding: 2px 10px; min-width: 28px; max-width: 36px;
}
QToolButton#新建分类:hover { color: #ffffff; border-color: #4f8cff; background: rgba(79,140,255,30); }
/* 分类标签栏（QTabBar） */
QTabBar#分类标签栏 { background: transparent; }
QTabBar#分类标签栏::tab {
    color: #9aa4b8; background: rgba(255,255,255,10);
    border: none; border-top-left-radius: 8px; border-top-right-radius: 8px;
    padding: 5px 12px; margin-right: 4px; font-size: 9pt;
}
QTabBar#分类标签栏::tab:hover { color: #eef1f7; background: rgba(255,255,255,18); }
QTabBar#分类标签栏::tab:selected {
    color: #ffffff; background: rgba(79,140,255,70);
    border-bottom: 2px solid #4f8cff;
}
/* 标签栏左右翻页箭头（标签多时出现） */
QTabBar#分类标签栏::scroller { background: transparent; border: none; }
QTabBar#分类标签栏::left-arrow, QTabBar#分类标签栏::right-arrow {
    width: 16px; height: 16px; background: rgba(255,255,255,20); border-radius: 4px;
}
/* 分类页面栈容器 */
QStackedWidget#分类页面栈 { background: rgba(255,255,255,8); border-radius: 10px; }
QScrollArea { border: none; background: transparent; }
QScrollArea > QWidget > QWidget { background: transparent; }
QScrollBar:vertical { background: transparent; width: 6px; margin: 2px 1px; }
QScrollBar::handle:vertical { background: rgba(255,255,255,70); border-radius: 3px; min-height: 30px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; border: none; height: 0; width: 0; }
QMenu { background: #2b2e37; color: #e6e9f0; border: 1px solid rgba(255,255,255,35); border-radius: 8px; padding: 4px; }
QMenu::item { padding: 6px 20px; border-radius: 6px; }
QMenu::item:selected { background: #4f8cff; }
QMenu::separator { height: 1px; background: rgba(255,255,255,28); margin: 4px 8px; }
QMessageBox, QInputDialog, QDialog { background: #2b2e37; color: #e6e9f0; }
QMessageBox QLabel, QInputDialog QLabel, QDialog QLabel { color: #e6e9f0; background: transparent; }
QMessageBox QPushButton, QInputDialog QPushButton, QDialog QPushButton {
    background: #3a3f4c; color: #e6e9f0; border: none; border-radius: 6px; padding: 5px 16px;
}
QMessageBox QPushButton:hover, QInputDialog QPushButton:hover, QDialog QPushButton:hover { background: #4f8cff; }
QKeySequenceEdit { background: #22242c; color: #e6e9f0; border: 1px solid rgba(255,255,255,40); border-radius: 6px; padding: 4px; }
QSpinBox { background: #22242c; color: #e6e9f0; border: 1px solid rgba(255,255,255,40); border-radius: 6px; padding: 3px 6px; }
QSpinBox::up-button, QSpinBox::down-button { width: 18px; background: #3a3f4c; border: none; }
QSpinBox::up-button:hover, QSpinBox::down-button:hover { background: #4f8cff; }
QInputDialog QLineEdit { background: #22242c; color: #e6e9f0; border: 1px solid rgba(255,255,255,40); border-radius: 6px; padding: 4px; }
"""


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
    app.setStyleSheet(样式表)

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
