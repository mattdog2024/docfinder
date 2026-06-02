# -*- coding: utf-8 -*-
"""
文档搜索索引 - 程序入口
"""

import sys
import os

# 确保 Windows 下 DPI 缩放正常
if sys.platform == 'win32':
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon

# 添加项目根目录到路径（便携版兼容）
# 打包后 __file__ 在 _internal 目录里，需要特殊处理
if getattr(sys, 'frozen', False):
    # 打包后的 exe 运行环境
    _base_dir = sys._MEIPASS  # PyInstaller 解压目录
    _exe_dir = os.path.dirname(sys.executable)  # exe 所在目录
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))
    _exe_dir = _base_dir

if _base_dir not in sys.path:
    sys.path.insert(0, _base_dir)

from ui.main_window import MainWindow, STYLE_SHEET


def main():
    # 高 DPI 支持
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("文档搜索索引")
    app.setOrganizationName("DocFinder")

    # 设置默认字体（中文友好）
    # 明确指定微软雅黑，避免 Windows 字体回退导致汉字显示异常
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    font.setHintingPreference(QFont.PreferFullHinting)
    app.setFont(font)
    # 备用字体列表（如果微软雅黑不可用）
    from PyQt5.QtGui import QFontDatabase
    db = QFontDatabase()
    available = db.families()
    for fallback in ["Microsoft YaHei", "微软雅黑", "SimHei", "黑体", "SimSun", "宋体"]:
        if fallback in available:
            font.setFamily(fallback)
            app.setFont(font)
            break

    # 应用样式表
    app.setStyleSheet(STYLE_SHEET)

    # 设置应用图标（先找 _MEIPASS 里的 assets，再找 exe 旁边的 assets）
    icon_path = ''
    for search_dir in [_base_dir, _exe_dir]:
        for icon_name in ['logo_icon.ico', 'logo_icon.png']:
            candidate = os.path.join(search_dir, 'assets', icon_name)
            if os.path.exists(candidate):
                icon_path = candidate
                break
        if icon_path:
            break
    if icon_path:
        app_icon = QIcon(icon_path)
        app.setWindowIcon(app_icon)

    window = MainWindow()
    if icon_path:
        window.setWindowIcon(QIcon(icon_path))
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
