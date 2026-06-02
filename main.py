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
_base_dir = os.path.dirname(os.path.abspath(__file__))
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

    # 设置应用图标
    icon_path = os.path.join(_base_dir, 'assets', 'logo_icon.ico')
    if not os.path.exists(icon_path):
        icon_path = os.path.join(_base_dir, 'assets', 'logo_icon.png')
    if os.path.exists(icon_path):
        app_icon = QIcon(icon_path)
        app.setWindowIcon(app_icon)

    window = MainWindow()
    if os.path.exists(icon_path):
        window.setWindowIcon(QIcon(icon_path))
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
