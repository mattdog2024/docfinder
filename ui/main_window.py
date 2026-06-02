# -*- coding: utf-8 -*-
"""
主窗口 UI
明亮风格，基于 PyQt5
"""

import os
import sys
import time
import threading
import subprocess
import multiprocessing
from datetime import datetime
from typing import List, Dict, Optional

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QListWidget, QListWidgetItem,
    QSplitter, QTextEdit, QStatusBar, QProgressBar,
    QFileDialog, QMessageBox, QCheckBox, QGroupBox,
    QTabWidget, QFrame, QScrollArea, QSizePolicy,
    QApplication, QAction, QToolBar, QComboBox,
    QDialog, QDialogButtonBox, QGridLayout, QSpinBox,
    QTreeWidget, QTreeWidgetItem, QHeaderView, QAbstractItemView
)
from PyQt5.QtCore import (
    Qt, QThread, pyqtSignal, QTimer, QSize, QSettings
)
from PyQt5.QtGui import (
    QFont, QIcon, QColor, QPalette, QPixmap,
    QTextCharFormat, QTextCursor
)
from PyQt5.QtGui import QTextDocument

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.indexer import IndexEngine, IndexBuilder
from core.extractor import SUPPORTED_EXTENSIONS

APP_NAME = "文档搜索索引"
APP_VERSION = "1.9"

# ─── 样式表 ──────────────────────────────────────────────────────────────────

STYLE_SHEET = """
QMainWindow, QWidget {
    background-color: #F5F5F5;
    color: #212121;
    font-family: "Microsoft YaHei", "微软雅黑", Arial, sans-serif;
    font-size: 13px;
}

QMenuBar {
    background-color: #FFFFFF;
    border-bottom: 1px solid #E0E0E0;
    padding: 2px;
}
QMenuBar::item:selected {
    background-color: #E3F2FD;
    border-radius: 4px;
}
QMenu {
    background-color: #FFFFFF;
    border: 1px solid #E0E0E0;
    border-radius: 4px;
}
QMenu::item:selected {
    background-color: #E3F2FD;
}

QToolBar {
    background-color: #FFFFFF;
    border-bottom: 1px solid #E0E0E0;
    spacing: 4px;
    padding: 4px;
}

QPushButton {
    background-color: #1976D2;
    color: white;
    border: none;
    border-radius: 5px;
    padding: 7px 16px;
    font-weight: bold;
    min-height: 28px;
}
QPushButton:hover {
    background-color: #1565C0;
}
QPushButton:pressed {
    background-color: #0D47A1;
}
QPushButton:disabled {
    background-color: #BDBDBD;
    color: #757575;
}
QPushButton#btn_secondary {
    background-color: #FFFFFF;
    color: #1976D2;
    border: 1px solid #1976D2;
}
QPushButton#btn_secondary:hover {
    background-color: #E3F2FD;
}
QPushButton#btn_danger {
    background-color: #D32F2F;
    color: white;
    font-weight: bold;
    font-size: 13px;
    border: 2px solid #B71C1C;
    padding: 6px 14px;
    min-width: 80px;
    min-height: 30px;
}
QPushButton#btn_danger:hover {
    background-color: #B71C1C;
    border: 2px solid #7F0000;
}
QPushButton#btn_danger:disabled {
    background-color: #EF9A9A;
    color: #FFFFFF;
    border: 2px solid #E57373;
}
QPushButton#btn_success {
    background-color: #388E3C;
}
QPushButton#btn_success:hover {
    background-color: #2E7D32;
}

QLineEdit {
    background-color: #FFFFFF;
    border: 1px solid #BDBDBD;
    border-radius: 5px;
    padding: 6px 10px;
    font-size: 14px;
    min-height: 30px;
}
QLineEdit:focus {
    border: 2px solid #1976D2;
}

QTextEdit {
    background-color: #FFFFFF;
    border: 1px solid #E0E0E0;
    border-radius: 5px;
    padding: 8px;
    font-size: 13px;
    line-height: 1.5;
}

QTreeWidget {
    background-color: #FFFFFF;
    border: 1px solid #E0E0E0;
    border-radius: 5px;
    alternate-background-color: #FAFAFA;
    show-decoration-selected: 1;
}
QTreeWidget::item {
    padding: 6px 4px;
    border-bottom: 1px solid #F5F5F5;
}
QTreeWidget::item:selected {
    background-color: #E3F2FD;
    color: #1565C0;
}
QTreeWidget::item:hover {
    background-color: #F5F5F5;
}
QHeaderView::section {
    background-color: #EEEEEE;
    border: none;
    border-right: 1px solid #E0E0E0;
    border-bottom: 1px solid #E0E0E0;
    padding: 6px 8px;
    font-weight: bold;
    color: #424242;
}

QGroupBox {
    border: 1px solid #E0E0E0;
    border-radius: 6px;
    margin-top: 12px;
    padding-top: 8px;
    background-color: #FFFFFF;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 6px;
    color: #1976D2;
    font-weight: bold;
    left: 10px;
}

QCheckBox {
    spacing: 6px;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 2px solid #BDBDBD;
    border-radius: 3px;
    background-color: #FFFFFF;
}
QCheckBox::indicator:checked {
    background-color: #1976D2;
    border-color: #1976D2;
    image: url(none);
}
QCheckBox::indicator:hover {
    border-color: #1976D2;
}

QProgressBar {
    border: none;
    border-radius: 4px;
    background-color: #E0E0E0;
    text-align: center;
    height: 8px;
    font-size: 11px;
}
QProgressBar::chunk {
    background-color: #1976D2;
    border-radius: 4px;
}

QStatusBar {
    background-color: #1976D2;
    color: white;
    font-size: 12px;
    padding: 2px 8px;
}

QTabWidget::pane {
    border: 1px solid #E0E0E0;
    border-radius: 0 5px 5px 5px;
    background-color: #FFFFFF;
}
QTabBar::tab {
    background-color: #EEEEEE;
    border: 1px solid #E0E0E0;
    border-bottom: none;
    border-radius: 5px 5px 0 0;
    padding: 6px 16px;
    margin-right: 2px;
    color: #616161;
}
QTabBar::tab:selected {
    background-color: #FFFFFF;
    color: #1976D2;
    font-weight: bold;
}
QTabBar::tab:hover {
    background-color: #F5F5F5;
}

QSplitter::handle {
    background-color: #E0E0E0;
}
QSplitter::handle:horizontal {
    width: 3px;
}
QSplitter::handle:vertical {
    height: 3px;
}

QScrollBar:vertical {
    border: none;
    background: #F5F5F5;
    width: 8px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #BDBDBD;
    border-radius: 4px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background: #9E9E9E;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}

QLabel#label_title {
    font-size: 16px;
    font-weight: bold;
    color: #1976D2;
    font-family: "Microsoft YaHei", "微软雅黑", Arial, sans-serif;
}
QLabel#label_subtitle {
    font-size: 12px;
    color: #757575;
}
QLabel#label_result_count {
    color: #1976D2;
    font-weight: bold;
}
QLabel#label_no_result {
    color: #9E9E9E;
    font-size: 15px;
}
"""


# ─── 工作线程 ─────────────────────────────────────────────────────────────────

class IndexWorker(QThread):
    """索引构建工作线程（v1.3：带速度统计）"""
    progress = pyqtSignal(int, int, str)   # current, total, filepath
    log_msg = pyqtSignal(str)              # 日志消息
    finished = pyqtSignal(dict)            # 完成，带统计信息
    error = pyqtSignal(str)                # 错误
    speed_update = pyqtSignal(float, float, int, int)  # speed(files/s), eta(s), current, total

    def __init__(self, engine: IndexEngine, root_dir: str,
                 enabled_exts: List[str], enable_pdf: bool,
                 enable_ocr: bool, max_workers: int):
        super().__init__()
        self.engine = engine
        self.root_dir = root_dir
        self.enabled_exts = enabled_exts
        self.enable_pdf = enable_pdf
        self.enable_ocr = enable_ocr
        self.max_workers = max_workers
        self.builder = IndexBuilder(engine)
        self._start_time = 0.0
        self._last_count = 0
        self._last_time = 0.0

    def _on_progress(self, cur: int, tot: int, fp: str):
        """进度回调，计算速度和剩余时间"""
        self.progress.emit(cur, tot, fp)
        now = time.time()
        # 每 0.5 秒更新一次速度
        elapsed_since_last = now - self._last_time
        if elapsed_since_last >= 0.5:
            delta_files = cur - self._last_count
            speed = delta_files / elapsed_since_last if elapsed_since_last > 0 else 0
            remaining = tot - cur
            eta = remaining / speed if speed > 0 else 0
            self.speed_update.emit(speed, eta, cur, tot)
            self._last_count = cur
            self._last_time = now

    def run(self):
        self._start_time = time.time()
        self._last_time = self._start_time
        self._last_count = 0
        try:
            def _speed_cb(spd, eta):
                cur = self._last_count
                tot = 0  # total not needed here, already tracked in progress
                self.speed_update.emit(spd, eta, cur, 0)

            stats = self.builder.build_index(
                root_dir=self.root_dir,
                enabled_extensions=self.enabled_exts,
                enable_pdf=self.enable_pdf,
                enable_ocr=self.enable_ocr,
                max_workers=self.max_workers,
                progress_callback=self._on_progress,
                log_callback=lambda msg: self.log_msg.emit(msg),
                speed_callback=_speed_cb,
            )
            self.finished.emit(stats)
        except Exception as e:
            self.error.emit(str(e))

    def stop(self):
        self.builder.stop()


# ─── 离线预览对话框 ────────────────────────────────────────────────────────────

class OfflinePreviewDialog(QDialog):
    """离线预览对话框：显示关键词命中段落 + 可展开全文"""

    def __init__(self, result: Dict, query: str, parent=None):
        super().__init__(parent)
        self.result = result
        self.query = query
        self.full_content_loaded = False

        self.setWindowTitle(f"离线预览 — {result['filename']}")
        self.setMinimumSize(700, 500)
        self.resize(800, 600)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # 文件信息栏
        info_frame = QFrame()
        info_frame.setStyleSheet(
            "QFrame { background: #E3F2FD; border-radius: 6px; padding: 8px; }"
        )
        info_layout = QVBoxLayout(info_frame)
        info_layout.setSpacing(4)

        filename_label = QLabel(f"📄  {self.result['filename']}")
        filename_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #1565C0;")
        info_layout.addWidget(filename_label)

        path_label = QLabel(f"路径：{self.result['filepath']}")
        path_label.setStyleSheet("color: #546E7A; font-size: 12px;")
        path_label.setWordWrap(True)
        info_layout.addWidget(path_label)

        size_str = self._format_size(self.result.get('filesize', 0))
        mtime_str = self._format_time(self.result.get('modified_time', 0))
        meta_label = QLabel(f"大小：{size_str}    修改时间：{mtime_str}")
        meta_label.setStyleSheet("color: #78909C; font-size: 12px;")
        info_layout.addWidget(meta_label)

        layout.addWidget(info_frame)

        # Tab 切换
        self.tabs = QTabWidget()

        # Tab 1: 关键词命中段落
        snippet_widget = QWidget()
        snippet_layout = QVBoxLayout(snippet_widget)
        snippet_layout.setContentsMargins(8, 8, 8, 8)

        snippets = self.result.get('snippets', [])
        if snippets:
            snippet_label = QLabel(f"找到 {len(snippets)} 处命中（搜索词：{self.query}）")
            snippet_label.setStyleSheet("color: #1976D2; font-weight: bold; margin-bottom: 4px;")
            snippet_layout.addWidget(snippet_label)

            for i, snippet in enumerate(snippets):
                snip_frame = QFrame()
                snip_frame.setStyleSheet(
                    "QFrame { background: #FFFDE7; border: 1px solid #FFF176; "
                    "border-radius: 4px; margin: 2px; }"
                )
                snip_inner = QVBoxLayout(snip_frame)
                snip_inner.setContentsMargins(8, 6, 8, 6)

                num_label = QLabel(f"第 {i+1} 处")
                num_label.setStyleSheet("color: #F57F17; font-size: 11px; font-weight: bold;")
                snip_inner.addWidget(num_label)

                text_label = QLabel(snippet.replace('**', ''))
                text_label.setWordWrap(True)
                text_label.setStyleSheet("color: #212121; line-height: 1.6;")
                snip_inner.addWidget(text_label)

                snippet_layout.addWidget(snip_frame)
        else:
            no_snip = QLabel("未找到关键词命中段落")
            no_snip.setStyleSheet("color: #9E9E9E; font-size: 14px;")
            no_snip.setAlignment(Qt.AlignCenter)
            snippet_layout.addWidget(no_snip)

        snippet_layout.addStretch()

        scroll1 = QScrollArea()
        scroll1.setWidget(snippet_widget)
        scroll1.setWidgetResizable(True)
        scroll1.setFrameShape(QFrame.NoFrame)
        self.tabs.addTab(scroll1, f"关键词命中 ({len(snippets)} 处)")

        # Tab 2: 全文内容
        self.full_text_edit = QTextEdit()
        self.full_text_edit.setReadOnly(True)
        self.full_text_edit.setPlaceholderText("点击此标签页加载全文内容...")
        self.tabs.addTab(self.full_text_edit, "查看全文")
        self.tabs.currentChanged.connect(self._on_tab_changed)

        layout.addWidget(self.tabs)

        # 底部按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        close_btn = QPushButton("关闭")
        close_btn.setObjectName("btn_secondary")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

    def _on_tab_changed(self, index):
        """切换到全文 Tab 时加载内容"""
        if index == 1 and not self.full_content_loaded:
            content = self.result.get('content', '')
            if content:
                self.full_text_edit.setPlainText(content)
                # 高亮关键词
                self._highlight_keywords(self.query)
            else:
                self.full_text_edit.setPlainText("（无法获取全文内容）")
            self.full_content_loaded = True

    def _highlight_keywords(self, query: str):
        """在全文中高亮显示关键词（不依赖 jieba，直接字符串搜索）"""
        if not query:
            return
        fmt = QTextCharFormat()
        fmt.setBackground(QColor('#FFF176'))
        fmt.setForeground(QColor('#E65100'))
        fmt.setFontWeight(QFont.Bold)

        cursor = self.full_text_edit.textCursor()
        cursor.movePosition(QTextCursor.Start)
        self.full_text_edit.setTextCursor(cursor)

        # 直接搜索关键词本身，不需要分词
        # 同时也搜索空格分隔的各个单词
        words = [query] + [w for w in query.split() if w.strip() and w != query]

        doc = self.full_text_edit.document()
        for word in words:
            find_flags = QTextDocument.FindFlags()
            cursor = doc.find(word, 0, find_flags)
            while not cursor.isNull():
                cursor.mergeCharFormat(fmt)
                cursor = doc.find(word, cursor, find_flags)

    def _format_size(self, size: int) -> str:
        if size < 1024:
            return f"{size} B"
        elif size < 1024 * 1024:
            return f"{size/1024:.1f} KB"
        else:
            return f"{size/1024/1024:.1f} MB"

    def _format_time(self, ts: float) -> str:
        if not ts:
            return "未知"
        return datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M')


# ─── 索引设置对话框 ────────────────────────────────────────────────────────────

class IndexSettingsDialog(QDialog):
    """索引设置对话框"""

    def __init__(self, parent=None, current_settings: Dict = None):
        super().__init__(parent)
        self.setWindowTitle("索引设置")
        self.setMinimumWidth(500)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.settings = current_settings or {}
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        # 索引目录
        dir_group = QGroupBox("索引目录")
        dir_layout = QHBoxLayout(dir_group)
        self.dir_edit = QLineEdit(self.settings.get('root_dir', ''))
        self.dir_edit.setPlaceholderText("选择要索引的文件夹...")
        dir_btn = QPushButton("浏览...")
        dir_btn.setObjectName("btn_secondary")
        dir_btn.setMaximumWidth(80)
        dir_btn.clicked.connect(self._browse_dir)
        dir_layout.addWidget(self.dir_edit)
        dir_layout.addWidget(dir_btn)
        layout.addWidget(dir_group)

        # 索引文件保存位置
        save_group = QGroupBox("索引文件保存位置")
        save_layout = QHBoxLayout(save_group)
        self.save_edit = QLineEdit(self.settings.get('db_path', ''))
        self.save_edit.setPlaceholderText("选择索引文件保存路径（.db 文件）...")
        save_btn = QPushButton("浏览...")
        save_btn.setObjectName("btn_secondary")
        save_btn.setMaximumWidth(80)
        save_btn.clicked.connect(self._browse_save)
        save_layout.addWidget(self.save_edit)
        save_layout.addWidget(save_btn)
        layout.addWidget(save_group)

        # 文件格式选择
        fmt_group = QGroupBox("索引文件格式")
        fmt_layout = QGridLayout(fmt_group)
        fmt_layout.setSpacing(8)

        self.ext_checks = {}
        ext_list = [
            ('.docx', 'Word 文档 (.docx)'),
            ('.doc', 'Word 文档 (.doc)'),
            ('.xlsx', 'Excel 表格 (.xlsx)'),
            ('.xls', 'Excel 表格 (.xls)'),
            ('.pptx', 'PowerPoint (.pptx)'),
            ('.ppt', 'PowerPoint (.ppt)'),
            ('.pdf', 'PDF 文档（可识别文字）'),
        ]
        enabled_exts = self.settings.get('enabled_extensions',
                                          ['.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt'])
        for i, (ext, label) in enumerate(ext_list):
            cb = QCheckBox(label)
            cb.setChecked(ext in enabled_exts)
            self.ext_checks[ext] = cb
            fmt_layout.addWidget(cb, i // 2, i % 2)

        layout.addWidget(fmt_group)

        # PDF OCR 选项
        ocr_group = QGroupBox("PDF OCR 设置（扫描版 PDF）")
        ocr_layout = QVBoxLayout(ocr_group)
        self.ocr_check = QCheckBox("启用 OCR 识别扫描版 PDF（速度较慢，需安装 Tesseract）")
        self.ocr_check.setChecked(self.settings.get('enable_ocr', False))
        ocr_layout.addWidget(self.ocr_check)
        ocr_note = QLabel("注意：OCR 会大幅增加索引时间，建议仅在需要时开启")
        ocr_note.setStyleSheet("color: #FF6F00; font-size: 11px;")
        ocr_layout.addWidget(ocr_note)
        layout.addWidget(ocr_group)

        # 线程数
        thread_group = QGroupBox("性能设置")
        thread_layout = QHBoxLayout(thread_group)
        thread_layout.addWidget(QLabel("并行处理线程数："))
        self.thread_spin = QSpinBox()
        self.thread_spin.setRange(1, 16)
        import os as _os
        default_workers = min(_os.cpu_count() or 4, 16)
        self.thread_spin.setValue(self.settings.get('max_workers', default_workers))
        self.thread_spin.setMaximumWidth(80)
        thread_layout.addWidget(self.thread_spin)
        thread_layout.addWidget(QLabel(f"（自动检测：{default_workers} 核，可根据需要调整）"))
        thread_layout.addStretch()
        layout.addWidget(thread_group)

        # 按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._validate_and_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _browse_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择要索引的文件夹",
                                                 self.dir_edit.text() or os.path.expanduser('~'))
        if path:
            self.dir_edit.setText(path)
            # 自动建议索引文件路径
            if not self.save_edit.text():
                suggested = os.path.join(path, 'docfinder_index.db')
                self.save_edit.setText(suggested)

    def _browse_save(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "选择索引文件保存位置",
            self.save_edit.text() or os.path.expanduser('~'),
            "索引文件 (*.db)"
        )
        if path:
            if not path.endswith('.db'):
                path += '.db'
            self.save_edit.setText(path)

    def _validate_and_accept(self):
        if not self.dir_edit.text().strip():
            QMessageBox.warning(self, "提示", "请选择要索引的文件夹")
            return
        if not self.save_edit.text().strip():
            QMessageBox.warning(self, "提示", "请选择索引文件保存位置")
            return
        enabled = [ext for ext, cb in self.ext_checks.items() if cb.isChecked()]
        if not enabled:
            QMessageBox.warning(self, "提示", "请至少选择一种文件格式")
            return
        self.accept()

    def get_settings(self) -> Dict:
        return {
            'root_dir': self.dir_edit.text().strip(),
            'db_path': self.save_edit.text().strip(),
            'enabled_extensions': [ext for ext, cb in self.ext_checks.items() if cb.isChecked()],
            'enable_pdf': '.pdf' in [ext for ext, cb in self.ext_checks.items() if cb.isChecked()],
            'enable_ocr': self.ocr_check.isChecked(),
            'max_workers': self.thread_spin.value(),
        }


# ─── 主窗口 ───────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.engine: Optional[IndexEngine] = None
        self.index_worker: Optional[IndexWorker] = None
        self.current_results: List[Dict] = []
        self.current_query = ''
        self.qsettings = QSettings('DocFinder', 'DocFinder')

        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(900, 650)
        self.resize(1100, 750)

        self._build_ui()
        self._build_menu()
        self._restore_state()
        self._update_status("就绪")

    def _build_ui(self):
        """构建主界面"""
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # 顶部搜索区域
        search_panel = self._build_search_panel()
        main_layout.addWidget(search_panel)

        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color: #E0E0E0;")
        main_layout.addWidget(line)

        # 主内容区（左：结果列表，右：详情）
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(4)

        # 左侧结果列表
        left_panel = self._build_result_panel()
        splitter.addWidget(left_panel)

        # 右侧详情面板
        right_panel = self._build_detail_panel()
        splitter.addWidget(right_panel)

        splitter.setSizes([500, 400])
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        main_layout.addWidget(splitter, 1)

        # 底部进度区域
        self.progress_panel = self._build_progress_panel()
        self.progress_panel.hide()
        main_layout.addWidget(self.progress_panel)

        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

    def _build_search_panel(self) -> QWidget:
        """构建顶部搜索面板"""
        panel = QWidget()
        panel.setStyleSheet("background-color: #FFFFFF; padding: 12px;")
        panel.setMaximumHeight(100)
        layout = QVBoxLayout(panel)
        layout.setSpacing(8)
        layout.setContentsMargins(16, 12, 16, 12)

        # 标题行
        title_row = QHBoxLayout()
        title_row.setSpacing(8)

        # Logo 图标
        # 标题图标：用 QPainter 在内存中直接绘制，不依赖外部文件
        logo_label = QLabel()
        # 先尝试加载外部图标文件
        _assets_candidates = []
        if getattr(sys, 'frozen', False):
            _assets_candidates.append(os.path.join(sys._MEIPASS, 'assets'))
            _assets_candidates.append(os.path.join(os.path.dirname(sys.executable), 'assets'))
        _assets_candidates.append(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'assets'))
        assets_dir = ''
        for _d in _assets_candidates:
            if os.path.isdir(_d):
                assets_dir = _d
                break
        logo_path = os.path.join(assets_dir, 'logo_icon.png') if assets_dir else ''
        _logo_loaded = False
        if logo_path and os.path.exists(logo_path):
            pix = QPixmap(logo_path).scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            if not pix.isNull():
                logo_label.setPixmap(pix)
                _logo_loaded = True
        if not _logo_loaded:
            # 用 QPainter 绘制一个简单的搜索图标
            from PyQt5.QtGui import QPainter, QColor, QPen, QBrush
            _icon_pix = QPixmap(32, 32)
            _icon_pix.fill(Qt.transparent)
            _p = QPainter(_icon_pix)
            _p.setRenderHint(QPainter.Antialiasing)
            _p.setBrush(QBrush(QColor('#1565C0')))
            _p.setPen(Qt.NoPen)
            _p.drawRoundedRect(0, 0, 32, 32, 6, 6)
            _p.setPen(QPen(QColor('white'), 2))
            _p.drawLine(8, 10, 24, 10)
            _p.drawLine(8, 16, 20, 16)
            _p.drawLine(8, 22, 22, 22)
            _p.end()
            logo_label.setPixmap(_icon_pix)
        logo_label.setFixedSize(36, 36)
        title_row.addWidget(logo_label)

        # 标题文字：直接用 QLabel 显示文字，不依赖图片文件
        # 这是最可靠的方式，Qt 会自动使用系统字体
        title_label = QLabel(APP_NAME)
        title_label.setObjectName("label_title")
        title_label.setStyleSheet(
            "color: #1565C0; font-size: 20px; font-weight: bold;"
            "font-family: '微软雅黑', 'Microsoft YaHei', 'SimHei', '黑体', sans-serif;"
        )
        title_row.addWidget(title_label)
        title_row.addStretch()

        # 索引信息标签
        self.index_info_label = QLabel("未加载索引")
        self.index_info_label.setStyleSheet(
            "color: #9E9E9E; font-size: 12px; "
            "background: #F5F5F5; border-radius: 4px; padding: 3px 8px;"
        )
        title_row.addWidget(self.index_info_label)
        layout.addLayout(title_row)

        # 搜索框行
        search_row = QHBoxLayout()
        search_row.setSpacing(8)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入关键词搜索文档内容...（支持中英文，多词用空格分隔）")
        self.search_input.setMinimumHeight(36)
        self.search_input.returnPressed.connect(self._do_search)
        search_row.addWidget(self.search_input, 1)

        self.search_btn = QPushButton("搜索")
        self.search_btn.setMinimumWidth(80)
        self.search_btn.setMinimumHeight(36)
        self.search_btn.clicked.connect(self._do_search)
        search_row.addWidget(self.search_btn)

        self.clear_btn = QPushButton("清空")
        self.clear_btn.setObjectName("btn_secondary")
        self.clear_btn.setMinimumWidth(60)
        self.clear_btn.setMinimumHeight(36)
        self.clear_btn.clicked.connect(self._clear_search)
        search_row.addWidget(self.clear_btn)

        layout.addLayout(search_row)
        return panel

    def _build_result_panel(self) -> QWidget:
        """构建左侧结果列表面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(6)
        layout.setContentsMargins(12, 10, 6, 10)

        # 结果数量标签
        result_header = QHBoxLayout()
        self.result_count_label = QLabel("搜索结果")
        self.result_count_label.setObjectName("label_result_count")
        result_header.addWidget(self.result_count_label)
        result_header.addStretch()
        layout.addLayout(result_header)

        # 结果树形列表
        self.result_tree = QTreeWidget()
        self.result_tree.setColumnCount(4)
        self.result_tree.setHeaderLabels(['文件名', '类型', '大小', '修改时间'])
        self.result_tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.result_tree.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.result_tree.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.result_tree.header().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.result_tree.setAlternatingRowColors(True)
        self.result_tree.setSelectionMode(QAbstractItemView.SingleSelection)
        self.result_tree.itemClicked.connect(self._on_result_clicked)
        self.result_tree.itemDoubleClicked.connect(self._on_result_double_clicked)
        self.result_tree.setToolTip("单击查看摘要，双击打开文件（文件不存在时显示离线预览）")
        # 开启右键菜单
        self.result_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.result_tree.customContextMenuRequested.connect(self._show_result_context_menu)
        layout.addWidget(self.result_tree, 1)

        # 空状态提示
        self.empty_label = QLabel("请输入关键词开始搜索")
        self.empty_label.setObjectName("label_no_result")
        self.empty_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.empty_label)

        return panel

    def _build_detail_panel(self) -> QWidget:
        """构建右侧详情面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(6)
        layout.setContentsMargins(6, 10, 12, 10)

        detail_label = QLabel("文档详情")
        detail_label.setObjectName("label_result_count")
        layout.addWidget(detail_label)

        self.detail_tabs = QTabWidget()

        # Tab 1: 摘要
        snippet_widget = QWidget()
        snippet_layout = QVBoxLayout(snippet_widget)
        snippet_layout.setContentsMargins(8, 8, 8, 8)

        self.snippet_text = QTextEdit()
        self.snippet_text.setReadOnly(True)
        self.snippet_text.setPlaceholderText("点击左侧搜索结果查看关键词命中摘要...")
        snippet_layout.addWidget(self.snippet_text)

        self.detail_tabs.addTab(snippet_widget, "关键词摘要")

        # Tab 2: 文件信息
        info_widget = QWidget()
        info_layout = QVBoxLayout(info_widget)
        info_layout.setContentsMargins(8, 8, 8, 8)
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        info_layout.addWidget(self.info_text)
        self.detail_tabs.addTab(info_widget, "文件信息")

        layout.addWidget(self.detail_tabs, 1)

        # 操作按钮
        btn_layout = QHBoxLayout()
        self.open_btn = QPushButton("打开文件")
        self.open_btn.setObjectName("btn_success")
        self.open_btn.clicked.connect(self._open_selected_file)
        self.open_btn.setEnabled(False)
        btn_layout.addWidget(self.open_btn)

        self.open_folder_btn = QPushButton("打开所在文件夹")
        self.open_folder_btn.setObjectName("btn_secondary")
        self.open_folder_btn.clicked.connect(self._open_selected_folder)
        self.open_folder_btn.setEnabled(False)
        btn_layout.addWidget(self.open_folder_btn)

        self.offline_btn = QPushButton("离线预览")
        self.offline_btn.setObjectName("btn_secondary")
        self.offline_btn.clicked.connect(self._offline_preview_selected)
        self.offline_btn.setEnabled(False)
        btn_layout.addWidget(self.offline_btn)

        layout.addLayout(btn_layout)
        return panel

    def _build_progress_panel(self) -> QWidget:
        """构建底部进度面板（v1.3：详细信息版）"""
        panel = QWidget()
        panel.setStyleSheet(
            "QWidget { background: #E8F4FD; border-top: 2px solid #1976D2; }"
        )
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 10, 16, 10)
        layout.setSpacing(6)

        # 第一行：标题 + 停止按钮
        header_row = QHBoxLayout()
        self.progress_title = QLabel("正在建立索引...")
        self.progress_title.setStyleSheet(
            "font-weight: bold; color: #1565C0; font-size: 13px;"
        )
        header_row.addWidget(self.progress_title)
        header_row.addStretch()

        self.stop_btn = QPushButton("⏹ 停止索引")
        self.stop_btn.setObjectName("btn_danger")
        self.stop_btn.setMinimumWidth(100)
        self.stop_btn.setMinimumHeight(32)
        self.stop_btn.clicked.connect(self._stop_indexing)
        header_row.addWidget(self.stop_btn)
        layout.addLayout(header_row)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumHeight(12)
        self.progress_bar.setStyleSheet(
            "QProgressBar { border-radius: 6px; background: #BBDEFB; }"
            "QProgressBar::chunk { background: #1976D2; border-radius: 6px; }"
        )
        layout.addWidget(self.progress_bar)

        # 第二行：统计数字（4列）
        stats_row = QHBoxLayout()
        stats_row.setSpacing(20)

        # 已处理 / 总数
        self.stat_count_label = QLabel("已处理：0 / 0")
        self.stat_count_label.setStyleSheet(
            "color: #1565C0; font-size: 12px; font-weight: bold;"
        )
        stats_row.addWidget(self.stat_count_label)

        # 索引速度
        self.stat_speed_label = QLabel("速度：-- 文件/秒")
        self.stat_speed_label.setStyleSheet(
            "color: #2E7D32; font-size: 12px; font-weight: bold;"
        )
        stats_row.addWidget(self.stat_speed_label)

        # 剩余时间
        self.stat_eta_label = QLabel("剩余：计算中...")
        self.stat_eta_label.setStyleSheet(
            "color: #E65100; font-size: 12px; font-weight: bold;"
        )
        stats_row.addWidget(self.stat_eta_label)

        # 已用时间
        self.stat_elapsed_label = QLabel("已用：0 秒")
        self.stat_elapsed_label.setStyleSheet(
            "color: #546E7A; font-size: 12px;"
        )
        stats_row.addWidget(self.stat_elapsed_label)

        stats_row.addStretch()
        layout.addLayout(stats_row)

        # 第三行：当前文件
        self.progress_label = QLabel("准备扫描文件...")
        self.progress_label.setStyleSheet(
            "color: #37474F; font-size: 11px; padding: 2px 0;"
        )
        self.progress_label.setWordWrap(False)
        layout.addWidget(self.progress_label)

        # 第四行：日志区（可折叠，默认显示最近3条）
        self.progress_log = QTextEdit()
        self.progress_log.setReadOnly(True)
        self.progress_log.setMaximumHeight(60)
        self.progress_log.setStyleSheet(
            "QTextEdit { background: #F0F8FF; border: 1px solid #BBDEFB; "
            "border-radius: 4px; font-size: 11px; color: #37474F; }"
        )
        layout.addWidget(self.progress_log)

        # 计时器：每秒更新已用时间
        self._elapsed_timer = QTimer(self)
        self._elapsed_timer.setInterval(1000)
        self._elapsed_timer.timeout.connect(self._update_elapsed_time)
        self._index_start_time = 0.0

        return panel

    def _build_menu(self):
        """构建菜单栏"""
        menubar = self.menuBar()

        # 索引菜单
        index_menu = menubar.addMenu("索引")

        new_index_action = QAction("新建索引...", self)
        new_index_action.setShortcut("Ctrl+N")
        new_index_action.triggered.connect(self._new_index)
        index_menu.addAction(new_index_action)

        open_index_action = QAction("加载索引...", self)
        open_index_action.setShortcut("Ctrl+O")
        open_index_action.triggered.connect(self._load_index)
        index_menu.addAction(open_index_action)

        index_menu.addSeparator()

        update_index_action = QAction("更新索引", self)
        update_index_action.setShortcut("F5")
        update_index_action.triggered.connect(self._update_index)
        index_menu.addAction(update_index_action)

        cleanup_action = QAction("清理孤立记录", self)
        cleanup_action.setToolTip("删除已不存在的文件的索引记录")
        cleanup_action.triggered.connect(self._cleanup_orphans)
        index_menu.addAction(cleanup_action)

        index_menu.addSeparator()

        exit_action = QAction("退出", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        index_menu.addAction(exit_action)

        # 帮助菜单
        help_menu = menubar.addMenu("帮助")
        about_action = QAction("关于", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    # ─── 索引操作 ─────────────────────────────────────────────────────────────

    def _cleanup_orphans(self):
        """手动清理孤立记录（删除已不存在的文件的索引记录）"""
        if not self.engine:
            QMessageBox.information(self, "提示", "请先加载一个索引")
            return
        try:
            deleted = self.engine.cleanup_orphans()
            info = self.engine.get_index_info()
            total = info.get('total_docs', 0)
            QMessageBox.information(
                self, "清理完成",
                f"共删除 {deleted} 条孤立记录（文件已不存在）\n"
                f"当前索引共 {total} 个文档"
            )
            self._refresh_index_info()
        except Exception as e:
            QMessageBox.critical(self, "清理失败", f"清理孤立记录时出错：\n{e}")

    def _new_index(self):
        """新建索引"""
        last_settings = {
            'root_dir': self.qsettings.value('last_root_dir', ''),
            'db_path': self.qsettings.value('last_db_path', ''),
            'enabled_extensions': self.qsettings.value(
                'enabled_extensions',
                ['.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt']
            ),
            'enable_ocr': self.qsettings.value('enable_ocr', False, type=bool),
            'max_workers': self.qsettings.value('max_workers', min(multiprocessing.cpu_count(), 16), type=int),
        }

        dlg = IndexSettingsDialog(self, last_settings)
        if dlg.exec_() != QDialog.Accepted:
            return

        settings = dlg.get_settings()
        self._save_index_settings(settings)
        self._start_indexing(settings)

    def _load_index(self):
        """加载已有索引文件"""
        path, _ = QFileDialog.getOpenFileName(
            self, "加载索引文件",
            self.qsettings.value('last_db_path', os.path.expanduser('~')),
            "索引文件 (*.db)"
        )
        if not path:
            return

        try:
            self.engine = IndexEngine(path)
            info = self.engine.get_index_info()
            total = info.get('total_docs', 0)
            root = info.get('index_root', '未知')
            last_time = info.get('last_index_time', '')
            if last_time:
                last_time = datetime.fromtimestamp(float(last_time)).strftime('%Y-%m-%d %H:%M')
            else:
                last_time = '未知'

            self.index_info_label.setText(
                f"已加载：{total} 个文档  |  上次索引：{last_time}"
            )
            self.index_info_label.setStyleSheet(
                "color: #2E7D32; font-size: 12px; font-weight: bold; "
                "background: #E8F5E9; border-radius: 4px; padding: 3px 8px;"
            )
            self.qsettings.setValue('last_db_path', path)
            self._update_status(f"已加载索引：{total} 个文档，来自 {root}")

            # 清空搜索结果
            self._clear_search()

        except Exception as e:
            QMessageBox.critical(self, "加载失败", f"无法加载索引文件：\n{e}")

    def _update_index(self):
        """更新当前索引"""
        if not self.engine:
            QMessageBox.information(self, "提示", "请先新建或加载一个索引")
            return

        info = self.engine.get_index_info()
        root_dir = info.get('index_root', '')

        if not root_dir or not os.path.exists(root_dir):
            reply = QMessageBox.question(
                self, "目录不存在",
                f"原索引目录不存在或未记录：\n{root_dir}\n\n是否重新选择目录？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self._new_index()
            return

        # 更新前先清理孤立记录（删除已不存在的文件的索引）
        try:
            deleted = self.engine.cleanup_orphans()
            if deleted > 0:
                self._update_status(f"已清理 {deleted} 条孤立记录")
        except Exception:
            pass

        # 使用上次的设置
        settings = {
            'root_dir': root_dir,
            'db_path': self.engine.db_path,
            'enabled_extensions': self.qsettings.value(
                'enabled_extensions',
                ['.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt']
            ),
            'enable_pdf': '.pdf' in self.qsettings.value('enabled_extensions', []),
            'enable_ocr': self.qsettings.value('enable_ocr', False, type=bool),
            'max_workers': self.qsettings.value('max_workers', 4, type=int),
        }
        self._start_indexing(settings)

    def _start_indexing(self, settings: Dict):
        """开始索引任务"""
        db_path = settings['db_path']
        if not db_path:
            QMessageBox.warning(self, "错误", "索引文件路径不能为空")
            return

        # 确保目录存在（处理只有文件名没有目录的情况）
        db_path = os.path.abspath(db_path)
        settings['db_path'] = db_path
        db_dir = os.path.dirname(db_path)
        if db_dir:
            try:
                os.makedirs(db_dir, exist_ok=True)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法创建目录：{db_dir}\n{e}")
                return

        # 初始化引擎
        self.engine = IndexEngine(db_path)

        # 启动工作线程
        self.index_worker = IndexWorker(
            engine=self.engine,
            root_dir=settings['root_dir'],
            enabled_exts=settings['enabled_extensions'],
            enable_pdf=settings.get('enable_pdf', True),
            enable_ocr=settings.get('enable_ocr', False),
            max_workers=settings.get('max_workers', min(multiprocessing.cpu_count(), 16)),
        )
        self.index_worker.progress.connect(self._on_index_progress)
        self.index_worker.log_msg.connect(self._on_index_log)
        self.index_worker.finished.connect(self._on_index_finished)
        self.index_worker.error.connect(self._on_index_error)
        self.index_worker.speed_update.connect(self._on_speed_update)
        self.index_worker.start()

        # 显示进度面板，重置所有状态
        self.progress_panel.show()
        self.progress_bar.setValue(0)
        self.progress_log.clear()
        self.stat_count_label.setText("已处理：0 / 0")
        self.stat_speed_label.setText("速度：-- 文件/秒")
        self.stat_eta_label.setText("剩余：计算中...")
        self.stat_elapsed_label.setText("已用：0 秒")
        self.progress_label.setText("正在扫描目录...")
        self.progress_title.setText(f"正在建立索引：{settings['root_dir']}")
        self.stop_btn.setEnabled(True)
        self.search_btn.setEnabled(False)
        # 启动计时器
        import time as _time
        self._index_start_time = _time.time()
        self._elapsed_timer.start()
        self._update_status("索引中...")

    def _stop_indexing(self):
        """停止索引（立即生效）"""
        if self.index_worker and self.index_worker.isRunning():
            self.progress_title.setText("正在停止，请稍候...")
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("停止中...")
            # 调用 IndexWorker.stop()，内部会调用 IndexBuilder.stop()
            # IndexBuilder.stop() 会设置 _stop_flag 并关闭线程池
            self.index_worker.stop()
            # 给 UI 一个反馈，让用户知道已经发出停止信号
            self._update_status("停止信号已发出，正在等待当前文件处理完成...")

    def _on_index_progress(self, current: int, total: int, filepath: str):
        """索引进度更新"""
        if total > 0:
            pct = int(current * 100 / total)
            self.progress_bar.setValue(pct)
            filename = os.path.basename(filepath)
            # 显示当前文件（路径过长则截断显示）
            if len(filepath) > 80:
                display_path = '...' + filepath[-77:]
            else:
                display_path = filepath
            self.progress_label.setText(f"当前：{display_path}")
            self.stat_count_label.setText(f"已处理：{current} / {total}")

    def _on_speed_update(self, speed: float, eta: float, current: int, total: int):
        """速度和剩余时间更新"""
        if speed > 0:
            self.stat_speed_label.setText(f"速度：{speed:.1f} 文件/秒")
        if eta > 0:
            if eta < 60:
                eta_str = f"{int(eta)} 秒"
            elif eta < 3600:
                eta_str = f"{int(eta // 60)} 分 {int(eta % 60)} 秒"
            else:
                h = int(eta // 3600)
                m = int((eta % 3600) // 60)
                eta_str = f"{h} 小时 {m} 分"
            self.stat_eta_label.setText(f"剩余：{eta_str}")
        else:
            self.stat_eta_label.setText("剩余：即将完成")

    def _update_elapsed_time(self):
        """每秒更新已用时间"""
        import time as _time
        if self._index_start_time > 0:
            elapsed = _time.time() - self._index_start_time
            if elapsed < 60:
                elapsed_str = f"{int(elapsed)} 秒"
            elif elapsed < 3600:
                elapsed_str = f"{int(elapsed // 60)} 分 {int(elapsed % 60)} 秒"
            else:
                h = int(elapsed // 3600)
                m = int((elapsed % 3600) // 60)
                elapsed_str = f"{h} 小时 {m} 分"
            self.stat_elapsed_label.setText(f"已用：{elapsed_str}")

    def _on_index_log(self, msg: str):
        """索引日志：同时显示在日志区和状态栏"""
        self._update_status(msg)
        # 日志区显示最新消息
        self.progress_log.append(msg)
        # 自动滚动到底部
        cursor = self.progress_log.textCursor()
        cursor.movePosition(cursor.End)
        self.progress_log.setTextCursor(cursor)

    def _on_index_finished(self, stats: Dict):
        """索引完成"""
        self._elapsed_timer.stop()
        import time as _time
        total_elapsed = _time.time() - self._index_start_time if self._index_start_time > 0 else 0
        self.progress_panel.hide()
        self.search_btn.setEnabled(True)

        info = self.engine.get_index_info()
        total = info.get('total_docs', 0)
        last_time = datetime.now().strftime('%Y-%m-%d %H:%M')

        self.index_info_label.setText(
            f"已索引：{total} 个文档  |  更新时间：{last_time}"
        )
        self.index_info_label.setStyleSheet(
            "color: #2E7D32; font-size: 12px; font-weight: bold; "
            "background: #E8F5E9; border-radius: 4px; padding: 3px 8px;"
        )

        # 格式化总耗时
        if total_elapsed < 60:
            elapsed_str = f"{total_elapsed:.1f} 秒"
        elif total_elapsed < 3600:
            elapsed_str = f"{int(total_elapsed // 60)} 分 {int(total_elapsed % 60)} 秒"
        else:
            h = int(total_elapsed // 3600)
            m = int((total_elapsed % 3600) // 60)
            elapsed_str = f"{h} 小时 {m} 分"

        avg_speed = stats['total'] / total_elapsed if total_elapsed > 0 else 0
        msg = (f"索引完成！共 {stats['total']} 个文件\n"
               f"新增 {stats['added']} 个，更新 {stats['updated']} 个，"
               f"跳过 {stats['skipped']} 个，失败 {stats['failed']} 个\n"
               f"总耗时：{elapsed_str}，平均速度：{avg_speed:.1f} 文件/秒")
        self._update_status(msg.replace('\n', ' '))
        QMessageBox.information(self, "索引完成", msg)

    def _on_index_error(self, error: str):
        """索引出错"""
        self._elapsed_timer.stop()
        self.progress_panel.hide()
        self.search_btn.setEnabled(True)
        QMessageBox.critical(self, "索引出错", f"索引过程中发生错误：\n{error}")
        self._update_status(f"索引出错：{error}")

    def _save_index_settings(self, settings: Dict):
        """保存索引设置"""
        self.qsettings.setValue('last_root_dir', settings.get('root_dir', ''))
        self.qsettings.setValue('last_db_path', settings.get('db_path', ''))
        self.qsettings.setValue('enabled_extensions', settings.get('enabled_extensions', []))
        self.qsettings.setValue('enable_ocr', settings.get('enable_ocr', False))
        self.qsettings.setValue('max_workers', settings.get('max_workers', min(multiprocessing.cpu_count(), 16)))

    # ─── 搜索操作 ─────────────────────────────────────────────────────────────

    def _do_search(self):
        """执行搜索"""
        query = self.search_input.text().strip()
        if not query:
            return

        if not self.engine:
            QMessageBox.information(self, "提示", "请先新建或加载一个索引文件")
            return

        self.current_query = query
        self._update_status(f"正在搜索：{query}")

        try:
            results = self.engine.search(query, limit=500)
            self.current_results = results
            self._display_results(results)
            self._update_status(
                f"搜索完成，找到 {len(results)} 个结果（关键词：{query}）"
            )
        except Exception as e:
            QMessageBox.critical(self, "搜索出错", f"搜索时发生错误：\n{e}")
            self._update_status(f"搜索出错：{e}")

    def _clear_search(self):
        """清空搜索"""
        self.search_input.clear()
        self.result_tree.clear()
        self.snippet_text.clear()
        self.info_text.clear()
        self.current_results = []
        self.current_query = ''
        self.result_count_label.setText("搜索结果")
        self.empty_label.show()
        self.open_btn.setEnabled(False)
        self.open_folder_btn.setEnabled(False)
        self.offline_btn.setEnabled(False)

    def _display_results(self, results: List[Dict]):
        """在列表中显示搜索结果"""
        self.result_tree.clear()
        self.snippet_text.clear()
        self.info_text.clear()

        if not results:
            self.empty_label.show()
            self.result_count_label.setText("搜索结果（0）")
            self.open_btn.setEnabled(False)
            self.open_folder_btn.setEnabled(False)
            self.offline_btn.setEnabled(False)
            return

        self.empty_label.hide()
        self.result_count_label.setText(f"搜索结果（{len(results)} 个）")

        for result in results:
            item = QTreeWidgetItem()
            item.setText(0, result['filename'])
            item.setText(1, result['extension'].upper().lstrip('.'))
            item.setText(2, self._format_size(result.get('filesize', 0)))
            item.setText(3, self._format_time(result.get('modified_time', 0)))

            # 文件不存在时显示灰色
            if not result['file_exists']:
                for col in range(4):
                    item.setForeground(col, QColor('#9E9E9E'))
                item.setToolTip(0, f"文件不存在（离线）：{result['filepath']}")
            else:
                item.setToolTip(0, result['filepath'])

            item.setData(0, Qt.UserRole, result)
            self.result_tree.addTopLevelItem(item)

    def _on_result_clicked(self, item: QTreeWidgetItem, column: int):
        """点击结果项：显示摘要"""
        result = item.data(0, Qt.UserRole)
        if not result:
            return

        self._show_snippets(result)
        self._show_file_info(result)

        file_exists = result.get('file_exists', False)
        self.open_btn.setEnabled(file_exists)
        self.open_folder_btn.setEnabled(file_exists)
        self.offline_btn.setEnabled(True)

    def _on_result_double_clicked(self, item: QTreeWidgetItem, column: int):
        """双击结果项：打开文件或离线预览"""
        result = item.data(0, Qt.UserRole)
        if not result:
            return

        if result.get('file_exists', False):
            self._open_file(result['filepath'])
        else:
            self._show_offline_preview(result)

    def _show_snippets(self, result: Dict):
        """在右侧显示关键词命中摘要"""
        self.snippet_text.clear()
        snippets = result.get('snippets', [])

        if not snippets:
            self.snippet_text.setPlainText("未找到关键词命中段落")
            return

        html_parts = [
            f'<div style="font-family: Microsoft YaHei, Arial; font-size: 13px; '
            f'line-height: 1.7; padding: 4px;">'
        ]
        for i, snippet in enumerate(snippets):
            # 将 **word** 转换为 HTML 高亮
            import re
            highlighted = re.sub(
                r'\*\*(.+?)\*\*',
                r'<span style="background:#FFF176; color:#E65100; '
                r'font-weight:bold; padding:1px 2px; border-radius:2px;">\1</span>',
                snippet
            )
            html_parts.append(
                f'<div style="margin-bottom: 10px; padding: 8px 10px; '
                f'background: #FFFDE7; border-left: 3px solid #FFC107; '
                f'border-radius: 0 4px 4px 0;">'
                f'<span style="color:#F57F17; font-size:11px; font-weight:bold;">'
                f'第 {i+1} 处命中</span><br>'
                f'{highlighted}</div>'
            )
        html_parts.append('</div>')
        self.snippet_text.setHtml(''.join(html_parts))

    def _show_file_info(self, result: Dict):
        """在右侧显示文件信息"""
        self.info_text.clear()
        status = "✅ 文件存在" if result.get('file_exists') else "⚠️ 文件不存在（离线）"
        size_str = self._format_size(result.get('filesize', 0))
        mtime_str = self._format_time(result.get('modified_time', 0))
        content_len = len(result.get('content', ''))

        html = f'''
        <div style="font-family: Microsoft YaHei, Arial; font-size: 13px; 
                    line-height: 1.8; padding: 8px;">
            <table style="width:100%; border-collapse: collapse;">
                <tr>
                    <td style="color:#757575; width:90px; padding:4px 0;">文件名</td>
                    <td style="font-weight:bold; color:#212121;">{result['filename']}</td>
                </tr>
                <tr>
                    <td style="color:#757575; padding:4px 0;">完整路径</td>
                    <td style="color:#546E7A; word-break:break-all;">{result['filepath']}</td>
                </tr>
                <tr>
                    <td style="color:#757575; padding:4px 0;">文件格式</td>
                    <td>{result['extension'].upper().lstrip('.')}</td>
                </tr>
                <tr>
                    <td style="color:#757575; padding:4px 0;">文件大小</td>
                    <td>{size_str}</td>
                </tr>
                <tr>
                    <td style="color:#757575; padding:4px 0;">修改时间</td>
                    <td>{mtime_str}</td>
                </tr>
                <tr>
                    <td style="color:#757575; padding:4px 0;">索引文字</td>
                    <td>{content_len:,} 字符</td>
                </tr>
                <tr>
                    <td style="color:#757575; padding:4px 0;">文件状态</td>
                    <td>{status}</td>
                </tr>
            </table>
        </div>
        '''
        self.info_text.setHtml(html)

    # ─── 文件操作 ─────────────────────────────────────────────────────────────

    def _get_selected_result(self) -> Optional[Dict]:
        items = self.result_tree.selectedItems()
        if not items:
            return None
        return items[0].data(0, Qt.UserRole)

    def _open_selected_file(self):
        result = self._get_selected_result()
        if result and result.get('file_exists'):
            self._open_file(result['filepath'])

    def _open_selected_folder(self):
        result = self._get_selected_result()
        if result and result.get('file_exists'):
            folder = os.path.dirname(result['filepath'])
            self._open_folder(folder)

    def _offline_preview_selected(self):
        result = self._get_selected_result()
        if result:
            self._show_offline_preview(result)

    def _open_file(self, filepath: str):
        """用系统默认程序打开文件"""
        try:
            if sys.platform == 'win32':
                os.startfile(filepath)
            elif sys.platform == 'darwin':
                subprocess.run(['open', filepath])
            else:
                subprocess.run(['xdg-open', filepath])
        except Exception as e:
            QMessageBox.warning(self, "打开失败", f"无法打开文件：\n{e}")

    def _open_folder(self, folder: str):
        """打开文件夹"""
        try:
            if sys.platform == 'win32':
                os.startfile(folder)
            elif sys.platform == 'darwin':
                subprocess.run(['open', folder])
            else:
                subprocess.run(['xdg-open', folder])
        except Exception as e:
            QMessageBox.warning(self, "打开失败", f"无法打开文件夹：\n{e}")

    def _show_offline_preview(self, result: Dict):
        """显示离线预览对话框"""
        dlg = OfflinePreviewDialog(result, self.current_query, self)
        dlg.exec_()

    # ─── 工具方法 ─────────────────────────────────────────────────────────────

    def _update_status(self, msg: str):
        self.status_bar.showMessage(msg)

    def _format_size(self, size: int) -> str:
        if size < 1024:
            return f"{size} B"
        elif size < 1024 * 1024:
            return f"{size/1024:.1f} KB"
        else:
            return f"{size/1024/1024:.1f} MB"

    def _format_time(self, ts: float) -> str:
        if not ts:
            return "—"
        return datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M')

    def _restore_state(self):
        """恢复上次窗口状态"""
        geometry = self.qsettings.value('geometry')
        if geometry:
            self.restoreGeometry(geometry)

        # 尝试自动加载上次的索引
        last_db = self.qsettings.value('last_db_path', '')
        if last_db and os.path.exists(last_db):
            try:
                self.engine = IndexEngine(last_db)
                info = self.engine.get_index_info()
                total = info.get('total_docs', 0)
                last_time = info.get('last_index_time', '')
                if last_time:
                    last_time = datetime.fromtimestamp(float(last_time)).strftime('%Y-%m-%d %H:%M')
                else:
                    last_time = '未知'
                self.index_info_label.setText(
                    f"已加载：{total} 个文档  |  上次索引：{last_time}"
                )
                self.index_info_label.setStyleSheet(
                    "color: #2E7D32; font-size: 12px; font-weight: bold; "
                    "background: #E8F5E9; border-radius: 4px; padding: 3px 8px;"
                )
                self._update_status(f"自动加载上次索引：{total} 个文档")
            except Exception:
                pass

    def _show_result_context_menu(self, pos):
        """显示搜索结果右键菜单"""
        item = self.result_tree.itemAt(pos)
        if not item:
            return
        result = item.data(0, Qt.UserRole)
        if not result:
            return

        from PyQt5.QtWidgets import QMenu
        menu = QMenu(self)

        file_exists = result.get('file_exists', False)

        open_action = menu.addAction("打开文件")
        open_action.setEnabled(file_exists)
        open_action.triggered.connect(lambda: self._open_file(result['filepath']))

        folder_action = menu.addAction("打开所在文件夹")
        folder_action.setEnabled(file_exists)
        folder_action.triggered.connect(
            lambda: self._open_folder(os.path.dirname(result['filepath']))
        )

        menu.addSeparator()

        preview_action = menu.addAction("离线预览")
        preview_action.triggered.connect(lambda: self._show_offline_preview(result))

        menu.addSeparator()

        copy_path_action = menu.addAction("复制完整路径")
        copy_path_action.triggered.connect(
            lambda: QApplication.clipboard().setText(result['filepath'])
        )

        copy_name_action = menu.addAction("复制文件名")
        copy_name_action.triggered.connect(
            lambda: QApplication.clipboard().setText(result['filename'])
        )

        menu.exec_(self.result_tree.viewport().mapToGlobal(pos))

    def _show_about(self):
        QMessageBox.about(
            self, f"关于 {APP_NAME}",
            f"<b>{APP_NAME}</b> v{APP_VERSION}<br><br>"
            "一款快速的文档全文搜索索引工具<br>"
            "支持 Word、Excel、PowerPoint、PDF 格式<br>"
            "使用 SQLite FTS5 全文索引，支持中英文搜索<br><br>"
            "单击结果：查看关键词命中摘要<br>"
            "双击结果：文件存在则打开，不存在则离线预览<br>"
            "右键结果：复制路径、打开文件夹等快捷操作"
        )

    def closeEvent(self, event):
        """关闭时保存状态"""
        self.qsettings.setValue('geometry', self.saveGeometry())
        if self.engine:
            self.engine.close()
        if self.index_worker and self.index_worker.isRunning():
            self.index_worker.stop()
            self.index_worker.wait(3000)
        event.accept()
