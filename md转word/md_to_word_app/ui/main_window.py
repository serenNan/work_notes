# -*- coding: utf-8 -*-
"""
主窗口 UI - Material Design 左右分栏布局
支持浅色/深色主题切换
"""

import os
import subprocess
import sys
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QSplitter, QApplication
)
from PyQt5.QtCore import Qt

from .styles import get_theme_manager, generate_stylesheet
from .styles.colors import WINDOW_SIZES
from .styles.scaling import scaled_size
from .components.sidebar import Sidebar
from .components.header_bar import HeaderBar
from .components.status_bar import StatusBar
from .pages import MdToWordPage, WordToMdPage, SettingsPage
from .pages.history_view import HistoryView
from core.settings_manager import get_settings_manager


class MainWindow(QMainWindow):
    """
    主窗口 - Material Design 风格
    左右分栏布局: Sidebar | HeaderBar + ContentArea + StatusBar
    """

    VERSION = "v2.0.0"

    def __init__(self):
        super().__init__()
        self._theme = get_theme_manager()
        self._settings = get_settings_manager()
        self._theme.theme_changed.connect(self._apply_theme)
        self._setup_ui()
        self._apply_theme(self._theme.colors)
        self._load_settings()

    def _setup_ui(self):
        # 窗口设置
        self.setWindowTitle("MD to Word")
        self.setMinimumSize(WINDOW_SIZES['min_width'], WINDOW_SIZES['min_height'])
        self.resize(WINDOW_SIZES['min_width'], WINDOW_SIZES['min_height'])

        # 无边框窗口 (可选, 暂时保留系统边框)
        # self.setWindowFlags(Qt.FramelessWindowHint)

        # 中央 Widget
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # 主布局: 水平 (Sidebar | 右侧区域)
        main_layout = QHBoxLayout(self.central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 侧边栏
        self.sidebar = Sidebar()
        self.sidebar.navigation_changed.connect(self._on_navigation_changed)
        main_layout.addWidget(self.sidebar)

        # 右侧区域 (垂直布局)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        # 顶部栏
        self.header_bar = HeaderBar()
        self.header_bar.theme_toggle_clicked.connect(self._toggle_theme)
        self.header_bar.preview_toggle_clicked.connect(self._toggle_preview)
        self.header_bar.minimize_clicked.connect(self.showMinimized)
        self.header_bar.maximize_clicked.connect(self._toggle_maximize)
        self.header_bar.close_clicked.connect(self.close)
        right_layout.addWidget(self.header_bar)

        # 内容区域 (堆叠)
        self.content_stack = QStackedWidget()
        self.content_stack.setObjectName('content_area')

        # 页面
        self.md_to_word_page = MdToWordPage()
        self.md_to_word_page.conversion_finished.connect(self._on_conversion_finished)
        self.content_stack.addWidget(self.md_to_word_page)

        self.word_to_md_page = WordToMdPage()
        self.word_to_md_page.conversion_finished.connect(self._on_conversion_finished)
        self.content_stack.addWidget(self.word_to_md_page)

        self.history_view = HistoryView()
        self.history_view.open_file_requested.connect(self._open_file)
        self.content_stack.addWidget(self.history_view)

        # 设置页面
        self.settings_page = SettingsPage()
        self.settings_page.theme_change_requested.connect(self._on_theme_change_requested)
        self.settings_page.history_cleared.connect(self._on_history_cleared)
        self.content_stack.addWidget(self.settings_page)

        right_layout.addWidget(self.content_stack, 1)

        # 底部状态栏
        self.status_bar = StatusBar(version=self.VERSION)
        right_layout.addWidget(self.status_bar)

        main_layout.addWidget(right_widget, 1)

        # 页面映射
        self._page_map = {
            'md_to_word': 0,
            'word_to_md': 1,
            'history': 2,
            'settings': 3,
        }

    def _toggle_theme(self):
        """切换主题"""
        self._theme.toggle()

    def _toggle_maximize(self):
        """切换最大化状态"""
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def _on_navigation_changed(self, key: str):
        """导航切换处理"""
        if key in self._page_map:
            self.content_stack.setCurrentIndex(self._page_map[key])

            # 更新标题
            titles = {
                'md_to_word': 'MD to Word',
                'word_to_md': 'Word to MD',
                'history': '历史记录',
                'settings': '设置',
            }
            self.header_bar.set_title(titles.get(key, 'MD to Word'))

    def _on_conversion_finished(self, success: bool, message: str):
        """转换完成处理"""
        if success:
            self.status_bar.show_success(message, 5000)
        else:
            self.status_bar.show_error(message, 5000)

    def _open_file(self, file_path: str):
        """打开文件"""
        if not os.path.exists(file_path):
            self.status_bar.show_error(f"文件不存在: {file_path}", 3000)
            return

        try:
            if sys.platform == 'win32':
                os.startfile(file_path)
            elif sys.platform == 'darwin':
                subprocess.run(['open', file_path])
            else:
                subprocess.run(['xdg-open', file_path])
        except Exception as e:
            self.status_bar.show_error(f"无法打开文件: {str(e)}", 3000)

    def _apply_theme(self, colors: dict):
        """应用主题样式"""
        is_dark = self._theme.is_dark

        # 全局样式表
        self.setStyleSheet(generate_stylesheet(colors))
        self.central_widget.setStyleSheet(
            f"background-color: {colors['bg_dark']};"
        )

        # 内容区域背景
        self.content_stack.setStyleSheet(f"""
            #content_area {{
                background-color: {colors['surface']};
            }}
        """)

        # 更新各组件主题
        self.sidebar.set_theme(is_dark, colors)
        self.header_bar.set_theme(is_dark, colors)
        self.status_bar.set_theme(is_dark, colors)

        # 更新页面主题
        self.md_to_word_page.set_theme(is_dark, colors)
        self.word_to_md_page.set_theme(is_dark, colors)
        self.history_view.set_theme(is_dark, colors)
        self.settings_page.set_theme(is_dark, colors)

    def show_status(self, message: str, status_type: str = 'normal', duration: int = 3000):
        """显示状态消息"""
        self.status_bar.set_status(message, status_type, duration)

    def _load_settings(self):
        """加载并应用设置"""
        # 加载预览面板可见性设置
        preview_visible = self._settings.preview_visible
        self._apply_preview_visibility(preview_visible)

    def _toggle_preview(self):
        """切换预览面板显示/隐藏"""
        new_state = self._settings.toggle_preview()
        self._apply_preview_visibility(new_state)

    def _apply_preview_visibility(self, visible: bool):
        """应用预览面板可见性"""
        self.md_to_word_page.set_preview_visible(visible)
        self.word_to_md_page.set_preview_visible(visible)
        self.header_bar.set_preview_visible(visible)

    def _on_theme_change_requested(self, is_dark: bool):
        """处理设置页面的主题切换请求"""
        if is_dark != self._theme.is_dark:
            self._theme.toggle()

    def _on_history_cleared(self):
        """处理历史记录清除"""
        self.history_view.refresh()
        self.status_bar.show_success("历史记录已清除", 3000)
