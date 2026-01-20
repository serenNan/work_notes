# -*- coding: utf-8 -*-
"""
主窗口 UI - 支持浅色/深色主题切换和双标签页
"""

from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTabWidget

from .styles import get_theme_manager, generate_stylesheet, generate_tab_widget_style, scaled_font, scaled_size, scaled_spacing
from .components import ThemeToggleButton
from .pages import MdToWordPage, WordToMdPage


class MainWindow(QMainWindow):
    """主窗口 - 支持主题切换和双标签页"""

    def __init__(self):
        super().__init__()
        self._theme = get_theme_manager()
        self._theme.theme_changed.connect(self._apply_theme)
        self._setup_ui()
        self._apply_theme(self._theme.colors)

    def _setup_ui(self):
        self.setWindowTitle("Markdown to Word")
        self.setMinimumSize(scaled_size(380), scaled_size(520))
        self.resize(scaled_size(420), scaled_size(600))

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setContentsMargins(scaled_spacing(20), scaled_spacing(16), scaled_spacing(20), scaled_spacing(16))
        main_layout.setSpacing(scaled_spacing(12))

        # 顶部栏: 标题 + 主题切换按钮
        top_bar = QHBoxLayout()
        self.title_label = QLabel("MD to Word")
        top_bar.addWidget(self.title_label)
        top_bar.addStretch()

        self.theme_btn = ThemeToggleButton()
        self.theme_btn.clicked.connect(self._toggle_theme)
        top_bar.addWidget(self.theme_btn)
        main_layout.addLayout(top_bar)

        # 标签页
        self.tab_widget = QTabWidget()

        # MD 转 Word 标签页
        self.md_to_word_page = MdToWordPage()
        self.tab_widget.addTab(self.md_to_word_page, "MD to Word")

        # Word 转 MD 标签页
        self.word_to_md_page = WordToMdPage()
        self.tab_widget.addTab(self.word_to_md_page, "Word to MD")

        main_layout.addWidget(self.tab_widget)

    def _toggle_theme(self):
        """切换主题"""
        self._theme.toggle()

    def _apply_theme(self, colors: dict):
        """应用主题样式"""
        # 全局样式表
        self.setStyleSheet(generate_stylesheet(colors))
        self.central_widget.setStyleSheet(f"background-color: {colors['bg_dark']};")

        # 标题样式
        self.title_label.setStyleSheet(f"""
            font-size: {scaled_font(20)}px;
            font-weight: 700;
            color: {colors['text_primary']};
            background: transparent;
        """)

        # 标签页样式
        self.tab_widget.setStyleSheet(generate_tab_widget_style(colors))
