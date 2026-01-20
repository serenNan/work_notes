# -*- coding: utf-8 -*-
"""
预览面板组件
实时预览 Markdown/Word 文件内容
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QTextEdit, QScrollArea, QPushButton, QStackedWidget
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor

from ..styles.scaling import scaled_size, scaled_spacing, scaled_font
from .material_card import MaterialCard


class PreviewPanel(QWidget):
    """
    预览面板
    支持 Markdown 源码预览和渲染预览
    """
    refresh_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._is_dark = True
        self._current_file = None
        self._preview_mode = 'source'  # source, rendered

        self._bg_color = '#242424'
        self._text_color = '#E3E3E3'
        self._muted_color = '#948F99'
        self._border_color = '#383838'

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 顶部工具栏
        toolbar = QWidget()
        toolbar.setObjectName('preview_toolbar')
        toolbar.setFixedHeight(scaled_size(40))

        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(
            scaled_spacing(12), 0, scaled_spacing(12), 0
        )
        toolbar_layout.setSpacing(scaled_spacing(8))

        # 标题
        self._title_label = QLabel("预览")
        self._title_label.setObjectName('preview_title')
        toolbar_layout.addWidget(self._title_label)

        # 文件名
        self._filename_label = QLabel()
        self._filename_label.setObjectName('preview_filename')
        toolbar_layout.addWidget(self._filename_label)

        toolbar_layout.addStretch()

        # 刷新按钮
        self._refresh_btn = QPushButton("刷新")
        self._refresh_btn.setObjectName('preview_refresh_btn')
        self._refresh_btn.clicked.connect(self.refresh_requested.emit)
        toolbar_layout.addWidget(self._refresh_btn)

        layout.addWidget(toolbar)

        # 分割线
        separator = QWidget()
        separator.setFixedHeight(1)
        separator.setObjectName('preview_separator')
        layout.addWidget(separator)

        # 内容区域
        self._content_stack = QStackedWidget()

        # 空状态
        self._empty_widget = QWidget()
        empty_layout = QVBoxLayout(self._empty_widget)
        empty_layout.setAlignment(Qt.AlignCenter)

        empty_icon = QLabel("📄")
        empty_icon.setStyleSheet(f"font-size: {scaled_font(48)}px;")
        empty_icon.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_icon)

        empty_text = QLabel("选择文件后在此预览")
        empty_text.setObjectName('empty_text')
        empty_text.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_text)

        self._content_stack.addWidget(self._empty_widget)

        # 源码预览
        self._source_view = QTextEdit()
        self._source_view.setReadOnly(True)
        self._source_view.setObjectName('source_view')
        self._source_view.setFont(QFont('Consolas', scaled_font(12)))
        self._source_view.setLineWrapMode(QTextEdit.NoWrap)
        self._content_stack.addWidget(self._source_view)

        layout.addWidget(self._content_stack, 1)

    def set_file(self, file_path: str):
        """设置要预览的文件"""
        self._current_file = file_path

        if not file_path or not os.path.exists(file_path):
            self._content_stack.setCurrentWidget(self._empty_widget)
            self._filename_label.setText("")
            return

        # 显示文件名
        filename = os.path.basename(file_path)
        self._filename_label.setText(filename)

        # 读取文件内容
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            self._source_view.setPlainText(content)
            self._content_stack.setCurrentWidget(self._source_view)
        except Exception as e:
            self._source_view.setPlainText(f"无法读取文件: {str(e)}")
            self._content_stack.setCurrentWidget(self._source_view)

    def clear(self):
        """清除预览"""
        self._current_file = None
        self._filename_label.setText("")
        self._source_view.clear()
        self._content_stack.setCurrentWidget(self._empty_widget)

    def refresh(self):
        """刷新预览"""
        if self._current_file:
            self.set_file(self._current_file)

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark

        if is_dark:
            self._bg_color = colors.get('surface_container', '#242424')
            self._text_color = colors.get('on_surface', '#E3E3E3')
            self._muted_color = colors.get('outline', '#948F99')
            self._border_color = colors.get('outline_variant', '#49454F')
            self._surface_color = colors.get('surface_container_low', '#1a1a1a')
        else:
            self._bg_color = colors.get('surface_container', '#F3EDF7')
            self._text_color = colors.get('on_surface', '#1C1B1F')
            self._muted_color = colors.get('outline', '#79747E')
            self._border_color = colors.get('outline_variant', '#CAC4D0')
            self._surface_color = colors.get('surface_container_low', '#F7F2FA')

        self.setStyleSheet(f"""
            PreviewPanel {{
                background-color: {self._bg_color};
                border-radius: {scaled_size(12)}px;
            }}
            #preview_toolbar {{
                background-color: transparent;
            }}
            #preview_title {{
                font-size: {scaled_font(14)}px;
                font-weight: 600;
                color: {self._text_color};
                background: transparent;
            }}
            #preview_filename {{
                font-size: {scaled_font(12)}px;
                color: {self._muted_color};
                background: transparent;
                padding-left: {scaled_spacing(8)}px;
            }}
            #preview_separator {{
                background-color: {self._border_color};
            }}
            #preview_refresh_btn {{
                background-color: transparent;
                border: 1px solid {self._border_color};
                border-radius: {scaled_size(6)}px;
                padding: {scaled_spacing(4)}px {scaled_spacing(12)}px;
                color: {self._text_color};
                font-size: {scaled_font(12)}px;
            }}
            #preview_refresh_btn:hover {{
                background-color: {self._surface_color};
            }}
            #empty_text {{
                font-size: {scaled_font(14)}px;
                color: {self._muted_color};
                background: transparent;
            }}
            #source_view {{
                background-color: {self._surface_color};
                color: {self._text_color};
                border: none;
                padding: {scaled_spacing(12)}px;
                selection-background-color: {colors.get('primary_container', '#004A77')};
            }}
        """)

        self.update()


class PreviewCard(MaterialCard):
    """
    预览卡片
    带标题的预览面板容器
    """

    def __init__(self, title: str = "预览", parent=None):
        super().__init__(elevation=1, parent=parent)
        self._title = title

        # 标题
        self._title_label = QLabel(title)
        self._title_label.setObjectName('card_title')
        self._layout.insertWidget(0, self._title_label)

        # 预览面板
        self._preview = PreviewPanel()
        self._layout.addWidget(self._preview, 1)

    def preview(self) -> PreviewPanel:
        """返回预览面板"""
        return self._preview

    def set_file(self, file_path: str):
        """设置预览文件"""
        self._preview.set_file(file_path)

    def clear(self):
        """清除预览"""
        self._preview.clear()

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        super().set_theme(is_dark, colors)

        self._title_label.setStyleSheet(f"""
            font-size: {scaled_font(16)}px;
            font-weight: 600;
            color: {colors.get('on_surface', '#E3E3E3')};
            background: transparent;
            padding-bottom: {scaled_spacing(8)}px;
        """)

        self._preview.set_theme(is_dark, colors)
