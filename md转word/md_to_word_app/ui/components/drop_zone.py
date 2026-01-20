# -*- coding: utf-8 -*-
"""
拖拽区域组件 - 文件拖放功能
"""

import os
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QLabel, QFrame
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QFont
from .base import ThemedFrame
from ..styles import scaled_font, scaled_size, scaled_spacing


class FileDropZone(ThemedFrame):
    """
    通用文件拖拽区域组件

    支持文件拖放, 显示文件选中状态, 并发送信号通知文件变更
    """
    file_dropped = pyqtSignal(str)

    def __init__(self, icon_text="FILE", hint_text="Drag file here",
                 extensions=None, parent=None):
        super().__init__(parent)
        self.has_file = False
        self.icon_text = icon_text
        self.hint_text = hint_text
        self.extensions = extensions or []
        self.setAcceptDrops(True)
        self._setup_ui()
        self._apply_theme(self.colors)

    def _setup_ui(self):
        self.setMinimumHeight(scaled_size(100))
        self.setFrameStyle(QFrame.NoFrame)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(scaled_spacing(6))
        layout.setContentsMargins(scaled_spacing(12), scaled_spacing(12), scaled_spacing(12), scaled_spacing(12))

        # 图标容器
        icon_size = scaled_size(48)
        self.icon_container = QFrame()
        self.icon_container.setFixedSize(icon_size, icon_size)
        icon_layout = QVBoxLayout(self.icon_container)
        icon_layout.setAlignment(Qt.AlignCenter)
        icon_layout.setContentsMargins(0, 0, 0, 0)

        self.icon_label = QLabel(self.icon_text)
        self.icon_label.setFont(QFont("Segoe UI", scaled_font(10), QFont.Bold))
        self.icon_label.setStyleSheet("color: white; background: transparent;")
        self.icon_label.setAlignment(Qt.AlignCenter)
        icon_layout.addWidget(self.icon_label)

        icon_h_layout = QHBoxLayout()
        icon_h_layout.addStretch()
        icon_h_layout.addWidget(self.icon_container)
        icon_h_layout.addStretch()
        layout.addLayout(icon_h_layout)

        self.hint_label = QLabel(self.hint_text)
        self.hint_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.hint_label)

        self.file_label = QLabel("")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.hide()
        layout.addWidget(self.file_label)

    def _apply_theme(self, colors: dict):
        # 图标容器渐变
        icon_radius = scaled_size(14)
        self.icon_container.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 {colors['accent_primary']}, stop:1 {colors['accent_glow']});
                border-radius: {icon_radius}px;
            }}
        """)

        # 提示标签
        fs_hint = scaled_font(13)
        self.hint_label.setStyleSheet(f"""
            color: {colors['text_primary']};
            font-size: {fs_hint}px;
            font-weight: 500;
            background: transparent;
        """)

        # 文件名标签
        fs_file = scaled_font(11)
        self.file_label.setStyleSheet(f"""
            color: {colors['accent_secondary']};
            font-size: {fs_file}px;
            font-weight: 600;
            background: transparent;
        """)

        # 边框样式
        if self.has_file:
            self._apply_file_selected_style()
        else:
            self._apply_normal_style()

    def _apply_normal_style(self):
        colors = self.colors
        radius = scaled_size(10)
        self.setStyleSheet(f"""
            FileDropZone, MarkdownDropZone, WordDropZone {{
                background-color: {colors['bg_primary']};
                border: 2px dashed {colors['border_default']};
                border-radius: {radius}px;
            }}
            FileDropZone:hover, MarkdownDropZone:hover, WordDropZone:hover {{
                border-color: {colors['accent_primary']};
                background-color: {colors['bg_secondary']};
            }}
        """)

    def _apply_drag_style(self):
        colors = self.colors
        radius = scaled_size(10)
        self.setStyleSheet(f"""
            FileDropZone, MarkdownDropZone, WordDropZone {{
                background-color: {colors['drag_bg']};
                border: 2px dashed {colors['accent_primary']};
                border-radius: {radius}px;
            }}
        """)

    def _apply_file_selected_style(self):
        colors = self.colors
        radius = scaled_size(10)
        self.setStyleSheet(f"""
            FileDropZone, MarkdownDropZone, WordDropZone {{
                background-color: {colors['bg_primary']};
                border: 2px solid {colors['accent_primary']};
                border-radius: {radius}px;
            }}
        """)

    def _check_extension(self, file_path: str) -> bool:
        return file_path.lower().endswith(tuple(self.extensions))

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1:
                file_path = urls[0].toLocalFile()
                if self._check_extension(file_path):
                    event.acceptProposedAction()
                    self._apply_drag_style()
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        if self.has_file:
            self._apply_file_selected_style()
        else:
            self._apply_normal_style()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if len(urls) == 1:
            file_path = urls[0].toLocalFile()
            if self._check_extension(file_path):
                self.set_file(file_path)
                self.file_dropped.emit(file_path)
                event.acceptProposedAction()
                return
        event.ignore()
        self._apply_normal_style()

    def set_file(self, file_path: str):
        """设置或清除选中的文件"""
        if file_path:
            self.has_file = True
            filename = os.path.basename(file_path)
            self.file_label.setText(filename)
            self.file_label.show()
            self.hint_label.setText("已选择文件")
            self._apply_file_selected_style()
        else:
            self.has_file = False
            self.file_label.hide()
            self.hint_label.setText(self.hint_text)
            self._apply_normal_style()

    def clear(self):
        """清除选中的文件"""
        self.set_file(None)


class MarkdownDropZone(FileDropZone):
    """Markdown 文件拖拽区域"""

    def __init__(self, hint_text="拖拽 Markdown 文件到这里", parent=None):
        super().__init__(
            icon_text="MD",
            hint_text=hint_text,
            extensions=['.md', '.markdown'],
            parent=parent
        )


class WordDropZone(FileDropZone):
    """Word 文件拖拽区域"""

    def __init__(self, hint_text="拖拽 Word 文件到这里", parent=None):
        super().__init__(
            icon_text="DOCX",
            hint_text=hint_text,
            extensions=['.docx'],
            parent=parent
        )


# 向后兼容别名
DropZone = MarkdownDropZone
