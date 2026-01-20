# -*- coding: utf-8 -*-
"""
进度条组件 - 带动画效果的进度指示器
"""

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from .base import ThemedWidget
from ..styles import scaled_font, scaled_size, scaled_spacing


class AnimatedProgressBar(ThemedWidget):
    """
    带动画效果的进度条

    支持不确定进度 (脉冲动画) 和确定进度两种模式
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self._apply_theme(self.colors)
        self.hide()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, scaled_spacing(6), 0, scaled_spacing(6))
        layout.setSpacing(scaled_spacing(3))

        # 状态标签
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(scaled_size(3))
        self.progress_bar.setRange(0, 0)  # 不确定模式
        layout.addWidget(self.progress_bar)

    def _apply_theme(self, colors: dict):
        fs = scaled_font(11)
        self.status_label.setStyleSheet(f"""
            color: {colors['text_muted']};
            font-size: {fs}px;
            background: transparent;
        """)

        radius = scaled_size(2)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                background-color: {colors['bg_secondary']};
                border: none;
                border-radius: {radius}px;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {colors['accent_primary']}, stop:1 {colors['accent_glow']});
                border-radius: {radius}px;
            }}
        """)

    def start(self, text: str = "Processing..."):
        """
        开始显示进度条 (不确定模式)

        Args:
            text: 状态文本
        """
        self.status_label.setText(text)
        self.progress_bar.setRange(0, 0)  # 不确定模式
        self.show()

    def set_progress(self, value: int, maximum: int = 100, text: str = None):
        """
        设置确定进度

        Args:
            value: 当前进度值
            maximum: 最大值
            text: 状态文本 (可选)
        """
        self.progress_bar.setRange(0, maximum)
        self.progress_bar.setValue(value)
        if text:
            self.status_label.setText(text)
        self.show()

    def stop(self):
        """停止并隐藏进度条"""
        self.hide()
        self.status_label.setText("")
        self.progress_bar.setRange(0, 0)


class StatusLabel(ThemedWidget):
    """
    状态标签组件

    简单的状态文本显示, 支持不同状态颜色
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self._apply_theme(self.colors)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.label = QLabel("")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

    def _apply_theme(self, colors: dict):
        self._colors = colors
        self.set_text(self.label.text())

    def set_text(self, text: str, status: str = "normal"):
        """
        设置状态文本

        Args:
            text: 文本内容
            status: 状态类型 ("normal", "success", "error", "accent")
        """
        self.label.setText(text)

        colors = self.colors
        color_map = {
            "normal": colors['text_muted'],
            "success": colors['success'],
            "error": colors['error'],
            "accent": colors['accent_secondary'],
        }
        color = color_map.get(status, colors['text_muted'])
        fs = scaled_font(11)

        self.label.setStyleSheet(f"""
            color: {color};
            font-size: {fs}px;
            background: transparent;
        """)

    def clear(self):
        """清除状态文本"""
        self.label.setText("")
