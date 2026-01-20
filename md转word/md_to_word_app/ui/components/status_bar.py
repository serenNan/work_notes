# -*- coding: utf-8 -*-
"""
Material Design 状态栏组件
显示转换状态和版本信息
"""

from PyQt5.QtWidgets import QWidget, QHBoxLayout, QLabel
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QPainter, QColor

from ..styles.colors import WINDOW_SIZES
from ..styles.scaling import scaled_size, scaled_spacing, scaled_font


class StatusBar(QWidget):
    """
    底部状态栏
    左侧: 状态信息
    右侧: 版本号
    """

    def __init__(self, version: str = "v1.0.0", parent=None):
        super().__init__(parent)
        self._is_dark = True
        self._bg_color = QColor('#1a1a1a')
        self._border_color = QColor('#2e2e2e')
        self._version = version
        self._status_type = 'normal'  # normal, success, error, warning

        self.setFixedHeight(WINDOW_SIZES['status_bar_height'])
        self._setup_ui()

        # 状态消息自动清除定时器
        self._clear_timer = QTimer(self)
        self._clear_timer.setSingleShot(True)
        self._clear_timer.timeout.connect(self._clear_status)

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(
            scaled_spacing(16), 0, scaled_spacing(16), 0
        )
        layout.setSpacing(scaled_spacing(12))

        # 状态图标 (可选)
        self._status_icon = QLabel()
        self._status_icon.setFixedSize(scaled_size(16), scaled_size(16))
        self._status_icon.setVisible(False)
        layout.addWidget(self._status_icon)

        # 状态文本
        self._status_label = QLabel("就绪")
        self._status_label.setObjectName('status_text')
        layout.addWidget(self._status_label)

        # 弹性空间
        layout.addStretch()

        # 版本号
        self._version_label = QLabel(self._version)
        self._version_label.setObjectName('version_text')
        layout.addWidget(self._version_label)

    def set_status(self, message: str, status_type: str = 'normal', duration: int = 0):
        """
        设置状态信息

        Args:
            message: 状态消息
            status_type: 状态类型 (normal, success, error, warning)
            duration: 自动清除时间 (毫秒), 0 表示不自动清除
        """
        self._status_label.setText(message)
        self._status_type = status_type
        self._update_status_style()

        if duration > 0:
            self._clear_timer.start(duration)

    def _clear_status(self):
        """清除状态, 恢复默认"""
        self._status_label.setText("就绪")
        self._status_type = 'normal'
        self._update_status_style()

    def _update_status_style(self):
        """更新状态样式"""
        colors = {
            'normal': self._text_color,
            'success': '#A5D6A7' if self._is_dark else '#2E7D32',
            'error': '#F2B8B5' if self._is_dark else '#B3261E',
            'warning': '#FFE082' if self._is_dark else '#F9A825',
        }

        color = colors.get(self._status_type, self._text_color)
        self._status_label.setStyleSheet(f"""
            color: {color};
            font-size: {scaled_font(12)}px;
            background: transparent;
        """)

    def show_success(self, message: str, duration: int = 5000):
        """显示成功消息"""
        self.set_status(message, 'success', duration)

    def show_error(self, message: str, duration: int = 5000):
        """显示错误消息"""
        self.set_status(message, 'error', duration)

    def show_warning(self, message: str, duration: int = 5000):
        """显示警告消息"""
        self.set_status(message, 'warning', duration)

    def set_version(self, version: str):
        """设置版本号"""
        self._version = version
        self._version_label.setText(version)

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark

        if is_dark:
            self._bg_color = QColor(colors.get('surface_container_low', '#1a1a1a'))
            self._border_color = QColor(colors.get('outline_variant', '#49454F'))
            self._text_color = colors.get('on_surface_variant', '#CAC4CF')
            self._muted_color = colors.get('outline', '#948F99')
        else:
            self._bg_color = QColor(colors.get('surface_container_low', '#F7F2FA'))
            self._border_color = QColor(colors.get('outline_variant', '#CAC4D0'))
            self._text_color = colors.get('on_surface_variant', '#49454F')
            self._muted_color = colors.get('outline', '#79747E')

        # 更新样式
        self.setStyleSheet(f"""
            StatusBar {{
                background-color: {self._bg_color.name()};
                border-top: 1px solid {self._border_color.name()};
            }}
            #status_text {{
                color: {self._text_color};
                font-size: {scaled_font(12)}px;
                background: transparent;
            }}
            #version_text {{
                color: {self._muted_color};
                font-size: {scaled_font(11)}px;
                background: transparent;
            }}
        """)

        self._update_status_style()
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制背景
        painter.fillRect(self.rect(), self._bg_color)

        # 绘制顶边框
        painter.setPen(self._border_color)
        painter.drawLine(0, 0, self.width(), 0)
