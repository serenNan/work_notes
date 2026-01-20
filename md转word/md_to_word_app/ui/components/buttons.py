# -*- coding: utf-8 -*-
"""
按钮组件 - 渐变主按钮, 次要按钮, 主题切换按钮
"""

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPainter, QColor
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtCore import QByteArray
from PyQt5.QtGui import QPixmap
from .base import ThemedButton
from ..styles import scaled_font, scaled_size, scaled_spacing


# SVG 图标定义
SUN_ICON_SVG = """<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="5"/>
  <line x1="12" y1="1" x2="12" y2="3"/>
  <line x1="12" y1="21" x2="12" y2="23"/>
  <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/>
  <line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/>
  <line x1="1" y1="12" x2="3" y2="12"/>
  <line x1="21" y1="12" x2="23" y2="12"/>
  <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/>
  <line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>
</svg>"""

MOON_ICON_SVG = """<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>
</svg>"""


def create_svg_icon(svg_template: str, color: str, size: int = 20) -> QIcon:
    """从 SVG 模板创建图标"""
    svg_data = svg_template.format(color=color)
    renderer = QSvgRenderer(QByteArray(svg_data.encode()))
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)


class GradientButton(ThemedButton):
    """渐变主按钮"""

    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFixedHeight(scaled_size(38))
        self.setCursor(Qt.PointingHandCursor)
        self._apply_theme(self.colors)

    def _apply_theme(self, colors: dict):
        fs = scaled_font(14)
        radius = scaled_size(7)
        pad_v = scaled_spacing(8)
        pad_h = scaled_spacing(16)
        self.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {colors['accent_primary']}, stop:1 {colors['accent_glow']});
                color: white;
                font-size: {fs}px;
                font-weight: 600;
                border: none;
                border-radius: {radius}px;
                padding: {pad_v}px {pad_h}px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {colors['accent_secondary']}, stop:1 {colors['accent_primary']});
            }}
            QPushButton:disabled {{
                background: {colors['bg_card']};
                color: {colors['text_muted']};
            }}
        """)


class SecondaryButton(ThemedButton):
    """次要按钮"""

    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFixedHeight(scaled_size(32))
        self.setCursor(Qt.PointingHandCursor)
        self._apply_theme(self.colors)

    def _apply_theme(self, colors: dict):
        fs = scaled_font(12)
        radius = scaled_size(5)
        pad_v = scaled_spacing(5)
        pad_h = scaled_spacing(10)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {colors['bg_card']};
                color: {colors['text_primary']};
                font-size: {fs}px;
                font-weight: 500;
                border: 1px solid {colors['border_default']};
                border-radius: {radius}px;
                padding: {pad_v}px {pad_h}px;
            }}
            QPushButton:hover {{
                background-color: {colors['bg_elevated']};
                border-color: {colors['accent_primary']};
            }}
        """)


class ThemeToggleButton(ThemedButton):
    """
    主题切换按钮

    深色模式显示月亮图标, 浅色模式显示太阳图标
    """

    def __init__(self, parent=None):
        super().__init__("", parent)
        btn_size = scaled_size(32)
        self.setFixedSize(btn_size, btn_size)
        self.setCursor(Qt.PointingHandCursor)
        self._apply_theme(self.colors)

    def _apply_theme(self, colors: dict):
        is_dark = self._theme.is_dark
        tooltip = "Switch to light theme" if is_dark else "Switch to dark theme"
        self.setToolTip(tooltip)

        # 根据当前主题选择图标
        icon_color = colors['text_primary']
        icon_size = scaled_size(18)
        if is_dark:
            icon = create_svg_icon(MOON_ICON_SVG, icon_color, icon_size)
        else:
            icon = create_svg_icon(SUN_ICON_SVG, icon_color, icon_size)

        self.setIcon(icon)

        btn_size = scaled_size(32)
        radius = btn_size // 2
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {colors['bg_card']};
                border: 1px solid {colors['border_default']};
                border-radius: {radius}px;
            }}
            QPushButton:hover {{
                background-color: {colors['bg_elevated']};
                border-color: {colors['accent_primary']};
            }}
        """)
