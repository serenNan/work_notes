# -*- coding: utf-8 -*-
"""
Material Design 3 阴影层级工具
使用 QGraphicsDropShadowEffect 实现
"""

from PyQt5.QtWidgets import QWidget, QGraphicsDropShadowEffect
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt


# Material Design 阴影层级配置
# (blur_radius, offset_y, color_alpha)
ELEVATION_LEVELS = {
    0: (0, 0, 0),
    1: (4, 2, 30),
    2: (8, 4, 35),
    3: (12, 6, 40),
    4: (16, 8, 45),
    5: (24, 12, 50),
}


def apply_elevation(widget: QWidget, level: int, is_dark: bool = True) -> QGraphicsDropShadowEffect:
    """
    为 widget 应用 Material Design 阴影效果

    Args:
        widget: 目标控件
        level: 阴影层级 (0-5)
        is_dark: 是否深色主题

    Returns:
        QGraphicsDropShadowEffect 实例
    """
    level = max(0, min(5, level))
    blur, offset_y, alpha = ELEVATION_LEVELS[level]

    if level == 0:
        widget.setGraphicsEffect(None)
        return None

    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(blur)
    shadow.setOffset(0, offset_y)

    # 深色主题使用更深的阴影
    if is_dark:
        shadow.setColor(QColor(0, 0, 0, alpha + 40))
    else:
        shadow.setColor(QColor(0, 0, 0, alpha))

    widget.setGraphicsEffect(shadow)
    return shadow


def get_elevation_shadow_css(level: int, is_dark: bool = True) -> str:
    """
    获取阴影的伪 CSS 描述 (用于调试/文档)
    注意: QSS 不支持 box-shadow, 此函数仅用于参考
    """
    level = max(0, min(5, level))
    blur, offset_y, alpha = ELEVATION_LEVELS[level]

    if is_dark:
        alpha += 40

    return f"/* elevation-{level}: blur={blur}px, offset-y={offset_y}px, alpha={alpha} */"


class ElevatedWidget(QWidget):
    """
    带阴影效果的基类 Widget
    支持动态调整阴影层级
    """

    def __init__(self, elevation: int = 1, parent=None):
        super().__init__(parent)
        self._elevation = elevation
        self._is_dark = True
        self._shadow_effect = None
        self._apply_elevation()

    def _apply_elevation(self):
        """应用当前阴影层级"""
        self._shadow_effect = apply_elevation(self, self._elevation, self._is_dark)

    def set_elevation(self, level: int):
        """设置阴影层级"""
        if self._elevation != level:
            self._elevation = level
            self._apply_elevation()

    def get_elevation(self) -> int:
        """获取当前阴影层级"""
        return self._elevation

    def set_theme_dark(self, is_dark: bool):
        """设置主题模式"""
        if self._is_dark != is_dark:
            self._is_dark = is_dark
            self._apply_elevation()

    elevation = property(get_elevation, set_elevation)
