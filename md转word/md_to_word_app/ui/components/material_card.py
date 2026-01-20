# -*- coding: utf-8 -*-
"""
Material Design 卡片容器组件
带阴影和圆角的容器
"""

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QGraphicsDropShadowEffect
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter, QColor, QPainterPath

from ..styles.elevation import apply_elevation
from ..styles.scaling import scaled_size, scaled_spacing


class MaterialCard(QWidget):
    """
    Material Design 风格的卡片容器
    特点:
    - 圆角边框
    - 阴影效果
    - 支持主题切换
    """

    def __init__(self, elevation: int = 1, parent=None):
        super().__init__(parent)
        self._elevation = elevation
        self._is_dark = True
        self._bg_color = QColor('#2e2e2e')
        self._border_color = QColor('#383838')
        self._border_radius = scaled_size(12)

        # 内部布局
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(
            scaled_spacing(16), scaled_spacing(16),
            scaled_spacing(16), scaled_spacing(16)
        )
        self._layout.setSpacing(scaled_spacing(12))

        # 应用阴影
        self._apply_elevation()

    def _apply_elevation(self):
        """应用阴影效果"""
        apply_elevation(self, self._elevation, self._is_dark)

    def set_elevation(self, level: int):
        """设置阴影层级"""
        if self._elevation != level:
            self._elevation = level
            self._apply_elevation()

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark
        self._bg_color = QColor(colors.get('bg_card', '#2e2e2e'))
        self._border_color = QColor(colors.get('border_subtle', '#383838'))
        self._apply_elevation()
        self.update()

    def set_border_radius(self, radius: int):
        """设置圆角半径"""
        self._border_radius = radius
        self.update()

    def layout(self):
        """返回内部布局"""
        return self._layout

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制圆角背景
        path = QPainterPath()
        path.addRoundedRect(
            0, 0, self.width(), self.height(),
            self._border_radius, self._border_radius
        )

        painter.fillPath(path, self._bg_color)

        # 绘制边框
        painter.setPen(self._border_color)
        painter.drawPath(path)


class MaterialCardOutlined(MaterialCard):
    """
    带边框的 Material 卡片
    无阴影, 使用边框区分
    """

    def __init__(self, parent=None):
        super().__init__(elevation=0, parent=parent)
        self._border_width = 1

    def set_border_width(self, width: int):
        """设置边框宽度"""
        self._border_width = width
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制圆角背景
        path = QPainterPath()
        half_border = self._border_width / 2
        path.addRoundedRect(
            half_border, half_border,
            self.width() - self._border_width,
            self.height() - self._border_width,
            self._border_radius, self._border_radius
        )

        painter.fillPath(path, self._bg_color)

        # 绘制边框
        from PyQt5.QtGui import QPen
        pen = QPen(self._border_color)
        pen.setWidth(self._border_width)
        painter.setPen(pen)
        painter.drawPath(path)


class MaterialCardElevated(MaterialCard):
    """
    带悬停效果的卡片
    悬停时提升阴影层级
    """

    def __init__(self, elevation: int = 1, parent=None):
        super().__init__(elevation, parent)
        self._base_elevation = elevation
        self._hover_elevation = min(elevation + 2, 5)

    def enterEvent(self, event):
        self.set_elevation(self._hover_elevation)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.set_elevation(self._base_elevation)
        super().leaveEvent(event)


class ContentCard(MaterialCard):
    """
    内容区域卡片
    用于包裹主要内容区域
    """

    def __init__(self, title: str = None, parent=None):
        super().__init__(elevation=1, parent=parent)
        self._title = title

        if title:
            from PyQt5.QtWidgets import QLabel
            self._title_label = QLabel(title)
            self._title_label.setObjectName('card_title')
            self._layout.insertWidget(0, self._title_label)

    def set_title(self, title: str):
        """设置标题"""
        self._title = title
        if hasattr(self, '_title_label'):
            self._title_label.setText(title)

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        super().set_theme(is_dark, colors)

        if hasattr(self, '_title_label'):
            self._title_label.setStyleSheet(f"""
                font-size: 16px;
                font-weight: 600;
                color: {colors.get('text_primary', '#E3E3E3')};
                background: transparent;
                padding-bottom: 8px;
            """)
