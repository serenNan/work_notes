# -*- coding: utf-8 -*-
"""
Material Design 侧边栏导航组件
支持折叠/展开动画
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QSpacerItem, QSizePolicy
)
from PyQt5.QtCore import Qt, pyqtSignal, QPropertyAnimation, QEasingCurve, QSize
from PyQt5.QtGui import QPainter, QColor, QPainterPath, QIcon

from ..styles.colors import WINDOW_SIZES
from ..styles.scaling import scaled_size, scaled_spacing
from ..styles.icons import render_icon_to_pixmap
from .ripple import RippleWidget


class NavItem(RippleWidget):
    """
    侧边栏导航项
    点击时显示涟漪效果
    """
    clicked = pyqtSignal(str)

    def __init__(self, key: str, icon_name: str, text: str, parent=None):
        super().__init__(parent)
        self._key = key
        self._icon_name = icon_name
        self._text = text
        self._selected = False
        self._expanded = True
        self._is_dark = True

        # 颜色
        self._bg_color = QColor('transparent')
        self._selected_bg = QColor('#004A77')
        self._text_color = QColor('#CAC4CF')
        self._selected_text_color = QColor('#90CAF9')
        self._icon_color = '#CAC4CF'
        self._selected_icon_color = '#90CAF9'

        self.setFixedHeight(scaled_size(48))
        self.setCursor(Qt.PointingHandCursor)

        # 布局
        self._layout = QHBoxLayout(self)
        self._layout.setContentsMargins(
            scaled_spacing(12), 0, scaled_spacing(12), 0
        )
        self._layout.setSpacing(scaled_spacing(12))

        # 图标容器
        self._icon_label = QLabel()
        self._icon_label.setFixedSize(scaled_size(24), scaled_size(24))
        self._icon_label.setAlignment(Qt.AlignCenter)
        self._layout.addWidget(self._icon_label)

        # 文字标签
        self._text_label = QLabel(text)
        self._text_label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self._layout.addWidget(self._text_label, 1)

        self._update_icon()

    def _update_icon(self):
        """更新图标"""
        color = self._selected_icon_color if self._selected else self._icon_color
        size = scaled_size(24)
        pixmap = render_icon_to_pixmap(self._icon_name, size, color)
        self._icon_label.setPixmap(pixmap)

    def set_selected(self, selected: bool):
        """设置选中状态"""
        if self._selected != selected:
            self._selected = selected
            self._update_icon()
            self.update()

    def is_selected(self) -> bool:
        return self._selected

    def set_expanded(self, expanded: bool):
        """设置展开状态"""
        self._expanded = expanded
        self._text_label.setVisible(expanded)

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark

        if is_dark:
            self._bg_color = QColor('transparent')
            self._selected_bg = QColor(colors.get('primary_container', '#004A77'))
            self._text_color = QColor(colors.get('on_surface_variant', '#CAC4CF'))
            self._selected_text_color = QColor(colors.get('primary', '#90CAF9'))
            self._icon_color = colors.get('on_surface_variant', '#CAC4CF')
            self._selected_icon_color = colors.get('primary', '#90CAF9')
            self.set_ripple_color(QColor(colors.get('primary', '#90CAF9')))
        else:
            self._bg_color = QColor('transparent')
            self._selected_bg = QColor(colors.get('primary_container', '#EADDFF'))
            self._text_color = QColor(colors.get('on_surface_variant', '#49454F'))
            self._selected_text_color = QColor(colors.get('primary', '#6750A4'))
            self._icon_color = colors.get('on_surface_variant', '#49454F')
            self._selected_icon_color = colors.get('primary', '#6750A4')
            self.set_ripple_color(QColor(colors.get('primary', '#6750A4')))

        # 更新文字颜色
        text_color = self._selected_text_color if self._selected else self._text_color
        self._text_label.setStyleSheet(f"""
            color: {text_color.name()};
            font-size: {scaled_size(14)}px;
            font-weight: {'600' if self._selected else '400'};
            background: transparent;
        """)

        self._update_icon()
        self.update()

    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)
        if event.button() == Qt.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit(self._key)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制背景
        radius = scaled_size(24)
        path = QPainterPath()
        path.addRoundedRect(
            scaled_spacing(8), 0,
            self.width() - scaled_spacing(16), self.height(),
            radius, radius
        )

        if self._selected:
            painter.fillPath(path, self._selected_bg)

        # 绘制涟漪
        super().paintEvent(event)


class Sidebar(QWidget):
    """
    侧边栏导航
    支持折叠/展开动画
    """
    navigation_changed = pyqtSignal(str)
    collapse_toggled = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._expanded = True
        self._is_dark = True
        self._current_key = 'md_to_word'
        self._nav_items = {}

        self._expanded_width = WINDOW_SIZES['sidebar_expanded']
        self._collapsed_width = WINDOW_SIZES['sidebar_collapsed']

        self.setFixedWidth(self._expanded_width)
        self.setMinimumHeight(WINDOW_SIZES['min_height'])

        self._bg_color = QColor('#1a1a1a')
        self._border_color = QColor('#2e2e2e')

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, scaled_spacing(8), 0, scaled_spacing(8))
        layout.setSpacing(scaled_spacing(4))

        # 折叠按钮
        self._toggle_btn = QPushButton()
        self._toggle_btn.setFixedSize(scaled_size(40), scaled_size(40))
        self._toggle_btn.setCursor(Qt.PointingHandCursor)
        self._toggle_btn.clicked.connect(self._toggle_expand)
        self._toggle_btn.setObjectName('sidebar_toggle')

        toggle_container = QHBoxLayout()
        toggle_container.setContentsMargins(scaled_spacing(12), 0, scaled_spacing(12), 0)
        toggle_container.addWidget(self._toggle_btn)
        toggle_container.addStretch()
        layout.addLayout(toggle_container)

        layout.addSpacing(scaled_spacing(12))

        # 主导航项
        main_nav_items = [
            ('md_to_word', 'md_to_word', 'MD to Word'),
            ('word_to_md', 'word_to_md', 'Word to MD'),
            ('history', 'history', '历史记录'),
        ]

        for key, icon, text in main_nav_items:
            item = NavItem(key, icon, text)
            item.clicked.connect(self._on_nav_clicked)
            self._nav_items[key] = item
            layout.addWidget(item)

        # 弹性空间
        layout.addStretch()

        # 分割线
        separator = QWidget()
        separator.setFixedHeight(1)
        separator.setObjectName('sidebar_separator')
        layout.addWidget(separator)

        # 底部导航项
        bottom_nav_items = [
            ('settings', 'settings', '设置'),
        ]

        for key, icon, text in bottom_nav_items:
            item = NavItem(key, icon, text)
            item.clicked.connect(self._on_nav_clicked)
            self._nav_items[key] = item
            layout.addWidget(item)

        # 默认选中
        self._nav_items['md_to_word'].set_selected(True)

    def _toggle_expand(self):
        """切换展开/折叠状态"""
        self._expanded = not self._expanded

        # 动画
        self._animation = QPropertyAnimation(self, b"minimumWidth")
        self._animation.setDuration(200)
        self._animation.setEasingCurve(QEasingCurve.OutCubic)

        if self._expanded:
            self._animation.setStartValue(self._collapsed_width)
            self._animation.setEndValue(self._expanded_width)
        else:
            self._animation.setStartValue(self._expanded_width)
            self._animation.setEndValue(self._collapsed_width)

        self._animation.start()

        # 同步 fixedWidth
        self._animation.finished.connect(
            lambda: self.setFixedWidth(
                self._expanded_width if self._expanded else self._collapsed_width
            )
        )

        # 更新导航项
        for item in self._nav_items.values():
            item.set_expanded(self._expanded)

        self.collapse_toggled.emit(self._expanded)

    def _on_nav_clicked(self, key: str):
        """导航项点击处理"""
        if key == self._current_key:
            return

        # 更新选中状态
        if self._current_key in self._nav_items:
            self._nav_items[self._current_key].set_selected(False)

        self._current_key = key
        self._nav_items[key].set_selected(True)

        self.navigation_changed.emit(key)

    def set_current(self, key: str):
        """设置当前选中项"""
        if key in self._nav_items:
            self._on_nav_clicked(key)

    def get_current(self) -> str:
        """获取当前选中项"""
        return self._current_key

    def is_expanded(self) -> bool:
        """是否展开"""
        return self._expanded

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark

        if is_dark:
            self._bg_color = QColor(colors.get('surface_container_low', '#1a1a1a'))
            self._border_color = QColor(colors.get('outline_variant', '#49454F'))
        else:
            self._bg_color = QColor(colors.get('surface_container_low', '#F7F2FA'))
            self._border_color = QColor(colors.get('outline_variant', '#CAC4D0'))

        # 更新样式
        self.setStyleSheet(f"""
            Sidebar {{
                background-color: {self._bg_color.name()};
                border-right: 1px solid {self._border_color.name()};
            }}
            #sidebar_toggle {{
                background-color: transparent;
                border: none;
                border-radius: {scaled_size(8)}px;
            }}
            #sidebar_toggle:hover {{
                background-color: {colors.get('surface_container', '#242424')};
            }}
            #sidebar_separator {{
                background-color: {self._border_color.name()};
                margin: {scaled_spacing(8)}px {scaled_spacing(16)}px;
            }}
        """)

        # 更新折叠按钮图标
        icon_color = colors.get('on_surface_variant', '#CAC4CF')
        icon_name = 'chevron_left' if self._expanded else 'chevron_right'
        size = scaled_size(24)
        pixmap = render_icon_to_pixmap(icon_name, size, icon_color)

        self._toggle_btn.setIcon(QIcon(pixmap))
        self._toggle_btn.setIconSize(QSize(size, size))

        # 更新导航项
        for item in self._nav_items.values():
            item.set_theme(is_dark, colors)

        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制背景
        painter.fillRect(self.rect(), self._bg_color)

        # 绘制右边框
        painter.setPen(self._border_color)
        painter.drawLine(self.width() - 1, 0, self.width() - 1, self.height())
