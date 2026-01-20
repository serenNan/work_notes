# -*- coding: utf-8 -*-
"""
Material Design 顶部栏组件
包含标题、主题切换、窗口控制按钮
"""

from PyQt5.QtWidgets import (
    QWidget, QHBoxLayout, QLabel, QPushButton, QSizePolicy
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize
from PyQt5.QtGui import QPainter, QColor, QIcon, QPixmap
from PyQt5.QtSvg import QSvgRenderer

from ..styles.colors import WINDOW_SIZES
from ..styles.scaling import scaled_size, scaled_spacing, scaled_font
from ..styles.icons import get_icon_svg


class HeaderButton(QPushButton):
    """
    顶部栏按钮
    带悬停效果
    """

    def __init__(self, icon_name: str, parent=None):
        super().__init__(parent)
        self._icon_name = icon_name
        self._icon_color = '#CAC4CF'
        self._hover_bg = '#383838'

        self.setFixedSize(scaled_size(40), scaled_size(40))
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName('header_button')

        self._update_icon()

    def _update_icon(self):
        """更新图标"""
        svg_data = get_icon_svg(self._icon_name, self._icon_color)

        renderer = QSvgRenderer()
        renderer.load(svg_data.encode('utf-8'))

        size = scaled_size(20)
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()

        self.setIcon(QIcon(pixmap))
        self.setIconSize(QSize(size, size))

    def set_icon_color(self, color: str):
        """设置图标颜色"""
        self._icon_color = color
        self._update_icon()

    def set_hover_bg(self, color: str):
        """设置悬停背景色"""
        self._hover_bg = color


class HeaderBar(QWidget):
    """
    顶部栏
    包含: 标题 | 弹性空间 | 主题切换 | 窗口控制按钮
    """
    theme_toggle_clicked = pyqtSignal()
    minimize_clicked = pyqtSignal()
    maximize_clicked = pyqtSignal()
    close_clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._is_dark = True
        self._bg_color = QColor('#1e1e1e')
        self._border_color = QColor('#2e2e2e')
        self._title = "MD to Word"

        self.setFixedHeight(WINDOW_SIZES['header_height'])
        self._setup_ui()

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(
            scaled_spacing(16), 0, scaled_spacing(8), 0
        )
        layout.setSpacing(scaled_spacing(8))

        # 标题
        self._title_label = QLabel(self._title)
        self._title_label.setObjectName('header_title')
        layout.addWidget(self._title_label)

        # 弹性空间
        layout.addStretch()

        # 主题切换按钮
        self._theme_btn = HeaderButton('moon')
        self._theme_btn.setToolTip('切换主题')
        self._theme_btn.clicked.connect(self.theme_toggle_clicked.emit)
        layout.addWidget(self._theme_btn)

        # 分隔
        layout.addSpacing(scaled_spacing(8))

        # 最小化按钮
        self._minimize_btn = HeaderButton('minimize')
        self._minimize_btn.setToolTip('最小化')
        self._minimize_btn.clicked.connect(self.minimize_clicked.emit)
        layout.addWidget(self._minimize_btn)

        # 最大化按钮
        self._maximize_btn = HeaderButton('maximize')
        self._maximize_btn.setToolTip('最大化')
        self._maximize_btn.clicked.connect(self.maximize_clicked.emit)
        layout.addWidget(self._maximize_btn)

        # 关闭按钮
        self._close_btn = HeaderButton('close')
        self._close_btn.setToolTip('关闭')
        self._close_btn.setObjectName('close_button')
        self._close_btn.clicked.connect(self.close_clicked.emit)
        layout.addWidget(self._close_btn)

    def set_title(self, title: str):
        """设置标题"""
        self._title = title
        self._title_label.setText(title)

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark

        if is_dark:
            self._bg_color = QColor(colors.get('surface', '#1e1e1e'))
            self._border_color = QColor(colors.get('outline_variant', '#49454F'))
            icon_color = colors.get('on_surface_variant', '#CAC4CF')
            hover_bg = colors.get('surface_container_high', '#2e2e2e')
            text_color = colors.get('on_surface', '#E3E3E3')
            theme_icon = 'sun'
        else:
            self._bg_color = QColor(colors.get('surface', '#FFFBFE'))
            self._border_color = QColor(colors.get('outline_variant', '#CAC4D0'))
            icon_color = colors.get('on_surface_variant', '#49454F')
            hover_bg = colors.get('surface_container_high', '#ECE6F0')
            text_color = colors.get('on_surface', '#1C1B1F')
            theme_icon = 'moon'

        # 更新样式
        self.setStyleSheet(f"""
            HeaderBar {{
                background-color: {self._bg_color.name()};
                border-bottom: 1px solid {self._border_color.name()};
            }}
            #header_title {{
                font-size: {scaled_font(18)}px;
                font-weight: 600;
                color: {text_color};
                background: transparent;
            }}
            #header_button {{
                background-color: transparent;
                border: none;
                border-radius: {scaled_size(8)}px;
            }}
            #header_button:hover {{
                background-color: {hover_bg};
            }}
            #close_button:hover {{
                background-color: #C42B1C;
            }}
        """)

        # 更新按钮图标
        self._theme_btn._icon_name = theme_icon
        for btn in [self._theme_btn, self._minimize_btn, self._maximize_btn, self._close_btn]:
            btn.set_icon_color(icon_color)
            btn.set_hover_bg(hover_bg)

        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制背景
        painter.fillRect(self.rect(), self._bg_color)

        # 绘制底边框
        painter.setPen(self._border_color)
        painter.drawLine(0, self.height() - 1, self.width(), self.height() - 1)
