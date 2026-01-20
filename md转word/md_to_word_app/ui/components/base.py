# -*- coding: utf-8 -*-
"""
基础组件 - 支持主题自动更新的基类
"""

from PyQt5.QtWidgets import QWidget, QFrame, QPushButton
from ..styles import get_theme_manager


class ThemedMixin:
    """
    主题混入类

    为组件提供自动主题订阅和更新能力
    子类应实现 _apply_theme(colors) 方法
    """

    def _init_theme(self):
        """初始化主题连接, 在子类 __init__ 末尾调用"""
        self._theme = get_theme_manager()
        self._theme.theme_changed.connect(self._on_theme_changed)

    def _on_theme_changed(self, colors: dict):
        """主题变更回调"""
        self._apply_theme(colors)

    def _apply_theme(self, colors: dict):
        """
        应用主题样式 (子类应重写此方法)

        Args:
            colors: 当前主题颜色配置
        """
        pass

    @property
    def colors(self) -> dict:
        """获取当前颜色配置"""
        return self._theme.colors


class ThemedWidget(QWidget, ThemedMixin):
    """支持主题自动更新的 QWidget 基类"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_theme()


class ThemedFrame(QFrame, ThemedMixin):
    """支持主题自动更新的 QFrame 基类"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_theme()


class ThemedButton(QPushButton, ThemedMixin):
    """支持主题自动更新的 QPushButton 基类"""

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self._init_theme()
