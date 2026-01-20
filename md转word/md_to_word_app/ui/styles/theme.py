# -*- coding: utf-8 -*-
"""
主题管理器 - 单例模式, 支持信号驱动的主题切换
"""

from PyQt5.QtCore import QObject, pyqtSignal, QSettings
from .colors import DARK_COLORS, LIGHT_COLORS


class ThemeManager(QObject):
    """
    主题管理器单例

    使用信号广播主题变更, 所有订阅的组件自动更新
    """
    theme_changed = pyqtSignal(dict)  # 广播当前颜色配置

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        super().__init__()
        self._initialized = True

        # 加载用户主题偏好
        self._settings = QSettings("MD2Word", "MarkdownToWord")
        self._is_dark = self._settings.value("dark_theme", True, type=bool)

    @property
    def is_dark(self) -> bool:
        """当前是否为深色主题"""
        return self._is_dark

    @property
    def colors(self) -> dict:
        """获取当前主题颜色配置"""
        return DARK_COLORS if self._is_dark else LIGHT_COLORS

    def toggle(self):
        """切换主题"""
        self._is_dark = not self._is_dark
        self._settings.setValue("dark_theme", self._is_dark)
        self.theme_changed.emit(self.colors)

    def set_dark(self, is_dark: bool):
        """设置主题"""
        if self._is_dark != is_dark:
            self._is_dark = is_dark
            self._settings.setValue("dark_theme", is_dark)
            self.theme_changed.emit(self.colors)


# 全局实例快捷访问
_theme_manager = None

def get_theme_manager() -> ThemeManager:
    """获取主题管理器实例"""
    global _theme_manager
    if _theme_manager is None:
        _theme_manager = ThemeManager()
    return _theme_manager
