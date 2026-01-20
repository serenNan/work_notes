# -*- coding: utf-8 -*-
"""
设置管理器 - 使用 QSettings 持久化应用设置
"""

from PyQt5.QtCore import QSettings, QObject, pyqtSignal


class SettingsManager(QObject):
    """
    应用设置管理器
    使用 QSettings 持久化存储
    """
    # 信号
    preview_visible_changed = pyqtSignal(bool)

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
        self._settings = QSettings('MD2Word', 'MDtoWordApp')

    # ========== 预览面板设置 ==========

    @property
    def preview_visible(self) -> bool:
        """预览面板是否可见"""
        return self._settings.value('ui/preview_visible', True, type=bool)

    @preview_visible.setter
    def preview_visible(self, value: bool):
        """设置预览面板可见性"""
        if self.preview_visible != value:
            self._settings.setValue('ui/preview_visible', value)
            self.preview_visible_changed.emit(value)

    def toggle_preview(self) -> bool:
        """切换预览面板可见性, 返回新状态"""
        new_value = not self.preview_visible
        self.preview_visible = new_value
        return new_value

    # ========== 主题设置 ==========

    @property
    def is_dark_theme(self) -> bool:
        """是否深色主题"""
        return self._settings.value('ui/dark_theme', True, type=bool)

    @is_dark_theme.setter
    def is_dark_theme(self, value: bool):
        """设置主题"""
        self._settings.setValue('ui/dark_theme', value)

    # ========== 窗口设置 ==========

    def save_window_geometry(self, geometry):
        """保存窗口几何信息"""
        self._settings.setValue('window/geometry', geometry)

    def load_window_geometry(self):
        """加载窗口几何信息"""
        return self._settings.value('window/geometry')

    def save_window_state(self, state):
        """保存窗口状态"""
        self._settings.setValue('window/state', state)

    def load_window_state(self):
        """加载窗口状态"""
        return self._settings.value('window/state')


# 全局实例
_settings_manager = None


def get_settings_manager() -> SettingsManager:
    """获取设置管理器单例"""
    global _settings_manager
    if _settings_manager is None:
        _settings_manager = SettingsManager()
    return _settings_manager
