# -*- coding: utf-8 -*-
"""
屏幕缩放系统 - 根据屏幕分辨率自动调整 UI 尺寸
"""

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QObject, pyqtSignal


class ScaleManager(QObject):
    """
    缩放管理器单例

    根据屏幕高度自动计算缩放比例:
    - 小屏幕 (<=900px): 缩放 0.85
    - 中等屏幕 (<=1080px): 缩放 1.0
    - 大屏幕 (>1080px): 缩放 1.15
    """
    _instance = None
    scale_changed = pyqtSignal(float)

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
        self._scale = 1.0
        self._detect_scale()

    def _detect_scale(self):
        """检测屏幕大小并设置缩放比例"""
        app = QApplication.instance()
        if app:
            screen = app.primaryScreen()
            if screen:
                height = screen.availableGeometry().height()
                if height <= 900:
                    self._scale = 0.8
                elif height <= 1080:
                    self._scale = 0.9
                else:
                    self._scale = 1.0

    @property
    def scale(self) -> float:
        return self._scale

    def font_size(self, base_size: int) -> int:
        """计算缩放后的字体大小"""
        return max(8, int(base_size * self._scale))

    def size(self, base_size: int) -> int:
        """计算缩放后的尺寸"""
        return max(4, int(base_size * self._scale))

    def spacing(self, base_spacing: int) -> int:
        """计算缩放后的间距"""
        return max(2, int(base_spacing * self._scale))


# 全局缩放管理器实例
_scale_manager = None


def get_scale_manager() -> ScaleManager:
    """获取缩放管理器单例"""
    global _scale_manager
    if _scale_manager is None:
        _scale_manager = ScaleManager()
    return _scale_manager


# 便捷函数
def scaled_font(base_size: int) -> int:
    """获取缩放后的字体大小"""
    return get_scale_manager().font_size(base_size)


def scaled_size(base_size: int) -> int:
    """获取缩放后的尺寸"""
    return get_scale_manager().size(base_size)


def scaled_spacing(base_spacing: int) -> int:
    """获取缩放后的间距"""
    return get_scale_manager().spacing(base_spacing)
