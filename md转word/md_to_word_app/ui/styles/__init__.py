# -*- coding: utf-8 -*-
"""
样式模块 - 提供主题管理, 颜色配置和样式表生成
"""

from .colors import DARK_COLORS, LIGHT_COLORS
from .theme import ThemeManager, get_theme_manager
from .stylesheet import generate_stylesheet, generate_tab_widget_style
from .scaling import (
    ScaleManager, get_scale_manager,
    scaled_font, scaled_size, scaled_spacing
)

__all__ = [
    'DARK_COLORS',
    'LIGHT_COLORS',
    'ThemeManager',
    'get_theme_manager',
    'generate_stylesheet',
    'generate_tab_widget_style',
    'ScaleManager',
    'get_scale_manager',
    'scaled_font',
    'scaled_size',
    'scaled_spacing',
]
