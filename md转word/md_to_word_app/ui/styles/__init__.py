# -*- coding: utf-8 -*-
"""
样式模块 - 提供主题管理, 颜色配置和样式表生成
"""

from .colors import DARK_COLORS, LIGHT_COLORS, WINDOW_SIZES
from .theme import ThemeManager, get_theme_manager
from .stylesheet import generate_stylesheet, generate_tab_widget_style
from .scaling import (
    ScaleManager, get_scale_manager,
    scaled_font, scaled_size, scaled_spacing
)
from .elevation import apply_elevation, ElevatedWidget, ELEVATION_LEVELS
from .icons import ICONS, get_icon_svg, get_icon_data_uri

__all__ = [
    # Colors
    'DARK_COLORS',
    'LIGHT_COLORS',
    'WINDOW_SIZES',
    # Theme
    'ThemeManager',
    'get_theme_manager',
    # Stylesheet
    'generate_stylesheet',
    'generate_tab_widget_style',
    # Scaling
    'ScaleManager',
    'get_scale_manager',
    'scaled_font',
    'scaled_size',
    'scaled_spacing',
    # Elevation
    'apply_elevation',
    'ElevatedWidget',
    'ELEVATION_LEVELS',
    # Icons
    'ICONS',
    'get_icon_svg',
    'get_icon_data_uri',
]
