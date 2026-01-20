# -*- coding: utf-8 -*-
"""
颜色定义 - Material Design 3 配色方案
支持深色和浅色主题
"""

# Material Design 3 深色主题配色
DARK_COLORS = {
    # Surface 层级 (从低到高)
    'surface_dim': '#121212',
    'surface': '#1e1e1e',
    'surface_bright': '#2c2c2c',
    'surface_container_lowest': '#0f0f0f',
    'surface_container_low': '#1a1a1a',
    'surface_container': '#242424',
    'surface_container_high': '#2e2e2e',
    'surface_container_highest': '#383838',

    # 主色 (蓝色系)
    'primary': '#90CAF9',
    'primary_container': '#004A77',
    'on_primary': '#003258',
    'on_primary_container': '#CDE5FF',

    # 次要色
    'secondary': '#BBC7DB',
    'secondary_container': '#3B4858',
    'on_secondary': '#253140',
    'on_secondary_container': '#D7E3F8',

    # 文字
    'on_surface': '#E3E3E3',
    'on_surface_variant': '#CAC4CF',
    'outline': '#948F99',
    'outline_variant': '#49454F',

    # 状态色
    'error': '#F2B8B5',
    'error_container': '#8C1D18',
    'on_error': '#601410',
    'success': '#A5D6A7',
    'success_container': '#1B5E20',
    'warning': '#FFE082',
    'warning_container': '#F57F17',

    # 向后兼容的别名
    'bg_dark': '#121212',
    'bg_primary': '#1e1e1e',
    'bg_secondary': '#242424',
    'bg_elevated': '#2c2c2c',
    'bg_card': '#2e2e2e',
    'border_subtle': '#383838',
    'border_default': '#49454F',
    'border_hover': '#948F99',
    'text_primary': '#E3E3E3',
    'text_secondary': '#CAC4CF',
    'text_muted': '#948F99',
    'accent_primary': '#90CAF9',
    'accent_secondary': '#64B5F6',
    'accent_glow': '#42A5F5',
    'success_bg': 'rgba(165, 214, 167, 0.12)',
    'success_border': 'rgba(165, 214, 167, 0.3)',
    'drag_bg': 'rgba(144, 202, 249, 0.12)',
}

# Material Design 3 浅色主题配色
LIGHT_COLORS = {
    # Surface 层级 (从低到高)
    'surface_dim': '#DED8E1',
    'surface': '#FFFBFE',
    'surface_bright': '#FFFBFE',
    'surface_container_lowest': '#FFFFFF',
    'surface_container_low': '#F7F2FA',
    'surface_container': '#F3EDF7',
    'surface_container_high': '#ECE6F0',
    'surface_container_highest': '#E6E0E9',

    # 主色 (紫色系)
    'primary': '#6750A4',
    'primary_container': '#EADDFF',
    'on_primary': '#FFFFFF',
    'on_primary_container': '#21005E',

    # 次要色
    'secondary': '#625B71',
    'secondary_container': '#E8DEF8',
    'on_secondary': '#FFFFFF',
    'on_secondary_container': '#1E192B',

    # 文字
    'on_surface': '#1C1B1F',
    'on_surface_variant': '#49454F',
    'outline': '#79747E',
    'outline_variant': '#CAC4D0',

    # 状态色
    'error': '#B3261E',
    'error_container': '#F9DEDC',
    'on_error': '#FFFFFF',
    'success': '#2E7D32',
    'success_container': '#C8E6C9',
    'warning': '#F9A825',
    'warning_container': '#FFF8E1',

    # 向后兼容的别名
    'bg_dark': '#F7F2FA',
    'bg_primary': '#FFFBFE',
    'bg_secondary': '#F3EDF7',
    'bg_elevated': '#ECE6F0',
    'bg_card': '#FFFFFF',
    'border_subtle': '#E6E0E9',
    'border_default': '#CAC4D0',
    'border_hover': '#79747E',
    'text_primary': '#1C1B1F',
    'text_secondary': '#49454F',
    'text_muted': '#79747E',
    'accent_primary': '#6750A4',
    'accent_secondary': '#7E67C3',
    'accent_glow': '#9A82DB',
    'success_bg': 'rgba(46, 125, 50, 0.08)',
    'success_border': 'rgba(46, 125, 50, 0.3)',
    'drag_bg': 'rgba(103, 80, 164, 0.08)',
}

# 窗口尺寸常量
WINDOW_SIZES = {
    'min_width': 1000,
    'min_height': 680,
    'default_width': 1200,
    'default_height': 800,
    'sidebar_collapsed': 64,
    'sidebar_expanded': 220,
    'header_height': 56,
    'status_bar_height': 32,
}
