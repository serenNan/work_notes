# -*- coding: utf-8 -*-
"""
样式表生成器 - 根据颜色配置生成 QSS, 支持屏幕缩放
"""

from .scaling import scaled_font, scaled_size, scaled_spacing


def generate_stylesheet(colors: dict) -> str:
    """根据配色方案生成全局样式表"""
    # 缩放后的字体和尺寸
    fs_base = scaled_font(14)
    fs_large = scaled_font(15)
    fs_small = scaled_font(13)
    fs_title = scaled_font(14)

    pad_btn = scaled_spacing(10)
    pad_input = scaled_spacing(8)
    radius = scaled_size(8)
    radius_lg = scaled_size(12)
    checkbox_size = scaled_size(18)

    return f"""
/* 全局样式 */
QMainWindow {{
    background-color: {colors['bg_dark']};
}}

QWidget {{
    background-color: transparent;
    color: {colors['text_primary']};
    font-family: "Segoe UI", "Microsoft YaHei UI", "PingFang SC", sans-serif;
    font-size: {fs_base}px;
}}

QLabel {{
    color: {colors['text_primary']};
    background-color: transparent;
}}

/* 按钮基础样式 */
QPushButton {{
    background-color: {colors['bg_card']};
    color: {colors['text_primary']};
    border: 1px solid {colors['border_default']};
    border-radius: {radius}px;
    padding: {pad_btn}px {pad_btn * 2}px;
    font-size: {fs_base}px;
    font-weight: 500;
}}

QPushButton:hover {{
    background-color: {colors['bg_elevated']};
    border-color: {colors['border_hover']};
}}

QPushButton:pressed {{
    background-color: {colors['bg_secondary']};
    border-color: {colors['accent_primary']};
}}

QPushButton:disabled {{
    background-color: {colors['bg_secondary']};
    color: {colors['text_muted']};
    border-color: {colors['border_subtle']};
}}

/* 配置组样式 */
QGroupBox {{
    background-color: {colors['bg_primary']};
    border: 1px solid {colors['border_subtle']};
    border-radius: {radius_lg}px;
    margin-top: {scaled_spacing(14)}px;
    padding: {scaled_spacing(16)}px {scaled_spacing(14)}px {scaled_spacing(14)}px {scaled_spacing(14)}px;
    font-weight: 600;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: {scaled_spacing(14)}px;
    top: {scaled_spacing(4)}px;
    padding: 0 {scaled_spacing(6)}px;
    color: {colors['text_secondary']};
    font-size: {fs_small}px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
}}

/* 复选框样式 */
QCheckBox {{
    color: {colors['text_primary']};
    spacing: {scaled_spacing(8)}px;
    font-size: {fs_base}px;
}}

QCheckBox::indicator {{
    width: {checkbox_size}px;
    height: {checkbox_size}px;
    border-radius: {scaled_size(5)}px;
    border: 2px solid {colors['border_default']};
    background-color: {colors['bg_secondary']};
}}

QCheckBox::indicator:hover {{
    border-color: {colors['accent_primary']};
    background-color: {colors['bg_card']};
}}

QCheckBox::indicator:checked {{
    background-color: {colors['accent_primary']};
    border-color: {colors['accent_primary']};
    image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxMiIgaGVpZ2h0PSIxMiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjMiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBvbHlsaW5lIHBvaW50cz0iMjAgNiA5IDE3IDQgMTIiPjwvcG9seWxpbmU+PC9zdmc+);
}}

/* 下拉框样式 */
QComboBox {{
    background-color: {colors['bg_card']};
    border: 1px solid {colors['border_default']};
    border-radius: {radius}px;
    padding: {pad_input}px {scaled_spacing(28)}px {pad_input}px {scaled_spacing(10)}px;
    color: {colors['text_primary']};
    font-size: {fs_base}px;
    min-width: {scaled_size(100)}px;
}}

QComboBox:hover {{
    border-color: {colors['accent_primary']};
    background-color: {colors['bg_elevated']};
}}

QComboBox:focus {{
    border-color: {colors['accent_primary']};
    outline: none;
}}

QComboBox QAbstractItemView {{
    background-color: {colors['bg_card']};
    border: 1px solid {colors['border_default']};
    border-radius: {radius}px;
    selection-background-color: {colors['accent_primary']};
    selection-color: white;
    padding: {scaled_spacing(4)}px;
    outline: none;
}}

QComboBox QAbstractItemView::item {{
    padding: {pad_input}px {scaled_spacing(10)}px;
    border-radius: {scaled_size(4)}px;
    margin: {scaled_spacing(2)}px;
}}

QComboBox QAbstractItemView::item:hover {{
    background-color: {colors['bg_elevated']};
}}

/* 数字输入框样式 - 隐藏上下箭头 */
QSpinBox {{
    background-color: {colors['bg_card']};
    border: 1px solid {colors['border_default']};
    border-radius: {radius}px;
    padding: {pad_input}px {scaled_spacing(10)}px;
    color: {colors['text_primary']};
    font-size: {fs_base}px;
    min-width: {scaled_size(50)}px;
}}

QSpinBox:hover {{
    border-color: {colors['accent_primary']};
}}

QSpinBox:focus {{
    border-color: {colors['accent_primary']};
}}

QSpinBox::up-button, QSpinBox::down-button {{
    width: 0;
    height: 0;
    border: none;
}}

/* 消息框样式 */
QMessageBox {{
    background-color: {colors['bg_card']};
}}

QMessageBox QLabel {{
    color: {colors['text_primary']};
    font-size: {fs_base}px;
}}

QMessageBox QPushButton {{
    min-width: {scaled_size(80)}px;
    padding: {scaled_spacing(7)}px {scaled_spacing(14)}px;
}}
"""


def generate_tab_widget_style(colors: dict) -> str:
    """生成标签页样式"""
    fs_tab = scaled_font(13)
    pad_tab = scaled_spacing(8)
    radius = scaled_size(8)

    return f"""
QTabWidget::pane {{
    border: 1px solid {colors['border_subtle']};
    border-radius: {radius}px;
    background-color: {colors['bg_primary']};
}}
QTabBar::tab {{
    background-color: {colors['bg_secondary']};
    color: {colors['text_secondary']};
    padding: {pad_tab}px {pad_tab * 2}px;
    margin-right: {scaled_spacing(4)}px;
    border-top-left-radius: {radius}px;
    border-top-right-radius: {radius}px;
    font-size: {fs_tab}px;
    font-weight: 500;
}}
QTabBar::tab:selected {{
    background-color: {colors['bg_primary']};
    color: {colors['text_primary']};
    font-weight: 600;
}}
QTabBar::tab:hover:!selected {{
    background-color: {colors['bg_elevated']};
}}
"""
