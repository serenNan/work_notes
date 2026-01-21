# -*- coding: utf-8 -*-
"""
选项面板组件 - MD 转 Word 转换选项配置
"""

import os
from PyQt5.QtWidgets import (
    QGroupBox, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QCheckBox, QComboBox, QSpinBox, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QPixmap
from .base import ThemedMixin
from ..styles import get_theme_manager, scaled_font, scaled_size, scaled_spacing
from core.settings_manager import get_settings_manager


class OptionsPanel(QGroupBox, ThemedMixin):
    """
    转换选项面板

    包含目录生成, 字体设置, 代码高亮等配置选项
    """
    options_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__("转换选项", parent)
        self._settings = get_settings_manager()
        self._init_theme()
        self._setup_ui()
        self._apply_theme(self.colors)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(scaled_spacing(10))
        layout.setContentsMargins(scaled_spacing(12), scaled_spacing(16), scaled_spacing(12), scaled_spacing(12))

        # 生成目录 - 使用设置中的默认值
        self.toc_checkbox = QCheckBox("生成目录")
        self.toc_checkbox.setChecked(self._settings.default_toc_enabled)
        self.toc_checkbox.stateChanged.connect(self._on_toc_changed)
        layout.addWidget(self.toc_checkbox)

        # 目录深度 - 使用设置中的默认值
        toc_depth_layout = QHBoxLayout()
        toc_depth_layout.setSpacing(scaled_spacing(8))
        self.toc_depth_label = QLabel("目录深度")
        toc_depth_layout.addWidget(self.toc_depth_label)
        self.toc_depth_spin = QSpinBox()
        self.toc_depth_spin.setRange(1, 6)
        self.toc_depth_spin.setValue(self._settings.default_toc_depth)
        self.toc_depth_spin.setFixedWidth(scaled_size(60))
        self.toc_depth_spin.valueChanged.connect(self._emit_changed)
        toc_depth_layout.addWidget(self.toc_depth_spin)
        toc_depth_layout.addStretch()
        layout.addLayout(toc_depth_layout)

        # 分割线 1
        self.separator1 = QFrame()
        self.separator1.setFrameShape(QFrame.HLine)
        layout.addWidget(self.separator1)

        # 字体设置 - 2x2 网格布局
        font_grid = QGridLayout()
        font_grid.setSpacing(scaled_spacing(8))
        font_grid.setColumnStretch(1, 1)
        font_grid.setColumnStretch(3, 1)

        # 第一行: 中文字体 | 代码字体 - 使用设置中的默认值
        self.chinese_font_label = QLabel("中文字体")
        font_grid.addWidget(self.chinese_font_label, 0, 0)
        self.chinese_font_combo = QComboBox()
        self.chinese_font_combo.addItems(['宋体', '黑体', '微软雅黑', '楷体', '仿宋', '华文中宋'])
        self.chinese_font_combo.setCurrentText(self._settings.default_chinese_font)
        self.chinese_font_combo.currentTextChanged.connect(self._emit_changed)
        font_grid.addWidget(self.chinese_font_combo, 0, 1)

        self.code_font_label = QLabel("代码字体")
        font_grid.addWidget(self.code_font_label, 0, 2)
        self.code_font_combo = QComboBox()
        self.code_font_combo.addItems(['Times New Roman', 'Consolas', 'Courier New', 'Source Code Pro', 'Monaco', 'Fira Code'])
        self.code_font_combo.setCurrentText(self._settings.default_code_font)
        self.code_font_combo.currentTextChanged.connect(self._emit_changed)
        font_grid.addWidget(self.code_font_combo, 0, 3)

        # 第二行: 字体大小 | 行间距 - 使用设置中的默认值
        self.font_size_label = QLabel("字体大小")
        font_grid.addWidget(self.font_size_label, 1, 0)
        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(['10', '10.5', '11', '12', '14', '16'])
        self.font_size_combo.setCurrentText(self._settings.default_font_size)
        self.font_size_combo.currentTextChanged.connect(self._emit_changed)
        font_grid.addWidget(self.font_size_combo, 1, 1)

        self.line_spacing_label = QLabel("行间距")
        font_grid.addWidget(self.line_spacing_label, 1, 2)
        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems(['1.0', '1.15', '1.5', '2.0'])
        self.line_spacing_combo.setCurrentText(self._settings.default_line_spacing)
        self.line_spacing_combo.currentTextChanged.connect(self._emit_changed)
        font_grid.addWidget(self.line_spacing_combo, 1, 3)

        layout.addLayout(font_grid)

        # 分割线 2
        self.separator2 = QFrame()
        self.separator2.setFrameShape(QFrame.HLine)
        layout.addWidget(self.separator2)

        # 代码高亮 - 使用设置中的默认值
        highlight_layout = QHBoxLayout()
        highlight_layout.setSpacing(scaled_spacing(8))
        self.highlight_label = QLabel("代码高亮")
        highlight_layout.addWidget(self.highlight_label)
        self.highlight_combo = QComboBox()
        self.highlight_combo.addItems(['tango', 'pygments', 'espresso', 'monochrome', 'kate', 'zenburn'])
        self.highlight_combo.setCurrentText(self._settings.default_highlight_style)
        self.highlight_combo.setFixedWidth(scaled_size(100))
        self.highlight_combo.currentTextChanged.connect(self._on_highlight_changed)
        highlight_layout.addWidget(self.highlight_combo)
        highlight_layout.addStretch()
        layout.addLayout(highlight_layout)

        # 代码高亮预览图
        self.highlight_preview = QLabel()
        self.highlight_preview.setAlignment(Qt.AlignCenter)
        self.highlight_preview.setMinimumHeight(scaled_size(60))
        layout.addWidget(self.highlight_preview)
        self._update_highlight_preview(self._settings.default_highlight_style)

    def _apply_theme(self, colors: dict):
        # 组框样式
        radius = scaled_size(8)
        self.setStyleSheet(f"""
            QGroupBox {{
                background-color: {colors['bg_primary']};
                border: 1px solid {colors['border_subtle']};
                border-radius: {radius}px;
            }}
        """)

        # 标签样式
        fs_label = scaled_font(12)
        label_style = f"""
            color: {colors['text_secondary']};
            font-size: {fs_label}px;
            background: transparent;
        """
        for label in [self.toc_depth_label, self.chinese_font_label,
                      self.code_font_label, self.font_size_label,
                      self.line_spacing_label, self.highlight_label]:
            label.setStyleSheet(label_style)

        # 分割线样式
        separator_style = f"background-color: {colors['border_subtle']}; max-height: 1px;"
        self.separator1.setStyleSheet(separator_style)
        self.separator2.setStyleSheet(separator_style)

        # 高亮预览框样式
        preview_radius = scaled_size(5)
        preview_pad = scaled_spacing(3)
        self.highlight_preview.setStyleSheet(f"""
            background-color: {colors['bg_secondary']};
            border: 1px solid {colors['border_subtle']};
            border-radius: {preview_radius}px;
            padding: {preview_pad}px;
        """)

    def _on_toc_changed(self):
        """目录选项变更"""
        self.toc_depth_spin.setEnabled(self.toc_checkbox.isChecked())
        self._emit_changed()

    def _on_highlight_changed(self, style: str):
        """高亮样式变更"""
        self._update_highlight_preview(style)
        self._emit_changed()

    def _update_highlight_preview(self, style: str):
        """更新高亮预览图"""
        # 获取图片路径
        base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        # 处理 zenburn 拼写错误的文件名
        filename = 'zneburn.png' if style == 'zenburn' else f'{style}.png'
        image_path = os.path.join(base_dir, 'assets', '代码高亮图', filename)

        if os.path.exists(image_path):
            from PyQt5.QtWidgets import QApplication

            # 获取设备像素比
            dpr = 1.0
            app = QApplication.instance()
            if app:
                screen = app.primaryScreen()
                if screen:
                    dpr = screen.devicePixelRatio()

            pixmap = QPixmap(image_path)

            # 逻辑尺寸
            max_width = scaled_size(320)
            max_height = scaled_size(150)

            # 物理尺寸 (考虑高 DPI)
            physical_width = int(max_width * dpr)
            physical_height = int(max_height * dpr)

            # 缩放到物理尺寸, 保持宽高比
            scaled_pixmap = pixmap.scaledToWidth(physical_width, Qt.SmoothTransformation)
            if scaled_pixmap.height() > physical_height:
                scaled_pixmap = pixmap.scaledToHeight(physical_height, Qt.SmoothTransformation)

            # 设置设备像素比, 使显示清晰
            scaled_pixmap.setDevicePixelRatio(dpr)
            self.highlight_preview.setPixmap(scaled_pixmap)
        else:
            self.highlight_preview.setText(f"Preview not found: {style}")

    def _emit_changed(self):
        """发送选项变更信号"""
        self.options_changed.emit()

    def get_options(self) -> dict:
        """获取当前配置选项"""
        return {
            'generate_toc': self.toc_checkbox.isChecked(),
            'toc_depth': self.toc_depth_spin.value(),
            'highlight_style': self.highlight_combo.currentText(),
            'chinese_font': self.chinese_font_combo.currentText(),
            'code_font': self.code_font_combo.currentText(),
            'font_size': float(self.font_size_combo.currentText()),
            'line_spacing': float(self.line_spacing_combo.currentText())
        }
