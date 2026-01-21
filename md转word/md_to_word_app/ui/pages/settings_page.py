# -*- coding: utf-8 -*-
"""
设置页面 - 应用配置界面
"""

import os
import webbrowser
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QScrollArea,
    QLabel, QComboBox, QSpinBox, QCheckBox, QPushButton,
    QGroupBox, QGridLayout, QLineEdit, QFileDialog,
    QMessageBox, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal

from ..components.material_card import MaterialCard
from ..components.buttons import SecondaryButton, GradientButton
from ..styles.scaling import scaled_spacing, scaled_font, scaled_size
from core.settings_manager import get_settings_manager
from core.history_manager import get_history_manager


class SettingsSection(QGroupBox):
    """设置分组"""

    def __init__(self, title: str, parent=None):
        super().__init__(title, parent)
        self._colors = {}

    def set_theme(self, is_dark: bool, colors: dict):
        self._colors = colors
        radius = scaled_size(8)
        fs_title = scaled_font(13)
        self.setStyleSheet(f"""
            QGroupBox {{
                background-color: {colors['bg_primary']};
                border: 1px solid {colors['border_subtle']};
                border-radius: {radius}px;
                margin-top: {scaled_spacing(12)}px;
                padding-top: {scaled_spacing(8)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {scaled_spacing(12)}px;
                padding: 0 {scaled_spacing(6)}px;
                color: {colors['text_primary']};
                font-size: {fs_title}px;
                font-weight: 600;
            }}
        """)


class SettingsPage(QWidget):
    """设置页面"""

    # 信号
    theme_change_requested = pyqtSignal(bool)  # is_dark
    history_cleared = pyqtSignal()

    VERSION = "v2.0.0"
    GITHUB_URL = "https://github.com/example/md-to-word"

    def __init__(self, parent=None):
        super().__init__(parent)
        self._settings = get_settings_manager()
        self._history = get_history_manager()
        self._colors = {}
        self._is_dark = True
        self._setup_ui()

    def _setup_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(
            scaled_spacing(16), scaled_spacing(16),
            scaled_spacing(16), scaled_spacing(16)
        )

        # 使用 MaterialCard 包裹
        card = MaterialCard(elevation=1)
        self._card = card
        card_layout = card.layout()

        # 滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll = scroll

        # 滚动内容
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(scaled_spacing(16))
        layout.setContentsMargins(scaled_spacing(16), scaled_spacing(8), scaled_spacing(16), scaled_spacing(16))

        # 页面标题
        title = QLabel("Settings")
        title.setObjectName('page_title')
        layout.addWidget(title)
        self._title = title

        # 外观设置
        self._appearance_section = self._create_appearance_section()
        layout.addWidget(self._appearance_section)

        # 转换默认值
        self._convert_section = self._create_convert_section()
        layout.addWidget(self._convert_section)

        # 路径设置
        self._path_section = self._create_path_section()
        layout.addWidget(self._path_section)

        # 历史记录设置
        self._history_section = self._create_history_section()
        layout.addWidget(self._history_section)

        # 关于
        self._about_section = self._create_about_section()
        layout.addWidget(self._about_section)

        layout.addStretch()

        scroll.setWidget(content)
        card_layout.addWidget(scroll)
        main_layout.addWidget(card)

    def _create_appearance_section(self) -> SettingsSection:
        """创建外观设置分组"""
        section = SettingsSection("外观设置")
        layout = QGridLayout(section)
        layout.setSpacing(scaled_spacing(12))
        layout.setContentsMargins(scaled_spacing(16), scaled_spacing(24), scaled_spacing(16), scaled_spacing(16))

        # 主题
        row = 0
        theme_label = QLabel("主题")
        layout.addWidget(theme_label, row, 0)
        self._labels = [theme_label]

        self._theme_combo = QComboBox()
        self._theme_combo.addItems(['深色', '浅色'])
        self._theme_combo.setCurrentIndex(0 if self._settings.is_dark_theme else 1)
        self._theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        self._theme_combo.setFixedWidth(scaled_size(150))
        layout.addWidget(self._theme_combo, row, 1)

        # 预览面板默认显示
        row += 1
        preview_label = QLabel("预览面板默认显示")
        layout.addWidget(preview_label, row, 0)
        self._labels.append(preview_label)

        self._preview_check = QCheckBox()
        self._preview_check.setChecked(self._settings.preview_visible)
        self._preview_check.stateChanged.connect(self._on_preview_changed)
        layout.addWidget(self._preview_check, row, 1)

        layout.setColumnStretch(2, 1)
        return section

    def _create_convert_section(self) -> SettingsSection:
        """创建转换默认值分组"""
        section = SettingsSection("转换默认值")
        layout = QGridLayout(section)
        layout.setSpacing(scaled_spacing(12))
        layout.setContentsMargins(scaled_spacing(16), scaled_spacing(24), scaled_spacing(16), scaled_spacing(16))

        row = 0

        # 生成目录
        toc_label = QLabel("生成目录")
        layout.addWidget(toc_label, row, 0)
        self._labels.append(toc_label)

        self._toc_check = QCheckBox()
        self._toc_check.setChecked(self._settings.default_toc_enabled)
        self._toc_check.stateChanged.connect(self._on_toc_changed)
        layout.addWidget(self._toc_check, row, 1)

        # 目录深度
        toc_depth_label = QLabel("目录深度")
        layout.addWidget(toc_depth_label, row, 2)
        self._labels.append(toc_depth_label)

        self._toc_depth_spin = QSpinBox()
        self._toc_depth_spin.setRange(1, 6)
        self._toc_depth_spin.setValue(self._settings.default_toc_depth)
        self._toc_depth_spin.valueChanged.connect(self._on_toc_depth_changed)
        self._toc_depth_spin.setFixedWidth(scaled_size(80))
        layout.addWidget(self._toc_depth_spin, row, 3)

        # 中文字体
        row += 1
        cn_font_label = QLabel("中文字体")
        layout.addWidget(cn_font_label, row, 0)
        self._labels.append(cn_font_label)

        self._cn_font_combo = QComboBox()
        self._cn_font_combo.addItems(['宋体', '黑体', '微软雅黑', '楷体', '仿宋', '华文中宋'])
        self._cn_font_combo.setCurrentText(self._settings.default_chinese_font)
        self._cn_font_combo.currentTextChanged.connect(self._on_cn_font_changed)
        self._cn_font_combo.setFixedWidth(scaled_size(150))
        layout.addWidget(self._cn_font_combo, row, 1)

        # 代码字体
        code_font_label = QLabel("代码字体")
        layout.addWidget(code_font_label, row, 2)
        self._labels.append(code_font_label)

        self._code_font_combo = QComboBox()
        self._code_font_combo.addItems(['Times New Roman', 'Consolas', 'Courier New', 'Source Code Pro', 'Monaco', 'Fira Code'])
        self._code_font_combo.setCurrentText(self._settings.default_code_font)
        self._code_font_combo.currentTextChanged.connect(self._on_code_font_changed)
        self._code_font_combo.setFixedWidth(scaled_size(150))
        layout.addWidget(self._code_font_combo, row, 3)

        # 字体大小
        row += 1
        font_size_label = QLabel("字体大小")
        layout.addWidget(font_size_label, row, 0)
        self._labels.append(font_size_label)

        self._font_size_combo = QComboBox()
        self._font_size_combo.addItems(['10', '10.5', '11', '12', '14', '16'])
        self._font_size_combo.setCurrentText(self._settings.default_font_size)
        self._font_size_combo.currentTextChanged.connect(self._on_font_size_changed)
        self._font_size_combo.setFixedWidth(scaled_size(150))
        layout.addWidget(self._font_size_combo, row, 1)

        # 行间距
        line_spacing_label = QLabel("行间距")
        layout.addWidget(line_spacing_label, row, 2)
        self._labels.append(line_spacing_label)

        self._line_spacing_combo = QComboBox()
        self._line_spacing_combo.addItems(['1.0', '1.15', '1.5', '2.0'])
        self._line_spacing_combo.setCurrentText(self._settings.default_line_spacing)
        self._line_spacing_combo.currentTextChanged.connect(self._on_line_spacing_changed)
        self._line_spacing_combo.setFixedWidth(scaled_size(150))
        layout.addWidget(self._line_spacing_combo, row, 3)

        # 代码高亮风格
        row += 1
        highlight_label = QLabel("代码高亮风格")
        layout.addWidget(highlight_label, row, 0)
        self._labels.append(highlight_label)

        self._highlight_combo = QComboBox()
        self._highlight_combo.addItems(['tango', 'pygments', 'espresso', 'monochrome', 'kate', 'zenburn'])
        self._highlight_combo.setCurrentText(self._settings.default_highlight_style)
        self._highlight_combo.currentTextChanged.connect(self._on_highlight_changed)
        self._highlight_combo.setFixedWidth(scaled_size(150))
        layout.addWidget(self._highlight_combo, row, 1)

        layout.setColumnStretch(4, 1)
        return section

    def _create_path_section(self) -> SettingsSection:
        """创建路径设置分组"""
        section = SettingsSection("路径设置")
        layout = QGridLayout(section)
        layout.setSpacing(scaled_spacing(12))
        layout.setContentsMargins(scaled_spacing(16), scaled_spacing(24), scaled_spacing(16), scaled_spacing(16))

        row = 0

        # Pandoc 路径
        pandoc_label = QLabel("Pandoc 路径")
        layout.addWidget(pandoc_label, row, 0)
        self._labels.append(pandoc_label)

        path_layout = QHBoxLayout()
        path_layout.setSpacing(scaled_spacing(8))

        self._pandoc_edit = QLineEdit()
        self._pandoc_edit.setText(self._settings.pandoc_path)
        self._pandoc_edit.textChanged.connect(self._on_pandoc_path_changed)
        path_layout.addWidget(self._pandoc_edit)

        self._pandoc_browse_btn = SecondaryButton("浏览...")
        self._pandoc_browse_btn.clicked.connect(self._browse_pandoc)
        self._pandoc_browse_btn.setFixedWidth(scaled_size(80))
        path_layout.addWidget(self._pandoc_browse_btn)

        layout.addLayout(path_layout, row, 1, 1, 3)

        # 默认输出目录
        row += 1
        output_label = QLabel("默认输出目录")
        layout.addWidget(output_label, row, 0)
        self._labels.append(output_label)

        output_layout = QHBoxLayout()
        output_layout.setSpacing(scaled_spacing(8))

        self._output_edit = QLineEdit()
        self._output_edit.setText(self._settings.default_output_dir)
        self._output_edit.setPlaceholderText("留空则与源文件同目录")
        self._output_edit.textChanged.connect(self._on_output_dir_changed)
        output_layout.addWidget(self._output_edit)

        self._output_browse_btn = SecondaryButton("浏览...")
        self._output_browse_btn.clicked.connect(self._browse_output_dir)
        self._output_browse_btn.setFixedWidth(scaled_size(80))
        output_layout.addWidget(self._output_browse_btn)

        layout.addLayout(output_layout, row, 1, 1, 3)

        return section

    def _create_history_section(self) -> SettingsSection:
        """创建历史记录设置分组"""
        section = SettingsSection("历史记录")
        layout = QGridLayout(section)
        layout.setSpacing(scaled_spacing(12))
        layout.setContentsMargins(scaled_spacing(16), scaled_spacing(24), scaled_spacing(16), scaled_spacing(16))

        row = 0

        # 最大保留条数
        max_label = QLabel("最大保留条数")
        layout.addWidget(max_label, row, 0)
        self._labels.append(max_label)

        self._max_history_spin = QSpinBox()
        self._max_history_spin.setRange(10, 500)
        self._max_history_spin.setValue(self._settings.max_history_count)
        self._max_history_spin.valueChanged.connect(self._on_max_history_changed)
        self._max_history_spin.setFixedWidth(scaled_size(100))
        layout.addWidget(self._max_history_spin, row, 1)

        # 清除历史记录按钮
        self._clear_history_btn = SecondaryButton("清除所有历史记录")
        self._clear_history_btn.clicked.connect(self._on_clear_history)
        layout.addWidget(self._clear_history_btn, row, 2)

        layout.setColumnStretch(3, 1)
        return section

    def _create_about_section(self) -> SettingsSection:
        """创建关于分组"""
        section = SettingsSection("关于")
        layout = QVBoxLayout(section)
        layout.setSpacing(scaled_spacing(12))
        layout.setContentsMargins(scaled_spacing(16), scaled_spacing(24), scaled_spacing(16), scaled_spacing(16))

        # 版本信息
        version_label = QLabel(f"MD to Word {self.VERSION}")
        version_label.setObjectName('about_version')
        layout.addWidget(version_label)
        self._version_label = version_label

        # 描述
        desc_label = QLabel("Markdown 转 Word 文档工具, 基于 Pandoc 转换引擎")
        desc_label.setObjectName('about_desc')
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)
        self._desc_label = desc_label

        # 按钮行
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(scaled_spacing(12))

        self._github_btn = SecondaryButton("GitHub")
        self._github_btn.clicked.connect(self._open_github)
        btn_layout.addWidget(self._github_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        return section

    # ========== 事件处理 ==========

    def _on_theme_changed(self, index: int):
        is_dark = (index == 0)
        self._settings.is_dark_theme = is_dark
        self.theme_change_requested.emit(is_dark)

    def _on_preview_changed(self, state: int):
        self._settings.preview_visible = (state == Qt.Checked)

    def _on_toc_changed(self, state: int):
        self._settings.default_toc_enabled = (state == Qt.Checked)

    def _on_toc_depth_changed(self, value: int):
        self._settings.default_toc_depth = value

    def _on_cn_font_changed(self, font: str):
        self._settings.default_chinese_font = font

    def _on_code_font_changed(self, font: str):
        self._settings.default_code_font = font

    def _on_font_size_changed(self, size: str):
        self._settings.default_font_size = size

    def _on_line_spacing_changed(self, spacing: str):
        self._settings.default_line_spacing = spacing

    def _on_highlight_changed(self, style: str):
        self._settings.default_highlight_style = style

    def _on_pandoc_path_changed(self, path: str):
        self._settings.pandoc_path = path

    def _on_output_dir_changed(self, path: str):
        self._settings.default_output_dir = path

    def _on_max_history_changed(self, value: int):
        self._settings.max_history_count = value

    def _browse_pandoc(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Pandoc 可执行文件", "",
            "可执行文件 (*.exe);;所有文件 (*.*)"
        )
        if file_path:
            self._pandoc_edit.setText(file_path)

    def _browse_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "选择默认输出目录"
        )
        if dir_path:
            self._output_edit.setText(dir_path)

    def _on_clear_history(self):
        reply = QMessageBox.question(
            self, "确认清除",
            "确定要清除所有历史记录吗?\n此操作不可撤销。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self._history.clear_history()
            self.history_cleared.emit()
            QMessageBox.information(self, "完成", "历史记录已清除")

    def _open_github(self):
        webbrowser.open(self.GITHUB_URL)

    # ========== 主题 ==========

    def set_theme(self, is_dark: bool, colors: dict):
        self._is_dark = is_dark
        self._colors = colors

        # 更新卡片主题
        self._card.set_theme(is_dark, colors)

        # 更新标题样式
        fs_title = scaled_font(18)
        self._title.setStyleSheet(f"""
            color: {colors['text_primary']};
            font-size: {fs_title}px;
            font-weight: 600;
            background: transparent;
        """)

        # 更新分组主题
        for section in [self._appearance_section, self._convert_section,
                        self._path_section, self._history_section, self._about_section]:
            section.set_theme(is_dark, colors)

        # 更新标签样式
        fs_label = scaled_font(12)
        label_style = f"""
            color: {colors['text_secondary']};
            font-size: {fs_label}px;
            background: transparent;
        """
        for label in self._labels:
            label.setStyleSheet(label_style)

        # 更新版本和描述样式
        fs_version = scaled_font(14)
        self._version_label.setStyleSheet(f"""
            color: {colors['text_primary']};
            font-size: {fs_version}px;
            font-weight: 600;
            background: transparent;
        """)

        fs_desc = scaled_font(12)
        self._desc_label.setStyleSheet(f"""
            color: {colors['text_secondary']};
            font-size: {fs_desc}px;
            background: transparent;
        """)

        # 更新输入框样式
        fs_input = scaled_font(12)
        radius = scaled_size(4)
        pad = scaled_spacing(6)
        input_style = f"""
            QLineEdit {{
                background-color: {colors['bg_secondary']};
                color: {colors['text_primary']};
                border: 1px solid {colors['border_subtle']};
                border-radius: {radius}px;
                padding: {pad}px;
                font-size: {fs_input}px;
            }}
            QLineEdit:focus {{
                border-color: {colors['primary']};
            }}
        """
        self._pandoc_edit.setStyleSheet(input_style)
        self._output_edit.setStyleSheet(input_style)

        # 更新滚动区域样式
        self._scroll.setStyleSheet(f"""
            QScrollArea {{
                background: transparent;
                border: none;
            }}
            QScrollBar:vertical {{
                background: {colors['bg_secondary']};
                width: {scaled_size(8)}px;
                border-radius: {scaled_size(4)}px;
            }}
            QScrollBar::handle:vertical {{
                background: {colors['border_subtle']};
                border-radius: {scaled_size(4)}px;
                min-height: {scaled_size(30)}px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {colors['text_muted']};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
        """)

        # 同步主题选择器
        self._theme_combo.blockSignals(True)
        self._theme_combo.setCurrentIndex(0 if is_dark else 1)
        self._theme_combo.blockSignals(False)

    def refresh_settings(self):
        """刷新设置显示 (从设置管理器重新加载)"""
        self._theme_combo.setCurrentIndex(0 if self._settings.is_dark_theme else 1)
        self._preview_check.setChecked(self._settings.preview_visible)
        self._toc_check.setChecked(self._settings.default_toc_enabled)
        self._toc_depth_spin.setValue(self._settings.default_toc_depth)
        self._cn_font_combo.setCurrentText(self._settings.default_chinese_font)
        self._code_font_combo.setCurrentText(self._settings.default_code_font)
        self._font_size_combo.setCurrentText(self._settings.default_font_size)
        self._line_spacing_combo.setCurrentText(self._settings.default_line_spacing)
        self._highlight_combo.setCurrentText(self._settings.default_highlight_style)
        self._pandoc_edit.setText(self._settings.pandoc_path)
        self._output_edit.setText(self._settings.default_output_dir)
        self._max_history_spin.setValue(self._settings.max_history_count)
