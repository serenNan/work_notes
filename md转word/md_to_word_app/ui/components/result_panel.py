# -*- coding: utf-8 -*-
"""
结果面板组件 - 统一的转换结果显示
"""

import os
import sys
import subprocess
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QLabel
from PyQt5.QtCore import Qt
from .base import ThemedFrame
from .buttons import SecondaryButton
from ..styles import scaled_font, scaled_size, scaled_spacing


class ResultPanel(ThemedFrame):
    """
    转换结果面板

    显示成功状态, 输出文件名, 以及打开文件/文件夹按钮
    """

    def __init__(self, success_text="转换成功!", parent=None):
        super().__init__(parent)
        self.success_text_str = success_text
        self.output_file = None
        self._setup_ui()
        self._apply_theme(self.colors)
        self.hide()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        pad = scaled_spacing(10)
        layout.setContentsMargins(pad, pad, pad, pad)
        layout.setSpacing(scaled_spacing(6))

        # 成功标题行
        header_layout = QHBoxLayout()
        header_layout.setAlignment(Qt.AlignCenter)

        icon_size = scaled_size(18)
        self.success_icon = QLabel("OK")
        self.success_icon.setFixedSize(icon_size, icon_size)
        self.success_icon.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(self.success_icon)

        self.success_label = QLabel(self.success_text_str)
        header_layout.addWidget(self.success_label)

        layout.addLayout(header_layout)

        # 结果文件标签
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label)

        # 按钮行
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(scaled_spacing(6))

        self.open_file_btn = SecondaryButton("打开文件")
        self.open_file_btn.clicked.connect(self._on_open_file)
        btn_layout.addWidget(self.open_file_btn)

        self.open_folder_btn = SecondaryButton("打开文件夹")
        self.open_folder_btn.clicked.connect(self._on_open_folder)
        btn_layout.addWidget(self.open_folder_btn)

        layout.addLayout(btn_layout)

    def _apply_theme(self, colors: dict):
        # 面板背景
        radius = scaled_size(8)
        self.setStyleSheet(f"""
            ResultPanel {{
                background-color: {colors['success_bg']};
                border: 1px solid {colors['success_border']};
                border-radius: {radius}px;
            }}
        """)

        # 成功图标
        icon_size = scaled_size(18)
        icon_fs = scaled_font(8)
        self.success_icon.setStyleSheet(f"""
            background-color: {colors['success']};
            color: white;
            font-size: {icon_fs}px;
            font-weight: bold;
            border-radius: {icon_size // 2}px;
        """)

        # 成功文本
        fs_success = scaled_font(12)
        self.success_label.setStyleSheet(f"""
            color: {colors['success']};
            font-size: {fs_success}px;
            font-weight: 600;
            background: transparent;
        """)

        # 结果标签
        fs_result = scaled_font(11)
        self.result_label.setStyleSheet(f"""
            color: {colors['text_secondary']};
            font-size: {fs_result}px;
            background: transparent;
        """)

    def show_result(self, output_file: str, extra_info: str = ""):
        """
        显示转换结果

        Args:
            output_file: 输出文件路径
            extra_info: 额外信息 (可选)
        """
        self.output_file = output_file
        filename = os.path.basename(output_file)
        text = f"输出: {filename}"
        if extra_info:
            text += f"\n{extra_info}"
        self.result_label.setText(text)
        self.show()

    def clear(self):
        """清除结果并隐藏面板"""
        self.output_file = None
        self.result_label.setText("")
        self.hide()

    def _on_open_file(self):
        """打开输出文件"""
        if self.output_file and os.path.exists(self.output_file):
            if sys.platform == 'win32':
                os.startfile(self.output_file)
            elif sys.platform == 'darwin':
                subprocess.run(['open', self.output_file])
            else:
                subprocess.run(['xdg-open', self.output_file])

    def _on_open_folder(self):
        """打开输出文件所在文件夹"""
        if self.output_file and os.path.exists(self.output_file):
            if sys.platform == 'win32':
                subprocess.run(f'explorer /select,"{self.output_file}"', shell=True)
            elif sys.platform == 'darwin':
                subprocess.run(['open', '-R', self.output_file])
            else:
                subprocess.run(['xdg-open', os.path.dirname(self.output_file)])
