# -*- coding: utf-8 -*-
"""
MD 转 Word 页面 - Markdown 到 Word 文档转换
"""

import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QFileDialog, QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal

from ..components import (
    MarkdownDropZone, SecondaryButton, GradientButton,
    OptionsPanel, ResultPanel, AnimatedProgressBar
)
from ..styles import scaled_spacing
from core.converter import ConverterService


class ConverterThread(QThread):
    """异步转换线程"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, input_file, output_file, options):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.options = options
        self.converter = ConverterService()

    def run(self):
        try:
            self.progress.emit("Converting...")
            success, message = self.converter.convert(
                self.input_file,
                self.output_file,
                self.options
            )
            if success:
                self.finished.emit(self.output_file)
            else:
                self.error.emit(message)
        except Exception as e:
            self.error.emit(str(e))


class MdToWordPage(QWidget):
    """MD 转 Word 页面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_file = None
        self.converter_thread = None
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        pad = scaled_spacing(12)
        layout.setContentsMargins(pad, pad, pad, pad)
        layout.setSpacing(scaled_spacing(10))

        # 拖拽区域
        self.drop_zone = MarkdownDropZone(hint_text="拖拽 Markdown 文件到这里")
        self.drop_zone.file_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_zone)

        # 选择文件按钮
        self.select_btn = SecondaryButton("浏览文件...")
        self.select_btn.clicked.connect(self._on_select_file)
        layout.addWidget(self.select_btn)

        # 配置选项组
        self.options_panel = OptionsPanel()
        layout.addWidget(self.options_panel)

        # 转换按钮
        self.convert_btn = GradientButton("开始转换")
        self.convert_btn.clicked.connect(self._on_convert)
        self.convert_btn.setEnabled(False)
        layout.addWidget(self.convert_btn)

        # 进度条
        self.progress_bar = AnimatedProgressBar()
        layout.addWidget(self.progress_bar)

        # 结果区域
        self.result_panel = ResultPanel(success_text="转换成功!")
        layout.addWidget(self.result_panel)

        layout.addStretch()

    def _on_file_dropped(self, file_path: str):
        """文件拖拽处理"""
        self.current_file = file_path
        self.convert_btn.setEnabled(True)
        self.result_panel.clear()

    def _on_select_file(self):
        """选择文件处理"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Markdown 文件", "",
            "Markdown 文件 (*.md *.markdown);;所有文件 (*.*)"
        )
        if file_path:
            self.current_file = file_path
            self.drop_zone.set_file(file_path)
            self.convert_btn.setEnabled(True)
            self.result_panel.clear()

    def _on_convert(self):
        """开始转换处理"""
        if not self.current_file:
            QMessageBox.warning(self, "警告", "请先选择要转换的文件")
            return

        if not os.path.exists(self.current_file):
            QMessageBox.warning(self, "警告", "文件不存在或已被删除")
            return

        output_file = os.path.splitext(self.current_file)[0] + '.docx'
        options = self.options_panel.get_options()

        self.convert_btn.setEnabled(False)
        self.convert_btn.setText("转换中...")
        self.progress_bar.start("正在处理...")
        self.result_panel.clear()

        self.converter_thread = ConverterThread(self.current_file, output_file, options)
        self.converter_thread.progress.connect(self._on_progress)
        self.converter_thread.finished.connect(self._on_finished)
        self.converter_thread.error.connect(self._on_error)
        self.converter_thread.start()

    def _on_progress(self, message: str):
        """进度更新"""
        self.progress_bar.start(message)

    def _on_finished(self, output_file: str):
        """转换完成"""
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("开始转换")
        self.progress_bar.stop()
        self.result_panel.show_result(output_file)

    def _on_error(self, error_message: str):
        """转换错误"""
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("开始转换")
        self.progress_bar.stop()
        QMessageBox.critical(self, "转换失败", error_message)
