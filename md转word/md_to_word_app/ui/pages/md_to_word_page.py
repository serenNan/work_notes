# -*- coding: utf-8 -*-
"""
MD 转 Word 页面 - 双栏布局
左侧: 工作面板 (文件选择、配置、转换)
右侧: 预览面板 (Markdown 源码预览)
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QFileDialog, QMessageBox, QLabel
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt

from ..components import (
    MarkdownDropZone, SecondaryButton, GradientButton,
    OptionsPanel, ResultPanel, AnimatedProgressBar
)
from ..components.material_card import MaterialCard
from ..components.preview_panel import PreviewPanel
from ..styles.scaling import scaled_spacing, scaled_font
from core.converter import ConverterService
from core.history_manager import get_history_manager


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
    """
    MD 转 Word 页面
    双栏布局: 工作面板 | 预览面板
    """
    conversion_finished = pyqtSignal(bool, str)  # success, message

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_file = None
        self.converter_thread = None
        self._is_dark = True
        self._colors = {}
        self._history = get_history_manager()
        self._setup_ui()

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(
            scaled_spacing(16), scaled_spacing(16),
            scaled_spacing(16), scaled_spacing(16)
        )
        layout.setSpacing(scaled_spacing(16))

        # 使用 QSplitter 实现可调整的双栏
        self._splitter = QSplitter(Qt.Horizontal)
        self._splitter.setHandleWidth(scaled_spacing(8))

        # 左侧: 工作面板
        self._work_panel = self._create_work_panel()
        self._splitter.addWidget(self._work_panel)

        # 右侧: 预览面板
        self._preview_panel = self._create_preview_panel()
        self._splitter.addWidget(self._preview_panel)

        # 设置初始比例 55:45
        self._splitter.setSizes([550, 450])

        layout.addWidget(self._splitter)

    def _create_work_panel(self) -> QWidget:
        """创建工作面板"""
        card = MaterialCard(elevation=1)
        layout = card.layout()

        # 标题
        title = QLabel("Markdown to Word")
        title.setObjectName('panel_title')
        layout.addWidget(title)

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

        return card

    def _create_preview_panel(self) -> QWidget:
        """创建预览面板"""
        card = MaterialCard(elevation=1)
        layout = card.layout()

        # 标题
        title = QLabel("预览")
        title.setObjectName('panel_title')
        layout.addWidget(title)

        # 预览区域
        self._preview = PreviewPanel()
        self._preview.refresh_requested.connect(self._refresh_preview)
        layout.addWidget(self._preview, 1)

        return card

    def _on_file_dropped(self, file_path: str):
        """文件拖拽处理"""
        self.current_file = file_path
        self.convert_btn.setEnabled(True)
        self.result_panel.clear()
        self._preview.set_file(file_path)

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
            self._preview.set_file(file_path)

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

        # 添加历史记录
        self._history.add_record(
            source_file=self.current_file,
            output_file=output_file,
            conversion_type='md_to_word',
            status='success'
        )

        self.conversion_finished.emit(True, f"转换成功: {os.path.basename(output_file)}")

    def _on_error(self, error_message: str):
        """转换错误"""
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("开始转换")
        self.progress_bar.stop()

        # 添加历史记录
        if self.current_file:
            self._history.add_record(
                source_file=self.current_file,
                output_file='',
                conversion_type='md_to_word',
                status='error',
                message=error_message
            )

        self.conversion_finished.emit(False, f"转换失败: {error_message}")
        QMessageBox.critical(self, "转换失败", error_message)

    def _refresh_preview(self):
        """刷新预览"""
        if self.current_file:
            self._preview.set_file(self.current_file)

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark
        self._colors = colors

        # 更新卡片主题
        self._work_panel.set_theme(is_dark, colors)
        self._preview_panel.set_theme(is_dark, colors)
        self._preview.set_theme(is_dark, colors)

        # 更新标题样式
        text_color = colors.get('on_surface', '#E3E3E3')
        self.setStyleSheet(f"""
            #panel_title {{
                font-size: {scaled_font(18)}px;
                font-weight: 600;
                color: {text_color};
                background: transparent;
                padding-bottom: {scaled_spacing(12)}px;
            }}
        """)

        # 更新 splitter 样式
        handle_color = colors.get('outline_variant', '#49454F')
        self._splitter.setStyleSheet(f"""
            QSplitter::handle {{
                background-color: {handle_color};
                border-radius: 2px;
            }}
            QSplitter::handle:hover {{
                background-color: {colors.get('outline', '#948F99')};
            }}
        """)
