# -*- coding: utf-8 -*-
"""
Word 转 MD 页面 - Word 文档与 Markdown 双向转换
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt

from ..components import (
    WordDropZone, MarkdownDropZone, SecondaryButton, GradientButton,
    ResultPanel, AnimatedProgressBar
)
from ..components.base import ThemedMixin
from ..styles import get_theme_manager, scaled_font, scaled_size, scaled_spacing
from core.word_md_bridge import word_to_markdown, markdown_to_word, get_template_source


class WordToMdPage(QWidget, ThemedMixin):
    """Word 转 MD 页面 - 包含双向转换功能"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_theme()
        self.word_file_path = None
        self.md_file_path = None
        self._setup_ui()
        self._apply_theme(self.colors)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        pad = scaled_spacing(12)
        layout.setContentsMargins(pad, pad, pad, pad)
        layout.setSpacing(scaled_spacing(10))

        # ========== 上半部分: Word -> MD ==========
        self.word_to_md_label = QLabel("Word -> Markdown")
        layout.addWidget(self.word_to_md_label)

        # Word 文件拖拽区域
        self.word_drop_zone = WordDropZone(hint_text="拖拽 Word 文件到这里")
        self.word_drop_zone.file_dropped.connect(self._on_word_file_dropped)
        layout.addWidget(self.word_drop_zone)

        # Word 文件选择按钮和转换按钮
        word_btn_layout = QHBoxLayout()
        word_btn_layout.setSpacing(scaled_spacing(6))
        self.select_word_btn = SecondaryButton("浏览文件...")
        self.select_word_btn.clicked.connect(self._on_select_word_file)
        word_btn_layout.addWidget(self.select_word_btn)
        self.word_to_md_btn = GradientButton("转换为 MD")
        self.word_to_md_btn.clicked.connect(self._on_convert_word_to_md)
        self.word_to_md_btn.setEnabled(False)
        word_btn_layout.addWidget(self.word_to_md_btn)
        layout.addLayout(word_btn_layout)

        # Word->MD 进度条
        self.w2m_progress = AnimatedProgressBar()
        layout.addWidget(self.w2m_progress)

        # Word->MD 结果区域
        self.word_to_md_result = ResultPanel(success_text="转换成功!")
        layout.addWidget(self.word_to_md_result)

        # 分隔线
        self.separator = QFrame()
        self.separator.setFrameShape(QFrame.HLine)
        layout.addWidget(self.separator)

        # ========== 下半部分: MD -> Word ==========
        self.md_to_word_label = QLabel("Markdown -> Word")
        layout.addWidget(self.md_to_word_label)

        # MD 文件拖拽区域
        self.md_drop_zone = MarkdownDropZone(hint_text="拖拽填充好的 MD 文件到这里")
        self.md_drop_zone.file_dropped.connect(self._on_md_file_dropped)
        layout.addWidget(self.md_drop_zone)

        # MD 文件选择按钮和转换按钮
        md_btn_layout = QHBoxLayout()
        md_btn_layout.setSpacing(scaled_spacing(6))
        self.select_md_btn = SecondaryButton("浏览文件...")
        self.select_md_btn.clicked.connect(self._on_select_md_file)
        md_btn_layout.addWidget(self.select_md_btn)
        self.md_to_word_btn = GradientButton("填充到 Word")
        self.md_to_word_btn.clicked.connect(self._on_convert_md_to_word)
        self.md_to_word_btn.setEnabled(False)
        md_btn_layout.addWidget(self.md_to_word_btn)
        layout.addLayout(md_btn_layout)

        # MD->Word 进度条
        self.m2w_progress = AnimatedProgressBar()
        layout.addWidget(self.m2w_progress)

        # MD->Word 结果区域
        self.md_to_word_result = ResultPanel(success_text="填充成功!")
        layout.addWidget(self.md_to_word_result)

        layout.addStretch()

    def _apply_theme(self, colors: dict):
        # 标题标签样式
        fs_title = scaled_font(14)
        title_style = f"""
            color: {colors['text_primary']};
            font-size: {fs_title}px;
            font-weight: 700;
            background: transparent;
        """
        self.word_to_md_label.setStyleSheet(title_style)
        self.md_to_word_label.setStyleSheet(title_style)

        # 分隔线样式
        self.separator.setStyleSheet(
            f"background-color: {colors['border_default']}; max-height: 2px;"
        )

    # ========== Word -> MD 事件处理 ==========

    def _on_word_file_dropped(self, file_path: str):
        """Word 文件拖拽"""
        self.word_file_path = file_path
        self.word_to_md_btn.setEnabled(True)
        self.word_to_md_result.clear()

    def _on_select_word_file(self):
        """选择 Word 文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Word 文件", "",
            "Word 文件 (*.docx);;所有文件 (*.*)"
        )
        if file_path:
            self.word_file_path = file_path
            self.word_drop_zone.set_file(file_path)
            self.word_to_md_btn.setEnabled(True)
            self.word_to_md_result.clear()

    def _on_convert_word_to_md(self):
        """将 Word 转换为 MD"""
        if not self.word_file_path:
            QMessageBox.warning(self, "警告", "请先选择 Word 文件")
            return

        if not os.path.exists(self.word_file_path):
            QMessageBox.warning(self, "警告", "Word 文件不存在")
            return

        # 生成输出文件路径
        base_name = os.path.splitext(self.word_file_path)[0]
        output_path = base_name + '.md'

        self.word_to_md_btn.setEnabled(False)
        self.word_to_md_btn.setText("转换中...")
        self.w2m_progress.start("转换中...")

        success, result = word_to_markdown(self.word_file_path, output_path)

        self.word_to_md_btn.setEnabled(True)
        self.word_to_md_btn.setText("转换为 MD")
        self.w2m_progress.stop()

        if success:
            self.word_to_md_result.show_result(output_path)
        else:
            QMessageBox.critical(self, "转换失败", result)

    # ========== MD -> Word 事件处理 ==========

    def _on_md_file_dropped(self, file_path: str):
        """MD 文件拖拽"""
        self.md_file_path = file_path
        self.md_to_word_btn.setEnabled(True)
        self.md_to_word_result.clear()

    def _on_select_md_file(self):
        """选择填充好的 MD 文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Markdown File", "",
            "Markdown Files (*.md);;All Files (*.*)"
        )
        if file_path:
            self.md_file_path = file_path
            self.md_drop_zone.set_file(file_path)
            self.md_to_word_btn.setEnabled(True)
            self.md_to_word_result.clear()

    def _on_convert_md_to_word(self):
        """将 MD 内容填充回 Word"""
        if not self.md_file_path:
            QMessageBox.warning(self, "Warning", "Please select an MD file first")
            return

        if not os.path.exists(self.md_file_path):
            QMessageBox.warning(self, "Warning", "MD file does not exist")
            return

        # 从 MD 文件中获取源 Word 路径
        source_path = get_template_source(self.md_file_path)
        if not source_path:
            QMessageBox.warning(
                self, "Warning",
                "Source Word template path not found in MD file\n"
                "Please ensure MD file contains <!-- source: ... --> comment"
            )
            return

        if not os.path.exists(source_path):
            QMessageBox.warning(self, "Warning", f"Source Word template not found:\n{source_path}")
            return

        # 生成输出文件路径
        base_name = os.path.splitext(source_path)[0]
        output_path = base_name + '_filled.docx'

        self.md_to_word_btn.setEnabled(False)
        self.md_to_word_btn.setText("Filling...")
        self.m2w_progress.start("Filling...")

        success, result = markdown_to_word(self.md_file_path, source_path, output_path)

        self.md_to_word_btn.setEnabled(True)
        self.md_to_word_btn.setText("Fill to Word")
        self.m2w_progress.stop()

        if success:
            self.md_to_word_result.show_result(output_path, result)
        else:
            QMessageBox.critical(self, "Fill Failed", result)
