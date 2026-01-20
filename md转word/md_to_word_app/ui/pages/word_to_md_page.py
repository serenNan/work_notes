# -*- coding: utf-8 -*-
"""
Word 转 MD 页面 - 双栏布局
左侧: 工作面板 (Word 与 Markdown 双向转换)
右侧: 预览面板 (文件内容预览)
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QLabel, QFrame, QFileDialog, QMessageBox, QScrollArea
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QPalette

from ..components import (
    WordDropZone, MarkdownDropZone, SecondaryButton, GradientButton,
    ResultPanel, AnimatedProgressBar
)
from ..components.material_card import MaterialCard
from ..components.preview_panel import PreviewPanel
from ..styles.scaling import scaled_font, scaled_spacing, scaled_size
from core.word_md_bridge import word_to_markdown, markdown_to_word, get_template_source
from core.history_manager import get_history_manager


class WordToMdPage(QWidget):
    """
    Word 转 MD 页面 - 双栏布局
    包含 Word->MD 和 MD->Word 双向转换功能
    """
    conversion_finished = pyqtSignal(bool, str)  # success, message

    def __init__(self, parent=None):
        super().__init__(parent)
        self.word_file_path = None
        self.md_file_path = None
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
        card_layout = card.layout()

        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QScrollArea.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setObjectName('work_scroll_area')
        self._work_scroll_area = scroll_area

        # 滚动区域内容容器
        scroll_content = QWidget()
        scroll_content.setObjectName('scroll_content')
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(scaled_spacing(12))

        # ========== Word -> MD ==========
        self.word_to_md_label = QLabel("Word -> Markdown")
        self.word_to_md_label.setObjectName('section_title')
        layout.addWidget(self.word_to_md_label)

        # Word 文件拖拽区域
        self.word_drop_zone = WordDropZone(hint_text="拖拽 Word 文件到这里")
        self.word_drop_zone.file_dropped.connect(self._on_word_file_dropped)
        layout.addWidget(self.word_drop_zone)

        # Word 文件选择和转换按钮
        word_btn_layout = QHBoxLayout()
        word_btn_layout.setSpacing(scaled_spacing(8))

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
        self.word_to_md_result.setMinimumHeight(scaled_size(80))
        layout.addWidget(self.word_to_md_result)

        # 分隔线
        self.separator = QFrame()
        self.separator.setFrameShape(QFrame.HLine)
        self.separator.setObjectName('section_separator')
        layout.addWidget(self.separator)

        # ========== MD -> Word ==========
        self.md_to_word_label = QLabel("Markdown -> Word")
        self.md_to_word_label.setObjectName('section_title')
        layout.addWidget(self.md_to_word_label)

        # MD 文件拖拽区域
        self.md_drop_zone = MarkdownDropZone(hint_text="拖拽填充好的 MD 文件到这里")
        self.md_drop_zone.file_dropped.connect(self._on_md_file_dropped)
        layout.addWidget(self.md_drop_zone)

        # MD 文件选择和转换按钮
        md_btn_layout = QHBoxLayout()
        md_btn_layout.setSpacing(scaled_spacing(8))

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
        self.md_to_word_result.setMinimumHeight(scaled_size(80))
        layout.addWidget(self.md_to_word_result)

        layout.addStretch()

        scroll_area.setWidget(scroll_content)
        card_layout.addWidget(scroll_area)

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
            self._preview.set_file(output_path)

            # 添加历史记录
            self._history.add_record(
                source_file=self.word_file_path,
                output_file=output_path,
                conversion_type='word_to_md',
                status='success'
            )

            self.conversion_finished.emit(True, f"转换成功: {os.path.basename(output_path)}")
        else:
            # 添加历史记录
            self._history.add_record(
                source_file=self.word_file_path,
                output_file='',
                conversion_type='word_to_md',
                status='error',
                message=result
            )

            self.conversion_finished.emit(False, f"转换失败: {result}")
            QMessageBox.critical(self, "转换失败", result)

    # ========== MD -> Word 事件处理 ==========

    def _on_md_file_dropped(self, file_path: str):
        """MD 文件拖拽"""
        self.md_file_path = file_path
        self.md_to_word_btn.setEnabled(True)
        self.md_to_word_result.clear()
        self._preview.set_file(file_path)

    def _on_select_md_file(self):
        """选择填充好的 MD 文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Markdown 文件", "",
            "Markdown 文件 (*.md);;所有文件 (*.*)"
        )
        if file_path:
            self.md_file_path = file_path
            self.md_drop_zone.set_file(file_path)
            self.md_to_word_btn.setEnabled(True)
            self.md_to_word_result.clear()
            self._preview.set_file(file_path)

    def _on_convert_md_to_word(self):
        """将 MD 内容填充回 Word"""
        if not self.md_file_path:
            QMessageBox.warning(self, "警告", "请先选择 MD 文件")
            return

        if not os.path.exists(self.md_file_path):
            QMessageBox.warning(self, "警告", "MD 文件不存在")
            return

        # 从 MD 文件中获取源 Word 路径
        source_path = get_template_source(self.md_file_path)
        if not source_path:
            QMessageBox.warning(
                self, "警告",
                "MD 文件中未找到源 Word 模板路径\n"
                "请确保 MD 文件包含 <!-- source: ... --> 注释"
            )
            return

        if not os.path.exists(source_path):
            QMessageBox.warning(self, "警告", f"源 Word 模板不存在:\n{source_path}")
            return

        # 生成输出文件路径
        base_name = os.path.splitext(source_path)[0]
        output_path = base_name + '_filled.docx'

        self.md_to_word_btn.setEnabled(False)
        self.md_to_word_btn.setText("填充中...")
        self.m2w_progress.start("填充中...")

        success, result = markdown_to_word(self.md_file_path, source_path, output_path)

        self.md_to_word_btn.setEnabled(True)
        self.md_to_word_btn.setText("填充到 Word")
        self.m2w_progress.stop()

        if success:
            self.md_to_word_result.show_result(output_path, result)

            # 添加历史记录
            self._history.add_record(
                source_file=self.md_file_path,
                output_file=output_path,
                conversion_type='word_to_md',
                status='success'
            )

            self.conversion_finished.emit(True, f"填充成功: {os.path.basename(output_path)}")
        else:
            self.conversion_finished.emit(False, f"填充失败: {result}")
            QMessageBox.critical(self, "填充失败", result)

    def _refresh_preview(self):
        """刷新预览"""
        if self.md_file_path:
            self._preview.set_file(self.md_file_path)

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark
        self._colors = colors

        # 更新卡片主题
        self._work_panel.set_theme(is_dark, colors)
        self._preview_panel.set_theme(is_dark, colors)
        self._preview.set_theme(is_dark, colors)

        # 更新样式
        text_color = colors.get('on_surface', '#E3E3E3')
        border_color = colors.get('outline_variant', '#49454F')

        self.setStyleSheet(f"""
            #panel_title {{
                font-size: {scaled_font(18)}px;
                font-weight: 600;
                color: {text_color};
                background: transparent;
                padding-bottom: {scaled_spacing(12)}px;
            }}
            #section_title {{
                font-size: {scaled_font(14)}px;
                font-weight: 600;
                color: {text_color};
                background: transparent;
                padding-bottom: {scaled_spacing(8)}px;
            }}
            #section_separator {{
                background-color: {border_color};
                max-height: 1px;
                margin: {scaled_spacing(12)}px 0;
            }}
            #work_scroll_area {{
                background: transparent;
                border: none;
            }}
            #scroll_content {{
                background: transparent;
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

    def set_preview_visible(self, visible: bool):
        """显示或隐藏预览面板"""
        self._preview_panel.setVisible(visible)
        if visible:
            # 恢复 55:45 比例
            self._splitter.setSizes([550, 450])

    def is_preview_visible(self) -> bool:
        """预览面板是否可见"""
        return self._preview_panel.isVisible()
