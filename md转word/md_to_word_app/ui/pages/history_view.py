# -*- coding: utf-8 -*-
"""
历史记录页面
显示最近的转换记录
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QPushButton,
    QMessageBox, QMenu, QAction
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize
from PyQt5.QtGui import QColor

from ..styles.scaling import scaled_size, scaled_spacing, scaled_font
from ..styles.icons import render_icon_to_pixmap
from ..components.material_card import MaterialCard
from core.history_manager import get_history_manager, HistoryItem


class HistoryItemWidget(QWidget):
    """历史记录项 Widget"""

    def __init__(self, item: HistoryItem, parent=None):
        super().__init__(parent)
        self._item = item
        self._is_dark = True
        self._setup_ui()

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(
            scaled_spacing(12), scaled_spacing(8),
            scaled_spacing(12), scaled_spacing(8)
        )
        layout.setSpacing(scaled_spacing(12))

        # 图标
        self._icon_label = QLabel()
        self._icon_label.setFixedSize(scaled_size(32), scaled_size(32))
        self._icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self._icon_label)

        # 文件信息
        info_layout = QVBoxLayout()
        info_layout.setSpacing(scaled_spacing(2))

        # 文件名 - 显示输出文件名 (用户关心的是转换结果)
        output_filename = self._item.get_output_filename()
        if not output_filename:
            output_filename = self._item.get_source_filename()
        self._filename_label = QLabel(output_filename)
        self._filename_label.setObjectName('history_filename')
        info_layout.addWidget(self._filename_label)

        # 转换类型和时间
        type_text = "MD -> Word" if self._item.conversion_type == 'md_to_word' else "Word -> MD"
        detail_text = f"{type_text} | {self._item.get_formatted_time()}"
        self._detail_label = QLabel(detail_text)
        self._detail_label.setObjectName('history_detail')
        info_layout.addWidget(self._detail_label)

        layout.addLayout(info_layout, 1)

        # 状态
        self._status_label = QLabel()
        self._status_label.setFixedSize(scaled_size(20), scaled_size(20))
        layout.addWidget(self._status_label)

    def get_item(self) -> HistoryItem:
        return self._item

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark

        # 更新图标
        icon_name = 'md_to_word' if self._item.conversion_type == 'md_to_word' else 'word_to_md'
        icon_color = colors.get('on_surface_variant', '#CAC4CF')
        size = scaled_size(24)
        pixmap = render_icon_to_pixmap(icon_name, size, icon_color)
        self._icon_label.setPixmap(pixmap)

        # 更新状态图标
        if self._item.status == 'success':
            status_icon = 'check'
            status_color = colors.get('success', '#A5D6A7')
        else:
            status_icon = 'error'
            status_color = colors.get('error', '#F2B8B5')

        size = scaled_size(16)
        pixmap = render_icon_to_pixmap(status_icon, size, status_color)
        self._status_label.setPixmap(pixmap)

        # 更新样式
        text_color = colors.get('on_surface', '#E3E3E3')
        muted_color = colors.get('on_surface_variant', '#CAC4CF')

        self._filename_label.setStyleSheet(f"""
            font-size: {scaled_font(14)}px;
            font-weight: 500;
            color: {text_color};
            background: transparent;
        """)

        self._detail_label.setStyleSheet(f"""
            font-size: {scaled_font(12)}px;
            color: {muted_color};
            background: transparent;
        """)


class HistoryView(QWidget):
    """
    历史记录页面
    显示转换历史列表
    """
    file_selected = pyqtSignal(str, str)  # source_file, output_file
    open_file_requested = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._is_dark = True
        self._history_manager = get_history_manager()
        self._history_manager.history_changed.connect(self._refresh_list)

        self._setup_ui()
        self._refresh_list()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            scaled_spacing(16), scaled_spacing(16),
            scaled_spacing(16), scaled_spacing(16)
        )
        layout.setSpacing(scaled_spacing(12))

        # 标题栏
        header = QHBoxLayout()

        title = QLabel("转换历史")
        title.setObjectName('page_title')
        header.addWidget(title)

        header.addStretch()

        # 清空按钮
        self._clear_btn = QPushButton("清空历史")
        self._clear_btn.setObjectName('clear_btn')
        self._clear_btn.clicked.connect(self._on_clear_history)
        header.addWidget(self._clear_btn)

        layout.addLayout(header)

        # 历史列表卡片
        self._card = MaterialCard(elevation=1)

        # 列表
        self._list_widget = QListWidget()
        self._list_widget.setObjectName('history_list')
        self._list_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self._list_widget.customContextMenuRequested.connect(self._show_context_menu)
        self._list_widget.itemDoubleClicked.connect(self._on_item_double_clicked)

        self._card.layout().addWidget(self._list_widget)
        layout.addWidget(self._card, 1)

        # 空状态提示
        self._empty_label = QLabel("暂无转换记录")
        self._empty_label.setObjectName('empty_label')
        self._empty_label.setAlignment(Qt.AlignCenter)
        self._empty_label.hide()
        layout.addWidget(self._empty_label)

    def _refresh_list(self):
        """刷新历史列表"""
        self._list_widget.clear()

        history = self._history_manager.get_history()

        if not history:
            self._list_widget.hide()
            self._card.hide()
            self._empty_label.show()
            return

        self._empty_label.hide()
        self._card.show()
        self._list_widget.show()

        for item in history:
            widget = HistoryItemWidget(item)
            if hasattr(self, '_colors'):
                widget.set_theme(self._is_dark, self._colors)

            list_item = QListWidgetItem()
            list_item.setSizeHint(QSize(0, scaled_size(56)))

            self._list_widget.addItem(list_item)
            self._list_widget.setItemWidget(list_item, widget)

    def _on_clear_history(self):
        """清空历史"""
        reply = QMessageBox.question(
            self, "确认",
            "确定要清空所有历史记录吗?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self._history_manager.clear_history()

    def _show_context_menu(self, pos):
        """显示右键菜单"""
        item = self._list_widget.itemAt(pos)
        if not item:
            return

        widget = self._list_widget.itemWidget(item)
        if not widget:
            return

        history_item = widget.get_item()

        menu = QMenu(self)

        # 打开源文件
        if os.path.exists(history_item.source_file):
            open_source = QAction("打开源文件", self)
            open_source.triggered.connect(
                lambda: self.open_file_requested.emit(history_item.source_file)
            )
            menu.addAction(open_source)

        # 打开输出文件
        if os.path.exists(history_item.output_file):
            open_output = QAction("打开输出文件", self)
            open_output.triggered.connect(
                lambda: self.open_file_requested.emit(history_item.output_file)
            )
            menu.addAction(open_output)

        menu.addSeparator()

        # 删除记录
        delete_action = QAction("删除记录", self)
        delete_action.triggered.connect(
            lambda: self._delete_item(self._list_widget.row(item))
        )
        menu.addAction(delete_action)

        menu.exec_(self._list_widget.mapToGlobal(pos))

    def _delete_item(self, index: int):
        """删除指定记录"""
        self._history_manager.remove_item(index)

    def _on_item_double_clicked(self, item):
        """双击打开文件"""
        widget = self._list_widget.itemWidget(item)
        if widget:
            history_item = widget.get_item()
            if os.path.exists(history_item.output_file):
                self.open_file_requested.emit(history_item.output_file)

    def set_theme(self, is_dark: bool, colors: dict):
        """设置主题"""
        self._is_dark = is_dark
        self._colors = colors

        text_color = colors.get('on_surface', '#E3E3E3')
        muted_color = colors.get('on_surface_variant', '#CAC4CF')
        surface_color = colors.get('surface_container', '#242424')
        border_color = colors.get('outline_variant', '#49454F')
        hover_color = colors.get('surface_container_high', '#2e2e2e')

        self.setStyleSheet(f"""
            #page_title {{
                font-size: {scaled_font(20)}px;
                font-weight: 600;
                color: {text_color};
                background: transparent;
            }}
            #clear_btn {{
                background-color: transparent;
                border: 1px solid {border_color};
                border-radius: {scaled_size(6)}px;
                padding: {scaled_spacing(6)}px {scaled_spacing(16)}px;
                color: {text_color};
                font-size: {scaled_font(13)}px;
            }}
            #clear_btn:hover {{
                background-color: {hover_color};
            }}
            #history_list {{
                background-color: transparent;
                border: none;
                outline: none;
            }}
            #history_list::item {{
                background-color: transparent;
                border-bottom: 1px solid {border_color};
            }}
            #history_list::item:hover {{
                background-color: {hover_color};
            }}
            #history_list::item:selected {{
                background-color: {colors.get('primary_container', '#004A77')};
            }}
            #empty_label {{
                font-size: {scaled_font(14)}px;
                color: {muted_color};
                background: transparent;
            }}
        """)

        self._card.set_theme(is_dark, colors)

        # 更新列表项样式
        for i in range(self._list_widget.count()):
            item = self._list_widget.item(i)
            widget = self._list_widget.itemWidget(item)
            if widget:
                widget.set_theme(is_dark, colors)
