# -*- coding: utf-8 -*-
"""
主窗口 UI - 支持浅色/深色主题切换和双标签页
"""

import os
import subprocess
import sys
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QCheckBox, QComboBox, QSpinBox,
    QGroupBox, QFileDialog, QMessageBox, QFrame, QTabWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QFont, QPixmap

# 导入转换服务
from core.converter import ConverterService
from core.word_md_bridge import word_to_markdown, markdown_to_word, get_template_source


# 深色主题配色
DARK_COLORS = {
    'bg_dark': '#0d0d0d',
    'bg_primary': '#141414',
    'bg_secondary': '#1a1a1a',
    'bg_elevated': '#1f1f1f',
    'bg_card': '#242424',
    'border_subtle': '#2a2a2a',
    'border_default': '#333333',
    'border_hover': '#404040',
    'text_primary': '#f5f5f5',
    'text_secondary': '#a0a0a0',
    'text_muted': '#666666',
    'accent_primary': '#3b82f6',
    'accent_secondary': '#60a5fa',
    'accent_glow': '#2563eb',
    'success': '#22c55e',
    'success_bg': 'rgba(34, 197, 94, 0.08)',
    'success_border': 'rgba(34, 197, 94, 0.3)',
    'error': '#ef4444',
    'drag_bg': 'rgba(59, 130, 246, 0.1)',
}

# 浅色主题配色
LIGHT_COLORS = {
    'bg_dark': '#f8fafc',
    'bg_primary': '#ffffff',
    'bg_secondary': '#f1f5f9',
    'bg_elevated': '#e2e8f0',
    'bg_card': '#ffffff',
    'border_subtle': '#e2e8f0',
    'border_default': '#cbd5e1',
    'border_hover': '#94a3b8',
    'text_primary': '#1e293b',
    'text_secondary': '#64748b',
    'text_muted': '#94a3b8',
    'accent_primary': '#2563eb',
    'accent_secondary': '#3b82f6',
    'accent_glow': '#1d4ed8',
    'success': '#16a34a',
    'success_bg': 'rgba(22, 163, 74, 0.08)',
    'success_border': 'rgba(22, 163, 74, 0.3)',
    'error': '#dc2626',
    'drag_bg': 'rgba(37, 99, 235, 0.08)',
}


def generate_stylesheet(colors):
    """根据配色方案生成样式表"""
    return f"""
/* 全局样式 */
QMainWindow {{
    background-color: {colors['bg_dark']};
}}

QWidget {{
    background-color: transparent;
    color: {colors['text_primary']};
    font-family: "Segoe UI", "Microsoft YaHei UI", "PingFang SC", sans-serif;
    font-size: 15px;
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
    border-radius: 8px;
    padding: 10px 20px;
    font-size: 15px;
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
    border-radius: 12px;
    margin-top: 16px;
    padding: 20px 16px 16px 16px;
    font-weight: 600;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 16px;
    top: 4px;
    padding: 0 8px;
    color: {colors['text_secondary']};
    font-size: 14px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
}}

/* 复选框样式 */
QCheckBox {{
    color: {colors['text_primary']};
    spacing: 10px;
    font-size: 15px;
}}

QCheckBox::indicator {{
    width: 20px;
    height: 20px;
    border-radius: 6px;
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
    border-radius: 8px;
    padding: 8px 30px 8px 12px;
    color: {colors['text_primary']};
    min-width: 120px;
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
    border-radius: 8px;
    selection-background-color: {colors['accent_primary']};
    selection-color: white;
    padding: 4px;
    outline: none;
}}

QComboBox QAbstractItemView::item {{
    padding: 8px 12px;
    border-radius: 4px;
    margin: 2px;
}}

QComboBox QAbstractItemView::item:hover {{
    background-color: {colors['bg_elevated']};
}}

/* 数字输入框样式 - 隐藏上下箭头 */
QSpinBox {{
    background-color: {colors['bg_card']};
    border: 1px solid {colors['border_default']};
    border-radius: 8px;
    padding: 8px 12px;
    color: {colors['text_primary']};
    min-width: 60px;
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
    font-size: 15px;
}}

QMessageBox QPushButton {{
    min-width: 90px;
    padding: 8px 16px;
}}
"""


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
            self.progress.emit("正在转换...")
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


class FileDropZone(QFrame):
    """通用拖拽区域组件"""
    file_dropped = pyqtSignal(str)

    def __init__(self, colors, icon_text="FILE", hint_text="拖拽文件到这里", extensions=None, parent=None):
        super().__init__(parent)
        self.colors = colors
        self.has_file = False
        self.icon_text = icon_text
        self.hint_text = hint_text
        self.extensions = extensions or []
        self.setAcceptDrops(True)
        self.setup_ui()

    def setup_ui(self):
        self.setMinimumHeight(120)
        self.setFrameStyle(QFrame.NoFrame)
        self.apply_normal_style()

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(8)
        layout.setContentsMargins(16, 16, 16, 16)

        # 图标容器
        self.icon_container = QFrame()
        self.icon_container.setFixedSize(56, 56)
        self.update_icon_style()
        icon_layout = QVBoxLayout(self.icon_container)
        icon_layout.setAlignment(Qt.AlignCenter)
        icon_layout.setContentsMargins(0, 0, 0, 0)

        self.icon_label = QLabel(self.icon_text)
        self.icon_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.icon_label.setStyleSheet("color: white; background: transparent;")
        self.icon_label.setAlignment(Qt.AlignCenter)
        icon_layout.addWidget(self.icon_label)

        icon_h_layout = QHBoxLayout()
        icon_h_layout.addStretch()
        icon_h_layout.addWidget(self.icon_container)
        icon_h_layout.addStretch()
        layout.addLayout(icon_h_layout)

        self.hint_label = QLabel(self.hint_text)
        self.hint_label.setAlignment(Qt.AlignCenter)
        self.hint_label.setStyleSheet(f"""
            color: {self.colors['text_primary']};
            font-size: 15px;
            font-weight: 500;
            background: transparent;
        """)
        layout.addWidget(self.hint_label)

        self.file_label = QLabel("")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setStyleSheet(f"""
            color: {self.colors['accent_secondary']};
            font-size: 13px;
            font-weight: 600;
            background: transparent;
        """)
        self.file_label.hide()
        layout.addWidget(self.file_label)

    def update_icon_style(self):
        self.icon_container.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 {self.colors['accent_primary']}, stop:1 {self.colors['accent_glow']});
                border-radius: 16px;
            }}
        """)

    def update_theme(self, colors):
        self.colors = colors
        self.update_icon_style()
        self.hint_label.setStyleSheet(f"""
            color: {colors['text_primary']};
            font-size: 15px;
            font-weight: 500;
            background: transparent;
        """)
        self.file_label.setStyleSheet(f"""
            color: {colors['accent_secondary']};
            font-size: 13px;
            font-weight: 600;
            background: transparent;
        """)
        if self.has_file:
            self.apply_file_selected_style()
        else:
            self.apply_normal_style()

    def apply_normal_style(self):
        self.setStyleSheet(f"""
            FileDropZone {{
                background-color: {self.colors['bg_primary']};
                border: 2px dashed {self.colors['border_default']};
                border-radius: 12px;
            }}
            FileDropZone:hover {{
                border-color: {self.colors['accent_primary']};
                background-color: {self.colors['bg_secondary']};
            }}
        """)

    def apply_drag_style(self):
        self.setStyleSheet(f"""
            FileDropZone {{
                background-color: {self.colors['drag_bg']};
                border: 2px dashed {self.colors['accent_primary']};
                border-radius: 12px;
            }}
        """)

    def apply_file_selected_style(self):
        self.setStyleSheet(f"""
            FileDropZone {{
                background-color: {self.colors['bg_primary']};
                border: 2px solid {self.colors['accent_primary']};
                border-radius: 12px;
            }}
        """)

    def _check_extension(self, file_path):
        return file_path.lower().endswith(tuple(self.extensions))

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1:
                file_path = urls[0].toLocalFile()
                if self._check_extension(file_path):
                    event.acceptProposedAction()
                    self.apply_drag_style()
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        if self.has_file:
            self.apply_file_selected_style()
        else:
            self.apply_normal_style()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if len(urls) == 1:
            file_path = urls[0].toLocalFile()
            if self._check_extension(file_path):
                self.set_file(file_path)
                self.file_dropped.emit(file_path)
                event.acceptProposedAction()
                return
        event.ignore()
        self.apply_normal_style()

    def set_file(self, file_path):
        if file_path:
            self.has_file = True
            filename = os.path.basename(file_path)
            self.file_label.setText(filename)
            self.file_label.show()
            self.hint_label.setText("已选择文件")
            self.apply_file_selected_style()
        else:
            self.has_file = False
            self.file_label.hide()
            self.hint_label.setText(self.hint_text)
            self.apply_normal_style()


# 为了向后兼容保留 DropZone 别名
class DropZone(FileDropZone):
    def __init__(self, colors, parent=None):
        super().__init__(colors, icon_text="MD", hint_text="拖拽 Markdown 文件到这里",
                         extensions=['.md', '.markdown'], parent=parent)


class GradientButton(QPushButton):
    """渐变主按钮"""

    def __init__(self, text, colors, parent=None):
        super().__init__(text, parent)
        self.colors = colors
        self.setFixedHeight(44)
        self.setCursor(Qt.PointingHandCursor)
        self.apply_style()

    def apply_style(self):
        self.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {self.colors['accent_primary']}, stop:1 {self.colors['accent_glow']});
                color: white;
                font-size: 15px;
                font-weight: 600;
                border: none;
                border-radius: 8px;
                padding: 10px 20px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {self.colors['accent_secondary']}, stop:1 {self.colors['accent_primary']});
            }}
            QPushButton:disabled {{
                background: {self.colors['bg_card']};
                color: {self.colors['text_muted']};
            }}
        """)

    def update_theme(self, colors):
        self.colors = colors
        self.apply_style()


class SecondaryButton(QPushButton):
    """次要按钮"""

    def __init__(self, text, colors, parent=None):
        super().__init__(text, parent)
        self.colors = colors
        self.setFixedHeight(36)
        self.setCursor(Qt.PointingHandCursor)
        self.apply_style()

    def apply_style(self):
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.colors['bg_card']};
                color: {self.colors['text_primary']};
                font-size: 13px;
                font-weight: 500;
                border: 1px solid {self.colors['border_default']};
                border-radius: 6px;
                padding: 6px 12px;
            }}
            QPushButton:hover {{
                background-color: {self.colors['bg_elevated']};
                border-color: {self.colors['accent_primary']};
            }}
        """)

    def update_theme(self, colors):
        self.colors = colors
        self.apply_style()


class ThemeToggleButton(QPushButton):
    """主题切换按钮"""

    def __init__(self, is_dark, colors, parent=None):
        super().__init__(parent)
        self.is_dark = is_dark
        self.colors = colors
        self.setFixedSize(36, 36)
        self.setCursor(Qt.PointingHandCursor)
        self.update_icon()

    def update_icon(self):
        tooltip = "切换到浅色主题" if self.is_dark else "切换到深色主题"
        self.setToolTip(tooltip)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.colors['bg_card']};
                border: 1px solid {self.colors['border_default']};
                border-radius: 18px;
                font-size: 14px;
            }}
            QPushButton:hover {{
                background-color: {self.colors['bg_elevated']};
                border-color: {self.colors['accent_primary']};
            }}
        """)
        self.setText("dark" if self.is_dark else "light")

    def set_dark_mode(self, is_dark, colors):
        self.is_dark = is_dark
        self.colors = colors
        self.update_icon()


class MainWindow(QMainWindow):
    """主窗口 - 支持主题切换和双标签页"""

    def __init__(self):
        super().__init__()
        self.current_file = None
        self.converter_thread = None
        # Word转MD 相关
        self.word_template_path = None
        self.md_file_path = None

        # 加载用户主题偏好
        self.settings = QSettings("MD2Word", "MarkdownToWord")
        self.is_dark_theme = self.settings.value("dark_theme", True, type=bool)
        self.colors = DARK_COLORS if self.is_dark_theme else LIGHT_COLORS

        self.setup_ui()
        self.apply_theme()

    def apply_theme(self):
        self.colors = DARK_COLORS if self.is_dark_theme else LIGHT_COLORS
        self.setStyleSheet(generate_stylesheet(self.colors))
        self.central_widget.setStyleSheet(f"background-color: {self.colors['bg_dark']};")
        self.update_component_themes()

    def update_component_themes(self):
        # 标题
        self.title_label.setStyleSheet(f"""
            font-size: 24px;
            font-weight: 700;
            color: {self.colors['text_primary']};
            background: transparent;
        """)

        # 标签页样式
        self.tab_widget.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {self.colors['border_subtle']};
                border-radius: 8px;
                background-color: {self.colors['bg_primary']};
            }}
            QTabBar::tab {{
                background-color: {self.colors['bg_secondary']};
                color: {self.colors['text_secondary']};
                padding: 10px 20px;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-weight: 500;
            }}
            QTabBar::tab:selected {{
                background-color: {self.colors['bg_primary']};
                color: {self.colors['text_primary']};
                font-weight: 600;
            }}
            QTabBar::tab:hover:!selected {{
                background-color: {self.colors['bg_elevated']};
            }}
        """)

        # MD转Word标签页组件更新
        self.drop_zone.update_theme(self.colors)
        self.select_btn.update_theme(self.colors)
        self.convert_btn.update_theme(self.colors)
        self.open_file_btn.update_theme(self.colors)
        self.open_folder_btn.update_theme(self.colors)
        self.theme_btn.set_dark_mode(self.is_dark_theme, self.colors)

        # 选项组
        self.options_group.setStyleSheet(f"""
            QGroupBox {{
                background-color: {self.colors['bg_primary']};
                border: 1px solid {self.colors['border_subtle']};
                border-radius: 10px;
            }}
        """)

        # 标签
        for label in [self.toc_depth_label, self.chinese_font_label, self.code_font_label, self.font_size_label, self.line_spacing_label, self.highlight_label]:
            label.setStyleSheet(f"""
                color: {self.colors['text_secondary']};
                font-size: 14px;
                background: transparent;
            """)

        self.separator.setStyleSheet(f"background-color: {self.colors['border_subtle']}; max-height: 1px;")
        self.separator2.setStyleSheet(f"background-color: {self.colors['border_subtle']}; max-height: 1px;")

        # 高亮预览
        self.highlight_preview.setStyleSheet(f"""
            background-color: {self.colors['bg_secondary']};
            border: 1px solid {self.colors['border_subtle']};
            border-radius: 6px;
            padding: 4px;
        """)

        self.status_label.setStyleSheet(f"""
            color: {self.colors['text_muted']};
            font-size: 13px;
            background: transparent;
        """)

        # 结果区域
        self.result_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['success_bg']};
                border: 1px solid {self.colors['success_border']};
                border-radius: 10px;
            }}
        """)
        self.success_icon.setStyleSheet(f"""
            background-color: {self.colors['success']};
            color: white;
            font-size: 9px;
            font-weight: bold;
            border-radius: 10px;
        """)
        self.success_text.setStyleSheet(f"""
            color: {self.colors['success']};
            font-size: 14px;
            font-weight: 600;
            background: transparent;
        """)
        self.result_label.setStyleSheet(f"""
            color: {self.colors['text_secondary']};
            font-size: 12px;
            background: transparent;
        """)

        # Word转MD标签页组件更新
        self.word_drop_zone.update_theme(self.colors)
        self.md_drop_zone.update_theme(self.colors)
        self.select_word_btn.update_theme(self.colors)
        self.select_md_btn.update_theme(self.colors)
        self.word_to_md_btn.update_theme(self.colors)
        self.md_to_word_btn.update_theme(self.colors)
        self.w2m_open_file_btn.update_theme(self.colors)
        self.w2m_open_folder_btn.update_theme(self.colors)
        self.m2w_open_file_btn.update_theme(self.colors)
        self.m2w_open_folder_btn.update_theme(self.colors)

        # Word转MD标签页标签更新
        for label in [self.word_to_md_label, self.md_to_word_label]:
            label.setStyleSheet(f"""
                color: {self.colors['text_primary']};
                font-size: 16px;
                font-weight: 700;
                background: transparent;
            """)

        # 分隔线
        self.tab2_separator.setStyleSheet(f"background-color: {self.colors['border_default']}; max-height: 2px;")

    def toggle_theme(self):
        self.is_dark_theme = not self.is_dark_theme
        self.settings.setValue("dark_theme", self.is_dark_theme)
        self.apply_theme()

    def setup_ui(self):
        self.setWindowTitle("Markdown to Word")
        self.setMinimumSize(520, 900)
        self.resize(540, 1050)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(16)

        # 顶部栏
        top_bar = QHBoxLayout()
        self.title_label = QLabel("MD to Word")
        top_bar.addWidget(self.title_label)
        top_bar.addStretch()
        self.theme_btn = ThemeToggleButton(self.is_dark_theme, self.colors)
        self.theme_btn.clicked.connect(self.toggle_theme)
        top_bar.addWidget(self.theme_btn)
        main_layout.addLayout(top_bar)

        # 创建标签页
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {self.colors['border_subtle']};
                border-radius: 8px;
                background-color: {self.colors['bg_primary']};
            }}
            QTabBar::tab {{
                background-color: {self.colors['bg_secondary']};
                color: {self.colors['text_secondary']};
                padding: 10px 20px;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-weight: 500;
            }}
            QTabBar::tab:selected {{
                background-color: {self.colors['bg_primary']};
                color: {self.colors['text_primary']};
                font-weight: 600;
            }}
            QTabBar::tab:hover:!selected {{
                background-color: {self.colors['bg_elevated']};
            }}
        """)

        # 创建 MD转Word 标签页
        self.md_to_word_tab = QWidget()
        self.setup_md_to_word_tab()
        self.tab_widget.addTab(self.md_to_word_tab, "MD转Word")

        # 创建 Word转MD 标签页
        self.word_to_md_tab = QWidget()
        self.setup_word_to_md_tab()
        self.tab_widget.addTab(self.word_to_md_tab, "Word转MD")

        main_layout.addWidget(self.tab_widget)

    def setup_md_to_word_tab(self):
        """设置 MD转Word 标签页"""
        tab_layout = QVBoxLayout(self.md_to_word_tab)
        tab_layout.setContentsMargins(16, 16, 16, 16)
        tab_layout.setSpacing(12)

        # 拖拽区域
        self.drop_zone = DropZone(self.colors)
        self.drop_zone.file_dropped.connect(self.on_file_dropped)
        tab_layout.addWidget(self.drop_zone)

        # 选择文件按钮
        self.select_btn = SecondaryButton("浏览文件...", self.colors)
        self.select_btn.clicked.connect(self.on_select_file)
        tab_layout.addWidget(self.select_btn)

        # 配置选项组
        self.options_group = QGroupBox("转换选项")
        options_layout = QVBoxLayout(self.options_group)
        options_layout.setSpacing(12)
        options_layout.setContentsMargins(14, 20, 14, 14)

        # 生成目录
        self.toc_checkbox = QCheckBox("生成目录")
        self.toc_checkbox.setChecked(True)
        self.toc_checkbox.stateChanged.connect(self.on_toc_changed)
        options_layout.addWidget(self.toc_checkbox)

        # 目录深度
        toc_depth_layout = QHBoxLayout()
        toc_depth_layout.setSpacing(10)
        self.toc_depth_label = QLabel("目录深度")
        toc_depth_layout.addWidget(self.toc_depth_label)
        self.toc_depth_spin = QSpinBox()
        self.toc_depth_spin.setRange(1, 6)
        self.toc_depth_spin.setValue(3)
        self.toc_depth_spin.setFixedWidth(70)
        toc_depth_layout.addWidget(self.toc_depth_spin)
        toc_depth_layout.addStretch()
        options_layout.addLayout(toc_depth_layout)

        # 分割线
        self.separator = QFrame()
        self.separator.setFrameShape(QFrame.HLine)
        options_layout.addWidget(self.separator)

        # 中文字体
        chinese_font_layout = QHBoxLayout()
        chinese_font_layout.setSpacing(10)
        self.chinese_font_label = QLabel("中文字体")
        chinese_font_layout.addWidget(self.chinese_font_label)
        self.chinese_font_combo = QComboBox()
        self.chinese_font_combo.addItems(['宋体', '黑体', '微软雅黑', '楷体', '仿宋', '华文中宋'])
        self.chinese_font_combo.setFixedWidth(120)
        chinese_font_layout.addWidget(self.chinese_font_combo)
        chinese_font_layout.addStretch()
        options_layout.addLayout(chinese_font_layout)

        # 代码字体
        code_font_layout = QHBoxLayout()
        code_font_layout.setSpacing(10)
        self.code_font_label = QLabel("代码字体")
        code_font_layout.addWidget(self.code_font_label)
        self.code_font_combo = QComboBox()
        self.code_font_combo.addItems(['Times New Roman', 'Consolas', 'Courier New', 'Source Code Pro', 'Monaco', 'Fira Code'])
        self.code_font_combo.setFixedWidth(150)
        code_font_layout.addWidget(self.code_font_combo)
        code_font_layout.addStretch()
        options_layout.addLayout(code_font_layout)

        # 字体大小
        font_size_layout = QHBoxLayout()
        font_size_layout.setSpacing(10)
        self.font_size_label = QLabel("字体大小")
        font_size_layout.addWidget(self.font_size_label)
        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(['10', '10.5', '11', '12', '14', '16'])
        self.font_size_combo.setCurrentText('12')
        self.font_size_combo.setFixedWidth(80)
        font_size_layout.addWidget(self.font_size_combo)
        font_size_layout.addStretch()
        options_layout.addLayout(font_size_layout)

        # 行间距
        line_spacing_layout = QHBoxLayout()
        line_spacing_layout.setSpacing(10)
        self.line_spacing_label = QLabel("行间距")
        line_spacing_layout.addWidget(self.line_spacing_label)
        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems(['1.0', '1.15', '1.5', '2.0'])
        self.line_spacing_combo.setCurrentText('1.5')
        self.line_spacing_combo.setFixedWidth(80)
        line_spacing_layout.addWidget(self.line_spacing_combo)
        line_spacing_layout.addStretch()
        options_layout.addLayout(line_spacing_layout)

        # 分割线2
        self.separator2 = QFrame()
        self.separator2.setFrameShape(QFrame.HLine)
        options_layout.addWidget(self.separator2)

        # 代码高亮
        highlight_layout = QHBoxLayout()
        highlight_layout.setSpacing(10)
        self.highlight_label = QLabel("代码高亮")
        highlight_layout.addWidget(self.highlight_label)
        self.highlight_combo = QComboBox()
        self.highlight_combo.addItems(['tango', 'pygments', 'espresso', 'monochrome', 'kate', 'zenburn'])
        self.highlight_combo.setFixedWidth(120)
        self.highlight_combo.currentTextChanged.connect(self.on_highlight_changed)
        highlight_layout.addWidget(self.highlight_combo)
        highlight_layout.addStretch()
        options_layout.addLayout(highlight_layout)

        # 代码高亮预览图
        self.highlight_preview = QLabel()
        self.highlight_preview.setAlignment(Qt.AlignCenter)
        self.highlight_preview.setStyleSheet(f"""
            background-color: {self.colors['bg_secondary']};
            border: 1px solid {self.colors['border_subtle']};
            border-radius: 6px;
            padding: 4px;
        """)
        self.highlight_preview.setMinimumHeight(80)
        options_layout.addWidget(self.highlight_preview)
        self.update_highlight_preview('tango')

        tab_layout.addWidget(self.options_group)

        # 转换按钮
        self.convert_btn = GradientButton("开始转换", self.colors)
        self.convert_btn.clicked.connect(self.on_convert)
        self.convert_btn.setEnabled(False)
        tab_layout.addWidget(self.convert_btn)

        # 状态标签
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        tab_layout.addWidget(self.status_label)

        # 结果区域
        self.result_frame = QFrame()
        result_layout = QVBoxLayout(self.result_frame)
        result_layout.setContentsMargins(12, 12, 12, 12)
        result_layout.setSpacing(8)

        success_header = QHBoxLayout()
        success_header.setAlignment(Qt.AlignCenter)
        self.success_icon = QLabel("OK")
        self.success_icon.setFixedSize(20, 20)
        self.success_icon.setAlignment(Qt.AlignCenter)
        success_header.addWidget(self.success_icon)
        self.success_text = QLabel("转换成功!")
        success_header.addWidget(self.success_text)
        result_layout.addLayout(success_header)

        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        self.result_label.setWordWrap(True)
        result_layout.addWidget(self.result_label)

        result_btn_layout = QHBoxLayout()
        result_btn_layout.setSpacing(8)
        self.open_file_btn = SecondaryButton("打开文件", self.colors)
        self.open_file_btn.clicked.connect(self.on_open_file)
        result_btn_layout.addWidget(self.open_file_btn)
        self.open_folder_btn = SecondaryButton("打开文件夹", self.colors)
        self.open_folder_btn.clicked.connect(self.on_open_folder)
        result_btn_layout.addWidget(self.open_folder_btn)
        result_layout.addLayout(result_btn_layout)

        self.result_frame.hide()
        tab_layout.addWidget(self.result_frame)

        tab_layout.addStretch()

    def setup_word_to_md_tab(self):
        """设置 Word转MD 标签页 - 上下两个独立功能"""
        tab_layout = QVBoxLayout(self.word_to_md_tab)
        tab_layout.setContentsMargins(16, 16, 16, 16)
        tab_layout.setSpacing(12)

        # ========== 上半部分：Word → MD ==========
        self.word_to_md_label = QLabel("Word → Markdown")
        self.word_to_md_label.setStyleSheet(f"""
            color: {self.colors['text_primary']};
            font-size: 16px;
            font-weight: 700;
            background: transparent;
        """)
        tab_layout.addWidget(self.word_to_md_label)

        # Word 文件拖拽区域
        self.word_drop_zone = FileDropZone(
            self.colors,
            icon_text="DOCX",
            hint_text="拖拽 Word 文件到这里",
            extensions=['.docx']
        )
        self.word_drop_zone.file_dropped.connect(self.on_word_file_dropped)
        tab_layout.addWidget(self.word_drop_zone)

        # Word 文件选择按钮和转换按钮
        word_btn_layout = QHBoxLayout()
        word_btn_layout.setSpacing(8)
        self.select_word_btn = SecondaryButton("浏览文件...", self.colors)
        self.select_word_btn.clicked.connect(self.on_select_word_file)
        word_btn_layout.addWidget(self.select_word_btn)
        self.word_to_md_btn = GradientButton("转换为 MD", self.colors)
        self.word_to_md_btn.clicked.connect(self.on_convert_word_to_md)
        self.word_to_md_btn.setEnabled(False)
        word_btn_layout.addWidget(self.word_to_md_btn)
        tab_layout.addLayout(word_btn_layout)

        # Word→MD 结果区域
        self.word_to_md_result_frame = QFrame()
        w2m_result_layout = QVBoxLayout(self.word_to_md_result_frame)
        w2m_result_layout.setContentsMargins(12, 10, 12, 10)
        w2m_result_layout.setSpacing(6)

        w2m_header = QHBoxLayout()
        w2m_header.setAlignment(Qt.AlignCenter)
        self.w2m_success_icon = QLabel("OK")
        self.w2m_success_icon.setFixedSize(20, 20)
        self.w2m_success_icon.setAlignment(Qt.AlignCenter)
        w2m_header.addWidget(self.w2m_success_icon)
        self.w2m_success_text = QLabel("转换成功!")
        w2m_header.addWidget(self.w2m_success_text)
        w2m_result_layout.addLayout(w2m_header)

        self.w2m_result_label = QLabel("")
        self.w2m_result_label.setAlignment(Qt.AlignCenter)
        self.w2m_result_label.setWordWrap(True)
        w2m_result_layout.addWidget(self.w2m_result_label)

        w2m_btn_layout = QHBoxLayout()
        w2m_btn_layout.setSpacing(8)
        self.w2m_open_file_btn = SecondaryButton("打开文件", self.colors)
        self.w2m_open_file_btn.clicked.connect(self.on_open_w2m_file)
        w2m_btn_layout.addWidget(self.w2m_open_file_btn)
        self.w2m_open_folder_btn = SecondaryButton("打开文件夹", self.colors)
        self.w2m_open_folder_btn.clicked.connect(self.on_open_w2m_folder)
        w2m_btn_layout.addWidget(self.w2m_open_folder_btn)
        w2m_result_layout.addLayout(w2m_btn_layout)

        self.word_to_md_result_frame.hide()
        tab_layout.addWidget(self.word_to_md_result_frame)

        # 分隔线
        self.tab2_separator = QFrame()
        self.tab2_separator.setFrameShape(QFrame.HLine)
        self.tab2_separator.setStyleSheet(f"background-color: {self.colors['border_default']}; max-height: 2px;")
        tab_layout.addWidget(self.tab2_separator)

        # ========== 下半部分：MD → Word ==========
        self.md_to_word_label = QLabel("Markdown → Word")
        self.md_to_word_label.setStyleSheet(f"""
            color: {self.colors['text_primary']};
            font-size: 16px;
            font-weight: 700;
            background: transparent;
        """)
        tab_layout.addWidget(self.md_to_word_label)

        # MD 文件拖拽区域
        self.md_drop_zone = FileDropZone(
            self.colors,
            icon_text="MD",
            hint_text="拖拽填充好的 MD 文件到这里",
            extensions=['.md', '.markdown']
        )
        self.md_drop_zone.file_dropped.connect(self.on_md_file_dropped)
        tab_layout.addWidget(self.md_drop_zone)

        # MD 文件选择按钮和转换按钮
        md_btn_layout = QHBoxLayout()
        md_btn_layout.setSpacing(8)
        self.select_md_btn = SecondaryButton("浏览文件...", self.colors)
        self.select_md_btn.clicked.connect(self.on_select_md_file)
        md_btn_layout.addWidget(self.select_md_btn)
        self.md_to_word_btn = GradientButton("填充到 Word", self.colors)
        self.md_to_word_btn.clicked.connect(self.on_convert_md_to_word)
        self.md_to_word_btn.setEnabled(False)
        md_btn_layout.addWidget(self.md_to_word_btn)
        tab_layout.addLayout(md_btn_layout)

        # MD→Word 结果区域
        self.md_to_word_result_frame = QFrame()
        m2w_result_layout = QVBoxLayout(self.md_to_word_result_frame)
        m2w_result_layout.setContentsMargins(12, 10, 12, 10)
        m2w_result_layout.setSpacing(6)

        m2w_header = QHBoxLayout()
        m2w_header.setAlignment(Qt.AlignCenter)
        self.m2w_success_icon = QLabel("OK")
        self.m2w_success_icon.setFixedSize(20, 20)
        self.m2w_success_icon.setAlignment(Qt.AlignCenter)
        m2w_header.addWidget(self.m2w_success_icon)
        self.m2w_success_text = QLabel("填充成功!")
        m2w_header.addWidget(self.m2w_success_text)
        m2w_result_layout.addLayout(m2w_header)

        self.m2w_result_label = QLabel("")
        self.m2w_result_label.setAlignment(Qt.AlignCenter)
        self.m2w_result_label.setWordWrap(True)
        m2w_result_layout.addWidget(self.m2w_result_label)

        m2w_btn_layout = QHBoxLayout()
        m2w_btn_layout.setSpacing(8)
        self.m2w_open_file_btn = SecondaryButton("打开文件", self.colors)
        self.m2w_open_file_btn.clicked.connect(self.on_open_m2w_file)
        m2w_btn_layout.addWidget(self.m2w_open_file_btn)
        self.m2w_open_folder_btn = SecondaryButton("打开文件夹", self.colors)
        self.m2w_open_folder_btn.clicked.connect(self.on_open_m2w_folder)
        m2w_btn_layout.addWidget(self.m2w_open_folder_btn)
        m2w_result_layout.addLayout(m2w_btn_layout)

        self.md_to_word_result_frame.hide()
        tab_layout.addWidget(self.md_to_word_result_frame)

        tab_layout.addStretch()

    def on_toc_changed(self):
        """目录选项变更"""
        self.toc_depth_spin.setEnabled(self.toc_checkbox.isChecked())

    def on_highlight_changed(self, style):
        """高亮样式变更"""
        self.update_highlight_preview(style)

    def update_highlight_preview(self, style):
        """更新高亮预览图"""
        # 获取图片路径
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        # 处理 zenburn 拼写错误的文件名
        filename = 'zneburn.png' if style == 'zenburn' else f'{style}.png'
        image_path = os.path.join(base_dir, 'assets', '代码高亮图', filename)

        if os.path.exists(image_path):
            pixmap = QPixmap(image_path)
            # 缩放图片以适应预览区域，保持宽高比
            scaled = pixmap.scaledToWidth(400, Qt.SmoothTransformation)
            if scaled.height() > 200:
                scaled = pixmap.scaledToHeight(200, Qt.SmoothTransformation)
            self.highlight_preview.setPixmap(scaled)
        else:
            self.highlight_preview.setText(f"预览图不存在: {style}")

    def on_file_dropped(self, file_path):
        self.current_file = file_path
        self.convert_btn.setEnabled(True)
        self.status_label.setText(f"已选择: {os.path.basename(file_path)}")
        self.status_label.setStyleSheet(f"""
            color: {self.colors['accent_secondary']};
            font-size: 13px;
            background: transparent;
        """)
        self.result_frame.hide()

    def on_select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Markdown 文件", "",
            "Markdown 文件 (*.md *.markdown);;所有文件 (*.*)"
        )
        if file_path:
            self.current_file = file_path
            self.drop_zone.set_file(file_path)
            self.convert_btn.setEnabled(True)
            self.status_label.setText(f"已选择: {os.path.basename(file_path)}")
            self.status_label.setStyleSheet(f"""
                color: {self.colors['accent_secondary']};
                font-size: 13px;
                background: transparent;
            """)
            self.result_frame.hide()

    def on_convert(self):
        if not self.current_file:
            QMessageBox.warning(self, "警告", "请先选择要转换的文件")
            return

        if not os.path.exists(self.current_file):
            QMessageBox.warning(self, "警告", "文件不存在或已被删除")
            return

        output_file = os.path.splitext(self.current_file)[0] + '.docx'

        options = {
            'generate_toc': self.toc_checkbox.isChecked(),
            'toc_depth': self.toc_depth_spin.value(),
            'highlight_style': self.highlight_combo.currentText(),
            'chinese_font': self.chinese_font_combo.currentText(),
            'code_font': self.code_font_combo.currentText(),
            'font_size': float(self.font_size_combo.currentText()),
            'line_spacing': float(self.line_spacing_combo.currentText())
        }

        self.convert_btn.setEnabled(False)
        self.convert_btn.setText("转换中...")
        self.status_label.setText("正在处理...")
        self.result_frame.hide()

        self.converter_thread = ConverterThread(self.current_file, output_file, options)
        self.converter_thread.progress.connect(self.on_progress)
        self.converter_thread.finished.connect(self.on_finished)
        self.converter_thread.error.connect(self.on_error)
        self.converter_thread.start()

    def on_progress(self, message):
        self.status_label.setText(message)

    def on_finished(self, output_file):
        self.output_file = output_file
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("开始转换")
        self.status_label.setText("")
        self.result_label.setText(f"输出: {os.path.basename(output_file)}")
        self.result_frame.show()

    def on_error(self, error_message):
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("开始转换")
        self.status_label.setText("转换失败")
        self.status_label.setStyleSheet(f"""
            color: {self.colors['error']};
            font-size: 13px;
            background: transparent;
        """)
        QMessageBox.critical(self, "转换失败", error_message)

    def on_open_file(self):
        if hasattr(self, 'output_file') and os.path.exists(self.output_file):
            if sys.platform == 'win32':
                os.startfile(self.output_file)
            elif sys.platform == 'darwin':
                subprocess.run(['open', self.output_file])
            else:
                subprocess.run(['xdg-open', self.output_file])

    def on_open_folder(self):
        if hasattr(self, 'output_file') and os.path.exists(self.output_file):
            if sys.platform == 'win32':
                subprocess.run(['explorer', '/select,', self.output_file])
            elif sys.platform == 'darwin':
                subprocess.run(['open', '-R', self.output_file])
            else:
                subprocess.run(['xdg-open', os.path.dirname(self.output_file)])

    # ========== Word转MD 标签页事件处理 ==========

    # ----- 上半部分：Word → MD -----

    def on_word_file_dropped(self, file_path):
        """Word 文件拖拽（用于转换为 MD）"""
        self.word_file_path = file_path
        self.word_to_md_btn.setEnabled(True)
        self.word_to_md_result_frame.hide()

    def on_select_word_file(self):
        """选择 Word 文件（用于转换为 MD）"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Word 文件", "",
            "Word 文件 (*.docx);;所有文件 (*.*)"
        )
        if file_path:
            self.word_file_path = file_path
            self.word_drop_zone.set_file(file_path)
            self.word_to_md_btn.setEnabled(True)
            self.word_to_md_result_frame.hide()

    def on_convert_word_to_md(self):
        """将 Word 转换为 MD"""
        if not hasattr(self, 'word_file_path') or not self.word_file_path:
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

        success, result = word_to_markdown(self.word_file_path, output_path)

        self.word_to_md_btn.setEnabled(True)
        self.word_to_md_btn.setText("转换为 MD")

        if success:
            self.w2m_output_file = output_path
            self.w2m_result_label.setText(f"输出: {os.path.basename(output_path)}")
            self.word_to_md_result_frame.show()
            self._apply_success_style_w2m()
        else:
            QMessageBox.critical(self, "转换失败", result)

    def _apply_success_style_w2m(self):
        """应用 Word→MD 成功样式"""
        self.word_to_md_result_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['success_bg']};
                border: 1px solid {self.colors['success_border']};
                border-radius: 10px;
            }}
        """)
        self.w2m_success_icon.setStyleSheet(f"""
            background-color: {self.colors['success']};
            color: white;
            font-size: 9px;
            font-weight: bold;
            border-radius: 10px;
        """)
        self.w2m_success_text.setStyleSheet(f"""
            color: {self.colors['success']};
            font-size: 14px;
            font-weight: 600;
            background: transparent;
        """)
        self.w2m_result_label.setStyleSheet(f"""
            color: {self.colors['text_secondary']};
            font-size: 12px;
            background: transparent;
        """)

    def on_open_w2m_file(self):
        """打开转换出的 MD 文件"""
        if hasattr(self, 'w2m_output_file') and os.path.exists(self.w2m_output_file):
            if sys.platform == 'win32':
                os.startfile(self.w2m_output_file)
            elif sys.platform == 'darwin':
                subprocess.run(['open', self.w2m_output_file])
            else:
                subprocess.run(['xdg-open', self.w2m_output_file])

    def on_open_w2m_folder(self):
        """打开 MD 文件所在文件夹"""
        if hasattr(self, 'w2m_output_file') and os.path.exists(self.w2m_output_file):
            if sys.platform == 'win32':
                subprocess.run(['explorer', '/select,', self.w2m_output_file])
            elif sys.platform == 'darwin':
                subprocess.run(['open', '-R', self.w2m_output_file])
            else:
                subprocess.run(['xdg-open', os.path.dirname(self.w2m_output_file)])

    # ----- 下半部分：MD → Word -----

    def on_md_file_dropped(self, file_path):
        """MD 文件拖拽（用于填充回 Word）"""
        self.md_file_path = file_path
        self.md_to_word_btn.setEnabled(True)
        self.md_to_word_result_frame.hide()

    def on_select_md_file(self):
        """选择填充好的 MD 文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Markdown 文件", "",
            "Markdown 文件 (*.md);;所有文件 (*.*)"
        )
        if file_path:
            self.md_file_path = file_path
            self.md_drop_zone.set_file(file_path)
            self.md_to_word_btn.setEnabled(True)
            self.md_to_word_result_frame.hide()

    def on_convert_md_to_word(self):
        """将 MD 内容填充回 Word"""
        if not hasattr(self, 'md_file_path') or not self.md_file_path:
            QMessageBox.warning(self, "警告", "请先选择 MD 文件")
            return

        if not os.path.exists(self.md_file_path):
            QMessageBox.warning(self, "警告", "MD 文件不存在")
            return

        # 从 MD 文件中获取源 Word 路径
        source_path = get_template_source(self.md_file_path)
        if not source_path:
            QMessageBox.warning(self, "警告", "MD 文件中没有找到源 Word 模板路径\n请确保 MD 文件包含 <!-- source: ... --> 注释")
            return

        if not os.path.exists(source_path):
            QMessageBox.warning(self, "警告", f"源 Word 模板不存在:\n{source_path}")
            return

        # 生成输出文件路径
        base_name = os.path.splitext(source_path)[0]
        output_path = base_name + '_filled.docx'

        self.md_to_word_btn.setEnabled(False)
        self.md_to_word_btn.setText("填充中...")

        success, result = markdown_to_word(self.md_file_path, source_path, output_path)

        self.md_to_word_btn.setEnabled(True)
        self.md_to_word_btn.setText("填充到 Word")

        if success:
            self.m2w_output_file = output_path
            self.m2w_result_label.setText(f"输出: {os.path.basename(output_path)}\n{result}")
            self.md_to_word_result_frame.show()
            self._apply_success_style_m2w()
        else:
            QMessageBox.critical(self, "填充失败", result)

    def _apply_success_style_m2w(self):
        """应用 MD→Word 成功样式"""
        self.md_to_word_result_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {self.colors['success_bg']};
                border: 1px solid {self.colors['success_border']};
                border-radius: 10px;
            }}
        """)
        self.m2w_success_icon.setStyleSheet(f"""
            background-color: {self.colors['success']};
            color: white;
            font-size: 9px;
            font-weight: bold;
            border-radius: 10px;
        """)
        self.m2w_success_text.setStyleSheet(f"""
            color: {self.colors['success']};
            font-size: 14px;
            font-weight: 600;
            background: transparent;
        """)
        self.m2w_result_label.setStyleSheet(f"""
            color: {self.colors['text_secondary']};
            font-size: 12px;
            background: transparent;
        """)

    def on_open_m2w_file(self):
        """打开填充后的 Word 文件"""
        if hasattr(self, 'm2w_output_file') and os.path.exists(self.m2w_output_file):
            if sys.platform == 'win32':
                os.startfile(self.m2w_output_file)
            elif sys.platform == 'darwin':
                subprocess.run(['open', self.m2w_output_file])
            else:
                subprocess.run(['xdg-open', self.m2w_output_file])

    def on_open_m2w_folder(self):
        """打开填充后的 Word 文件所在文件夹"""
        if hasattr(self, 'm2w_output_file') and os.path.exists(self.m2w_output_file):
            if sys.platform == 'win32':
                subprocess.run(['explorer', '/select,', self.m2w_output_file])
            elif sys.platform == 'darwin':
                subprocess.run(['open', '-R', self.m2w_output_file])
            else:
                subprocess.run(['xdg-open', os.path.dirname(self.m2w_output_file)])
