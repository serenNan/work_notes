# -*- coding: utf-8 -*-
"""
主窗口 UI - 支持浅色/深色主题切换 + 实时预览
"""

import os
import subprocess
import sys
import tempfile
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QCheckBox, QComboBox, QSpinBox,
    QGroupBox, QFileDialog, QMessageBox, QFrame, QGraphicsDropShadowEffect,
    QSplitter, QTextBrowser, QScrollArea, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QFont, QColor

# 导入转换服务
from core.converter import ConverterService

# 获取 assets 目录
ASSETS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
DEMO_MD_PATH = os.path.join(ASSETS_DIR, "demo.md")

# Pandoc 路径
PANDOC_PATH = r'S:\Tools\Miniconda\envs\pandoc\Library\bin\pandoc.exe'


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
    'warning': '#f59e0b',
    'shadow_color': 'rgba(0, 0, 0, 80)',
    'drag_bg': 'rgba(59, 130, 246, 0.1)',
    'preview_bg': '#1a1a1a',
    'code_bg': '#2d2d2d',
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
    'warning': '#d97706',
    'shadow_color': 'rgba(0, 0, 0, 40)',
    'drag_bg': 'rgba(37, 99, 235, 0.08)',
    'preview_bg': '#ffffff',
    'code_bg': '#f5f5f5',
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

/* 分割器样式 */
QSplitter::handle {{
    background-color: {colors['border_subtle']};
}}

QSplitter::handle:horizontal {{
    width: 2px;
}}

/* 文本浏览器样式 */
QTextBrowser {{
    background-color: {colors['preview_bg']};
    border: 1px solid {colors['border_subtle']};
    border-radius: 8px;
    padding: 16px;
    color: {colors['text_primary']};
}}

/* 滚动区域样式 */
QScrollArea {{
    border: none;
    background-color: transparent;
}}

QScrollBar:vertical {{
    background-color: {colors['bg_secondary']};
    width: 10px;
    border-radius: 5px;
}}

QScrollBar::handle:vertical {{
    background-color: {colors['border_default']};
    border-radius: 5px;
    min-height: 30px;
}}

QScrollBar::handle:vertical:hover {{
    background-color: {colors['border_hover']};
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
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


class PreviewThread(QThread):
    """预览生成线程"""
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, md_path, options, colors):
        super().__init__()
        self.md_path = md_path
        self.options = options
        self.colors = colors

    def run(self):
        try:
            # 使用 pandoc 转换为 HTML
            cmd = [
                PANDOC_PATH,
                self.md_path,
                '-f', 'markdown+tex_math_dollars+raw_tex',
                '-t', 'html5',
                '--standalone',
                f'--highlight-style={self.options.get("highlight_style", "tango")}',
                '--mathjax',
            ]

            # 添加目录
            if self.options.get('generate_toc', True):
                cmd.append('--toc')
                cmd.append(f'--toc-depth={self.options.get("toc_depth", 3)}')

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                cwd=os.path.dirname(self.md_path)
            )

            if result.returncode == 0:
                # 添加自定义样式
                html = self.inject_custom_styles(result.stdout)
                self.finished.emit(html)
            else:
                self.error.emit(result.stderr)

        except Exception as e:
            self.error.emit(str(e))

    def inject_custom_styles(self, html):
        """注入自定义样式以匹配主题"""
        colors = self.colors
        custom_css = f"""
        <style>
        body {{
            font-family: "Segoe UI", "Microsoft YaHei UI", sans-serif;
            background-color: {colors['preview_bg']};
            color: {colors['text_primary']};
            padding: 20px;
            line-height: 1.6;
            max-width: 100%;
        }}
        h1, h2, h3, h4, h5, h6 {{
            color: {colors['text_primary']};
            margin-top: 1.5em;
            margin-bottom: 0.5em;
        }}
        h1 {{ font-size: 1.8em; border-bottom: 2px solid {colors['accent_primary']}; padding-bottom: 0.3em; }}
        h2 {{ font-size: 1.5em; border-bottom: 1px solid {colors['border_default']}; padding-bottom: 0.2em; }}
        h3 {{ font-size: 1.2em; }}
        code {{
            background-color: {colors['code_bg']};
            padding: 2px 6px;
            border-radius: 4px;
            font-family: "Consolas", "Monaco", monospace;
            font-size: 0.9em;
        }}
        pre {{
            background-color: {colors['code_bg']};
            padding: 16px;
            border-radius: 8px;
            overflow-x: auto;
            border: 1px solid {colors['border_subtle']};
        }}
        pre code {{
            background-color: transparent;
            padding: 0;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
        }}
        th, td {{
            border: 1px solid {colors['border_default']};
            padding: 10px 12px;
            text-align: left;
        }}
        th {{
            background-color: {colors['bg_elevated']};
            font-weight: 600;
        }}
        blockquote {{
            border-left: 4px solid {colors['accent_primary']};
            margin: 1em 0;
            padding: 0.5em 1em;
            background-color: {colors['bg_secondary']};
            color: {colors['text_secondary']};
        }}
        a {{
            color: {colors['accent_primary']};
            text-decoration: none;
        }}
        a:hover {{
            text-decoration: underline;
        }}
        hr {{
            border: none;
            border-top: 1px solid {colors['border_default']};
            margin: 2em 0;
        }}
        #TOC {{
            background-color: {colors['bg_secondary']};
            padding: 16px;
            border-radius: 8px;
            margin-bottom: 2em;
        }}
        #TOC ul {{
            list-style-type: none;
            padding-left: 1em;
        }}
        #TOC > ul {{
            padding-left: 0;
        }}
        #TOC a {{
            color: {colors['text_secondary']};
        }}
        </style>
        """
        # 在 </head> 前插入样式
        if '</head>' in html:
            html = html.replace('</head>', custom_css + '</head>')
        else:
            html = custom_css + html
        return html


class DropZone(QFrame):
    """拖拽区域组件"""
    file_dropped = pyqtSignal(str)

    def __init__(self, colors, parent=None):
        super().__init__(parent)
        self.colors = colors
        self.has_file = False
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

        icon_label = QLabel("MD")
        icon_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        icon_label.setStyleSheet("color: white; background: transparent;")
        icon_label.setAlignment(Qt.AlignCenter)
        icon_layout.addWidget(icon_label)

        icon_h_layout = QHBoxLayout()
        icon_h_layout.addStretch()
        icon_h_layout.addWidget(self.icon_container)
        icon_h_layout.addStretch()
        layout.addLayout(icon_h_layout)

        self.hint_label = QLabel("拖拽 Markdown 文件到这里")
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
            DropZone {{
                background-color: {self.colors['bg_primary']};
                border: 2px dashed {self.colors['border_default']};
                border-radius: 12px;
            }}
            DropZone:hover {{
                border-color: {self.colors['accent_primary']};
                background-color: {self.colors['bg_secondary']};
            }}
        """)

    def apply_drag_style(self):
        self.setStyleSheet(f"""
            DropZone {{
                background-color: {self.colors['drag_bg']};
                border: 2px dashed {self.colors['accent_primary']};
                border-radius: 12px;
            }}
        """)

    def apply_file_selected_style(self):
        self.setStyleSheet(f"""
            DropZone {{
                background-color: {self.colors['bg_primary']};
                border: 2px solid {self.colors['accent_primary']};
                border-radius: 12px;
            }}
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1:
                file_path = urls[0].toLocalFile()
                if file_path.lower().endswith(('.md', '.markdown')):
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
            if file_path.lower().endswith(('.md', '.markdown')):
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
            self.hint_label.setText("拖拽 Markdown 文件到这里")
            self.apply_normal_style()


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
    """主窗口 - 支持主题切换和实时预览"""

    def __init__(self):
        super().__init__()
        self.current_file = None
        self.converter_thread = None
        self.preview_thread = None
        self.preview_timer = QTimer()
        self.preview_timer.setSingleShot(True)
        self.preview_timer.timeout.connect(self.generate_preview)

        # 加载用户主题偏好
        self.settings = QSettings("MD2Word", "MarkdownToWord")
        self.is_dark_theme = self.settings.value("dark_theme", True, type=bool)
        self.colors = DARK_COLORS if self.is_dark_theme else LIGHT_COLORS

        self.setup_ui()
        self.apply_theme()

        # 初始加载预览
        QTimer.singleShot(500, self.generate_preview)

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

        # 组件更新
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
        for label in [self.toc_depth_label, self.highlight_label]:
            label.setStyleSheet(f"""
                color: {self.colors['text_secondary']};
                font-size: 14px;
                background: transparent;
            """)

        self.separator.setStyleSheet(f"background-color: {self.colors['border_subtle']}; max-height: 1px;")

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

        # 预览区域
        self.preview_title.setStyleSheet(f"""
            font-size: 14px;
            font-weight: 600;
            color: {self.colors['text_secondary']};
            background: transparent;
        """)
        self.preview_browser.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {self.colors['preview_bg']};
                border: 1px solid {self.colors['border_subtle']};
                border-radius: 8px;
            }}
        """)

        # 重新生成预览以更新主题
        self.schedule_preview_update()

    def toggle_theme(self):
        self.is_dark_theme = not self.is_dark_theme
        self.settings.setValue("dark_theme", self.is_dark_theme)
        self.apply_theme()

    def setup_ui(self):
        self.setWindowTitle("Markdown to Word")
        self.setMinimumSize(1100, 750)
        self.resize(1200, 850)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        main_layout = QHBoxLayout(self.central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(1)

        # 左侧面板 - 设置
        left_panel = QWidget()
        left_panel.setMinimumWidth(380)
        left_panel.setMaximumWidth(450)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(24, 20, 12, 20)
        left_layout.setSpacing(16)

        # 顶部栏
        top_bar = QHBoxLayout()
        self.title_label = QLabel("MD to Word")
        top_bar.addWidget(self.title_label)
        top_bar.addStretch()
        self.theme_btn = ThemeToggleButton(self.is_dark_theme, self.colors)
        self.theme_btn.clicked.connect(self.toggle_theme)
        top_bar.addWidget(self.theme_btn)
        left_layout.addLayout(top_bar)

        # 拖拽区域
        self.drop_zone = DropZone(self.colors)
        self.drop_zone.file_dropped.connect(self.on_file_dropped)
        left_layout.addWidget(self.drop_zone)

        # 选择文件按钮
        self.select_btn = SecondaryButton("浏览文件...", self.colors)
        self.select_btn.clicked.connect(self.on_select_file)
        left_layout.addWidget(self.select_btn)

        # 配置选项组
        self.options_group = QGroupBox("转换选项")
        options_layout = QVBoxLayout(self.options_group)
        options_layout.setSpacing(12)
        options_layout.setContentsMargins(14, 20, 14, 14)

        # 生成目录
        self.toc_checkbox = QCheckBox("生成目录")
        self.toc_checkbox.setChecked(True)
        self.toc_checkbox.stateChanged.connect(self.on_setting_changed)
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
        self.toc_depth_spin.valueChanged.connect(self.on_setting_changed)
        toc_depth_layout.addWidget(self.toc_depth_spin)
        toc_depth_layout.addStretch()
        options_layout.addLayout(toc_depth_layout)

        # 分割线
        self.separator = QFrame()
        self.separator.setFrameShape(QFrame.HLine)
        options_layout.addWidget(self.separator)

        # 代码高亮
        highlight_layout = QHBoxLayout()
        highlight_layout.setSpacing(10)
        self.highlight_label = QLabel("代码高亮")
        highlight_layout.addWidget(self.highlight_label)
        self.highlight_combo = QComboBox()
        self.highlight_combo.addItems(['tango', 'pygments', 'espresso', 'monochrome', 'kate', 'zenburn'])
        self.highlight_combo.setFixedWidth(120)
        self.highlight_combo.currentTextChanged.connect(self.on_setting_changed)
        highlight_layout.addWidget(self.highlight_combo)
        highlight_layout.addStretch()
        options_layout.addLayout(highlight_layout)

        left_layout.addWidget(self.options_group)

        # 转换按钮
        self.convert_btn = GradientButton("开始转换", self.colors)
        self.convert_btn.clicked.connect(self.on_convert)
        self.convert_btn.setEnabled(False)
        left_layout.addWidget(self.convert_btn)

        # 状态标签
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(self.status_label)

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
        left_layout.addWidget(self.result_frame)

        left_layout.addStretch()

        # 右侧面板 - 预览
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(12, 20, 24, 20)
        right_layout.setSpacing(12)

        # 预览标题
        preview_header = QHBoxLayout()
        self.preview_title = QLabel("实时预览")
        preview_header.addWidget(self.preview_title)
        preview_header.addStretch()
        right_layout.addLayout(preview_header)

        # 预览浏览器
        self.preview_browser = QTextBrowser()
        self.preview_browser.setOpenExternalLinks(True)
        self.preview_browser.setMinimumWidth(500)
        right_layout.addWidget(self.preview_browser)

        # 添加到分割器
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

        main_layout.addWidget(splitter)

    def on_setting_changed(self):
        """设置变更时触发预览更新"""
        self.toc_depth_spin.setEnabled(self.toc_checkbox.isChecked())
        self.schedule_preview_update()

    def schedule_preview_update(self):
        """延迟更新预览，避免频繁刷新"""
        self.preview_timer.start(300)

    def generate_preview(self):
        """生成预览"""
        if not os.path.exists(DEMO_MD_PATH):
            self.preview_browser.setHtml("<p>Demo file not found</p>")
            return

        options = {
            'generate_toc': self.toc_checkbox.isChecked(),
            'toc_depth': self.toc_depth_spin.value(),
            'highlight_style': self.highlight_combo.currentText()
        }

        self.preview_thread = PreviewThread(DEMO_MD_PATH, options, self.colors)
        self.preview_thread.finished.connect(self.on_preview_finished)
        self.preview_thread.error.connect(self.on_preview_error)
        self.preview_thread.start()

    def on_preview_finished(self, html):
        """预览生成完成"""
        self.preview_browser.setHtml(html)

    def on_preview_error(self, error):
        """预览生成错误"""
        self.preview_browser.setHtml(f"<p style='color: red;'>Preview error: {error}</p>")

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
            'highlight_style': self.highlight_combo.currentText()
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
