# -*- coding: utf-8 -*-
"""
设置管理器 - 使用 QSettings 持久化应用设置
"""

from PyQt5.QtCore import QSettings, QObject, pyqtSignal


class SettingsManager(QObject):
    """
    应用设置管理器
    使用 QSettings 持久化存储
    """
    # 信号
    preview_visible_changed = pyqtSignal(bool)
    settings_changed = pyqtSignal()

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        super().__init__()
        self._initialized = True
        self._settings = QSettings('MD2Word', 'MDtoWordApp')

    # ========== 预览面板设置 ==========

    @property
    def preview_visible(self) -> bool:
        """预览面板是否可见"""
        return self._settings.value('ui/preview_visible', True, type=bool)

    @preview_visible.setter
    def preview_visible(self, value: bool):
        """设置预览面板可见性"""
        if self.preview_visible != value:
            self._settings.setValue('ui/preview_visible', value)
            self.preview_visible_changed.emit(value)

    def toggle_preview(self) -> bool:
        """切换预览面板可见性, 返回新状态"""
        new_value = not self.preview_visible
        self.preview_visible = new_value
        return new_value

    # ========== 主题设置 ==========

    @property
    def is_dark_theme(self) -> bool:
        """是否深色主题"""
        return self._settings.value('ui/dark_theme', True, type=bool)

    @is_dark_theme.setter
    def is_dark_theme(self, value: bool):
        """设置主题"""
        self._settings.setValue('ui/dark_theme', value)
        self.settings_changed.emit()

    # ========== 窗口设置 ==========

    def save_window_geometry(self, geometry):
        """保存窗口几何信息"""
        self._settings.setValue('window/geometry', geometry)

    def load_window_geometry(self):
        """加载窗口几何信息"""
        return self._settings.value('window/geometry')

    def save_window_state(self, state):
        """保存窗口状态"""
        self._settings.setValue('window/state', state)

    def load_window_state(self):
        """加载窗口状态"""
        return self._settings.value('window/state')

    # ========== 转换默认值设置 ==========

    @property
    def default_toc_enabled(self) -> bool:
        """默认是否生成目录"""
        return self._settings.value('convert/toc_enabled', True, type=bool)

    @default_toc_enabled.setter
    def default_toc_enabled(self, value: bool):
        self._settings.setValue('convert/toc_enabled', value)
        self.settings_changed.emit()

    @property
    def default_toc_depth(self) -> int:
        """默认目录深度"""
        return self._settings.value('convert/toc_depth', 3, type=int)

    @default_toc_depth.setter
    def default_toc_depth(self, value: int):
        self._settings.setValue('convert/toc_depth', value)
        self.settings_changed.emit()

    @property
    def default_chinese_font(self) -> str:
        """默认中文字体"""
        return self._settings.value('convert/chinese_font', '宋体', type=str)

    @default_chinese_font.setter
    def default_chinese_font(self, value: str):
        self._settings.setValue('convert/chinese_font', value)
        self.settings_changed.emit()

    @property
    def default_code_font(self) -> str:
        """默认代码字体"""
        return self._settings.value('convert/code_font', 'Times New Roman', type=str)

    @default_code_font.setter
    def default_code_font(self, value: str):
        self._settings.setValue('convert/code_font', value)
        self.settings_changed.emit()

    @property
    def default_font_size(self) -> str:
        """默认字体大小"""
        return self._settings.value('convert/font_size', '12', type=str)

    @default_font_size.setter
    def default_font_size(self, value: str):
        self._settings.setValue('convert/font_size', value)
        self.settings_changed.emit()

    @property
    def default_line_spacing(self) -> str:
        """默认行间距"""
        return self._settings.value('convert/line_spacing', '1.5', type=str)

    @default_line_spacing.setter
    def default_line_spacing(self, value: str):
        self._settings.setValue('convert/line_spacing', value)
        self.settings_changed.emit()

    @property
    def default_highlight_style(self) -> str:
        """默认代码高亮风格"""
        return self._settings.value('convert/highlight_style', 'tango', type=str)

    @default_highlight_style.setter
    def default_highlight_style(self, value: str):
        self._settings.setValue('convert/highlight_style', value)
        self.settings_changed.emit()

    # ========== 路径设置 ==========

    @property
    def pandoc_path(self) -> str:
        """Pandoc 路径"""
        # 检查是否有用户配置的路径
        saved_path = self._settings.value('paths/pandoc', '', type=str)
        if saved_path:
            return saved_path
        # 没有配置则使用自动检测
        from .converter import find_pandoc, get_default_pandoc_path
        return get_default_pandoc_path()

    @pandoc_path.setter
    def pandoc_path(self, value: str):
        self._settings.setValue('paths/pandoc', value)
        self.settings_changed.emit()

    @property
    def default_output_dir(self) -> str:
        """默认输出目录 (空字符串表示与源文件同目录)"""
        return self._settings.value('paths/output_dir', '', type=str)

    @default_output_dir.setter
    def default_output_dir(self, value: str):
        self._settings.setValue('paths/output_dir', value)
        self.settings_changed.emit()

    # ========== 历史记录设置 ==========

    @property
    def max_history_count(self) -> int:
        """最大历史记录数量"""
        return self._settings.value('history/max_count', 50, type=int)

    @max_history_count.setter
    def max_history_count(self, value: int):
        self._settings.setValue('history/max_count', value)
        self.settings_changed.emit()

    # ========== 获取所有转换默认值 ==========

    def get_default_options(self) -> dict:
        """获取所有转换默认选项"""
        return {
            'toc': self.default_toc_enabled,
            'toc_depth': self.default_toc_depth,
            'chinese_font': self.default_chinese_font,
            'code_font': self.default_code_font,
            'font_size': self.default_font_size,
            'line_spacing': self.default_line_spacing,
            'highlight_style': self.default_highlight_style,
        }


# 全局实例
_settings_manager = None


def get_settings_manager() -> SettingsManager:
    """获取设置管理器单例"""
    global _settings_manager
    if _settings_manager is None:
        _settings_manager = SettingsManager()
    return _settings_manager
