# -*- coding: utf-8 -*-
"""
历史记录管理器
使用 QSettings 持久化存储转换历史
"""

import os
from datetime import datetime
from typing import List, Dict, Optional
from PyQt5.QtCore import QSettings, QObject, pyqtSignal


class HistoryItem:
    """历史记录项"""

    def __init__(self, source_file: str, output_file: str,
                 conversion_type: str, timestamp: str = None,
                 status: str = 'success', message: str = ''):
        self.source_file = source_file
        self.output_file = output_file
        self.conversion_type = conversion_type  # 'md_to_word' or 'word_to_md'
        self.timestamp = timestamp or datetime.now().isoformat()
        self.status = status  # 'success', 'error'
        self.message = message

    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'source_file': self.source_file,
            'output_file': self.output_file,
            'conversion_type': self.conversion_type,
            'timestamp': self.timestamp,
            'status': self.status,
            'message': self.message,
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'HistoryItem':
        """从字典创建"""
        return cls(
            source_file=data.get('source_file', ''),
            output_file=data.get('output_file', ''),
            conversion_type=data.get('conversion_type', ''),
            timestamp=data.get('timestamp', ''),
            status=data.get('status', 'success'),
            message=data.get('message', ''),
        )

    def get_source_filename(self) -> str:
        """获取源文件名"""
        return os.path.basename(self.source_file)

    def get_output_filename(self) -> str:
        """获取输出文件名"""
        return os.path.basename(self.output_file)

    def get_formatted_time(self) -> str:
        """获取格式化的时间"""
        try:
            dt = datetime.fromisoformat(self.timestamp)
            return dt.strftime('%Y-%m-%d %H:%M')
        except:
            return self.timestamp


class HistoryManager(QObject):
    """
    历史记录管理器
    使用 QSettings 持久化存储
    """
    history_changed = pyqtSignal()

    MAX_HISTORY_ITEMS = 50

    def __init__(self, parent=None):
        super().__init__(parent)
        self._settings = QSettings('MDtoWord', 'ConversionHistory')
        self._history: List[HistoryItem] = []
        self._load_history()

    def _load_history(self):
        """从存储加载历史记录"""
        count = self._settings.beginReadArray('history')
        self._history = []

        for i in range(count):
            self._settings.setArrayIndex(i)
            data = {
                'source_file': self._settings.value('source_file', ''),
                'output_file': self._settings.value('output_file', ''),
                'conversion_type': self._settings.value('conversion_type', ''),
                'timestamp': self._settings.value('timestamp', ''),
                'status': self._settings.value('status', 'success'),
                'message': self._settings.value('message', ''),
            }
            self._history.append(HistoryItem.from_dict(data))

        self._settings.endArray()

    def _save_history(self):
        """保存历史记录到存储"""
        self._settings.beginWriteArray('history')

        for i, item in enumerate(self._history):
            self._settings.setArrayIndex(i)
            data = item.to_dict()
            for key, value in data.items():
                self._settings.setValue(key, value)

        self._settings.endArray()
        self._settings.sync()

    def add_record(self, source_file: str, output_file: str,
                   conversion_type: str, status: str = 'success',
                   message: str = ''):
        """
        添加历史记录

        Args:
            source_file: 源文件路径
            output_file: 输出文件路径
            conversion_type: 转换类型 ('md_to_word' or 'word_to_md')
            status: 状态 ('success' or 'error')
            message: 消息
        """
        item = HistoryItem(
            source_file=source_file,
            output_file=output_file,
            conversion_type=conversion_type,
            status=status,
            message=message,
        )

        # 插入到列表开头
        self._history.insert(0, item)

        # 限制最大数量
        if len(self._history) > self.MAX_HISTORY_ITEMS:
            self._history = self._history[:self.MAX_HISTORY_ITEMS]

        self._save_history()
        self.history_changed.emit()

    def get_history(self, conversion_type: str = None,
                    limit: int = None) -> List[HistoryItem]:
        """
        获取历史记录

        Args:
            conversion_type: 过滤转换类型 (可选)
            limit: 限制返回数量 (可选)

        Returns:
            历史记录列表
        """
        history = self._history

        if conversion_type:
            history = [h for h in history if h.conversion_type == conversion_type]

        if limit:
            history = history[:limit]

        return history

    def get_recent(self, count: int = 10) -> List[HistoryItem]:
        """获取最近的历史记录"""
        return self._history[:count]

    def clear_history(self):
        """清空历史记录"""
        self._history = []
        self._save_history()
        self.history_changed.emit()

    def remove_item(self, index: int):
        """删除指定索引的记录"""
        if 0 <= index < len(self._history):
            del self._history[index]
            self._save_history()
            self.history_changed.emit()

    def get_count(self) -> int:
        """获取历史记录数量"""
        return len(self._history)


# 全局单例
_history_manager: Optional[HistoryManager] = None


def get_history_manager() -> HistoryManager:
    """获取历史记录管理器单例"""
    global _history_manager
    if _history_manager is None:
        _history_manager = HistoryManager()
    return _history_manager
