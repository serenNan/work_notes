# -*- coding: utf-8 -*-
"""
页面模块 - 应用程序的各个功能页面
"""

from .md_to_word_page import MdToWordPage
from .word_to_md_page import WordToMdPage
from .history_view import HistoryView

__all__ = [
    'MdToWordPage',
    'WordToMdPage',
    'HistoryView',
]
