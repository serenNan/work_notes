# -*- coding: utf-8 -*-
"""
UI 模块 - 提供应用程序的图形用户界面

模块结构:
- styles/: 主题管理, 颜色配置, 样式表生成
- components/: 可复用的 UI 组件
- pages/: 功能页面
- main_window.py: 主窗口
"""

from .main_window import MainWindow

__all__ = ['MainWindow']
