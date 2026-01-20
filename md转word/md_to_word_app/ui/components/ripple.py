# -*- coding: utf-8 -*-
"""
Material Design 涟漪动画组件
点击时产生水波纹扩散效果
"""

from PyQt5.QtWidgets import QWidget
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve, pyqtProperty, QPoint, QRect
from PyQt5.QtGui import QPainter, QColor, QPainterPath


class RippleEffect:
    """单个涟漪效果"""

    def __init__(self, center: QPoint, color: QColor, max_radius: float):
        self.center = center
        self.color = QColor(color)
        self.max_radius = max_radius
        self.radius = 0.0
        self.opacity = 0.3
        self.finished = False

    def update(self, progress: float):
        """更新涟漪状态 (progress: 0.0 ~ 1.0)"""
        self.radius = self.max_radius * progress
        # 后半段开始淡出
        if progress > 0.4:
            self.opacity = 0.3 * (1 - (progress - 0.4) / 0.6)
        else:
            self.opacity = 0.3

        if progress >= 1.0:
            self.finished = True

    def paint(self, painter: QPainter):
        """绘制涟漪"""
        if self.opacity <= 0:
            return

        painter.save()
        color = QColor(self.color)
        color.setAlphaF(self.opacity)
        painter.setBrush(color)
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(self.center, int(self.radius), int(self.radius))
        painter.restore()


class RippleWidget(QWidget):
    """
    带涟漪动画效果的基类 Widget
    继承此类并重写 paintEvent 时需调用 super().paintEvent(event)
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._ripples = []
        self._ripple_color = QColor(255, 255, 255)
        self._animation_timer = QTimer(self)
        self._animation_timer.timeout.connect(self._update_ripples)
        self._animation_progress = 0.0
        self._current_ripple = None

        # 启用鼠标追踪
        self.setMouseTracking(True)

    def set_ripple_color(self, color: QColor):
        """设置涟漪颜色"""
        self._ripple_color = QColor(color)

    def _start_ripple(self, pos: QPoint):
        """在指定位置开始涟漪动画"""
        # 计算最大半径 (到最远角落的距离)
        rect = self.rect()
        corners = [
            rect.topLeft(), rect.topRight(),
            rect.bottomLeft(), rect.bottomRight()
        ]
        max_dist = 0
        for corner in corners:
            dist = ((corner.x() - pos.x()) ** 2 + (corner.y() - pos.y()) ** 2) ** 0.5
            max_dist = max(max_dist, dist)

        ripple = RippleEffect(pos, self._ripple_color, max_dist)
        self._ripples.append(ripple)
        self._current_ripple = ripple
        self._animation_progress = 0.0

        if not self._animation_timer.isActive():
            self._animation_timer.start(16)  # ~60fps

    def _update_ripples(self):
        """更新所有涟漪动画"""
        if not self._ripples:
            self._animation_timer.stop()
            return

        # 更新进度
        self._animation_progress += 0.04  # 约 400ms 完成

        # 更新所有涟漪
        for ripple in self._ripples:
            if not ripple.finished:
                ripple.update(self._animation_progress)

        # 移除已完成的涟漪
        self._ripples = [r for r in self._ripples if not r.finished]

        if not self._ripples:
            self._animation_timer.stop()

        self.update()

    def mousePressEvent(self, event):
        """鼠标按下时启动涟漪"""
        if event.button() == Qt.LeftButton:
            self._start_ripple(event.pos())
        super().mousePressEvent(event)

    def paintEvent(self, event):
        """绘制涟漪效果"""
        super().paintEvent(event)

        if not self._ripples:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 裁剪到控件边界 (圆角)
        path = QPainterPath()
        path.addRoundedRect(0, 0, self.width(), self.height(), 8, 8)
        painter.setClipPath(path)

        # 绘制所有涟漪
        for ripple in self._ripples:
            ripple.paint(painter)


class RippleButton(RippleWidget):
    """
    带涟漪效果的按钮基类
    提供悬停和按下状态的视觉反馈
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._hovered = False
        self._pressed = False
        self._bg_color = QColor(30, 30, 30)
        self._hover_color = QColor(45, 45, 45)
        self._press_color = QColor(60, 60, 60)

        self.setCursor(Qt.PointingHandCursor)

    def set_colors(self, bg: str, hover: str, press: str, ripple: str):
        """设置颜色方案"""
        self._bg_color = QColor(bg)
        self._hover_color = QColor(hover)
        self._press_color = QColor(press)
        self.set_ripple_color(QColor(ripple))

    def enterEvent(self, event):
        self._hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hovered = False
        self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        self._pressed = True
        self.update()
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        self._pressed = False
        self.update()
        super().mouseReleaseEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制背景
        if self._pressed:
            painter.setBrush(self._press_color)
        elif self._hovered:
            painter.setBrush(self._hover_color)
        else:
            painter.setBrush(self._bg_color)

        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(self.rect(), 8, 8)

        # 绘制涟漪
        super().paintEvent(event)
