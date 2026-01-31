"""计算器模块 - 用于测试 LSP 功能"""

from typing import List, Optional


class Calculator:
    """简单计算器类"""

    def __init__(self, name: str = "Calculator"):
        self.name = name
        self.history: List[str] = []

    def add(self, a: float, b: float) -> float:
        """加法运算"""
        result = a + b
        self._record(f"{a} + {b} = {result}")
        return result

    def subtract(self, a: float, b: float) -> float:
        """减法运算"""
        result = a - b
        self._record(f"{a} - {b} = {result}")
        return result

    def multiply(self, a: float, b: float) -> float:
        """乘法运算"""
        result = a * b
        self._record(f"{a} * {b} = {result}")
        return result

    def divide(self, a: float, b: float) -> Optional[float]:
        """除法运算"""
        if b == 0:
            self._record(f"{a} / {b} = Error (division by zero)")
            return None
        result = a / b
        self._record(f"{a} / {b} = {result}")
        return result

    def _record(self, operation: str) -> None:
        """记录操作历史"""
        self.history.append(operation)

    def get_history(self) -> List[str]:
        """获取历史记录"""
        return self.history.copy()

    def clear_history(self) -> None:
        """清空历史记录"""
        self.history.clear()


def create_calculator(name: str) -> Calculator:
    """工厂函数: 创建计算器实例"""
    return Calculator(name)


PI = 3.14159
E = 2.71828
