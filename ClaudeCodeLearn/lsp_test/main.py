"""主程序 - 测试 LSP 跨文件功能"""

from calculator import Calculator, create_calculator, PI, E


def main():
    # 测试 LSP: 悬停查看类型提示
    calc = Calculator("MyCalc")

    # 测试 LSP: 方法自动补全和跳转到定义
    result1 = calc.add(10, 5)
    result2 = calc.subtract(10, 5)
    result3 = calc.multiply(3, 4)
    result4 = calc.divide(20, 4)

    print(f"Calculator: {calc.name}")
    print(f"10 + 5 = {result1}")
    print(f"10 - 5 = {result2}")
    print(f"3 * 4 = {result3}")
    print(f"20 / 4 = {result4}")

    # 测试 LSP: 查看属性引用
    history = calc.get_history()
    print(f"\nHistory ({len(history)} operations):")
    for op in history:
        print(f"  - {op}")

    # 测试 LSP: 工厂函数和常量引用
    calc2 = create_calculator("SciCalc")
    print(f"\nCreated: {calc2.name}")
    circle_area = PI * 5 * 5
    print(f"\nCircle area (r=5): {circle_area}")
    print(f"Euler's number: {E}")

    # 测试除零
    zero_result = calc.divide(10, 0)
    print(f"\n10 / 0 = {zero_result}")


if __name__ == "__main__":
    main()
