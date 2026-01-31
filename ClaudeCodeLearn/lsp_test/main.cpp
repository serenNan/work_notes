#include "calculator.hpp"
#include <iostream>

int main() {
    // 测试 LSP: 命名空间和类引用
    math::Calculator calc("MyCalc");

    // 测试 LSP: 方法调用和自动补全
    double r1 = calc.add(10, 5);
    double r2 = calc.subtract(10, 5);
    double r3 = calc.multiply(3, 4);
    auto r4 = calc.divide(20, 4);

    std::cout << "Calculator: " << calc.getName() << "\n";
    std::cout << "10 + 5 = " << r1 << "\n";
    std::cout << "10 - 5 = " << r2 << "\n";
    std::cout << "3 * 4 = " << r3 << "\n";

    if (r4.has_value()) {
        std::cout << "20 / 4 = " << r4.value() << "\n";
    }

    // 测试 LSP: 查看历史记录
    const auto& history = calc.getHistory();
    std::cout << "\nHistory (" << history.size() << " operations):\n";
    for (const auto& op : history) {
        std::cout << "  - " << op << "\n";
    }

    // 测试 LSP: 工厂函数和常量
    auto calc2 = math::createCalculator("SciCalc");
    std::cout << "\nCreated: " << calc2.getName() << "\n";

    double circleArea = math::PI * 5 * 5;
    std::cout << "Circle area (r=5): " << circleArea << "\n";
    std::cout << "Euler's number: " << math::E << "\n";

    // 测试除零
    auto zeroResult = calc.divide(10, 0);
    std::cout << "\n10 / 0 = ";
    if (zeroResult.has_value()) {
        std::cout << zeroResult.value() << "\n";
    } else {
        std::cout << "Error\n";
    }

    return 0;
}
