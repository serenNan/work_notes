#include <iostream>
#include <windows.h>
#include "calculator.h"
#include "utils.h"

int main() {
    SetConsoleOutputCP(936);  // GBK

    std::cout << "=== 多文件项目演示 ===" << std::endl;
    std::cout << std::endl;

    // 使用计算器类
    Calculator calc;
    std::cout << "计算器测试:" << std::endl;
    std::cout << "  10 + 5 = " << calc.add(10, 5) << std::endl;
    std::cout << "  10 - 5 = " << calc.subtract(10, 5) << std::endl;
    std::cout << "  10 * 5 = " << calc.multiply(10, 5) << std::endl;
    std::cout << "  10 / 5 = " << calc.divide(10, 5) << std::endl;
    std::cout << std::endl;

    // 使用工具函数
    std::cout << "工具函数测试:" << std::endl;
    print_greeting("世界");
    std::cout << "  字符串长度: " << get_string_length("你好世界") << std::endl;
    std::cout << std::endl;

    std::cout << "程序运行完成!" << std::endl;
    return 0;
}
