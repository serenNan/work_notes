#include "calculator.h"
#include <iostream>

Calculator::Calculator() {
    std::cout << "  [计算器初始化]" << std::endl;
}

Calculator::~Calculator() {
    std::cout << "  [计算器销毁]" << std::endl;
}

int Calculator::add(int a, int b) {
    return a + b;
}

int Calculator::subtract(int a, int b) {
    return a - b;
}

int Calculator::multiply(int a, int b) {
    return a * b;
}

double Calculator::divide(int a, int b) {
    if (b == 0) {
        std::cout << "  错误: 除数不能为零!" << std::endl;
        return 0;
    }
    return static_cast<double>(a) / b;
}
