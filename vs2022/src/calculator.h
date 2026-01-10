#ifndef CALCULATOR_H
#define CALCULATOR_H

// 简单计算器类
class Calculator {
public:
    Calculator();
    ~Calculator();

    int add(int a, int b);       // 加法
    int subtract(int a, int b);  // 减法
    int multiply(int a, int b);  // 乘法
    double divide(int a, int b); // 除法
};

#endif // CALCULATOR_H
