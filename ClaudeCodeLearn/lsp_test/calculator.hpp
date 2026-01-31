#pragma once

#include <string>
#include <vector>
#include <optional>

namespace math {

class Calculator {
public:
    explicit Calculator(const std::string& name = "Calculator");
    ~Calculator() = default;

    // 基本运算
    double add(double a, double b);
    double subtract(double a, double b);
    double multiply(double a, double b);
    std::optional<double> divide(double a, double b);

    // 历史记录
    const std::vector<std::string>& getHistory() const;
    void clearHistory();
    std::string getName() const;

private:
    void record(const std::string& operation);

    std::string name_;
    std::vector<std::string> history_;
};

// 工厂函数
Calculator createCalculator(const std::string& name);

// 常量
constexpr double PI = 3.14159265358979;
constexpr double E = 2.71828182845905;

} // namespace math
