#include "calculator.hpp"
#include <sstream>

namespace math {

Calculator::Calculator(const std::string& name) : name_(name) {}

double Calculator::add(double a, double b) {
    double result = a + b;
    std::ostringstream oss;
    oss << a << " + " << b << " = " << result;
    record(oss.str());
    return result;
}

double Calculator::subtract(double a, double b) {
    double result = a - b;
    std::ostringstream oss;
    oss << a << " - " << b << " = " << result;
    record(oss.str());
    return result;
}

double Calculator::multiply(double a, double b) {
    double result = a * b;
    std::ostringstream oss;
    oss << a << " * " << b << " = " << result;
    record(oss.str());
    return result;
}

std::optional<double> Calculator::divide(double a, double b) {
    std::ostringstream oss;
    if (b == 0.0) {
        oss << a << " / " << b << " = Error (division by zero)";
        record(oss.str());
        return std::nullopt;
    }
    double result = a / b;
    oss << a << " / " << b << " = " << result;
    record(oss.str());
    return result;
}

void Calculator::record(const std::string& operation) {
    history_.push_back(operation);
}

const std::vector<std::string>& Calculator::getHistory() const {
    return history_;
}

void Calculator::clearHistory() {
    history_.clear();
}

std::string Calculator::getName() const {
    return name_;
}

Calculator createCalculator(const std::string& name) {
    return Calculator(name);
}

} // namespace math
