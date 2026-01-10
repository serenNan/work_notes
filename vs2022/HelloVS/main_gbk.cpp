#include <iostream>
#include <windows.h>

int main() {
    SetConsoleOutputCP(936);

    std::cout << "你好, Visual Studio 2022!" << std::endl;
    std::cout << "这是一个简单的测试项目." << std::endl;

    int a = 10, b = 20;
    std::cout << a << " + " << b << " = " << (a + b) << std::endl;

    return 0;
}
