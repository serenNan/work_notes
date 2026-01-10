#include "utils.h"
#include <iostream>

void print_greeting(const std::string& name) {
    std::cout << "  你好, " << name << "!" << std::endl;
}

size_t get_string_length(const std::string& str) {
    return str.length();
}
