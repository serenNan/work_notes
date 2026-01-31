cmake_minimum_required(VERSION 3.16)
project(lsp_test LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_EXPORT_COMPILE_COMMANDS ON)

add_executable(calculator_test
    main.cpp
    calculator.cpp
)
