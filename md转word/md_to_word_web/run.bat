@echo off
REM Markdown 转 Word Web 应用启动脚本
REM 支持 Conda 虚拟环境

chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

echo.
echo =============================================
echo  Markdown 转 Word Web 应用
echo =============================================
echo.

REM 检查当前目录
cd /d "%~dp0"
echo 工作目录: %cd%
echo.

REM 尝试使用 base 环境的 Python 直接运行
echo 检查 Python...
python --version 2>nul
if errorlevel 1 (
    echo 错误：找不到 Python
    echo 请确保已安装 Miniconda 或 Anaconda
    echo.
    pause
    exit /b 1
)

echo.
echo 检查依赖...
pip --version 2>nul
if errorlevel 1 (
    echo 错误：找不到 pip
    pause
    exit /b 1
)

echo.
echo 安装依赖...
pip install -r requirements.txt
if errorlevel 1 (
    echo 错误：安装依赖失败
    echo.
    pause
    exit /b 1
)

echo.
echo 检查 Pandoc...
where pandoc >nul 2>&1
if errorlevel 1 (
    echo 未检测到 Pandoc
    echo.
    echo 请运行以下命令安装 Pandoc:
    echo   conda install -c conda-forge pandoc
    echo.
    echo 或访问: https://pandoc.org/installing.html
    echo.
) else (
    echo Pandoc 已安装
)

echo.
echo =============================================
echo 启动应用...
echo 打开浏览器访问: http://localhost:5000
echo =============================================
echo.

REM 启动应用
python app.py

echo.
echo 应用已关闭
pause
endlocal
