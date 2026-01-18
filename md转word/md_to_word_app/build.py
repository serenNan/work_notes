# -*- coding: utf-8 -*-
"""
PyInstaller 打包脚本
使用方法: python build.py
"""

import os
import sys
import subprocess
import shutil

def check_pyinstaller():
    """检查 PyInstaller 是否已安装"""
    try:
        import PyInstaller
        return True
    except ImportError:
        return False

def install_pyinstaller():
    """安装 PyInstaller"""
    print("Installing PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

def find_pyqt5_plugins():
    """查找 PyQt5 插件目录"""
    try:
        from PyQt5.QtCore import QLibraryInfo
        plugins_path = QLibraryInfo.location(QLibraryInfo.PluginsPath)
        if os.path.exists(plugins_path):
            return plugins_path
    except:
        pass

    # 备选路径
    import PyQt5
    pyqt5_dir = os.path.dirname(PyQt5.__file__)
    for subdir in ['Qt5/plugins', 'Qt/plugins', 'plugins']:
        path = os.path.join(pyqt5_dir, subdir)
        if os.path.exists(path):
            return path

    return None

def build():
    """执行打包"""
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 应用名称 (使用英文避免编码问题)
    app_name = "MD2Word"

    # 主程序入口
    main_script = os.path.join(current_dir, "main.py")

    # assets 目录
    assets_dir = os.path.join(current_dir, "assets")

    # 输出目录
    dist_dir = os.path.join(current_dir, "dist")
    build_dir = os.path.join(current_dir, "build")

    # 查找 PyQt5 插件
    plugins_path = find_pyqt5_plugins()

    # PyInstaller 命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", app_name,
        "--windowed",  # 不显示控制台窗口
        "--onefile",   # 打包成单个 exe
        "--noconfirm", # 不询问确认
        "--clean",     # 清理临时文件
        # 添加 assets 目录
        "--add-data", f"{assets_dir};assets",
        # 隐藏导入
        "--hidden-import", "PyQt5.sip",
        "--hidden-import", "docx",
        "--hidden-import", "docx.oxml",
        "--hidden-import", "docx.oxml.ns",
        "--hidden-import", "lxml",
        "--hidden-import", "lxml.etree",
        "--hidden-import", "lxml._elementpath",
        # 输出目录
        "--distpath", dist_dir,
        "--workpath", build_dir,
    ]

    # 如果找到插件目录，添加到打包
    if plugins_path:
        cmd.extend(["--add-data", f"{plugins_path};PyQt5/Qt5/plugins"])
        print(f"Found PyQt5 plugins: {plugins_path}")

    # 主程序
    cmd.append(main_script)

    print(f"Building {app_name}...")
    print(f"Command: {' '.join(cmd)}")

    # 执行打包
    result = subprocess.run(cmd, cwd=current_dir)

    if result.returncode == 0:
        exe_path = os.path.join(dist_dir, f"{app_name}.exe")
        print(f"\nBuild successful!")
        print(f"Executable: {exe_path}")
        print(f"\nNote: Pandoc is required to run this application")
        print(f"Download Pandoc: https://pandoc.org/installing.html")
    else:
        print(f"\nBuild failed, exit code: {result.returncode}")
        return False

    return True

def main():
    print("=" * 50)
    print("MD2Word - Build Script")
    print("=" * 50)

    # 检查并安装 PyInstaller
    if not check_pyinstaller():
        install_pyinstaller()

    # 执行打包
    success = build()

    if success:
        print("\nBuild complete!")
    else:
        print("\nBuild failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()
