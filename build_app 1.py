#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Prince应用打包脚本
用于将Prince_Operate_theme.py及相关文件打包成可执行文件
"""

import PyInstaller.__main__
import shutil
import os
import sys
import time
import subprocess
import platform
import argparse
from datetime import datetime

# 命令行参数解析
parser = argparse.ArgumentParser(description='Prince应用打包工具')
parser.add_argument('--name', default='PrinceOperate', help='应用程序名称')
parser.add_argument('--onefile', action='store_true', default=True, help='是否打包为单文件')
parser.add_argument('--console', action='store_true', default=False, help='是否显示控制台')
parser.add_argument('--icon', default='ch-2.ico', help='应用图标路径')
parser.add_argument('--version', default='1.0.0', help='应用版本号')
args = parser.parse_args()

# 打包配置参数
APP_NAME = args.name
APP_VERSION = args.version
ENTRY_FILE = 'Prince_Operate_theme.py'
ICON_PATH = os.path.abspath(args.icon)
WORK_PATH = os.path.abspath('./build')
DIST_PATH = os.path.abspath('./dist')
SPEC_PATH = os.path.abspath('./spec')
OUTPUT_PATH = os.path.join(DIST_PATH, f"{APP_NAME}-{APP_VERSION}")

# 日志函数
def log(message, level="INFO"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    level_colors = {
        "INFO": "\033[92m",  # 绿色
        "WARNING": "\033[93m",  # 黄色
        "ERROR": "\033[91m",  # 红色
        "SUCCESS": "\033[96m",  # 青色
    }
    reset_color = "\033[0m"
    
    if level in level_colors:
        print(f"{level_colors[level]}[{timestamp}] [{level}] {message}{reset_color}")
    else:
        print(f"[{timestamp}] [{level}] {message}")

# 检查环境
def check_environment():
    log("检查Python环境...")
    python_version = platform.python_version()
    log(f"Python版本: {python_version}")
    
    # 检查PyInstaller
    try:
        import PyInstaller
        pyinstaller_version = PyInstaller.__version__
        log(f"PyInstaller版本: {pyinstaller_version}")
    except ImportError:
        log("未安装PyInstaller，请先安装: pip install pyinstaller", "ERROR")
        return False
    
    # 检查必要文件
    required_files = [
        ENTRY_FILE,
        ICON_PATH,
        'Prince_Operate_Ui.ui',
        'Table_Ui.ui',
        'requirements.txt'
    ]
    
    missing_files = []
    for file_path in required_files:
        if not os.path.exists(file_path):
            missing_files.append(file_path)
    
    if missing_files:
        log(f"缺少必要文件: {', '.join(missing_files)}", "ERROR")
        return False
    
    log("环境检查通过", "SUCCESS")
    return True

# 安装依赖
def install_dependencies():
    log("安装项目依赖...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        log("依赖安装完成", "SUCCESS")
        return True
    except subprocess.CalledProcessError as e:
        log(f"依赖安装失败: {e}", "ERROR")
        return False

# 清理历史构建文件
def clean_build_files():
    log("清理历史构建文件...")
    paths_to_clean = [
        (WORK_PATH, "构建目录"),
        (DIST_PATH, "分发目录"),
        (SPEC_PATH, "规格文件目录")
    ]
    
    for path, desc in paths_to_clean:
        if os.path.exists(path):
            try:
                shutil.rmtree(path)
                log(f"已清理{desc}: {path}")
            except Exception as e:
                log(f"清理{desc}失败: {e}", "ERROR")
                return False
    
    log("清理完成", "SUCCESS")
    return True

# 收集资源文件
def collect_resources():
    log("收集资源文件...")
    resource_files = [
        'ch-2.ico',
        'Prince_Operate_Ui.ui',
        'Table_Ui.ui'
    ]
    
    datas = []
    for file_name in resource_files:
        if os.path.exists(file_name):
            datas.append((os.path.abspath(file_name), '.'))
            log(f"添加资源文件: {file_name}")
        else:
            log(f"警告: 找不到资源文件 '{file_name}'", "WARNING")
    
    return datas

# 构建PyInstaller参数
def build_pyinstaller_args(datas):
    log("构建PyInstaller参数...")
    
    # 核心模块
    hidden_imports = [
        # 核心UI模块
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'qt_material',
        
        # pandas核心模块
        'pandas',
        'pandas.plotting',
        'pandas.core.frame',
        'pandas.core.series',
        'openpyxl',
        'xlrd',
        'xlwt',
        'xlutils',
        'xlwings',
        'xlsxwriter',

        # 项目自身模块
        'theme_manager_theme',
        'Browser_operation',
        'Get_Data',
        'Prince_Operate_Ui',
        'Table_Operate',
        'Logger',
        'File_Operate',
        'chicon',
        
        # Playwright核心模块
        'playwright.sync_api',
    ]
    
    # 排除模块
    excludes = [
        'pandas.tests',
        'numpy.tests',
        'numpy.random.tests',
        'numpy.distutils.tests',
        'playwright.driver',
        'matplotlib',
        'scipy',
        'tkinter',
        'PySide2',
        'IPython',
        'jupyter',
        'zmq',
        'qtconsole',
        'node.exe',
        'ffmpeg.dll',
    ]
    
    # 构建参数
    pyinstaller_args = [
        f'--name={APP_NAME}',
        '--onefile' if args.onefile else '--onedir',
        '' if args.console else '--noconsole',
        f'--icon={ICON_PATH}',
        f'--workpath={WORK_PATH}',
        f'--distpath={DIST_PATH}',
        f'--specpath={SPEC_PATH}',
        '--clean',
        '--log-level=INFO',
    ]
    
    # 添加数据文件
    for src, dst in datas:
        pyinstaller_args.append(f'--add-data={src};{dst}')
    
    # 添加隐藏导入
    for imp in hidden_imports:
        pyinstaller_args.append(f'--hidden-import={imp}')
    
    # 添加排除模块
    for exc in excludes:
        pyinstaller_args.append(f'--exclude-module={exc}')
    
    # 添加collect-all，确保关键模块的所有必要组件都被包含
    key_modules = ['pandas.core', 'playwright.sync_api']
    for module in key_modules:
        pyinstaller_args.append(f'--collect-all={module}')
    
    # 添加入口文件
    pyinstaller_args.append(ENTRY_FILE)
    
    return pyinstaller_args

# 执行PyInstaller打包
def run_pyinstaller(pyinstaller_args):
    log("开始执行PyInstaller打包...")
    
    # 打印完整命令（便于调试）
    log("PyInstaller命令参数:")
    for arg in pyinstaller_args:
        log(f"  {arg}")
    
    try:
        PyInstaller.__main__.run(pyinstaller_args)
        log("PyInstaller打包完成", "SUCCESS")
        return True
    except Exception as e:
        log(f"PyInstaller打包失败: {e}", "ERROR")
        return False

# 创建版本信息文件
def create_version_file():
    version_info = f"""
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({APP_VERSION.replace('.', ', ')}, 0),
    prodvers=({APP_VERSION.replace('.', ', ')}, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          u'080404b0',
          [StringStruct(u'CompanyName', u'Chen-Huai'),
           StringStruct(u'FileDescription', u'Prince自动化操作工具'),
           StringStruct(u'FileVersion', u'{APP_VERSION}'),
           StringStruct(u'InternalName', u'{APP_NAME}'),
           StringStruct(u'LegalCopyright', u'Copyright (C) 2025 Chen-Huai'),
           StringStruct(u'OriginalFilename', u'{APP_NAME}.exe'),
           StringStruct(u'ProductName', u'Prince自动化操作工具'),
           StringStruct(u'ProductVersion', u'{APP_VERSION}')])
      ]
    ),
    VarFileInfo([VarStruct(u'Translation', [2052, 1200])])
  ]
)
"""
    version_file_path = os.path.abspath(f"version_{APP_NAME}.txt")
    with open(version_file_path, "w", encoding="utf-8") as f:
        f.write(version_info)
    
    log(f"创建版本信息文件: {version_file_path}")
    return version_file_path

# 创建README文件
def create_readme():
    readme_content = f"""# Prince自动化操作工具 v{APP_VERSION}

## 应用说明
这是一个网页自动化操作工具，用于Prince系统的自动化操作。

## 使用方法
1. 双击运行 {APP_NAME}.exe
2. 按照界面提示进行操作

## 配置文件
配置文件位于用户桌面的config文件夹中，名为config_prince.csv。
首次运行时会自动创建配置文件。

## 技术支持
如有问题，请联系开发者。

## 版本历史
- v{APP_VERSION}: 当前版本
"""
    
    readme_path = os.path.join(DIST_PATH, "README.txt")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(readme_content)
    
    log(f"创建README文件: {readme_path}")

# 主函数
def main():
    log(f"===== 开始打包 Prince自动化操作工具 v{APP_VERSION} =====", "INFO")
    start_time = time.time()
    
    # 检查环境
    if not check_environment():
        log("环境检查失败，打包终止", "ERROR")
        return False
    
    # 安装依赖
    if not install_dependencies():
        log("依赖安装失败，打包终止", "ERROR")
        return False
    
    # 清理历史构建文件
    if not clean_build_files():
        log("清理历史构建文件失败，打包终止", "ERROR")
        return False
    
    # 创建版本信息文件 - 仅创建但不使用，因为PyInstaller的version-file参数在Windows上有问题
    create_version_file()
    
    # 收集资源文件
    datas = collect_resources()
    
    # 构建PyInstaller参数
    pyinstaller_args = build_pyinstaller_args(datas)
    
    # 执行PyInstaller打包
    if not run_pyinstaller(pyinstaller_args):
        log("PyInstaller打包失败，打包终止", "ERROR")
        return False
    
    # 创建README文件
    create_readme()
    
    # 计算耗时
    end_time = time.time()
    elapsed_time = end_time - start_time
    minutes, seconds = divmod(elapsed_time, 60)
    
    log(f"===== 打包完成 =====", "SUCCESS")
    log(f"总耗时: {int(minutes)}分{int(seconds)}秒")
    log(f"可执行文件位置: {os.path.join(DIST_PATH, APP_NAME + '.exe')}")
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        if success:
            sys.exit(0)
        else:
            sys.exit(1)
    except KeyboardInterrupt:
        log("用户中断打包过程", "WARNING")
        sys.exit(1)
    except Exception as e:
        log(f"打包过程发生未知错误: {e}", "ERROR")
        sys.exit(1)
