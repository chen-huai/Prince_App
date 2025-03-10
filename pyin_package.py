import PyInstaller.__main__
import shutil
import os
import sys
import time

# 打包配置参数
APP_NAME = 'PrinceOperate'
ENTRY_FILE = 'Prince_Operate_theme.py'
ICON_PATH = os.path.abspath('ch-2.ico')
WORK_PATH = os.path.abspath('./build')
DIST_PATH = os.path.abspath('./dist')
SPEC_PATH = os.path.abspath('./spec')

print("===== 开始执行自动打包 =====")
start_time = time.time()

# 清理历史构建文件
if os.path.exists(WORK_PATH):
    print(f"清理历史构建目录: {WORK_PATH}")
    shutil.rmtree(WORK_PATH)
if os.path.exists(DIST_PATH):
    print(f"清理历史分发目录: {DIST_PATH}")
    shutil.rmtree(DIST_PATH)
if os.path.exists(SPEC_PATH):
    print(f"清理历史规格文件目录: {SPEC_PATH}")
    shutil.rmtree(SPEC_PATH)

# 验证资源文件路径
for file_path in ['ch-2.ico', 'Prince_Operate_Ui.ui', 'Table_Ui.ui']:
    if not os.path.exists(file_path):
        print(f"警告: 找不到资源文件 '{file_path}'")

# 收集需要的UI资源文件 - 使用绝对路径
datas = []
for file_name in ['ch-2.ico', 'Prince_Operate_Ui.ui', 'Table_Ui.ui']:
    if os.path.exists(file_name):
        datas.append((os.path.abspath(file_name), '.'))
        print(f"添加资源文件: {file_name}")

# 必要的隐藏导入 - 确保包含所有核心依赖
hidden_imports = [
    # 核心UI模块
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'qt_material',
    
    # pandas核心模块
    'pandas',
    'pandas.plotting',  # 确保包含这个核心依赖
    'pandas.core.frame',
    'pandas.core.series',
    
    # 项目自身模块
    'theme_manager_theme',
    'Browser_operation',
    'Get_Data',
    'Prince_Operate_Ui',
    'Table_Operate',
    'Logger',
    'File_Operate',
    'chicon',
    
    # Playwright核心模块 - 但不包括浏览器
    'playwright.sync_api',
]

# 构建排除模块列表 - 更保守的排除策略
excludes = [
    # 只排除确定不需要的测试模块
    'pandas.tests',
    'numpy.tests',
    'numpy.random.tests',
    'numpy.distutils.tests',
    
    # Playwright浏览器相关
    'playwright.driver',  # 浏览器驱动
    
    # 其他大型库
    'matplotlib',
    'scipy',
    'tkinter',
    'PySide2',
    'IPython',
    'jupyter',
    'zmq',
    'qtconsole',
]

print("===== 开始 PyInstaller 打包 =====")

# 构建 PyInstaller 命令行参数
pyinstaller_args = [
    '--name=%s' % APP_NAME,
    '--onefile',       # 单文件模式
    '--noconsole',     # 隐藏控制台
    '--icon=%s' % ICON_PATH,
    '--workpath=%s' % WORK_PATH,
    '--distpath=%s' % DIST_PATH,
    '--specpath=%s' % SPEC_PATH,
    '--clean',         # 清理缓存
    '--log-level=INFO', # 增加日志级别，以便查看更多信息
]

# 添加数据文件
for src, dst in datas:
    pyinstaller_args.append('--add-data=%s;%s' % (src, dst))

# 添加隐藏导入
for imp in hidden_imports:
    pyinstaller_args.append('--hidden-import=%s' % imp)

# 添加排除模块
for exc in excludes:
    pyinstaller_args.append('--exclude-module=%s' % exc)

# 添加collecct-all，确保关键模块的所有必要组件都被包含
key_modules = ['pandas.core', 'playwright.sync_api']
for module in key_modules:
    pyinstaller_args.append('--collect-all=%s' % module)

# 添加入口文件
pyinstaller_args.append(ENTRY_FILE)

# 打印完整的命令（便于调试）
print("\n最终 PyInstaller 命令参数:")
for arg in pyinstaller_args:
    print(f"  {arg}")

# 执行 PyInstaller
try:
    print("\n开始执行 PyInstaller...")
    PyInstaller.__main__.run(pyinstaller_args)
    print("PyInstaller 打包完成")
except Exception as e:
    print(f"错误: PyInstaller 打包失败: {e}")
    sys.exit(1)

# 计算耗时
end_time = time.time()
elapsed_time = end_time - start_time
minutes, seconds = divmod(elapsed_time, 60)

print(f"\n===== 打包完成 =====")
print(f"总耗时: {int(minutes)}分{int(seconds)}秒")
print(f"可执行文件位置: {os.path.join(DIST_PATH, APP_NAME + '.exe')}")