import PyInstaller.__main__
import os
import shutil

"""

导出为exe程序，在终端运行 python build_exe.py

"""

# 清理之前的构建文件
if os.path.exists("dist"):
    shutil.rmtree("dist")
if os.path.exists("build"):
    shutil.rmtree("build")
if os.path.exists("excel_modifier_simple.spec"):
    os.remove("excel_modifier_simple.spec")

# 使用PyInstaller打包
PyInstaller.__main__.run([
    'excel_modifier_simple.py',
    '--name=财务报表智能生成',
    '--windowed',  # 不显示控制台窗口
    '--onefile',   # 打包成单个exe文件
    '--clean',     # 清理临时文件
])