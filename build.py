"""
打包脚本说明:

功能:
- 将 Python 项目打包成单个可执行文件(.exe)
- 自动生成多尺寸图标文件
- 配置打包参数和依赖项

主要步骤:
1. 图标处理
   - 解码 base64 格式的图标数据
   - 生成多种尺寸的图标(16x16 到 256x256)
   - 保存为临时 .ico 文件

2. 打包配置
   - 设置主程序文件(main.py)
   - 配置输出文件名(PlantMark)
   - 设置 GUI 模式
   - 打包为单文件
   - 配置工作目录和输出目录
   - 添加资源文件和模块
   - 添加隐式依赖

3. 执行打包
   - 调用 PyInstaller 进行打包
   - 完成后清理临时文件
"""

import PyInstaller.__main__
import os
import base64
import tempfile
from PIL import Image
from assets.icon import ICON
from io import BytesIO

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 创建临时图标文件
icon_path = None
try:
    # 解码base64数据
    icon_data = base64.b64decode(ICON)
    
    # 使用 PIL 打开图像
    image = Image.open(BytesIO(icon_data))
    
    # 确保图像是正确的格式
    image = image.convert('RGBA')
    
    # 创建所需的图标尺寸
    sizes = [(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]
    icon_images = []
    
    for size in sizes:
        # 调整图像大小
        resized_image = image.resize(size, Image.Resampling.LANCZOS)
        icon_images.append(resized_image)
    
    # 创建临时 ICO 文件
    icon_path = os.path.join(current_dir, 'temp_icon.ico')
    icon_images[0].save(icon_path, format='ICO', sizes=sizes, append_images=icon_images[1:])
    
except Exception as e:
    print(f"创建临时图标文件时出错: {str(e)}")
    icon_path = None

# 定义打包参数
params = [
    'main.py',  # 主程序文件
    '--name=PlantMark',  # 生成的 exe 文件名
    '--windowed',  # 使用 GUI 模式
    '--onefile',  # 打包成单个文件
    '--noconfirm',  # 如果 dist 文件夹存在则覆盖
    '--clean',  # 清理临时文件
    f'--workpath={os.path.join(current_dir, "build")}',  # 工作目录
    f'--distpath={os.path.join(current_dir, "dist")}',  # 输出目录
    '--add-data=assets;assets',  # 添加资源文件
    '--add-data=ui;ui',  # 添加 UI 模块
    '--add-data=cad;cad',  # 添加 CAD 模块
    '--add-data=utils;utils',  # 添加工具模块
    '--hidden-import=PIL',  # 添加隐式依赖
    '--hidden-import=PIL._tkinter_finder',
    '--hidden-import=win32com.client',
]

# 如果成功创建了临时图标文件，添加图标参数
if icon_path:
    params.append(f'--icon={icon_path}')

try:
    # 运行打包命令
    PyInstaller.__main__.run(params)
finally:
    # 清理临时图标文件
    if icon_path and os.path.exists(icon_path):
        os.unlink(icon_path)

print("打包完成！exe 文件位于 dist 目录下") 