"""
PlantMark 项目结构

项目概述:
PlantMark 是一个CAD图纸处理工具，用于计算和标注CAD图纸中的面积，支持导出到Office文档。

目录结构:
PlantMark/
├── main.py # 程序入口文件
├── assets/ # 资源文件目录
│ ├── icon.py # 程序图标数据
│ ├── qr_codes.py # 收款码数据
│ └── convert_icon.py # 图标转换工具
├── ui/ # 用户界面相关模块
│ ├── plant_mark_ui.py # 主界面实现
│ ├── ui_components.py # UI组件
│ ├── window_manager.py # 窗口管理
│ └── export_manager.py # 导出功能管理
├── cad/ # CAD相关模块
│ ├── plant_mark.py # CAD操作核心功能
│ └── cad_detector.py # CAD软件检测模块
├── utils/ # 工具模块
│ └── settings_manager.py # 设置管理
└── README.md # 项目说明文档

功能说明:
1. CAD面积计算与标注
   - 自动计算CAD图纸中选定区域的面积
   - 支持多种标注样式和格式
   - 批量处理多个区域

2. 文档导出功能
   - 支持导出到Word文档
   - 支持导出到Excel表格
   - 支持导出到PPT演示文稿
   - 兼容WPS和Microsoft Office

3. 设置管理
   - 保存用户偏好设置
   - 自定义标注样式
   - 导出模板配置

模块说明:
- main.py: 程序入口点，负责启动检查和初始化
- ui/plant_mark_ui.py: 主界面实现，包含用户交互逻辑
- ui/ui_components.py: 自定义UI组件库
- ui/window_manager.py: 窗口管理器，控制程序窗口行为
- ui/export_manager.py: 文档导出功能实现
- cad/plant_mark.py: CAD交互核心，处理面积计算和标注
- cad/cad_detector.py: 检测CAD软件运行状态
- utils/settings_manager.py: 用户设置管理工具
- utils/wps_path_finder.py: WPS软件路径检测工具

"""

import os
from ui.plant_mark_ui import PlantMarkUI
from cad.cad_detector import check_cad_running
from tkinter import messagebox
import sys
import tkinter as tk
import base64
from assets.icon import ICON

def main():
    # 检测CAD软件是否运行
    running_cad_list = check_cad_running()
    
    if not running_cad_list:
        # 如果没有检测到CAD软件运行，显示提示框并退出程序
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        # 设置消息框图标
        try:
            import tempfile
            from PIL import Image, ImageTk
            from io import BytesIO
            
            # 创建临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_file:
                # 解码base64数据并写入临时文件
                tmp_file.write(base64.b64decode(ICON))
                tmp_file.flush()
                
                # 使用PIL打开图像
                image = Image.open(tmp_file.name)
                # 确保图像是正确的格式
                image = image.convert('RGBA')
                # 创建合适大小的图标
                image = image.resize((32, 32), Image.Resampling.LANCZOS)
                
                # 转换为PhotoImage
                photo = ImageTk.PhotoImage(image)
                
                # 设置图标
                root.iconphoto(True, photo)
                
            # 删除临时文件
            os.unlink(tmp_file.name)
            
        except Exception as e:
            print(f"设置图标时出错: {str(e)}")
            
        messagebox.showwarning(
            title="提醒",
            message="未检测到任何CAD软件运行！\n请先启动CAD软件后再运行本程序。"
        )
        root.destroy()
        sys.exit()
    
    # CAD软件正在运行，继续启动程序
    ui = PlantMarkUI()
    
    # 初始化标题（如果有上次的图纸记录）
    if hasattr(ui, 'components'):
        ui.components.update_title()
    
    ui.run()

if __name__ == '__main__':
    main() 