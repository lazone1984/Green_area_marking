"""
CAD软件检测模块

功能说明:
- 检测当前系统中运行的CAD软件
- 支持检测多种常见CAD软件,包括:
  * AutoCAD
  * ZWCAD
  * GstarCAD 
  * SolidWorks
  * CATIA
  * Siemens NX
  * Autodesk Inventor

使用方法:
1. 直接运行此模块可以打印当前运行的CAD软件列表
2. 作为模块导入时,可以调用check_cad_running()函数获取运行中的CAD软件列表

依赖:
- psutil: 用于获取系统进程信息
"""

import psutil

def check_cad_running():
    # 定义常见CAD软件的进程名列表
    cad_process_names = [
        'ACAD.EXE',        # AutoCAD
        'ZWCAD.EXE',       # ZWCAD
        'GCAD.EXE',        # GstarCAD
        'SOLIDWORKS.EXE',  # SolidWorks
        'CATIA.EXE',       # CATIA
        'nx.exe',          # Siemens NX
        'inventor.exe'     # Autodesk Inventor
    ]
    
    # 获取当前运行的所有进程
    running_cad = []
    for proc in psutil.process_iter(['name']):
        try:
            # 检查进程名是否在CAD软件列表中
            if proc.info['name'].upper() in [x.upper() for x in cad_process_names]:
                running_cad.append(proc.info['name'])
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    
    return running_cad

def main():
    # 检测运行中的CAD软件
    running_cad_list = check_cad_running()
    
    if running_cad_list:
        # print("检测到以下CAD软件正在运行：")
        for cad in running_cad_list:
            print(f"- {cad}")
    else:
        print("未检测到任何CAD软件正在运行")

if __name__ == "__main__":
    main() 