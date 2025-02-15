import os
import winreg
import logging
from typing import Optional
import glob
import subprocess
from utils.settings_manager import SettingsManager  # 添加导入

class WPSPathFinder:
    """查找WPS安装路径的工具类"""
    
    @staticmethod
    def get_wps_path() -> Optional[str]:
        """获取WPS路径，优先从设置文件读取，如果没有则搜索并保存"""
        # 先从设置文件中读取
        settings_manager = SettingsManager()
        settings = settings_manager.load_settings()
        
        wps_path = settings.get('wps_path')
        if wps_path and os.path.exists(wps_path):
            return wps_path
            
        # 如果设置文件中没有或路径无效，则搜索
        wps_path = WPSPathFinder.find_wps_path()
        if wps_path:
            # 找到后保存到设置文件
            settings['wps_path'] = wps_path
            settings_manager.save_settings(settings)
            
        return wps_path
    
    @staticmethod
    def find_wps_path() -> Optional[str]:
        """查找WPS可执行文件的路径"""
        try:
            # 1. 尝试使用where命令查找
            wps_path = WPSPathFinder._find_using_where()
            if wps_path:
                return wps_path
                
            # 2. 尝试使用进程查找
            wps_path = WPSPathFinder._find_from_process()
            if wps_path:
                return wps_path
                
            # 3. 尝试从快捷方式查找
            wps_path = WPSPathFinder._find_from_shortcuts()
            if wps_path:
                return wps_path
                
            # 4. 尝试从固定路径查找
            wps_path = WPSPathFinder._find_from_fixed_paths()
            if wps_path:
                return wps_path
            
            return None
            
        except Exception as e:
            logging.error(f"查找WPS路径时出错: {str(e)}")
            return None

    @staticmethod
    def _find_using_where() -> Optional[str]:
        """使用where命令查找wps.exe"""
        try:
            result = subprocess.run(['where', 'wps.exe'], 
                                 capture_output=True, 
                                 text=True)
            if result.returncode == 0:
                paths = result.stdout.strip().split('\n')
                for path in paths:
                    if os.path.exists(path):
                        return path
        except:
            pass
        return None

    @staticmethod
    def _find_from_process() -> Optional[str]:
        """从运行的进程中查找WPS路径"""
        try:
            # 使用wmic命令查找运行的WPS进程
            result = subprocess.run(
                ['wmic', 'process', 'where', 'name like "%wps.exe%"', 'get', 'ExecutablePath'],
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                paths = result.stdout.strip().split('\n')[1:]  # 跳过标题行
                for path in paths:
                    path = path.strip()
                    if path and os.path.exists(path):
                        return path
        except:
            pass
        return None

    @staticmethod
    def _find_from_shortcuts() -> Optional[str]:
        """从快捷方式查找WPS路径"""
        try:
            # 常见的快捷方式位置
            shortcut_locations = [
                os.path.join(os.environ['ProgramData'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
                os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
                os.path.join(os.environ['PUBLIC'], 'Desktop'),
                os.path.join(os.environ['USERPROFILE'], 'Desktop'),
            ]
            
            for location in shortcut_locations:
                if not os.path.exists(location):
                    continue
                    
                # 递归搜索.lnk文件
                for root, _, files in os.walk(location):
                    for file in files:
                        if 'wps' in file.lower() and file.endswith('.lnk'):
                            try:
                                # 使用PowerShell解析快捷方式
                                cmd = f'powershell -command "(New-Object -COM WScript.Shell).CreateShortcut(\'{os.path.join(root, file)}\').TargetPath"'
                                result = subprocess.run(cmd, capture_output=True, text=True)
                                if result.returncode == 0:
                                    target = result.stdout.strip()
                                    if target and os.path.exists(target) and target.lower().endswith('wps.exe'):
                                        return target
                            except:
                                continue
        except:
            pass
        return None

    @staticmethod
    def _find_from_fixed_paths() -> Optional[str]:
        """从固定路径查找WPS"""
        # 扩展搜索路径列表
        search_paths = [
            r"C:\Users\Public\WPS Cloud Files\WPS Office",
            r"C:\Program Files\WPS Office",
            r"C:\Program Files (x86)\WPS Office",
            r"C:\Program Files\Kingsoft\WPS Office",
            r"C:\Program Files (x86)\Kingsoft\WPS Office",
            # 添加用户目录下的可能位置
            os.path.join(os.environ['LOCALAPPDATA'], 'Kingsoft', 'WPS Office'),
            os.path.join(os.environ['LOCALAPPDATA'], 'WPS Office'),
            # 添加其他可能的驱动器
            r"D:\Program Files\WPS Office",
            r"D:\Program Files (x86)\WPS Office",
            r"E:\Program Files\WPS Office",
            r"E:\Program Files (x86)\WPS Office",
        ]
        
        # 遍历所有用户的Program Files目录
        users_dir = os.path.join(os.environ['SystemDrive'] + '\\', 'Users')
        if os.path.exists(users_dir):
            for user in os.listdir(users_dir):
                user_program_files = os.path.join(users_dir, user, 'AppData', 'Local')
                if os.path.exists(user_program_files):
                    search_paths.append(os.path.join(user_program_files, 'Kingsoft', 'WPS Office'))
                    search_paths.append(os.path.join(user_program_files, 'WPS Office'))

        for base_path in search_paths:
            if not os.path.exists(base_path):
                continue
                
            # 在每个基础路径下搜索可能的子目录
            for root, dirs, files in os.walk(base_path):
                if 'wps.exe' in files:
                    return os.path.join(root, 'wps.exe')
                
                # 检查特定的子目录
                office6_path = os.path.join(root, 'office6', 'wps.exe')
                if os.path.exists(office6_path):
                    return office6_path
                    
                bin_path = os.path.join(root, 'bin', 'wps.exe')
                if os.path.exists(bin_path):
                    return bin_path

        return None

def main():
    """测试WPS路径查找功能"""
    logging.basicConfig(level=logging.INFO)
    print("开始查找WPS路径...")
    
    # 尝试所有查找方法并显示结果
    methods = [
        ('Where命令查找', WPSPathFinder._find_using_where),
        ('进程查找', WPSPathFinder._find_from_process),
        ('快捷方式查找', WPSPathFinder._find_from_shortcuts),
        ('固定路径查找', WPSPathFinder._find_from_fixed_paths),
    ]
    
    for method_name, method in methods:
        print(f"\n尝试{method_name}...")
        try:
            result = method()
            if result:
                print(f"✓ 成功: {result}")
                print(f"文件存在: {os.path.exists(result)}")
            else:
                print("✗ 未找到")
        except Exception as e:
            print(f"✗ 错误: {str(e)}")
    
    # 使用主方法进行查找
    print("\n使用综合查找方法:")
    wps_path = WPSPathFinder.find_wps_path()
    if wps_path:
        print(f"最终找到WPS路径: {wps_path}")
        print(f"文件存在: {os.path.exists(wps_path)}")
    else:
        print("未找到WPS安装路径")

if __name__ == "__main__":
    main() 