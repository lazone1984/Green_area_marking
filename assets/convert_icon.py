import base64
import os
from PIL import Image

def convert_icon_to_base64(icon_path):
    # 打开图标文件并转换为PNG格式
    with Image.open(icon_path) as img:
        # 确保图像是RGBA模式
        img = img.convert('RGBA')
        # 调整大小为32x32
        img = img.resize((32, 32), Image.Resampling.LANCZOS)
        # 创建临时文件保存PNG
        temp_path = icon_path + '.png'
        img.save(temp_path, 'PNG')
        
        # 读取PNG文件并转换为base64
        with open(temp_path, 'rb') as f:
            base64_str = base64.b64encode(f.read()).decode()
            
        # 删除临时文件
        os.remove(temp_path)
        return base64_str

def main():
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 构建 logo.ico 的完整路径
    icon_path = os.path.join(current_dir, 'logo.ico')
    
    # 检查文件是否存在
    if not os.path.exists(icon_path):
        print(f"错误: 找不到图标文件 {icon_path}")
        return
    
    try:
        # 转换图标为base64
        icon_base64 = convert_icon_to_base64(icon_path)
        
        # 构建 icon.py 的完整路径
        output_path = os.path.join(current_dir, 'icon.py')
        
        # 创建 icon.py 文件
        with open(output_path, 'w') as f:
            f.write('# logo.ico 的 base64 编码 (PNG格式)\n')
            f.write('ICON = """\n')
            f.write(icon_base64)
            f.write('\n"""')
        
        print(f"转换完成！图标已保存到 {output_path}")
        
    except Exception as e:
        print(f"转换过程中出错: {str(e)}")

if __name__ == '__main__':
    main() 