import json
import os

class SettingsManager:
    def __init__(self):
        self.settings_file = "user_settings.json"
        self.default_settings = {
            "layer": "全部图层",
            "unit": "毫米",
            "text_height": "3.0",
            "redline_area": 0,
            "has_redline": False,
            "has_garage": False,
            "garage_points": [],
            "original_areas": [],  # 添加框选的面积数据
            "center_points": [],   # 添加中心点数据
            "cad_filename": ""     # 添加CAD文件名
        }
        
    def save_settings(self, settings):
        """保存设置"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            return True
        except:
            return False
    
    def load_settings(self):
        """加载设置"""
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return {}
        return {} 