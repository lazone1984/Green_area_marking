import tkinter as tk
from tkinter import ttk, messagebox
from .window_manager import WindowManager
from .ui_components import UIComponents
from cad.plant_mark import PlantMark
import math
from .export_manager import ExportManager
from utils.settings_manager import SettingsManager
import os
import sys
import base64
from io import BytesIO
from assets.icon import ICON  # 从新模块导入图标
from PIL import Image, ImageTk
from assets.qr_codes import WECHAT_QR, ALIPAY_QR

class PlantMarkUI(UIComponents, WindowManager):
    def __init__(self):
        super().__init__()
        self.root = tk.Tk()
        self.root.title("CAD面积标注工具")
        self.root.geometry("460x550")
        self.root.resizable(False, False)  # 禁止调整窗口大小
        
        # 设置窗口图标
        try:
            # 解码base64图标数据
            icon_data = base64.b64decode(ICON)
            # 创建PhotoImage对象
            icon = tk.PhotoImage(data=icon_data)
            # 设置图标
            self.root.tk.call('wm', 'iconphoto', self.root._w, icon)
        except Exception as e:
            print(f"设置图标时出错: {str(e)}")
        
        # 初始化设置管理器
        self.settings_manager = SettingsManager()
        
        # 设置默认值
        self.default_settings = {
            "layer": "全部图层",
            "unit": "毫米",
            "text_height": "3.0",
            "redline_area": 0,
            "has_redline": False,
            "has_garage": False,
            "garage_points": [],
            "original_areas": [],
            "center_points": [],
            "last_drawing": "",
            "hatch_settings": {
                "pattern": "CROSS",
                "color": "绿",
                "angle": "0",
                "scale": "1"
            }
        }
        
        # 加载设置，如果没有则使用默认值
        self.settings = self.settings_manager.load_settings()
        for key, value in self.default_settings.items():
            if key not in self.settings:
                self.settings[key] = value
        
        # 创建 CAD 实例
        self.cad = PlantMark("AutoCAD.Application")
        
        # 设置 CAD 实例到组件中
        self.set_cad_instance(self.cad)
        self.cad.ui = self
        
        # 调用父类的初始化方法
        self.init_ui()
        self.init_window_handles()
        self.export_manager = ExportManager()
        
        # 恢复保存的设置
        self.restore_settings()
        
        # 绑定关闭窗口事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.components = UIComponents()
        self.components.root = self.root  # 确保 UIComponents 能访问到 root

    def restore_settings(self):
        """恢复保存的设置"""
        try:
            # 恢复图层选择
            if self.settings.get("layer") in self.layer_combo['values']:
                self.layer_var.set(self.settings["layer"])
            else:
                self.layer_var.set("全部图层")
            
            # 恢复单位选择
            self.unit_var.set(self.settings.get("unit", "毫米"))
            
            # 恢复字高
            self.text_height_var.set(self.settings.get("text_height", "3.0"))
            
            # 恢复红线面积
            self.redline_area = self.settings.get("redline_area", 0)
            
            # 恢复地库线坐标
            self.garage_points = self.settings.get("garage_points", [])
            
            # 恢复框选数据
            self.original_areas = self.settings.get("original_areas", [])
            self.center_points = self.settings.get("center_points", [])
            
            # 更新按钮状态
            if self.settings.get("has_redline", False):
                self.select_redline_button.configure(text="选择红线(已加载)")
            
            if self.settings.get("has_garage", False) and self.garage_points:
                self.select_garage_button.configure(text="选地库线(已加载)")
                # 如果有原始数据，重新计算折算系数
                if hasattr(self, 'original_areas') and self.original_areas:
                    self.update_area_list(self.original_areas)
                
        except Exception as e:
            print(f"恢复设置时出错: {str(e)}")

    def on_closing(self):
        """关闭窗口时保存设置"""
        self.settings.update({
            "layer": self.layer_var.get(),
            "unit": self.unit_var.get(),
            "text_height": self.text_height_var.get(),
            "redline_area": self.redline_area,
            "has_redline": self.redline_area > 0,
            "has_garage": bool(self.garage_points),
            "garage_points": self.garage_points,  # 保存地库线坐标
            "original_areas": self.original_areas,
            "center_points": self.center_points
        })
        
        self.settings_manager.save_settings(self.settings)
        self.root.destroy()

    def init_ui(self):
        """初始化UI组件"""
        self.create_main_frame()
        self.create_layer_selection()
        self.create_unit_selection()
        self.create_buttons()
        self.create_area_list()
        self.create_export_frame()
        
        # 获取CAD图层
        self.update_layer_list()

    def run(self):
        """运行UI主循环"""
        self.root.mainloop()

    def start_marking(self):
        """开始标注的回调函数"""
        try:
            selected_layer = self.layer_var.get()
            if selected_layer == "全部图层":
                layer_name = ["全部图层"]
            else:
                layer_name = [selected_layer]
                
            self.root.iconify()
            
            plant = PlantMark("Autocad.Application")
            plant.ui = self
            
            # 获取CAD文件名并保存
            cad_filename = os.path.basename(plant.doc.FullName)
            self.settings["cad_filename"] = cad_filename
            self.settings_manager.save_settings(self.settings)
            
            # 将CAD窗口置于最前
            plant.wincad.Visible = True
            plant.wincad.WindowState = 3  # 3 = 最大化
            self.switch_to_cad()
            
            areas, center_points = plant.applicate(layer_name)
            self.center_points = center_points
            
            if areas:
                self.original_areas = areas
                self.update_area_list(areas)
                
                # 保存框选数据
                self.settings["original_areas"] = areas
                self.settings["center_points"] = center_points
                self.settings_manager.save_settings(self.settings)
            
            # 设置到导出管理器
            self.export_manager.set_cad_name(cad_filename)
            
            self.switch_to_ui()
            
        except Exception as e:
            messagebox.showerror("错误", f"标注过程出错：{str(e)}")
            self.switch_to_ui()

    def select_redline(self):
        """选择红线并计算面积"""
        try:
            self.root.iconify()
            
            plant = PlantMark("Autocad.Application")
            
            # 将CAD窗口置于最前
            plant.wincad.Visible = True
            plant.wincad.WindowState = 3  # 3 = 最大化
            self.switch_to_cad()
            
            # 清空现有选择集
            while plant.doc.SelectionSets.Count > 0:
                plant.doc.SelectionSets.Item(0).Delete()
            
            # 创建选择集
            redline_select = plant.doc.SelectionSets.Add("redline_select")
            
            plant.doc.Utility.Prompt("请选择红线范围的多段线...")
            redline_select.SelectOnScreen()
            
            total_area = 0
            for obj in redline_select:
                if obj.ObjectName == "AcDbPolyline":
                    if obj.Closed:  # 确保多段线是闭合的
                        total_area = obj.Area
                        break
                elif obj.ObjectName == "AcDbCircle":
                    total_area = 3.14159 * obj.Radius * obj.Radius
                    break
                elif obj.ObjectName == "AcDbEllipse":
                    major_axis = obj.MajorAxis
                    major_radius = math.sqrt(major_axis[0]**2 + major_axis[1]**2)
                    minor_radius = major_radius * obj.RadiusRatio
                    total_area = 3.14159 * major_radius * minor_radius
                    break
            
            redline_select.Delete()
            
            if total_area > 0:
                self.redline_area = total_area
                if self.original_areas:
                    self.update_area_list(self.original_areas)
                # 更新按钮文本
                self.select_redline_button.configure(text="选择红线(已加载)")
            
            self.switch_to_ui()
            
        except Exception as e:
            messagebox.showerror("错误", f"选择红线时出错：{str(e)}")
            self.switch_to_ui()

    def select_garage(self):
        """选择地下车库线并记录边界"""
        try:
            self.root.iconify()
            
            plant = PlantMark("Autocad.Application")
            
            # 将CAD窗口置于最前
            plant.wincad.Visible = True
            plant.wincad.WindowState = 3  # 3 = 最大化
            self.switch_to_cad()
            
            # 清空现有选择集
            while plant.doc.SelectionSets.Count > 0:
                plant.doc.SelectionSets.Item(0).Delete()
            
            # 创建选择集
            garage_select = plant.doc.SelectionSets.Add("garage_select")
            
            plant.doc.Utility.Prompt("请选择地下车库范围的多段线...")
            garage_select.SelectOnScreen()
            
            # 获取地下车库边界点
            self.garage_points = []
            for obj in garage_select:
                if obj.ObjectName == "AcDbPolyline":
                    if obj.Closed:
                        coords = obj.Coordinates
                        points = []
                        for i in range(0, len(coords), 2):
                            points.append([coords[i], coords[i+1]])
                        self.garage_points = points
                        break
                elif obj.ObjectName == "AcDbCircle":
                    # 对于圆形，创建一个近似的多边形
                    center = obj.Center
                    radius = obj.Radius
                    points = []
                    for i in range(36):  # 36个点形成一个近似圆
                        angle = i * 10 * math.pi / 180
                        x = center[0] + radius * math.cos(angle)
                        y = center[1] + radius * math.sin(angle)
                        points.append([x, y])
                    self.garage_points = points
                    break
            
            garage_select.Delete()
            
            if self.garage_points:
                if self.original_areas:
                    self.update_area_list(self.original_areas)
                self.select_garage_button.configure(text="选地库线(已加载)")
                
                # 保存设置
                self.settings["has_garage"] = True
                self.settings["garage_points"] = self.garage_points
                self.settings_manager.save_settings(self.settings)
            
            self.switch_to_ui()
            
        except Exception as e:
            messagebox.showerror("错误", f"选择地下车库线时出错：{str(e)}")
            self.switch_to_ui()

    def update_area_list(self, areas):
        """更新面积列表"""
        try:
            # 清除现有列表
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            
            # 获取单位转换系数
            unit = self.unit_var.get()
            conversion = 0.000001 if unit == "米" else 1
            unit_symbol = "㎡" if unit == "米" else "㎡"
            
            # 添加每个面积项
            for i, area in enumerate(areas, 1):
                row_frame = ttk.Frame(self.scrollable_frame)
                row_frame.pack(fill=tk.X, pady=2)
                
                # 序号
                ttk.Label(row_frame, text=str(i), width=6).pack(side=tk.LEFT)
                
                # 实测面积（可编辑）
                actual_area = area * conversion
                area_var = tk.StringVar(value=f"{actual_area:.2f}")
                area_entry = ttk.Entry(row_frame, textvariable=area_var, width=15)
                area_entry.pack(side=tk.LEFT, padx=2)
                
                # 折算系数下拉框
                initial_factor = "80%" if (self.garage_points and 
                                        len(self.center_points) > i-1 and 
                                        self.is_point_in_garage(self.center_points[i-1])) else "100%"
                factor_var = tk.StringVar(value=initial_factor)
                
                self.factor_vars.append((factor_var, area_var))
                factor_combo = ttk.Combobox(row_frame, textvariable=factor_var,
                                          values=["100%", "80%", "50%", "30%", "10%"],
                                          width=8, state="readonly")
                factor_combo.pack(side=tk.LEFT, padx=5)
                
                # 禁用下拉框的鼠标滚轮事件
                def prevent_mousewheel(event):
                    return "break"
                
                factor_combo.bind("<MouseWheel>", prevent_mousewheel)
                
                # 折算面积（自动计算）
                converted_area = actual_area * float(factor_var.get().strip('%')) / 100
                converted_var = tk.StringVar(value=f"{converted_area:.2f}")
                ttk.Entry(row_frame, textvariable=converted_var, width=15,
                         state="readonly").pack(side=tk.LEFT, padx=2)
                
                # 绑定更新事件
                def update_converted(event, area_v=area_var, factor_v=factor_var, conv_v=converted_var):
                    try:
                        area = float(area_v.get())
                        factor = float(factor_v.get().strip('%')) / 100
                        conv_v.set(f"{area * factor:.2f}")
                        # 保存修改后的值到original_areas
                        if unit == "米":
                            area = area * 1000000  # 转回毫米
                        idx = int(event.widget.master.pack_info()['in'].winfo_children().index(event.widget.master)) - 1
                        if 0 <= idx < len(self.original_areas):
                            self.original_areas[idx] = area
                        self.calculate_total()
                    except ValueError:
                        pass
                
                # 绑定实时更新事件
                area_entry.bind('<KeyRelease>', update_converted)
                area_entry.bind('<FocusOut>', update_converted)
                factor_combo.bind('<<ComboboxSelected>>', update_converted)
            
            # 计算总计
            self.calculate_total()
            
            # 更新画布高度
            self.on_frame_configure()
            
        except Exception as e:
            messagebox.showerror("错误", f"更新面积列表时出错：{str(e)}")

    def calculate_total(self):
        """计算总计"""
        try:
            unit = self.unit_var.get()
            conversion = 0.000001 if unit == "米" else 1  # 转换系数
            unit_symbol = "㎡" if unit == "米" else "㎡"
            
            # 清除现有总计显示
            for widget in self.total_frame.winfo_children():
                widget.destroy()
            
            total_area = 0
            total_converted_area = 0
            
            # 获取当前显示的所有面积数据
            for widget in self.scrollable_frame.winfo_children():
                try:
                    # 获取每行的控件
                    children = widget.winfo_children()
                    if len(children) >= 4:  # 确保有足够的控件（序号、面积、系数、折算面积）
                        area_entry = children[1]  # 实测面积输入框
                        factor_combo = children[2]  # 折算系数下拉框
                        
                        actual_area = float(area_entry.get())
                        factor = float(factor_combo.get().strip('%')) / 100
                        
                        # 累加总计
                        total_area += actual_area
                        total_converted_area += actual_area * factor
                except (ValueError, IndexError):
                    continue
            
            # 显示总计
            ttk.Label(self.total_frame, text=f"总实测面积: {total_area:.2f}{unit_symbol}").pack(pady=2)
            ttk.Label(self.total_frame, text=f"总折算面积: {total_converted_area:.2f}{unit_symbol}").pack(pady=2)
            
            # 显示红线面积和绿地率
            if self.redline_area > 0:
                # 将红线面积转换为当前单位
                redline_area_display = self.redline_area * conversion
                ttk.Label(self.total_frame, text=f"红线面积: {redline_area_display:.2f}{unit_symbol}").pack(pady=2)
                
                # 计算绿地率（使用当前单位的值计算）
                green_ratio = (total_converted_area / redline_area_display) * 100
                ttk.Label(self.total_frame, text=f"绿地率 = 折算面积/红线面积 × 100% = {green_ratio:.2f}%").pack(pady=2)
                
        except Exception as e:
            print(f"计算总计时出错: {str(e)}")

    def show_help(self):
        """显示帮助说明窗口"""
        help_window = tk.Toplevel(self.root)
        help_window.title("使用帮助")
        help_window.geometry("700x600")  # 调整窗口高度
        
        # 创建标签页控件
        notebook = ttk.Notebook(help_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5, 0))  # 减小底部padding
        
        # 创建折算系数说明标签页
        factor_frame = ttk.Frame(notebook)
        notebook.add(factor_frame, text="折算系数说明")
        
        # 创建文本框和滚动条的容器
        factor_container = ttk.Frame(factor_frame)
        factor_container.pack(fill=tk.BOTH, expand=True)
        
        factor_text = tk.Text(factor_container, wrap=tk.WORD, padx=10, pady=10)
        factor_scroll = ttk.Scrollbar(factor_container, orient="vertical", command=factor_text.yview)
        
        # 先放置滚动条，再放置文本框
        factor_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        factor_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        factor_text.configure(yscrollcommand=factor_scroll.set)
        
        factor_content = """根据《园林绿化工程施工及验收规范》(CJJ 82-2012)第5.3.2.5条规定：

住宅地块屋顶绿化不计入绿地面积。

除住宅以外的建设工程项目按要求实施屋顶绿化面积的，屋顶绿化面积可按照下列比例计算为绿地面积：

a) 覆土厚度1.5m以上的，按100%计算
b) 覆土厚度1.0m以上不足1.5m的，按80%计算绿地面积
c) 覆土厚度0.5m以上不足1.0m的，按50%计算绿地面积
d) 覆土厚度0.3m以上不足0.5m的，按30%计算绿地面积
e) 覆土厚度0.1m以上不足0.3m的，按10%计算绿地面积
f) 覆土厚度不足0.1m的，不计算绿地面积"""
        
        factor_text.insert(tk.END, factor_content)
        factor_text.config(state=tk.DISABLED)
        
        # 创建软件使用说明标签页
        guide_frame = ttk.Frame(notebook)
        notebook.add(guide_frame, text="软件使用说明")
        
        # 创建文本框和滚动条的容器
        guide_container = ttk.Frame(guide_frame)
        guide_container.pack(fill=tk.BOTH, expand=True)
        
        guide_text = tk.Text(guide_container, wrap=tk.WORD, padx=10, pady=10)
        guide_scroll = ttk.Scrollbar(guide_container, orient="vertical", command=guide_text.yview)
        
        # 先放置滚动条，再放置文本框
        guide_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        guide_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        guide_text.configure(yscrollcommand=guide_scroll.set)
        
        guide_content = """1. 基本操作流程
   a) 启动前准备：
      - 确保CAD软件已经打开并加载了需要标注的图纸
      - 程序会自动检测CAD软件是否运行
   
   b) 界面设置：
      - 从下拉菜单选择要标注的CAD图层（可选择"全部图层"）
      - 选择计量单位（米或毫米）
      - 设置标注文字的高度
   
   c) 面积标注：
      - 点击"开始标注"按钮
      - 在CAD图中框选需要计算面积的区域
      - 程序会自动计算面积并显示在列表中

2. 红线面积设置
   a) 点击"选择红线"按钮
   b) 在CAD图中选择红线范围（支持多段线、圆形、椭圆形）
   c) 程序自动计算红线面积并用于绿地率计算
   d) 红线数据会自动保存，下次打开时自动加载

3. 地库线设置
   a) 点击"选地库线"按钮
   b) 在CAD图中选择地库范围（支持多段线、圆形）
   c) 程序自动识别位于地库范围内的绿地
   d) 对地库范围内的绿地自动设置80%的折算系数
   e) 地库线数据会自动保存，下次打开时自动加载

4. 面积数据管理
   a) 实测面积：
      - 可直接编辑数值
      - 支持米/毫米单位自动转换
   
   b) 折算系数：
      - 提供5档预设：100%、80%、50%、30%、10%
      - 根据《园林绿化工程施工及验收规范》自动计算
   
   c) 自动计算：
      - 实时计算折算面积
      - 自动统计总实测面积和总折算面积
      - 自动计算绿地率

5. 数据导出功能
   a) 支持多种格式：
      - Word文档
      - Excel表格
      - PowerPoint演示文稿
      - WPS文档
      - 插入CAD图纸
   
   b) 导出内容包括：
      - 序号和实测面积
      - 折算系数和折算面积
      - 面积汇总数据
      - 红线面积和绿地率
      - 自动带入CAD图纸名称

6. 数据保存功能
   程序会自动保存以下设置：
   - 选择的图层
   - 单位选择和字高
   - 红线面积数据
   - 地库线范围数据
   - 已标注的面积数据
   - 中心点坐标信息
   - CAD文件名

7. 其他功能
   a) 窗口管理：
      - 自动切换CAD和程序窗口
      - 支持最小化和还原
   
   b) 图层管理：
      - 自动获取CAD图层列表
      - 支持选择特定图层或全部图层
   
   c) 错误处理：
      - 提供详细的错误提示
      - 自动保存防止数据丢失

8. 使用技巧
   a) 标注过程中：
      - 程序会自动最小化，方便CAD操作
      - 完成后自动恢复窗口
   
   b) 数据编辑：
      - 双击数值可直接编辑
      - 修改后自动重新计算
   
   c) 快捷操作：
      - Tab键快速切换输入框
      - 鼠标滚轮浏览面积列表"""
        
        guide_text.insert(tk.END, guide_content)
        guide_text.config(state=tk.DISABLED)
        
        # 创建投币说明标签页
        donate_frame = ttk.Frame(notebook)
        notebook.add(donate_frame, text="投币说明")
        
        # 创建文本框和滚动条的容器
        donate_container = ttk.Frame(donate_frame)
        donate_container.pack(fill=tk.X)  # 移除expand=True，只填充水平方向
        
        donate_text = tk.Text(donate_container, wrap=tk.WORD, padx=10, pady=20)
        donate_text.configure(height=20)  # 设置文本框高度为20行
        donate_scroll = ttk.Scrollbar(donate_container, orient="vertical", command=donate_text.yview)
        
        donate_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        donate_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        donate_text.configure(yscrollcommand=donate_scroll.set)
        
        # 添加说明文字
        donate_content = """感谢您使用本软件！

本软件完全免费使用，没有任何强制付费要求。如果您觉得这个工具对您有帮助，欢迎随心投币支持开发者继续改进完善。

您的支持将用于：
- 持续优化软件功能
- 修复已知问题
- 开发新的实用功能
- 提供更好的技术支持

有任何问题欢迎联系本人。
QQ邮箱：lazone@qq.com

无论您是否投币，都可以免费使用本软件的全部功能。感谢您的支持与理解！

扫描下方二维码进行投币："""
        
        donate_text.insert(tk.END, donate_content)
        
        # 添加收款码图片
        try:
            def load_and_resize_image(base64_data, size=(180, 180)):
                # 解码 base64 数据
                img_data = base64.b64decode(base64_data)
                # 使用 BytesIO 创建内存文件对象
                img_file = BytesIO(img_data)
                # 使用 PIL 打开并缩放图片
                img = Image.open(img_file)
                # 计算等比缩放尺寸
                width, height = img.size
                ratio = min(size[0]/width, size[1]/height)
                new_size = (int(width*ratio), int(height*ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                return ImageTk.PhotoImage(img)
            
            # 加载并缩放图片
            wechat_img = load_and_resize_image(WECHAT_QR)
            alipay_img = load_and_resize_image(ALIPAY_QR)
            
            # 创建图片容器框架，放在文字下方
            img_frame = ttk.Frame(donate_frame)
            img_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
            
            # 创建水平排列的容器
            qr_container = ttk.Frame(img_frame)
            qr_container.pack()
            
            # 微信收款码（左侧）
            wechat_container = ttk.Frame(qr_container)
            wechat_container.pack(side=tk.LEFT, padx=20)
            wechat_label = ttk.Label(wechat_container, image=wechat_img)
            wechat_label.image = wechat_img  # 保持引用防止被回收
            wechat_label.pack()
            ttk.Label(wechat_container, text="微信收款码").pack(pady=(5, 0))
            
            # 支付宝收款码（右侧）
            alipay_container = ttk.Frame(qr_container)
            alipay_container.pack(side=tk.LEFT, padx=20)
            alipay_label = ttk.Label(alipay_container, image=alipay_img)
            alipay_label.image = alipay_img  # 保持引用防止被回收
            alipay_label.pack()
            ttk.Label(alipay_container, text="支付宝收款码").pack(pady=(5, 0))
            
        except Exception as e:
            donate_text.insert(tk.END, "\n\n抱歉，加载收款码图片失败。")
            print(f"加载收款码图片出错: {str(e)}")
        
        donate_text.config(state=tk.DISABLED)
        
        # 添加确定按钮
        ok_button = ttk.Button(help_window, text="确定", command=help_window.destroy)
        ok_button.pack(pady=(5, 10))  # 调整按钮的上下间距
        
        # 设置窗口模态
        help_window.transient(self.root)
        help_window.grab_set()
        self.root.wait_window(help_window)

    def is_point_in_garage(self, point):
        """判断点是否在地下车库范围内"""
        if not self.garage_points:
            return False
            
        # 使用射线法判断点是否在多边形内
        x, y = point[0], point[1]
        inside = False
        j = len(self.garage_points) - 1
        
        for i in range(len(self.garage_points)):
            if ((self.garage_points[i][1] > y) != (self.garage_points[j][1] > y) and
                x < (self.garage_points[j][0] - self.garage_points[i][0]) * 
                (y - self.garage_points[i][1]) / 
                (self.garage_points[j][1] - self.garage_points[i][1]) + 
                self.garage_points[i][0]):
                inside = not inside
            j = i
            
        return inside 

    def on_unit_change(self, event=None):
        """处理单位切换"""
        try:
            # 更新字高
            current_height = float(self.text_height_var.get())
            if self.unit_var.get() == "米":
                # 从毫米转换到米
                new_height = current_height * 1000  # 毫米转米时扩大1000倍
            else:
                # 从米转换到毫米
                new_height = current_height / 1000  # 米转毫米时缩小1000倍
            self.text_height_var.set(f"{new_height:.1f}")
            
            # 更新面积列表
            if self.original_areas:
                self.update_area_list(self.original_areas)
        except ValueError:
            pass

    def update_layer_list(self):
        """更新图层列表"""
        try:
            # 获取所有图层名称
            layer_names = ["全部图层"]  # 默认选项
            for i in range(self.cad.doc.Layers.Count):
                layer = self.cad.doc.Layers.Item(i)
                if not layer.Name.startswith("*"):  # 排除系统图层
                    layer_names.append(layer.Name)
            
            # 更新下拉列表
            self.layer_combo['values'] = layer_names
            
            # 如果当前选择的图层不在列表中，重置为"全部图层"
            if self.layer_var.get() not in layer_names:
                self.layer_var.set("全部图层")
                
        except Exception as e:
            print(f"获取CAD图层失败: {str(e)}")
            # 设置默认值
            self.layer_combo['values'] = ["全部图层"]
            self.layer_var.set("全部图层") 

    def on_export(self, event=None):
        """处理导出事件"""
        try:
            # 准备导出数据
            export_data = []
            unit = self.unit_var.get()
            unit_symbol = "㎡" if unit == "米" else "㎡"
            
            # 获取当前显示的所有面积数据
            for widget in self.scrollable_frame.winfo_children():
                try:
                    children = widget.winfo_children()
                    if len(children) >= 4:
                        area_entry = children[1]
                        factor_combo = children[2]
                        
                        actual_area = float(area_entry.get())
                        factor = factor_combo.get()
                        factor_value = float(factor.strip('%')) / 100
                        converted_area = actual_area * factor_value
                        
                        export_data.append({
                            'actual_area': f"{actual_area:.2f}",  # 移除单位
                            'factor': factor,
                            'converted_area': f"{converted_area:.2f}"  # 移除单位
                        })
                except (ValueError, IndexError):
                    continue
            
            # 添加总计和其他数据（合并后三列）
            summary_data = []
            total_area = sum(float(d['actual_area']) for d in export_data)
            total_converted = sum(float(d['converted_area']) for d in export_data)
            
            summary_data.append({
                'actual_area': f"{total_area:.2f}",
                'merged_content': f"总实测面积: {total_area:.2f}{unit_symbol}"
            })
            
            summary_data.append({
                'actual_area': f"{total_converted:.2f}",
                'merged_content': f"总折算面积: {total_converted:.2f}{unit_symbol}"
            })
            
            if self.redline_area > 0:
                redline_area_converted = self.redline_area * (0.000001 if unit == "米" else 1)
                summary_data.append({
                    'actual_area': f"{redline_area_converted:.2f}",
                    'merged_content': f"红线面积: {redline_area_converted:.2f}{unit_symbol}"
                })
                
                green_ratio = (total_converted / redline_area_converted) * 100
                summary_data.append({
                    'actual_area': "",
                    'merged_content': f"绿地率 = 折算面积/红线面积 × 100% = {green_ratio:.2f}%"
                })
            
            # 根据选择的格式导出
            export_type = self.export_var.get()
            if export_type == "导出到Word":
                self.export_manager.export_to_word(export_data, summary_data, unit_symbol)
            elif export_type == "导出到Excel":
                self.export_manager.export_to_excel(export_data, summary_data, unit_symbol)
            elif export_type == "导出到PowerPoint":
                self.export_manager.export_to_ppt(export_data, summary_data, unit_symbol)
            
            # 重置下拉列表
            self.export_var.set("选择导出格式")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}") 

    def set_icon(self, icon_path):
        """此方法不再需要，可以删除或保留为空"""
        pass 