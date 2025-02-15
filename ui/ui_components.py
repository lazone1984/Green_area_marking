"""
## UI 组件模块

本模块定义了 `UIComponents` 类，它负责 CAD 区域计算和标注工具的图形用户界面。它提供以下功能：

- 图层选择和管理
- 单位选择和文字高度设置
- 面积测量和计算
- 导出为各种格式（Word、Excel、PowerPoint、CAD）的功能
- 填充图案控制和设置
- 带滚动的面积列表显示
- 总面积计算

UI 使用 tkinter 和 ttk 部件构建，组织成以下框架：
- 图层选择
- 单位和文字设置
- 控制按钮
- 带有测量值的面积列表
- 导出和填充图案控制

主要特点：
- 支持公制单位（mm、m）
- 可配置的文字高度和标注图层
- 带有转换因子的交互式面积列表
- 多种导出格式选项
- 可定制的填充图案和颜色
- 会话间设置持久性

依赖项：
- tkinter
- cad.plant_mark
- ui.export_manager

"""

import tkinter as tk
from tkinter import ttk
from cad.plant_mark import PlantMark
from tkinter import messagebox  # 添加在文件开头的导入部分

class UIComponents:
    def __init__(self):
        # 初始化变量
        self.redline_area = 0
        self.original_areas = []
        self.center_points = []
        self.garage_points = []
        self.factor_vars = []
        self.cad = None
        
        # 添加 export_manager 的初始化
        from ui.export_manager import ExportManager
        self.export_manager = ExportManager()

        # 添加填充相关的变量初始化
        self.hatch_pattern_var = None
        self.hatch_pattern_combo = None
        self.hatch_color_var = None
        self.hatch_scale_var = None
        
        # 从设置中加载填充配置
        self.load_hatch_settings()
        
        # 添加颜色映射
        self.color_map = {
            "红": 1,
            "黄": 2,
            "绿": 3,
            "青": 4,
            "蓝": 5,
            "洋红": 6,
            "白": 7,
            "灰": 8,
            "默认": 3  # 默认使用绿色
        }

        self.current_dwg = ""  # 添加当前图纸名称变量
        
        # 从设置中加载上次的图纸名称
        self.load_last_drawing()

    def create_main_frame(self):
        """创建主框架"""
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def create_layer_selection(self):
        """创建图层选择组件"""
        # 创建一个框架来容纳图层选择和标注类型
        self.layer_frame = ttk.Frame(self.main_frame)
        self.layer_frame.grid(row=0, column=0, pady=5, sticky=(tk.W, tk.E))
        
        # 图层标签和下拉框
        self.layer_label = ttk.Label(self.layer_frame, text="选择图层:")
        self.layer_label.pack(side=tk.LEFT, padx=(0,3))
        
        self.layer_var = tk.StringVar(value="全部图层")
        self.layer_combo = ttk.Combobox(self.layer_frame, textvariable=self.layer_var,
                                      width=30, state="readonly")
        self.layer_combo.pack(side=tk.LEFT, padx=3)
        
        # 添加标注类型选择
        self.mark_type_label = ttk.Label(self.layer_frame, text="标注类型:")
        self.mark_type_label.pack(side=tk.LEFT, padx=(10,3))
        
        self.mark_type_var = tk.StringVar(value="标记")
        self.mark_type_combo = ttk.Combobox(self.layer_frame, 
                                          textvariable=self.mark_type_var,
                                          values=["标记", "数字", "综合"],
                                          width=6,
                                          state="readonly")
        self.mark_type_combo.pack(side=tk.LEFT, padx=3)

    def create_unit_selection(self):
        """创建单位选择组件"""
        self.unit_frame = ttk.Frame(self.main_frame)
        self.unit_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        
        # 单位选择
        self.unit_label = ttk.Label(self.unit_frame, text="单位:")
        self.unit_label.pack(side=tk.LEFT, padx=3)  # 减小边距
        
        self.unit_var = tk.StringVar(value="毫米")
        self.unit_combo = ttk.Combobox(self.unit_frame, textvariable=self.unit_var,
                                     values=["毫米", "米"], width=6,  # 减小宽度
                                     state="readonly")
        self.unit_combo.pack(side=tk.LEFT, padx=3)  # 减小边距
        
        # 绑定单位切换事件
        self.unit_combo.bind('<<ComboboxSelected>>', self.on_unit_change)
        
        # 字高输入
        self.text_height_label = ttk.Label(self.unit_frame, text="字高:")
        self.text_height_label.pack(side=tk.LEFT, padx=(10, 3))  # 调整间距
        
        self.text_height_var = tk.StringVar(value="3.0")
        self.text_height_entry = ttk.Entry(self.unit_frame, textvariable=self.text_height_var,
                                         width=4)  # 减小宽度
        self.text_height_entry.pack(side=tk.LEFT, padx=3)

        # 添加标注图层选择
        self.annotation_layer_label = ttk.Label(self.unit_frame, text="标注图层:")
        self.annotation_layer_label.pack(side=tk.LEFT, padx=(10, 3))  # 调整间距
        
        self.annotation_layer_var = tk.StringVar(value="0-绿化面积标注")
        self.annotation_layer_combo = ttk.Combobox(self.unit_frame, 
                                                  textvariable=self.annotation_layer_var,
                                                  values=["0-绿化面积标注", "0-面积标注", "0-绿化标注"],
                                                  width=12,  # 减小宽度
                                                  state="readonly")
        self.annotation_layer_combo.pack(side=tk.LEFT, padx=3)

        # 添加新建图层按钮
        self.new_layer_button = ttk.Button(self.unit_frame, 
                                          text="+",
                                          width=2,  # 减小宽度
                                          command=self.create_new_layer)  # 添加回命令绑定
        self.new_layer_button.pack(side=tk.LEFT, padx=3)

    def create_buttons(self):
        """创建按钮组件"""
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        # 创建选择红线按钮
        self.select_redline_button = ttk.Button(self.button_frame, text="选择红线",
                                              command=self.select_redline)
        self.select_redline_button.pack(side=tk.LEFT, padx=5)
        
        # 创建选择地下车库线按钮
        self.select_garage_button = ttk.Button(self.button_frame, text="选地库线",
                                            command=self.select_garage)
        self.select_garage_button.pack(side=tk.LEFT, padx=5)
        
        # 创建开始标注按钮
        self.start_button = ttk.Button(self.button_frame, text="框选标注",
                                    command=self.start_marking)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        # 创建帮助按钮
        self.help_button = ttk.Button(self.button_frame, text="帮助说明",
                                    command=self.show_help)
        self.help_button.pack(side=tk.LEFT, padx=5)

    def create_area_list(self):
        """创建面积列表区域"""
        # 创建一个主容器来包含列表和总计区域
        self.list_container = ttk.Frame(self.main_frame)
        self.list_container.grid(row=4, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        
        # 创建列表区域
        self.list_frame = ttk.Frame(self.list_container)
        self.list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表头框架
        self.header_frame = ttk.Frame(self.list_frame)
        self.header_frame.pack(fill=tk.X)
        
        # 设置固定宽度
        seq_width = 6
        area_width = 15
        factor_width = 15
        converted_width = 15
        
        # 创建表头
        ttk.Label(self.header_frame, text="序号", width=seq_width).pack(side=tk.LEFT)
        ttk.Label(self.header_frame, text="实测面积", width=area_width).pack(side=tk.LEFT)
        ttk.Label(self.header_frame, text="折算系数", width=factor_width).pack(side=tk.LEFT)
        ttk.Label(self.header_frame, text="折算面积", width=converted_width).pack(side=tk.LEFT)
        
        # 创建带滚动条的列表区域
        list_container = ttk.Frame(self.list_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        # 设置固定高度（7行 * 每行高度30像素）
        fixed_height = 7 * 30
        
        # 创建画布和滚动条
        self.canvas = tk.Canvas(list_container, height=fixed_height)
        self.scrollbar = ttk.Scrollbar(list_container, orient="vertical",
                                     command=self.canvas.yview)
        
        # 创建可滚动框架
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # 设置画布的固定宽度
        total_width = (seq_width + area_width + factor_width + converted_width) * 8 + 20
        self.canvas.configure(width=total_width)
        
        # 创建画布窗口
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw",
                                width=total_width)
        
        # 配置滚动
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 放置画布和滚动条
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定鼠标滚轮事件
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # 添加总计区域
        self.total_frame = ttk.Frame(self.list_container)
        self.total_frame.pack(fill=tk.X, pady=(5, 0))

    def _on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        if self.canvas.winfo_exists():
            # 检查鼠标是否在折算系数下拉框上
            widget = event.widget
            if isinstance(widget, ttk.Combobox):
                return  # 如果是下拉框，不处理滚轮事件
            
            # 检查鼠标位置是否在任何下拉框上
            x, y = event.x_root, event.y_root
            for combo in self.scrollable_frame.winfo_children():
                for child in combo.winfo_children():
                    if isinstance(child, ttk.Combobox):
                        try:
                            combo_x = child.winfo_rootx()
                            combo_y = child.winfo_rooty()
                            combo_width = child.winfo_width()
                            combo_height = child.winfo_height()
                            
                            if (combo_x <= x <= combo_x + combo_width and 
                                combo_y <= y <= combo_y + combo_height):
                                return  # 如果鼠标在下拉框上，不处理滚轮事件
                        except:
                            continue
            
            # 处理画布的滚动
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def on_frame_configure(self, event=None):
        """处理滚动区域大小变化"""
        # 只更新滚动区域，不改变画布高度
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def create_export_frame(self):
        """创建导出和填充控件"""
        self.export_frame = ttk.Frame(self.main_frame)
        self.export_frame.grid(row=6, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        
        # 创建导出按钮
        ttk.Label(self.export_frame, text="导出结果:").pack(side=tk.LEFT, padx=3)
        
        # 更新导出选项下拉列表
        self.export_var = tk.StringVar(value="选择导出格式")
        self.export_combo = ttk.Combobox(self.export_frame, 
                                       textvariable=self.export_var,
                                       values=["导出到Word", "导出到WPS", "导出到Excel", "导出到PowerPoint", "插入到CAD"],
                                       width=12,
                                       state="readonly")
        self.export_combo.pack(side=tk.LEFT, padx=3)
        
        # 修改绑定方式
        self.export_combo.bind('<<ComboboxSelected>>', lambda e: self.handle_export())

        # 添加分隔符和填充标题
        ttk.Label(self.export_frame, text="|").pack(side=tk.LEFT, padx=5)
        ttk.Label(self.export_frame, text="填充控制:").pack(side=tk.LEFT, padx=3)
        
        # 创建填充按钮（移到这里）
        self.hatch_button = ttk.Button(self.export_frame,
                                    text="应用填充",
                                    command=self.apply_hatch,
                                    width=8)
        self.hatch_button.pack(side=tk.LEFT, padx=3)
        
        # 创建填充控制框架
        self.fill_control_frame = ttk.Frame(self.main_frame)
        self.fill_control_frame.grid(row=7, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        
        # 创建填充样式选择
        ttk.Label(self.fill_control_frame, text="填充样式:").grid(row=0, column=0, padx=3, pady=2)
        self.hatch_pattern_var = tk.StringVar(value=self._hatch_settings['pattern'])
        self.hatch_pattern_combo = ttk.Combobox(self.fill_control_frame,
                                              textvariable=self.hatch_pattern_var,
                                              width=6,
                                              state="readonly")
        self.hatch_pattern_combo.grid(row=0, column=1, padx=3, pady=2)
        
        # 创建颜色选择
        ttk.Label(self.fill_control_frame, text="颜色:").grid(row=0, column=2, padx=(10,3), pady=2)
        self.hatch_color_var = tk.StringVar(value=self._hatch_settings['color'])
        self.color_combo = ttk.Combobox(self.fill_control_frame,
                                      textvariable=self.hatch_color_var,
                                      values=list(self.color_map.keys()),
                                      width=3,
                                      state="readonly")
        self.color_combo.grid(row=0, column=3, padx=3, pady=2)
        
        # 创建填充角度输入
        ttk.Label(self.fill_control_frame, text="角度:").grid(row=0, column=4, padx=(10,3), pady=2)
        self.hatch_angle_var = tk.StringVar(value=self._hatch_settings['angle'])
        angle_entry = ttk.Entry(self.fill_control_frame, textvariable=self.hatch_angle_var, width=3)
        angle_entry.grid(row=0, column=5, padx=3, pady=2)
        
        # 创建填充比例输入
        ttk.Label(self.fill_control_frame, text="比例:").grid(row=0, column=6, padx=(10,3), pady=2)
        self.hatch_scale_var = tk.StringVar(value=self._hatch_settings['scale'])
        self.hatch_scale_combo = ttk.Combobox(self.fill_control_frame,
                                            textvariable=self.hatch_scale_var,
                                            values=["10000","1000", "500", "100", "50", "10", "5", "1", "0.5", "0.1", "0.05", "0.01", "0.001"],
                                            width=6,
                                            state="readonly")
        self.hatch_scale_combo.grid(row=0, column=7, padx=3, pady=2)
        
        # 绑定值改变事件来保存设置
        self.hatch_pattern_combo.bind('<<ComboboxSelected>>', lambda e: self.save_hatch_settings())
        self.color_combo.bind('<<ComboboxSelected>>', lambda e: self.save_hatch_settings())
        self.hatch_angle_var.trace_add('write', lambda *args: self.save_hatch_settings())
        self.hatch_scale_combo.bind('<<ComboboxSelected>>', lambda e: self.save_hatch_settings())
        
        # 初始化时获取CAD中的填充样式
        self.update_hatch_patterns()

    def create_new_layer(self):
        """创建新的标注图层"""
        # 创建一个顶层窗口
        dialog = tk.Toplevel(self.root)
        dialog.title("新建标注图层")
        dialog.geometry("300x100")  # 减小窗口高度，因为只需要一个输入框
        dialog.resizable(False, False)
        
        # 使对话框模态
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 创建输入框架
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建图层名称输入
        ttk.Label(frame, text="图层名称:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        name_var = tk.StringVar(value="绿化面积标注")
        name_entry = ttk.Entry(frame, textvariable=name_var, width=25)
        name_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        name_entry.select_range(0, tk.END)  # 选中默认文本
        name_entry.focus()  # 设置焦点
        
        def confirm():
            """确认创建新图层"""
            # 获取输入值并处理
            name = name_var.get().strip()
            
            if not name:  # 如果名称为空
                tk.messagebox.showwarning("警告", "图层名称不能为空！")
                return
            
            # 检查是否已存在相同名称的图层
            current_layers = list(self.annotation_layer_combo['values'])
            if name in current_layers:
                tk.messagebox.showwarning("警告", f"图层 '{name}' 已存在！")
                return
            
            # 更新图层列表
            current_layers.append(name)
            self.annotation_layer_combo['values'] = current_layers
            
            # 选中新建的图层
            self.annotation_layer_var.set(name)
            
            # 关闭对话框
            dialog.destroy()
        
        def cancel():
            """取消创建"""
            dialog.destroy()
        
        # 创建按钮框架
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        # 添加确认和取消按钮
        ttk.Button(button_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=cancel).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键到确认按钮
        dialog.bind('<Return>', lambda e: confirm())
        dialog.bind('<Escape>', lambda e: cancel())
        
        # 设置对话框位置为屏幕中心
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f'{width}x{height}+{x}+{y}')

    def set_cad_instance(self, cad_instance):
        """设置 CAD 实例"""
        self.cad = cad_instance
        if self.cad and hasattr(self.cad, 'doc'):  # 确保 CAD 实例有效
            self.export_manager.set_cad_instance(cad_instance)
            # 确保界面元素都已创建后再更新填充样式
            if hasattr(self, 'root'):
                self.root.after(100, self.update_hatch_patterns)

    def update_hatch_patterns(self):
        """设置默认填充样式"""
        try:
            if not hasattr(self, 'hatch_pattern_combo') or not self.hatch_pattern_combo:
                print("填充样式下拉框未初始化")
                return
            
            # 使用固定的三种样式，将 SOLID 改为 GRASS
            default_patterns = ["ANSI31", "CROSS", "GRASS"]
            self.hatch_pattern_combo['values'] = default_patterns
            
            # 设置默认值为 CROSS
            self.hatch_pattern_var.set("CROSS")
            
        except Exception as e:
            print(f"设置填充样式时出错: {str(e)}")

    def apply_hatch(self):
        """应用填充"""
        try:
            if self.cad:  # 使用 self.cad 而不是 hasattr 检查
                pattern = self.hatch_pattern_var.get()
                scale = float(self.hatch_scale_var.get())
                self.cad.apply_hatch(pattern, scale)
        except Exception as e:
            tk.messagebox.showerror("错误", f"应用填充时出错: {str(e)}") 

    def keep_hatch(self):
        """保留当前填充"""
        try:
            if self.cad:
                # 提示用户操作已完成
                self.cad.doc.Utility.Prompt("\n填充已保留")
        except Exception as e:
            tk.messagebox.showerror("错误", f"保留填充时出错: {str(e)}") 

    def handle_export(self):
        """处理导出事件"""
        export_type = self.export_var.get()
        if not export_type or export_type == "选择导出格式":
            return
        
        try:
            # 获取当前数据
            data = []
            for frame in self.scrollable_frame.winfo_children():
                if len(frame.winfo_children()) >= 4:  # 确保有足够的子组件
                    area_entry = frame.winfo_children()[1]
                    factor_combo = frame.winfo_children()[2]
                    converted_label = frame.winfo_children()[3]
                    
                    data.append({
                        'actual_area': area_entry.get(),
                        'factor': factor_combo.get(),
                        'converted_area': converted_label.cget('text')
                    })
            
            # 获取汇总数据
            summary_data = []
            for child in self.total_frame.winfo_children():
                if isinstance(child, ttk.Label):
                    text = child.cget('text')
                    if text:  # 只添加非空文本
                        summary_data.append({
                            'merged_content': text
                        })
            
            # 获取单位符号
            unit_symbol = "㎡" if self.unit_var.get() == "米" else "㎟"
            
            # 根据选择调用相应的导出方法
            if export_type == "导出到Word":
                self.export_manager.export_to_word(data, summary_data, unit_symbol)
            elif export_type == "导出到WPS":
                self.export_manager.export_to_wps(data, summary_data, unit_symbol)
            elif export_type == "导出到Excel":
                self.export_manager.export_to_excel(data, summary_data, unit_symbol)
            elif export_type == "导出到PowerPoint":
                self.export_manager.export_to_ppt(data, summary_data, unit_symbol)
            elif export_type == "插入到CAD":
                if not self.cad:
                    messagebox.showerror("错误", "未找到活动的CAD实例")
                    return
                self.export_manager.export_to_cad(data, summary_data, unit_symbol)
            
            # 重置选择
            self.export_var.set("选择导出格式")
            
        except Exception as e:
            messagebox.showerror("导出错误", f"导出失败：{str(e)}")
            # 打印详细错误信息到控制台
            import traceback
            traceback.print_exc() 

    def load_hatch_settings(self):
        """从配置文件加载填充设置"""
        try:
            from utils.settings_manager import SettingsManager
            # 创建 SettingsManager 实例
            settings_manager = SettingsManager()
            settings = settings_manager.load_settings()
            
            # 获取填充设置，如果不存在则使用默认值
            hatch_settings = settings.get('hatch_settings', {
                'pattern': 'CROSS',
                'color': '绿',
                'angle': '0',
                'scale': '1'
            })
            
            # 保存设置以供后续使用
            self._hatch_settings = hatch_settings
            
        except Exception as e:
            print(f"加载填充设置时出错: {str(e)}")
            # 使用默认值
            self._hatch_settings = {
                'pattern': 'CROSS',
                'color': '绿',
                'angle': '0',
                'scale': '1'
            }

    def save_hatch_settings(self):
        """保存填充设置到配置文件"""
        try:
            from utils.settings_manager import SettingsManager
            # 创建 SettingsManager 实例
            settings_manager = SettingsManager()
            settings = settings_manager.load_settings()
            
            # 更新填充设置
            settings['hatch_settings'] = {
                'pattern': self.hatch_pattern_var.get(),
                'color': self.hatch_color_var.get(),
                'angle': self.hatch_angle_var.get(),
                'scale': self.hatch_scale_var.get()
            }
            
            # 保存设置
            settings_manager.save_settings(settings)
            
        except Exception as e:
            print(f"保存填充设置时出错: {str(e)}") 

    def load_last_drawing(self):
        """从设置中加载上次打开的图纸名称"""
        try:
            from utils.settings_manager import SettingsManager
            settings_manager = SettingsManager()  # 创建实例
            settings = settings_manager.load_settings()  # 调用实例方法
            self.current_dwg = settings.get('last_drawing', '')
        except Exception as e:
            print(f"加载上次图纸名称时出错: {str(e)}")
            self.current_dwg = ""

    def save_current_drawing(self, drawing_name):
        """保存当前图纸名称到设置"""
        try:
            from utils.settings_manager import SettingsManager
            settings_manager = SettingsManager()  # 创建实例
            settings = settings_manager.load_settings()
            settings['last_drawing'] = drawing_name
            settings_manager.save_settings(settings)  # 调用实例方法
        except Exception as e:
            print(f"保存当前图纸名称时出错: {str(e)}")

    def update_title(self, drawing_name=None):
        """更新程序标题"""
        if drawing_name:
            self.current_dwg = drawing_name
            self.save_current_drawing(drawing_name)
            self.root.title(f"CAD面积标注工具 - {drawing_name}")
        elif self.current_dwg:
            self.root.title(f"CAD面积标注工具 - {self.current_dwg}")
        else:
            self.root.title("CAD面积标注工具")

    def start_marking(self):
        """开始标注"""
        try:
            selected_layer = self.layer_var.get()
            if selected_layer == "全部图层":
                layer_name = ["全部图层"]
            else:
                layer_name = [selected_layer]
            
            self.root.iconify()
            
            plant = PlantMark("Autocad.Application")
            plant.ui = self
            
            # 获取CAD文件名并更新标题
            current_drawing = plant.get_current_drawing_name()
            if current_drawing:
                self.update_title(current_drawing)
            
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
            self.export_manager.set_cad_name(current_drawing)
            
            self.switch_to_ui()
            
        except Exception as e:
            messagebox.showerror("错误", f"标注过程出错：{str(e)}")
            self.switch_to_ui() 