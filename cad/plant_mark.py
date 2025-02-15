import math
import time
from datetime import datetime
import os
from .cad_utils import CadUtils

class PlantMark(CadUtils):
    def __init__(self, app_name):
        super().__init__(app_name)
        self.ui = None
        self.ucs_matrix = None  # 存储UCS变换矩阵

    def get_ucs_matrix(self):
        """获取当前UCS变换矩阵"""
        try:
            # 获取当前UCS
            current_ucs = self.doc.GetVariable("UCSORG")  # UCS原点
            ucs_xaxis = self.doc.GetVariable("UCSXDIR")   # UCS X轴方向
            ucs_yaxis = self.doc.GetVariable("UCSYDIR")   # UCS Y轴方向
            
            # 构建变换矩阵
            self.ucs_matrix = {
                'origin': [current_ucs[0], current_ucs[1]],
                'xaxis': [ucs_xaxis[0], ucs_xaxis[1]],
                'yaxis': [ucs_yaxis[0], ucs_yaxis[1]]
            }
            return True
        except Exception as e:
            print(f"获取UCS信息时出错: {str(e)}")
            self.ucs_matrix = None
            return False

    def transform_point(self, point):
        """将点从WCS转换到UCS"""
        if not self.ucs_matrix:
            return point
        
        try:
            # 将点相对于UCS原点进行平移
            x = point[0] - self.ucs_matrix['origin'][0]
            y = point[1] - self.ucs_matrix['origin'][1]
            
            # 计算在UCS中的坐标
            ux = (x * self.ucs_matrix['xaxis'][0] + y * self.ucs_matrix['xaxis'][1])
            uy = (x * self.ucs_matrix['yaxis'][0] + y * self.ucs_matrix['yaxis'][1])
            
            return [ux, uy]
        except Exception as e:
            print(f"坐标转换时出错: {str(e)}")
            return point

    def transform_point_to_wcs(self, point):
        """将点从UCS转换回WCS"""
        if not self.ucs_matrix:
            return point
        
        try:
            # 计算在WCS中的坐标
            x = (point[0] * self.ucs_matrix['xaxis'][0] + 
                 point[1] * self.ucs_matrix['yaxis'][0] + 
                 self.ucs_matrix['origin'][0])
            y = (point[0] * self.ucs_matrix['xaxis'][1] + 
                 point[1] * self.ucs_matrix['yaxis'][1] + 
                 self.ucs_matrix['origin'][1])
            
            return [x, y]
        except Exception as e:
            print(f"坐标转换回WCS时出错: {str(e)}")
            return point

    def applicate(self, layer_name):
        """应用标注"""
        try:
            # 调用绘制方法
            areas, center_points = self.draw_leader(layer_name)
            if hasattr(self, 'ui'):
                self.ui.update_area_list(areas)
            return areas, center_points
        except Exception as e:
            print(f"应用标注时出错: {str(e)}")
            return [], []

    def draw_leader(self, layer_name):
        """绘制引线和标注的主方法"""
        try:
            # 获取UCS信息
            self.get_ucs_matrix()
            
            # 获取当前选择的标注图层名称
            annotation_layer_name = None
            if hasattr(self, 'ui') and hasattr(self.ui, 'annotation_layer_var'):
                annotation_layer_name = self.ui.annotation_layer_var.get()
            
            # 创建或获取标注图层
            try:
                new_layer = self.doc.Layers.Add(annotation_layer_name)
                new_layer.Color = 3  # 设置图层颜色为绿色
            except:
                # 如果图层已存在，获取该图层
                new_layer = self.doc.Layers.Item(annotation_layer_name)
            
            count = 0
            polyline_averge_xy = []
            area_set = []
            direct_fontsize = []
            unclose_line_number = 0
            lenth_set = []
            
            # 清空选择集
            while self.doc.SelectionSets.Count > 0:
                self.doc.SelectionSets.Item(0).Delete()
                
            # 创建选择集
            object_select = self.doc.SelectionSets.Add("cad_object_select")

            # 提示用户选择对象
            self.doc.Utility.Prompt("请选择要标注的图形...")
            object_select.SelectOnScreen()
            
            # 处理选择的对象
            filtered_objects = []
            self.ui.original_objects = []  # 新增：保存原始对象和它们的图层
            for obj in object_select:
                # 如果指定了特定图层且对象不在该图层上，则跳过
                if layer_name != ["全部图层"] and obj.Layer != layer_name[0]:
                    continue
                filtered_objects.append(obj)
                self.ui.original_objects.append({
                    'object': obj,
                    'layer': obj.Layer,
                    'type': obj.ObjectName
                })
            
            # 获取地库线范围（如果存在）
            basement_bounds = None
            if hasattr(self, 'ui') and hasattr(self.ui, 'basement_bounds'):
                basement_bounds = self.ui.basement_bounds

            # 在绘制序号时使用选定的标注图层
            current_layer = self.doc.ActiveLayer
            self.doc.ActiveLayer = self.doc.Layers.Item(annotation_layer_name)
            
            # 处理过滤后的对象
            unclosed_count = 0
            for obj in filtered_objects:
                area = 0
                center_point = None
                
                try:
                    # 处理多段线
                    if obj.ObjectName == "AcDbPolyline":
                        if not obj.Closed:
                            # 检查是否有面积
                            if obj.Area > 1:
                                unclosed_count += 1
                                area = obj.Area
                                center = self.calculate_center(obj.Coordinates)
                                # 转换中心点坐标到UCS
                                center_point = self.transform_point([center[0], center[1]])
                                polyline_averge_xy.append(center_point)
                                direct_fontsize.append(math.sqrt(area) / 20)
                            else:
                                continue
                        else:
                            if obj.Area < 1:
                                continue
                            area = obj.Area
                            center = self.calculate_center(obj.Coordinates)
                            # 转换中心点坐标到UCS
                            center_point = self.transform_point([center[0], center[1]])
                            polyline_averge_xy.append(center_point)
                            direct_fontsize.append(math.sqrt(area) / 20)
                    
                    # 处理圆
                    elif obj.ObjectName == "AcDbCircle":
                        area = 3.14159 * obj.Radius * obj.Radius
                        center = obj.Center
                        # 转换中心点坐标到UCS
                        center_point = self.transform_point([center[0], center[1]])
                        polyline_averge_xy.append(center_point)
                        # 计算字体大小
                        direct_fontsize.append(obj.Radius / 5)
                    
                    # 处理椭圆
                    elif obj.ObjectName == "AcDbEllipse":
                        major_axis = obj.MajorAxis
                        major_radius = math.sqrt(major_axis[0]**2 + major_axis[1]**2)
                        minor_radius = major_radius * obj.RadiusRatio
                        area = 3.14159 * major_radius * minor_radius
                        
                        center = obj.Center
                        # 转换中心点坐标到UCS
                        center_point = self.transform_point([center[0], center[1]])
                        polyline_averge_xy.append(center_point)
                        # 计算字体大小
                        direct_fontsize.append(min(major_radius, minor_radius) / 5)
                    
                    # 如果找到有效图形，添加到面积集合
                    if area > 0 and center_point is not None:
                        area_set.append([obj.ObjectID, area])
                        count += 1

                except Exception as e:
                    print(f"处理图形时出错: {str(e)}")
                    continue
            
            # 如果有未闭合的多段线，显示提示
            if unclosed_count > 0:
                self.doc.Utility.Prompt(f"\n注意：发现{unclosed_count}条未闭合的多段线，但仍计入面积计算。")
            
            # 进行标注
            if count > 0:
                # 对中心点进行排序
                sorted_points = self.sort_points(polyline_averge_xy)
                
                # 创建映射关系
                point_to_index = {}
                for i, point in enumerate(polyline_averge_xy):
                    point_to_index[tuple(point)] = i
                
                # 获取排序后的索引
                sorted_indices = [point_to_index[tuple(p)] for p in sorted_points]
                
                # 获取标注类型
                mark_type = "标记"  # 默认值
                if hasattr(self, 'ui') and hasattr(self.ui, 'mark_type_var'):
                    mark_type = self.ui.mark_type_var.get()
                
                # 按标注顺序重新排列面积和中心点
                sorted_areas = [area_set[i][1] for i in sorted_indices]
                sorted_centers = [polyline_averge_xy[i] for i in sorted_indices]
                
                # 根据不同标注类型进行标注
                for i, (center, area) in enumerate(zip(sorted_centers, sorted_areas), 1):
                    if mark_type == "标记":
                        # 只绘制圆圈和序号
                        self.draw_circle_number(center, i)
                    elif mark_type == "数字":
                        # 只标注面积数值
                        self.draw_area_text(center, area)
                    elif mark_type == "综合":
                        # 先绘制面积数值（在圆的左边），再绘制圆圈序号
                        self.draw_area_text(center, area, is_combined=True)
                        self.draw_circle_number(center, i)
                
                object_select.Delete()
                return sorted_areas, sorted_centers
                               
            object_select.Delete()
            return [], []
            
        except Exception as e:
            print(f"绘制过程出错: {str(e)}")
            return [], []

    def calculate_center(self, coords):
        """计算多段线的中心点"""
        x_coords = []
        y_coords = []
        for i in range(0, len(coords), 2):
            x_coords.append(coords[i])
            y_coords.append(coords[i + 1])
        return [sum(x_coords) / len(x_coords), sum(y_coords) / len(y_coords)]

    def get_ucs_rotation(self):
        """计算UCS相对于WCS的旋转角度"""
        if not self.ucs_matrix:
            return 0
        
        try:
            # 计算UCS X轴与WCS X轴的夹角
            x_axis = self.ucs_matrix['xaxis']
            angle = math.atan2(x_axis[1], x_axis[0])  # 返回弧度
            return angle
        except Exception as e:
            print(f"计算UCS旋转角度时出错: {str(e)}")
            return 0

    def draw_circle_number(self, point, number):
        """绘制带序号的圆圈"""
        try:
            # 获取字高
            text_height = 3.0
            if hasattr(self, 'ui'):
                try:
                    text_height = float(self.ui.text_height_var.get())
                    if self.ui.unit_var.get() == "米":
                        text_height = text_height
                except ValueError:
                    text_height = 3.0
            
            # 确保字高不小于最小值
            MIN_TEXT_HEIGHT = 2.5
            if text_height < MIN_TEXT_HEIGHT:
                text_height = MIN_TEXT_HEIGHT
            
            # 将UCS坐标转换回WCS用于绘制
            wcs_point = self.transform_point_to_wcs(point)
            center = self.vtpnt(wcs_point[0], wcs_point[1])
            radius = max(text_height * 0.6, 1.5)
            
            # 获取UCS旋转角度（弧度）
            rotation_angle = self.get_ucs_rotation()
            
            # 获取当前选择的标注图层名称
            annotation_layer_name = None
            if hasattr(self, 'ui') and hasattr(self.ui, 'annotation_layer_var'):
                annotation_layer_name = self.ui.annotation_layer_var.get()
            else:
                annotation_layer_name = "绿化面积标注"  # 默认图层名称
            
            # 设置当前图层
            current_layer = self.doc.ActiveLayer
            self.doc.ActiveLayer = self.doc.Layers.Item(annotation_layer_name)
            
            try:
                # 创建圆
                circle = self.msp.AddCircle(center, radius)
                circle.Color = 1
                
                # 创建文字并根据UCS旋转角度调整
                text = self.msp.AddText(str(number), center, text_height)
                text.Color = 1
                text.Alignment = 4  # 中心对齐
                text.TextAlignmentPoint = center
                text.Rotation = -rotation_angle * 180 / math.pi
                
                # 刷新显示
                self.doc.Regen(1)
                
                return True
            finally:
                self.doc.ActiveLayer = current_layer
                
        except Exception as e:
            print(f"绘制序号圆圈时出错: {str(e)}")
            try:
                if 'circle' in locals():
                    circle.Delete()
                if 'text' in locals():
                    text.Delete()
            except:
                pass
            return False

    def sort_points(self, points):
        """按从上到下，从左到右排序点"""
        if not points:
            return []
            
        # 计算所有点的Y坐标范围
        y_coords = [p[1] for p in points]
        y_min, y_max = min(y_coords), max(y_coords)
        y_range = y_max - y_min
        
        # 计算所有点的X坐标范围
        x_coords = [p[0] for p in points]
        x_min, x_max = min(x_coords), max(x_coords)
        x_range = x_max - x_min
        
        # 根据图形分布动态调整容差
        y_tolerance = y_range / 10  # 将Y轴范围分成10个区域
        y_tolerance = max(min(y_tolerance, y_range / 3), 100)  # 设置合理的容差范围
        
        # 将点按y坐标分组
        rows = {}
        for point in points:
            y = point[1]
            assigned = False
            
            # 查找最近的行，优先使用已存在的行
            closest_row = None
            min_distance = float('inf')
            for existing_y in rows.keys():
                distance = abs(y - existing_y)
                if distance < min_distance and distance < y_tolerance:
                    min_distance = distance
                    closest_row = existing_y
                    
            if closest_row is not None:
                rows[closest_row].append(point)
            else:
                # 创建新行
                rows[y] = [point]
        
        # 对每一行内的点按x坐标排序，并确保行与行之间有足够的距离
        sorted_points = []
        sorted_y = sorted(rows.keys(), reverse=True)  # 从上到下
        
        for y in sorted_y:
            # 对当前行的点按x坐标从左到右排序
            row_points = sorted(rows[y], key=lambda p: p[0])
            sorted_points.extend(row_points)
        
        return sorted_points 

    def check_basement_overlap(self, obj, basement_bounds):
        """检查对象是否与地库线重叠并计算折算系数"""
        try:
            # 获取对象的边界框
            obj_bounds = obj.GetBoundingBox()
            obj_min = [obj_bounds[0][0], obj_bounds[0][1]]
            obj_max = [obj_bounds[1][0], obj_bounds[1][1]]

            # 检查是否与地库线范围重叠
            if (obj_min[0] <= basement_bounds[1][0] and obj_max[0] >= basement_bounds[0][0] and
                obj_min[1] <= basement_bounds[1][1] and obj_max[1] >= basement_bounds[0][1]):
                # 如果重叠，返回配置的折算系数（可以从UI获取）
                if hasattr(self, 'ui') and hasattr(self.ui, 'basement_factor'):
                    return float(self.ui.basement_factor)
                return 0.5  # 默认折算系数
            
            return 1.0  # 如果不重叠，返回1.0（不进行折算）
        except Exception as e:
            print(f"检查地库线重叠时出错: {str(e)}")
            return 1.0 

    def get_hatch_patterns(self):
        """获取CAD中所有可用的填充样式"""
        try:
            # 定义常用填充样式
            common_patterns = ["CROSS", "GRASS", "ANSI31", "ANSI32", "ANSI33", "ANSI37", "AR-CONC", "AR-SAND"]
            
            # 获取所有预定义的填充样式
            patterns = []
            for i in range(self.doc.HatchPatterns.Count):
                pattern = self.doc.HatchPatterns.Item(i)
                patterns.append(pattern.Name)
            
            # 合并常用样式和系统样式，去重
            all_patterns = list(set(common_patterns + patterns))
            # 确保常用样式在列表前面
            for pattern in reversed(common_patterns):
                if pattern in all_patterns:
                    all_patterns.remove(pattern)
                    all_patterns.insert(0, pattern)
                
            return all_patterns
        except Exception as e:
            print(f"获取填充样式时出错: {str(e)}")
            return ["CROSS", "GRASS", "ANSI31"]  # 返回基本默认样式

    def apply_hatch(self, pattern, scale):
        """应用填充到已标注的对象"""
        try:
            if not hasattr(self, 'ui') or not hasattr(self.ui, 'original_objects'):
                self.doc.Utility.Prompt("请先进行面积标注！")
                return

            # 获取填充参数
            angle = float(self.ui.hatch_angle_var.get())
            
            # 获取颜色值
            color_name = self.ui.hatch_color_var.get()
            color = self.ui.color_map.get(color_name, self.ui.color_map["默认"])

            # 记录未闭合的多段线数量
            unclosed_count = 0

            try:
                # 创建填充
                [PatternType, patternName, bAss] = [1, pattern or "CROSS", True]  # 设置默认值为CROSS
                hatch = self.msp.AddHatch(PatternType, patternName, bAss)
                
                # 遍历所有原始对象
                for obj_info in self.ui.original_objects:
                    try:
                        obj = obj_info['object']
                        if obj_info['type'] not in ["AcDbPolyline", "AcDbCircle", "AcDbEllipse"]:
                            continue
                            
                        # 检查是否是未闭合的多段线
                        if obj_info['type'] == "AcDbPolyline" and not obj.Closed:
                            unclosed_count += 1
                            
                        # 获取对象中心点
                        center = None
                        if obj_info['type'] == "AcDbPolyline":
                            coords = obj.Coordinates
                            x_sum = sum(coords[i] for i in range(0, len(coords), 2))
                            y_sum = sum(coords[i+1] for i in range(0, len(coords), 2))
                            points_count = len(coords) // 2
                            center = [x_sum/points_count, y_sum/points_count]
                        elif obj_info['type'] == "AcDbCircle":
                            center = list(obj.Center)
                        elif obj_info['type'] == "AcDbEllipse":
                            center = list(obj.Center)

                        # 检查这个对象是否是我们标注的对象
                        if center and self.ui.center_points:
                            for marked_point in self.ui.center_points:
                                if (abs(center[0] - marked_point[0]) < 1 and 
                                    abs(center[1] - marked_point[1]) < 1):
                                    # 创建对象数组并添加边界
                                    outerloop = []
                                    outerloop.append(obj)
                                    outerloop = self.vtobj(outerloop)
                                    try:
                                        hatch.AppendInnerLoop(outerloop)
                                        # 设置填充到原始对象的图层
                                        hatch.Layer = obj_info['layer']
                                    except Exception as e:
                                        print(f"添加填充边界时出错: {str(e)}")
                                    break
                    
                    except Exception as e:
                        print(f"处理对象时出错: {str(e)}")
                        continue

                    # 设置填充属性
                    hatch.PatternAngle = math.radians(angle)
                    hatch.PatternScale = scale
                    hatch.Color = color

                    # 计算填充
                    try:
                        hatch.Evaluate()
                    except:
                        self.doc.Utility.Prompt("填充比例有误，请尝试其他数值\n")
                        if 'hatch' in locals():
                            hatch.Delete()
                        return

                    # 刷新显示
                    self.doc.Regen(1)
                    
                    # 显示未闭合多段线的提示
                    if unclosed_count > 0:
                        self.doc.Utility.Prompt(f"\n注意：发现{unclosed_count}条未闭合的多段线\n")
                    self.doc.Utility.Prompt("已完成填充\n")

            except Exception as e:
                print(f"填充过程出错: {str(e)}")
                if 'hatch' in locals():
                    try:
                        hatch.Delete()
                    except:
                        pass

        except Exception as e:
            print(f"填充过程出错: {str(e)}") 

    def draw_area_text(self, point, area, offset_y=0, offset_x=0, is_combined=False):
        """绘制面积数值"""
        try:
            # 获取字高
            text_height = 3.0
            if hasattr(self, 'ui'):
                try:
                    text_height = float(self.ui.text_height_var.get())
                except ValueError:
                    text_height = 3.0
            
            # 将UCS坐标转换回WCS用于绘制
            wcs_point = self.transform_point_to_wcs(point)
            
            # 处理单位转换和数值格式化
            if hasattr(self, 'ui') and self.ui.unit_var.get() == "米":
                # 将平方毫米转换为平方米（除以1,000,000）
                converted_area = area / 1000000
                area_text = f"{converted_area:.2f}㎡"
            else:
                area_text = f"{area:.2f}㎟"
            
            # 如果是综合模式，调整文本位置到圆的左侧
            if is_combined:
                # 计算圆的半径（与draw_circle_number中保持一致）
                radius = max(text_height * 0.6, 1.5)
                # 将文本位置向左偏移圆的直径加上一些间距
                offset_x =radius *2 + text_height
            
            # 添加偏移
            text_point = self.vtpnt(
                wcs_point[0] + offset_x,
                wcs_point[1] + offset_y
            )
            
            # 获取UCS旋转角度（弧度）
            rotation_angle = self.get_ucs_rotation()
            
            # 创建文字
            text = self.msp.AddText(area_text, text_point, text_height)
            text.Color = 1
            
            # 如果是综合模式，使用右对齐
            if is_combined:
                text.Alignment = 4  # 右对齐
            else:
                text.Alignment = 4  # 中心对齐
            
            text.TextAlignmentPoint = text_point
            text.Rotation = -rotation_angle * 180 / math.pi
            
            return True
        except Exception as e:
            print(f"绘制面积文本时出错: {str(e)}")
            if 'text' in locals():
                try:
                    text.Delete()
                except:
                    pass
            return False 

    def get_current_drawing_name(self):
        """获取当前CAD图纸的名称"""
        try:
            if self.doc:
                full_path = self.doc.FullName
                if full_path:
                    return os.path.basename(full_path)
        except:
            pass
        return "" 