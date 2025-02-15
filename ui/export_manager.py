import win32com.client
import os
from tkinter import messagebox, filedialog
import time
from utils.settings_manager import SettingsManager
import tkinter as tk
import pythoncom
import traceback
from win32com.client import VARIANT
from utils.wps_path_finder import WPSPathFinder

class ExportManager:
    def __init__(self):
        self.cad_name = ""
        self.settings_manager = SettingsManager()
        self.cad = None  # 添加 CAD 实例变量
        
    def set_cad_name(self, name):
        """设置CAD文件名"""
        self.cad_name = name.replace(".dwg", "") if name else ""
        
    def get_cad_name(self):
        """获取CAD文件名，如果当前没有则从设置中获取"""
        if not self.cad_name:
            settings = self.settings_manager.load_settings()
            self.cad_name = settings.get("cad_filename", "").replace(".dwg", "") or "未命名"
        return self.cad_name
        
    def set_cad_instance(self, cad_instance):
        """设置 CAD 实例"""
        self.cad = cad_instance
        
    def export_to_word(self, data, summary_data, unit_symbol):
        """导出到Word"""
        word = None
        doc = None
        try:
            # 获取主窗口并最小化
            root = tk._default_root
            if root:
                root.iconify()
            
            # 获取保存路径
            default_filename = f"{self.get_cad_name()}绿地率计算表.docx"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word文档", "*.docx")],
                title="保存Word文档",
                initialfile=default_filename
            )
            if not file_path:
                return
                
            # 确保文件路径是绝对路径
            file_path = os.path.abspath(file_path)
            
            # 检查文件是否被占用
            try:
                with open(file_path, 'a'):
                    pass
            except PermissionError:
                messagebox.showerror("错误", "文件被占用，请关闭已打开的文件后重试")
                return
            except:
                pass
            
            # 创建Word应用实例
            word = win32com.client.DispatchEx('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False
            
            # 创建新文档
            doc = word.Documents.Add()
            
            # 添加标题             
            selection = word.Selection
            selection.Font.Size = 22  
            selection.Font.Bold = True  
            
            # 添加主标题
            selection.TypeText("绿地面积统计表")
            selection.Font.Size = 22
            selection.Font.Bold = True  
            selection.ParagraphFormat.Alignment = 1
            selection.TypeParagraph()
            
            # 计算表格行数和列数
            rows = len(data) + len(summary_data) + 1
            cols = 4
            
            # 创建表格
            table = doc.Tables.Add(selection.Range, rows, cols)
            table.Borders.Enable = True
            
            # 设置表格宽度和列宽
            table.PreferredWidth = 450
            col_widths = [50, 130, 100, 130]
            for i, width in enumerate(col_widths):
                table.Columns(i + 1).Width = width
            
            # 设置表头行在每页重复显示
            table.Rows(1).HeadingFormat = True
            
            # 设置表头
            headers = [
                "序号", 
                f"实测面积({unit_symbol})", 
                "折算系数", 
                f"折算面积({unit_symbol})"
            ]
            for i, header in enumerate(headers):
                cell = table.Cell(1, i + 1)
                cell.Range.Text = header
                cell.Range.Font.Bold = True
                cell.Range.Font.Size = 12  # 小四号字体为12磅
                cell.Range.ParagraphFormat.Alignment = 1
            
            # 填充数据
            for row, item in enumerate(data, 2):
                # 计算折算面积
                actual_area = float(item['actual_area'])
                factor = float(item['factor'].replace('%', '')) / 100.0
                converted_area = round(actual_area * factor, 2)
                
                cells = [
                    str(row - 1),
                    item['actual_area'],
                    item['factor'],
                    f"{converted_area:.2f}"  # 格式化为两位小数
                ]
                for col, text in enumerate(cells, 1):
                    cell = table.Cell(row, col)
                    cell.Range.Text = text
                    cell.Range.Font.Bold = False
                    cell.Range.Font.Size = 12  # 小四号字体为12磅
                    cell.Range.ParagraphFormat.Alignment = 1
            
            # 添加汇总数据
            start_row = len(data) + 2
            for i, summary in enumerate(summary_data):
                row = start_row + i
                # 合并整行
                first_cell = table.Cell(row, 1)
                last_cell = table.Cell(row, cols)
                first_cell.Merge(last_cell)
                
                # 设置文本和格式
                first_cell.Range.Text = summary['merged_content']
                first_cell.Range.Font.Bold = False
                first_cell.Range.Font.Size = 12  # 小四号字体为12磅
                first_cell.Range.ParagraphFormat.Alignment = 1
                
            # 修改保存和关闭逻辑
            try:
                doc.SaveAs(file_path)
            except Exception as save_error:
                messagebox.showerror("保存错误", f"保存文件时出错: {str(save_error)}")
                raise save_error
            
            try:
                doc.Close(SaveChanges=False)
                doc = None
            except:
                pass
            
            try:
                word.Quit()
                word = None
            except:
                pass

            # 等待文件可用
            max_attempts = 10
            attempt = 0
            while attempt < max_attempts:
                try:
                    with open(file_path, 'r'):
                        break
                except:
                    time.sleep(0.5)
                    attempt += 1

            # 打开文件
            os.startfile(file_path)
            
        except Exception as e:
            error_msg = f"导出到Word失败：{str(e)}"
            messagebox.showerror("错误", error_msg)
        finally:
            # 确保资源被释放
            try:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except:
                        pass
            except:
                pass
            
            try:
                if word is not None:
                    try:
                        word.Quit()
                    except:
                        pass
            except:
                pass
            
            # 恢复窗口
            if root:
                root.deiconify()

    def export_to_excel(self, data, summary_data, unit_symbol):
        """导出到Excel"""
        excel = None
        wb = None
        try:
            # 获取主窗口并最小化
            root = tk._default_root
            if root:
                root.iconify()
            
            # 获取保存路径
            default_filename = f"{self.get_cad_name()}绿地率计算表.xlsx"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel工作表", "*.xlsx")],
                title="保存Excel文件",
                initialfile=default_filename
            )
            if not file_path:
                return
                
            # 确保文件路径是绝对路径
            file_path = os.path.abspath(file_path)
            
            # 创建Excel应用实例
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # 创建新工作簿
            wb = excel.Workbooks.Add()
            ws = wb.ActiveSheet
            
            try:
                # 添加标题
                ws.Range("A1:D1").Merge()
                ws.Range("A1").Value = "绿地面积统计表"
                ws.Range("A1").Font.Size = 14
                ws.Range("A1").Font.Bold = True
                ws.Range("A1").HorizontalAlignment = -4108  # 居中
                
                # 添加表头
                headers = [
                    "序号", 
                    f"实测面积({unit_symbol})", 
                    "折算系数", 
                    f"折算面积({unit_symbol})"
                ]
                for i, header in enumerate(headers):
                    cell = ws.Cells(2, i + 1)
                    cell.Value = header
                    cell.HorizontalAlignment = -4108  # 居中
                
                # 填充数据
                for row, item in enumerate(data, 3):
                    # 计算折算面积
                    actual_area = float(item['actual_area'])
                    factor = float(item['factor'].replace('%', '')) / 100.0
                    converted_area = round(actual_area * factor, 2)
                    
                    ws.Cells(row, 1).Value = row - 2
                    ws.Cells(row, 2).Value = item['actual_area']
                    ws.Cells(row, 3).Value = item['factor']
                    ws.Cells(row, 4).Value = f"{converted_area:.2f}"  # 格式化为两位小数
                    
                    # 设置居中对齐
                    for col in range(1, 5):
                        cell = ws.Cells(row, col)
                        cell.HorizontalAlignment = -4108  # 居中
                
                # 添加汇总数据
                start_row = len(data) + 3
                for i, summary in enumerate(summary_data):
                    row = start_row + i
                    # 合并整行(A-D列)
                    merge_range = ws.Range(
                        ws.Cells(row, 1),
                        ws.Cells(row, 4)
                    )
                    merge_range.Merge()
                    merge_range.Value = summary['merged_content']
                    merge_range.HorizontalAlignment = -4108  # 居中
                
                # 设置表格边框
                table_range = ws.Range(
                    ws.Cells(1, 1),
                    ws.Cells(start_row + len(summary_data) - 1, 4)
                )
                for border_id in range(7, 13):
                    table_range.Borders(border_id).LineStyle = 1
                    table_range.Borders(border_id).Weight = 2
                
                # 设置D列宽度为13.1
                ws.Columns("D").ColumnWidth = 13.1
                
                # 调整其他列宽和行高
                ws.Columns("A:C").AutoFit()
                ws.Rows.AutoFit()
                
                # 保存并关闭
                wb.SaveAs(file_path)
                wb.Close()
                excel.Quit()
                
                # 等待Excel完全关闭后再打开文件
                time.sleep(1)
                
                # 打开Excel文件查看
                os.system(f'start excel.exe "{file_path}"')
                
            except Exception as e:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        pass
                raise e
                
        except Exception as e:
            messagebox.showerror("错误", f"导出到Excel失败：{str(e)}")
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except:
                    pass
            # 恢复窗口
            if root:
                root.deiconify()

    def export_to_ppt(self, data, summary_data, unit_symbol):
        """导出到PowerPoint"""
        ppt = None
        presentation = None
        try:
            # 获取主窗口并最小化
            root = tk._default_root
            if root:
                root.iconify()
            
            # 获取保存路径
            default_filename = f"{self.get_cad_name()}绿地率计算表.pptx"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pptx",
                filetypes=[("PowerPoint演示文稿", "*.pptx")],
                title="保存PowerPoint文件",
                initialfile=default_filename
            )
            if not file_path:
                return
                
            # 确保文件路径是绝对路径
            file_path = os.path.abspath(file_path)
            
            # 创建PowerPoint应用实例
            ppt = win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
            ppt.Visible = True
            
            # 创建新演示文稿
            presentation = ppt.Presentations.Add(WithWindow=True)
            
            # 添加幻灯片
            slide = presentation.Slides.Add(1, 11)  # 11 = ppLayoutText
            
            # 设置标题
            title_shape = slide.Shapes.Title
            title_shape.TextFrame.TextRange.Text = "绿地面积统计表"
            
            # 根据数据量决定布局
            total_rows = len(data)
            if total_rows <= 20:
                # 单列布局
                cols = 1
                table_width = 600
                left_margin = 50
            else:
                # 双列布局
                cols = 2
                table_width = 380
                left_margin = 30
            
            # 计算每列的行数
            rows_per_col = (total_rows + cols - 1) // cols
            
            for col in range(cols):
                # 计算当前列的数据范围
                start_idx = col * rows_per_col
                end_idx = min((col + 1) * rows_per_col, total_rows)
                current_data = data[start_idx:end_idx]
                
                if not current_data:
                    continue
                
                # 创建表格（加1是为了表头）
                rows = len(current_data) + 1
                table_cols = 4
                
                # 计算表格位置
                left = left_margin + (table_width + 20) * col
                top = 120
                height = min(480, rows * 20)  # 限制最大高度
                
                table = slide.Shapes.AddTable(rows, table_cols, left, top, table_width, height).Table
                
                # 设置表头
                headers = ["序号", f"实测面积({unit_symbol})", "折算系数", f"折算面积({unit_symbol})"]
                
                # 设置列宽比例
                col_widths = [0.15, 0.3, 0.25, 0.3]
                for i, width_ratio in enumerate(col_widths):
                    table.Columns(i + 1).Width = table_width * width_ratio
                
                # 添加表头
                for i, header in enumerate(headers):
                    cell = table.Cell(1, i + 1)
                    cell.Shape.TextFrame.TextRange.Text = header
                    # 根据数据量调整字体大小
                    font_size = 12 if total_rows <= 30 else 10
                    cell.Shape.TextFrame.TextRange.Font.Size = font_size
                    cell.Shape.TextFrame.TextRange.Font.Bold = True
                    cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                
                # 填充数据
                for row, item in enumerate(current_data, 2):
                    # 计算折算面积
                    actual_area = float(item['actual_area'])
                    factor = float(item['factor'].replace('%', '')) / 100.0
                    converted_area = round(actual_area * factor, 2)
                    
                    # 序号要连续
                    table.Cell(row, 1).Shape.TextFrame.TextRange.Text = str(start_idx + row - 1)
                    table.Cell(row, 2).Shape.TextFrame.TextRange.Text = item['actual_area']
                    table.Cell(row, 3).Shape.TextFrame.TextRange.Text = item['factor']
                    table.Cell(row, 4).Shape.TextFrame.TextRange.Text = f"{converted_area:.2f}"  # 格式化为两位小数
                    
                    # 设置数据格式
                    for col_idx in range(1, 5):
                        cell = table.Cell(row, col_idx)
                        cell.Shape.TextFrame.TextRange.Font.Size = font_size - 1
                        cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            
            # 添加汇总信息
            if summary_data:
                # 计算汇总表格的位置
                summary_left = left_margin
                summary_top = top + height + 20
                summary_width = 600
                summary_height = min(100, len(summary_data) * 25)
                
                # 创建汇总表格
                sum_table = slide.Shapes.AddTable(
                    len(summary_data), 1,
                    summary_left, summary_top,
                    summary_width, summary_height
                ).Table
                
                # 填充汇总数据
                for i, summary in enumerate(summary_data):
                    cell = sum_table.Cell(i + 1, 1)
                    cell.Shape.TextFrame.TextRange.Text = summary['merged_content']
                    cell.Shape.TextFrame.TextRange.Font.Size = 12
                    cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            
            # 保存并关闭
            presentation.SaveAs(file_path)
            presentation.Close()
            ppt.Quit()
            
            # 等待PowerPoint完全关闭后再打开文件
            time.sleep(1)
            os.startfile(file_path)
            
        except Exception as e:
            messagebox.showerror("错误", f"导出到PowerPoint失败：{str(e)}")
        finally:
            try:
                if presentation is not None:
                    presentation.Close()
            except:
                pass
            try:
                if ppt is not None:
                    ppt.Quit()
            except:
                pass
            # 恢复窗口
            if root:
                root.deiconify()


    def export_to_cad(self, data, summary_data, unit_symbol):
        """导出表格到CAD"""
        try:
            # 尝试获取已在运行的 AutoCAD 实例
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
            doc = acad.ActiveDocument
            msp = doc.ModelSpace

            # ********************* 调试信息输出 *********************
            print("--- 开始执行 export_to_cad ---")
            print("  AutoCAD 连接成功")

            # ********************* 获取插入点 *********************
            try:
                prompt_message = "请在 CAD 中选择主数据表格插入点:"
                doc.Utility.Prompt(prompt_message)
                point = doc.Utility.GetPoint()
                insert_point_tuple = point #  **重要修改： 保存 point 为元组**
                insert_point = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(point[0]), float(point[1]), float(point[2])))
                print(f"用户选择的主数据表格插入点: {point}")
            except pythoncom.com_error as e:
                if e.hresult == -2147352565:
                    messagebox.showinfo("提示", "用户取消了点选择操作，表格导出已中止。")
                    print("用户取消了点选择操作，表格导出已中止。")
                    return #  用户取消，提前退出函数
                else:
                    raise #  其他 COM 错误，继续抛出

            # ********************* 创建主数据表格 *********************
            print("  --- 创建主数据表格 ---")
            # 主数据表格参数
            data_table_rows = len(data) + 2
            data_table_cols = 4
            row_height = 0.6   # 从100改为30
            col_width = 2.4   # 从400改为120

            print("  准备创建主数据表格，参数:")
            print(f"    插入点 (insert_point): {insert_point}")
            print(f"    行数 (data_table_rows): {data_table_rows}")
            print(f"    列数 (data_table_cols): {data_table_cols}")
            print(f"    行高 (row_height): {row_height}")
            print(f"    列宽 (col_width): {col_width}")

            # 创建主数据表格对象
            try:
                data_table = msp.AddTable(insert_point, data_table_rows, data_table_cols, row_height, col_width)
                print("  主数据表格创建命令 AddTable 执行完成，返回值:", data_table)
            except Exception as add_data_table_error:
                print("  **主数据表格创建命令 AddTable 发生错误:**")
                print(traceback.format_exc())
                raise add_data_table_error

            if not data_table:
                raise Exception("CAD 主数据表格对象创建失败，AddTable 返回 None")
            else:
                print("  主数据表格对象创建成功")

            if data_table is None:
                raise Exception("主数据表格对象创建失败 (data_table 为 None)， 无法继续设置标题和表头")

            # ********************* 设置主数据表格文字高度  *********************
            try:
                data_table.TextStyle.TextHeight = 350  # 从50改为350
                print("  主数据表格文字高度已设置为 350")
            except Exception as set_data_table_text_height_error:
                print("  **设置主数据表格文字高度 (TextStyle.TextHeight) 发生错误:**")
                print(traceback.format_exc())

            # 添加主数据表格标题
            try:
                data_table.SetText(0, 0, "绿地面积统计表")
                print("  主数据表格标题已设置为 '绿地面积统计表' (无 MergeCells)")
            except Exception as set_title_error:
                print("  **设置主数据表格标题 (SetText) 发生错误:**")
                print(traceback.format_exc())
                raise set_title_error

            # 添加主数据表格表头
            try:
                headers = [
                    "序号",
                    f"实测面积(㎡)", # 使用 self.unit_symbol
                    "折算系数",
                    f"折算面积(㎡)" # 使用 self.unit_symbol
                ]
                if len(headers) != data_table_cols:
                    raise ValueError("表头列表 'headers' 的长度必须与表格列数 {} 相同".format(data_table_cols))

                print("  准备设置主数据表格表头，表头内容:", headers)
                for col, header in enumerate(headers):
                    data_table.SetText(1, col, header)
                    print(f"    已设置主数据表格表头单元格 (行: 1, 列: {col}), 内容: '{header}'")
                print("  主数据表格表头已设置 (SetText)")

            except Exception as set_headers_error:
                print("  **设置主数据表格表头 (SetText 循环) 发生错误:**")
                print(traceback.format_exc())
                raise set_headers_error

            # 添加主数据表格数据行 - 计算折算面积和总面积
            total_actual_area = 0.0
            total_converted_area = 0.0
            try:
                print("  准备设置主数据表格数据行，数据行数:", len(data))
                for row_index, item in enumerate(data):
                    row = row_index + 2
                    if row >= data_table_rows:
                        break

                    actual_area_str = item.get('actual_area', '')
                    factor_str = item.get('factor', '')

                    actual_area = 0.0
                    factor = 0.0

                    try:
                        actual_area = float(actual_area_str)
                        total_actual_area += actual_area
                    except ValueError:
                        print(f"  警告: 行 {row}, 序号 {row_index + 1}, 实测面积 '{actual_area_str}' 无法转换为数字，使用默认值 0.0")

                    try:
                        factor = float(factor_str.replace('%', '')) / 100.0
                    except ValueError:
                        print(f"  警告: 行 {row}, 序号 {row_index + 1}, 折算系数 '{factor_str}' 无法转换为百分比，使用默认值 0.0")

                    converted_area = actual_area * factor
                    total_converted_area += converted_area
                    converted_area_str = "{:.2f}".format(converted_area)

                    data_table.SetText(row, 0, str(row_index + 1))
                    data_table.SetText(row, 1, actual_area_str)
                    data_table.SetText(row, 2, factor_str)
                    data_table.SetText(row, 3, converted_area_str)

                    print(f"    已设置主数据表格数据行 (行: {row}), 内容: {item}, 计算折算面积: {converted_area_str}, 累计总实测面积: {total_actual_area:.2f}, 累计总折算面积: {total_converted_area:.2f}")

                print("  主数据表格数据行已设置 (SetText 循环), 折算面积已自动计算, 总面积已累计")

            except Exception as set_data_rows_error:
                print("  **设置主数据表格数据行 (SetText 循环) 发生错误:**")
                print(traceback.format_exc())
                raise set_data_rows_error


            # ********************* 创建单列汇总表格 *********************
            print("  --- 创建单列汇总表格 ---")
            # 汇总表格参数
            summary_table_rows = len(summary_data)
            summary_table_cols = 1
            summary_table_col_width = col_width * 4

            # 计算汇总表格插入点 - 放在主数据表格下方
            # **重要修改： 使用 insert_point_tuple (元组) 进行计算**
            data_table_bottom_y = insert_point_tuple[1] - (data_table_rows * row_height)
            summary_insert_point_y = data_table_bottom_y - (row_height * 2)
            # **重要修改： 使用 insert_point_tuple 的 X 和 Z 坐标， 以及计算出的 summary_insert_point_y**
            summary_insert_point = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(insert_point_tuple[0]), float(summary_insert_point_y), float(insert_point_tuple[2])))

            print("  准备创建单列汇总表格，参数:")
            print(f"    插入点 (summary_insert_point): {summary_insert_point}")
            print(f"    行数 (summary_table_rows): {summary_table_rows}")
            print(f"    列数 (summary_table_cols): {summary_table_cols}")
            print(f"    行高 (row_height): {row_height}")
            print(f"    列宽 (summary_table_col_width): {summary_table_col_width}")

            # *********************  调试：创建汇总表格前  *********************
            print("  **调试信息: 准备调用 AddTable 创建单列汇总表格...**")

            # 创建单列汇总表格对象
            try:
                summary_table = msp.AddTable(summary_insert_point, summary_table_rows, summary_table_cols, row_height, summary_table_col_width)
                print("  单列汇总表格创建命令 AddTable 执行完成，返回值:", summary_table)
            except Exception as add_summary_table_error:
                print("  **单列汇总表格创建命令 AddTable 发生错误:**")
                print(traceback.format_exc())
                raise add_summary_table_error

            if not summary_table:
                raise Exception("CAD 单列汇总表格对象创建失败，AddTable 返回 None")
            else:
                print("  单列汇总表格对象创建成功")

            # *********************  设置单列汇总表格文字高度  *********************
            try:
                summary_table.TextStyle.TextHeight = 350  # 从50改为350
                print("  单列汇总表格文字高度已设置为 350")
            except Exception as set_summary_table_text_height_error:
                print("  **设置单列汇总表格文字高度 (TextStyle.TextHeight) 发生错误:**")
                print(traceback.format_exc())


            # ********************* 添加单列汇总表格数据行 *********************
            try:
                print("  准备设置单列汇总表格数据行，汇总行数:", len(summary_data))
                for row_index, summary_item in enumerate(summary_data):
                    row = row_index
                    if row >= summary_table_rows:
                        break
                    merged_content = summary_item.get('merged_content', '')

                    summary_table.SetText(row, 0, merged_content)

                    print(f"    已设置单列汇总表格数据行 (行: {row}), 内容: '{merged_content}' (无 MergeCells)")
                print("  单列汇总表格数据行已设置 (SetText 循环, 无 MergeCells)")

            except Exception as set_summary_rows_error:
                print("  **设置单列汇总表格数据行 (SetText 循环) 发生错误:**")
                print(traceback.format_exc())
                raise set_summary_rows_error

            # *********************  调试：添加汇总表格数据行后，刷新前  *********************
            print("  **调试信息: 单列汇总表格数据行设置完成，准备刷新视图...**")


            # ********************* 刷新视图和显示成功消息 *********************
            doc.Regen(True)
            # messagebox.showinfo("成功", "已创建主数据表格和单列汇总表格 (均无 MergeCells)，请检查CAD")
            print("--- export_to_cad 执行结束，未发生严重错误 ---")

        except ValueError as ve:
            messagebox.showerror("数据错误", str(ve))
            print(traceback.format_exc())
        except Exception as e:
            error_message = "在创建 CAD 表格或设置数据/汇总行时发生未知错误 (类方法版本, 无 MergeCells).\n详细错误信息: {}".format(e)
            messagebox.showerror("错误", error_message)
            print("--- export_to_cad 执行异常结束 ---")
            print(traceback.format_exc())

    def export_to_wps(self, data, summary_data, unit_symbol):
        """导出到WPS"""
        word = None
        doc = None
        try:
            # 获取主窗口并最小化
            root = tk._default_root
            if root:
                root.iconify()
            
            # 获取保存路径
            default_filename = f"{self.get_cad_name()}绿地率计算表.wps"  # 改为.wps格式
            file_path = filedialog.asksaveasfilename(
                defaultextension=".wps",  # 改为.wps格式
                filetypes=[("WPS文档", "*.wps")],  # 改为.wps格式
                title="保存WPS文档",
                initialfile=default_filename
            )
            if not file_path:
                return
                
            # 确保文件路径是绝对路径
            file_path = os.path.abspath(file_path)
            
            # 检查文件是否被占用
            try:
                with open(file_path, 'a'):
                    pass
            except PermissionError:
                messagebox.showerror("错误", "文件被占用，请关闭已打开的文件后重试")
                return
            except:
                pass
            
            # 创建Word应用实例
            word = win32com.client.DispatchEx('Word.Application')
            word.Visible = False
            word.DisplayAlerts = False
            
            # 创建新文档
            doc = word.Documents.Add()
            
            # 添加标题             
            selection = word.Selection
            selection.Font.Size = 22  
            selection.Font.Bold = True  
            
            # 添加主标题
            selection.TypeText("绿地面积统计表")
            selection.Font.Size = 22
            selection.Font.Bold = True  
            selection.ParagraphFormat.Alignment = 1
            selection.TypeParagraph()
            
            # 计算表格行数和列数
            rows = len(data) + len(summary_data) + 1
            cols = 4
            
            # 创建表格
            table = doc.Tables.Add(selection.Range, rows, cols)
            table.Borders.Enable = True
            
            # 设置表格宽度和列宽
            table.PreferredWidth = 450
            col_widths = [50, 130, 100, 130]
            for i, width in enumerate(col_widths):
                table.Columns(i + 1).Width = width
            
            # 设置表头行在每页重复显示
            table.Rows(1).HeadingFormat = True
            
            # 设置表头
            headers = [
                "序号", 
                f"实测面积({unit_symbol})", 
                "折算系数", 
                f"折算面积({unit_symbol})"
            ]
            for i, header in enumerate(headers):
                cell = table.Cell(1, i + 1)
                cell.Range.Text = header
                cell.Range.Font.Bold = True
                cell.Range.Font.Size = 12
                cell.Range.ParagraphFormat.Alignment = 1
            
            # 填充数据
            for row, item in enumerate(data, 2):
                # 计算折算面积
                actual_area = float(item['actual_area'])
                factor = float(item['factor'].replace('%', '')) / 100.0
                converted_area = round(actual_area * factor, 2)
                
                cells = [
                    str(row - 1),
                    item['actual_area'],
                    item['factor'],
                    f"{converted_area:.2f}"
                ]
                for col, text in enumerate(cells, 1):
                    cell = table.Cell(row, col)
                    cell.Range.Text = text
                    cell.Range.Font.Bold = False
                    cell.Range.Font.Size = 12
                    cell.Range.ParagraphFormat.Alignment = 1
            
            # 添加汇总数据
            start_row = len(data) + 2
            for i, summary in enumerate(summary_data):
                row = start_row + i
                # 合并整行
                first_cell = table.Cell(row, 1)
                last_cell = table.Cell(row, cols)
                first_cell.Merge(last_cell)
                
                # 设置文本和格式
                first_cell.Range.Text = summary['merged_content']
                first_cell.Range.Font.Bold = False
                first_cell.Range.Font.Size = 12
                first_cell.Range.ParagraphFormat.Alignment = 1
            
            # 保存为WPS格式
            try:
                # 使用数字常量12表示WPS格式
                doc.SaveAs(file_path, FileFormat=12)  # 12 对应 WPS 格式
            except Exception as save_error:
                messagebox.showerror("保存错误", f"保存文件时出错: {str(save_error)}")
                raise save_error
            
            # 关闭文档和应用程序
            try:
                doc.Close(SaveChanges=False)
                doc = None
            except:
                pass
            
            try:
                word.Quit()
                word = None
            except:
                pass

            # 等待文件可用
            max_attempts = 10
            attempt = 0
            while attempt < max_attempts:
                try:
                    with open(file_path, 'r'):
                        break
                except:
                    time.sleep(0.5)
                    attempt += 1

            # 使用WPS打开文件
            try:
                wps_path = WPSPathFinder.get_wps_path()  # 使用新的get_wps_path方法
                if wps_path:
                    os.system(f'start "" "{wps_path}" "{file_path}"')
                else:
                    # 如果找不到WPS，使用系统默认程序打开
                    os.startfile(file_path)
            except:
                # 如果出错，使用系统默认程序打开
                os.startfile(file_path)
        
        except Exception as e:
            error_msg = f"导出到WPS失败：{str(e)}"
            messagebox.showerror("错误", error_msg)
            print(f"导出到WPS时发生错误: {str(e)}")
            print(traceback.format_exc())
        finally:
            # 确保资源被释放
            try:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except:
                        pass
            except:
                pass
            
            try:
                if word is not None:
                    try:
                        word.Quit()
                    except:
                        pass
            except:
                pass
            
            # 恢复窗口
            if root:
                root.deiconify()
