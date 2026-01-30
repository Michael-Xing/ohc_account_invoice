"""产品环境评估要项书/结果书填充器"""

from pathlib import Path
from typing import Dict, Any, List
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, PatternFill
from openpyxl.utils import get_column_letter

from src.infrastructure.template_service import ExcelTemplateFiller


class ProductEnvironmentAssessmentFiller(ExcelTemplateFiller):
    """产品环境评估要项书/结果书填充器"""
    
    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        """
        填充产品环境评估要项书/结果书模板
        
        Args:
            template_path: 模板文件路径
            parameters: 填充参数
            output_path: 输出文件路径
            
        Returns:
            bool: 是否成功
        """
        try:
            workbook = load_workbook(template_path)
            worksheet = workbook.active
            
            # 1. 填充简单拼接字段
            self._fill_simple_fields(worksheet, parameters)
            
            # 2. 填充需要分割和合并单元格的字段
            max_row = self._fill_table_data(worksheet, parameters)
            
            # 替换其他占位符
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell.value = self._replace_placeholders(cell.value, parameters)
            
            workbook.save(output_path)
            return True
        except Exception as e:
            print(f"产品环境评估模板填充失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def _fill_simple_fields(self, worksheet, parameters: Dict[str, Any]):
        """填充简单拼接字段"""
        # theme_no 拼接到 B5 单元格内容后面
        if 'theme_no' in parameters and parameters['theme_no']:
            cell_b5 = worksheet['B5']
            original_value = cell_b5.value or ''
            cell_b5.value = str(original_value) + str(parameters['theme_no'])
        
        # product_model_name 拼接到 B7 单元格内容后面
        if 'product_model_name' in parameters and parameters['product_model_name']:
            cell_b7 = worksheet['B7']
            original_value = cell_b7.value or ''
            cell_b7.value = str(original_value) + str(parameters['product_model_name'])
        
        # product_name 拼接到 I5 单元格内容后面
        if 'product_name' in parameters and parameters['product_name']:
            cell_i5 = worksheet['I5']
            original_value = cell_i5.value or ''
            cell_i5.value = str(original_value) + str(parameters['product_name'])
        
        # production_area 拼接到 I6 单元格内容后面
        if 'production_area' in parameters and parameters['production_area']:
            cell_i6 = worksheet['I6']
            original_value = cell_i6.value or ''
            cell_i6.value = str(original_value) + str(parameters['production_area'])
    
    def _fill_table_data(self, worksheet, parameters: Dict[str, Any]) -> int:
        """
        填充表格数据（从22行开始）
        返回填充的最大行号
        """
        start_row = 22
        max_protected_row = 30  # 30行后的内容不要修改
        
        # 获取22行的样式作为参考（用于插入的行）
        ref_row_height = worksheet.row_dimensions[start_row].height
        
        # 获取22行的样式（用于合并单元格样式参考）
        ref_cell_b = worksheet[f'B{start_row}']
        ref_cell_d = worksheet[f'D{start_row}']
        ref_cell_i = worksheet[f'I{start_row}']
        ref_cell_k = worksheet[f'K{start_row}']
        
        # 准备数据
        product_models = []
        if 'product_model' in parameters and parameters['product_model']:
            product_models = [item.strip() for item in str(parameters['product_model']).split('/') if item.strip()]
        
        sales_names = []
        if 'sales_name' in parameters and parameters['sales_name']:
            sales_names = [item.strip() for item in str(parameters['sales_name']).split('/') if item.strip()]
        
        target_area = parameters.get('target_area', '')
        
        # 计算需要填充的最大行数：Max(product_model, sales_name)
        max_data_rows = max(len(product_models), len(sales_names), 1)
        
        # 第一步：计算需要插入的行数：Max(product_model, sales_name) - 8，如果最大行没超过8，不用插入
        rows_to_insert = max(0, max_data_rows - 8)
        
        # 如果超过8行，需要在22行下面插入行
        if rows_to_insert > 0:
            # 确保插入后不超过30行（不损坏后面的内容）
            # 计算插入后的最大行号
            final_max_row = start_row + max_data_rows - 1
            if final_max_row < max_protected_row:
                # 在22行下面插入行（从23行开始插入）
                # 从后往前插入，避免行号变化问题
                # 注意：openpyxl的insert_rows会在指定行之前插入，所以从后往前插入更安全
                for i in range(rows_to_insert - 1, -1, -1):
                    worksheet.insert_rows(start_row + 1)
                    # 调整受保护行的起始位置
                    max_protected_row += 1
        
        # 第一步完成后：为所有行设置合并单元格和样式（从22行到22+max_data_rows-1行）
        # 无论是否需要插入行，都需要设置合并单元格
        for row_idx in range(max_data_rows):
            row_num = start_row + row_idx
            if row_num >= max_protected_row:
                break
            
            # 合并 B~C列，样式和22行保持一致
            self._merge_and_copy_style(
                worksheet,
                f'B{row_num}',
                f'C{row_num}',
                ref_cell_b,
                ref_row_height
            )
            
            # 合并 D~H列，样式和22行保持一致
            self._merge_and_copy_style(
                worksheet,
                f'D{row_num}',
                f'H{row_num}',
                ref_cell_d,
                ref_row_height
            )
            
            # 合并 I~J列，样式和22行保持一致
            self._merge_and_copy_style(
                worksheet,
                f'I{row_num}',
                f'J{row_num}',
                ref_cell_i,
                ref_row_height
            )
            
            # 合并 K~M列，样式和22行保持一致
            self._merge_and_copy_style(
                worksheet,
                f'K{row_num}',
                f'M{row_num}',
                ref_cell_k,
                ref_row_height
            )
        
        # 确定实际填充的最大行
        actual_max_row = min(start_row + max_data_rows - 1, max_protected_row - 1)
        
        # 第二步：插入行后，再将product_model拆分的内容填充到D～H合并列，
        # sales_name内容填充到I～J合并列，target_area内容填充到K～M合并列
        # 填充时保持行高（已在_merge_and_copy_style中设置）
        for idx in range(max_data_rows):
            row_num = start_row + idx
            if row_num >= max_protected_row:
                break
            
            # 填充 product_model 拆分的内容到 D~H 合并列
            if idx < len(product_models):
                worksheet[f'D{row_num}'].value = product_models[idx]
            
            # 填充 sales_name 内容到 I~J 合并列
            if idx < len(sales_names):
                worksheet[f'I{row_num}'].value = sales_names[idx]
            
            # 填充 target_area 内容到 K~M 合并列
            if target_area:
                worksheet[f'K{row_num}'].value = target_area
        
        return actual_max_row
    
    def _merge_and_copy_style(self, worksheet, start_cell: str, end_cell: str, 
                              ref_cell, ref_row_height):
        """
        合并单元格并复制样式（不填充值，仅用于插入行时设置格式）
        
        Args:
            worksheet: 工作表对象
            start_cell: 起始单元格（如 'B22'）
            end_cell: 结束单元格（如 'C22'）
            ref_cell: 参考单元格（用于复制样式）
            ref_row_height: 参考行高
        """
        # 取消可能存在的合并
        merged_to_remove = []
        for merged_range in list(worksheet.merged_cells.ranges):
            merged_str = str(merged_range)
            if start_cell in merged_str or end_cell in merged_str:
                merged_to_remove.append(merged_str)
        
        for merged_str in merged_to_remove:
            try:
                worksheet.unmerge_cells(merged_str)
            except:
                pass
        
        # 合并单元格
        worksheet.merge_cells(f'{start_cell}:{end_cell}')
        
        # 复制样式（字号小1）
        if ref_cell.has_style:
            # 复制字体（字号小1）
            ref_font = ref_cell.font
            new_font = Font(
                name=ref_font.name,
                size=(ref_font.size or 11) - 1 if ref_font.size else 10,
                bold=ref_font.bold,
                italic=ref_font.italic,
                vertAlign=ref_font.vertAlign,
                underline=ref_font.underline,
                strike=ref_font.strike,
                color=ref_font.color
            )
            worksheet[start_cell].font = new_font
            
            # 复制对齐方式
            if ref_cell.alignment:
                worksheet[start_cell].alignment = Alignment(
                    horizontal=ref_cell.alignment.horizontal,
                    vertical=ref_cell.alignment.vertical,
                    wrap_text=ref_cell.alignment.wrap_text,
                    shrink_to_fit=ref_cell.alignment.shrink_to_fit,
                    indent=ref_cell.alignment.indent
                )
            
            # 复制边框
            if ref_cell.border:
                worksheet[start_cell].border = Border(
                    left=ref_cell.border.left,
                    right=ref_cell.border.right,
                    top=ref_cell.border.top,
                    bottom=ref_cell.border.bottom
                )
            
            # 复制填充
            if ref_cell.fill:
                worksheet[start_cell].fill = PatternFill(
                    fill_type=ref_cell.fill.fill_type,
                    start_color=ref_cell.fill.start_color,
                    end_color=ref_cell.fill.end_color
                )
        
        # 设置行高（保持行高）
        row_num = int(start_cell[1:])
        if ref_row_height:
            worksheet.row_dimensions[row_num].height = ref_row_height
    
    def _merge_and_fill_cells(self, worksheet, start_cell: str, end_cell: str, 
                              value: str, ref_cell, ref_row_height):
        """
        合并单元格并填充值，复制参考单元格的样式，字号小1
        
        Args:
            worksheet: 工作表对象
            start_cell: 起始单元格（如 'D22'）
            end_cell: 结束单元格（如 'H22'）
            value: 要填充的值
            ref_cell: 参考单元格（用于复制样式）
            ref_row_height: 参考行高
        """
        # 取消可能存在的合并
        for merged_range in list(worksheet.merged_cells.ranges):
            if start_cell in merged_range or end_cell in merged_range:
                worksheet.unmerge_cells(str(merged_range))
        
        # 合并单元格
        worksheet.merge_cells(f'{start_cell}:{end_cell}')
        
        # 填充值
        worksheet[start_cell].value = value
        
        # 复制样式（字号小1）
        if ref_cell.has_style:
            # 复制字体（字号小1）
            ref_font = ref_cell.font
            new_font = Font(
                name=ref_font.name,
                size=(ref_font.size or 11) - 1 if ref_font.size else 10,
                bold=ref_font.bold,
                italic=ref_font.italic,
                vertAlign=ref_font.vertAlign,
                underline=ref_font.underline,
                strike=ref_font.strike,
                color=ref_font.color
            )
            worksheet[start_cell].font = new_font
            
            # 复制对齐方式
            if ref_cell.alignment:
                worksheet[start_cell].alignment = Alignment(
                    horizontal=ref_cell.alignment.horizontal,
                    vertical=ref_cell.alignment.vertical,
                    wrap_text=ref_cell.alignment.wrap_text,
                    shrink_to_fit=ref_cell.alignment.shrink_to_fit,
                    indent=ref_cell.alignment.indent
                )
            
            # 复制边框
            if ref_cell.border:
                worksheet[start_cell].border = Border(
                    left=ref_cell.border.left,
                    right=ref_cell.border.right,
                    top=ref_cell.border.top,
                    bottom=ref_cell.border.bottom
                )
            
            # 复制填充
            if ref_cell.fill:
                worksheet[start_cell].fill = PatternFill(
                    fill_type=ref_cell.fill.fill_type,
                    start_color=ref_cell.fill.start_color,
                    end_color=ref_cell.fill.end_color
                )
        
        # 设置行高
        row_num = int(start_cell[1:])
        if ref_row_height:
            worksheet.row_dimensions[row_num].height = ref_row_height