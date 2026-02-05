from pathlib import Path
from typing import Dict, Any, List

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, PatternFill, Protection, Side
from openpyxl.cell.cell import MergedCell

from src.infrastructure.template_service import ExcelTemplateFiller


class DHFIndexFiller(ExcelTemplateFiller):
    """DHF INDEX填充器"""
    
    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        """
        填充DHF INDEX模板
        
        Args:
            template_path: 模板文件路径
            parameters: 填充参数
            output_path: 输出文件路径
            
        Returns:
            bool: 是否成功
        """
        try:
            # 加载模板文件
            workbook = load_workbook(template_path, data_only=False, keep_vba=False)
            worksheet = workbook.active
            
            # 填充基本信息单元格
            self._fill_basic_info(worksheet, parameters)
            
            # 填充文件列表
            self._fill_file_list(worksheet, parameters)
            
            # 替换其他占位符
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell.value = self._replace_placeholders(cell.value, parameters)
            
            # 保存工作簿
            workbook.save(output_path)
            
            return True
        except Exception as e:
            print(f"DHF INDEX模板填充失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def _fill_basic_info(self, worksheet, parameters: Dict[str, Any]):
        """填充基本信息到C3-C7单元格"""
        # theme_no: 填充到C3单元格
        if 'theme_no' in parameters and parameters['theme_no']:
            worksheet['C3'] = str(parameters['theme_no'])
        
        # theme_name: 填充到C4单元格
        if 'theme_name' in parameters and parameters['theme_name']:
            worksheet['C4'] = str(parameters['theme_name'])
        
        # product_model: 根据"/"分割，填充到C5单元格
        if 'product_model' in parameters and parameters['product_model']:
            product_model = str(parameters['product_model'])
            # 根据"/"分割，用换行符连接
            product_model_list = [item.strip() for item in product_model.split('/') if item.strip()]
            cell_c5 = worksheet['C5']
            cell_c5.value = '\n'.join(product_model_list)
            # 设置单元格为自动换行，保留原有对齐方式的其他属性
            if cell_c5.alignment:
                cell_c5.alignment = Alignment(
                    horizontal=cell_c5.alignment.horizontal,
                    vertical=cell_c5.alignment.vertical,
                    text_rotation=cell_c5.alignment.text_rotation,
                    wrap_text=True,
                    shrink_to_fit=cell_c5.alignment.shrink_to_fit,
                    indent=cell_c5.alignment.indent
                )
            else:
                cell_c5.alignment = Alignment(wrap_text=True)
        
        # sales_name: 根据"/"分割，填充到C6单元格
        if 'sales_name' in parameters and parameters['sales_name']:
            sales_name = str(parameters['sales_name'])
            # 根据"/"分割，用换行符连接
            sales_name_list = [item.strip() for item in sales_name.split('/') if item.strip()]
            cell_c6 = worksheet['C6']
            cell_c6.value = '\n'.join(sales_name_list)
            # 设置单元格为自动换行，保留原有对齐方式的其他属性
            if cell_c6.alignment:
                cell_c6.alignment = Alignment(
                    horizontal=cell_c6.alignment.horizontal,
                    vertical=cell_c6.alignment.vertical,
                    text_rotation=cell_c6.alignment.text_rotation,
                    wrap_text=True,
                    shrink_to_fit=cell_c6.alignment.shrink_to_fit,
                    indent=cell_c6.alignment.indent
                )
            else:
                cell_c6.alignment = Alignment(wrap_text=True)
        
        # stage: 拼接到C7单元格内容的后面
        if 'stage' in parameters and parameters['stage']:
            cell_c7 = worksheet['C7']
            original_value = cell_c7.value or ''
            cell_c7.value = str(original_value) + str(parameters['stage'])

    def _get_merged_cell_top_left(self, worksheet, row, col):
        """
        获取合并单元格的左上角单元格
        如果单元格不是合并单元格，返回该单元格本身
        """
        cell = worksheet.cell(row, col)
        
        # 如果是 MergedCell，找到它所属的合并区域
        if isinstance(cell, MergedCell):
            for merged_range in worksheet.merged_cells.ranges:
                if (merged_range.min_row <= row <= merged_range.max_row and
                    merged_range.min_col <= col <= merged_range.max_col):
                    # 返回合并区域的左上角单元格
                    return worksheet.cell(merged_range.min_row, merged_range.min_col)
        
        return cell

    def _copy_cell_style(self, source_cell, target_cell):
        """
        复制源单元格的样式到目标单元格
        
        Args:
            source_cell: 源单元格
            target_cell: 目标单元格
        """
        # 复制字体
        if source_cell.font:
            target_cell.font = Font(
                name=source_cell.font.name,
                size=source_cell.font.size,
                bold=source_cell.font.bold,
                italic=source_cell.font.italic,
                vertAlign=source_cell.font.vertAlign,
                underline=source_cell.font.underline,
                strike=source_cell.font.strike,
                color=source_cell.font.color
            )
        
        # 复制对齐方式
        if source_cell.alignment:
            target_cell.alignment = Alignment(
                horizontal=source_cell.alignment.horizontal,
                vertical=source_cell.alignment.vertical,
                text_rotation=source_cell.alignment.text_rotation,
                wrap_text=source_cell.alignment.wrap_text,
                shrink_to_fit=source_cell.alignment.shrink_to_fit,
                indent=source_cell.alignment.indent
            )
        
        # 复制边框（创建新的 Border 对象，避免 StyleProxy 问题）
        if source_cell.border:
            source_border = source_cell.border
            # 辅助函数：安全地复制 Side 对象
            def copy_side(side_obj):
                if side_obj:
                    return Side(
                        style=getattr(side_obj, 'style', None),
                        color=getattr(side_obj, 'color', None)
                    )
                return None
            
            target_cell.border = Border(
                left=copy_side(getattr(source_border, 'left', None)),
                right=copy_side(getattr(source_border, 'right', None)),
                top=copy_side(getattr(source_border, 'top', None)),
                bottom=copy_side(getattr(source_border, 'bottom', None)),
                diagonal=copy_side(getattr(source_border, 'diagonal', None)),
                diagonal_direction=getattr(source_border, 'diagonal_direction', None),
                outline=getattr(source_border, 'outline', None),
                vertical=copy_side(getattr(source_border, 'vertical', None)),
                horizontal=copy_side(getattr(source_border, 'horizontal', None))
            )
        
        # 复制填充（创建新的 PatternFill 对象，避免 StyleProxy 问题）
        if source_cell.fill and source_cell.fill.patternType:
            source_fill = source_cell.fill
            target_cell.fill = PatternFill(
                fill_type=source_fill.patternType,
                start_color=source_fill.start_color,
                end_color=source_fill.end_color,
                fgColor=source_fill.fgColor,
                bgColor=source_fill.bgColor
            )
        
        # 复制数字格式
        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format
        
        # 复制保护
        if source_cell.protection:
            target_cell.protection = Protection(
                locked=source_cell.protection.locked,
                hidden=source_cell.protection.hidden
            )

    def _fill_file_list(self, worksheet, parameters: Dict[str, Any]):
        """
        填充文件列表
        file_name关键词匹配B29～B196单元格，匹配上后填充number到G列，stage到H列
        填充时会复制B列的样式到G列和H列
        """
        if 'file_list' not in parameters or not parameters['file_list']:
            return
        
        file_list = parameters['file_list']
        
        # 遍历B29～B196单元格，进行关键词匹配
        for row_num in range(29, 197):  # B26到B196
            cell_b = worksheet.cell(row_num, 2)  # B列
            # 获取B列的实际单元格（如果是合并单元格，获取左上角的单元格）
            cell_b_actual = self._get_merged_cell_top_left(worksheet, row_num, 2)
            cell_b_value = str(cell_b_actual.value) if cell_b_actual.value else ''
            
            # 遍历file_list，查找匹配的file_name
            for file_item in file_list:
                # 处理file_item可能是字典或对象的情况
                if isinstance(file_item, dict):
                    file_name = str(file_item.get('file_name', ''))
                    number = str(file_item.get('number', ''))
                    stage = str(file_item.get('stage', ''))
                else:
                    # 假设是对象，尝试获取属性
                    file_name = str(getattr(file_item, 'file_name', ''))
                    number = str(getattr(file_item, 'number', ''))
                    stage = str(getattr(file_item, 'stage', ''))
                
                if not file_name:
                    continue
                
                # 关键词匹配：检查file_name是否在B列单元格内容中，或B列单元格内容是否包含file_name
                # 这里使用包含匹配（双向）
                if file_name in cell_b_value or cell_b_value in file_name:
                    # 匹配成功，填充number到G列，stage到H列
                    # 获取G列和H列的左上角单元格（如果是合并单元格）
                    cell_g = self._get_merged_cell_top_left(worksheet, row_num, 7)  # G列
                    cell_h = self._get_merged_cell_top_left(worksheet, row_num, 8)  # H列
                    
                    # 填充值
                    if number:
                        cell_g.value = number
                    if stage:
                        cell_h.value = stage
                    
                    # 复制B列的样式到G列和H列
                    self._copy_cell_style(cell_b_actual, cell_g)
                    self._copy_cell_style(cell_b_actual, cell_h)
                    
                    # 找到匹配后，跳出内层循环，继续下一行
                    break


