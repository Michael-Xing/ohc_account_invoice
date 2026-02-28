"""使用说明书仕样书填充器"""

import re
from pathlib import Path
from typing import Dict, Any, List
from openpyxl import load_workbook

from src.infrastructure.template_service import ExcelTemplateFiller

class UserManualSpecificationFiller(ExcelTemplateFiller):
    """使用说明书仕样书填充器"""

    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        """
        填充使用说明书仕样书模板

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

            # 填充字段
            self._fill_fields(worksheet, parameters)

            # 替换其他占位符
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell.value = self._replace_placeholders(cell.value, parameters)

            workbook.save(output_path)
            return True
        except Exception as e:
            print(f"使用说明书仕样书模板填充失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def _fill_fields(self, worksheet, parameters: Dict[str, Any]):
        """填充字段到指定单元格"""
        # theme_no 填入 B19 单元格
        if 'theme_no' in parameters and parameters['theme_no']:
            worksheet['B19'].value = str(parameters['theme_no'])

        # theme_name 填入 D19 单元格
        if 'theme_name' in parameters and parameters['theme_name']:
            worksheet['D19'].value = str(parameters['theme_name'])

        # product_model_name 填入 J19 单元格
        if 'product_model_name' in parameters and parameters['product_model_name']:
            worksheet['J19'].value = str(parameters['product_model_name'])

        # sales_name 填入 B21 单元格
        if 'sales_name' in parameters and parameters['sales_name']:
            worksheet['B21'].value = str(parameters['sales_name'])

        related_file_numbers: List[Dict[str, Any]] = parameters.get('related_file_numbers', [])
        if related_file_numbers:
            mapping = {item['short_name']: item for item in related_file_numbers}
            row = 25
            while True:
                cell_value = worksheet.cell(row=row, column=1).value  # A列
                if cell_value is None or str(cell_value).strip() == '':
                    break
                key = str(cell_value).strip()
                if key in mapping:
                    worksheet.cell(row=row, column=4).value = str(mapping[key]['file_number'])  # D列
                    worksheet.cell(row=row, column=7).value = str(mapping[key]['version'])  # G列
                row += 1


