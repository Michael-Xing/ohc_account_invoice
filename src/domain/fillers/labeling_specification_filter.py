"""标签仕样书-仕样确认书填充器"""

from pathlib import Path
from typing import Dict, Any
from openpyxl import load_workbook

from src.infrastructure.template_service import ExcelTemplateFiller


class LabelingSpecificationFiller(ExcelTemplateFiller):
    """标签仕样书-仕样确认书填充器"""

    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        """
        填充标签仕样书-仕样确认书模板

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
            print(f"标签仕样书模板填充失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def _fill_fields(self, worksheet, parameters: Dict[str, Any]):
        """填充字段到指定单元格"""
        # theme_no 填入 D5 单元格
        if 'theme_no' in parameters and parameters['theme_no']:
            worksheet['D5'].value = str(parameters['theme_no'])

        # theme_name 填入 K5 单元格
        if 'theme_name' in parameters and parameters['theme_name']:
            worksheet['K5'].value = str(parameters['theme_name'])

        # product_model_name 填入 D7 单元格
        if 'product_model_name' in parameters and parameters['product_model_name']:
            worksheet['D7'].value = str(parameters['product_model_name'])

