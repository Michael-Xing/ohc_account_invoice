"""包装设计仕样书填充器"""

from pathlib import Path
from typing import Dict, Any
from openpyxl import load_workbook

from src.infrastructure.template_service import ExcelTemplateFiller


class PackagingDesignSpecificationFiller(ExcelTemplateFiller):
    """包装设计仕样书填充器"""

    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        """
        填充包装设计仕样书模板

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

            # 从模板路径中提取语言
            language = self._extract_language_from_path(template_path)

            # 填充字段
            self._fill_fields(worksheet, parameters, language)

            # 替换其他占位符
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell.value = self._replace_placeholders(cell.value, parameters)

            workbook.save(output_path)
            return True
        except Exception as e:
            print(f"包装设计仕样书模板填充失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def _extract_language_from_path(self, template_path: Path) -> str:
        """从模板路径中提取语言代码"""
        parts = template_path.parts
        # 路径格式通常是: .../excel/zh/文件名.xlsx 或 .../excel/ja/文件名.xlsx
        for part in parts:
            if part in ['zh', 'ja', 'en']:
                return part
        return 'zh'  # 默认返回中文

    def _fill_fields(self, worksheet, parameters: Dict[str, Any], language: str = 'zh'):
        """填充字段到指定单元格"""
        # theme_no 填入 C21 单元格
        if 'theme_no' in parameters and parameters['theme_no']:
            worksheet['C21'].value = str(parameters['theme_no'])

        # theme_name 填入 E21 单元格
        if 'theme_name' in parameters and parameters['theme_name']:
            worksheet['E21'].value = str(parameters['theme_name'])

        # product_model_name 填入 L21 单元格
        if 'product_model_name' in parameters and parameters['product_model_name']:
            worksheet['L21'].value = str(parameters['product_model_name'])

        # sales_name 填入 C23 单元格
        if 'sales_name' in parameters and parameters['sales_name']:
            worksheet['C23'].value = str(parameters['sales_name'])

        texts = []
        # 根据语言填充文档类型列表到 B27 列往下
        if language == 'ja':
            texts = [
                '製品要件書',
                '要求仕様書',
                '製品設計仕様書',
                'リスクコントロール仕様書',
                'ユーザビリティ仕様書',
                '課題分析·対策書',
                '製品アセスメント要項書'
            ]
        elif language == 'zh':
            texts = [
                '产品要件书',
                '要求仕样书',
                '产品设计仕样书',
                '风险控制仕样书',
                '可用性仕样书',
                '课题分析/对策结果书',
                '产品环境评估要项书'
            ]

        # 从 B27 开始填充
        for idx, text in enumerate(texts):
            row = 27 + idx
            worksheet[f'B{row}'].value = text

