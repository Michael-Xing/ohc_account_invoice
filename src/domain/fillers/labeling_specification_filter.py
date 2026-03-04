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
            print(f"标签仕样书模板填充失败: {str(e)}")
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
        # theme_no 填入 D5 单元格
        if 'theme_no' in parameters and parameters['theme_no']:
            worksheet['D5'].value = str(parameters['theme_no'])

        # theme_name 填入 K5 单元格
        if 'theme_name' in parameters and parameters['theme_name']:
            worksheet['K5'].value = str(parameters['theme_name'])

        # product_model_name 填入 D7 单元格
        if 'product_model_name' in parameters and parameters['product_model_name']:
            worksheet['D7'].value = str(parameters['product_model_name'])

        # product_model_name 填入 G12 单元格 (商品型式名)
        if 'product_model' in parameters and parameters['product_model']:
            worksheet['G12'].value = str(parameters['product_model'])

        # product_name 填入 G13 单元格 (販売名称)
        if 'product_name' in parameters and parameters['product_name']:
            worksheet['G13'].value = str(parameters['product_name'])

        # sales_name 填入 G14 单元格 （販売形式）
        if 'sales_name' in parameters and parameters['sales_name']:
            worksheet['G14'].value = str(parameters['sales_name'])

        worksheet['G15'].value = "Intellisense"
        worksheet['G16'].value = "OMRON"

        # ohc_target 填入 G19 单元格, 如果是OHC向以外则保持空白。
        if 'ohc_target' in parameters and parameters['ohc_target']:
            worksheet['G19'].value = "OHC提供"
            worksheet['G26'].value = "欧姆龙健康医疗（中国）有限公司"

        # sales_channel 填入 E21 单元格 （販売形式）
        if 'sales_channel' in parameters and parameters['sales_channel']:
            if str(parameters['sales_channel']).strip() == "医療機関":
                worksheet['G21'].value = "400-889-0089"
            else:
                worksheet['G21'].value = "400-770-9988"

        worksheet['G25'].value = "注册证编号/产品技术要求编号"



        production_area_map = {
            "OMD": {
                "country": "中国",
                "address": "欧姆龙（大连）有限公司大连经济技术开发区松江路3号",
                "phone": "",
            },
            "OHZ": {
                "country": "日本",
                "address": "",
                "phone": "",
            },
            "OHV": {
                "country": "越南",
                "address": "",
                "phone": "",
            },
        }

        # production_area
        if 'production_area' in parameters and parameters['production_area']:
            info = production_area_map.get(parameters['production_area'], {})
            worksheet['G17'].value = info.get('address', '')
            worksheet['G18'].value = info.get('country', '')+'制造'

        texts = ""
        if language == 'ja':
            texts = [
                '代表型番',
                '商品型式名',
                '販売名称',
                '販売形式',
                'ｴﾘｱﾈｰﾐﾝｸﾞ（販売商品コード）',
                'OMRON　ロゴ',
                '製造元',
                '生産国',
                'JANコード',
                'ITFコート',
                'お問合せ先 ',
                '医療機器分類',
                '類別番号および類別名称',
                '使用目的/効能効果',
                '医療機器認証番号',
                '製造販売元'
            ]

        if language == 'zh':
            texts = [
                '代表型号',
                '商品型号名称',
                '销售名称',
                '销售形式',
                '区域命名（销售商品代码）',
                'OMRON 标志',
                '制造商',
                '生产国',
                'JAN 代码',
                'ITF 代码',
                '咨询联系方式',
                '医疗器械分类',
                '类别编号及类别名称',
                '使用目的／功效',
                '医疗器械认证编号',
                '制造销售商'
            ]

        if language == 'en':
            texts = [
                'Representative Model Number',
                'Product Model Name',
                'Sales Name',
                'Sales Format',
                'Area Naming (Sales Product Code)',
                'OMRON Logo',
                'Manufacturer',
                'Country of Origin',
                'JAN Code',
                'ITF Code',
                'Contact Information',
                'Medical Device Classification',
                'Class Number and Class Name',
                'Intended Use / Effectiveness',
                'Medical Device Certification Number',
                'Manufacturer and Distributor'
            ]

        for idx, text in enumerate(texts):
            row = 11 + idx  # 从第11行开始
            worksheet[f'C{row}'].value = text

