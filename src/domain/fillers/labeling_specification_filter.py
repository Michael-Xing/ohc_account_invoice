"""标签仕样书-仕样确认书填充器"""

from pathlib import Path
from typing import Any, Dict, Optional
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from src.infrastructure.template_service import ExcelTemplateFiller


class LabelingSpecificationFiller(ExcelTemplateFiller):
    """标签仕样书-仕样确认书填充器"""

    # 填充器写入/修改过的单元格背景色：RGB(115,159,215) -> HEX 0x739FD7
    _FILLED_BG_COLOR = "FF739FD7"  # ARGB format

    def _apply_filled_background(self, cell) -> None:
        """将单元格背景色设置为填充高亮色"""
        cell.fill = PatternFill(fill_type="solid", fgColor=self._FILLED_BG_COLOR)

    def fill_template(
        self,
        template_path: Path,
        parameters: Dict[str, Any],
        output_path: Path,
        language: Optional[str] = None,
    ) -> bool:
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
            # 优先使用外部传入语言，否则从模板路径提取
            resolved_language = (language or self._extract_language_from_path(template_path)).strip().lower() or "zh"
            # 设置语言（用于空值兜底）
            self._set_language(resolved_language)

            workbook = load_workbook(template_path)
            worksheet = workbook.active

            # 填充字段
            self._fill_fields(worksheet, parameters, resolved_language)

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

    def _ensure_background_fill(self, worksheet, cell_addr: str, fallback_fill: Optional[PatternFill] = None) -> None:
        """
        确保单元格有背景色：
        - 若当前单元格无填充，则优先复制左侧相邻单元格的填充
        - 若左侧也无填充，则使用 fallback_fill（默认填充高亮色）
        """
        cell = worksheet[cell_addr]
        has_fill = bool(getattr(cell.fill, "patternType", None))
        if has_fill:
            return

        # 尝试复制左侧相邻单元格填充
        try:
            col = cell.column  # 1-based
            row = cell.row
            if col > 1:
                left = worksheet.cell(row=row, column=col - 1)
                if bool(getattr(left.fill, "patternType", None)):
                    cell.fill = left.fill
                    return
        except Exception:
            # 忽略复制失败，走 fallback
            pass

        if fallback_fill is None:
            # 默认使用填充高亮色 FF739FD7
            fallback_fill = PatternFill(patternType="solid", fgColor=self._FILLED_BG_COLOR)
        cell.fill = fallback_fill

    def _fill_fields(self, worksheet, parameters: Dict[str, Any], language: str = 'zh'):
        """填充字段到指定单元格，空值时使用语言对应的兜底文本"""
        # 获取空值兜底文本
        missing_text = self._missing_text()

        # theme_no 填入 D5 单元格
        if 'theme_no' in parameters and parameters['theme_no']:
            worksheet['D5'].value = str(parameters['theme_no'])
        else:
            worksheet['D5'].value = missing_text

        # theme_name 填入 M5 单元格
        if 'theme_name' in parameters and parameters['theme_name']:
            worksheet['M5'].value = str(parameters['theme_name'])
        else:
            worksheet['M5'].value = missing_text

        # product_model_name 填入 D7 单元格
        if 'product_model_name' in parameters and parameters['product_model_name']:
            worksheet['D7'].value = str(parameters['product_model_name'])
        else:
            worksheet['D7'].value = missing_text

        # representative_model 填入 G11 单元格 (代表型号)
        if 'representative_model' in parameters and parameters['representative_model']:
            worksheet['G11'].value = str(parameters['representative_model'])
        else:
            worksheet['G11'].value = missing_text

        # product_model 填入 G12 单元格 (商品型式名)
        if 'product_model' in parameters and parameters['product_model']:
            worksheet['G12'].value = str(parameters['product_model'])
        else:
            worksheet['G12'].value = missing_text

        # product_name 填入 G13 单元格
        if 'product_name' in parameters and parameters['product_name']:
            worksheet['G13'].value = str(parameters['product_name'])
        else:
            worksheet['G13'].value = missing_text

        # sales_name 填入 G14 单元格
        if 'sales_name' in parameters and parameters['sales_name']:
            worksheet['G14'].value = str(parameters['sales_name'])
        else:
            worksheet['G14'].value = missing_text

        worksheet['G15'].value = "Intellisense"
        worksheet['G16'].value = "OMRON"

        # ohc_target 填入 G19 单元格, 如果是OHC向以外则保持空白。
        if 'ohc_target' in parameters and parameters['ohc_target']:
            worksheet['G19'].value = "OHC提供"
            worksheet['G26'].value = "欧姆龙健康医疗（中国）有限公司"

        # sales_channel 填入 E21 单元格
        if 'sales_channel' in parameters and parameters['sales_channel']:
            if str(parameters['sales_channel']).strip() == "医療機関":
                worksheet['G21'].value = "400-889-0089"
            else:
                worksheet['G21'].value = "400-770-9988"

        worksheet['G25'].value = "注册证编号/产品技术要求编号"

        # sales_channel：若未提供（None/空）则用兜底文本，避免空白
        if not ('sales_channel' in parameters and str(parameters['sales_channel']).strip()):
            worksheet['G21'].value = missing_text

        if 'address' in parameters and parameters['address']:
            worksheet['G17'].value =  parameters['address']
        if 'country' in parameters and parameters['country']:
            worksheet['G18'].value =  parameters['country'] + "制造"

        # production_area
        if 'address' in parameters and parameters['address']:
            worksheet['G17'].value = parameters['address']
        else:
            worksheet['G17'].value = missing_text

        if 'country' in parameters and parameters['country']:
            worksheet['G18'].value = parameters['country']
        else:
            worksheet['G18'].value = missing_text


        # 为所有“填写/覆盖”的单元格补背景色（若模板本身无背景色）
        # 目标区域：表头与主要填写区
        fallback = PatternFill(patternType="solid", fgColor=self._FILLED_BG_COLOR)
        fill_cells = [
            "D5", "K5", "D7",
            "G11", "G12", "G13", "G14", "G15", "G16",
            "G17", "G18", "G19", "G21", "G25", "G26",
        ]
        for addr in fill_cells:
            self._ensure_background_fill(worksheet, addr, fallback)

        # C11~C26 为“项目项名称”列，同样确保背景色一致
        for r in range(11, 27):
            self._ensure_background_fill(worksheet, f"C{r}", fallback)

