"""模板填充服务模块（已迁移到 infrastructure 层）"""

import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, Optional

from docx import Document
from docx.shared import RGBColor
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from src.config import settings


class TemplateFillerStrategy(ABC):
    """模板填充策略抽象基类"""
    
    @abstractmethod
    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        """填充模板"""
        pass
    
    def _replace_placeholders(self, text: str, parameters: Dict[str, Any]) -> str:
        """替换文本中的占位符"""
        if not text or not isinstance(text, str):
            return text
        
        def replace_double_brace(match):
            placeholder = match.group(1).strip()
            if placeholder in parameters:
                return str(parameters[placeholder])
            return match.group(0)
        
        text = re.sub(r'\{\{([^}]+)\}\}', replace_double_brace, text)
        
        def replace_single_brace(match):
            placeholder = match.group(1).strip()
            if placeholder in parameters:
                return str(parameters[placeholder])
            return match.group(0)
        
        text = re.sub(r'\{([^}]+)\}', replace_single_brace, text)
        
        return text


#: Excel/Word 统一使用的“已填充”主题蓝色（RGB 115,159,215）
WORD_FILLED_TEXT_COLOR = RGBColor(0x73, 0x9F, 0xD7)


class ExcelTemplateFiller(TemplateFillerStrategy):
    """Excel模板填充策略"""

    _FILLED_CELL_COLOR = "FF739FD7"
    
    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        try:
            workbook = load_workbook(template_path)
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                for row in worksheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            cell.value = self._replace_placeholders(cell.value, parameters)
            workbook.save(output_path)
            return True
        except Exception as e:
            print(f"Excel模板填充失败: {str(e)}")
            return False

    def _create_filled_cell_pattern(self) -> PatternFill:
        """创建统一的已填充单元格背景色样式（RGB 115,159,215）"""
        return PatternFill(fill_type="solid", fgColor=self._FILLED_CELL_COLOR)

    def _set_cell_value_with_fill(self, cell, value: Any) -> None:
        """设置单元格值并应用统一背景色"""
        cell.value = str(value)
        cell.fill = self._create_filled_cell_pattern()

    def _set_worksheet_cell_with_fill(self, worksheet, cell_ref: str, value: Any) -> None:
        """通过单元格引用设置值并应用统一背景色"""
        self._set_cell_value_with_fill(worksheet[cell_ref], value)

    def _set_cell_with_fill_by_position(self, worksheet, row: int, column: int, value: Any) -> None:
        """通过行列坐标设置值并应用统一背景色"""
        self._set_cell_value_with_fill(worksheet.cell(row=row, column=column), value)


class WordTemplateFiller(TemplateFillerStrategy):
    """Word模板填充策略"""
    
    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        try:
            doc = Document(template_path)
            for paragraph in doc.paragraphs:
                if paragraph.text:
                    paragraph.text = self._replace_placeholders(paragraph.text, parameters)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text:
                            cell.text = self._replace_placeholders(cell.text, parameters)
            for section in doc.sections:
                if section.header:
                    for paragraph in section.header.paragraphs:
                        if paragraph.text:
                            paragraph.text = self._replace_placeholders(paragraph.text, parameters)
                if section.footer:
                    for paragraph in section.footer.paragraphs:
                        if paragraph.text:
                            paragraph.text = self._replace_placeholders(paragraph.text, parameters)
            doc.save(output_path)
            return True
        except Exception as e:
            print(f"Word模板填充失败: {str(e)}")
            return False