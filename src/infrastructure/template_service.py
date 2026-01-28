"""模板填充服务模块（已迁移到 infrastructure 层）"""

import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, Optional

from docx import Document
from openpyxl import load_workbook

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


class ExcelTemplateFiller(TemplateFillerStrategy):
    """Excel模板填充策略"""
    
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