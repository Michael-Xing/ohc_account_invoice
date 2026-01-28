from pathlib import Path
from typing import Dict, Any

from openpyxl import load_workbook

from src.infrastructure.template_service import ExcelTemplateFiller


class DHFIndexFiller(ExcelTemplateFiller):
    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path) -> bool:
        try:
            workbook = load_workbook(template_path)
            worksheet = workbook.active
            self._fill_project_info(worksheet, parameters)
            self._fill_document_list(worksheet, parameters)
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell.value = self._replace_placeholders(cell.value, parameters)
            workbook.save(output_path)
            return True
        except Exception as e:
            print(f"DHF INDEX模板填充失败: {str(e)}")
            return False

    def _fill_project_info(self, worksheet, parameters: Dict[str, Any]):
        project_info = {
            'A1': parameters.get('project_name', ''),
            'A2': parameters.get('version', ''),
            'A3': parameters.get('date', ''),
            'A4': parameters.get('author', ''),
            'A5': parameters.get('department', ''),
            'A6': parameters.get('document_type', ''),
            'A7': parameters.get('reviewer', ''),
            'A8': parameters.get('approval_date', '')
        }
        for cell_ref, value in project_info.items():
            if value:
                worksheet[cell_ref] = value

    def _fill_document_list(self, worksheet, parameters: Dict[str, Any]):
        pass


