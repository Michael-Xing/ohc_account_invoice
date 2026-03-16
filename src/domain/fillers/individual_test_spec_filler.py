"""个别试验要项书填充器

要求：
- 仅填充指定单元格（含合并单元格的左上角单元格）
- 不影响其它单元格内容与样式
- 对 C37 / C44 设置左上对齐，并做 2 级缩进
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Optional

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment

from src.infrastructure.template_service import ExcelTemplateFiller


class IndividualTestSpecFiller(ExcelTemplateFiller):
    """个别试验要项书填充器（Excel）"""

    # 与 DHFIndexFiller 保持一致的高亮背景色（ARGB）
    _FILLED_BG_COLOR = "FF739FD7"  # RGB(115,159,215)

    def _apply_filled_background(self, cell) -> None:
        """将单元格背景色设置为填充高亮色。"""
        cell.fill = PatternFill(fill_type="solid", fgColor=self._FILLED_BG_COLOR)

    def fill_template(
        self,
        template_path: Path,
        parameters: Dict[str, Any],
        output_path: Path,
        language: Optional[str] = None,
    ) -> bool:
        try:
            # 设置语言（用于空值兜底）
            self._set_language(language)

            workbook = load_workbook(template_path)
            worksheet = workbook.active

            self._fill_fields(worksheet, parameters)

            # 替换其它占位符（如果模板里存在 {xxx} / {{xxx}}）
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        cell.value = self._replace_placeholders(cell.value, parameters)

            workbook.save(output_path)
            return True
        except Exception as e:
            print(f"个别试验要项书模板填充失败: {str(e)}")
            import traceback

            traceback.print_exc()
            return False

    def _fill_fields(self, worksheet, parameters: Dict[str, Any]) -> None:
        # 获取空值兜底文本
        missing_text = self._missing_text()

        # 1) 基础字段：只写值，不动样式
        mapping = {
            "test_name": "D3",  # 合并后单元格 D3
            "test_number": "K3",  # 合并后单元格 K3
            "theme_no": "D5",  # 合并后单元格 D5
            "product_model": "K5",  # 合并后单元格 K5
            "meas_temperature": "L8",  # 合并后单元格 L8
            "meas_humidity": "N8",  # 合并后单元格 N8
        }

        for key, addr in mapping.items():
            val = parameters.get(key, "")
            if val is None:
                val = ""
            # 空值时填充默认值
            if str(val) == "":
                cell = worksheet[addr]
                cell.value = missing_text
                self._apply_filled_background(cell)
            else:
                cell = worksheet[addr]
                cell.value = str(val)
                self._apply_filled_background(cell)

        # 2) 长文本：需要左上对齐 + 行首缩进两个字符（Excel 缩进）
        test_purpose = parameters.get("test_purpose")
        self._set_text_block(
            worksheet=worksheet,
            cell_addr="C37",
            text=str(test_purpose) if test_purpose else missing_text,
        )
        test_conditions = parameters.get("test_conditions")
        self._set_text_block(
            worksheet=worksheet,
            cell_addr="C44",
            text=str(test_conditions) if test_conditions else missing_text,
        )

    def _set_text_block(self, worksheet, cell_addr: str, text: str) -> None:
        if text == "":
            # 同上：为空时不覆盖，避免清掉模板内可能存在的提示文案/默认内容
            return

        cell = worksheet[cell_addr]
        cell.value = text

        # 写入长文本的同时也高亮
        self._apply_filled_background(cell)

        # 不改字体/边框/填充，只在现有 alignment 基础上做小改动。
        # 兼容不同版本的 openpyxl：有的没有 Alignment.copy() 方法，这里直接构造一个新的 Alignment。
        old_alignment = cell.alignment
        cell.alignment = Alignment(
            horizontal=old_alignment.horizontal or "left",
            vertical=old_alignment.vertical or "top",
            indent=1,
            wrap_text=True,
        )

