"""PTF INDEX填充器"""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, PatternFill

from src.infrastructure.template_service import ExcelTemplateFiller

logger = logging.getLogger(__name__)


class PTFIndexFiller(ExcelTemplateFiller):
    """PTF INDEX填充器"""

    # 填充器写入/修改过的单元格背景色：RGB(115,159,215) -> HEX 0x739FD7
    _FILLED_BG_COLOR = "FF739FD7"  # ARGB format

    def _apply_filled_background(self, cell) -> None:
        """将单元格背景色设置为填充高亮色"""
        cell.fill = PatternFill(fill_type="solid", fgColor=self._FILLED_BG_COLOR)

    def _get_merged_cell_top_left(self, worksheet, row, col):
        """
        获取合并单元格的左上角单元格
        如果单元格不是合并单元格，返回该单元格本身
        """
        cell = worksheet.cell(row, col)
        if isinstance(cell, MergedCell):
            for merged_range in worksheet.merged_cells.ranges:
                if (merged_range.min_row <= row <= merged_range.max_row and
                        merged_range.min_col <= col <= merged_range.max_col):
                    return worksheet.cell(merged_range.min_row, merged_range.min_col)
        return cell

    def fill_template(
        self,
        template_path: Path,
        parameters: Dict[str, Any],
        output_path: Path,
        language: Optional[str] = None,
    ) -> bool:
        """
        填充PTF INDEX模板

        Args:
            template_path: 模板文件路径
            parameters: 填充参数
            output_path: 输出文件路径

        Returns:
            bool: 是否成功
        """
        non_empty_fields = [k for k, v in parameters.items() if v]
        logger.info("[PTFIndexFiller] 填充字段: %s", non_empty_fields)
        try:
            # 设置语言（用于空值兜底）
            self._set_language(language)

            # 加载模板文件
            workbook = load_workbook(template_path, keep_vba=True)
            worksheet = workbook.active

            # 根据 target_area 匹配 D15:H15 表头，按顺序填充对应行列
            self._fill_data_by_area(worksheet, parameters)

            # 保存工作簿
            workbook.save(output_path)
            return True
        except Exception as e:
            logger.error("PTF INDEX模板填充失败: %s", str(e), exc_info=True)
            return False

    # file_numbers 按顺序填充的模板行号（非连续）
    _FILE_NUMBER_ROWS = [
        *range(19, 22),    # 19-21
        *range(23, 37),    # 23-36
        *range(38, 44),    # 38-43
        *range(45, 50),    # 45-49
        51,
        *range(53, 56),    # 53-55
        57, 58,
        60,
        *range(62, 66),    # 62-65
    ]

    def _fill_data_by_area(self, worksheet, parameters: Dict[str, Any]) -> None:
        """
        根据 target_area 匹配 D15~H15 表头，填充 sales_name 和 file_numbers 到匹配列

        target_area 支持多个值用半角逗号连接，例如 "日本,中国"
        D15~H15 为贩卖国家表头，匹配成功后：
        - file_numbers 按顺序依次填充到指定的数据行
        未匹配的列清空数据行内容
        """
        target_area_raw: str = parameters.get("target_area", "")
        file_numbers: List[str] = parameters.get("file_numbers", [])
        missing_text = self._missing_text()

        if not target_area_raw:
            return

        # 按半角逗号分割 target_area
        target_areas = [a.strip() for a in target_area_raw.split(",") if a.strip()]
        if not target_areas:
            return

        # 读取 D15~H15 表头，构建 表头文本 -> 列号 的映射
        header_map: Dict[str, int] = {}
        for col in range(4, 9):  # D=4, E=5, F=6, G=7, H=8
            cell = self._get_merged_cell_top_left(worksheet, 15, col)
            header_value = str(cell.value).strip() if cell.value else ""
            if header_value:
                header_map[header_value] = col

        # 匹配 target_area 与表头
        matched_cols: List[int] = []
        for area in target_areas:
            col = header_map.get(area)
            if col is not None:
                matched_cols.append(col)
            else:
                logger.warning(
                    "[PTFIndexFiller] target_area '%s' 未在 D15:H15 中匹配到（表头: %s）",
                    area, list(header_map.keys()),
                )

        if not matched_cols:
            return

        # 未匹配的列，清空数据行内容
        all_area_cols = set(header_map.values())
        unmatched_cols = all_area_cols - set(matched_cols)
        for col in unmatched_cols:
            for row in [16] + self._FILE_NUMBER_ROWS:
                cell = self._get_merged_cell_top_left(worksheet, row, col)
                if cell.value and isinstance(cell.value, str):
                    cell.value = ""

        # 填充匹配的列
        for col in matched_cols:
            # 按顺序依次填充 file_numbers，空值时使用兜底文本
            for idx, row in enumerate(self._FILE_NUMBER_ROWS):
                cell = self._get_merged_cell_top_left(worksheet, row, col)
                if idx < len(file_numbers) and file_numbers[idx]:
                    cell.value = file_numbers[idx]
                else:
                    cell.value = missing_text
                self._apply_filled_background(cell)
                self._set_wrap_text(cell)

    def _set_wrap_text(self, cell) -> None:
        """设置单元格自动换行，保留原有对齐方式"""
        if cell.alignment:
            cell.alignment = Alignment(
                horizontal=cell.alignment.horizontal,
                vertical=cell.alignment.vertical,
                text_rotation=cell.alignment.text_rotation,
                wrap_text=True,
                shrink_to_fit=cell.alignment.shrink_to_fit,
                indent=cell.alignment.indent,
            )
        else:
            cell.alignment = Alignment(wrap_text=True)
