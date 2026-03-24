"""个别试验要项书填充器

要求：
- 仅填充指定单元格（含合并单元格的左上角单元格）
- 不影响其它单元格内容与样式
- 对 C37 / C44 设置左上对齐，并做 2 级缩进
- 支持动态调整行高以完整展示内容
- 支持在 C44 位置处理 markdown 表格，将其转换为 Excel 表格插入
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side

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

            # 为 product_model (K5) 动态调整行高
            if key == "product_model" and val:
                self._set_wrap_text_and_adjust_row_height(worksheet, cell)

        # 2) 长文本：需要左上对齐 + 行首缩进两个字符（Excel 缩进）
        test_purpose = parameters.get("test_purpose")
        self._set_text_block_with_row_height(
            worksheet=worksheet,
            cell_addr="C37",
            text=str(test_purpose) if test_purpose else missing_text,
        )
        test_conditions = parameters.get("test_conditions")
        self._set_test_conditions_with_tables(
            worksheet=worksheet,
            cell_addr="C44",
            text=str(test_conditions) if test_conditions else missing_text,
        )

    def _set_wrap_text_and_adjust_row_height(self, worksheet, cell) -> None:
        """设置单元格自动换行并动态调整行高"""
        if cell.alignment:
            cell.alignment = Alignment(
                horizontal=cell.alignment.horizontal,
                vertical=cell.alignment.vertical,
                text_rotation=cell.alignment.text_rotation,
                wrap_text=True,
                shrink_to_fit=cell.alignment.shrink_to_fit,
                indent=cell.alignment.indent
            )
        else:
            cell.alignment = Alignment(wrap_text=True)
        self._adjust_row_height_for_text(worksheet, cell.row, cell.column)

    def _adjust_row_height_for_text(self, worksheet, row: int, column: int) -> None:
        """
        根据文本长度自适应调整行高
        
        Args:
            worksheet: 工作表对象
            row: 行号
            column: 列号
        """
        cell = worksheet.cell(row, column)
        text = cell.value
        if not text:
            return
        
        # 获取列宽（字符数），如果未设置则使用默认值
        column_letter = cell.column_letter
        column_width = worksheet.column_dimensions[column_letter].width
        if not column_width or column_width == 0:
            column_width = 15  # 默认列宽
        
        # 获取字体大小，如果未设置则使用默认值
        font_size = 11  # 默认字体大小
        if cell.font and cell.font.size:
            font_size = cell.font.size
        
        # 估算每行可容纳的字符数（考虑中文字符占2个位置，英文占1个位置）
        chars_per_line = int(column_width * 1.7)
        if chars_per_line <= 0:
            chars_per_line = 15
        
        # 计算需要的行数
        lines = text.split('\n')
        total_lines = 0
        for line in lines:
            if line:
                line_count = (len(line) + chars_per_line - 1) // chars_per_line
                total_lines += max(1, line_count)
            else:
                total_lines += 1
        
        if total_lines == 0:
            total_lines = 1
        
        # 计算行高（点）
        row_height = total_lines * font_size * 1.5
        
        # 设置最小和最大行高限制
        min_height = font_size * 1.2
        max_height = font_size * 25
        row_height = max(min_height, min(row_height, max_height))
        
        # 设置行高（如果当前行高小于计算出的行高，则更新）
        current_height = worksheet.row_dimensions[row].height
        if not current_height or current_height < row_height:
            worksheet.row_dimensions[row].height = row_height

    def _set_text_block_with_row_height(self, worksheet, cell_addr: str, text: str) -> None:
        """设置长文本块并动态调整行高"""
        if text == "":
            return

        cell = worksheet[cell_addr]
        cell.value = text
        self._apply_filled_background(cell)

        # 设置对齐和换行
        old_alignment = cell.alignment
        cell.alignment = Alignment(
            horizontal=old_alignment.horizontal or "left",
            vertical=old_alignment.vertical or "top",
            indent=1,
            wrap_text=True,
        )

        # 动态调整行高
        self._adjust_row_height_for_text(worksheet, cell.row, cell.column)

    def _set_test_conditions_with_tables(self, worksheet, cell_addr: str, text: str) -> None:
        """
        设置试验条件文本，支持 markdown 表格转换为 Excel 表格
        内容结构：文本 + 表格 + 文本 + 表格 ...
        """
        if text == "":
            return

        # 解析混合内容（文本和表格）
        parts = self._parse_mixed_content(text)
        
        if not parts:
            return
        
        cell = worksheet[cell_addr]
        self._apply_filled_background(cell)

        # 如果只有纯文本，没有表格，则使用原有的单单元格填充方式
        has_tables = any(part["type"] == "table" for part in parts)
        
        if not has_tables:
            cell.value = text
            old_alignment = cell.alignment
            cell.alignment = Alignment(
                horizontal=old_alignment.horizontal or "left",
                vertical=old_alignment.vertical or "top",
                indent=1,
                wrap_text=True,
            )
            self._adjust_row_height_for_text(worksheet, cell.row, cell.column)
            return

        # 处理混合内容（文本 + 表格）
        self._fill_mixed_content_with_tables(worksheet, cell, parts)

    def _fill_mixed_content_with_tables(
        self, worksheet, first_cell, parts: List[Dict[str, Any]]
    ) -> None:
        """
        填充混合内容（文本+表格）
        
        结构：
        - 第一个文本写在 C44 单元格（取消原有合并，重新合并 C-R 列）
        - 第一个表格从 C46 开始插入
        - 后续文本在表格后空两行再插入（C-R 列合并）
        - 后续表格继续按顺序插入
        """
        # 先取消 C44 相关的合并，然后重新合并 C-R 列
        self._unmerge_cells_in_range(worksheet, first_cell.row, first_cell.row, 3, 18)
        try:
            worksheet.unmerge_cells(f"C{first_cell.row}:R{first_cell.row}")
        except (ValueError, KeyError):
            pass
        worksheet.merge_cells(f"C{first_cell.row}:R{first_cell.row}")
        
        current_row = first_cell.row
        
        for part in parts:
            if part["type"] == "text":
                if part["content"].strip():
                    if current_row == first_cell.row:
                        # 第一个文本直接写入 C44
                        first_cell.value = part["content"]
                        old_alignment = first_cell.alignment
                        first_cell.alignment = Alignment(
                            horizontal=old_alignment.horizontal or "left",
                            vertical=old_alignment.vertical or "top",
                            indent=1,
                            wrap_text=True,
                        )
                        self._adjust_row_height_for_text(worksheet, current_row, first_cell.column)
                    else:
                        # 后续文本：从当前行开始，跨越 C-R 列合并
                        start_row = current_row + 2  # 空两行
                        cell_range = f"C{start_row}:R{start_row}"
                        self._merge_and_set_text(worksheet, cell_range, part["content"])
                        current_row = start_row
                        self._adjust_row_height_for_text(worksheet, current_row, 3)  # C列
                current_row += 1  # 移动到下一行
                
            elif part["type"] == "table":
                # 插入 Excel 表格：从当前行+2开始插入（间隔两行）
                table_start_row = current_row + 2
                self._insert_excel_table(worksheet, table_start_row, part["content"])
                # 计算表格占用的行数
                table_rows_count = self._get_table_row_count(part["content"])
                current_row = table_start_row + table_rows_count

    def _merge_and_set_text(self, worksheet, cell_range: str, text: str) -> None:
        """
        合并单元格区域并设置文本
        
        Args:
            worksheet: 工作表对象
            cell_range: 单元格范围，如 "C5:R5"
            text: 要设置的文本
        """
        # 取消可能存在的合并
        try:
            worksheet.unmerge_cells(cell_range)
        except (ValueError, KeyError):
            pass
        
        # 合并单元格
        worksheet.merge_cells(cell_range)
        
        # 获取合并后的单元格（左上角）
        cell = worksheet[cell_range.split(":")[0]]
        cell.value = text
        self._apply_filled_background(cell)
        
        # 设置对齐
        cell.alignment = Alignment(
            horizontal="left",
            vertical="top",
            indent=1,
            wrap_text=True,
        )
        
        # 获取合并区域的起始行
        start_row = int(cell_range.split(":")[0][1:])
        worksheet.row_dimensions[start_row].height = 30  # 设置默认行高

    def _insert_excel_table(self, worksheet, start_row: int, markdown_text: str) -> None:
        """
        在指定位置插入 Excel 表格
        
        Args:
            worksheet: 工作表对象
            start_row: 起始行号
            markdown_text: markdown 格式的表格文本
        """
        headers, rows = self._parse_markdown_table(markdown_text)
        if not headers:
            return
        
        # 表格列数
        num_cols = len(headers)
        
        # 确保列数不超过可用列数（C到R = 18列）
        max_cols = 18
        if num_cols > max_cols:
            num_cols = max_cols
            headers = headers[:max_cols]
        
        # 计算结束行
        end_row = start_row + len(rows)  # 表头 + 数据行
        
        # 先取消目标区域的合并单元格（C-R 列，从 start_row 到 end_row）
        self._unmerge_cells_in_range(worksheet, start_row, end_row, 3, 18)
        
        # 设置边框样式
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 写入表头
        for col_idx, header in enumerate(headers, start=3):  # 从C列开始
            cell = worksheet.cell(start_row, col_idx)
            cell.value = header
            cell.font = Font(bold=True)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # 写入数据行
        for row_idx, row_data in enumerate(rows, start=1):
            for col_idx, header in enumerate(headers, start=3):
                cell = worksheet.cell(start_row + row_idx, col_idx)
                val = row_data[col_idx - 3] if col_idx - 3 < len(row_data) else ""
                cell.value = val
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        
        # 设置表格行高
        header_row_height = 20
        data_row_height = 18
        worksheet.row_dimensions[start_row].height = header_row_height
        for row_idx in range(1, len(rows) + 1):
            worksheet.row_dimensions[start_row + row_idx].height = data_row_height
        
        # 应用高亮背景（只应用到有效列）
        for row_idx in range(len(rows) + 1):
            for col_idx in range(3, 3 + num_cols):
                cell = worksheet.cell(start_row + row_idx, col_idx)
                self._apply_filled_background(cell)
        for row_idx in range(len(rows) + 1):
            for col_idx in range(3, 3 + num_cols):
                cell = worksheet.cell(start_row + row_idx, col_idx)
                self._apply_filled_background(cell)

    def _unmerge_cells_in_range(
        self, worksheet, start_row: int, end_row: int, start_col: int, end_col: int
    ) -> None:
        """
        取消指定行范围内涉及特定列区域的合并单元格
        
        Args:
            worksheet: 工作表对象
            start_row: 起始行
            end_row: 结束行
            start_col: 起始列
            end_col: 结束列
        """
        # 复制一份 ranges，避免迭代时修改集合
        merged_ranges_to_remove = []
        for merged_range in list(worksheet.merged_cells.ranges):
            # 检查合并区域是否与目标区域重叠
            if (merged_range.min_row <= end_row and merged_range.max_row >= start_row and
                merged_range.min_col <= end_col and merged_range.max_col >= start_col):
                merged_ranges_to_remove.append(str(merged_range))
        
        # 取消合并
        for range_str in merged_ranges_to_remove:
            try:
                worksheet.unmerge_cells(range_str)
            except (ValueError, KeyError):
                pass

    def _get_table_row_count(self, markdown_text: str) -> int:
        """获取 markdown 表格的行数（不含分隔符行）"""
        headers, rows = self._parse_markdown_table(markdown_text)
        return 1 + len(rows)  # 表头行 + 数据行

    def _parse_mixed_content(self, text: str) -> List[Dict[str, Any]]:
        """
        解析混合的 markdown 内容，识别文本段落和表格部分
        
        返回一个列表，每个元素是一个字典：
        - type: 'text' 或 'table'
        - content: 对应的内容（文本字符串或表格的 markdown 字符串）
        """
        if not text:
            return []
        
        lines = text.splitlines()
        parts = []
        current_text_lines = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            line_stripped = line.strip()
            
            # 检查从当前行开始是否是表格的开始
            # 表格需要至少2行：表头行和分隔符行
            if i + 1 < len(lines):
                next_line_stripped = lines[i + 1].strip()
                # 检查是否是表格的开始（表头行和分隔符行）
                if ("|" in line_stripped and "|" in next_line_stripped and 
                    re.search(r'[-:]+', next_line_stripped)):
                    # 先保存之前的文本内容
                    if current_text_lines:
                        text_content = "\n".join(current_text_lines)
                        if text_content.strip():
                            parts.append({"type": "text", "content": text_content})
                        current_text_lines = []
                    
                    # 收集表格的所有行
                    table_lines = [line]
                    i += 1
                    table_lines.append(lines[i])
                    i += 1
                    
                    # 继续收集表格的数据行（直到遇到非表格行或空行）
                    while i < len(lines):
                        current_line = lines[i]
                        current_line_stripped = current_line.strip()
                        
                        # 如果遇到空行，检查下一行是否是表格的继续
                        if not current_line_stripped:
                            if i + 1 < len(lines):
                                next_line_stripped = lines[i + 1].strip()
                                if "|" in next_line_stripped:
                                    # 下一行是表格行，空行也是表格的一部分
                                    table_lines.append(current_line)
                                    i += 1
                                    continue
                                else:
                                    break
                            else:
                                break
                        
                        # 如果当前行包含 |，可能是表格的数据行
                        if "|" in current_line_stripped:
                            table_lines.append(current_line)
                            i += 1
                        else:
                            break
                    
                    # 保存表格内容
                    table_content = "\n".join(table_lines)
                    if self._is_markdown_table(table_content):
                        parts.append({"type": "table", "content": table_content})
                    else:
                        # 如果不是有效的表格，将其作为文本处理
                        current_text_lines.extend(table_lines)
                    continue
            
            # 不是表格，作为文本处理
            current_text_lines.append(line)
            i += 1
        
        # 保存最后的文本内容
        if current_text_lines:
            text_content = "\n".join(current_text_lines)
            if text_content.strip():
                parts.append({"type": "text", "content": text_content})
        
        return parts

    def _is_markdown_table(self, markdown_text: str) -> bool:
        """
        检测文本是否是 markdown 表格格式
        
        要求：
        - 至少包含2行（表头和分隔符）
        - 表头和分隔符都包含 | 分隔符
        - 分隔符行包含 - 或 : 用于对齐
        """
        if not markdown_text:
            return False
        
        lines = [line.strip() for line in markdown_text.splitlines() if line.strip()]
        if len(lines) < 2:
            return False
        
        header_line = lines[0]
        separator_line = lines[1]
        
        # 必须包含 | 分隔符
        if "|" not in header_line or "|" not in separator_line:
            return False
        
        # 表头行应该包含至少一个 |
        header_pipes = header_line.count("|")
        if header_pipes < 2:
            return False
        
        # 分隔符行应该包含 - 或 : 用于对齐
        if not re.search(r'[-:]+', separator_line):
            return False
        
        # 验证分隔符行的格式
        separator_clean = separator_line.replace("|", "").replace(" ", "")
        if not re.search(r'[-:]{2,}', separator_clean):
            return False
        
        return True

    def _parse_markdown_table(self, markdown_text: str) -> Tuple[List[str], List[List[str]]]:
        """
        解析 markdown 表格文本为表头和数据行
        
        Args:
            markdown_text: markdown 格式的表格文本
            
        Returns:
            Tuple[List[str], List[List[str]]]: (表头列表, 数据行列表)
        """
        lines = [line.strip() for line in markdown_text.splitlines() if line.strip()]
        if len(lines) < 2:
            return [], []

        header_line = lines[0]
        separator_line = lines[1]
        if "|" not in header_line or "|" not in separator_line:
            return [], []

        headers = [cell.strip() for cell in header_line.strip("|").split("|")]
        rows: List[List[str]] = []
        for line in lines[2:]:
            if "|" not in line:
                continue
            cells = [cell.strip() for cell in line.strip("|").split("|")]
            # 对齐列数
            if len(cells) < len(headers):
                cells.extend([""] * (len(headers) - len(cells)))
            elif len(cells) > len(headers):
                cells = cells[: len(headers)]
            rows.append(cells)

        return headers, rows

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
