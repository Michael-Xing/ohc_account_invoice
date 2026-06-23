"""项目计划书（Word）填充器"""

from copy import deepcopy
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from src.infrastructure.template_service import TemplateFillerStrategy
from src.infrastructure.word_image_utils import (
    clear_cell,
    pure_image_urls,
    replace_placeholder_with_images,
)

logger = logging.getLogger(__name__)


class ProjectPlanFiller(TemplateFillerStrategy):
    """项目计划书专用填充器（docx）"""

    MARKDOWN_TABLE_FIELDS = {"function_module"}

    def fill_template(self, template_path: Path, parameters: Dict[str, Any], output_path: Path, language: Optional[str] = None) -> bool:
        """填充项目计划书模板。

        整体流程：
        Step 1: markdown 混合字段（function_module 等）→ 文本 + 表格渲染
        Step 2: 图片 URL 字段 → 定位占位符后下载并插图（或下载失败兜底）
        Step 3: 其余字段 → 普通 {{key}} 文本替换（已处理图片字段除外）
        """
        self._set_language(language)
        non_empty_fields = [k for k, v in parameters.items() if v]
        logger.info("[ProjectPlanFiller] 填充字段: %s", non_empty_fields)
        try:
            doc = Document(template_path)

            # Step 1: markdown 字段（文本 + 表格，不走图片流程）
            self._process_markdown_table_fields(doc, parameters)

            # Step 2: 识别纯 URL 图片字段，替换模板占位符
            processed_image_keys = self._process_image_placeholders(doc, parameters)

            # Step 3: 其余占位符文本替换（exclude 已插图字段，避免 URL 被当文字写入）
            self._fallback_text_replace(
                doc, self._flatten_parameters(parameters, exclude=processed_image_keys)
            )

            doc.save(output_path)
            return True
        except Exception as e:
            logger.error("项目计划书模板填充失败: %s", str(e), exc_info=True)
            return False

    # ------------------------------------------------------------------
    # Markdown 混合内容字段处理（文本 + 表格）
    # ------------------------------------------------------------------

    def _process_markdown_table_fields(self, doc: Document, parameters: Dict[str, Any]) -> None:
        """处理 MARKDOWN_TABLE_FIELDS（如 function_module）：将 markdown 文本/表格渲染到 Word"""
        for field in self.MARKDOWN_TABLE_FIELDS:
            placeholder = f"{{{{{field}}}}}"
            md_text = str(parameters.get(field) or "").strip()
            if not md_text:
                continue

            # 拆分为 text / table 片段
            parts = self._parse_mixed_markdown(md_text)
            if not parts:
                continue

            # 优先在正文段落中定位占位符
            for paragraph in list(doc.paragraphs):
                if placeholder in paragraph.text:
                    parent = paragraph._element.getparent()
                    idx = list(parent).index(paragraph._element)
                    parent.remove(paragraph._element)
                    self._render_mixed_at_position(doc, parent, idx, parts)
                    break
            else:
                # 段落未命中则在表格单元格中查找
                for table in doc.tables:
                    found = False
                    for row in table.rows:
                        for cell in row.cells:
                            if placeholder in (cell.text or ""):
                                clear_cell(cell)
                                self._render_mixed_into_cell(doc, cell, parts)
                                found = True
                                break
                        if found:
                            break
                    if found:
                        break

    # ------------------------------------------------------------------
    # 图片占位符处理（先识别图片字段，再替换模板占位符）
    # ------------------------------------------------------------------

    def _process_image_placeholders(self, doc: Document, parameters: Dict[str, Any]) -> set:
        """筛出图片 URL 字段，在模板中定位 {{key}} 并插入图片，返回已成功处理的字段名。

        对每个参数字段：
        Step 1: 跳过 MARKDOWN_TABLE_FIELDS（如 function_module，走 markdown 渲染）
        Step 2: pure_image_urls 判断并解析纯 URL / URL 列表
        Step 3: replace_placeholder_with_images 定位占位符 → 下载 → 插图或兜底
        Step 4: 成功处理则加入 processed_keys，后续文本替换时跳过
        """
        processed_keys: set = set()
        for key, raw in self._build_param_lookup(parameters).items():
            if key in self.MARKDOWN_TABLE_FIELDS:
                continue
            urls = pure_image_urls(raw)
            if not urls:
                continue
            placeholder = f"{{{{{key}}}}}"
            if replace_placeholder_with_images(
                doc,
                placeholder,
                urls,
                self._language,
                replace_in_paragraph_fn=self._replace_image_fallback_in_paragraph,
                replace_in_cell_fn=self._replace_image_fallback_in_cell,
                insert_block_paragraph_fn=self._insert_image_fallback_paragraph,
                insert_cell_paragraph_fn=self._insert_image_fallback_in_cell,
            ):
                processed_keys.add(key)
                logger.info("[ProjectPlanFiller] 字段 %s 识别为图片，已处理 %d 个URL", key, len(urls))
        return processed_keys

    def _replace_image_fallback_in_paragraph(self, paragraph, placeholder: str, text: str) -> None:
        """图片下载失败兜底：与普通字段填充相同的字体 + 蓝色样式"""
        key = placeholder[2:-2]
        self._replace_in_paragraph(paragraph, {key: text})

    def _replace_image_fallback_in_cell(self, cell, placeholder: str, text: str) -> None:
        """图片下载失败兜底（单元格）：与普通字段填充相同的字体 + 蓝色样式"""
        for para in cell.paragraphs:
            if placeholder in (para.text or ""):
                self._replace_image_fallback_in_paragraph(para, placeholder, text)
                return
        if placeholder in (cell.text or ""):
            self._set_cell_value(cell, cell.text.replace(placeholder, text))

    def _insert_image_fallback_paragraph(
        self, doc: Document, parent, insert_idx: int, text: str, reference_paragraph=None
    ) -> None:
        """部分 URL 下载失败：在图片后追加与普通填充值同风格的兜底段落"""
        paragraph = doc.add_paragraph()
        parent.insert(insert_idx, paragraph._element)
        self._write_image_fallback_text(paragraph, text, reference_paragraph)

    def _insert_image_fallback_in_cell(self, cell, text: str, reference_paragraph=None) -> None:
        """部分 URL 下载失败：在单元格图片后追加与普通填充值同风格的兜底段落"""
        self._write_image_fallback_text(cell.add_paragraph(), text, reference_paragraph)

    def _write_image_fallback_text(self, paragraph, text: str, reference_paragraph=None) -> None:
        font = {}
        if reference_paragraph is not None and reference_paragraph.runs:
            font = self._extract_run_font_from_run(reference_paragraph.runs[0])
        self._write_filled_text_in_paragraph(paragraph, text, font)

    def _write_filled_text_in_paragraph(self, paragraph, text: str, font: Optional[Dict] = None) -> None:
        """将整段文字按填充值样式写入（继承字体 + 蓝色）"""
        for run in list(paragraph.runs):
            run.clear()
        new_run = paragraph.add_run(text)
        self._apply_font(new_run, font or {})
        new_run.font.color.rgb = RGBColor(115, 159, 215)

    def _build_param_lookup(self, parameters: Dict[str, Any]) -> Dict[str, Any]:
        """将参数展平为一层 dict，供占位符 key 查找（支持一层嵌套）"""
        lookup: Dict[str, Any] = {}
        for key, value in parameters.items():
            if isinstance(value, dict):
                for sub_key, sub_val in value.items():
                    lookup[sub_key] = sub_val
            else:
                lookup[key] = value
        return lookup

    # ------------------------------------------------------------------
    # 混合 Markdown 解析
    # ------------------------------------------------------------------

    def _parse_mixed_markdown(self, markdown_text: str) -> List[Dict[str, Any]]:
        """将 markdown 文本拆分为 text / table 片段列表"""
        if not markdown_text:
            return []

        lines = markdown_text.splitlines()
        parts: List[Dict[str, Any]] = []
        text_buf: List[str] = []
        i = 0

        while i < len(lines):
            line_stripped = lines[i].strip()

            # 识别 markdown 表格：表头行 + 分隔符行（|---|）
            if i + 1 < len(lines):
                next_stripped = lines[i + 1].strip()
                if ("|" in line_stripped and "|" in next_stripped
                        and re.search(r'[-:]+', next_stripped)):
                    # 表格前的文本先入缓冲
                    if text_buf:
                        joined = "\n".join(text_buf)
                        if joined.strip():
                            parts.append({"type": "text", "content": joined})
                        text_buf = []

                    # 收集表格所有行
                    table_lines = [lines[i], lines[i + 1]]
                    i += 2
                    while i < len(lines):
                        cur_stripped = lines[i].strip()
                        if not cur_stripped:
                            if i + 1 < len(lines) and "|" in lines[i + 1].strip():
                                table_lines.append(lines[i])
                                i += 1
                                continue
                            break
                        if "|" in cur_stripped:
                            table_lines.append(lines[i])
                            i += 1
                        else:
                            break

                    parts.append({"type": "table", "content": "\n".join(table_lines)})
                    continue

            text_buf.append(lines[i])
            i += 1

        # 收尾：剩余文本作为最后一个 text 片段
        if text_buf:
            joined = "\n".join(text_buf)
            if joined.strip():
                parts.append({"type": "text", "content": joined})

        return parts

    # ------------------------------------------------------------------
    # 混合内容渲染
    # ------------------------------------------------------------------

    def _render_mixed_at_position(self, doc: Document, parent, insert_idx: int,
                                  parts: List[Dict[str, Any]]) -> None:
        """在文档块级位置依次渲染 text/table 片段"""
        cur = insert_idx
        for part in parts:
            if part["type"] == "table":
                headers, rows = self._parse_markdown_table(part["content"])
                if headers:
                    self._insert_md_table_at(doc, parent, cur, headers, rows)
                    cur += 1
            else:
                # 文本片段：逐段插入并推进索引
                elements = self._render_text_to_elements(doc, part["content"])
                for el in elements:
                    parent.insert(cur, el)
                    cur += 1

    def _render_mixed_into_cell(self, doc: Document, cell,
                                parts: List[Dict[str, Any]]) -> None:
        """在单元格内依次渲染 text/table 片段"""
        for part in parts:
            if part["type"] == "table":
                headers, rows = self._parse_markdown_table(part["content"])
                if headers:
                    self._insert_md_table_into_cell(cell, headers, rows)
            else:
                for raw_line in part["content"].splitlines():
                    line = raw_line.strip()
                    if not line:
                        continue
                    p = cell.add_paragraph("")
                    self._append_markdown_inline(p, line)

    # ------------------------------------------------------------------
    # Markdown 表格解析与插入
    # ------------------------------------------------------------------

    def _parse_markdown_table(self, markdown_text: str) -> Tuple[List[str], List[List[str]]]:
        lines = [line.strip() for line in markdown_text.splitlines() if line.strip()]
        if len(lines) < 2:
            return [], []
        # 第 1 行表头，第 2 行分隔符
        header_line = lines[0]
        separator_line = lines[1]
        if "|" not in header_line or "|" not in separator_line:
            return [], []
        headers = [c.strip() for c in header_line.strip("|").split("|")]
        rows: List[List[str]] = []
        # 第 3 行起为数据行
        for line in lines[2:]:
            if "|" not in line:
                continue
            cells = [c.strip() for c in line.strip("|").split("|")]
            if len(cells) < len(headers):
                cells.extend([""] * (len(headers) - len(cells)))
            elif len(cells) > len(headers):
                cells = cells[:len(headers)]
            rows.append(cells)
        return headers, rows

    def _insert_md_table_at(self, doc: Document, parent, insert_idx: int,
                            headers: List[str], rows: List[List[str]]) -> None:
        tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
        # 表头行
        for i, h in enumerate(headers):
            run = tbl.rows[0].cells[i].paragraphs[0].add_run(h)
            self._apply_table_font(run, bold=True)
        self._apply_header_row_style(tbl.rows[0])
        # 数据行
        for r_idx, vals in enumerate(rows, start=1):
            for c_idx, val in enumerate(vals):
                self._append_markdown_inline(tbl.rows[r_idx].cells[c_idx].paragraphs[0], val)
        self._apply_table_border(tbl)
        parent.insert(insert_idx, tbl._element)

    def _insert_md_table_into_cell(self, cell,
                                   headers: List[str], rows: List[List[str]]) -> None:
        tbl = cell.add_table(rows=1 + len(rows), cols=len(headers))
        for i, h in enumerate(headers):
            run = tbl.rows[0].cells[i].paragraphs[0].add_run(h)
            self._apply_table_font(run, bold=True)
        self._apply_header_row_style(tbl.rows[0])
        for r_idx, vals in enumerate(rows, start=1):
            for c_idx, val in enumerate(vals):
                self._append_markdown_inline(tbl.rows[r_idx].cells[c_idx].paragraphs[0], val)
        self._apply_table_border(tbl)

    # ------------------------------------------------------------------
    # 文本渲染
    # ------------------------------------------------------------------

    def _render_text_to_elements(self, doc: Document, text: str) -> List:
        """将文本内容渲染为段落元素列表，支持标题/列表/加粗"""
        elements = []
        for raw_line in text.splitlines():
            line = raw_line.rstrip()
            if not line:
                p = doc.add_paragraph("")
                elements.append(p._element)
                continue

            if line.startswith("#"):
                level = len(line) - len(line.lstrip("#"))
                content = line[level:].strip()
                p = doc.add_paragraph("")
                style_name = f"Heading {level}"
                if style_name in doc.styles:
                    p.style = doc.styles[style_name]
                self._append_markdown_inline(p, content)
                elements.append(p._element)
                continue

            if re.match(r"^[-*]\s+", line):
                content = re.sub(r"^[-*]\s+", "", line)
                p = doc.add_paragraph(style="List Bullet" if "List Bullet" in doc.styles else None)
                self._append_markdown_inline(p, content)
                elements.append(p._element)
                continue

            if re.match(r"^\d+\.\s+", line):
                content = re.sub(r"^\d+\.\s+", "", line)
                p = doc.add_paragraph(style="List Number" if "List Number" in doc.styles else None)
                self._append_markdown_inline(p, content)
                elements.append(p._element)
                continue

            p = doc.add_paragraph("")
            self._append_markdown_inline(p, line)
            elements.append(p._element)

        return elements

    def _append_markdown_inline(self, paragraph, text: str) -> None:
        """处理行内 markdown（**加粗**）"""
        pos = 0
        for match in re.finditer(r"\*\*(.+?)\*\*", text):
            if match.start() > pos:
                run = paragraph.add_run(text[pos:match.start()])
                self._apply_table_font(run)
            run = paragraph.add_run(match.group(1))
            self._apply_table_font(run, bold=True)
            pos = match.end()
        if pos < len(text):
            run = paragraph.add_run(text[pos:])
            self._apply_table_font(run)

    # ------------------------------------------------------------------
    # 单元格/样式工具
    # ------------------------------------------------------------------

    def _apply_table_font(self, run, bold: bool = False) -> None:
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(115, 159, 215)
        run.font.bold = bold

    def _apply_header_row_style(self, row) -> None:
        for cell in row.cells:
            shading = OxmlElement("w:shd")
            shading.set(qn("w:fill"), "D9D9D9")
            cell._tc.get_or_add_tcPr().append(shading)

    def _apply_table_border(self, table) -> None:
        tbl_pr = table._element.tblPr
        if tbl_pr is None:
            tbl_pr = OxmlElement("w:tblPr")
            table._element.insert(0, tbl_pr)
        borders = OxmlElement("w:tblBorders")
        for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b = OxmlElement(f"w:{name}")
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), "4")
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), "000000")
            borders.append(b)
        tbl_pr.append(borders)

    # ------------------------------------------------------------------
    # 兜底占位符替换（正文 + 表格 + 页眉页脚）
    # ------------------------------------------------------------------

    def _fallback_text_replace(self, doc: Document, flat_parameters: Dict[str, str]) -> None:
        # 正文
        for paragraph in doc.paragraphs:
            self._replace_in_paragraph(paragraph, flat_parameters)

        # 表格单元格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, flat_parameters)

        # 页眉 / 页脚
        for section in doc.sections:
            if section.header:
                for paragraph in section.header.paragraphs:
                    self._replace_in_paragraph(paragraph, flat_parameters)
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    self._replace_in_paragraph(paragraph, flat_parameters)

    def _replace_in_paragraph(self, paragraph, flat_parameters: Dict[str, str]) -> None:
        """在段落中替换占位符，拆分 run 使得只有填充值着色，前后缀保持原样。

        无论占位符在单个 run 内还是跨 run，都按"前缀（原色）+ 值（蓝色）+ 后缀（原色）"拆分。
        """
        full_text = paragraph.text or ""
        if not full_text:
            return

        ideal_text = full_text
        for key, value in flat_parameters.items():
            ideal_text = ideal_text.replace(f"{{{{{key}}}}}", value)

        if ideal_text == full_text:
            return

        placeholder_re = re.compile(
            r"\{\{(" + "|".join(re.escape(k) for k in flat_parameters) + r")\}\}"
        )

        # 1. 按占位符拆分：普通文字 / 填充值
        parts = placeholder_re.split(full_text)
        segments: List[Tuple[str, bool]] = []
        for i, part in enumerate(parts):
            if i % 2 == 0:
                if part:
                    segments.append((part, False))
            else:
                value = flat_parameters.get(part, "")
                if value:
                    segments.append((value, True))

        font = self._extract_run_font_from_run(paragraph.runs[0]) if paragraph.runs else {}

        # 2. 清空原 runs，按段重建（填充值着蓝色）
        for run in list(paragraph.runs):
            run.clear()

        for text, is_filled in segments:
            new_run = paragraph.add_run(text)
            self._apply_font(new_run, font)
            if is_filled:
                new_run.font.color.rgb = RGBColor(115, 159, 215)

    # ------------------------------------------------------------------
    # 单元格写值
    # ------------------------------------------------------------------

    def _set_cell_value(self, cell, text: str, style: Dict = None) -> None:
        """写入单元格文本：清空所有 run 后用 add_run 写入，并恢复字体 + 段落对齐"""
        paragraph = cell.paragraphs[0] if cell.paragraphs else None
        if paragraph and paragraph.runs:
            if not style:
                style = self._extract_cell_style(cell)
            for run in list(paragraph.runs):
                run.clear()
            new_run = paragraph.add_run(text)
            self._apply_font(new_run, style)
            new_run.font.color.rgb = RGBColor(115, 159, 215)
            if style.get("alignment") is not None:
                paragraph.alignment = style["alignment"]
        else:
            cell.text = text
            p = cell.paragraphs[0] if cell.paragraphs else None
            if p:
                if style and style.get("alignment") is not None:
                    p.alignment = style["alignment"]
                if p.runs:
                    self._apply_font(p.runs[0], style)
                    p.runs[0].font.color.rgb = RGBColor(115, 159, 215)

    # ------------------------------------------------------------------
    # 行操作
    # ------------------------------------------------------------------

    def _clone_last_row(self, table) -> None:
        new_tr = deepcopy(table.rows[-1]._tr)
        table._tbl.append(new_tr)
        for cell in table.rows[-1].cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.text = ""

    # ------------------------------------------------------------------
    # 字体工具
    # ------------------------------------------------------------------

    def _extract_cell_style(self, cell) -> Dict:
        """从单元格首段提取字体属性 + 段落对齐方式"""
        if not cell.paragraphs:
            return {}
        paragraph = cell.paragraphs[0]
        style: Dict[str, Any] = {}
        style["alignment"] = paragraph.alignment
        if paragraph.runs:
            style.update(self._extract_run_font_from_run(paragraph.runs[0]))
        return style

    def _extract_run_font_from_run(self, run) -> Dict:
        props: Dict[str, Any] = {}
        if run.font.name:
            props["name"] = run.font.name
        if run.font.size:
            props["size"] = run.font.size
        props["bold"] = run.font.bold
        props["italic"] = run.font.italic
        if run.font.color and run.font.color.rgb is not None:
            props["color_rgb"] = run.font.color.rgb
        elif run.font.color and run.font.color.theme_color is not None:
            props["color_theme"] = run.font.color.theme_color
        return props

    def _apply_font(self, run, font: Dict = None) -> None:
        """统一设置 run 字体，并尽量恢复原有显式颜色"""
        if not font:
            return
        if font.get("name"):
            run.font.name = font["name"]
            run._element.rPr.rFonts.set(qn("w:eastAsia"), font["name"])
        if font.get("size"):
            run.font.size = font["size"]
        if font.get("bold") is not None:
            run.font.bold = bool(font["bold"])
        if font.get("italic") is not None:
            run.font.italic = bool(font["italic"])
        if font.get("color_rgb") is not None:
            run.font.color.rgb = font["color_rgb"]
        elif font.get("color_theme") is not None:
            run.font.color.theme_color = font["color_theme"]

    # ------------------------------------------------------------------
    # 参数工具
    # ------------------------------------------------------------------

    def _flatten_parameters(
        self, parameters: Dict[str, Any], exclude: Optional[set] = None
    ) -> Dict[str, str]:
        skip = exclude or set()
        flat: Dict[str, str] = {}
        for key, value in parameters.items():
            # markdown 混合字段已在前面单独处理
            if key in self.MARKDOWN_TABLE_FIELDS or key in skip:
                continue
            # 一层嵌套对象：子键展平到顶层
            if isinstance(value, dict):
                for sub_key, sub_val in value.items():
                    if sub_key in skip:
                        continue
                    if isinstance(sub_val, (str, int, float)):
                        if isinstance(sub_val, str) and sub_val.strip() == "":
                            flat[sub_key] = self._missing_text()
                        else:
                            flat[sub_key] = str(sub_val)
                    elif sub_val is None:
                        flat[sub_key] = self._missing_text()
            elif isinstance(value, (str, int, float)):
                if isinstance(value, str) and value.strip() == "":
                    flat[key] = self._missing_text()
                else:
                    flat[key] = str(value)
            elif value is None:
                flat[key] = self._missing_text()
        return flat
