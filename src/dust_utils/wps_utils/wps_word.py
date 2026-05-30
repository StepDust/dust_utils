from .wps_base import WPSBase, auto_before_call
from .wps_utils import WPSUtils

import os
import pywintypes
import tempfile
import shutil
from typing import List, Any, Optional
from PyPDF2 import PdfReader

from loguru import logger

# Word 常量定义
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_HEADER_FOOTER_PRIMARY = 1
WD_FIELD_EMPTY = -1
WD_FORMAT_PDF = 17
WD_PREFERRED_WIDTH_POINTS = 2


@auto_before_call(before_func="available")
class WPSWord(WPSBase):
    """
    Word/WPS文档操作类，基于COM接口实现。
    提供表格操作、段落操作、页面设置等功能。
    """

    def __init__(
        self,
        file_path: str,
        family_name: str = None,
        use_wps: bool = True,
        is_debug: bool = False,
    ):
        """
        初始化 WPSWord 类，打开指定 Word 文件。

        参数:
        - file_path: str, Word 文件路径
        - family_name: str = None, 如果是创建文件，则生效的默认字体
        - use_wps: bool = True, True 表示使用 WPS Office，False 表示使用 MS Word
        - is_debug: bool = False, 是否进入调试模式
        """
        # 确定使用 WPS 还是 Word
        prog_id = "KWps.Application" if use_wps else "Word.Application"
        # 设置 logger 级别
        logger.setLevel(logging.DEBUG if is_debug else logging.INFO)
        super().__init__(file_path, prog_id, is_debug)
        try:
            self.family_name = family_name
            # 等待 COM 对象完全就绪
            if not hasattr(self.office, "Documents"):
                raise AttributeError("COM 对象尚未准备好，缺少 Documents 属性")
            if not os.path.exists(self.file_path):
                logger.debug(f"文件 {self.file_path} 不存在，创建新文件")
                self.word = self.office.Documents.Add()
                self.word.SaveAs(self.file_path)
            else:
                logger.debug(f"文件 {self.file_path} 已存在，打开文件")
                self.word = self.office.Documents.Open(self.file_path)
        except AttributeError as ae:
            logger.error("COM 对象初始化失败: %s", str(ae))
            self.quit()
            raise RuntimeError(f"COM 对象初始化失败: {str(ae)}")
        except Exception as e:
            logger.error("无法打开或创建文件 %s: %s", self.file_path, str(e))
            self.quit()
            raise RuntimeError(f"无法打开或创建文件 {self.file_path}: {str(e)}")

    # region 表格操作

    def get_tables(self) -> List[Any]:
        """
        获取文档中的所有表格。

        返回:
        - List[Any], 包含所有表格对象的列表。
        """
        return self.word.Tables

    def insert_table(self, para, data):
        """
        将包含 {{附录数据.I7:K9}} 的段落替换为表格
        占位符不会出现在表格中
        """
        if not data or not isinstance(data, list) or not isinstance(data[0], list):
            raise ValueError("data 必须是二维列表")

        rows = len(data)
        cols = len(data[0])
        doc = para.Range.Document

        # 记录段落起始位置
        start_pos = para.Range.Start

        # 保存段落格式
        format_dict = self.save_format(para)

        # 清空段落内容(避免占位符进入表格第一个单元格)
        para.Range.Text = ""

        # 在原位置创建 Range
        rng = doc.Range(Start=start_pos, End=start_pos)

        # 插入表格
        table = doc.Tables.Add(rng, rows, cols)
        table.Style = None  # 取消表格样式

        # 填充数据
        for r in range(rows):
            for c in range(cols):
                table.Cell(r + 1, c + 1).Range.Text = str(data[r][c])

        # 自动调整表格
        try:
            table.AutoFitBehavior(1)  # 自动适应内容
        except Exception:
            pass

        try:
            table.AutoFitBehavior(0)  # 固定列宽
            table.PreferredWidthType = 1  # wdPreferredWidthPoints
            table.PreferredWidth = self.get_page_width()
        except Exception:
            pass

        # 恢复表格自身的格式（应用到整个表格）
        try:
            # 遍历表格所有单元格，将保存的段落格式应用到每个单元格的第一个段落
            for r in range(1, table.Rows.Count + 1):
                for c in range(1, table.Columns.Count + 1):
                    cell = table.Cell(r, c)
                    self.restore_format(cell, format_dict)
                    cell.VerticalAlignment = (
                        1  # 1 对应 wdCellAlignVerticalCenter，垂直居中
                    )

        except Exception as e:
            logger.warning(f"恢复表格格式时出错: {e}")

        return table

    @staticmethod
    def get_real_col_index(grid, row_index, col_index):
        """
        根据逻辑列号计算当前表格中的真实列索引（考虑已合并单元格偏移）
        """
        row = grid[row_index - 1]
        unique_ids = []
        for i in range(col_index):
            if row[i] not in unique_ids:
                unique_ids.append(row[i])
        return len(unique_ids)

    @staticmethod
    def update_grid_merge(grid, start_row, start_col, end_row, end_col):
        """
        在 grid 矩阵中标记合并区域
        - grid: 逻辑矩阵
        - start_row, start_col: 起始行列（1-based）
        - end_row, end_col: 结束行列（1-based）
        """
        base = grid[start_row - 1][start_col - 1]
        for r in range(start_row - 1, end_row):
            for c in range(start_col - 1, end_col):
                grid[r][c] = base

    def init_table_grid(self, table):
        """
        初始化表格的逻辑矩阵（用于合并单元格跟踪）
        - table: Word/WPS 表格对象
        """
        if not hasattr(self, "_table_index_map"):
            self._table_index_map = {}

        tid = id(table)
        total_rows = table.Rows.Count
        total_cols = table.Columns.Count
        # 初始化表格结构
        if tid not in self._table_index_map:
            index = 1
            self._table_index_map[tid] = [
                [index + c + r * total_cols for c in range(total_cols)]
                for r in range(total_rows)
            ]

        return self._table_index_map[tid]

    def table_merge_cells(self, table, row_index, col_index, rowspan=1, colspan=1):
        """
        合并表格中的单元格（支持行列合并与动态索引偏移）
        参数:
        - table: Word/WPS 表格对象
        - row_index, col_index: 起始单元格（从1开始）
        - rowspan, colspan: 合并范围
        - grid: 当前表格对应的二维逻辑矩阵（会自动更新）
        """
        try:
            total_rows = table.Rows.Count
            total_cols = table.Columns.Count
            grid = self.init_table_grid(table)
            logger.debug(f"[表格] 行数={total_rows}, 列数={total_cols}")

            if (
                row_index < 1
                or col_index < 1
                or row_index > total_rows
                or col_index > total_cols
            ):
                raise ValueError(f"指定单元格 ({row_index},{col_index}) 超出表格范围")

            # 计算逻辑合并范围
            end_row = min(row_index + rowspan - 1, total_rows)
            end_col = min(col_index + colspan - 1, total_cols)
            logger.debug(
                f"[请求合并] 起点=({row_index},{col_index}) → 终点=({end_row},{end_col})"
            )

            # 计算真实位置（考虑已合并单元格）
            real_start_col = self.get_real_col_index(grid, row_index, col_index)
            real_end_col = self.get_real_col_index(grid, end_row, end_col)

            logger.debug(
                f"[修正索引] 真实起点=({row_index},{real_start_col}) → 真实终点=({end_row},{real_end_col})"
            )

            # 获取起始与结束单元格（安全访问）
            start_cell = table.Cell(row_index, real_start_col)
            end_cell = table.Cell(end_row, real_end_col)

            logger.debug(
                f"[执行合并] 从 ({row_index},{real_start_col}) 到 ({end_row},{real_end_col})"
            )
            merged_cell = start_cell.Merge(end_cell)
            logger.debug(
                f"[完成合并] 行 {row_index}-{end_row}, 列 {real_start_col}-{real_end_col}"
            )

            # 自动更新 grid 矩阵
            self.update_grid_merge(grid, row_index, col_index, end_row, end_col)
            logger.debug(
                f"[更新grid] 已同步合并区域: ({row_index},{col_index}) → ({end_row},{end_col})"
            )

            # 清理空白段落
            return self.merged_cell_clear_blank_paragraph(merged_cell)

        except Exception as e:
            logger.error(f"[错误] 合并单元格失败: {e}", exc_info=True)
            return None

    def set_cell_border(
        self,
        table: Any,
        row_index: int,
        col_index: int,
        direction: list = ["top", "bottom", "left", "right"],
        border_type="Single",
        border_width=1,
    ):
        """
        设置单元格边框
        参数:
        - table: 表格对象
        - row_index: 行索引（从1开始）
        - col_index: 列索引（从1开始）
        - direction: 边框方向，可选"top"、"bottom"、"left"、"right"
        - border_type: 边框类型，可选"Single"（单边框）、"Double"（双边框）、"Dot"（点线）、"Dash"（短虚线）
        - border_width: 边框宽度，单位磅（默认1磅）
        """
        # Word 边框常量
        wdBorderTop = 1
        wdBorderBottom = 3
        wdBorderLeft = 2
        wdBorderRight = 4

        grid = self.init_table_grid(table)
        real_col = self.get_real_col_index(grid, row_index, col_index)
        cell = table.Cell(row_index, real_col)

        WD_LINE_STYLE = {
            "NONE": 0,  # wdLineStyleNone,无边框(修正为0)
            "SINGLE": 1,  # wdLineStyleSingle,单实线
            "DOUBLE": 7,  # wdLineStyleDouble,双线
            "DOT": 2,  # wdLineStyleDot,点线
            "DASH": 3,  # wdLineStyleDashSmallGap,短虚线
        }

        # 边框样式
        border_type = WD_LINE_STYLE.get(
            border_type.upper() if border_type else "NONE", None
        )
        if border_type == None:
            border_type = WD_LINE_STYLE["NONE"]

        # 循环遍历方向
        for d in direction:
            if d == "top":
                cell.Borders(wdBorderTop).LineStyle = border_type
                cell.Borders(wdBorderTop).LineWidth = border_width
            elif d == "bottom":
                cell.Borders(wdBorderBottom).LineStyle = border_type
                cell.Borders(wdBorderBottom).LineWidth = border_width
            elif d == "left":
                cell.Borders(wdBorderLeft).LineStyle = border_type
                cell.Borders(wdBorderLeft).LineWidth = border_width
            elif d == "right":
                cell.Borders(wdBorderRight).LineStyle = border_type
                cell.Borders(wdBorderRight).LineWidth = border_width

    def set_cell_alignment(
        self, table: Any, row_index: int, col_index: int, alignment="center"
    ) -> None:
        """
        设置单元格内容水平对齐方式

        参数:
        - table: 表格对象
        - row_index: 行索引（从1开始）
        - col_index: 列索引（从1开始）
        - alignment: 对齐方式，可选"center"（居中）、"left"（左对齐）、"right"（右对齐）
        """
        grid = self.init_table_grid(table)
        real_col = self.get_real_col_index(grid, row_index, col_index)

        alignment = alignment.lower().strip()
        alignment_map = {
            "center": 1,
            "left": 0,
            "right": 2,
        }
        if alignment not in alignment_map:
            raise ValueError(f"不支持的对齐方式: {alignment}")
        cell = table.Cell(row_index, real_col)
        # 1 对应 wdAlignParagraphCenter，水平居中
        cell.Range.ParagraphFormat.Alignment = alignment_map[alignment]

    def table_merge_cells_by_column(self, table: Any) -> None:
        """
        将表格中连续的空单元格合并到上一个非空单元格（逐列纵向合并）

        参数:
        - table: 表格对象
        """
        col_count = table.Columns.Count
        row_count = table.Rows.Count

        for col_idx in range(1, col_count + 1):
            anchor_cell = None  # 记录最近的非空单元格
            for row_idx in range(1, row_count + 1):
                cell = table.Cell(row_idx, col_idx)
                value = WPSUtils.remove_non_printable(cell.Range.Text.strip()).strip()
                if value:  # 非空
                    anchor_cell = cell
                else:  # 空单元格
                    if anchor_cell is not None:
                        try:
                            merged_cell = anchor_cell.Merge(cell)
                            # 删除合并后单元格中“完全空白”的段落
                            merged_cell = self.merged_cell_clear_blank_paragraph(
                                merged_cell
                            )
                        except ValueError:
                            # 已经被合并过的单元格跳过
                            continue

    def merged_cell_clear_blank_paragraph(self, cell: Any) -> Any:
        """
        仅删除单元格中"完全空白"的段落（保留首尾空格与所有格式）

        参数:
        - cell: 单元格对象

        返回:
        - 清理后的单元格对象
        """
        return cell
        # 防御式检查：确保 cell 及其 Range 有效
        if cell is None or getattr(cell, "Range", None) is None:
            logger.warning(
                "merged_cell_clear_blank_paragraph 接收到无效 cell，直接返回"
            )
            return cell

        # 从后往前遍历段落，避免索引错乱
        for i in range(cell.Range.Paragraphs.Count, 0, -1):
            try:
                para = cell.Range.Paragraphs(i)
            except Exception:
                # 段落索引失效时跳过
                continue
            # 判断段落是否“完全空白”
            if not para.Range.Text.strip():
                # 如果该段落是最后一个段落且删除后会导致单元格无段落，则保留
                if cell.Range.Paragraphs.Count == 1:
                    continue
                try:
                    para.Range.Delete()
                except Exception:
                    # 删除失败时忽略
                    pass
        return cell

    def set_table_row_height(self, table, height=25, rule=1):
        """
        设置 Word 表格的行高

        参数:
            table: 表格对象 (Word.Table)
            height: 行高
            rule: 行高规则
                1 = 最小行高 (wdRowHeightAtLeast)
                2 = 固定行高 (wdRowHeightExactly)
        """

        try:
            for r in range(1, table.Rows.Count + 1):
                row = table.Rows(r)
                row.HeightRule = rule  # 设置规则
                row.Height = height  # 设置行高
                # rule=2 时允许内容撑高
                if rule == 2:
                    row.AllowBreakAcrossPages = True
        except Exception as e:
            print(f"设置表格行高失败: {e}")

    def set_table_column_width(self, table, src_index, target_index, width=150):
        """
        在不影响table总宽度的情况下设置 Word 表格的列宽

        参数:
            table: 表格对象 (Word.Table)
            src_index: 源列索引，从1开始，target_index列宽增减的差额从该列宽度计算
            target_index: 目标列索引，从1开始，设置该列宽度后，src_index列宽度会自动调整
            width: 列宽，单位为磅（point）
        """

        # 获取目标列与源列对象
        target_col = table.Columns(target_index)
        src_col = table.Columns(src_index)

        # 记录原目标列宽度
        old_width = target_col.Width

        # 计算宽度差额
        delta = width - old_width

        # 先调整源列宽度，保持总宽不变
        src_col.PreferredWidthType = WD_PREFERRED_WIDTH_POINTS
        new_src_width = src_col.Width - delta
        if new_src_width < 0:
            new_src_width = 0

        # 调整源列宽度，确保不小于0
        src_col.Width = new_src_width

        # 再设置目标列新宽度
        target_col.Width = width

    # endregion

    # region 段落操作

    def get_paragraphs(self) -> List[Any]:
        """
        获取文档中的所有段落。

        返回:
        - List[Any], 包含所有段落对象的列表。
        """
        return self.word.Paragraphs

    def get_para_info(self, para: Any):
        """
        获取段落的详细信息。

        参数:
        - para: COM段落对象。

        返回:
        - Dict[str, Any], 包含段落信息的字典，键值对包括：
            - "text": 段落文本内容。
            - "style": 段落样式名称。
            - "alignment": 段落对齐方式（中文描述）。
            - "indent": 段落缩进值（磅）。
            - "space_before": 段落前间距（磅）。
            - "space_after": 段落后间距（磅）。
            - "bg_color_hex": 段落背景颜色（16进制）。
        """
        # 获取样式名称
        try:
            style_name = getattr(para.Style, "Name", None)
        except Exception:
            style_name = None

        # 获取缩进（Word/WPS 兼容）
        indent = getattr(para, "LeftIndent", None)
        if indent is None:
            indent = getattr(getattr(para, "Indentation", None), "Left", 0)
        if indent is None:
            indent = 0

        # 对齐方式映射
        alignment_map = {
            0: "左对齐",
            1: "居中",
            2: "右对齐",
            3: "两端对齐",
            4: "分散对齐",
        }
        alignment = alignment_map.get(getattr(para, "Alignment", -1), "未知")

        # 获取背景颜色（Shading.BackgroundPatternColor）
        try:
            bg_color_int = getattr(para.Range, "HighlightColorIndex", None)
            # 映射颜色索引到RGB值
            highlight_colors = {
                0: None,  # wdNoHighlight
                1: (0, 0, 0),  # wdBlack
                2: (0, 0, 255),  # wdBlue
                3: (0, 255, 255),  # wdTurquoise
                4: (0, 255, 0),  # wdBrightGreen
                5: (255, 0, 255),  # wdPink
                6: (255, 0, 0),  # wdRed
                7: (255, 255, 0),  # wdYellow
                8: (255, 255, 255),  # wdWhite
                9: (0, 0, 128),  # wdDarkBlue
                10: (0, 128, 128),  # wdTeal
                11: (0, 128, 0),  # wdGreen
                12: (128, 0, 128),  # wdViolet
                13: (128, 0, 0),  # wdDarkRed
                14: (128, 128, 0),  # wdDarkYellow
                15: (128, 128, 128),  # wdGray50
                16: (192, 192, 192),  # wdGray25
            }
        except Exception:
            bg_color_hex = None

        para_info = {
            "text": getattr(getattr(para, "Range", None), "Text", "").strip(),
            "style": style_name,
            "alignment": alignment,
            "indent": indent,
            "space_before": getattr(para, "SpaceBefore", 0),
            "space_after": getattr(para, "SpaceAfter", 0),
            "bg_color_hex": bg_color_hex,
        }

        logger.info(f"段落信息: {para_info}")
        return para_info

    def set_para_text(self, para, text: str, is_remove_empty=True):
        """
        安全替换段落或单元格文本，不破坏样式结构。

        参数：
        - para: 段落（或单元格中的段落）COM 对象
        - text: 要替换的新文本
        - is_remove_empty: 是否移除文本的左右空格
        """
        doc = para.Range.Document
        start = para.Range.Start
        end = para.Range.End

        # 去除不可见字符、左右空格等
        text = WPSUtils.remove_non_printable(text)
        if is_remove_empty:
            text = text.strip()

        # 删除段落内容，但保留段落结束符（¶）
        if end > start:
            doc.Range(start, end - 1).Delete()

        # 在段落开头插入新文本（保留原样式、run结构）
        insert_point = doc.Range(start, start)
        insert_point.InsertAfter(text)

        return para

    def copy_para(self, para, position="after"):
        """
        复制段落到指定位置，包括文本、格式和样式。

        参数：
        - para: 源段落（或单元格中的段落）COM 对象
        - position: 插入位置，可选值：
            - "before": 在该段落前插入
            - "after": 在该段落后插入（默认）
            - "end": 在文档末尾插入

        返回:
        - 新插入的段落对象
        """
        try:
            doc = para.Range.Document
            source_range = para.Range

            # 获取源段落的文本
            text = source_range.Text.rstrip("\r")  # 移除段落结束符

            # 确定插入位置
            if position == "before":
                # 在该段落前插入
                insert_range = source_range.Duplicate
                insert_range.Collapse(0)  # 折叠到开头
            elif position == "after":
                # 在该段落后插入
                insert_range = source_range.Duplicate
                insert_range.Collapse(1)  # 折叠到末尾
            elif position == "end":
                # 在文档末尾插入
                insert_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            else:
                raise ValueError(
                    f"不支持的位置参数: {position}，应为 'before'、'after' 或 'end'"
                )

            # 插入段落
            insert_range.InsertAfter("\r")
            new_para = insert_range.Paragraphs(1)

            # 将文本设置到新段落
            self.set_para_text(new_para, text, is_remove_empty=False)

            # 恢复格式到新段落
            self.restore_format(new_para.Range, fmt)

            logger.info(f"段落已复制到 {position} 位置")
            return new_para

        except Exception as e:
            logger.error(f"复制段落失败: {e}", exc_info=True)
            raise RuntimeError(f"复制段落失败: {e}")

    # endregion

    # region 页面操作

    def get_page_width(self) -> float:
        """
        获取当前页面的宽度（以磅为单位）。

        返回:
        - float, 页面宽度（磅）。
        """
        page_setup = self.word.PageSetup
        page_width = page_setup.PageWidth
        left_margin = page_setup.LeftMargin
        right_margin = page_setup.RightMargin
        usable_width = page_width - left_margin - right_margin
        return usable_width

    def set_page_start_number(self, start_number: Optional[int]) -> None:
        """
        设置页脚页码起始编号。

        参数:
        - start_number: int/None，页码起始值；传入 None 则清除页码与页脚。
        """
        # 获取第一个 section
        _section = self.word.Sections(1)
        footer = _section.Footers(1)  # wdHeaderFooterPrimary = 1
        page_numbers = footer.PageNumbers

        if start_number is None:
            # 清空页脚
            footer.Range.Delete()
            footer.LinkToPrevious = False
            try:
                page_numbers.RestartNumberingAtSection = False
            except (AttributeError, pywintypes.com_error) as e:
                logger.debug(f"PageNumbers.RestartNumberingAtSection 重置跳过: {e}")
            try:
                page_numbers.StartingNumber = 1
            except (AttributeError, pywintypes.com_error) as e:
                logger.debug(f"PageNumbers.StartingNumber 重置跳过: {e}")
            return

        # 设置页脚内容
        footer.LinkToPrevious = False
        footer_range = footer.Range
        footer_range.Delete()  # 先彻底清空

        # 设置页码属性
        try:
            page_numbers.NumberStyle = 0  # wdPageNumberStyleArabic = 0
        except (AttributeError, pywintypes.com_error) as e:
            logger.debug(f"PageNumbers.NumberStyle 跳过: {e}")
        try:
            page_numbers.IncludeChapterNumber = False
        except (AttributeError, pywintypes.com_error) as e:
            logger.debug(f"PageNumbers.IncludeChapterNumber 跳过: {e}")
        try:
            page_numbers.RestartNumberingAtSection = True
        except (AttributeError, pywintypes.com_error) as e:
            logger.debug(f"PageNumbers.RestartNumberingAtSection 跳过: {e}")
        try:
            page_numbers.StartingNumber = start_number
        except (AttributeError, pywintypes.com_error) as e:
            logger.debug(f"PageNumbers.StartingNumber 跳过: {e}")

        # 插入段落并居中
        footer_range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter = 1
        para = footer_range.Paragraphs.Add()

        # 插入左横线
        para.Range.Text = "- "
        para.Range.Font.Name = "宋体"
        para.Range.Font.Size = 9

        # 插入页码域（在段落末尾）
        end_range = para.Range.Duplicate
        end_range.Collapse(0)  # 折叠到末尾
        fld = footer_range.Fields.Add(end_range, -1, "PAGE", False)  # wdFieldEmpty = -1
        fld.Result.Font.Name = "宋体"
        fld.Result.Font.Size = 9

        # 插入右横线
        end_range = para.Range.Duplicate
        end_range.Collapse(0)  # 折叠到末尾
        end_range.InsertAfter(" -")
        end_range.Font.Name = "宋体"
        end_range.Font.Size = 9

        # 更新所有字段以应用变化
        self.word.Fields.Update()

        # 强制重分页以应用变化
        self.word.Repaginate()

    def get_last_page_number(self) -> int:
        """
        获取 Word 文档的最后显示页码：
        起始页码 + PDF总页数 - 1。兼容 Word 和 WPS，转换失败会抛出异常。

        Args:
            file_path: Word 文档路径（.docx）

        Returns:
            int: 文档最后一页显示的页码（估算）

        Raises:
            FileNotFoundError: 文件不存在
            RuntimeError: 转换或读取失败
        """

        # 1. 获取文档起始页码
        start_page = self.get_first_section_start()

        # 2. 在英文路径下用 win32com 转换 PDF
        with tempfile.TemporaryDirectory() as tmp_folder:
            try:
                # 拷贝到英文路径，避免中文路径导致 Word/WPS 崩溃
                temp_docx = os.path.join(tmp_folder, "input.docx")
                shutil.copyfile(self.file_path, temp_docx)

                # 转换路径
                pdf_path = os.path.join(tmp_folder, "output.pdf")

                # 调用 Word 或 WPS 转换
                self.convert_to_pdf(temp_docx, pdf_path)
                logger.debug("PDF 转换成功")

                # 3. 读取 PDF 页数
                try:
                    reader = PdfReader(pdf_path)
                    total_pages = len(reader.pages)
                    logger.debug(f"PDF 总页数: {total_pages}")
                except Exception as e:
                    logger.error(f"读取 PDF 页数失败: {str(e)}")
                    raise FileNotFoundError(f"读取 PDF 页数失败: {e}")

            except Exception as e:
                logger.error(f"PDF 转换失败: {str(e)}")
                raise RuntimeError(f"Word/WPS 转换 PDF 失败: {e}")

        # 4. 计算显示的最后一页码
        last_displayed_page = start_page + total_pages - 1
        logger.info(f"文档最后显示页码为: {last_displayed_page}")
        return last_displayed_page

    def get_first_section_start(self) -> int:
        """
        通过 COM 接口读取文档第一节的页码起始值，并判断后续节是否重置页码。

        参数:
            file_path: Word 文档路径

        返回:
            int: 文档第一节的起始页码

        异常:
            RuntimeError: COM 接口操作失败
        """
        try:

            first_section = self.word.Sections(1)
            restart = first_section.PageSetup.RestartNumberingAtSection
            start_num = first_section.PageSetup.StartingNumber

            return start_num if restart else 1
        except Exception as e:
            logger.error(f"通过 COM 读取页码失败: {e}")
            raise RuntimeError(f"通过 COM 读取页码失败: {e}")

    def convert_to_pdf(self, pdf_path: str) -> None:
        """
        使用 COM 接口将 docx 转换为 pdf。

        参数:
            docx_path: Word 文档路径
            pdf_path: 输出的 PDF 文件路径

        异常:
            RuntimeError: COM 接口操作失败
        """
        try:
            self.word.SaveAs(pdf_path, FileFormat=WD_FORMAT_PDF)
        except Exception as e:
            logger.error(f"COM 转换 PDF 失败: {e}")
            raise RuntimeError(f"COM 转换 PDF 失败: {e}")

    # endregion

    # region 文档状态管理

    def available(self) -> bool:
        """
        判断word对象是否可用，如果可用，则不做处理，不可用，则重新创建

        返回:
        - bool: 对象是否可用
        """
        try:
            # 尝试访问文档的Name属性，若抛出异常说明COM对象已失效
            _ = self.word.Name
            return True
        except Exception as e:
            logger.warning(f"Word COM 对象不可用，准备重新创建: {e}")
            # 释放旧对象
            self.word = None
            # 重新打开文档
            try:
                self.word = self.office.Documents.Open(self.file_path)
                logger.info("Word 文档已重新打开")
                return True
            except Exception as reopen_err:
                logger.error(f"重新创建文档失败: {reopen_err}")
                raise RuntimeError(f"无法重新创建文档: {reopen_err}")

    def save(
        self, output_path: Optional[str] = None, is_transfer: bool = False
    ) -> None:
        """
        保存 Word 文件。

        参数:
        - output_path: str, 可选，指定保存路径。如果未提供，将使用当前文件路径。
        - is_transfer: bool, 可选，是否将当前工作簿设为新打开的工作簿。默认值为 False。

        异常:
        - 如果保存失败，抛出 RuntimeError 异常。
        """
        save_path = (
            os.path.normpath(os.path.abspath(output_path))
            if output_path
            else self.file_path
        )
        # 确保保存目录存在
        save_folder = os.path.dirname(save_path)
        if save_folder and not os.path.exists(save_folder):
            os.makedirs(save_folder)
            logger.debug(f"已创建保存目录：{save_folder}")

        try:
            self.word.SaveAs(save_path)
            if is_transfer:
                self.word = self.office.Documents.Open(save_path)
                self.file_path = save_path
                logger.debug(f"当前处理文档更新为：{save_path}")
        except Exception as e:
            logger.error(f"保存Word失败 {save_path}: {str(e)}")
            raise RuntimeError(f"保存Word失败 {save_path}: {str(e)}")

    def close(self) -> None:
        """
        关闭 Word 文件。
        """
        logger.info(f"关闭 Word 文件：{self.file_path}")
        if self.word:
            try:
                # 检查COM对象是否有效
                # 可以通过尝试访问一个简单属性来测试
                _ = self.word.Name  # 如果对象无效，这里会抛出异常
                self.word.Close(SaveChanges=False)
            except:
                logger.warning(f"Word 文件 {self.file_path} 可能已关闭或无效")
            finally:
                self.word = None
        self.quit()

    # endregion

    def save_format(self, para):
        """
        保存段落格式，用于应用到单元格。
        """
        try:
            rng = para.Range
            fmt = {}

            font = getattr(rng, "Font", None)
            pfmt = (
                para.Format if hasattr(para, "Format") else para.Range.ParagraphFormat
            )

            # 字体属性
            fmt["font_name"] = getattr(font, "Name", None)
            fmt["font_size"] = getattr(font, "Size", None)
            fmt["font_bold"] = getattr(font, "Bold", None)
            fmt["font_italic"] = getattr(font, "Italic", None)
            fmt["font_underline"] = getattr(font, "Underline", None)
            fmt["font_color"] = getattr(font, "Color", None)

            # 高亮
            try:
                fmt["highlight"] = getattr(rng, "HighlightColorIndex", None)

                if fmt["highlight"] is not None and fmt["highlight"] != 0:
                    # 高亮颜色字典
                    highlight_colors = {
                        0: None,  # wdNoHighlight
                        1: (0, 0, 0),  # wdBlack
                        2: (0, 0, 255),  # wdBlue
                        3: (0, 255, 255),  # wdTurquoise
                        4: (0, 255, 0),  # wdBrightGreen
                        5: (255, 0, 255),  # wdPink
                        6: (255, 0, 0),  # wdRed
                        7: (255, 255, 0),  # wdYellow
                        8: (255, 255, 255),  # wdWhite
                        9: (0, 0, 128),  # wdDarkBlue
                        10: (0, 128, 128),  # wdTeal
                        11: (0, 128, 0),  # wdGreen
                        12: (128, 0, 128),  # wdViolet
                        13: (128, 0, 0),  # wdDarkRed
                        14: (128, 128, 0),  # wdDarkYellow
                        15: (128, 128, 128),  # wdGray50
                        16: (192, 192, 192),  # wdGray25
                    }
                    fmt["cell_bg"] = highlight_colors.get(fmt["highlight"], None)
            except Exception:
                fmt["highlight"] = None

            # 段落格式
            fmt["alignment"] = getattr(pfmt, "Alignment", None)
            fmt["line_spacing"] = getattr(pfmt, "LineSpacing", None)
            fmt["line_spacing_rule"] = getattr(pfmt, "LineSpacingRule", None)
            fmt["space_before"] = getattr(pfmt, "SpaceBefore", None)
            fmt["space_after"] = getattr(pfmt, "SpaceAfter", None)

            return fmt
        except Exception as e:
            logger.warning(f"[save_format] 保存段落格式失败: {e}")
            return {}

    def restore_format(self, cell, fmt):
        """
        将段落格式应用到 table.Cell 对象。
        """
        try:
            para = cell.Range.Paragraphs(1)
            cell_rng = cell.Range
            font = getattr(cell_rng, "Font", None)
            pfmt = (
                para.Format if hasattr(para, "Format") else para.Range.ParagraphFormat
            )

            # --- 字体 ---
            if font is not None:
                if fmt.get("font_name"):
                    font.Name = fmt["font_name"]
                if fmt.get("font_size"):
                    font.Size = fmt["font_size"]
                if fmt.get("font_bold") is not None:
                    font.Bold = fmt["font_bold"]
                if fmt.get("font_italic") is not None:
                    font.Italic = fmt["font_italic"]
                if fmt.get("font_underline") is not None:
                    font.Underline = fmt["font_underline"]
                if fmt.get("font_color"):
                    font.Color = fmt["font_color"]

            # --- 单元格文本高亮 ---
            if fmt.get("highlight") is not None:
                cell_rng.HighlightColorIndex = fmt["highlight"]
            else:
                cell_rng.HighlightColorIndex = 0  # wdNoHighlight

            # --- 单元格底色 ---
            if fmt.get("cell_bg") is not None:
                try:
                    r, g, b = fmt["cell_bg"]
                    cell.Shading.Texture = 0
                    cell.Shading.BackgroundPatternColor = b * 65536 + g * 256 + r
                except:
                    pass

            # --- 段落格式 ---
            if fmt.get("line_spacing") is not None:
                pfmt.LineSpacing = fmt["line_spacing"]
            if fmt.get("line_spacing_rule") is not None:
                pfmt.LineSpacingRule = fmt["line_spacing_rule"]
            if fmt.get("space_before") is not None:
                pfmt.SpaceBefore = fmt["space_before"]
            if fmt.get("space_after") is not None:
                pfmt.SpaceAfter = fmt["space_after"]

            # --- 水平居中 ---
            if fmt.get("alignment") is not None:
                try:
                    for p in cell.Range.Paragraphs:
                        p.Format.Alignment = fmt["alignment"]
                    cell_rng.ParagraphFormat.Alignment = fmt["alignment"]
                except:
                    pass

        except Exception as e:
            logger.warning(f"[restore_format] 应用格式失败: {e}")
