from .wps_base import WPSBase, auto_before_call
from .wps_utils import WPSUtils

import os
from typing import List, Any
from typing import overload  # 用于重载方法，根据参数类型选择不同的实现
import re
from datetime import datetime
from typing import Optional

import types

from loguru import logger


@auto_before_call(before_func="available")
class WPSExcel(WPSBase):
    def __init__(
        self,
        file_path: str,
        family_name: str = None,
        use_wps: bool = True,
        is_debug: bool = False,
    ):
        """
        初始化 ExcelUtils 类，打开指定 Excel 文件。

        参数:
        - file_path: str, Excel 文件路径
        - family_name: str = "微软雅黑", 如果是创建文件，则生效
        - use_wps: bool = True, True 表示使用 WPS Office，False 表示使用 Excel
        - is_debug: bool = False, 是否进入调试模式
        """
        # 确定使用 WPS 还是 Excel
        prog_id = "Ket.Application" if use_wps else "Excel.Application"
        # 设置 logger 级别
        logger.setLevel(logging.DEBUG if is_debug else logging.INFO)
        super().__init__(file_path, prog_id, is_debug)
        try:
            self.family_name = family_name
            self.border_color = 0x16B777
            if not os.path.exists(self.file_path):
                logger.debug(f"文件 {self.file_path} 不存在，创建新文件")
                self.excel = self.office.Workbooks.Add()
                self.excel.SaveAs(self.file_path)
            else:
                logger.debug(f"文件 {self.file_path} 已存在，打开文件")
                self.excel = self.office.Workbooks.Open(self.file_path)
        except Exception as e:
            logger.error("无法打开或创建文件 %s: %s", self.file_path, str(e))
            self.quit()
            raise RuntimeError(f"无法打开或创建文件 {self.file_path}: {str(e)}")

    # region 工作簿操作

    def get_sheet_names(self) -> list[str]:
        """
        获取工作簿中所有工作簿的名称列表。

        返回:
        - list: 工作簿名称列表
        """
        sheet_names = []
        if self.excel:
            sheet_names = [ws.Name for ws in self.excel.Worksheets]
        logger.debug(f"工作簿 {self.file_path} 中的工作表名称: {sheet_names}")
        return sheet_names

    def get_sheet(self, sheet_name: str) -> "WorksheetWrapper":
        """
        通过 sheet_name 获取工作簿对象。

        参数:
        - sheet_name: str, 工作簿名称

        返回:
        - WorksheetWrapper: 包装的工作簿对象
        """
        try:
            ws = self.excel.Worksheets(sheet_name)
            return WorksheetWrapper(ws)
        except Exception:
            logger.error(f"工作表 {sheet_name} 不存在")
            raise RuntimeError(f"工作表 {sheet_name} 不存在")

    def read_sheet(
        self,
        sheet_name: str,
        skip_row_count: int = 0,
        skip_col_count: int = 0,
        is_value: bool = False,
    ) -> list[list]:
        """
        读取指定工作表的所有单元格值。

        参数:
        - sheet_name: str, 工作表名称
        - skip_row_count: int = 0，跳过前几行
        - skip_col_count: int = 0，跳过前几列
        - is_value: bool = False，为True时，只读取值，不读取公式

        返回:
        - list of list: 二维列表，包含所有单元格的值
        """
        ws = self.get_sheet(sheet_name)
        values = ws.get_UsedRange_value(is_value=is_value)

        # 应用 skip_row_count 跳过指定行数
        if skip_row_count > 0 and skip_row_count < len(values):
            values = values[skip_row_count:]
            logger.debug(f"跳过前 {skip_row_count} 行，剩余 {len(values)} 行")

        # 应用 skip_col_count 跳过指定列数
        if skip_col_count > 0 and values and skip_col_count < len(values[0]):
            values = [row[skip_col_count:] for row in values]
            logger.debug(
                f"跳过前 {skip_col_count} 列，剩余 {len(values[0]) if values else 0} 列"
            )

        logger.debug(f"工作表 {sheet_name} 的所有单元格值: {values}")
        return values

    def copy_sheet(
        self, src_sheet_name: str, target_path: str, target_sheet_name: str
    ) -> None:
        """
        将指定工作簿复制到 target_file 中，并重命名为 target_sheet_name。
        - 如果 target_file 不存在，则创建新文件。
        - 如果 target_file 中已存在 target_sheet_name，则先删除。
        - 不删除新工作簿的默认工作簿，保留至少一张工作簿。
        - 复制工作簿中的线条、图表和图片。
        - 如果 src_sheet_name 不存在，直接跳过。

        参数:
        - src_sheet_name: str, 源工作簿名
        - target_file: str, 目标文件路径
        - target_sheet_name: str, 目标工作簿名
        """
        target_path = os.path.normpath(os.path.abspath(target_path))
        wb_dest = None

        try:
            # 检查源工作簿是否存在
            try:
                ws_src = self.excel.Worksheets(src_sheet_name)
            except Exception as e:
                logger.info(f"源工作簿 {src_sheet_name} 不存在，直接跳过")
                return

            # 打开或创建Excel文件
            if not os.path.exists(target_path):
                logger.debug(f"目标文件 {target_path} 不存在，创建新文件")
                wb_dest = self.office.Workbooks.Add()
                wb_dest.SaveAs(target_path)
            else:
                logger.debug(f"目标文件 {target_path} 已存在，打开文件")
                wb_dest = self.office.Workbooks.Open(target_path)

            # 输出目标工作簿信息
            sheet_count = wb_dest.Worksheets.Count
            sheet_names = [ws.Name for ws in wb_dest.Worksheets]
            logger.debug(f"目标工作簿包含 {sheet_count} 个工作簿：{sheet_names}")

            # 删除目标工作簿中已存在的目标工作簿
            try:
                ws_to_delete = wb_dest.Worksheets(target_sheet_name)
                logger.debug(f"目标工作簿中存在 {target_sheet_name}，删除该工作簿")
                ws_to_delete.Delete()
            except:
                logger.debug(f"目标工作簿中不存在 {target_sheet_name}，无需删除")

            # 复制源工作簿到目标工作簿
            logger.debug(f"复制工作簿 {src_sheet_name} 到目标工作簿")
            ws_src.Copy(
                Before=wb_dest.Worksheets(1) if wb_dest.Worksheets.Count > 0 else None
            )

            # 重命名新工作簿
            new_ws = wb_dest.Worksheets(1)
            new_ws.Name = target_sheet_name
            logger.info(
                f"工作簿复制成功：{target_sheet_name}",
            )

            # 复制形状和图表
            try:
                for shape in ws_src.Shapes:
                    logger.debug(
                        f"工作簿 {target_sheet_name} 复制形状：{shape.Name} (类型：{shape.Type})",
                    )
                    shape.Copy()
                    new_ws.Paste()
                    new_shape = new_ws.Shapes(new_ws.Shapes.Count)
                    new_shape.Left = shape.Left
                    new_shape.Top = shape.Top
                    logger.debug("形状 %s 已复制并定位", shape.Name)

                for chart in ws_src.ChartObjects():
                    logger.debug(
                        f"工作簿 {target_sheet_name} 复制图表：{chart.Name}",
                    )
                    chart.Copy()
                    new_ws.Paste()
                    new_chart = new_ws.ChartObjects(new_ws.ChartObjects.Count)
                    new_chart.Left = chart.Left
                    new_chart.Top = chart.Top
                    logger.debug("图表 %s 已复制并定位", chart.Name)
            except Exception as e:
                logger.warning(
                    f"工作簿 {target_sheet_name} 复制图表时出错: {str(e)}，继续执行",
                )

            # 保存并关闭目标工作簿
            wb_dest.Save()
            wb_dest.Close()
            logger.debug(
                f"保存Excel：{target_path}",
            )

        except Exception as e:
            logger.error(f"复制工作簿 {src_sheet_name} 到 {target_path} 失败: {str(e)}")
            try:
                if wb_dest:
                    wb_dest.Close(SaveChanges=False)
                    logger.warning("异常处理：关闭目标工作簿")
            except:
                pass
            raise RuntimeError(f"复制工作簿 {src_sheet_name} 失败: {str(e)}")
        finally:
            if wb_dest:
                try:
                    wb_dest.Close(SaveChanges=False)
                    logger.debug("确保目标工作簿已关闭")
                except:
                    pass

    def remove_sheet(self, sheet_name: str):
        """
        删除指定工作簿。

        参数:
        - sheet_name: str, 要删除的工作簿名称

        异常:
        - 如果工作簿仅剩一张工作簿，抛出异常以防止删除。
        - 如果工作簿不存在，记录日志。
        """
        sheet_count = self.excel.Worksheets.Count
        if sheet_count <= 1:
            logger.error(f"工作簿仅剩一张工作簿，无法删除{sheet_name}")
            raise RuntimeError("工作簿至少需要保留一张工作簿")
        try:
            ws_to_delete = self.excel.Worksheets(sheet_name)
            ws_to_delete.Delete()
            self.excel.Save()  # 保存更改
            logger.info(f"工作簿删除成功：{sheet_name}")
        except:
            logger.info(f"工作簿删除失败：{sheet_name} 不存在")

    def create_sheet(self, sheet_name: str, is_delete: bool = False):
        """
        创建一个新的工作表。

        参数:
        - sheet_name: str, 新工作表名称
        - is_delete: bool = False，是否删除已存在的工作表
        """

        if sheet_name in self.get_sheet_names():
            if is_delete:
                self.read_sheet(sheet_name=sheet_name)
            else:
                raise RuntimeError(f"工作表 '{sheet_name}' 已存在")

        new_sheet = self.excel.Worksheets.Add()
        new_sheet.Name = sheet_name
        self.excel.Save()
        logger.debug(f"成功创建新工作表：{sheet_name}")

    # engregion

    # region 单元格操作

    @overload
    def get_cell(self, sheet_name: str, cell_ref: str):
        """
        (重载)获取指定工作表中指定位置的单元格。
        参数:
        - sheet_name: str, 工作表名称
        - cell_ref: str, 单元格位置(如A1、B10、AB200等)

        返回:
        - Cell, 指定位置的单元格COM对象
        """

        ...

    @overload
    def get_cell(self, sheet_name: str, row: int, col: int):
        """
        (重载)获取指定工作表中指定位置的单元格。

        参数:
        - sheet_name: str, 工作表名称
        - row: int, 行索引（从1开始）
        - col: int, 列索引（从1开始）

        返回:
        - Cell, 指定位置的单元格COM对象
        """

        ...

    def get_cell(self, *args):
        """
        (重载)获取指定工作表中指定位置的单元格。

        返回:
        - Cell, 指定位置的单元格COM对象
        """

        row_index, col_index = (0, 0)
        sheet_name = None
        cell_ref = None

        if len(args) == 2:
            sheet_name, cell_ref = args
            row_index, col_index = self.cell_to_row_col(cell_ref)
        elif len(args) == 3:
            sheet_name, row_index, col_index = args
        else:
            logger.error("get_cell 方法参数错误")
            raise ValueError("get_cell 方法参数错误")
        worksheet = self.get_sheet(sheet_name)
        cell = worksheet.get_cell(row_index, col_index)
        logger.debug(f"访问单元格 {cell_ref}：({row_index}, {col_index})")
        return cell

    def get_values(self, sheet_cell_ref: str, is_value: bool = True):
        """获取单元格或区域数据

        Args:
            key: 单元格坐标，如'Sheet1.A1'或'Sheet1.A1:C4'

        Returns:
            单元格数据或数据数组
        """
        # 解析工作表名称和单元格坐标
        parts = sheet_cell_ref.strip().split(".")
        if len(parts) != 2:
            raise ValueError(
                f"当前坐标：{sheet_cell_ref}\n单元格坐标格式错误，正确格式为'Sheet1.A1'或'Sheet1.A1:C4'"
            )

        sheet_name, cell_ref = parts

        # 检查工作表是否存在
        if sheet_name not in self.get_sheet_names():
            raise ValueError(f"工作表 '{sheet_name}' 不存在")

        # 获取指定工作表
        sheet = self.get_sheet(sheet_name=sheet_name)
        # 检查是否为范围格式（包含:）
        if ":" in cell_ref:
            return sheet.get_range_values(cell_range=cell_ref, is_value=is_value)
        else:
            return sheet.get_cell_value(cell_ref=cell_ref, is_value=is_value)

    def set_cell_value(
        self,
        sheet_name: str,
        cell_ref: str,
        value: Any,
        number_format: str = None,
        is_bold: bool = False,
        auto_save: bool = True,
    ):
        """
        写入单元格值。

        参数:
        - sheet_name: str, 工作表名称
        - cell_ref: str, 单元格位置(如A1、B10、AB200等)
        - value: Any, 要写入的单元格值
        - number_format: str, 单元格数字格式（如"#,##0.00"）
        - is_bold: bool, 是否加粗文字
        - auto_save: bool, 是否自动保存工作簿
        """
        worksheet = self.get_sheet(sheet_name)
        cell = worksheet.get_cell(cell_ref)
        cell.Value = value
        if isinstance(value, str) and value.startswith("="):
            cell.Borders.Color = self.border_color
        logger.debug(f"写入单元格 {cell_ref}: {value}")

        # 设置数字格式
        if number_format:
            cell.NumberFormat = number_format

        # 设置字体加粗
        if is_bold:
            cell.Font.Bold = True

        if self.family_name:
            cell.Font.Name = self.family_name

        # 自动保存
        if auto_save:
            self.save()

    def set_merge_cell(
        self,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        value: str,
        is_bold=False,
        auto_save: bool = True,
    ):
        """
        合并指定区域的单元格并写入值。

        参数:
        - sheet_name: str, 工作表名称
        - start_cell: str, 起始单元格（如 A1）
        - end_cell: str, 结束单元格（如 C3）
        - value: str, 要写入的值
        - is_bold: bool = False, 是否加粗文字
        """
        # 确保工作表存在
        if sheet_name not in self.get_sheet_names():
            self.create_sheet(sheet_name=sheet_name)
        ws = self.get_sheet(sheet_name)
        merge_range = ws.worksheet.Range(f"{start_cell}:{end_cell}")
        merge_range.Merge()
        merge_range.Value = value
        if is_bold:
            merge_range.Font.Bold = True
        if self.family_name:
            merge_range.Font.Name = self.family_name
        # 如果以=开头，设置绿色边框
        if isinstance(value, str) and value.startswith("="):
            merge_range.Borders.Color = self.border_color
        # 自动保存
        if auto_save:
            self.save()
        logger.debug(
            f"合并单元格 {sheet_name}!{start_cell}:{end_cell} 并写入值: {value}"
        )

    def set_range_values(
        self,
        sheet_name: str,
        start_cell: str,
        data: List[List[Any]],
        auto_save: bool = True,
        number_format: str = None,
    ):
        """
        将二维数据写入指定工作表的连续区域。
        二维数组每行数量必须一致，否则会写入失败，会自动补充默认值（默认 ""）。

        参数:
        - sheet_name: str, 目标工作表名称
        - start_cell: str, 起始单元格（如 A1）
        - data: List[List[Any]], 要写入的二维数据
        - auto_save: bool = True, 写入后是否自动保存
        - number_format: str, 可选，统一设置数字格式
        """
        if not data or not data[0]:
            logger.warning("写入数据为空，跳过写入")
            return
        # 确保工作表存在
        if sheet_name not in self.get_sheet_names():
            self.create_sheet(sheet_name=sheet_name)

        ws = self.get_sheet(sheet_name)
        rows, cols = len(data), len(data[0])

        # 计算结束单元格
        end_cell = self.calculate_end_cell(start_cell, data)

        # 获取目标区域
        target_range = ws.worksheet.Range(f"{start_cell}:{end_cell}")

        # 规范二维数组每行长度一致
        data = WPSUtils.normalize_row_lengths(data)

        # 设置字体格式
        if self.family_name:
            target_range.Font.Name = self.family_name

        # 遍历数据，若单元格值以=开头，则设置绿色边框
        start_row_letter, start_row_num = self.cell_to_row_col(start_cell)
        start_row = int(start_row_num)
        start_col = 0
        for ch in start_row_letter:
            start_col = start_col * 26 + (ord(ch) - ord("A") + 1)
        for r_idx, row in enumerate(data):
            for c_idx, cell_value in enumerate(row):
                cell = ws.worksheet.Cells(start_row + r_idx, start_col + c_idx)
                # ========= 1、公式 =========
                if isinstance(cell_value, str) and cell_value.startswith("="):
                    if number_format:
                        cell.NumberFormat = number_format
                    cell.Formula = cell_value
                    cell.Borders.Color = self.border_color
                    continue

                # ========= 2、数字 =========
                if WPSUtils.is_number(cell_value):
                    if number_format:
                        cell.NumberFormat = number_format
                    cell.Value = cell_value
                    continue

                # ========= 3、字符串 =========
                if isinstance(cell_value, str):
                    # 日期字符串
                    is_date, date_format = WPSUtils.get_date_format(cell_value)
                    if is_date:
                        cell.NumberFormat = date_format
                        cell.Value = cell_value
                        continue

                    # 普通文本（重点：先设置格式）
                    cell.NumberFormat = "@"
                    cell.Value = str(cell_value)
                    continue

                # ========= 4、其他 =========
                cell.Value = cell_value

        # 自动保存
        if auto_save:
            self.save()

    def set_range_color(
        self,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        color: str,
        auto_save=True,
    ):
        """
        设置指定区域内单元格的背景颜色。

        参数:
        - sheet_name: str, 工作表名称
        - start_cell: str, 起始单元格（如 A1）
        - end_cell: str, 结束单元格（如 C3）
        - color: str, 颜色值，支持十六进制（如 "#FF0000"）
        - auto_save: bool = True, 写入后是否自动保存
        """
        # 获取工作表
        ws = self.get_sheet(sheet_name)

        # 获取区域
        target_range = ws.worksheet.Range(f"{start_cell}:{end_cell}")

        # 解析颜色
        rgb_color = WPSUtils.hex_to_bgr(color)

        # 设置背景颜色
        target_range.Interior.Color = rgb_color
        logger.debug(
            f"设置工作表 {sheet_name} 区域 {start_cell}:{end_cell} 背景颜色为 {color}"
        )

        # 自动保存
        if auto_save:
            self.save()

    def add_annotation(
        self,
        sheet_name: str,
        cell_ref: str,
        annotation: str,
        author: str = "系统",
        auto_save: bool = True,
    ):
        """添加批注"""
        ws = self.get_sheet(sheet_name)
        cell_obj = ws.get_cell(cell_ref)
        # 如果已有批注，先删除旧的
        if cell_obj.Comment is not None:
            cell_obj.Comment.Delete()

        # 添加批注（AddComment 返回 Comment 对象），wps下无法设置author
        cell_obj.AddComment(annotation)
        # 自动保存
        if auto_save:
            self.save()
        logger.debug(
            f"工作表 {sheet_name} 单元格 {cell_ref} 添加批注：{annotation} (作者：{author})"
        )

    def set_conditional_format(
        self,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        condition_type: str,
        condition_value: Any,
        background_color: str = "FF0000",
        font_color: str = "FFFFFF",
        auto_save: bool = True,
    ):
        """
        为指定区域设置条件格式。

        参数:
        - sheet_name: str, 工作表名称
        - start_cell: str, 起始单元格（如 A1）
        - end_cell: str, 结束单元格（如 C3）
        - condition_type: str, 条件类型，支持 greater/less/equal/not_equal/between/contains
        - condition_value: Any, 条件值；between 时传入 (min, max) 元组
        - background_color: str, 背景颜色，6 位十六进制，默认 FF0000（红）
        - font_color: str, 字体颜色，6 位十六进制，默认 FFFFFF（白）
        - auto_save: bool, 是否自动保存
        """
        if sheet_name not in self.get_sheet_names():
            raise ValueError(f"工作表 '{sheet_name}' 不存在")

        ws = self.get_sheet(sheet_name).worksheet
        cell_range = f"{start_cell}:{end_cell}"

        bg_color = WPSUtils.hex_to_bgr(background_color)
        ft_color = WPSUtils.hex_to_bgr(font_color)

        # 删除同区域旧规则，避免叠加
        try:
            fc = ws.Range(cell_range).FormatConditions
            for i in range(fc.Count, 0, -1):
                try:
                    fc.Item(i).Delete()
                except Exception:
                    pass
        except Exception:
            pass

        # 新建条件格式
        operator = self.__map_operator(condition_type)
        formula1 = self.__build_formula1(condition_type, condition_value, start_cell)
        formula2 = self.__build_formula2(condition_type, condition_value)

        if condition_type == "contains":
            # 使用公式类型的条件格式
            fmt = ws.Range(cell_range).FormatConditions.Add(
                Type=2,  # xlExpression
                Operator=None,
                Formula1=formula1,
            )
        elif condition_type == "between":
            fmt = ws.Range(cell_range).FormatConditions.Add(
                Type=1,  # xlCellValue
                Operator=operator,
                Formula1=formula1,
                Formula2=formula2,
            )
        else:
            fmt = ws.Range(cell_range).FormatConditions.Add(
                Type=1,  # xlCellValue
                Operator=operator,
                Formula1=formula1,
            )

        # 设置样式
        fmt.Interior.Color = bg_color
        fmt.Font.Color = ft_color

        logger.debug(
            f"条件格式设置成功：{sheet_name}!{cell_range} "
            f"类型={condition_type} 值={condition_value} "
            f"背景={background_color} 字体={font_color}"
        )

        if auto_save:
            self.save()

    # region 内部辅助
    @staticmethod
    def __map_operator(condition_type: str) -> int:
        """将条件类型映射为 Excel COM 常量"""
        mapping = {
            "between": 1,  # xlBetween
            "not_between": 2,  # xlNotBetween
            "equal": 3,  # xlEqual
            "not_equal": 4,  # xlNotEqual
            "greater": 5,  # xlGreater
            "less": 6,  # xlLess
            "greater_equal": 7,  # xlGreaterEqual
            "less_equal": 8,  # xlLessEqual
            "contains": -1,  # 特殊处理，用公式
        }
        if condition_type not in mapping:
            raise ValueError(f"不支持的条件类型: {condition_type}")
        return mapping[condition_type]

    @staticmethod
    def __build_formula1(
        condition_type: str, value: Any, start_cell: str = "A1"
    ) -> str:
        if condition_type == "contains":
            return f'=NOT(ISERROR(SEARCH("{value}",{start_cell})))'
        if condition_type == "between":
            return str(value[0])
        if isinstance(value, str):
            return f'"{value}"'
        return str(value)

    @staticmethod
    def __build_formula2(condition_type: str, value: Any) -> Optional[str]:
        return str(value[1]) if condition_type == "between" else None

    # endregion
    # endregion

    # region 工具方法

    @staticmethod
    def calculate_end_cell(
        start_cell: str,
        data: List[List[Any]] = None,
        add_row_count: int = 0,
        add_col_count: int = 0,
    ) -> str:
        """根据起始单元格和二维数组计算结束单元格位置

        Args:
            start_cell: 起始单元格位置(如A1)
            data: 二维数组数据
            add_row_count: 额外增加的行数
            add_col_count: 额外增加的列数

        Returns:
            str: 结束单元格位置(如B2)
        """
        # 处理data为None或空的情况
        if data is None or not data:
            data = [[]]

        # 使用正则提取起始列字母与行号
        m = re.match(r"([A-Z]+)(\d+)", start_cell.upper())
        if not m:
            raise ValueError("起始单元格格式错误，应为如 A1 的格式")
        col_part, row_part = m.groups()
        start_row = int(row_part)

        # 列字母转数字（A->1, B->2, ..., Z->26, AA->27...）
        start_col = 0
        for ch in col_part:
            start_col = start_col * 26 + (ord(ch) - ord("A") + 1)

        # 计算结束行列
        end_row = start_row + max(len(data) - 1, 0) + add_row_count
        max_col_length = max(len(row) if row else 0 for row in data)
        end_col_num = start_col + max(max_col_length - 1, 0) + add_col_count

        end_col_letter = WPSExcel.col_num_to_letter(end_col_num)
        return f"{end_col_letter}{end_row}"

    @staticmethod
    def cell_to_row_col(cell_ref: str) -> tuple[int, int]:
        """
        将 Excel 单元格坐标（如 A1、AA1、AZ4）转换为 (行号, 列号) 的元组，行列均从 1 开始。

        参数:
        - cell_ref: str, 单元格地址，形如 "A1"

        返回:
        - tuple[int, int]: (行号, 列号)
        """
        m = re.match(r"([A-Z]+)(\d+)", cell_ref.upper())
        if not m:
            raise ValueError("单元格格式错误，应为如 A1 的格式")
        col_part, row_part = m.groups()
        row = int(row_part)
        col = 0
        for ch in col_part:
            col = col * 26 + (ord(ch) - ord("A") + 1)
        return row, col

    def auto_adjust_columns(
        self,
        sheet_name: str,
        columns: List[str] = None,
        padding: int = 4,
        max_width: int = None,
    ):
        """
        自动调整指定工作表的列宽

        参数:
         - sheet_name: str， 工作表名称
         - columns: List[str] = None，需要调整的列名列表(如['A', 'B'])，为None时调整所有列
         - padding: int = 4，列宽边距
         - max_width: int = None，最大列宽限制，为None时不限制
        """
        ws = self.get_sheet(sheet_name)
        used_range = ws.worksheet.UsedRange

        # 如果未指定列，则调整所有已用列
        if columns is None:
            # 使用静态方法 col_num_to_letter 将列号转为字母
            columns = [self.col_num_to_letter(c.Column) for c in used_range.Columns]

        for col_letter in columns:
            try:
                col_range = ws.worksheet.Range(f"{col_letter}:{col_letter}")
                # 自动适应内容
                col_range.AutoFit()
                # 获取自动适应后的宽度
                current_width = col_range.ColumnWidth
                # 增加边距
                new_width = current_width + padding
                # 应用最大宽度限制
                if max_width is not None and new_width > max_width:
                    new_width = max_width
                col_range.ColumnWidth = new_width
                logger.debug(
                    f"工作表 {sheet_name} 列 {col_letter} 宽度调整为 {new_width}"
                )
            except Exception as e:
                logger.warning(
                    f"调整工作表 {sheet_name} 列 {col_letter} 宽度失败: {str(e)}"
                )

        self.save()

    @staticmethod
    def col_num_to_letter(n: int) -> str:
        """
        将列号（1 基）转换为 Excel 列字母（如 1 -> A，28 -> AB）

        参数:
        - n: int, 列号，从 1 开始

        返回:
        - str, 对应的列字母
        """
        letter = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letter = chr(ord("A") + rem) + letter
        return letter or "A"

    @staticmethod
    def cell_to_row_col(cell_ref: str) -> tuple[str, str]:
        """
        将 Excel 单元格坐标（如 A1、AA1、AZ4）转换为 (列字母, 行号) 的元组。

        参数:
        - cell_ref: str, 单元格坐标，如 "A1"

        返回:
        - tuple[str, str]: (列字母, 行号)
        """
        cell_ref = cell_ref.upper()
        match = re.match(r"([A-Z]+)(\d+)", cell_ref)
        if not match:
            raise ValueError(f"单元格坐标格式错误: {cell_ref}")

        col_str, row_str = match.groups()
        return col_str, row_str

    def available(self):
        """
        判断excel对象是否可用，如果可用，则不做处理，不可用，则重新创建
        """
        try:
            # 尝试访问工作簿的Name属性，若抛出异常说明COM对象已失效
            _ = self.excel.Name
            return True
        except Exception as e:
            logger.warning(f"Excel COM 对象不可用，准备重新创建: {e}")
            # 释放旧对象
            self.excel = None
            # 重新打开工作簿
            try:
                self.excel = self.office.Workbooks.Open(self.file_path)
                logger.info("Excel 工作簿已重新打开")
                return True
            except Exception as reopen_err:
                logger.error(f"重新创建工作簿失败: {reopen_err}")
                raise RuntimeError(f"无法重新创建工作簿: {reopen_err}")

    # endregion

    # region 退出操作
    def save(self, output_path=None, is_transfer: bool = False):
        """

        保存 Excel 文件。

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
            self.excel.SaveAs(save_path)
            if is_transfer:
                self.excel = self.office.Workbooks.Open(save_path)
                self.file_path = save_path
                logger.debug(f"当前处理工作簿更新为：{save_path}")
        except Exception as e:
            logger.error(f"保存Excel失败 {save_path}: {str(e)}")
            raise RuntimeError(f"保存Excel失败 {save_path}: {str(e)}")
        try:
            self.excel.SaveAs(save_path)
            if is_transfer:
                self.excel = self.office.Workbooks.Open(save_path)
                self.file_path = save_path
                logger.debug(f"当前处理工作簿更新为：{save_path}")
        except Exception as e:
            logger.error(f"保存Excel失败 {save_path}: {str(e)}")
            raise RuntimeError(f"保存Excel失败 {save_path}: {str(e)}")

    def close(self):
        """
        关闭工作簿。
        """
        logger.info(f"关闭工作簿：{self.file_path}")
        if self.excel:
            try:
                # 检查COM对象是否有效
                # 可以通过尝试访问一个简单属性来测试
                _ = self.excel.Name  # 如果对象无效，这里会抛出异常
                self.excel.Close(SaveChanges=False)
            except:
                logger.warning(f"工作簿 {self.file_path} 可能已关闭或无效")
            finally:
                self.excel = None
        self.quit()

    # endregion


class WorksheetWrapper:
    def __init__(self, worksheet):
        """
        包装 COM 工作簿对象工作簿。

        参数:
        - worksheet: COM Worksheet 对象
        """
        self.worksheet = worksheet
        logger.debug(f"初始化工作簿：{self.worksheet.Name}")

    def get_UsedRange_value(self, is_value: bool):
        """
        获取已使用范围的所有单元格值。

        参数:
        - is_value: bool，默认False，为True时，只读取值，不读取公式

        返回:
        - list of list: 二维列表，包含所有单元格的值
        """
        used_range = self.worksheet.UsedRange
        values = [[]]
        if is_value:
            values = used_range.Value or [[]]
        else:
            values = used_range.Formula or [[]]

        # 如果只有一行或一列，确保是二维列表
        if not isinstance(values[0], (list, tuple)):
            values = [list(values)]

        # 去掉末尾全 None、空字符串或仅含空格的列
        if values:
            max_col = len(values[0])
            for col_idx in reversed(range(max_col)):
                if all(
                    row[col_idx] is None or str(row[col_idx]).strip() == ""
                    for row in values
                ):
                    max_col -= 1
                else:
                    break
            values = [row[:max_col] for row in values]

        # 去掉末尾全 None、空字符串或仅含空格的行
        while values and all(
            cell is None or str(cell).strip() == "" for cell in values[-1]
        ):
            values.pop()

        logger.debug(f"获取已使用范围值：{values}")
        return values

    @overload
    def get_cell(self, cell_ref: str):
        """
        (重载)获取指定单元格对象。

        参数:
        - cell_ref: str, 单元格地址，形如 "A1"

        返回:
        - CellWrapper: 包装的单元格对象
        """
        ...

    @overload
    def get_cell(self, row: int, col: int):
        """
        (重载)获取指定单元格对象。

        参数:
        - row: int, 行索引，从1开始
        - col: int, 列索引，从1开始

        返回:
        - CellWrapper: 包装的单元格对象
        """

    def get_cell(self, *args):
        """
        (重载)获取指定单元格对象。

        返回:
        - CellWrapper: 包装的单元格对象
        """
        # 处理单元格引用字符串
        if len(args) == 1 and isinstance(args[0], str):
            cell_ref = args[0]
            cell = self.worksheet.Range(cell_ref)
            logger.debug(f"访问单元格：{cell_ref}")
            return cell
        # 处理行索引和列索引
        if len(args) == 2 and all(isinstance(x, int) for x in args):
            row, col = args
            cell = self.worksheet.Cells(row, col)
            logger.debug(f"访问单元格：{row},{col}")
            return cell
        raise TypeError("get_cell 参数不合法")

    def get_cell_value(self, cell_ref: str, is_value: bool = True):
        """
        获取指定单元格的值或公式。

        参数:
        - cell_ref: str, 单元格地址，形如 "A1"
        - is_value: bool, 为 True 时返回单元格的值，为 False 时返回单元格的公式，默认为 True

        返回:
        - 单元格的值或公式；若值为日期或时间，则按 Excel 文件中的显示格式返回文本
        """
        # 通过 get_cell 获取单元格 COM 对象
        cell = self.get_cell(cell_ref)
        data = ""
        # 根据 is_value 决定返回 Value 还是 Formula
        if is_value:
            # Value 返回单元格的实际值（公式计算结果或常量值）
            data = cell.Value
            if isinstance(data, (int, float)) and cell.NumberFormat == "@":
                data = str(cell.Text)

        else:
            data = cell.Formula

        # 若 data 为日期或时间类型，则按 Excel 显示格式返回文本
        if isinstance(data, datetime):
            data = str(cell.Text)  # Text 属性即为 Excel 中显示的格式

        # 将None、纯空格处理为空字符串
        data = (
            ""
            if data is None or (isinstance(data, str) and data.strip() == "")
            else data
        )
        return data

    def get_range_values(self, cell_range: str, is_value: bool = True):
        """
        获取指定区域所有单元格的值或公式。

        参数:
        - cell_range: str, 区域地址，形如 "A1:C7"
        - is_value: bool, 为 True 时返回单元格的值，为 False 时返回单元格的公式，默认为 True

        返回:
        - list[list]: 二维列表，包含区域内所有单元格的值或公式
        """
        # 获取区域 COM 对象
        rng = self.worksheet.Range(cell_range)
        # 根据 is_value 决定返回 Value 还是 Formula
        if is_value:
            data = rng.Value
        else:
            data = rng.Formula

        # 如果只有一行或一列，确保返回二维列表
        if not isinstance(data[0], (list, tuple)):
            data = [list(data)]
        else:
            # 将外层元组转为列表，避免后续赋值时报错
            data = [list(row) for row in data]

        # 将None、纯空格处理为空字符串
        for i in range(len(data)):
            for j in range(len(data[i])):
                value = data[i][j]
                if value is None or (isinstance(value, str) and value.strip() == ""):
                    data[i][j] = ""
                elif isinstance(value, datetime):
                    # 获取对应单元格的自定义格式
                    cell = self.worksheet.Cells(rng.Row + i, rng.Column + j)
                    fmt = cell.NumberFormatLocal
                    try:
                        # 用 Excel 的格式字符串将 datetime 转为 str
                        data[i][j] = value.strftime(
                            WPSUtils.excel_format_to_python(fmt)
                        )
                    except Exception:
                        # 转换失败时退化为默认字符串
                        data[i][j] = str(value)

        return data
