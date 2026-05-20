from functools import cache

import re
import os
from .wps_excel import WPSExcel
from .wps_word import WPSWord

from loguru import logger


class BaseExport:
    def __init__(self, data: dict) -> None:
        self.data = data

    def get_data(self, text: str, is_replace=True):
        """处理文本中的{{xxx}}占位符,从data_excel中获取对应的数据进行替换"""
        result_str = text
        result_list = []
        # 查找所有{{xxx}}格式的内容
        matches = re.finditer(r"{{(.*?)}}", text)

        for match in matches:
            placeholder = match.group(0)  # 完整的占位符 {{xxx}}
            key = match.group(1).strip()  # 提取占位符中的内容
            value = self.data.get(key, "")
            # 替换占位符为实际数据
            if is_replace:
                result_str = result_str.replace(
                    placeholder,
                    f"{value:,.2f}" if isinstance(value, (int, float)) else str(value),
                )
            result_list.append(
                {
                    "placeholder": placeholder,
                    "key": key,
                    "value": value,
                }
            )

        return result_str if is_replace else result_list

    @staticmethod
    def set_number_decimal_places(value, decimal_places=2):
        """
        设置数值的小数位数
        """
        if isinstance(value, (int, float)):
            return f"{value:,.{decimal_places}f}"
        return str(value)


class WordExport(BaseExport):
    def __init__(self, template_path: str, data: dict) -> None:
        super().__init__(data)
        self.template_word = WPSWord(template_path)

    def export(self, output_path: str, page_number_start: int | None = 1) -> None:
        """
        导出Word，基于模板复制并填充数据到新文件
        """
        self.template_word.save(output_path, is_transfer=True)

        # 替换表格中的占位符，必须在段落替换之前，因为段落替换会改变表格中的占位符
        for table in self.template_word.get_tables():
            for row in table.rows:
                for cell in row.cells:
                    self._replace_placeholder(cell, is_last=True)

        # 替换段落中的占位符
        self._replace_paragraphs()

        output_word = WPSWord(output_path)
        output_word.set_page_start_number(page_number_start)
        output_word.save(output_path)
        output_word.close()

    def _replace_paragraphs(self):
        self.in_region = False
        paragraphs = self.template_word.get_paragraphs()
        for idx, para in enumerate(paragraphs):
            full_text = para.Range.Text if hasattr(para, "Range") else ""
            if not full_text.strip():
                continue  # 空行直接跳过

            is_last = idx == len(paragraphs) - 1

            # 处理段落文本，移除首尾空格
            tmp_full_text = full_text.strip()
            # 检查是否是区域开始标记
            if tmp_full_text.startswith("[[START;") and tmp_full_text.endswith("]]"):
                self.in_region = True
                self.region_buffer = [para]
                region_info = self._parse_region(tmp_full_text)
                continue

            # 检查是否是区域结束标记
            if tmp_full_text.startswith("[[END]]"):
                # 清除标记
                self._remove_paragraph(self.region_buffer[0])
                self._remove_paragraph(para)
                self.region_buffer.remove(self.region_buffer[0])
                # 调用函数决定是否删除区域
                if not self._show_fun(region_info):
                    for p in self.region_buffer:
                        self._remove_paragraph(p)
                else:
                    for p in self.region_buffer:
                        self._replace_placeholder(p, is_last=is_last)
                self.in_region = False
                self.region_buffer = []
                continue

            # 区域内内容 → 加入缓存，只能放到最后
            if self.in_region:
                self.region_buffer.append(para)
                continue

            # 区域外内容 → 处理占位符
            if "{{" in full_text and "}}" in full_text:
                self._replace_placeholder(para, full_text, is_last=is_last)

    def _replace_placeholder(self, para, full_text=None, is_last=False):
        """用一个 run 替换段落中的占位符，保留原 run 样式"""
        # 拼接整段落文本
        if full_text is None:
            full_text = para.Range.Text

        if str(full_text).replace(" ", "").startswith("{{FUNC."):
            return self.add_func_paragraph(para)

        # 如果段落没有占位符，直接跳过
        if "{{" not in full_text or "}}" not in full_text:
            return

        # 仅替换文本，不改动任何样式
        data_list = self.get_data(full_text, is_replace=False)
        for index, data in enumerate(data_list):
            placeholder = data["placeholder"]
            value = data["value"]
            if not isinstance(value, list):
                if (value is None or str(value).strip() == "") and len(data_list) == 1:
                    self._remove_paragraph(para=para)
                    return

                # 替换占位符为对应值
                full_text = full_text.replace(
                    placeholder,
                    f"{value:,.2f}" if isinstance(value, (int, float)) else str(value),
                )
                if is_last:
                    full_text = full_text.rstrip()
            else:
                if index != 0:
                    # 表格类型数据必须独占一行，否则会导致格式错误
                    raise ValueError(f"表格类型数据必须独占一行：\n{full_text}")
                self.add_table(para, value)
                return  # 直接返回，不执行后续的 set_pars_text

        # 将替换后的文本写回 Range，但不改变原有样式
        self.template_word.set_para_text(para=para, text=full_text)

    def _remove_paragraph(self, para):
        """删除整个段落"""
        para.Range.Delete()

    def _parse_region(self, text):
        """解析区域标记 [[START;show_fun=合计;show_param=附录数据.Q67:R69]]"""
        text = text[2:-2]
        info = {}
        for item in text.split(";"):
            if "=" in item:
                key, value = item.split("=", 1)
                info[key.strip()] = value.strip()
        return info

    def _show_fun(self, region_info):
        """是否删除整个区域"""
        if region_info is None:
            return True

        if "show_fun" in region_info and "show_param" in region_info:
            show_param = region_info["show_param"].strip()
            data_value = self.get_data("{{" + show_param + "}}", is_replace=False)[0][
                "value"
            ]
            total = 0
            # 如果data_value是数组
            if isinstance(data_value, list):
                # 遍历二维数组,将所有单元格的值转换为小数并求和
                for row in data_value:
                    for cell in row:
                        try:
                            # 尝试将单元格值转换为float类型
                            total += float(str(cell).replace(",", ""))
                        except (ValueError, TypeError):
                            # 如果转换失败则跳过该单元格
                            continue
            else:
                total = data_value
            match region_info["show_fun"]:
                case "不为零":
                    return total != 0
                case "大于零":
                    return total > 0
                case "小于零":
                    return total < 0
                case "等于零":
                    return total == 0
                case "有值":
                    try:
                        # 先检查是否为None
                        if total is None:
                            return False
                        # 尝试转换为字符串并检查是否为空
                        str_total = str(total).strip()
                        if str_total == "":
                            return False
                        # 尝试转换为数字并检查是否为0
                        try:
                            num_total = float(str_total.replace(",", ""))
                            return num_total != 0
                        except (ValueError, TypeError):
                            # 如果无法转换为数字，只要字符串不为空即可
                            return len(str_total) > 0
                    except:
                        return False
                case "无值":
                    try:
                        # 先检查是否为None
                        if total is None:
                            return True
                        # 尝试转换为字符串并检查是否为空
                        str_total = str(total).strip()
                        if str_total == "":
                            return True
                        # 尝试转换为数字并检查是否为0
                        try:
                            num_total = float(str_total.replace(",", ""))
                            return num_total == 0
                        except (ValueError, TypeError):
                            # 如果无法转换为数字，只要字符串为空即可
                            return len(str_total) == 0
                    except:
                        return False

        return False

    def add_table(self, paragraph, value, is_remove_empty_row=True):
        """
        向段落中添加表格
        """
        # 移除空行
        if is_remove_empty_row:
            value = [
                row for row in value if not all(cell == "" or cell == 0 for cell in row)
            ]

        rng = paragraph.Range

        # 如果value长度<=2，删除当前段落
        if len(value) <= 2:
            rng.Delete()
            return

        # 包含以下内容的列，左对齐
        left_align_content = ["固定资产", "无形资产", "未分配利润"]
        left_align_index = []

        # 将value的数值转为字符串，保留两位小数
        for row in value:
            for i in range(len(row)):
                # 左对齐内容，添加到索引列表
                if isinstance(row[i], str) and any(
                    content in row[i] for content in left_align_content
                ):
                    left_align_index.append(i)
                # 第一列，整数部分不添加千分位分隔，保留0位小数
                elif isinstance(row[i], (int, float)) and i == 0:
                    row[i] = f"{row[i]:.0f}"
                # 其他列，保留2位小数
                elif isinstance(row[i], (int, float)):
                    row[i] = f"{row[i]:,.2f}"

        # 插入表格
        table = self.template_word.insert_table(
            para=paragraph,
            data=value,
        )

        if left_align_index:
            # 左对齐索引列表中的列
            for index, row in enumerate(table.Rows):
                if index == 0:
                    continue
                for index in left_align_index:
                    row.Cells(index + 1).Range.ParagraphFormat.Alignment = (
                        0  # wdAlignParagraphLeft
                    )

        # 保留上下外边框，移除左右外边框
        table.Borders.Enable = True
        table.Borders.InsideLineStyle = 1  # 内边框保持实线
        table.Borders.OutsideLineStyle = 0  # 先关闭整体外边框
        # 单独开启上下边框
        table.Borders.Item(1).LineStyle = 1  # wdBorderTop
        table.Borders.Item(3).LineStyle = 1  # wdBorderBottom

        # 设置表格行高
        self.template_word.set_table_row_height(table, height=20)

        # 设置表头重复
        table.Rows(1).HeadingFormat = True

        return table
