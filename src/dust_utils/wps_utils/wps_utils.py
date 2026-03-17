import re
from datetime import datetime
from typing import List, Any
import logging

logger = logging.getLogger(__name__)


class WPSUtils:

    @staticmethod
    def excel_format_to_python(fmt: str) -> str:
        """
        将 Excel 自定义日期格式字符串转换为 Python 的 strftime 格式字符串
        """
        replacements = {
            "yyyy": "%Y",
            "yy": "%y",
            "mmmm": "%B",
            "mmm": "%b",
            "mm": "%m",
            "m": "%m",
            "dddd": "%A",
            "ddd": "%a",
            "dd": "%d",
            "d": "%d",
            "hh": "%H",
            "h": "%H",
            "nn": "%M",  # Excel 中 n 表示分钟
            "ss": "%S",
        }

        # 先转成普通字符串，确保不会被格式化系统干扰
        result = str(fmt)

        # 从长到短替换
        for k in sorted(replacements.keys(), key=len, reverse=True):
            result = result.replace(k, replacements[k])

        # 最关键：去掉多余的 %%
        result = result.replace("%%", "%")

        return result

    @staticmethod
    def get_date_format(s: str) -> tuple[bool, str]:
        """
        判断字符串是否可以解析为日期。
        自动兼容多种日期格式（中英文、数字、带时间等）。
        返回:
        - tuple[bool, str]: (是否匹配成功, Excel自定义格式字符串)
        """

        if not s or not isinstance(s, str):
            return (False, "")

        s = s.strip()
        if not s:
            return (False, "")

        # ✅ 1. 快速过滤不可能的字符串
        if not re.search(r"\d", s):
            return (False, "")

        # ✅ 2. 常见日期格式（按优先顺序尝试）
        date_formats = [
            ("%Y-%m-%d", "yyyy-mm-dd"),
            ("%Y/%m/%d", "yyyy/mm/dd"),
            ("%Y.%m.%d", "yyyy.mm.dd"),
            ("%Y-%m-%d %H:%M:%S", "yyyy-mm-dd hh:mm:ss"),
            ("%Y/%m/%d %H:%M:%S", "yyyy/mm/dd hh:mm:ss"),
            ("%Y.%m.%d %H:%M:%S", "yyyy.mm.dd hh:mm:ss"),
            ("%Y-%m-%d %H:%M", "yyyy-mm-dd hh:mm"),
            ("%Y/%m/%d %H:%M", "yyyy/mm/dd hh:mm"),
            ("%Y.%m.%d %H:%M", "yyyy.mm.dd hh:mm"),
            # ("%Y%m%d", "yyyymmdd"),
            # ("%Y%m", "yyyymm"),
            ("%Y-%m", "yyyy-mm"),
            ("%Y-%m-%dT%H:%M:%S", "yyyy-mm-ddThh:mm:ss"),
            ("%Y-%m-%dT%H:%M:%S%z", "yyyy-mm-ddThh:mm:ss"),
            ("%d/%m/%Y", "dd/mm/yyyy"),
            ("%d-%m-%Y", "dd-mm-yyyy"),
            ("%d.%m.%Y", "dd.mm.yyyy"),
            ("%m/%d/%Y", "mm/dd/yyyy"),
            ("%m-%d-%Y", "mm-dd-yyyy"),
            ("%b %d, %Y", "mmm dd, yyyy"),
            ("%B %d, %Y", "mmmm dd, yyyy"),
            ("%d %b %Y", "dd mmm yyyy"),
            ("%d %B %Y", "dd mmmm yyyy"),
        ]

        for fmt, excel_fmt in date_formats:
            try:
                datetime.strptime(s, fmt)
                return (True, excel_fmt)
            except ValueError:
                continue

        # ✅ 3. 处理中文日期（年/月/日/时/分/秒）
        zh_pattern = re.compile(
            r"^\s*(\d{2,4})年(\d{1,2})月(\d{1,2})日"
            r"(\s*(\d{1,2})[时:：](\d{1,2})([分:：](\d{1,2})秒?)?)?\s*$"
        )
        if zh_pattern.match(s):
            return (True, 'yyyy"年"mm"月"dd"日"')

        # ✅ 4. 处理 ISO8601 格式（含Z或时区）
        if re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(Z|[+-]\d{2}:?\d{2})?$", s):
            return (True, "yyyy-mm-ddThh:mm:ss")

        # ✅ 5. 处理英文月份（如 "October 24th, 2025"）
        s_clean = re.sub(r"(st|nd|rd|th)", "", s, flags=re.IGNORECASE)
        try:
            datetime.strptime(s_clean, "%B %d, %Y")
            return (True, "mmmm dd, yyyy")
        except ValueError:
            try:
                datetime.strptime(s_clean, "%b %d, %Y")
                return (True, "mmm dd, yyyy")
            except ValueError:
                pass

        return (False, "")

    @staticmethod
    def hex_to_bgr(color: str) -> int:
        """
        将十六进制颜色转换为 BGR 整数，莫名其妙要这个格式，而不是rgb格式。

        参数:
        - color: str, 颜色值，支持十六进制或常用颜色名

        返回:
        - int, BGR 整数
        """
        # 统一小写
        color = color.lower().strip()

        # 解析十六进制颜色
        if color.startswith("#") and len(color) == 7:
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
            bgr_value = (b << 16) | (g << 8) | r
            return bgr_value
        else:
            logger.warning(f"不支持的颜色格式: {color}")
            raise ValueError(f"不支持的颜色格式: {color}")

    @staticmethod
    def normalize_row_lengths(arr: List[List[Any]], fill: Any = "") -> List[List[Any]]:
        """
        将二维数组每一行长度补齐，使用指定填充值（默认 ""）。

        参数：
            arr: List[List[Any]]，二维数组
            fill: Any，可选，补充的默认值

        返回:
            List[List[Any]]，每行长度一致的二维数组
        """
        if not arr:
            return []

        max_len = max(len(row) for row in arr)
        return [row + [fill] * (max_len - len(row)) for row in arr]

    @staticmethod
    def remove_non_printable(s: str) -> str:
        """
        移除字符串中的不可打印字符（ASCII 码 0-31 和 127）。

        参数:
        - s: str, 输入字符串

        返回:
        - str, 移除不可打印字符后的字符串
        """
        if not s:
            return ""
        return re.sub(r"[\x00-\x1F\x7F]", "", s)
