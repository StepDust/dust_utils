from datetime import date, time, datetime, timedelta

from collections.abc import Sequence
from typing import TypeVar, Any
from loguru import logger

T = TypeVar("T")


class CommonUtils:

    # region 日期工具
    @staticmethod
    def parse_date(
        value: int | str | None,
        format: str = "yyyy-MM-dd",
        default: str | date | None = None,
    ) -> str:
        """
        日期解析，返回指定格式的日期字符串。

        支持解析：
            2026-09-20
            2026/09/20
            2026.09.20
            2026年09月20日
            2026年09月20

        负整数：
            -1  表示昨天
            -7  表示7天前

        Args:
            value: 待解析的日期。
            format: 返回的日期格式，例如：
                yyyy-MM-dd
                yyyy/MM/dd
                yyyy.MM.dd
                yyyy年MM月dd日
            default: value 为 None 时的默认日期。
                不指定时使用今天。

        Returns:
            按指定 format 格式化后的日期字符串。
        """
        import re

        if not default:
            default = date.today()

        if not value:
            return default.strftime(
                format.replace("yyyy", "%Y").replace("MM", "%m").replace("dd", "%d")
            )

        value = str(value).strip()

        # 整数：表示相对于今天偏移多少天
        if re.fullmatch(r"-?\d+", value):
            result = date.today() + timedelta(days=int(value))
        else:
            # 日期格式
            match = re.fullmatch(
                r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?",
                value,
            )

            if not match:
                raise ValueError(f"无法识别的日期格式：{value}")

            year, month, day = map(int, match.groups())

            try:
                result = date(year, month, day)
            except ValueError:
                raise ValueError(f"无效的日期：{value}")

        # yyyy 格式转 Python strftime 格式
        format = format.replace("yyyy", "%Y").replace("MM", "%m").replace("dd", "%d")

        return result.strftime(format)

    # endregion

    # region 数组工具

    @staticmethod
    def dict_list_to_2d(
        list_data: Sequence[dict[str, Any]],
        convert_complex: bool = False,
    ):
        """字典列表转换为二维数组。"""
        if not list_data:
            return []

        headers = list(list_data[0].keys())

        def convert(value: Any):
            if not convert_complex:
                return value

            if isinstance(value, datetime):
                return value.strftime("%Y-%m-%d %H:%M:%S")

            if isinstance(value, date):
                return value.strftime("%Y-%m-%d")

            if isinstance(value, time):
                return value.strftime("%H:%M:%S")

            if isinstance(value, (str, int, float, bool)) or value is None:
                return value

            return str(value)

        return [
            headers,
            *[[convert(item.get(header)) for header in headers] for item in list_data],
        ]

    # endregion

    # region 字符工具

    @staticmethod
    def mask_secret(
        value: str, show_start: int = 4, show_end: int = 4, mask: str = "****"
    ) -> str:
        """
        密钥脱敏

         Args:
            value (str): 原始密钥
            show_start (int): 保留开头字符数，默认4
            show_end (int): 保留结尾字符数，默认4
            mask (str): 脱敏字符，默认****
        """
        if not value:
            return value

        value = str(value)

        length = len(value)

        # 太短直接隐藏
        if length <= show_start + show_end:
            return mask

        return value[:show_start] + mask + value[-show_end:]

    @staticmethod
    def format_str(content: str | None, parameters: dict = {}) -> str:
        """
        格式化数据，替换字符串中的模板占位符

        Args:
            content (str): 需要格式化的字符串，支持 {{变量名}} 形式的占位符
            parameters (dict): 额外的参数字典，用于替换content中对应的占位符
                             格式为 {key: value}

        Returns:
            str: 替换占位符后的字符串
        """
        import re

        if not content:
            return ""

        def replace_var(match) -> str:
            var_name = match.group(1).strip()
            value = CommonUtils.get_value(parameters, var_name, match.group(0))

            if value is None:
                return ""

            return str(value)

        pattern = r"\{\{\s*([^{}]+?)\s*\}\}"
        # 这种替换可以兼容包含特殊字符的content
        content = re.sub(pattern, replace_var, content)

        # 返回格式化后的字符串
        return content

    # endregion

    # region dict工具

    @staticmethod
    def get_value(obj: dict, key: str, default: T = None) -> T:
        """
        根据键名获取字典中的值。

        键名匹配时忽略大小写、下划线和连字符。

        例如：
            username
            user-name
            user_name
            Username
            USER_NAME

        会被视为相同的键。

        Args:
            obj: 要查询的字典。
            key: 要查询的键名。
            default: 未找到或值为 None 时返回的默认值。

        Returns:
            找到的值，否则返回 default。
        """

        def normalize(value):
            return str(value).replace("_", "").replace("-", "").lower()

        target = normalize(key)

        for k, v in obj.items():
            if normalize(k) == target:
                return default if v is None else v

        return default

    # endregion

    # region 网络工具

    @staticmethod
    def download_file(
        url: str, save_folder: str, save_path: str = None, is_cover: bool = True
    ) -> str:
        """
        下载文件

        Args:
            url: 要查询的字典。
            save_dir: 保存的文件夹

        Returns:
            文件下载后的路径
        """
        import os
        import requests
        from urllib.parse import unquote

        os.makedirs(save_folder, exist_ok=True)

        filename = unquote(os.path.basename(url.split("?", 1)[0]))
        if not save_path:
            save_path = os.path.join(save_folder, filename)

        if os.path.exists(save_path) and not is_cover:
            return save_path

        file_info = CommonUtils.get_file_info(url)

        if not file_info.get("exists", False):
            logger.warning(f"文件不存在：{url}")
            return ""

        response = requests.get(url, timeout=60)
        response.raise_for_status()

        with open(save_path, "wb") as f:
            f.write(response.content)

        return save_path

    @staticmethod
    def get_file_info(url):

        import os
        import requests
        from urllib.parse import urlparse, unquote

        response = requests.head(
            url,
            allow_redirects=True,
            timeout=10,
        )

        headers = response.headers

        # 文件是否存在
        exists = response.status_code == 200

        # 文件名
        filename = None
        content_disposition = headers.get("Content-Disposition")

        if content_disposition:
            for part in content_disposition.split(";"):
                part = part.strip()

                if part.startswith("filename*="):
                    filename = unquote(part.split("=", 1)[1].strip().strip('"'))
                    if "''" in filename:
                        filename = filename.split("''", 1)[1]
                    break

                if part.startswith("filename="):
                    filename = part.split("=", 1)[1].strip().strip('"')
                    break

        if not filename:
            filename = os.path.basename(unquote(urlparse(response.url).path))

        content_length = headers.get("Content-Length")
        size = int(content_length) if content_length else None

        return {
            "exists": exists,
            "status_code": response.status_code,
            "filename": filename or None,
            "size": size,
            "size_mb": round(size / 1024 / 1024, 2) if size is not None else None,
            "content_type": headers.get("Content-Type"),
            "last_modified": headers.get("Last-Modified"),
            "etag": headers.get("ETag"),
            "accept_ranges": headers.get("Accept-Ranges"),
            "content_disposition": content_disposition,
            "url": url,
            "final_url": response.url,
        }

    # endregion
