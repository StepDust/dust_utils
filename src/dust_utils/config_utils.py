import os
import sys
from dataclasses import fields, is_dataclass, dataclass
from .file_utils import PipUtils
import json
from typing import TypeVar

T = TypeVar("T")


class ConfigUtils:

    @staticmethod
    def from_dict(target_class, data: dict):
        """从字典加载配置"""
        result = {}

        for field in fields(target_class):
            value = data[field.name]
            field_type = field.type

            if is_dataclass(field_type):
                value = ConfigUtils.from_dict(field_type, value)

            result[field.name] = value

        return target_class(**result)

    @staticmethod
    def load_config(target_class: type[T], config_path) -> T:
        """加载配置文件，映射为指定类型的对象"""

        if not os.path.exists(config_path):
            raise FileNotFoundError(f"配置文件不存在: {config_path}")

        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        return ConfigUtils.from_dict(target_class, data)

    @staticmethod
    def resolve_path(path: str) -> str:
        """把配置里的相对路径解析成绝对路径。

        基准是 base_path（正常运行时=项目目录，PyInstaller 打包后=exe 所在目录），
        这样从任意工作目录启动服务，temp/ 和 logs/ 都不会跑偏。
        """
        if not path:
            return path
        if os.path.isabs(path):
            return path
        return os.path.join(PipUtils.get_base_path(), path)
