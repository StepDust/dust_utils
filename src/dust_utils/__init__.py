"""
dust-utils
通用工具库（按需加载）
"""

__version__ = "0.1.1"

# 只暴露“完全无重依赖”的工具
from .loguru_setup import setup_loguru

__all__ = [
    "setup_loguru",
]
