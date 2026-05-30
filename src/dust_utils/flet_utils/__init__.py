try:
    import flet
except ImportError:
    raise RuntimeError("fletUtils 需要 flet库")

from .flet_utils import FletUtils, LogConsole

__all__ = ["FletUtils", "LogConsole"]
