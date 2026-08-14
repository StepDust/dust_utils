import sys
import os
from typing import Any
import json
import time
from loguru import logger
import loguru

from typing import Protocol, cast


def safe_to_dict(
    obj: Any, seen: set | None = None, max_depth: int = 3, current_depth: int = 0
) -> Any:
    """
    把对象转成可序列化的 dict，支持嵌套、列表、循环引用检测
    """
    if seen is None:
        seen = set()

    obj_id = id(obj)
    # print(f"对象：{obj}\t深度：{current_depth}")
    if obj_id in seen:
        return f"<循环引用: {type(obj).__name__}>"

    if current_depth > max_depth * 2 + 1:
        return f"<深度超出 {max_depth}>"

    seen.add(obj_id)

    if isinstance(obj, (list, tuple)):
        return [safe_to_dict(x, seen.copy(), max_depth, current_depth + 1) for x in obj]

    if isinstance(obj, dict):
        return {
            k: safe_to_dict(v, seen.copy(), max_depth, current_depth + 1)
            for k, v in obj.items()
        }

    if hasattr(obj, "__dict__") and not isinstance(obj, type):
        d = {}
        for k, v in vars(obj).items():
            d[k] = safe_to_dict(v, seen.copy(), max_depth, current_depth + 1)
        return {**d, "__class__": obj.__class__.__name__}

    if hasattr(obj, "_asdict"):  # dataclass / namedtuple
        return safe_to_dict(obj._asdict(), seen.copy(), max_depth, current_depth + 1)

    return obj


def logger_divider(message="", max_len=50, char="=", *args, **kwargs):
    """记录 DIVIDER 分隔线日志"""
    import wcwidth

    message = message.rstrip()
    if len(message) > 0:
        message = f" {message} "

    msg_width = wcwidth.wcswidth(message)
    if msg_width >= max_len:
        logger.opt(depth=1).log("DIVIDER", message)
        return
    if msg_width >= max_len - 5:
        padding = char * (max_len - msg_width - 1)
        show_msg = f"{padding} {message}"
    else:
        left_padding = char * 5
        right_padding = char * (max_len - msg_width - 7)
        show_msg = f"{left_padding}{message}{right_padding}"
    # 跳过当前函数、再跳过包装函数，定位到调用的代码处
    logger.opt(depth=1).log("DIVIDER", show_msg)


def logger_object(object: dict | list, message="变量值如下：", *args, **kwargs):
    """记录对象日志"""
    if object is None:
        self.log(logging.INFO, "这是一个空对象", *args, stacklevel=3, **kwargs)
        return
    data = safe_to_dict(object, max_depth=5)
    # 跳过当前函数、再跳过包装函数，定位到调用的代码处
    logger.opt(depth=1).info(
        f"{message}\n{json.dumps(data, ensure_ascii=False, indent=2, default=repr)}"
    )


def get_log_path():
    global log_file
    logger.opt(depth=1).debug(f"当前日志文件路径：{log_file}")
    return log_file


def color_msg(msg, color, log_type="info"):
    log_method = getattr(
        logger.bind(color=color).opt(colors=True),
        log_type.lower(),
        None,
    )

    if not callable(log_method):
        raise ValueError(f"不支持的日志类型: {log_type}")

    log_method(f"<fg {color}>{msg}</>")


def get_pack_config():
    # 是否打包（PyInstaller）
    is_frozen = getattr(sys, "frozen", False)

    if is_frozen:
        # ✔ 打包环境：exe 所在目录
        base_folder = os.path.dirname(sys.executable)
        format_rule = (
            "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
            "<level>{level: <8}</level> | "
            "<cyan>{name}</cyan>:"
            "<cyan>{function}</cyan>:"
            "<cyan>{line}</cyan> - "
            "<level>{message}</level>"
        )
    else:
        # ✔ 开发环境：项目根目录（而不是 cwd）
        base_folder = os.path.dirname(os.path.abspath(sys.argv[0]))
        format_rule = (
            "<green>{time:HH:mm:ss}</green> | "
            "<level>{level: <8}</level> | "
            "<level>{message}</level>"
        )

    return base_folder, format_rule


def setup_loguru(log_folder="logs", disabled_list=[]):

    os.makedirs(log_folder, exist_ok=True)

    logger.remove()

    # 自定义日志级别
    logger.level(
        "DIVIDER",
        no=15,
        color="<white>",
    )

    filter_list = [
        "requests",
        "urllib3",
        "chardet",
        "charset_normalizer",
        "httpcore._backends.sync",
        "httpx._client",
        "httpx",
        "openai",
        "openrouter",
        "stainless",
        "httpcore",
        "pywebview",
        "webview",
    ]

    filter_list.extend(disabled_list)
    filter_tuple = tuple(filter_list)

    # 用于过滤日志记录器，排除掉不需要的库的日志输出
    filter_lambda = lambda record: not (record["name"] or "").startswith(filter_tuple)

    base_folder, stdout_format = get_pack_config()

    if sys.stdout:
        logger.add(
            sys.stdout,
            colorize=True,
            backtrace=False,
            format=stdout_format,
            filter=filter_lambda,
        )

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    global log_file
    log_file = os.path.join(base_folder, log_folder, f"{timestamp}.log")

    # 文件输出（自动切割）
    logger.add(
        log_file,
        rotation="10 MB",
        retention="7 days",
        encoding="utf-8",
        backtrace=False,
        filter=filter_lambda,
        format=(
            "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
            "<level>{level: <8}</level> | "
            "<cyan>{name}</cyan>:"
            "<cyan>{function}</cyan>:"
            "<cyan>{line}</cyan> - "
            "<level>{message}</level>"
        ),
    )
    logger.divider = logger_divider
    logger.object = logger_object
    logger.get_log_path = get_log_path
    logger.color_msg = color_msg

    # 核心语法，将自定义的logger对象赋值给 loguru.logger
    # 这样在其他模块中直接使用 loguru.logger 就能获得增强功能，同时保持原有的 import 方式不变
    loguru.logger = logger
    return logger


# =========================
# IDE 类型提示（关键）
# =========================
class LoggerExtension(Protocol):
    def debug(self, msg, *args, **kwargs): ...
    def info(self, msg, *args, **kwargs): ...
    def warning(self, msg, *args, **kwargs): ...
    def error(self, msg, *args, **kwargs): ...
    def success(self, msg, *args, **kwargs): ...
    def critical(self, msg, *args, **kwargs): ...

    def divider(self, msg, *args, **kwargs): ...
    def object(self, object, msg="变量值如下：", *args, **kwargs): ...
    def get_log_path(self) -> str: ...
    def color_msg(self, msg, color, log_type="info") -> str: ...


# 👉 关键：不改 import 方式，但增强 IDE
logger = cast(LoggerExtension, logger)
