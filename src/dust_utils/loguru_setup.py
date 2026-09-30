import sys
import os
from typing import Any
import json
import time
from loguru import logger
import loguru
import uuid
from contextlib import contextmanager, AbstractContextManager
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


def logger_divider(msg="", max_len=50, char="=", *args, **kwargs):
    """记录 DIVIDER 分隔线日志"""
    import wcwidth

    msg = str(msg).rstrip()
    if len(msg) > 0:
        msg = f" {msg} "

    msg_width = wcwidth.wcswidth(msg)
    if msg_width >= max_len:
        logger.opt(depth=1).log("DIVIDER", msg)
        return
    if msg_width >= max_len - 5:
        padding = char * (max_len - msg_width - 1)
        show_msg = f"{padding} {msg}"
    else:
        left_padding = char * 5
        right_padding = char * (max_len - msg_width - 7)
        show_msg = f"{left_padding}{msg}{right_padding}"
    # 跳过当前函数、再跳过包装函数，定位到调用的代码处
    logger.opt(depth=1).log("DIVIDER", show_msg)


def logger_object(object: dict | list, msg="变量值如下：", *args, **kwargs):
    """记录对象日志"""
    if object is None:
        logger.log(logging.INFO, "这是一个空对象", *args, stacklevel=3, **kwargs)
        return
    data = safe_to_dict(object, max_depth=5)
    # 跳过当前函数、再跳过包装函数，定位到调用的代码处
    logger.opt(depth=1).info(
        f"{msg}\n{json.dumps(data, ensure_ascii=False, indent=2, default=repr)}"
    )


def get_log_path():
    global log_file
    logger.opt(depth=1).debug(f"当前日志文件路径：{log_file}")
    return log_file


def escape_color_tags(msg) -> str:
    """把内容里的 "<" 转义为 loguru 认可的字面量写法。

    仅用于无法使用参数插值的场景（见 color_msg 的说明）。
    """
    return str(msg).replace("<", "\\<")


def color_msg(msg, color, log_type="info"):
    """按指定颜色输出日志，msg 中的 <> 一律按字面量处理。

    原理：loguru 在 colors=True 时，只对日志字符串中的 **字面部分** 解析颜色标签；
    通过 {...} 插值进来的参数走 AnsiParser.feed(value, raw=True)，
    会被原样插入、不参与标签解析（见 loguru/_colorizer.py:436）。
    所以把用户内容作为参数传入，而不是拼进字面串，就无需手工转义，
    也不会出现 "a < b > c" 里孤立 "<" 转义后残留反斜杠的问题。
    """
    log_method = getattr(
        logger.bind(color=color).opt(colors=True),
        log_type.lower(),
        None,
    )

    if not callable(log_method):
        raise ValueError(f"不支持的日志类型: {log_type}")

    # 字面部分只保留我们自己写的颜色标签，用户内容走参数插值
    log_method("<fg " + str(color) + ">{}</>", msg)


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


# =========================
# 子任务日志路由
# =========================
# 每个 bind_log 上下文都会把标记写入日志记录的 extra：
#   _bind_log_id        -> 当前（最内层）路由 id
#   _bind_log_main_log  -> 该路由是否仍保留在主日志
# 子任务文件 sink 只放行与本路由 id 一致的记录；主日志 sink 则拦截
# 所有标记为「不写主日志」的记录（见 _make_main_filter）。


def _make_route_filter(route_id: str):
    """生成子任务文件 sink 的过滤器：只放行当前上下文的日志。"""

    def route_filter(record) -> bool:
        return record["extra"].get("_bind_log_id") == route_id

    return route_filter


def _make_main_filter(base_filter=None):
    """包装主日志 sink 的过滤器，避免子任务日志重复写入主日志。

    只有 bind_log(..., main_log=False)（默认值）开启的上下文会被拦截；
    未进入任何 bind_log 上下文的普通日志不受影响。
    """

    def main_filter(record) -> bool:
        # 当前上下文显式声明“不写主日志”则拦截
        if record["extra"].get("_bind_log_main_log") is False:
            return False

        return True if base_filter is None else bool(base_filter(record))

    return main_filter


@contextmanager
def bind_log(
    file_path: str,
    context: dict | None = None,
    main_log: bool = False,
    *,
    format: str | None = None,
):
    """
    将当前上下文中的日志输出到指定文件，实现子任务日志分离。

    Args:
        file_path: 日志文件路径（相对路径按当前工作目录解析）
        context: 绑定到日志记录中的上下文，例如 {"task_id": "xxx"}
        main_log: 是否同时输出到主日志，默认 False（子任务日志不污染主日志）
        format: 文件日志格式，为 None 时使用默认格式

    说明:
        - 日志按「最内层」上下文路由：嵌套调用时，内层的所有日志只写内层
          文件，外层文件在内层期间不再接收日志，内层退出后自动恢复。
        - 主日志 sink 由 setup_loguru() 统一加上过滤，因此 main_log=False
          时才真正生效；自行 logger.add 的 sink 不受本参数控制。
    """
    # 归一化为绝对路径，避免结果随当前工作目录漂移
    file_path = os.path.abspath(file_path)
    file_dir = os.path.dirname(file_path)
    if file_dir:
        os.makedirs(file_dir, exist_ok=True)

    route_id = uuid.uuid4().hex

    context = {
        **(context or {}),
        "_bind_log_id": route_id,
        "_bind_log_main_log": main_log,
    }

    if format is None:
        base_folder, format = get_pack_config()

    sink_id = logger.add(
        file_path,
        encoding="utf-8",
        backtrace=False,
        filter=_make_route_filter(route_id),
        format=format,
    )

    try:
        with logger.contextualize(**context):
            yield logger
    finally:
        logger.remove(sink_id)


def setup_loguru(log_folder="logs", disabled_list=[], file_name: str = ""):
    """ """

    base_folder, stdout_format = get_pack_config()

    # 与主日志文件同根目录，避免创建目录和写文件落在不同位置
    log_dir = os.path.join(base_folder, log_folder)
    os.makedirs(log_dir, exist_ok=True)

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

    if sys.stdout:
        logger.add(
            sys.stdout,
            colorize=True,
            backtrace=False,
            format=stdout_format,
            filter=filter_lambda,
        )

    if not file_name:
        file_name = time.strftime("%Y%m%d_%H%M%S")

    global log_file
    log_file = os.path.join(base_folder, log_folder, f"{file_name}.log")

    # 文件输出（自动切割）
    # filter 包装后：标记 main_log=False 的子任务日志不会再写进主日志，
    # 只落在各自 bind_log 指定的文件里
    logger.add(
        log_file,
        rotation="10 MB",
        retention="7 days",
        encoding="utf-8",
        backtrace=False,
        filter=_make_main_filter(filter_lambda),
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
    logger.bind_log = bind_log

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

    def divider(self, msg, max_len=50, char="=", *args, **kwargs): ...
    def object(self, object, msg, *args, **kwargs): ...
    def get_log_path(self) -> str: ...
    def color_msg(self, msg, color, log_type="info"): ...
    def bind_log(
        self,
        file_path: str,
        context: dict[str, Any] | None = None,
        main_log: bool = False,
        *,
        format: str | None = None,
        encoding: str = "utf-8",
    ) -> AbstractContextManager: ...


# 👉 关键：不改 import 方式，但增强 IDE
logger = cast(LoggerExtension, logger)
