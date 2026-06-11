import os
import re
import win32com.client as win32
import pythoncom

from loguru import logger


def auto_before_call(before_func, skip_names=None):
    """
    类装饰器：为类中所有非私有“实例方法”自动加上调用指定前置方法的逻辑。
    静态函数、类方法将被跳过。

    参数：
        before_func: str, 前置方法名（例如 "check"、"ensure_ready"）
        skip_names: set[str]，可选，跳过的函数名集合
    """

    def decorator(cls):
        if skip_names is None:
            skips = {"close", "quit", before_func}
        else:
            skips = set(skip_names) | {before_func}

        for name, method in list(cls.__dict__.items()):
            # 仅处理可调用、非私有、不在跳过名单中的实例方法
            if (
                callable(method)
                and not name.startswith("__")
                and name not in skips
                and not isinstance(method, (staticmethod, classmethod))
            ):

                def make_wrapper(m=method):
                    def wrapper(self, *args, **kwargs):
                        # 获取前置函数
                        before = getattr(self, before_func, None)
                        if callable(before):
                            before()
                        return m(self, *args, **kwargs)

                    return wrapper

                setattr(cls, name, make_wrapper())

        return cls

    return decorator


class WPSBase:
    def __init__(self, file_path: str, prog_id: str, is_debug: bool = False):
        """
        初始化 OfficeBase 类，启动指定的 office 应用程序。

        参数:
        - file_path: str, Office 文件路径
        - prog_id: str, 要启动的应用程序 ID，例如 "Ket.Application" 或 "Excel.Application"
        """
        # 初始化 COM 库，确保在多线程环境下正常工作
        pythoncom.CoInitialize()
        self.office = None
        # 设置 logger 级别
        # logger.setLevel(logging.DEBUG if is_debug else logging.INFO)
        try:
            self.office = win32.Dispatch(prog_id)
            self.office.Visible = False  # 是否展示窗口，调试时可改为 True
            self.office.DisplayAlerts = False  # 关闭警报弹窗
            self.file_path = os.path.normpath(os.path.abspath(file_path))
            logger.info(f"Office启动成功：{prog_id}")
        except Exception as e:
            logger.error(f"Office启动失败 {prog_id}: {str(e)}")
            raise RuntimeError(f"Office启动失败 ：{prog_id}")

    def __del__(self):
        """
        析构函数，确保在对象销毁时关闭 Office 应用程序。
        """
        if self.office:
            self.quit()

    def quit(self):
        """
        关闭 Office 应用程序并释放资源。
        """
        if self.office is None:
            logger.debug("Office 对象已为空，无需再次关闭")
            return
        try:
            self.office.Quit()
        except Exception as e:
            logger.error(
                f"关闭Office时出错: {str(e)}",
            )
        finally:
            self.office = None
            pythoncom.CoUninitialize()  # 释放 COM 库资源
            logger.info("Office资源已释放")
