import importlib.util

import importlib.metadata  # 新增导入
import sys
import os
import inspect
import logging

# 配置日志
logger = logging.getLogger(__name__)


class PipUtils:
    @staticmethod
    def load_module(name, rel_path):
        """
        动态加载Python模块的工具函数。

        该函数支持在打包环境和开发环境下动态加载Python模块文件。它会根据运行环境自动判断
        正确的基础路径，并使用Python的importlib机制来加载模块。加载后的模块会被添加到
        sys.modules中，使其可以被其他代码import。

        Args:
            name (str): 模块别名，例如 "auto_packager"
            rel_path (str): 相对路径，例如 "02.自动录屏截屏打包/auto_packager.py"

        Returns:
            module: 加载完成的模块对象

        Example:
            auto_packager = PipUtils.load_module(
                "auto_packager", "02.自动录屏截屏打包/auto_packager.py"
            )
        """

        # 检查是否在打包环境中运行
        if getattr(sys, "frozen", False):
            # 打包后 exe
            base_path = sys._MEIPASS
        else:
            # 调试环境
            base_path = os.path.abspath(".")  # 当前项目根目录

        # 构建模块文件的完整路径
        file_path = os.path.join(base_path, rel_path)

        # 从文件路径创建模块规范对象
        spec = importlib.util.spec_from_file_location(name, file_path)

        # 根据规范创建新的模块对象
        module = importlib.util.module_from_spec(spec)

        # 将模块添加到sys.modules中以便能够被import语句引用
        sys.modules[name] = module

        # 执行模块代码
        spec.loader.exec_module(module)

        # 返回加载好的模块对象
        return module

    @staticmethod
    def is_development_mode():
        """
        判断当前是否在开发模式下运行。

        该方法会检查脚本是否在打包环境中运行。如果在打包环境中运行（如使用pyinstaller打包），
        则返回False；如果在调试环境中运行，则返回True。

        Returns:
            bool: 如果在开发模式下运行（调试环境），则返回True；否则返回False。
        """
        # 检查是否在打包环境中运行
        return not getattr(sys, "frozen", False)

    @staticmethod
    def get_base_path():
        """
        获取当前脚本的基础路径。

        该方法会根据脚本是否在打包环境中运行，返回不同的基础路径。
        如果在打包环境中运行（如使用pyinstaller打包），则返回sys._MEIPASS；
        如果在调试环境中运行，则返回当前脚本所在目录的绝对路径。

        Returns:
            str: 当前脚本的基础路径
        """
        # 检查是否在打包环境中运行
        if not PipUtils.is_development_mode():
            # 打包后的 exe，__file__ 指向 exe 路径
            return os.path.dirname(sys.executable)
        else:
            # 开发环境
            return os.path.dirname(
                os.path.abspath(inspect.stack()[1].filename)
            )  # 调用者所在目录

    @staticmethod
    def check_pip_module(module_name: str):
        """
        检查指定模块是否安装：
        - 已安装 → 输出版本号
        - 未安装 → 输出提示并退出程序
        """
        spec = importlib.util.find_spec(module_name)
        if spec is None:
            logger.error(
                f"模块 '{module_name}' 未安装，请执行: pip install {module_name}"
            )
            sys.exit(1)
        else:
            try:
                version = importlib.metadata.version(module_name)
                logger.debug(
                    f"模块 '{module_name}' 已安装，版本: {version}（位于: {spec.origin}）"
                )
            except importlib.metadata.PackageNotFoundError:
                # 极少数情况：find_spec 找到但 metadata 没版本（如内置模块或特殊安装）
                logger.debug(
                    f"模块 '{module_name}' 已安装（位于: {spec.origin}），但无法获取版本信息"
                )

    @staticmethod
    def get_running_name(has_suffix=True):
        """
        获取当前运行的脚本名称。

        该方法会根据脚本是否在打包环境中运行，返回不同的脚本名称。
        如果在打包环境中运行（如使用pyinstaller打包），则返回sys.executable的文件名；
        如果在调试环境中运行，则返回当前脚本的文件名。

        Returns:
            str: 当前运行的脚本名称
        """
        name = ""
        if not PipUtils.is_development_mode():
            # 打包后的 exe
            name = os.path.basename(sys.executable)
        else:
            # 开发环境
            name = os.path.basename(
                os.path.abspath(inspect.stack()[1].filename)
            )  # 调用者所在目录

        if has_suffix:
            return name

        name = os.path.splitext(name)[0]  # 去掉扩展名
        return name


if __name__ == "__main__":
    logger.info("开始检查PyInstaller模块")
    name = PipUtils.get_running_name(False)
    logger.info(name)
