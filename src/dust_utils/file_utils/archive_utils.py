import os
import shutil
import zipfile
import tarfile
import gzip
from pathlib import Path

from loguru import logger


class ArchiveUtils:
    """
    档案工具类
    """

    # region  解压压缩包

    @staticmethod
    def extract_archive(archive_path, extract_path="", is_delete=False):
        """
        解压压缩文件到指定目录，支持 zip、gz、tar.gz(tgz) 格式

        Args:
            archive_path (str): 压缩文件路径
            extract_path (str): 解压目标路径，默认为空（解压到压缩包同名目录）
            is_delete (bool): 是否在解压前删除已存在的目标路径

        Raises:
            FileNotFoundError: 压缩文件不存在
            ValueError: 不支持的文件格式
            Exception: 其他解压异常

        Returns:
            None
        """
        logger.debug(f"解压文件：{archive_path}")
        # 1. 检查文件是否存在
        if not os.path.exists(archive_path):
            logger.error(f"压缩文件 {archive_path} 不存在")
            raise FileNotFoundError(f"压缩文件 {archive_path} 不存在")

        # 2. 确定解压目标目录
        if not extract_path:
            # 对于 tar.gz/tgz 需要去掉双重后缀，gz 只去掉 .gz，zip 去掉 .zip
            p = Path(archive_path)
            if archive_path.endswith((".tar.gz", ".tgz")):
                # 去掉 .tar.gz 或 .tgz
                name = p.name
                if name.endswith(".tar.gz"):
                    base = name[:-7]
                else:  # .tgz
                    base = name[:-4]
                extract_path = str(p.parent / base)
            elif archive_path.endswith(".gz"):
                # 单文件 gz 也解压到同名目录（保持行为一致）
                extract_path = str(p.with_suffix(""))
            else:
                # 默认按 zip 处理
                extract_path = str(p.with_suffix(""))

        # 3. 处理删除标志
        if is_delete and os.path.exists(extract_path):
            try:
                shutil.rmtree(extract_path)
            except Exception as e:
                logger.error(f"删除目录 {extract_path} 失败: {str(e)}")
                raise

        os.makedirs(extract_path, exist_ok=True)

        # 4. 根据文件真实内容特征解压
        if zipfile.is_zipfile(archive_path):
            ArchiveUtils._extract_zip(archive_path, extract_path)
        elif tarfile.is_tarfile(archive_path):
            ArchiveUtils._extract_tar(archive_path, extract_path)
        else:
            # 检查是否为 gzip 压缩（单文件或未识别的tar.gz）
            try:
                with gzip.open(archive_path, "rb") as f:
                    f.read(1)  # 尝试读取1字节，若不报错则说明是 gzip 格式
                ArchiveUtils._extract_gz_single(archive_path, extract_path)
            except Exception:
                raise ValueError(f"不支持或已损坏的压缩格式: {archive_path}")

        ArchiveUtils.remove_empty_folders(extract_path)
        logger.debug(f"已解压至：{extract_path}")
        return extract_path

    @staticmethod
    def _extract_zip(zip_path, extract_path):
        """解压 zip 文件，处理文件名编码"""
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            for file in zip_ref.namelist():
                try:
                    filename = file.encode("cp437").decode("gbk")
                except UnicodeEncodeError:
                    filename = file

                zip_ref.extract(file, extract_path)
                # 重命名解码后的文件名
                if file != filename:
                    os.rename(
                        os.path.join(extract_path, file),
                        os.path.join(extract_path, filename),
                    )

    @staticmethod
    def _extract_tar(tar_path, extract_path):
        """
        解压 tar / tar.gz / tgz / tar.bz2 / tar.xz

        特性：
        - 自动识别压缩格式
        """

        # tarfile.open 的 mode='r:*' 会自动根据文件头识别压缩格式
        with tarfile.open(tar_path, mode="r:*") as tar:
            tar.extractall(path=extract_path)

    @staticmethod
    def _extract_gz_single(gz_path, extract_path):
        """解压 .gz 单文件，文件名为去掉 .gz 的部分"""

        with gzip.open(gz_path, mode="rb") as f_in:
            with open(extract_path, "wb") as f_out:
                shutil.copyfileobj(f_in, f_out)

    # endregion

    @staticmethod
    def zip_add_files(zip_path: str, files: list, is_repeat_skip: bool = True):
        """
        向已存在的zip文件中添加文件

        Args:
            zip_path str: zip文件的路径
            files list: 要添加的文件路径列表，每个元素可以是字符串路径或者(文件路径, zip内路径)的元组
            is_repeat_skip bool: 重复是否跳过，默认为True，跳过已存在的文件；为False时会先删除已存在的同名文件再添加

        Raises:
            FileNotFoundError: 当zip文件不存在时抛出此异常
            ValueError: 当files参数格式不正确时抛出此异常

        Returns:
            None
        """
        # 检查zip文件是否存在
        if not os.path.exists(zip_path):
            logger.error(f"zip文件 {zip_path} 不存在")
            raise FileNotFoundError(f"zip文件 {zip_path} 不存在")

        # 以追加模式打开zip文件
        with zipfile.ZipFile(zip_path, "a") as zip_ref:
            for file_item in files:
                # 处理输入参数，支持字符串路径或元组格式
                if isinstance(file_item, tuple):
                    file_path, arcname = file_item
                elif isinstance(file_item, str):
                    file_path = file_item
                    arcname = os.path.basename(file_path)
                else:
                    raise ValueError(
                        "files列表中的元素必须是字符串路径或(文件路径, zip内路径)的元组"
                    )

                # 检查要添加的文件是否存在
                if not os.path.exists(file_path):
                    logger.warning(f"要添加的文件 {file_path} 不存在，已跳过")
                    continue

                # 根据is_repeat_skip参数处理已存在的文件
                if arcname in zip_ref.namelist():
                    if is_repeat_skip:
                        logger.info(f"zip中已存在文件 {arcname}，已跳过")
                        continue
                    else:
                        # 先删除已存在的文件
                        try:
                            # 创建临时zip，排除要删除的文件
                            temp_zip_path = zip_path + ".tmp"
                            with zipfile.ZipFile(temp_zip_path, "w") as temp_zip:
                                for item in zip_ref.infolist():
                                    if item.filename != arcname:
                                        temp_zip.writestr(
                                            item, zip_ref.read(item.filename)
                                        )
                            # 替换原zip
                            zip_ref.close()
                            os.replace(temp_zip_path, zip_path)
                            # 重新打开zip文件以继续添加
                            zip_ref = zipfile.ZipFile(zip_path, "a")
                        except Exception as e:
                            logger.error(
                                f"删除zip中已存在文件 {arcname} 失败: {str(e)}"
                            )
                            raise

                try:
                    # 将文件添加到zip中
                    zip_ref.write(file_path, arcname)
                    # logger.info(f"已将文件 {file_path} 添加到zip中，存储为 {arcname}")
                except Exception as e:
                    logger.error(f"添加文件 {file_path} 到zip失败: {str(e)}")
                    raise

    @staticmethod
    def zip_folder(source_path: str, zip_path: str, is_delete: bool = False):
        """
        将单个文件或整个文件夹打包成 zip 压缩包
        - 压缩文件夹时：文件夹内的内容直接放在 zip 根目录（不额外嵌套一层文件夹）
        - 压缩单个文件时：文件直接放在 zip 根目录

        Args:
            source_path (str): 要压缩的源路径（可以是单个文件或文件夹）
            zip_path (str): 输出 zip 文件的完整路径（含文件名）
            is_delete (bool, optional): 如果目标 zip 已存在，是否先删除。默认为 False

        Raises:
            FileNotFoundError: 当 source_path 不存在时
        """
        if not os.path.exists(source_path):
            logger.error(f"要压缩的路径不存在: {source_path}")
            raise FileNotFoundError(f"要压缩的路径不存在: {source_path}")

        # 处理目标 zip 已存在的情况
        if os.path.exists(zip_path):
            if is_delete:
                try:
                    os.remove(zip_path)
                    logger.info(f"已删除已存在的 zip 文件: {zip_path}")
                except Exception as e:
                    logger.error(f"删除已存在 zip 文件失败: {e}")
                    raise
            else:
                logger.warning(f"目标 zip 文件已存在，将被覆盖: {zip_path}")

        # 确保输出目录存在
        os.makedirs(os.path.dirname(zip_path), exist_ok=True)

        try:
            with zipfile.ZipFile(
                zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
            ) as zip_ref:
                source_path = os.path.abspath(source_path)

                if os.path.isfile(source_path):
                    # ==================== 单个文件 ====================
                    arcname = os.path.basename(source_path)
                    zip_ref.write(source_path, arcname)
                    logger.info(f"已打包单个文件: {arcname} → {zip_path}")

                elif os.path.isdir(source_path):
                    # ==================== 整个文件夹（扁平化处理） ====================
                    logger.info(
                        f"正在打包文件夹（内容直接置于 zip 根目录）: {source_path}"
                    )

                    for root, dirs, files in os.walk(source_path):
                        # 计算相对路径（去掉 source_path 这一层）
                        rel_root = os.path.relpath(root, source_path)

                        # 添加目录（保留子目录结构，但不在最外层多一层）
                        if rel_root != ".":
                            for dir_name in dirs:
                                dir_path = os.path.join(root, dir_name)
                                arc_folder = os.path.join(rel_root, dir_name)
                                zip_ref.write(dir_path, arc_folder)  # 写入空目录

                        # 添加文件
                        for file in files:
                            file_path = os.path.join(root, file)
                            # 如果在根目录，arcname 就是文件名；否则保留相对子路径
                            arcname = (
                                file
                                if rel_root == "."
                                else os.path.join(rel_root, file)
                            )
                            zip_ref.write(file_path, arcname)

                    logger.info(f"文件夹内容已直接打包到 zip 根目录 → {zip_path}")
                else:
                    raise ValueError(f"不支持的路径类型: {source_path}")

        except Exception as e:
            logger.error(f"打包 zip 失败: {str(e)}")
            if os.path.exists(zip_path):
                try:
                    os.remove(zip_path)
                except:
                    pass
            raise

        logger.success(f"压缩完成 → {zip_path}")
        return zip_path

    @staticmethod
    def remove_empty_folders(folder_path):
        items = os.listdir(folder_path)
        # 如果目录中只有一个子目录，则继续处理
        if len(items) == 1 and os.path.isdir(os.path.join(folder_path, items[0])):
            source_dir = os.path.join(folder_path, items[0])
            # 将子目录中的所有内容移动到父目录
            for item in os.listdir(source_dir):
                source_item = os.path.join(source_dir, item)
                dest_item = os.path.join(folder_path, item)
                shutil.move(source_item, dest_item)
            # 清理空的子目录
            os.rmdir(source_dir)
            # 递归处理，以防还有更深层的单一子目录
            ArchiveUtils.remove_empty_folders(folder_path)

    @staticmethod
    def _normalize_suffix_list(suffix_list):
        """
        标准化后缀列表

        支持：
        - .py
        - *.py
        - py
        - .d.ts
        - *.d.ts
        - 带空格
        """

        result = set()

        for item in suffix_list:
            if not item:
                continue

            let_suffix = str(item).strip().lower()

            # 去掉 *
            if let_suffix.startswith("*"):
                let_suffix = let_suffix[1:]

            # 补 .
            if not let_suffix.startswith("."):
                let_suffix = "." + let_suffix

            result.add(let_suffix)

        return result

    @staticmethod
    def find_files(folder_path, suffix_list, is_recursive: bool = False) -> list[Path]:
        """
        遍历指定文件夹下的文件

        :param folder_path: 文件夹路径
        :param suffix_list: 后缀列表，例如 ['.py', '.js', '.vue', '.d.ts']
        :param is_recursive: 是否递归子目录
        :return: 文件路径列表
        """

        let_path = Path(folder_path)
        let_suffix_set = ArchiveUtils._normalize_suffix_list(suffix_list)

        result = []

        # 选择遍历方式
        if is_recursive:
            iterator = let_path.rglob("*")
        else:
            iterator = let_path.iterdir()

        for file in iterator:

            if not file.is_file():
                continue

            let_name = file.name.lower()

            if any(let_name.endswith(suffix) for suffix in let_suffix_set):
                result.append(file)

        return result
