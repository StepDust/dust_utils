import os
import shutil
import zipfile
import logging

logger = logging.getLogger(__name__)


class ZipUtils:
    """
    文件工具类，提供文件和目录操作功能
    """

    # 解压压缩包
    @staticmethod
    def extract_zip(zip_path, extract_path, is_delete=False):
        """
        解压zip格式的压缩文件到指定目录

        Args:
            zip_path (str): zip压缩文件的路径
            extract_path (str): 解压目标路径
            is_delete (bool, optional): 是否删除目标路径已存在的文件。默认为False

        Raises:
            FileNotFoundError: 当zip文件不存在时抛出此异常
            Exception: 当删除目标目录失败时抛出相应异常

        Returns:
            None

        功能说明:
        1. 检查压缩文件是否存在
        2. 根据is_delete参数决定是否删除已存在的目标目录
        3. 创建解压目标目录
        4. 解压文件
        5. 如果解压后只有一个文件夹，会将该文件夹中的内容移动到目标目录
        """
        # 第一步：验证输入文件的存在性
        if not os.path.exists(zip_path):
            logger.error(f"压缩文件 {zip_path} 不存在")
            raise FileNotFoundError(f"压缩文件 {zip_path} 不存在")

        # 第二步：处理目标路径
        # 如果设置了删除标志且目标路径存在，则删除目标路径
        if is_delete and os.path.exists(extract_path):
            try:
                shutil.rmtree(extract_path)
            except Exception as e:
                logger.error(f"删除目录 {extract_path} 失败: {str(e)}")
                raise
        # 确保目标路径存在
        os.makedirs(extract_path, exist_ok=True)

        # 第三步：解压文件处理
        # 使用zipfile打开文件时指定编码格式，避免中文乱码
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            # 遍历压缩包内的所有文件
            for file in zip_ref.namelist():
                # 处理文件名编码问题，优先尝试GBK解码
                try:
                    filename = file.encode("cp437").decode("gbk")
                except UnicodeEncodeError:
                    # 解码失败时保持原文件名
                    filename = file

                # 处理目标路径上的已存在文件
                target_path = os.path.join(extract_path, filename)
                if os.path.exists(target_path):
                    # 根据文件类型选择删除方式
                    if os.path.isdir(target_path):
                        shutil.rmtree(target_path)
                    else:
                        os.remove(target_path)

                # 执行解压操作
                zip_ref.extract(file, extract_path)
                # 如果文件名发生了变化，进行重命名
                if file != filename:
                    os.rename(
                        os.path.join(extract_path, file),
                        os.path.join(extract_path, filename),
                    )

        # 第四步：优化解压结果的目录结构
        # 递归处理单一子目录的情况
        def flatten_single_dir(dir_path):
            items = os.listdir(dir_path)
            # 如果目录中只有一个子目录，则继续处理
            if len(items) == 1 and os.path.isdir(os.path.join(dir_path, items[0])):
                source_dir = os.path.join(dir_path, items[0])
                # 将子目录中的所有内容移动到父目录
                for item in os.listdir(source_dir):
                    source_item = os.path.join(source_dir, item)
                    dest_item = os.path.join(dir_path, item)
                    shutil.move(source_item, dest_item)
                # 清理空的子目录
                os.rmdir(source_dir)
                # 递归处理，以防还有更深层的单一子目录
                flatten_single_dir(dir_path)

        # 开始处理解压目录
        flatten_single_dir(extract_path)

        # 记录操作完成的日志
        logger.info(f"已解压文件到 {extract_path}")

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
                                arc_dir = os.path.join(rel_root, dir_name)
                                zip_ref.write(dir_path, arc_dir)  # 写入空目录

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
