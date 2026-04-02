import os
import logging
from typing import List, Optional

# 配置日志
logger = logging.getLogger(__name__)


class OSSUtils:
    """
    阿里云 OSS 工具类（使用 logging）
    支持文件和文件夹的上传、下载、删除操作
    """

    def __init__(
        self,
        access_key_id: str,
        access_key_secret: str,
        endpoint: str,
        bucket_name: str,
        is_cname=False,
    ):
        """
        初始化 OSS 客户端

        :param access_key_id: 阿里云 AccessKey ID
        :param access_key_secret: 阿里云 AccessKey Secret
        :param endpoint：
        :param bucket_name：
        :param is_cname：
        """
        import oss2

        endpoint = endpoint.strip().replace("\n", "").replace("\r", "")
        if not endpoint.startswith("http"):
            endpoint = "http://" + endpoint

        auth = oss2.Auth(access_key_id, access_key_secret)
        self.bucket = oss2.Bucket(
            auth=auth,
            endpoint=endpoint,
            bucket_name=bucket_name,
            is_cname=is_cname,
        )
        # self.bucket.session.mount("https://", oss2.adapters.HTTPAdapter(max_retries=3))
        self.bucket_name = bucket_name
        self.endpoint = endpoint

        print(self.bucket.get_bucket_info())
        logger.info(f"OSSUtils 初始化成功 - Bucket: {bucket_name}")

    # region 文件操作

    def upload_file(self, local_path: str, oss_key: str) -> bool:
        """
        上传单个文件

        :param oss_key: oss键
        :param local_path：本地保存路径
        """
        if not os.path.exists(local_path):
            logger.error(f"本地文件不存在: {local_path}")
            return False

        oss_key = self.format_key(oss_key)

        try:
            logger.info(
                f"开始上传文件: {local_path} → oss://{self.bucket_name}/{oss_key}"
            )

            self.bucket.put_object_from_file(oss_key, local_path)

            logger.info(f"文件上传成功: {oss_key}")
            return True
        except Exception as e:
            logger.error(f"上传文件失败 {oss_key}: {e}")
            return False

    def download_file(self, oss_key: str, local_path: str) -> bool:
        """
        下载单个文件

        :param oss_key: oss键
        :param local_path：本地保存路径
        """
        import oss2

        try:
            # 确保本地目录存在
            os.makedirs(os.path.dirname(local_path) or ".", exist_ok=True)

            oss_key = self.format_key(oss_key)

            logger.info(
                f"开始下载文件: oss://{self.bucket_name}/{oss_key} → {local_path}"
            )

            self.bucket.get_object_to_file(oss_key, local_path)

            logger.info(f"文件下载成功: {local_path}")
            return True
        except oss2.exceptions.NoSuchKey:
            logger.error(f"OSS中文件不存在: {oss_key}")
            return False
        except Exception as e:
            logger.error(f"下载文件失败 {oss_key}: {e}")
            return False

    def delete_file(self, oss_key: str) -> bool:
        """
        删除单个文件

        :param oss_key: oss键
        """
        try:
            oss_key = self.format_key(oss_key)
            self.bucket.delete_object(oss_key)
            logger.info(f"文件删除成功: {oss_key}")
            return True
        except Exception as e:
            logger.error(f"删除文件失败 {oss_key}: {e}")
            return False

    # endregion

    # region 文件夹操作

    def upload_folder(self, local_folder: str, oss_prefix: str = "") -> List[str]:
        """
        上传整个文件夹（递归）
        返回成功上传的文件 oss_key 列表

        :param local_folder：本地保存路径
        :param oss_prefix: oss前缀，可以理解为文件夹路径
        """
        if not os.path.exists(local_folder):
            logger.error(f"本地文件夹不存在: {local_folder}")
            return []

        oss_prefix = self.format_prefix(oss_prefix)

        uploaded_files = []
        local_folder = os.path.abspath(local_folder)

        logger.info(
            f"开始上传文件夹: {local_folder} → oss://{self.bucket_name}/{oss_prefix}"
        )

        for root, dirs, files in os.walk(local_folder):
            for file in files:
                local_path = os.path.join(root, file)

                # 计算 OSS key（保持相对路径结构）
                rel_path = os.path.relpath(local_path, local_folder)
                oss_key = (
                    os.path.join(oss_prefix, rel_path).replace("\\", "/").lstrip("/")
                )

                if self.upload_file(local_path=local_path, oss_key=oss_key):
                    uploaded_files.append(oss_key)

        logger.info(f"文件夹上传完成，共上传 {len(uploaded_files)} 个文件")
        return uploaded_files

    def download_folder(self, oss_prefix: str, local_folder: str) -> List[str]:
        """
        下载整个文件夹（前缀匹配）
        返回成功下载的文件本地路径列表

        :param local_folder：本地保存路径
        :param oss_prefix: oss前缀，可以理解为文件夹路径
        """
        import oss2

        os.makedirs(local_folder, exist_ok=True)
        downloaded_files = []

        logger.info(
            f"开始下载文件夹: oss://{self.bucket_name}/{oss_prefix} → {local_folder}"
        )

        oss_prefix = self.format_prefix(oss_prefix)

        try:
            for obj in oss2.ObjectIterator(self.bucket, prefix=oss_prefix):
                if obj.key.endswith("/"):  # 跳过文件夹标记
                    continue

                # 计算本地保存路径
                rel_path = obj.key[len(oss_prefix) :].lstrip("/")
                local_path = os.path.join(local_folder, rel_path)

                if self.download_file(oss_key=obj.key, local_path=local_path):
                    downloaded_files.append(local_path)

            logger.info(f"文件夹下载完成，共下载 {len(downloaded_files)} 个文件")
            return downloaded_files
        except Exception as e:
            logger.error(f"下载文件夹失败: {e}")
            return downloaded_files

    def delete_folder(self, oss_prefix: str) -> int:
        """
        删除整个文件夹（前缀匹配）
        返回删除的文件数量

        :param oss_prefix: oss前缀，可以理解为文件夹路径
        """
        import oss2

        deleted_count = 0
        objects_to_delete = []

        logger.info(f"开始删除文件夹: oss://{self.bucket_name}/{oss_prefix}")
        oss_prefix = self.format_prefix(oss_prefix)
        try:
            for obj in oss2.ObjectIterator(self.bucket, prefix=oss_prefix):
                objects_to_delete.append(obj.key)

                # 批量删除（OSS 每次最多支持 1000 个）
                if len(objects_to_delete) >= 1000:
                    self.bucket.batch_delete_objects(objects_to_delete)
                    deleted_count += len(objects_to_delete)
                    objects_to_delete = []

            # 删除剩余对象
            if objects_to_delete:
                self.bucket.batch_delete_objects(objects_to_delete)
                deleted_count += len(objects_to_delete)

            logger.info(f"文件夹删除完成，共删除 {deleted_count} 个对象")
            return deleted_count
        except Exception as e:
            logger.error(f"删除文件夹失败: {e}")
            return deleted_count

    # endregion

    # region 辅助方法

    def get_oss_url(self, oss_prefix: str = "", oss_key: str = ""):
        """
        返回oss的url
        :param oss_prefix: oss前缀，可以理解为文件夹路径
        :param oss_key: oss键，可以理解为文件名称
        :return: oss的url
        """

        return f"https://{self.bucket_name}.{self.endpoint}{self.format_prefix(oss_prefix)}{self.format_key(oss_key)}"

    def format_prefix(self, oss_prefix: str) -> str:
        """
        标准化 OSS 前缀（文件夹路径）
        确保以 '/' 结尾，且去除多余的 '/' 和空格

        示例：
            "projects"      → "projects/"
            "projects/v1"   → "projects/v1/"
            "projects/v1/"  → "projects/v1/"
            ""              → ""
            "/"             → ""
            " /a/b/ "       → "a/b/"
        """
        if not oss_prefix or oss_prefix.strip() in ("", "/"):
            return ""

        # 去掉首尾空格
        prefix = oss_prefix.strip()

        # 移除开头和结尾的 '/'
        prefix = prefix.strip("/")

        if not prefix:
            return ""

        # 替换多个 '/' 为单个，并确保以 '/' 结尾
        prefix = "/".join(filter(None, prefix.split("/")))

        return prefix + "/"

    def format_key(self, oss_key: str) -> str:
        """
        标准化 OSS Key（文件完整路径）
        去除首尾空格和多余的 '/'，但不强制加末尾 '/'

        示例：
            "output.txt"           → "output.txt"
            "/output.txt"          → "output.txt"
            "projects/output.txt"  → "projects/output.txt"
            "projects//output.txt" → "projects/output.txt"
            " /a/b/c.txt "         → "a/b/c.txt"
        """
        if not oss_key or oss_key.strip() in ("", "/"):
            logger.warning("oss_key 为空，返回空字符串")
            return ""

        # 去掉首尾空格
        key = oss_key.strip()

        # 移除开头和结尾的 '/'
        key = key.strip("/")

        if not key:
            return ""

        # 替换多个连续的 '/' 为单个 '/'
        key = "/".join(filter(None, key.split("/")))

        return key

    def get_prefix_list(self, prefix: str = "", max_keys: int = 200) -> List[dict]:
        """列出指定前缀下的对象"""
        import oss2

        prefix = self.format_prefix(prefix)
        objects = []
        try:
            for obj in oss2.ObjectIterator(
                self.bucket, prefix=prefix, max_keys=max_keys
            ):
                objects.append(
                    {
                        "key": obj.key,
                        "size": obj.size,
                        "last_modified": obj.last_modified,
                        "type": "folder" if obj.key.endswith("/") else "file",
                    }
                )
            logger.info(f"列出对象完成，前缀: {prefix}, 数量: {len(objects)}")
            return objects
        except Exception as e:
            logger.error(f"列出对象失败: {e}")
            return objects

    def exists(self, oss_key: str) -> bool:
        """判断文件或文件夹是否存在"""

        import oss2

        try:
            self.bucket.get_object_meta(oss_key)
            return True
        except oss2.exceptions.NoSuchKey:
            # 检查是否为文件夹
            for _ in oss2.ObjectIterator(self.bucket, prefix=oss_key, max_keys=1):
                return True
            return False
        except Exception:
            return False

    # endregion


# ====================== 使用示例 ======================

if __name__ == "__main__":
    # 推荐使用环境变量配置密钥
    import os

    oss_utils = OSSUtils(
        access_key_id=os.getenv("ALIYUN_ACCESS_KEY_ID", "your_access_key_id"),
        access_key_secret=os.getenv(
            "ALIYUN_ACCESS_KEY_SECRET", "your_access_key_secret"
        ),
        endpoint="oss-cn-hangzhou.aliyuncs.com",  # 根据你的地域修改
        bucket_name="your-bucket-name",
    )

    # 示例使用：
    # oss_utils.upload_file("local.txt", "test/remote.txt")
    # oss_utils.upload_folder("./my_folder", "projects/v1/")
    # oss_utils.download_folder("projects/v1/", "./download/")
    # oss_utils.delete_folder("temp/")
