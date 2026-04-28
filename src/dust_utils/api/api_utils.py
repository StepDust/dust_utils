import base64
import os  # 用于文件和目录操作
import requests  # 用于发送HTTP请求
from urllib.parse import urlparse  # 用于URL解析
from urllib.parse import quote, urlunparse
import logging  # 用于日志记录

# 创建模块专用记录器
logger = logging.getLogger(__name__)


class ApiUtils:

    @staticmethod
    def get_img(img_file):
        """将本地图片转成base64编码的字符串，或者直接返回图片链接"""
        # 简单判断是否为图片链接
        if img_file.startswith("http"):
            return img_file
        else:
            with open(os.path.expanduser(img_file), "rb") as f:  # 以二进制读取本地图片
                data = f.read()
        try:
            encodestr = str(base64.b64encode(data), "utf-8")
        except TypeError:
            encodestr = base64.b64encode(data)

        return encodestr

    @staticmethod
    def combined_params(params, default_params):
        """
        合并默认参数和用户提供的参数。

        :param params: 用户提供的参数字典
        :param default_params: 默认参数字典
        :return: 合并后的参数字典
        """
        # 初始化参数字典
        if params is None:
            params = {}
        if default_params is None:
            default_params = {}

        # 遍历默认参数，填充缺失值
        for key, default_value in default_params.items():
            if key not in params or params[key] is None:
                params[key] = default_value

        return params

    @staticmethod
    def get_filename_by_url(url, has_suffix=True):
        """
        从URL中提取文件名。

        :param url: 完整的URL地址
        :param has_suffix: 是否保留文件后缀
        :return: 提取的文件名
        """
        # 解析URL获取路径
        parsed_url = urlparse(url)
        path = parsed_url.path
        # 获取文件名
        filename = os.path.basename(path)

        # 根据需要处理文件后缀
        if not has_suffix:
            # 分割文件名和后缀
            name, ext = os.path.splitext(filename)
            filename = name  # 返回不含后缀的文件名

        return filename

    @staticmethod
    def url_concat(base, path):
        """
        拼接URL路径，自动处理斜杠问题。
        不使用urljoin，因为urljoin会处理相对路径和绝对路径，
        而我们这里只需要简单的拼接，避免复杂的逻辑。

        :param base: 基础URL
        :param path: 路径
        :return: 拼接后的URL
        """
        if base.endswith("/"):
            base = base[:-1]
        if path.startswith("/"):
            path = path[1:]
        return f"{base}/{path}"

    @staticmethod
    def url_encode(url):
        parsed = urlparse(url)
        safe_path = quote(parsed.path)  # 只编码 path
        return urlunparse(
            (
                parsed.scheme,
                parsed.netloc,
                safe_path,
                parsed.params,
                parsed.query,  # 保留原始 query
                parsed.fragment,
            )
        )

    @staticmethod
    def download_file(
        download_base_urls, file_url, file_path, check_exists=True, desc="文件"
    ):
        """
        从多个服务器的指定URL下载文件并保存到本地。

        :param file_url: 相对文件路径（不含服务器地址）
        :param file_path: 本地保存文件的路径
        :param desc: 文件描述
        :return: 下载成功返回文件路径，失败返回空字符串
        """
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        proxies = {"http": None, "https": None}
        timeout = 10

        # 统一清理相对路径
        relative_url: str = file_url.strip("/")

        for i, server_url in enumerate(download_base_urls):
            is_last_attempt = i == len(download_base_urls) - 1
            if relative_url.startswith(server_url):
                full_url = relative_url
            else:
                full_url = ApiUtils.url_concat(server_url, relative_url)

            encoded_url = full_url

            logger.info(f"尝试访问服务器: {encoded_url}")

            try:
                # 改用 GET 方式探测，避免 HEAD 被误判为 200

                if check_exists:
                    try:
                        result, status_code = ApiUtils.check_file_exists(
                            full_url=encoded_url
                        )
                        if not result:
                            logger.error(
                                f"文件不存在，HTTP状态码: {status_code} → {encoded_url}"
                            )
                            if not is_last_attempt:
                                continue
                        else:
                            logger.success(f"文件存在于: {encoded_url}")
                    except requests.RequestException as e:
                        logger.error(f"探测文件失败: {encoded_url}, 错误: {e}")
                        if not is_last_attempt:
                            continue
                        else:
                            return ""

                # 确保目录存在
                dir_path = os.path.dirname(file_path)
                if not dir_path:
                    raise ValueError("filePath 无效，未指定目录路径")
                os.makedirs(dir_path, exist_ok=True)

                # 下载文件
                with requests.get(
                    encoded_url,
                    stream=True,
                    headers=headers,
                    timeout=30,
                    proxies=proxies,
                ) as r:
                    if r.status_code == 200:
                        logger.info(f"开始下载 {desc}: {encoded_url}")
                        with open(file_path, "wb") as f:
                            for chunk in r.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)
                        logger.success(f"{desc} 下载成功 → {file_path}")
                        return file_path
                    else:
                        logger.error(f"{desc} 下载失败，HTTP状态码: {r.status_code}")
                        if is_last_attempt:
                            return ""

                # 确保目录存在
                dir_path = os.path.dirname(file_path)
                if not dir_path:
                    raise ValueError("filePath 无效，未指定目录路径")
                os.makedirs(dir_path, exist_ok=True)

                # 下载文件
                with requests.get(
                    encoded_url,
                    stream=True,
                    headers=headers,
                    timeout=30,
                    proxies=proxies,
                ) as r:
                    if r.status_code == 200:
                        logger.info(f"开始下载 {desc}: {encoded_url}")
                        with open(file_path, "wb") as f:
                            for chunk in r.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)
                        logger.success(f"{desc} 下载成功 → {file_path}")
                        return file_path
                    else:
                        logger.error(f"{desc} 下载失败，HTTP状态码: {r.status_code}")
                        if is_last_attempt:
                            return ""
            except requests.RequestException as e:
                logger.error(f"访问服务器失败: {encoded_url}, 错误: {e}")
                if is_last_attempt:
                    return ""

        return ""

    @staticmethod
    def check_file_exists(full_url, debug=False):
        """
        判断远程 URL 是否存在文件
        自动兼容 HEAD/GET，支持 PDF/ZIP/图片/Office 文件判断
        """
        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "Accept": "*/*",
            }
            encoded_url = full_url  # quote(full_url, safe=":/")

            # 1️⃣ 先尝试 HEAD 请求快速判断
            try:
                head_resp = requests.head(
                    encoded_url,
                    headers=headers,
                    timeout=10,
                    allow_redirects=True,
                    proxies={"http": None, "https": None},
                )
                status = head_resp.status_code
                content_type = head_resp.headers.get("Content-Type", "").lower()
                content_length = head_resp.headers.get("Content-Length")
                content_length = (
                    int(content_length)
                    if content_length and content_length.isdigit()
                    else 0
                )
                ext = os.path.splitext(head_resp.url.split("?")[0])[1].lower()
                text_like = any(
                    t in content_type for t in ["html", "json", "text", "xml"]
                )

                # 简单判断 HEAD 可以直接确认
                if status in (200, 206):
                    if content_length > 0 and (
                        not text_like
                        or ext
                        in [
                            ".pdf",
                            ".zip",
                            ".rar",
                            ".jpg",
                            ".jpeg",
                            ".png",
                            ".gif",
                            ".mp4",
                            ".json",
                            ".exe",
                            ".md",
                        ]
                    ):
                        if debug:
                            logger.info(f"HEAD判断文件存在: {full_url}")
                        return True, status
            except requests.RequestException:
                # HEAD 失败则继续 GET
                pass

            # 2️⃣ 使用 GET 请求读取前几个字节判断文件类型
            get_resp = requests.get(
                encoded_url,
                headers=headers,
                timeout=10,
                allow_redirects=True,
                stream=True,
                proxies={"http": None, "https": None},
            )

            if get_resp.status_code not in (200, 206):
                return False, get_resp.status_code

            # 尝试读取前 64 字节判断文件类型
            preview = get_resp.raw.read(64, decode_content=True)

            # 文件头判断
            file_signatures = {
                b"%PDF": "PDF",
                b"\xff\xd8\xff": "JPEG",
                b"\x89PNG": "PNG",
                b"GIF8": "GIF",
                b"PK": "ZIP/Office",
            }
            for sig in file_signatures:
                if preview.startswith(sig):
                    if debug:
                        logger.info(f"GET文件头判断: {file_signatures[sig]}")
                    return True, get_resp.status_code

            not_file_signatures = [b"{"]

            for sig in not_file_signatures:
                if preview.startswith(sig):
                    if debug:
                        logger.info(f"GET文件头判断: {file_signatures[sig]}")
                    return False, get_resp.status_code

            # 如果有内容但不匹配文件头，也认为可能是文件
            if len(preview) > 0:
                if debug:
                    logger.info("GET判断有内容，可能是其他类型文件")
                return True, get_resp.status_code

            return False, get_resp.status_code

        except Exception as e:
            logger.error(f"判断文件存在性时出错: {full_url}, 错误: {e}")
            return False, get_resp.status_code
