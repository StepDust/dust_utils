# 导入必要的模块
import random  # 用于生成随机数
import requests  # 用于发送HTTP请求
import os  # 用于文件和目录操作
import re
import json
from .api_utils import ApiUtils

# 创建模块专用记录器
from loguru import logger


class RHLTAPI:
    def __init__(self, token_str: str):
        """
        初始化RHLTAPI类，设置API访问所需的基本配置。

        :param token_str: 用于API请求的授权令牌字符串，用于身份验证。
        :raises ValueError: 如果token_str为空，则抛出此异常。
        """
        # 禁用所有代理设置，确保直接连接
        os.environ.pop("HTTP_PROXY", None)
        os.environ.pop("HTTPS_PROXY", None)
        os.environ.pop("http_proxy", None)
        os.environ.pop("https_proxy", None)
        os.environ["NO_PROXY"] = "*"
        os.environ.pop("ALL_PROXY", None)

        if not token_str:
            raise ValueError("请求api时，token不能为空")
        self.token_str = token_str  # 存储token
        self.headers = {"Authorization": token_str}  # 设置请求头
        self.base_url = "https://renhelitai.com/prod-api"  # API基础URL
        # 软著名称后缀映射关系，用于标识不同类型的软件
        self.sys_soft_file_type = {
            "1": "系统",
            "2": "平台",
            "3": "软件",
            "4": "小程序",
            "5": "APP",
        }
        self.download_base_urls = [
            "https://rhlt.oss-cn-beijing.aliyuncs.com",
            "https://renhelitai.com/prod-api/profile/upload",
        ]

    def get_ccDetailList(self, params: dict = {}) -> dict:
        """
        获取软著详情列表，用于流程状态管理。

        :param params: 查询参数字典，支持分页和过滤条件
        :return: API响应的JSON数据，包含软著详情列表
        """
        # 构建API请求URL
        url = f"{self.base_url}/system/distributer/ccDetailList"

        # 设置默认的分页参数
        default_params = {
            "pageNum": 1,  # 默认第一页
            "pageSize": 50,  # 每页50条记录
            "todoFlag": "N",  # 默认非待办
        }
        # 合并用户参数和默认参数
        params = ApiUtils.combined_params(params=params, default_params=default_params)

        # 发送GET请求并返回JSON响应
        response = requests.get(url, params=params, headers=self.headers)
        return response.json()

    def post_importSignatureFile(self, busDistributerId: str, filePath: str) -> dict:
        """
        上传已盖章签章页文件。

        :param busDistributerId: 业务分发ID，用于标识具体业务
        :param filePath: 签章页文件的本地路径
        :return: 上传结果的JSON响应数据
        """
        # 构建上传URL
        url = f"{self.base_url}/system/distributer/importSignatureFile"
        # 设置请求参数
        params = {
            "busDistributerId": busDistributerId,
            "fileType": "signaturedPath",  # 指定文件类型为签章页
        }
        # 验证文件是否存在
        if not os.path.exists(filePath):
            logger.error(f"错误：文件 {filePath} 不存在")
            return {}

        # 打开文件并发送POST请求
        files = {"file": open(filePath, "rb")}
        response = requests.post(url, params=params, files=files, headers=self.headers)
        return response.json()

    def post_handleAproval(self, busDistributerId: str, caseStatus: int) -> dict:
        """
        处理签章页的审批操作。

        :param busDistributerId: 业务分发ID，标识待审批的业务
        :param caseStatus: 审批状态码，表示审批结果
        :return: 审批操作的响应结果
        """
        # 构建审批URL
        url = f"{self.base_url}/system/distributer/handleAproval"
        # 设置审批参数
        params = {
            "id": busDistributerId,
            "caseStatus": caseStatus,
        }
        # 发送审批请求
        response = requests.post(url, json=params, headers=self.headers)
        return response.json()

    def get_work_make_list(self, params: dict = {}) -> dict:
        """
        获取制件任务列表。

        :param params: 查询参数，支持自定义过滤条件
        :return: 包含制件列表的JSON响应
        """
        # 构建列表查询URL
        url = f"{self.base_url}/work/make/list"

        # 设置默认分页参数
        default_params = {
            "pageNum": 1,
            "pageSize": 50,
        }
        # 合并查询参数
        params = ApiUtils.combined_params(params, default_params)

        # 发送查询请求
        response = requests.get(url, params=params, headers=self.headers)
        return response.json()

    def download_case(self, case_id, output_folder):
        """
        下载软著材料并保存到指定目录。

        :param case_id: 案件ID，用于标识具体案件
        :param output_folder: 下载文件的保存目录
        :return: 保存的文件路径，下载失败返回空字符串
        """
        # 获取案件详情
        url = f"{self.base_url}/work/make/getFinalData/{case_id}"
        response = requests.get(url, headers=self.headers)

        # 解析响应获取下载路径
        case_info = response.json()
        finalZipPath = f"{case_info['data']['finalZipPath']}?timestamp={random.randint(0, 1000000000)}"

        save_path = os.path.join(
            output_folder, ApiUtils.get_filename_by_url(finalZipPath)
        )
        save_path = ApiUtils.download_file(
            self.download_base_urls, finalZipPath, save_path, desc="软著材料"
        )

        return save_path

    def download_code(self, case_id, case_name, output_folder):
        """
        下载软著源代码并保存到本地。

        :param case_id: 案件ID
        :param case_name: 案件名称，用于生成文件名
        :param output_folder: 下载文件的保存目录
        :return: 保存的文件路径，下载失败返回空字符串
        """
        # 获取案件详情
        url = f"{self.base_url}/work/make/getCodeDataInfo/{case_id}"
        response = requests.get(url, headers=self.headers)

        # 解析响应获取下载路径
        case_info = response.json()
        regex = r"^(/)*profile/upload"
        codePath = (
            f"{case_info['data']['codePath']}?timestamp={random.randint(0, 1000000000)}"
        )
        # 处理开头文件夹，这里已经写入self.download_base_urls，所以这里直接去掉
        if re.match(regex, codePath):
            codePath = re.sub(regex, "", codePath, count=1)

        save_path = os.path.join(output_folder, f"{case_name}_code.zip")
        save_path = ApiUtils.download_file(
            self.download_base_urls, codePath, save_path, desc="源码"
        )

        return save_path

    def post_work_make_importData(self, params, filePath):
        """
        上传软著材料文件。

        :param params: 上传参数，包含业务相关信息
        :param filePath: 待上传文件的本地路径
        :return: 上传响应结果
        """
        # 构建上传URL
        url = f"{self.base_url}/work/make/importData"
        default_params = {}
        # 合并上传参数
        params = ApiUtils.combined_params(params, default_params)
        # 检查文件是否存在
        if not os.path.exists(filePath):
            logger.error(f"错误：文件 {filePath} 不存在")
            return

        # 打开文件并发送上传请求
        files = {"file": open(filePath, "rb")}
        response = requests.post(
            url, params=params, files=files, headers=self.headers, timeout=600
        )
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 504:
            logger.warning(f"上传文件 {filePath} 超时，正常现象")
            return {"code": 504, "msg": "上传超时"}

    # region 使用自定义token调用，适用于外部接口

    def get_case_info_by_custom(self, case_id, case_name, remark):
        """
        返回软著信息
        :param case_id: 软著信息管理页面--件编号
        :param case_name: 软著名称
        :param remark: 属地信息
        """
        # 构建API请求URL
        url = f"{self.base_url}/outService/queryCaseInfo"

        # 设置默认的分页参数
        params = {
            "customToken": self.token_str,  # 使用实例的token进行查询
            "caseId": case_id,  # 默认第一页
            "caseName": case_name,  # 每页50条记录
            "remark": remark,  # 默认非待办
        }

        # 发送GET请求并返回JSON响应
        response = requests.get(url, params=params, headers=self.headers)
        return response.json()

    def get_oss_prefix_list(self, remote_path, remark):
        """
        获取远程路径下的文件前缀列表
        :param remote_path: 远程路径
        :param remark: 属地信息
        :return: 文件前缀列表
        """
        url = f"{self.base_url}/outService/listFiles"
        params = {
            "customToken": self.token_str,
            "remotePath": remote_path,
            "remark": remark,
        }
        response = requests.get(url, params=params, headers=self.headers)
        return response.json()

    def post_oss_upload_file(self, remote_path, file_path, remark):
        """
        上传文件到OSS
        :param remote_path: 远程路径
        :param file_path: 本地文件路径
        :param remark: 属地信息
        :return: 上传响应结果
        """

        # 构建上传URL
        url = f"{self.base_url}/outService/uploadSingleFile"
        # 设置请求参数
        params = {
            "customToken": self.token_str,
            "remotePath": remote_path,
            "remark": remark,
        }
        # 验证文件是否存在
        if not os.path.exists(file_path):
            logger.error(f"错误：文件 {file_path} 不存在")
            return {}

        # 打开文件并发送POST请求
        files = {"file": open(file_path, "rb")}
        response = requests.post(url, params=params, files=files, headers=self.headers)
        return response.json()

    def post_import_data(self, case_id, file_type, file_path, remark):
        # 构建上传URL
        url = f"{self.base_url}/outService/importData"
        # 设置请求参数
        params = {
            "customToken": self.token_str,
            "caseId": case_id,
            "fileType": file_type,
            "remark": remark,
        }
        # 验证文件是否存在
        if not os.path.exists(file_path):
            logger.error(f"错误：文件 {file_path} 不存在")
            return {}

        # 打开文件并发送POST请求
        files = {"file": open(file_path, "rb")}
        response = requests.post(url, params=params, files=files, headers=self.headers)
        return response.json()

    def download_case_custom(self, case_id, remark, output_folder):
        # 获取案件详情
        url = f"{self.base_url}/outService/getFinalData"
        params = {
            "customToken": self.token_str,
            "caseId": case_id,
            "remark": remark,
        }

        response = requests.get(url, params=params, headers=self.headers)

        # 解析响应获取下载路径
        case_info = response.json()
        finalZipPath = f"{case_info['data']['finalZipPath']}?timestamp={random.randint(0, 1000000000)}"

        save_path = os.path.join(
            output_folder, ApiUtils.get_filename_by_url(finalZipPath)
        )
        save_path = ApiUtils.download_file(
            self.download_base_urls, finalZipPath, save_path, desc="软著材料"
        )

        return save_path

    # endregion
