import os
import json
import logging

logger = logging.getLogger(__name__)


class CacheUtils:

    @staticmethod
    def get_cache_json(is_test, json_path, data_lambda: callable):
        """
        本地 JSON 缓存读取工具

        逻辑说明：
        1. 测试模式 且 缓存文件存在：
        - 直接从本地 JSON 读取
        - 如果 JSON 损坏，自动回退到重新获取数据
        2. 非测试模式 或 缓存不存在：
        - 调用 data_lambda() 获取数据
        - 将结果写入本地 JSON 文件作为缓存

        :param is_test: bool，是否为测试模式
        :param json_path: str，JSON 缓存文件路径
        :param data_lambda: callable，无参函数，用于获取原始数据
        :return: dict | list，最终数据结果
        """

        # 优先尝试从缓存读取（仅测试模式）
        if is_test and os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logger.warning(f"缓存 JSON 读取失败，准备重新生成：{json_path}")

        # 调用数据源获取数据
        data_result = data_lambda()

        # 确保父目录存在
        dir_path = os.path.dirname(json_path)
        if dir_path and not os.path.exists(dir_path):
            os.makedirs(dir_path, exist_ok=True)

        # 写入缓存（缓存失败不影响主流程）
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data_result, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logger.error(f"缓存 JSON 写入失败：{json_path}, 错误：{e}")

        return data_result

    @staticmethod
    def save_cache_file(data, file_path):
        """
        保存缓存文件

        :param data: 要缓存的数据
        :param file_path: 缓存文件路径
        """
        try:
            dir_path = os.path.dirname(file_path)
            if dir_path and not os.path.exists(dir_path):
                os.makedirs(dir_path, exist_ok=True)

            if isinstance(data, str):
                content = data
            elif isinstance(data, (dict, list)):
                content = json.dumps(data, ensure_ascii=False, indent=4)
            else:
                raise ValueError("数据格式错误")

            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)

            logger.info(f"内容已缓存：{file_path}")

        except Exception as e:
            logger.exception(f"缓存文件写入失败：{file_path}")
