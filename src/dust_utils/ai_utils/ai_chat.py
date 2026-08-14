import requests
import json
import time
import re
from enum import Enum
from urllib.parse import urlsplit, urlunsplit, quote

# 创建模块专用记录器
from loguru import logger

# 返回一个新的 Logger 实例，默认启用 ANSI
# logger = logger.opt(colors=True)


class ChatModel(Enum):
    """
    枚举，表示使用的请求模式
    """

    RESPONSES = 1
    CHAT_COMPLETIONS = 2


class AIChat:
    """
    这是一个AI聊天工具类，主要功能包括:
    1. 与AI大模型进行对话交互，支持文本和图像输入
    2. 支持图像输出
    3. 支持多种AI模型配置和服务商(OpenAI格式)
    4. 记录对话token使用量、费用和响应时间
    5. 提供JSON和代码格式修复功能

    主要方法:
    - send_message(): 发送消息并获取AI响应，支持文本和图像输入
    - clear_message(): 清空对话历史
    - fix_json(): 修复不规范的JSON字符串
    - fix_code(): 移除代码块标记，支持多种编程语言
    """

    def __init__(self, config):
        """
        初始化AI聊天实例

        Args:
            config: 包含AI配置信息的字典，需要包含hostsUrl、apiKey和model字段
        """
        try:
            from openai import OpenAI

            self.openai = OpenAI
        except ImportError:
            raise ImportError(
                "检测到未安装 openai。请执行 'pip install openai' 以使用此功能。"
            )

        self.base_url = config.get("baseUrl")
        self.api_key = config.get("apiKey")
        self.model = config.get("model")
        self.mask = config.get("mask")
        self.modelType = config.get("modelType", ["text"])
        self.temperature = config.get("temperature", 0.2)
        self.chat_model = ChatModel.RESPONSES  # 默认使用 Responses API

        # 初始化ai角色定义
        self.messageList = [
            {
                "role": "system",
                "content": self.mask,
            }
        ]
        self.imageMessageList = [
            {
                "role": "system",
                "content": "你是绘图提示词生成器，把用户要求转成英文prompt，不解释",
            }
        ]

        # 金额定价
        self.input_price = config.get("inputPrice", 0) / 1000  # 输入金额定价
        self.output_price = config.get("outputPrice", 0) / 1000  # 输出金额定价
        self.price = 0  # 已使用总金额
        self.useToken = 0  # 已使用总token
        self.useTime = 0  # 已使用总时间

        # 其他信息
        self.credits = None
        self.creditAlert = config.get("creditAlert", 0)
        self.sendCount = 0  # 发送次数

        self._init_colors()

    def _init_colors(self):
        self.input_color = "#31bdec"  # 输入消息颜色
        self.url_color = "#075F96"  # URL链接颜色
        self.log_color = "#ffb800"  # 输出消息颜色
        self.statistics_color = "#ff5722"  # 统计信息颜色

    def send_message(self, message, image_list=[]):
        """
        发送消息到AI服务并获取响应

        Args:
            message: 要发送给AI的消息内容

        Returns:
            str: AI的响应消息
        """
        try:

            client = self.openai(
                # 若没有配置环境变量,请用阿里云百炼API Key将下行替换为:api_key="sk-xxx",
                api_key=self.api_key,
                base_url=self.base_url,
            )

            print("")
            # 输入消息和图片列表
            logger.color_msg(f"{message}", color=self.input_color)
            if len(image_list) > 0:
                logger.color_msg(f"图片: {image_list}", color=self.url_color)

            # 记录开始时间
            start_time = time.time()

            # 自动切换调用，获取实际的回复内容、使用情况
            response_content, usage = self._call_api(client, message, image_list)

            # 计算响应时间
            response_time = time.time() - start_time
            self.useTime += response_time  # 累计使用时间

            # 计算本次对话的token使用量和金额
            input_token = usage["prompt_tokens"]
            output_token = usage["completion_tokens"]
            self.useToken += input_token + output_token  # 累计使用token
            this_send_price = (
                input_token * self.input_price + output_token * self.output_price
            )
            self.price += this_send_price  # 累计使用金额

            # 将大模型的回复信息添加到对话列表中
            self.messageList.append({"role": "assistant", "content": response_content})

            logger.info(response_content + "")
            # 输出黄色的token使用量和本次对话金额
            logger.color_msg(
                f"使用Token: {input_token + output_token}\t金额: {this_send_price:.6f}元\t响应时间: {response_time:.2f}秒\tAI模型: {self.model}\tbaseURL: {self.base_url}\t{str(self.chat_model.name)}",
                color=self.log_color,
            )

            self.sendCount += 1  # 发送次数加1
            return response_content

        except requests.exceptions.RequestException as e:
            logger.error(f"请求发生错误: {e}")
            return None
        except json.JSONDecodeError as e:
            logger.error(f"响应解析错误: {e}")
            return None
        except Exception as e:
            logger.error(f"发生未知错误: {e}")
            return None

    def _get_content(self, message, image_list=[]):
        content = []
        for url in image_list:
            parts = urlsplit(url)

            encoded_url = urlunsplit(
                (
                    parts.scheme,
                    parts.netloc,
                    quote(parts.path, safe="/"),
                    parts.query,
                    parts.fragment,
                )
            )

            # if "openrouter" in self.base_url.lower():
            #     content.append({"type": "image_url", "image_url": encoded_url})
            # else:

            if self.chat_model == ChatModel.RESPONSES:
                content.append({"type": "input_image", "image_url": encoded_url})
            else:
                content.append({"type": "image_url", "image_url": {"url": encoded_url}})

        if len(image_list) > 0:
            if self.chat_model == ChatModel.RESPONSES:
                content.append({"type": "input_text", "text": message})
            else:
                content.append({"type": "text", "text": message})
        else:
            content = message

        self.messageList.append({"role": "user", "content": content})

    def _call_api(self, client, message, image_list):
        """优先使用 Responses API，失败则自动降级到 Chat Completions"""
        # 1. 尝试 Responses API

        try:
            if self.chat_model == ChatModel.RESPONSES:
                # 配置消息内容
                self._get_content(message, image_list)
                resp = client.responses.create(
                    model=self.model,
                    input=self.messageList,
                    temperature=self.temperature,
                )
                text = "".join(
                    item.content[0].text
                    for item in resp.output
                    if item.type == "message" and item.content[0].type == "output_text"
                )
                usage = {
                    "prompt_tokens": resp.usage.input_tokens,
                    "completion_tokens": resp.usage.output_tokens,
                }
                return text, usage
        except Exception as e:
            logger.warning(f"Responses API 不可用 ({e})，降级到 Chat Completions")
            self.messageList.pop()  # 移除最新的内容
            self.chat_model = ChatModel.CHAT_COMPLETIONS  # 对话模式降级

        self._get_content(message, image_list)
        # 2. 降级到 Chat Completions
        resp = client.chat.completions.create(
            model=self.model,
            messages=self.messageList,
            temperature=self.temperature,
        )

        text = resp.choices[0].message.content
        usage = {
            "prompt_tokens": resp.usage.prompt_tokens,
            "completion_tokens": resp.usage.completion_tokens,
        }
        return text, usage

    def gen_image(
        self, message: str, output_path: str = "output.png", size: str = None
    ):
        try:
            import base64
            import requests

            client = self.openai(
                api_key=self.api_key,
                base_url=self.base_url,
            )

            logger.color_msg(f"[绘图] {message}", color=self.input_color)

            start_time = time.time()

            payload = {
                "model": self.model,
                "messages": [{"role": "user", "content": message}],
                "modalities": ["image"],  # 必须加上这个
            }

            # 处理 size（转为 aspect_ratio，Gemini 系列支持较好）
            extra_body = {}

            # 自动判断是否需要 16:9（优先级：参数 > prompt 中的关键词）
            aspect_ratio = None
            if size:
                aspect_ratio = self._size_to_aspect_ratio(size)
            elif any(
                x in message for x in ["16:9", "16：9", "宽屏", "横屏", "landscape"]
            ):
                aspect_ratio = "16:9"
            elif any(x in message for x in ["9:16", "竖屏", "portrait"]):
                aspect_ratio = "9:16"

            if aspect_ratio:
                extra_body["image_config"] = {"aspect_ratio": aspect_ratio}

            # 如果有 extra_body，就传进去
            if extra_body:
                payload["extra_body"] = extra_body

            response = client.chat.completions.create(**payload)

            response_time = time.time() - start_time
            self.useTime += response_time

            # 提取图片
            msg = response.choices[0].message
            image_saved = False

            if hasattr(msg, "images") and msg.images:
                for idx, item in enumerate(msg.images):
                    # item 是 dict 类型
                    if isinstance(item, dict):
                        image_data = item.get("image_url") or item.get("imageUrl")
                        if isinstance(image_data, dict):
                            url = image_data.get("url")
                        else:
                            url = None
                    else:
                        # 兼容对象类型（万一以后变了）
                        url = getattr(
                            getattr(item, "image_url", None), "url", None
                        ) or getattr(item, "url", None)

                    if url and isinstance(url, str):
                        if url.startswith("data:image"):
                            # base64 格式（Gemini 最常见）
                            try:
                                header, b64_data = url.split(",", 1)
                                with open(output_path, "wb") as f:
                                    f.write(base64.b64decode(b64_data))
                                image_saved = True
                                logger.info(
                                    f"第 {idx+1} 张图片已保存 (base64) → {output_path}"
                                )
                                break
                            except Exception as decode_err:
                                logger.error(f"base64 解码失败: {decode_err}")
                        else:
                            # 普通 URL
                            img_data = requests.get(url, timeout=60).content
                            with open(output_path, "wb") as f:
                                f.write(img_data)
                            image_saved = True
                            logger.info(
                                f"第 {idx+1} 张图片已保存 (url) → {output_path}"
                            )
                            break

            if not image_saved:
                logger.error("响应中未找到图片数据")
                logger.error(f"完整响应: {response}")
                return None

            logger.info(f"[图片] 保存成功: {output_path}")
            logger.color_msg(
                f"响应时间: {response_time:.2f}秒\t模型: {self.model}\tbaseURL: {self.base_url}",
                color=self.log_color,
            )
            return output_path

        except Exception as e:
            logger.error(f"绘图错误: {e}")
            import traceback

            logger.error(traceback.format_exc())
            return None

    def _size_to_aspect_ratio(self, size: str) -> str | None:
        """简单 size 转 aspect_ratio"""
        if not size:
            return None
        s = size.lower().replace(" ", "").replace("*", "x")
        mapping = {
            "1024x1024": "1:1",
            "1792x1024": "16:9",
            "1024x1792": "9:16",
            "1344x768": "16:9",
            "768x1344": "9:16",
        }
        return mapping.get(s)

    def clear_message(self):
        """
        清空消息列表
        """
        self.messageList = [
            {
                "role": "system",
                "content": self.mask,
            }
        ]

    def fix_json(self, json_str, out_obj=True):
        """
        修复不规范的JSON字符串，支持自动修复常见的JSON格式错误

        Args:
            json_str: 可能不规范的JSON字符串
            out_obj: 是否返回Python对象，True返回dict对象，False返回JSON字符串

        Returns:
            Union[dict, str]: 根据out_obj参数返回修复后的JSON对象或字符串
            - 当out_obj=True时返回dict对象
            - 当out_obj=False时返回格式化的JSON字符串
        """
        if not json_str:
            if out_obj:
                return {}
            else:
                return "{}"

        try_count = 0
        max_try_count = 3  # 最大重试次数

        while try_count < max_try_count:

            json_str = self.fix_code(json_str, ["json"]).replace("\n", "")

            # 移除所有 <style>...</style> 内容
            json_str = re.sub(r"<style>.*?</style>", "", json_str, flags=re.DOTALL)
            # 使用正则表达式查找缺少引号的键值对
            # 匹配模式: "key":value 其中value不是以引号、数字、{、[、true、false、null开头的
            pattern = r'("[^"]+":)\s*([^\s"\d\{\[trfn][^,\}\]]*)'  # 匹配没有引号的值
            json_str = re.sub(pattern, r'\1"\2"', json_str)

            # 修复没有使用双引号包裹的属性名
            pattern_unquoted_key = r"(\{|\,)\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*:"
            json_str = re.sub(pattern_unquoted_key, r'\1"\2":', json_str)

            try:
                jsonObj = json.loads(json_str)
                if out_obj:
                    return jsonObj
                return json.dumps(jsonObj, ensure_ascii=False)
            except json.JSONDecodeError:
                try_count += 1
                jsonErrorQuestion = f"```{json_str}```这是一个json格式错误的文本，请帮我修正，请注意属性应被双引号包裹，我只要修正后的json，不要输出其他内容，也不要增删属性，保持json数据结构不变，属性值中可能存在双引号，注意转义"
                json_str = self.send_message(jsonErrorQuestion)

        # 超过最大重试次数后抛出异常
        if try_count >= max_try_count:
            error_msg = f"JSON修复失败,已重试{max_try_count}次"
            logger.error(f"{error_msg}")  # 红色打印错误信息
            raise ValueError(error_msg)

    def fix_js(self, javascript_code):
        """
        修复JavaScript代码中的语法错误

        Args:
            javascript_code: 包含JavaScript代码的字符串

        Returns:
            str: 修复后的JavaScript代码字符串
        """
        try:
            import esprima
        except ImportError:
            raise ImportError("该功能需要 esprima，请执行：pip install esprima")

        if not javascript_code:
            return ""

        # ----------- 内部工具函数 -----------
        def strip_comments(code: str) -> str:
            """去掉 JS 单行和多行注释"""
            code = re.sub(r"/\*[\s\S]*?\*/", "", code)  # 多行注释
            code = re.sub(r"//[^\n]*", "", code)  # 单行注释
            return code

        def sanitize_js(code: str) -> str:
            """修复字符串字面量中被意外打断的换行，替换成 '\\n'"""
            code = re.sub(r"'[\r\n]+'", r"'\\n'", code)  # 单引号里的非法换行
            code = re.sub(r'"[\r\n]+"', r'"\\n"', code)  # 双引号里的非法换行
            return code

        def js_syntax_ok(code: str) -> bool:
            """仅做语法检查，返回 True/False"""
            try:
                esprima.parseScript(code, tolerant=False)
                return True
            except esprima.Error:
                return False

        # ----------- 内部工具函数结束 -----------

        try_count = 0
        max_try_count = 3

        while try_count < max_try_count:
            javascript_code = self.fix_code(javascript_code)  # 移除代码块标记
            no_comment_code = strip_comments(javascript_code)  # 1. 去注释
            sanitized_code = sanitize_js(no_comment_code)  # 2. 修非法换行

            if js_syntax_ok(sanitized_code):  # 3. 语法校验
                return javascript_code

            # 语法仍报错 → 交给 AI 修复
            try_count += 1
            js_error_question = (
                f"```{javascript_code}```\n"
                f"这是一个 JavaScript 代码，其中可能存在语法错误，请帮我修正。"
                f"我只要修正后的代码，不要输出其他内容，也不要改变代码逻辑或者修改变量、属性名称以及对应值。"
            )
            javascript_code = self.send_message(js_error_question)

        # 超过最大重试次数
        error_msg = f"JavaScript 代码修复失败，已重试 {max_try_count} 次"
        logger.error(error_msg)
        raise ValueError(error_msg)

    def fix_mermaid(self, mermaid_code):
        """
        修复Mermaid图表代码中的语法错误

        Args:
            mermaid_code: 包含Mermaid图表代码的字符串

        Returns:
            str: 修复后的Mermaid图表代码字符串
        """
        try:
            import mermaid as md
        except ImportError:
            raise ImportError(
                "生成 Mermaid 图表需要 mermaid，请执行：pip install mermaid-py"
            )

        if not mermaid_code:
            return ""

        try_count = 0
        max_try_count = 3  # 最大重试次数

        # while try_count < max_try_count:
        #     # 移除代码块标记
        #     mermaid_code = self.fix_code(mermaid_code, ["mermaid"])

        #     try:
        #         code = mermaid_code.replace("\\n", "\n")
        #         # 使用pymermaid检查Mermaid语法
        #         mermaid = md.Mermaid(code)
        #         if mermaid.svg_response.status_code != 200:
        #             raise ValueError(f"mermaid字符串异常:{mermaid_code}")
        #         return mermaid_code
        #     except Exception as e:
        #         try_count += 1
        #         # 发送修复请求给AI
        #         mermaid_error_question = f"```{mermaid_code}```这是一个Mermaid图表代码，其中可能存在语法错误，请帮我修正，我只要修正后的代码，不要输出其他内容，也不要改变图表逻辑或者修改节点、关系以及对应的描述"
        #         mermaid_code = self.send_message(mermaid_error_question)

        # # 超过最大重试次数后抛出异常
        # if try_count >= max_try_count:
        #     error_msg = f"Mermaid图表代码修复失败,已重试{max_try_count}次"
        #     logger.error(f"{error_msg}")
        #     raise ValueError(error_msg)

    def fix_code(self, code, additional_tags=[]):
        """
        移除代码字符串中的代码块标记（如```python等）

        Args:
            code: 需要处理的代码字符串，可能包含代码块标记
            additional_tags: 额外的编程语言标签列表，用于扩展默认支持的语言类型

        Returns:
            str: 移除代码块标记后的代码字符串，保持代码内容不变
        """
        # 定义常见编程语言列表
        languages = [
            # 后端语言
            "python",
            "java",
            "c",
            "c++",
            "c#",
            "csharp",
            "go",
            "rust",
            "php",
            "ruby",
            "kotlin",
            "scala",
            "perl",
            "r",
            # 前端语言
            "javascript",
            "typescript",
            "html",
            "css",
            "sass",
            "less",
            "vue",
            "react",
            "angular",
            # 数据库
            "sql",
            "mysql",
            "postgresql",
            "mongodb",
            # 标记语言
            "xml",
            "yaml",
            "json",
            "markdown",
            # 脚本语言
            "shell",
            "bash",
            "powershell",
            "batch",
            # 移动开发
            "swift",
            "objective-c",
            "dart",
            "flutter",
            # 其他语言
            "matlab",
            "assembly",
            "fortran",
            "cobol",
            "pascal",
            "ada",
            "lisp",
            "prolog",
            "haskell",
            "erlang",
            "elixir",
            "lua",
        ]

        if additional_tags:
            languages.extend(additional_tags)

        if not code:
            return ""

        # 使用正则表达式移除所有语言的代码块标记
        for lang in languages:
            # 使用re.escape转义语言名，避免元字符引发正则错误
            pattern = re.compile(rf"```{re.escape(lang)}[\s\n]", re.IGNORECASE)
            code = pattern.sub("", code)

        # 移除剩余的代码块标记和换行符
        code = code.replace("```", "")

        return code

    def print_statistics(self, prefix="", suffix=""):
        """
        输出当前对话的使用情况统计，包括使用的Token数量、金额、响应时间、AI模型和baseURL等信息
        """
        if prefix:
            prefix = prefix + "\t"
        if suffix:
            suffix = "\t" + suffix

        logger.info(
            f"<fg {self.statistics_color}>{prefix}总Token：{self.useToken}\t总金额: {(self.price):.6f}元\t总响应时间：{self.useTime:.2f}秒\tAI模型：{self.model}\t总次数: {(self.sendCount)}{suffix}</>"
        )
