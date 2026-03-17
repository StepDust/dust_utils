import requests
import hashlib
import random


class BaiduTranslateAPI:
    def __init__(self, app_id, app_key):
        """
        app_id: 你的百度翻译应用ID
        app_key: 你的密钥（secretKey）
        """
        self.app_id = app_id
        self.app_key = app_key
        self.base_url = "https://fanyi-api.baidu.com"

    def _generate_sign(self, q: str, salt: str) -> str:
        """
        生成百度翻译 sign 签名
        拼接顺序：appid + q + salt + secretKey
        然后做 md5，得到32位小写字符串
        """
        # 重要：拼接 sign 时，q 不要做 urlencode
        raw = f"{self.app_id}{q}{salt}{self.app_key}"
        # 使用 utf-8 编码后计算 md5
        md5_obj = hashlib.md5(raw.encode("utf-8"))
        return md5_obj.hexdigest()

    def ai_translate(self, text: str, to_lang: str, from_lang: str = "auto"):
        """
        调用百度通用翻译 API

        参数:
            text: 要翻译的文本
            to_lang: 目标语言 (zh, en, jp, kor 等)
            from_lang: 源语言，默认为 auto 自动检测

        返回:
            百度API的原始 json 响应
        """
        url = f"{self.base_url}/api/trans/vip/translate"

        # 生成随机 salt（推荐使用随机数，长度不限，但一般用 5-10 位）
        salt = str(random.randint(32768, 65536))

        # 生成正确的 sign
        sign = self._generate_sign(text, salt)

        # 注意：发送请求时，q 需要做 urlencode
        params = {
            "q": text,  # requests 会自动帮我们 urlencode
            "from": from_lang,
            "to": to_lang,
            "appid": self.app_id,
            "salt": salt,
            "sign": sign,
        }

        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()  # 非 200 抛出异常
            return response.json()
        except requests.exceptions.RequestException as e:
            return {"error": f"请求失败: {str(e)}"}


# 使用示例
if __name__ == "__main__":
    # 替换成你真实的 appid 和密钥
    APP_ID = ""
    APP_KEY = ""

    translator = BaiduTranslateAPI(APP_ID, APP_KEY)

    result = translator.ai_translate(text="今天天气很好，我们去散步吧", to_lang="en")

    print(result)
