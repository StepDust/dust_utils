from cryptography.fernet import Fernet
import base64
import hashlib


class SymCryptoUtils:
    def __init__(self, key: str = ""):
        """
        key: base64格式密钥字符串
        如果不传，会自动生成一个
        """
        if key:
            self.key = self._derive_key(key)
        else:
            self.key = Fernet.generate_key()

        self.cipher = Fernet(self.key)

    def _derive_key(self, password: str) -> bytes:
        """
        把任意字符串转换为合法 Fernet key
        """
        # SHA256 -> 32字节
        digest = hashlib.sha256(password.encode()).digest()

        # base64 urlsafe 编码
        return base64.urlsafe_b64encode(digest)

    def get_key(self) -> str:
        """
        获取当前密钥（用于保存）
        """
        return self.key.decode()

    def encrypt(self, text: str) -> str:
        """
        加密字符串 -> 返回密文字符串
        """
        encrypted = self.cipher.encrypt(text.encode("utf-8"))
        return encrypted.decode("utf-8")

    def decrypt(self, token: str) -> str:
        """
        解密字符串 -> 返回原文
        """
        decrypted = self.cipher.decrypt(token.encode("utf-8"))
        return decrypted.decode("utf-8")


if __name__ == "__main__":
    crypto = SymCryptoUtils("renhelitai")

    key = crypto.get_key()
    print("密钥：", key)

    text = "sk-9iKaz5hve9Qvc7ZNM4ReOLvgla0BdTax5Gum83IS2KnvVEkI"

    enc = crypto.encrypt(text)
    print("加密后：", enc)

    dec = crypto.decrypt(enc)
    print("解密后：", dec)
