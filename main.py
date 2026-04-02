from dust_utils.file_utils.md_to_docx import MdToDocx
from dust_utils.ai_utils import AIChat
from dust_utils import setup_logger
import os
import json
from dotenv import load_dotenv

logger = setup_logger()

load_dotenv(r"E:\Share\配置文件\.env")


def test_md_to_word():
    md_path = r"test_config\md_to_config\test.md"
    output_path = r"C:\Users\Administered\Desktop\test.docx"
    with open(r"test_config\md_to_config\style.json", "r", encoding="utf-8") as f:
        style = json.load(f)

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    md_to_docx = MdToDocx()
    md_to_docx.convert(md_text, output_path, style["default"])


def test_ai_chat():

    ai_path = r"E:\软著做件成品\rhlt_project\01.auto_soft_make\config\ai_list.json"
    with open(ai_path, "r", encoding="utf-8") as f:
        ai_config = json.load(f)

    ai_config = ai_config[-1]
    # 适配数据库字段和代码字段不一致的问题
    ai_config["baseUrl"] = ai_config["base_url"]
    ai_config["apiKey"] = ai_config["api_key"]
    ai_config["inputPrice"] = ai_config["input_price"]
    ai_config["outputPrice"] = ai_config["output_price"]
    ai_config["creditAlert"] = ai_config["credit_alert"]

    ai_chat = AIChat(ai_config)
    output_path = r"C:\Users\Administered\Desktop\output.png"
    question = "帮我生成一张水墨山水画，16：9的格式"
    response = ai_chat.gen_image(question, output_path=output_path, size="1792x1024")

    logger.info(response)


if __name__ == "__main__":
    picui_key = os.getenv("PICUI_KEY")
    logger.info(picui_key)
    # test_md_to_word()
    # test_ai_chat()
