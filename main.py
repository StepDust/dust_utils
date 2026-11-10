from dust_utils import setup_loguru
import random
from dust_utils.file_utils import WordUtils

logger = setup_loguru()

from dust_utils.file_utils.md_to_docx import MdToDocx
from dust_utils.ai_utils import AIChat
import os
import re
import json
from urllib.parse import unquote
from dotenv import load_dotenv

load_dotenv(r"E:\Share\配置文件\.env")


def test_md_to_word():
    md_path = (
        r"test_config\md_to_config\仓储库存数据监控与异动预警平台 使用说明书V1.0.md"
    )
    fm_folder = r"C:\Users\Administered\Desktop\封面"

    fm_list = [
        os.path.join(fm_folder, f)
        for f in os.listdir(fm_folder)
        if os.path.isfile(os.path.join(fm_folder, f)) and f.lower().endswith(".docx")
    ]

    with open(r"test_config\md_to_config\style.json", "r", encoding="utf-8") as f:
        styles = json.load(f)

    color_list = styles.get("color_list", {})

    cover_desc = styles.get("cover_desc", [])
    font_name_list = styles.get("font_name_list", [])
    case_name_colors = styles.get("case_name_colors", [])
    cover_desc_colors = styles.get("cover_desc_colors", [])
    page_number_formats = styles.get("page_number_formats", [])

    with open(md_path, "r", encoding="utf-8") as f:
        md_content = f.read().split("\n")
        md_text = "\n".join(md_content)

        md_to_docx = MdToDocx()

        for i, fm_path in enumerate(fm_list):

            logger.divider(f" {i + 1} / {len(fm_list)} ")
            name = "default"
            # 跳过颜色配置
            if name in ["color_list"]:
                continue

            color_name = random.choice(list(color_list.keys()))
            color = color_list.get(color_name)

            output_path = r"C:\Users\Administered\Desktop"
            # output_path = (
            #     rf"{output_path}\【{styles[name]['name']} - {color_name}】text.docx"
            # )
            output_path = os.path.join(output_path, os.path.basename(fm_path))

            if os.path.exists(output_path):
                # os.remove(output_path)
                continue

            style = styles[name]
            style["h1"]["font_color"] = color[0]
            style["h2"]["font_color"] = color[1]
            style["h3"]["font_color"] = color[2]
            style["text"]["font_color"] = color[3]
            style["li"]["font_color"] = color[3]

            logger.info(f"输出样式：[{color_name}] {style['name']}...")
            md_to_docx.convert(md_text, output_path, style)

            # 合并文档
            output_path = WordUtils.merge_docx(
                [fm_list[i % len(fm_list)], output_path], output_path
            )

            # 替换文档变量
            WordUtils.replace_vars(
                str(output_path.resolve()),
                {
                    "case_name": {
                        "value": "碳计量数据存储与安全溯源管理平台",
                        "style": {
                            "font_color": random.choice(case_name_colors),
                            "font_name": random.choice(font_name_list),
                        },
                    },
                    "version": "V1.2",
                    "cover_desc": {
                        "value": random.choice(cover_desc),
                        "style": {
                            "font_color": random.choice(cover_desc_colors),
                            "font_name": random.choice(font_name_list),
                        },
                    },
                },
            )

            # 页眉对齐
            header_align = random.choice(["left", "center"])

            # 页码位置
            page_number_position = random.choice(["header", "footer"])

            # 根据页码位置决定页码对齐
            if page_number_position == "header":
                # 页码在页眉时，只允许右对齐
                page_number_align = "right"
            else:
                # 页码在页脚时，可以右对齐或居中
                page_number_align = random.choice(
                    [
                        "right",
                        "center",
                    ]
                )

            logger.info(
                f"页眉对齐方式：{header_align}，页码位置：{page_number_position}，页码对齐方式：{page_number_align}"
            )
            # =========================================================
            # 增加页眉
            # =========================================================

            WordUtils.add_header(
                docx_path=output_path,
                align=header_align,
                text="碳计量数据存储与安全溯源管理平台 V1.2",
                style=style["text"],
            )

            # =========================================================
            # 增加页码
            # =========================================================
            # 随机前后缀内容
            number_format = random.choice(page_number_formats)
            WordUtils.add_page_number(
                docx_path=output_path,
                position=page_number_position,
                align=page_number_align,
                style=style["text"],
                prefix=number_format[0],
                suffix=number_format[1],
            )


def test_md_render():
    """验证 Mermaid 图表与 LaTeX 公式插入 Word"""
    md_path = r"test_config\md_to_config\test_render.md"
    style_path = r"test_config\md_to_config\style.json"
    output_path = r"test_config\md_to_config\test_render.docx"

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    with open(style_path, "r", encoding="utf-8") as f:
        styles = json.load(f)["docx_theme"]["default"]

    # convert 遇到已存在的文件会追加内容，演示时先清掉旧文件
    if os.path.exists(output_path):
        os.remove(output_path)

    md_to_docx = MdToDocx()
    md_to_docx.convert(
        md_text,
        output_path,
        styles,
        render={
            # 默认全离线：mmdc/browser 渲染图表，mathjax/matplotlib 渲染公式
            # 首次使用前先执行一次：python -m dust_utils.file_utils.diagram_render
            "mermaid_engine": "auto",
            "latex_engine": "auto",
            "allow_online": False,
        },
    )
    logger.success(f"已输出：{output_path}")


def test_txt_to_image():
    logger.info("<fg #ff0000>这是红色</>")

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


@logger.catch
def test_loguru():

    print(id(logger))
    logger.color_msg("111</>", color="#f00")  # 期望：111</>
    logger.color_msg("<red>假标签</red>", color="#f00")  # 期望：<red>假标签</red>
    logger.color_msg("a < b > c", color="#f00")  # 期望：a < b > c

    return
    logger.debug("这是一个调试日志")
    logger.info("这是一个信息日志")
    logger.warning("这是一个警告日志")
    logger.error("这是一个错误日志")
    logger.get_log_path()
    logger.success("这是一个成功日志")
    logger.critical("这是一个严重错误日志")
    logger.divider("这是一个分割线日志")
    logger.info("<fg #ff9000>main中的红色文字</>")
    logger.info("<red>main中的红色文字</red>")
    AIChat(None)


if __name__ == "__main__":
    # picui_key = os.getenv("PICUI_KEY")
    # logger.info(picui_key)
    test_md_render()
    # test_md_to_word()
    # test_txt_to_image()
    # test_dedup_md_images()
