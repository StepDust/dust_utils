from dust_utils import setup_loguru
import random
from dust_utils.file_utils import WordUtils

logger = setup_loguru()

from dust_utils.file_utils.md_to_docx import MdToDocx
from dust_utils.ai_utils import AIChat
import os
import json
from dotenv import load_dotenv

load_dotenv(r"E:\Share\配置文件\.env")


def test_md_to_word():
    md_path = r"C:\Users\Administered\Desktop\国土空间规划制图与数据分析系统 使用说明书V1.0.md"
    fm_folder = r"C:\Users\Administered\Desktop\封面"

    fm_list = [
        os.path.join(fm_folder, f)
        for f in os.listdir(fm_folder)
        if os.path.isfile(os.path.join(fm_folder, f)) and f.lower().endswith(".docx")
    ]

    with open(r"test_config\md_to_config\style.json", "r", encoding="utf-8") as f:
        styles = json.load(f)

    color_list = styles["color_list"]

    cover_desc = [
        "软著说明书",
        "说明书",
        "使用说明书",
        "软件说明书",
        "软件使用说明书",
        "软件操作说明书",
        "软件功能说明书",
        "软件技术说明书",
        "软件应用明书",
        "软件运行说明书",
        "软件使用手册",
        "软件操作手册",
        "软件用户手册",
        "软件功能手册",
        "软件应用手册",
        "用户使用手册",
        "用户操作手册",
        "用户说明书",
        "系统使用说明书",
        "系统操作说明书",
        "系统说明书",
        "系统功能说明书",
        "系统应用说明书",
        "系统运行说明书",
        "系统使用手册",
        "系统操作手册",
        "系统用户册",
        "软件使用指南",
        "软件操作指南",
        "软件用户指南",
        "软件应用指南",
        "软件使用指引",
        "软件操作指引",
        "系统使用指南",
        "系统操作指南",
        "系统用户指南",
        "软件说明",
        "软件使用说明",
        "软件操作说明",
        "软件功能说明",
    ]
    font_name_list = [
        "宋体",
        "黑体",
        "楷体",
        "仿宋",
        "新宋体",
        "微软雅黑",
        "等线",
        "等线 Light",
        "隶书",
        "幼圆",
    ]
    case_name_colors = [
        # 黑 / 深灰
        "#000000",
        "#1D1D1F",
        "#202124",
        "#24292F",
        "#37352F",
        "#334155",
        # 蓝
        "#1F4E79",
        "#2F75B5",
        "#1A73E8",
        "#0969DA",
        "#0066CC",
        "#1677FF",
        # 靛蓝 / 紫蓝
        "#283593",
        "#3949AB",
        # 绿
        "#385723",
        "#548235",
        # 深灰蓝
        "#1E293B",
        "#374151",
    ]
    cover_desc_colors = [
        # 黑灰
        "#333333",
        "#404040",
        "#555555",
        "#666666",
        "#787774",
        "#86868B",
        # 蓝
        "#5B9BD5",
        "#5C6BC0",
        "#4096FF",
        "#64748B",
        # 中灰蓝
        "#57606A",
        "#475569",
        # 绿
        "#548235",
        # 紫蓝
        "#3949AB",
    ]
    page_number_formats = [
        # 纯数字
        ("", ""),
        # 中文
        ("第 ", " 页"),
        ("第", "页"),
        ("第 ", "页"),
        ("页 ", ""),
        ("页码 ", ""),
        # 英文
        ("Page ", ""),
        ("No. ", ""),
        ("P. ", ""),
        ("Pg. ", ""),
        # 括号
        ("(", ")"),
        ("（", "）"),
        ("[", "]"),
        ("【", "】"),
        # 分隔符
        ("— ", " —"),
        ("- ", " -"),
        ("· ", " ·"),
        ("• ", " •"),
        ("/ ", " /"),
        ("| ", " |"),
        # 更偏文档风格
        ("第 ", " 页"),
        ("页码：", ""),
        ("页码 ", ""),
        ("Page ", ""),
        ("PAGE ", ""),
        # 带空格的简洁样式
        ("[ ", " ]"),
        ("( ", " )"),
        ("— ", " —"),
    ]

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

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
            header_align = random.choice(
                [
                    "left",
                    "center",
                ]
            )

            # 页码位置
            page_number_position = random.choice(
                [
                    "header",
                    "footer",
                ]
            )

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
    test_md_to_word()
    # test_txt_to_image()
    # test_loguru()
