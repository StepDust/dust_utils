import os
import json
import sys
import io

from loguru import logger

from .diagram_render import DiagramRenderer
from .md_math import MathItem, MathPreprocessor

# Mermaid / LaTeX / 代码块的默认渲染选项，可在 convert(render={...}) 中覆盖
DEFAULT_RENDER_OPTIONS = {
    "mermaid": True,  # 是否渲染 ```mermaid 代码块
    "math": True,  # 是否渲染 LaTeX 公式
    "code_block": True,  # 普通代码块是否按等宽文本写入
    "allow_online": False,  # 是否允许在线服务兜底，默认全离线
    "mermaid_engine": "auto",  # auto / mmdc / browser（kroki / mermaid_ink 需 allow_online）
    "mermaid_theme": "default",
    "mermaid_background": "white",
    "mermaid_scale": 2.0,  # 图表截图清晰度倍数
    "mermaid_width": None,  # 插入宽度（英寸），None 表示按图表自然尺寸
    "latex_engine": "auto",  # auto / mathjax / matplotlib（codecogs 需 allow_online）
    "math_dpi": 240,
    "mermaid_js": None,  # 本地 mermaid.min.js 路径
    "mathjax_js": None,  # 本地 tex-svg.js 路径
    "browser_path": None,  # 本地 Edge / Chrome 路径
    "browser_size": (1800, 1300),
    "code_font": "Consolas",
    "cache_dir": None,
    "timeout": 60,
}

# region 辅助类


class MarkdownAstParser:
    """
    Markdown AST 解析器类
    封装 markdown-it 库的功能，用于将 Markdown 文本解析为 token 流
    """

    def __init__(self):
        """
        初始化 Markdown 解析器实例
        """
        # 延迟初始化 MarkdownIt，只有在 parse 时才导入并创建实例
        self.md = None

    def parse(self, md_text: str):
        """
        解析 Markdown 文本为 token 列表

        Args:
            md_text (str): 输入的 Markdown 文本内容

        Returns:
            list: 解析后的 token 列表
        """
        if self.md is None:
            from markdown_it import MarkdownIt

            # 启用 GFM 表格规则，否则默认 commonmark 预设不解析表格
            self.md = MarkdownIt().enable("table")
        return self.md.parse(md_text)


class ListNode:
    """
    列表节点类
    用于构建和存储 Markdown 列表的树状结构
    """

    def __init__(self, content="", ordered=False, level=0):
        """
        初始化列表节点

        Args:
            content (str): 列表项的内容文本
            ordered (bool): 是否为有序列表
            level (int): 列表嵌套层级（从0开始）
        """
        self.content = content
        self.ordered = ordered
        self.level = level
        self.children = []


# endregion


class MdToDocx:
    """
    Markdown 转 Word 文档转换器核心类
    负责处理 Markdown 解析、样式应用和 Word 文档生成
    """

    def __init__(self):
        """
        初始化转换器
        创建 AST 解析器实例
        """
        import random

        self.parser = MarkdownAstParser()
        # {random.randint(0, 0xFFFFFFFF):08X}
        self.numbering_content = ""

        # 渲染相关状态
        self.options = dict(DEFAULT_RENDER_OPTIONS)
        self.math = None  # MathPreprocessor，转换时按需创建
        self.renderer = None  # DiagramRenderer，首次渲染图表/公式时创建

    def _ensure_docx(self):
        """
        延迟导入 python-docx 相关模块并绑定到实例属性，避免模块导入期间的开销。
        """
        if getattr(self, "_docx_loaded", False):
            return

        from docx import Document
        from docx.shared import Inches, Pt, RGBColor
        from docx.oxml.ns import qn
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
        from docx.oxml import OxmlElement

        self.Document = Document
        self.Inches = Inches
        self.Pt = Pt
        self.RGBColor = RGBColor
        self.qn = qn
        self.WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH
        self.WD_LINE_SPACING = WD_LINE_SPACING
        self.OxmlElement = OxmlElement
        self._docx_loaded = True

    def convert(
        self,
        md_text: str,
        output_path: str,
        styles: dict = None,
        render: dict = None,
    ):
        """
        执行转换流程：Markdown -> Word

        Args:
            md_text (str): 原始 Markdown 文本
            output_path (str): 输出 Word 文档的路径 (.docx)
            styles (dict, optional): 样式配置，形如 {"h1": {...}, "text": {...}}；
                传 None 时按输出目录、当前工作目录、程序目录依次查找 style.json
            render (dict, optional): 渲染选项，可覆盖 DEFAULT_RENDER_OPTIONS，
                常用项：mermaid / math / code_block / mermaid_engine / latex_engine
        """
        # 确保按需加载 python-docx
        self._ensure_docx()

        self.options = dict(DEFAULT_RENDER_OPTIONS)
        if render:
            self.options.update({k: v for k, v in render.items() if v is not None})

        # 每次转换都重新收集列表编号，避免同一实例多次转换时编号重复累积
        self.numbering_content = ""

        if os.path.exists(output_path):
            self.doc = self.Document(output_path)
        else:
            self.doc = self.Document()

        self._enable_doc_grid()
        self._ensure_default_styles()

        if styles is None:
            styles = self._load_style_file(output_path)
        self.styles = styles or {}

        # 公式先替换成占位符，解析完 token 后再还原为图片
        self.math = None
        if self.options.get("math"):
            self.math = MathPreprocessor()
            md_text = self.math.extract(md_text)

        logger.info("正在解析Markdown文档结构...")
        tokens = self.parser.parse(md_text)

        logger.info("正在写入Word文档...")
        self._write_tokens(tokens)
        self.doc.save(output_path)

        self._write_extra(output_path)
        logger.success("Word文档写入完毕！")

    def _load_style_file(self, output_path: str) -> dict:
        """
        查找并加载 style.json，兼容 {"default": {...}} 与 {"docx_theme": {"default": {...}}} 两种结构
        """
        candidates = [
            os.path.join(os.path.dirname(os.path.abspath(output_path)), "style.json"),
            os.path.join(os.getcwd(), "style.json"),
            os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "style.json"),
        ]

        for path in candidates:
            if not os.path.exists(path):
                continue
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as exc:
                logger.warning(f"样式文件解析失败：{path} -> {exc}")
                continue

            theme = data.get("default")
            if theme is None and isinstance(data.get("docx_theme"), dict):
                theme = data["docx_theme"].get("default")
            if isinstance(theme, dict) and theme:
                logger.info(f"已加载样式文件：{path}")
                return theme

        logger.warning("未找到可用的 style.json，将使用 Word 默认样式")
        return {}

    # region 写入内容

    def _write_tokens(self, tokens, paragraph=None, paragraph_style=None):
        """
        遍历并处理 Markdown token 流，将其转换为 Word 文档元素

        Args:
            tokens (list): token 列表
            paragraph (docx.text.paragraph.Paragraph, optional): 当前正在处理的段落对象. Defaults to None.
            paragraph_style (str, optional): 当前段落应用的样式名称. Defaults to None.
        """
        i = 0
        while i < len(tokens):
            t = tokens[i]

            # 解析标题级别
            if t.type == "heading_open":
                self._current_heading_level = int(t.tag[1])

            # 解析标题内容（标题内也可能有公式/加粗，统一交给行内处理）
            elif t.type == "inline" and hasattr(self, "_current_heading_level"):
                level = self._current_heading_level
                if paragraph is None:
                    paragraph = self.doc.add_heading("", level=level)
                self._set_paragraph_style(paragraph, f"h{level}")
                self._handle_inline(t, paragraph, f"h{level}")
                paragraph = None
                # 重置标题级别
                del self._current_heading_level

            # 列表处理
            elif t.type in ("bullet_list_open", "ordered_list_open"):
                # 先解析当前列表为树结构
                nodes, next_i = self._parse_list(tokens, i)
                # 再写入 Word
                self._write_list_to_word(nodes)
                i = next_i
                continue  # 跳过 i += 1

            # 表格处理
            elif t.type == "table_open":
                # 找到 table_close，整段交给表格写入
                end = i
                while end < len(tokens) and tokens[end].type != "table_close":
                    end += 1
                self._write_table(tokens[i : end + 1])
                i = end
                continue  # 跳过 i += 1

            # 代码块 / Mermaid 图表
            elif t.type in ("fence", "code_block"):
                self._handle_code_block(t)

            # 普通段落 + 图片 + 公式
            elif t.type == "inline":
                item = (
                    self.math.display_item(t.content)
                    if (self.math is not None and paragraph is None)
                    else None
                )
                if item is not None:
                    # 整段就是块级公式，图片独占一段并居中
                    self._add_display_math(item)
                else:
                    self._handle_inline(t, paragraph, paragraph_style)

            i += 1

    def _handle_inline(self, token, paragraph=None, paragraph_style=None):
        """
        处理行内元素：文本、加粗/斜体、行内代码、图片、行内公式、换行

        Args:
            token (Token): 当前处理的 inline token
            paragraph (docx.text.paragraph.Paragraph, optional): 目标段落对象；
                为 None 时新建段落，已有段落（标题/列表/单元格）则只追加 run
            paragraph_style (str, optional): 样式名称
        """
        style_name = paragraph_style or "text"
        style = self.styles.get(style_name) or {}

        if paragraph is None:
            paragraph = self.doc.add_paragraph()
            self._set_paragraph_style(paragraph, style_name)

        bold = None
        italic = None

        for child in token.children or []:
            ctype = child.type

            if ctype == "strong_open":
                bold = True
            elif ctype == "strong_close":
                bold = None
            elif ctype == "em_open":
                italic = True
            elif ctype == "em_close":
                italic = None
            elif ctype == "text":
                self._add_text(child.content, paragraph, style, bold, italic)
            elif ctype == "softbreak":
                self._add_run(paragraph, " ", style, bold, italic)
            elif ctype == "hardbreak":
                paragraph.add_run().add_break()
            elif ctype == "html_inline":
                if (child.content or "").strip().lower() in ("<br>", "<br/>", "<br />"):
                    paragraph.add_run().add_break()
            elif ctype == "code_inline":
                self._add_run(paragraph, child.content, style, bold, italic, mono=True)
            elif ctype == "image":
                self._add_image(child.attrs.get("src"), paragraph)

    def _add_text(self, text, paragraph, style, bold=None, italic=None):
        """写入一段文本，其中的公式占位符会被替换为公式图片。"""
        if not text:
            return

        for part in self.math.split(text) if self.math is not None else [text]:
            if isinstance(part, MathItem):
                self._add_math(part, paragraph, style)
            elif part:
                self._add_run(paragraph, part, style, bold, italic)

    def _add_run(
        self,
        paragraph,
        text,
        style,
        bold=None,
        italic=None,
        mono=False,
    ):
        """添加一个 run 并按样式配置设置字体。"""
        if not text:
            return None
        run = paragraph.add_run(text)
        self._format_run(run, style, bold=bold, italic=italic, mono=mono)
        return run

    # region 代码块 / 图表 / 公式

    def _handle_code_block(self, token):
        """处理 ``` 代码块：mermaid 渲染成图，其余按等宽文本写入。"""
        info = (getattr(token, "info", "") or "").strip().split(" ")[0].lower()
        content = token.content or ""
        if not content.strip():
            return

        if self.options.get("mermaid") and self._is_mermaid(info, content):
            self._add_mermaid(content)
        elif self.options.get("code_block"):
            self._add_code_block(content, info)

    @staticmethod
    def _is_mermaid(info: str, content: str) -> bool:
        """判断代码块是否为 Mermaid：优先看语言标记，没有标记时看首行关键字。"""
        if info in ("mermaid", "mmd"):
            return True
        if info:
            return False
        first = content.strip().split("\n", 1)[0].strip().lower()
        keywords = (
            "graph ",
            "flowchart ",
            "sequencediagram",
            "classdiagram",
            "statediagram",
            "erdiagram",
            "gantt",
            "pie",
            "mindmap",
            "timeline",
            "journey",
            "quadrantchart",
            "gitgraph",
        )
        return first.startswith(keywords)

    def _add_mermaid(self, code: str):
        """渲染 Mermaid 并插入为居中图片，失败时降级为代码文本。"""
        try:
            result = self._get_renderer().render_mermaid(
                code,
                engine=self.options.get("mermaid_engine", "auto"),
                theme=self.options.get("mermaid_theme", "default"),
                background=self.options.get("mermaid_background", "white"),
                scale=self.options.get("mermaid_scale", 2.0),
            )
        except Exception as exc:
            logger.warning(f"Mermaid 渲染失败，已降级为代码文本：{exc}")
            self._add_code_block(code, "mermaid")
            return

        self._insert_picture(
            result,
            self._new_centered_paragraph(),
            width_inch=self.options.get("mermaid_width"),
        )

    def _add_display_math(self, item: MathItem):
        """块级公式：独占一段并居中。"""
        paragraph = self._new_centered_paragraph()
        self._add_math(item, paragraph, self.styles.get("text") or {})

    def _add_math(self, item: MathItem, paragraph, style):
        """把公式渲染成图片插入段落，失败时退回原文。"""
        font_size = style.get("font_size", 11)
        try:
            result = self._get_renderer().render_latex(
                item.tex,
                engine=self.options.get("latex_engine", "auto"),
                display=item.display,
                font_size=font_size,
                dpi=self.options.get("math_dpi", 240),
            )
        except Exception as exc:
            logger.warning(f"公式渲染失败，已退回原文：{exc}")
            fallback = f"$${item.tex}$$" if item.display else f"${item.tex}$"
            self._add_run(paragraph, fallback, style)
            return

        self._insert_picture(result, paragraph)

    def _add_code_block(self, code: str, info: str = ""):
        """把代码块按等宽字体逐行写入（不缩进、单倍行距）。"""
        style_name = "code" if "code" in self.styles else "text"
        style = self.styles.get(style_name) or {}

        lines = code.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n").split("\n")
        for line in lines:
            paragraph = self.doc.add_paragraph()
            self._set_paragraph_style(paragraph, style_name)

            fmt = paragraph.paragraph_format
            fmt.first_line_indent = self.Pt(0)
            fmt.space_before = self.Pt(0)
            fmt.space_after = self.Pt(0)
            fmt.line_spacing = 1.0

            self._add_run(paragraph, line if line.strip() else " ", style, mono=True)

    def _new_centered_paragraph(self):
        """新建一个无首行缩进、居中的段落（用于图片 / 块级公式）。"""
        paragraph = self.doc.add_paragraph()
        self._set_paragraph_style(paragraph, "text")

        fmt = paragraph.paragraph_format
        fmt.first_line_indent = self.Pt(0)
        fmt.alignment = self.WD_ALIGN_PARAGRAPH.CENTER
        fmt.space_before = self.Pt(3)
        fmt.space_after = self.Pt(6)
        return paragraph

    def _insert_picture(self, result, paragraph, align: str = None, width_inch=None):
        """
        把渲染结果插入段落，按页面可用宽高等比缩放（不放大）。

        Args:
            result (RenderResult): 渲染结果
            paragraph: 目标段落
            align: left / center / right，None 表示不改动段落对齐
            width_inch: 指定插入宽度（英寸），None 表示按图片自然尺寸
        """
        max_width, max_height = self._get_available_size(paragraph)

        width = width_inch or result.width_inch or 1.0
        if result.width_inch:
            height = width * (result.height_inch / result.width_inch)
        else:
            height = result.height_inch or 1.0

        scale = 1.0
        if max_width > 0:
            scale = min(scale, max_width / width)
        if max_height > 0:
            scale = min(scale, max_height / height)
        scale = max(scale, 0.1)

        run = paragraph.add_run()
        run.add_picture(result.path, width=self.Inches(width * scale))

        if align:
            paragraph.paragraph_format.alignment = {
                "left": self.WD_ALIGN_PARAGRAPH.LEFT,
                "center": self.WD_ALIGN_PARAGRAPH.CENTER,
                "right": self.WD_ALIGN_PARAGRAPH.RIGHT,
            }.get(align)

    def _get_available_size(self, paragraph=None):
        """返回段落可用宽高（英寸），已扣除页边距与列表缩进。"""
        section = self.doc.sections[0]
        emu = 914400

        width = (
            (section.page_width or 0)
            - (section.left_margin or 0)
            - (section.right_margin or 0)
        )
        height = (
            (section.page_height or 0)
            - (section.top_margin or 0)
            - (section.bottom_margin or 0)
        )

        if paragraph is not None:
            level = self._get_list_level(paragraph)
            if level >= 0:
                width -= self.Inches(0.25 * (level + 1))

        return max(width / emu, 0.0), max(height / emu, 0.0)

    def _get_renderer(self) -> DiagramRenderer:
        """延迟创建 Mermaid / LaTeX 渲染器（默认全离线）。"""
        if self.renderer is None:
            self.renderer = DiagramRenderer(
                cache_dir=self.options.get("cache_dir"),
                timeout=self.options.get("timeout", 60),
                allow_online=bool(self.options.get("allow_online")),
                mermaid_js=self.options.get("mermaid_js"),
                mathjax_js=self.options.get("mathjax_js"),
                browser_path=self.options.get("browser_path"),
                mermaid_scale=self.options.get("mermaid_scale", 2.0),
                browser_size=self.options.get("browser_size", (1800, 1300)),
            )
        return self.renderer

    # endregion

    @logger.catch
    def _add_image(self, src: str, paragraph=None):
        """
        向段落添加图片，并自动计算合适的显示宽度
        支持列表内图片宽度正确计算

        Args:
            src (str): 图片来源 (路径/URL/Base64)
            paragraph (docx.text.paragraph.Paragraph, optional): 目标段落. Defaults to None.
        """

        # 延迟导入 requests/base64/tempfile 等
        self._ensure_docx()
        from urllib.parse import unquote
        import requests
        import base64
        import tempfile

        # ---------- 1️⃣ 获取图片 ----------
        if src.startswith("http"):
            src = unquote(src)
            r = requests.get(src, timeout=10)
            r.raise_for_status()
            img_bytes = r.content

        elif src.startswith("data:image"):
            img_bytes = base64.b64decode(src.split(",", 1)[1])

        else:
            src = unquote(src)

            with open(src, "rb") as f:
                img_bytes = f.read()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            tmp.write(img_bytes)
            tmp_path = tmp.name

        # ---------- 2️⃣ paragraph ----------
        if paragraph is None:
            paragraph = self.doc.add_paragraph()

        # ---------- 3️⃣ 计算宽度 ----------
        section = self.doc.sections[0]

        page_width = section.page_width
        margin_left = section.left_margin
        margin_right = section.right_margin

        # 🔥 核心：列表层级缩进
        list_level = self._get_list_level(paragraph)
        list_indent = self.Inches(0.25 * (list_level + 1))

        available_width = page_width - margin_left - margin_right - list_indent

        # 防御
        if isinstance(available_width, int):
            available_width = self.Inches(available_width / 914400)

        try:
            inches_val = available_width.inches
        except Exception:
            inches_val = None

        if inches_val is None or (
            isinstance(inches_val, (int, float)) and inches_val < 1
        ):
            available_width = self.Inches(1)

        # ---------- 4️⃣ 插入 ----------
        if len(paragraph.runs) > 0:
            paragraph.add_run("\n")

        image_path = self.clean_image(tmp_path)

        # 获取图片原始宽度
        from PIL import Image

        with Image.open(image_path) as img:
            width_px, height_px = img.size

            # 获取图片 DPI
            dpi = img.info.get("dpi", (96, 96))[0]

            # 转换为英寸
            original_width = width_px / dpi

        # 判断是否超过可用宽度，标准宽度界限为576px
        if original_width > available_width.inches:
            # 超出，限制为100%
            image_width = available_width
        else:
            # 未超出，保持原始尺寸
            image_width = self.Inches(original_width)

        run = paragraph.add_run()
        run.add_picture(image_path, width=image_width)

    def clean_image(self, path):
        """
        清理图片 EXIF 元数据，避免 Word 插入时出现尺寸异常等问题
        """
        from PIL import Image

        img = Image.open(path)

        # 关键：去掉 EXIF
        data = list(img.getdata())
        img2 = Image.new(img.mode, img.size)
        img2.putdata(data)

        buf = io.BytesIO()
        img2.save(buf, format="PNG")
        buf.seek(0)
        return buf

    def _write_extra(self, output_path):

        from zipfile import ZipFile, ZIP_DEFLATED
        import tempfile
        import os
        import shutil

        if not os.path.exists(output_path):
            return

        if not self.numbering_content:
            return

        logger.info("有序列表重新编号中...")

        temp_path = output_path + ".tmp"

        if self.numbering_content.strip():
            abstractNum = """
            
<w:abstractNum w:abstractNumId="999">
  <w:multiLevelType w:val="multilevel" />

  <!-- 一级 -->
  <w:lvl w:ilvl="0">
    <w:start w:val="1" />
    <w:numFmt w:val="decimal" />
    <w:lvlText w:val="%1." />
    <w:lvlJc w:val="left" />
    <w:pPr>
      <w:ind w:left="360" w:hanging="360" />
    </w:pPr>
  </w:lvl>

  <!-- 二级 -->
  <w:lvl w:ilvl="1">
    <w:start w:val="1" />
    <w:numFmt w:val="decimal" />
    <w:lvlText w:val="%1.%2" />
    <w:lvlJc w:val="left" />
    <w:pPr>
      <w:ind w:left="720" w:hanging="360" />
    </w:pPr>
  </w:lvl>

  <!-- 三级 -->
  <w:lvl w:ilvl="2">
    <w:start w:val="1" />
    <w:numFmt w:val="decimal" />
    <w:lvlText w:val="%1.%2.%3" />
    <w:lvlJc w:val="left" />
    <w:pPr>
      <w:ind w:left="1080" w:hanging="360" />
    </w:pPr>
  </w:lvl>

</w:abstractNum>

"""

            # 目前这个写法可以在打包后，更新docx的内容
            # with ZipFile(output_path, "a") as docx 打包后无法更新docx内容，需要重新打包
            with ZipFile(output_path, "r") as zin:
                with ZipFile(temp_path, "w", ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)

                        if item.filename == "word/numbering.xml":
                            text = data.decode("utf-8")

                            text = text.replace(
                                "</w:numbering>",
                                abstractNum
                                + self.numbering_content
                                + "\n</w:numbering>",
                            )

                            data = text.encode("utf-8")

                        zout.writestr(item, data)

        shutil.move(temp_path, output_path)

    # endregion

    # region 列表处理

    def _parse_list(self, tokens, start=0, level=0):
        """
        解析列表 tokens 为 ListNode 树状结构 (递归)

        Args:
            tokens (list): token 列表
            start (int): 当前解析起始索引
            level (int): 当前列表层级

        Returns:
            tuple: (解析得到的节点列表, 下一个处理索引)
        """
        nodes = []
        i = start
        ordered = tokens[start].type == "ordered_list_open"

        while i < len(tokens):
            t = tokens[i]

            if t.type == "list_item_open":
                node = ListNode(level=level, ordered=ordered)
                j = i + 1
                while j < len(tokens) and tokens[j].type != "list_item_close":
                    if tokens[j].type == "inline":
                        node.content = tokens[j].content
                    # 子列表
                    elif tokens[j].type in ("bullet_list_open", "ordered_list_open"):
                        child_nodes, next_j = self._parse_list(tokens, j, level + 1)
                        node.children.extend(child_nodes)
                        j = next_j - 1  # 解析完子列表后更新 j
                    j += 1
                nodes.append(node)
                i = j

            elif t.type in ("bullet_list_open", "ordered_list_open"):
                # 外层列表开始，更新 ordered
                ordered = t.type == "ordered_list_open"

            elif t.type in ("bullet_list_close", "ordered_list_close"):
                return nodes, i + 1

            i += 1

        return nodes, i

    def _write_list_to_word(self, nodes):
        """
        将列表节点树写入 Word 文档

        Args:
            nodes (list[ListNode]): 列表节点集合
        """

        def generate_unique_random():
            import time
            import random

            # 使用当前时间微秒 + random 组合，确保毫秒级唯一性
            timestamp = int(time.time() * 1000000) % 100000  # 取微秒部分
            rand_part = random.randint(0, 9999)
            num = 10000 + (timestamp + rand_part) % 90000
            return num

        def _write_nodes(nodes, level=0):
            for i, node in enumerate(nodes):
                num_id = generate_unique_random()
                style = self._get_list_style(node.ordered, level)

                paragraph = self.doc.add_paragraph(style=style)

                _write_numId(paragraph, num_id, level, i, style)

                self._set_paragraph_style(paragraph, f"li")
                tokens = self.parser.parse(node.content)
                self._write_tokens(tokens, paragraph=paragraph, paragraph_style=f"li")

                # 递归写入子列表
                if node.children:
                    children_num_id = generate_unique_random()
                    _write_nodes(node.children, level + 1)

        def _write_numId(paragraph, num_id, level, index, style):
            """
            写入段落的 numId，用于列表重新编号
            """
            if num_id is None or num_id == "":
                return

            if "number" not in style.lower():
                return

            p = paragraph._p
            pPr = p.get_or_add_pPr()

            ilvl = self.OxmlElement("w:ilvl")
            ilvl.set(self.qn("w:val"), str(level))

            numPr = self.OxmlElement("w:numPr")
            numId = self.OxmlElement("w:numId")
            numId.set(self.qn("w:val"), str(num_id))

            numPr.append(ilvl)
            numPr.append(numId)
            pPr.append(numPr)

            if index == 0:
                self.numbering_content += f"""
    <w:num w:numId="{num_id}">
        <w:abstractNumId w:val="999"/>
        <w:lvlOverride w:ilvl="{level}">
            <w:startOverride w:val="1"/>
        </w:lvlOverride>
    </w:num>

    """
            else:
                self.numbering_content += f"""
    <w:num w:numId="{num_id}">
        <w:abstractNumId w:val="999"/>
    </w:num>

    """

        _write_nodes(nodes)

    def _get_list_style(self, ordered, level):
        """
        根据列表类型和层级获取 Word 列表样式名

        Args:
            ordered (bool): 是否有序
            level (int): 层级

        Returns:
            str: 样式名称
        """

        if ordered:
            styles = [
                "List Number",
                "List Number 2",
                "List Number 3",
            ]

        else:
            styles = [
                "List Bullet",
                "List Bullet 2",
                "List Bullet 3",
            ]

        style = styles[min(level, len(styles) - 1)]

        return style

    def _get_list_level(self, paragraph):
        """
        获取段落的列表层级

        Args:
            paragraph (docx.text.paragraph.Paragraph): 段落对象

        Returns:
            int: 列表层级 (从0开始), 非列表返回 -1
        """

        p = paragraph._p
        pPr = p.pPr

        # 1️⃣ OOXML 编号列表
        if pPr is not None and pPr.numPr is not None:
            ilvl = pPr.numPr.ilvl
            if ilvl is not None:
                return int(ilvl.val)

        # 2️⃣ 样式列表（List Bullet / List Number）
        style = paragraph.style
        if style and style.name:
            name = style.name.lower()
            if name.startswith("list"):
                # 从样式名中提取数字作为层级
                import re

                match = re.search(r"(\d+)", name)
                return int(match.group(1)) - 1 if match else 0

        # 3️⃣ 非列表
        return -1

    # endregion

    # region 表格处理

    def _write_table(self, table_tokens):
        """
        将表格 tokens 写入 Word 文档

        单元格样式：优先使用样式配置中的 table 样式，未配置时回退到 text 样式；
        单元格内容固定左对齐、无首行缩进、无段前段后间距，表头行加粗。

        Args:
            table_tokens (list): table_open 到 table_close 之间的 token 列表
        """
        self._ensure_docx()
        rows = self._parse_table_rows(table_tokens)
        if not rows:
            return

        col_count = max(len(cells) for _, cells in rows)
        if col_count == 0:
            return

        table = self.doc.add_table(rows=len(rows), cols=col_count)
        try:
            table.style = "Table Grid"
        except KeyError:
            logger.warning("文档模板中不存在 Table Grid 样式，表格将不显示边框")

        for row_index, (is_header, cells) in enumerate(rows):
            for col_index, inline_token in enumerate(cells):
                paragraph = table.cell(row_index, col_index).paragraphs[0]
                if inline_token is not None:
                    self._handle_inline(
                        inline_token, paragraph=paragraph, paragraph_style="text"
                    )
                self._set_table_cell_style(paragraph)
                if is_header:
                    for run in paragraph.runs:
                        run.font.bold = True

    def _parse_table_rows(self, table_tokens):
        """
        解析表格 tokens 为行结构

        Args:
            table_tokens (list): table_open 到 table_close 之间的 token 列表

        Returns:
            list[tuple[bool, list]]: 每行为 (是否表头, 单元格 inline token 列表)
        """
        rows = []
        in_header = False

        i = 0
        while i < len(table_tokens):
            t = table_tokens[i]

            if t.type == "thead_open":
                in_header = True
            elif t.type == "thead_close":
                in_header = False
            elif t.type == "tr_open":
                cells = []
                j = i + 1
                while j < len(table_tokens) and table_tokens[j].type != "tr_close":
                    if table_tokens[j].type in ("th_open", "td_open"):
                        inline_token = None
                        j += 1
                        # 收集单元格内的 inline token
                        while j < len(table_tokens) and table_tokens[j].type not in (
                            "th_close",
                            "td_close",
                        ):
                            if table_tokens[j].type == "inline":
                                inline_token = table_tokens[j]
                            j += 1
                        cells.append(inline_token)
                    j += 1
                rows.append((in_header, cells))
                i = j

            i += 1

        return rows

    def _set_table_cell_style(self, paragraph):
        """
        表格单元格段落样式：套用 table 样式（未配置时回退 text 样式），
        并固定左对齐、无首行缩进、无段前段后间距

        Args:
            paragraph (docx.text.paragraph.Paragraph): 单元格段落对象
        """
        style_name = "table" if "table" in self.styles else "text"
        self._set_paragraph_style(paragraph, style_name)

        p_format = paragraph.paragraph_format
        p_format.alignment = self.WD_ALIGN_PARAGRAPH.LEFT
        p_format.first_line_indent = self.Pt(0)
        p_format.space_before = self.Pt(0)
        p_format.space_after = self.Pt(0)

    # endregion

    # region 样式设置

    def _ensure_default_styles(self):
        """补齐模板中缺失的 python-docx 默认样式。"""

        from docx import Document
        from docx.enum.style import WD_STYLE_TYPE

        styles = self.doc.styles
        default_styles = Document().styles

        for source_style in default_styles:
            try:
                styles[source_style.name]
                continue
            except KeyError:
                pass

            try:
                target_style = styles.add_style(
                    source_style.name,
                    source_style.type,
                )
            except Exception:
                continue

            # 基础样式
            if source_style.base_style is not None:
                try:
                    target_style.base_style = styles[source_style.base_style.name]
                except KeyError:
                    pass

            # 字体
            try:
                target_style.font.name = source_style.font.name
                target_style.font.size = source_style.font.size
                target_style.font.bold = source_style.font.bold
                target_style.font.italic = source_style.font.italic
                target_style.font.underline = source_style.font.underline
            except Exception:
                pass

            # 段落
            if source_style.type == WD_STYLE_TYPE.PARAGRAPH:
                try:
                    source_format = source_style.paragraph_format
                    target_format = target_style.paragraph_format

                    target_format.alignment = source_format.alignment
                    target_format.left_indent = source_format.left_indent
                    target_format.right_indent = source_format.right_indent
                    target_format.first_line_indent = source_format.first_line_indent
                    target_format.space_before = source_format.space_before
                    target_format.space_after = source_format.space_after
                    target_format.line_spacing = source_format.line_spacing
                    target_format.keep_together = source_format.keep_together
                    target_format.keep_with_next = source_format.keep_with_next
                    target_format.page_break_before = source_format.page_break_before
                    target_format.widow_control = source_format.widow_control
                except Exception:
                    pass

    def _set_paragraph_style(self, paragraph, style_name):
        """
        应用段落和字体样式
        完整版：
        - 支持任意倍数行距下的“视觉垂直居中”
        - 与人工 Word/WPS 文档行为一致
        - 不破坏列表（li）编号结构

        Args:
            paragraph (docx.text.paragraph.Paragraph): 目标段落
            style_name (str): 样式配置名称
        """
        # 确保 docx 系列符号已按需加载
        self._ensure_docx()
        style = self.styles.get(style_name)
        if not style:
            return

        # =====================================================
        # 1️⃣ 段落级格式（pPr）
        # =====================================================
        p_format = paragraph.paragraph_format
        pPr = paragraph._element.get_or_add_pPr()

        # ---- textAlignment = auto（与人工文档一致）
        text_align = pPr.find(self.qn("w:textAlignment"))
        if text_align is None:
            text_align = self.OxmlElement("w:textAlignment")
            pPr.append(text_align)
        text_align.set(self.qn("w:val"), "auto")

        # ---- snapToGrid = 1（关键：启用基线网格）
        snap = pPr.find(self.qn("w:snapToGrid"))
        if snap is None:
            snap = self.OxmlElement("w:snapToGrid")
            pPr.append(snap)
        snap.set(self.qn("w:val"), "1")

        # ---- 中文排版辅助属性（不影响西文）
        for tag in ["w:kinsoku", "w:overflowPunct", "w:adjustRightInd"]:
            if pPr.find(self.qn(tag)) is None:
                pPr.append(self.OxmlElement(tag))

        # ---- 对齐方式
        if "align" in style:
            align_map = {
                "left": self.WD_ALIGN_PARAGRAPH.LEFT,
                "center": self.WD_ALIGN_PARAGRAPH.CENTER,
                "right": self.WD_ALIGN_PARAGRAPH.RIGHT,
                "justify": self.WD_ALIGN_PARAGRAPH.JUSTIFY,
            }
            p_format.alignment = align_map.get(
                style["align"], self.WD_ALIGN_PARAGRAPH.LEFT
            )

        # ---- 行距（不固定倍数）
        if "line_spacing" in style:
            p_format.line_spacing = style["line_spacing"]
            p_format.line_spacing_rule = self.WD_LINE_SPACING.MULTIPLE

        p_format.space_before = self.Pt(style.get("space_before", 0))
        p_format.space_after = self.Pt(style.get("space_after", 0))

        # =====================================================
        # 2️⃣ 缩进规则
        # =====================================================
        if "first_line_indent" in style:
            p_format.first_line_indent = self.Pt(
                style.get("font_size", 11) * style["first_line_indent"]
            )

        # 🔥 li 特殊处理
        if style_name == "li":
            p_format.left_indent = None
            p_format.first_line_indent = None

        # =====================================================
        # 3️⃣ 字符级（rPr）
        # =====================================================
        for run in paragraph.runs:
            self._format_run(run, style)

    def _format_run(self, run, style, bold=None, italic=None, mono=False):
        """
        设置单个 run 的字体样式

        Args:
            run: docx 的 run 对象
            style (dict): 样式配置
            bold (bool, optional): 是否加粗，None 表示沿用样式配置
            italic (bool, optional): 是否斜体，None 表示沿用样式配置
            mono (bool): 是否使用等宽字体（行内代码 / 代码块）
        """
        self._ensure_docx()

        font_name = style.get("font_name", "微软雅黑")
        font_size = style.get("font_size", 11)

        if mono:
            font_name = self.options.get("code_font", "Consolas")
            font_size = max(font_size - 1, 8)

        font = run.font
        rPr = run._element.get_or_add_rPr()

        # ---- 字体四槽位（ascii / hAnsi / eastAsia / cs），
        # 解决“中文字体不生效”问题
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(self.qn("w:ascii"), font_name)
        rFonts.set(self.qn("w:hAnsi"), font_name)
        rFonts.set(self.qn("w:eastAsia"), font_name)
        rFonts.set(self.qn("w:cs"), font_name)

        # ---- 语言环境
        lang = rPr.find(self.qn("w:lang"))
        if lang is None:
            lang = self.OxmlElement("w:lang")
            rPr.append(lang)
        lang.set(self.qn("w:val"), "en-US")
        lang.set(self.qn("w:eastAsia"), "zh-CN")

        # ---- 基础样式
        font.size = self.Pt(font_size)
        font.bold = style.get("bold", False) if bold is None else bold
        font.italic = style.get("italic", False) if italic is None else italic
        font.underline = style.get("underline", False)

        if "font_color" in style:
            c = style["font_color"].lstrip("#")
            font.color.rgb = self.RGBColor(
                int(c[0:2], 16),
                int(c[2:4], 16),
                int(c[4:6], 16),
            )

    def _enable_doc_grid(self):
        """
        启用 Word 中文排版网格，实现段落文字垂直居中，只需要执行一次即可
        - 始终确保 w:type="lines"
        - 不强制 linePitch，由 Word 自动计算
        """
        # 确保 docx 相关对象已加载
        self._ensure_docx()
        section = self.doc.sections[0]
        sectPr = section._sectPr

        # 查找已有 docGrid
        docGrids = sectPr.xpath(".//w:docGrid")
        if docGrids:
            docGrid = docGrids[0]
        else:
            docGrid = self.OxmlElement("w:docGrid")
            sectPr.append(docGrid)

        # ✅ 核心：始终设置为 lines
        docGrid.set(self.qn("w:type"), "lines")

        # 可选：显式关闭字符网格（推荐）
        docGrid.set(self.qn("w:charSpace"), "0")

    # endregion
