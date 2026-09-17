import re
from pathlib import Path
import os
from typing import TYPE_CHECKING

# 初始化日志
from loguru import logger

# 用于类型注解
if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph
    from docx.table import Table, _Cell


class WordUtils:
    """
    word工具类
    使用库：python-docx
    """

    # region 变量替换

    @staticmethod
    def replace_vars(docx_path: str, params: dict):
        """
        替换word文档中的变量：支持 {{ var }} / {{ va r }} / {{Var}} 等模糊变量替换

        Args:
            docx_path (str): 输入 Word 文档的路径 (.docx)
            params (dict): 形如{{var}}，变量替换字典，键为变量名，值为替换值
        """
        logger.info(f"开始替换文档变量\n路径： {docx_path}\n变量：{params}")
        from docx import Document
        from docx.text.paragraph import Paragraph
        from docx.oxml.ns import qn

        doc = Document(docx_path)

        # 处理普通段落
        for p in doc.paragraphs:
            WordUtils._process_paragraph(p, params)

        # 处理表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    WordUtils._process_cell(cell, params)

        WordUtils._process_textboxes(doc, params)

        doc.save(docx_path)
        logger.success(f"替换文档变量完成\n路径： {docx_path}")

    @staticmethod
    def replace_vars_fuzzy(text: str, params: dict) -> str:
        """
        支持 {{ var }} / {{ va r }} / {{Var}} 等模糊变量替换

        Args:
            text (str): 输入文本实例
            params (dict): 变量替换字典，键为变量名，值为替换值
        """

        # 预处理参数：key 统一规范化
        norm_params = {re.sub(r"\s+", "", k).lower(): str(v) for k, v in params.items()}

        pattern = re.compile(r"\{\{(.*?)\}\}")

        def repl(match):
            raw_key = match.group(1)
            norm_key = re.sub(r"\s+", "", raw_key).lower()
            return norm_params.get(norm_key, match.group(0))

        return pattern.sub(repl, text)

    @staticmethod
    def _process_paragraph(p: "Paragraph", params: dict):
        """
        替换段落中的 {{var}}。

        特点：
        1. 支持变量跨 Run
        2. 不删除 Paragraph
        3. 可以新增 Run
        4. 新 Run 完整继承原 Run 的 XML 样式
        5. style 只覆盖明确指定的样式
        """
        from copy import deepcopy
        from docx.text.paragraph import Paragraph

        if not p.runs:
            return

        # ---------------------------------------------------------
        # 规范化参数
        # ---------------------------------------------------------

        norm_params = {}

        for key, value in params.items():
            norm_key = WordUtils._normalize_var_key(key)

            if isinstance(value, dict) and "value" in value:
                norm_params[norm_key] = {
                    "value": str(value["value"]),
                    "style": value.get("style", {}) or {},
                }
            else:
                norm_params[norm_key] = {
                    "value": str(value),
                    "style": {},
                }

        # ---------------------------------------------------------
        # 收集所有 Run 的文本
        # ---------------------------------------------------------

        runs = list(p.runs)

        full_text = "".join(run.text or "" for run in runs)

        if "{{" not in full_text:
            return

        pattern = re.compile(r"\{\{(.*?)\}\}")

        matches = list(pattern.finditer(full_text))

        if not matches:
            return

        # ---------------------------------------------------------
        # 建立字符位置 -> Run 映射
        # ---------------------------------------------------------

        run_ranges = []

        position = 0

        for run in runs:
            text = run.text or ""

            start = position
            end = position + len(text)

            run_ranges.append((start, end, run))

            position = end

        def get_source_run(position):
            """
            根据变量在 full_text 中的位置，
            找到它原来所在的 Run。
            """
            for start, end, run in run_ranges:
                if start <= position < end:
                    return run

            # 边界情况
            if run_ranges:
                return run_ranges[-1][2]

            return None

        # ---------------------------------------------------------
        # 生成替换后的片段
        # ---------------------------------------------------------

        parts = []

        current_pos = 0

        for match in matches:

            # 变量之前的普通文本
            if match.start() > current_pos:
                parts.append(
                    {
                        "text": full_text[current_pos : match.start()],
                        "source_run": get_source_run(current_pos),
                        "style": None,
                    }
                )

            raw_key = match.group(1)
            norm_key = WordUtils._normalize_var_key(raw_key)

            param = norm_params.get(norm_key)

            # 未找到变量，原样保留
            if param is None:
                parts.append(
                    {
                        "text": match.group(0),
                        "source_run": get_source_run(match.start()),
                        "style": None,
                    }
                )
            else:
                parts.append(
                    {
                        "text": param["value"],
                        "source_run": get_source_run(match.start()),
                        "style": param["style"],
                    }
                )

            current_pos = match.end()

        # 最后的普通文本
        if current_pos < len(full_text):
            parts.append(
                {
                    "text": full_text[current_pos:],
                    "source_run": get_source_run(current_pos),
                    "style": None,
                }
            )

        # ---------------------------------------------------------
        # 先清空原 Run
        #
        # 注意：
        # 不删除 Paragraph。
        # 原来的 Run 也不删除。
        # ---------------------------------------------------------

        for run in runs:
            run.text = ""

        # ---------------------------------------------------------
        # 使用第一个原 Run 作为第一个输出 Run
        #
        # 后面的部分允许新增 Run。
        # ---------------------------------------------------------

        first_run = runs[0]
        used_first_run = False

        for part in parts:

            text = part["text"]

            if text == "":
                continue

            source_run = part["source_run"]

            if source_run is None:
                source_run = first_run

            # 第一个片段直接复用原 Run
            if not used_first_run:
                target_run = first_run
                used_first_run = True

                # 先完整继承 source_run
                if source_run is not first_run:
                    target_run._r.get_or_add_rPr()
                    source_rPr = source_run._r.rPr

                    if source_rPr is not None:
                        target_rPr = target_run._r.get_or_add_rPr()

                        for child in list(target_rPr):
                            target_rPr.remove(child)

                        for child in source_rPr:
                            target_rPr.append(deepcopy(child))

            else:
                # 后续创建新的 Run
                target_run = WordUtils._copy_run(
                    source_run,
                    p,
                )

            # 设置文本
            target_run.text = text

            # 变量的 style 是增量覆盖
            if part["style"]:
                WordUtils._set_run_style(
                    target_run,
                    part["style"],
                )

    @staticmethod
    def _process_textboxes(doc, params):
        """
        处理 Word 文档中的文本框内容。
        """
        from docx.oxml.ns import qn
        from docx.text.paragraph import Paragraph

        body = doc.element.body

        # 查找所有文本框中的段落
        for txbx_content in body.iter(qn("w:txbxContent")):
            for p in txbx_content.iter(qn("w:p")):
                paragraph = Paragraph(p, txbx_content)
                WordUtils._process_paragraph(paragraph, params)

    @staticmethod
    def _process_cell(cell: "_Cell", params: dict):
        """
        处理 Word 文档中的单元格内容。
        """
        for p in cell.paragraphs:
            WordUtils._process_paragraph(p, params)

    @staticmethod
    def _copy_font_name(src_run, target_run):
        """
        复制字体
        """
        target_run.font.name = src_run.font.name

        src_rPr = src_run._element.rPr
        if src_rPr is None:
            return

        rFonts = src_rPr.rFonts
        if rFonts is None:
            return

        from docx.oxml.ns import qn

        east_asia_font = rFonts.get(qn("w:eastAsia"))
        if not east_asia_font:
            return

        target_rPr = target_run._element.get_or_add_rPr()
        target_rFonts = target_rPr.get_or_add_rFonts()
        target_rFonts.set(qn("w:eastAsia"), east_asia_font)

    @staticmethod
    def _normalize_var_key(key):
        """变量名规范化：去除空白并忽略大小写"""
        return re.sub(r"\s+", "", str(key)).lower()

    @staticmethod
    def _copy_run(run, paragraph):
        """
        基于原 Run 的 XML 完整复制一个新的 Run。

        原 Run 中存在的所有格式都会被保留。
        """
        from copy import deepcopy
        from docx.text.run import Run
        from docx.oxml.ns import qn

        new_r = deepcopy(run._r)

        # 删除原 Run 中的文本、Tab、Break 等内容
        for child in list(new_r):
            if child.tag in {
                qn("w:t"),
                qn("w:tab"),
                qn("w:br"),
                qn("w:cr"),
                qn("w:noBreakHyphen"),
                qn("w:softHyphen"),
            }:
                new_r.remove(child)

        paragraph._p.append(new_r)

        return Run(new_r, paragraph)

    @staticmethod
    def _set_run_style(run, style):
        """
        对 Run 的 XML 样式进行增量修改。

        style 中没有指定的属性完全保持原样。
        """
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        if not style:
            return

        rPr = run._r.get_or_add_rPr()

        # =========================================================
        # 字号
        # =========================================================
        if "font_size" in style:
            sz = rPr.find(qn("w:sz"))

            if sz is None:
                sz = OxmlElement("w:sz")
                rPr.append(sz)

            # Word 字号单位为 half-point
            value = int(float(style["font_size"]) * 2)
            sz.set(qn("w:val"), str(value))

            # 东亚字体有时还存在 w:szCs
            szCs = rPr.find(qn("w:szCs"))

            if szCs is None:
                szCs = OxmlElement("w:szCs")
                rPr.append(szCs)

            szCs.set(qn("w:val"), str(value))

        # =========================================================
        # 字体颜色
        # =========================================================
        if "font_color" in style:
            color = str(style["font_color"]).lstrip("#").upper()

            color_element = rPr.find(qn("w:color"))

            if color_element is None:
                color_element = OxmlElement("w:color")
                rPr.append(color_element)

            color_element.set(qn("w:val"), color)

        # =========================================================
        # 字体
        # =========================================================
        if "font_name" in style:
            font_name = str(style["font_name"])

            rFonts = rPr.find(qn("w:rFonts"))

            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                rPr.insert(0, rFonts)

            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                rFonts.set(qn(f"w:{attr}"), font_name)

        # =========================================================
        # 粗体
        # =========================================================
        if "bold" in style:
            element = rPr.find(qn("w:b"))

            if element is None:
                element = OxmlElement("w:b")
                rPr.append(element)

            if style["bold"]:
                element.set(qn("w:val"), "1")
            else:
                element.set(qn("w:val"), "0")

        # =========================================================
        # 斜体
        # =========================================================
        if "italic" in style:
            element = rPr.find(qn("w:i"))

            if element is None:
                element = OxmlElement("w:i")
                rPr.append(element)

            if style["italic"]:
                element.set(qn("w:val"), "1")
            else:
                element.set(qn("w:val"), "0")

        # =========================================================
        # 下划线
        # =========================================================
        if "underline" in style:
            element = rPr.find(qn("w:u"))

            if element is None:
                element = OxmlElement("w:u")
                rPr.append(element)

            if style["underline"]:
                element.set(qn("w:val"), "single")
            else:
                element.set(qn("w:val"), "none")

    # endregion

    @staticmethod
    def get_highlight(run):
        """
        安全获取 run 的高亮颜色
        返回 WD_COLOR_INDEX 枚举或 None
        """
        try:
            from docx.enum.text import WD_COLOR_INDEX

            val = run.font.highlight_color
            # 如果不是合法枚举，返回 None
            if val in WD_COLOR_INDEX.__members__.values():
                return val
        except ValueError:
            # 遇到 'none' 或非法值
            return None
        return None

    @staticmethod
    def merge_docx(file_list, output_path, page_break=True):
        """
        按顺序合并多个 docx 文件

        :param file_list: list[str] docx 文件路径
        :param output_path: 输出文件路径
        :param page_break: 是否在每个文档之间分页，默认 True
        :return: Path
        """
        # 懒加载第三方库
        from docx import Document
        from docxcompose.composer import Composer

        if not file_list:
            raise ValueError("合并文件列表不能为空")

        # 检测文件
        for file in file_list:
            path = Path(file)

            if not path.exists():
                raise FileNotFoundError(f"文件不存在: {file}")

            if path.suffix.lower() != ".docx":
                raise ValueError(f"不是docx文件: {file}")

        # 创建主文档
        master = Document(file_list[0])
        composer = Composer(master)

        # 依次追加
        for file in file_list[1:]:
            if page_break:
                master.add_page_break()

            doc = Document(file)
            composer.append(doc)

        # 保存
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        composer.save(str(output_path))

        return output_path

    @staticmethod
    def get_docx_text_content(docx_path: str) -> list[str]:
        """
        返回docx文档的文本列表
        """
        from docx import Document

        doc = Document(docx_path)

        return [
            paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()
        ]

    @staticmethod
    def append_content(docx_path: str, content_list: list[str]):
        """
        向 docx 末尾追加内容。
        注，需最后一行有内容，纯空格格式不生效

        每一项内容创建一个新段落，并继承文档最后一个段落的格式。

        Args:
            doc_path: docx 文件路径
            content_list: 要写入的内容，例如：
                [
                    "第一行内容",
                    "第二行内容",
                    "第三行内容"
                ]
        """
        from copy import deepcopy
        from docx import Document
        from docx.text.paragraph import Paragraph
        from docx.oxml.ns import qn

        doc = Document(docx_path)

        body = doc._body._element

        # 找 body 下最后一个 w:p
        paragraphs = body.findall(qn("w:p"))

        if not paragraphs:
            raise ValueError("文档中没有 w:p")

        source_p = paragraphs[-1]

        # 保存源段落最后一个 run 的格式
        source_rPr = None

        source_runs = source_p.findall(qn("w:r"))

        if source_runs:
            source_rPr = source_runs[-1].find(qn("w:rPr"))

        for text in content_list:

            # 完整复制最后一个 w:p
            new_p = deepcopy(source_p)

            # 删除复制出来的所有 run
            for element in list(new_p):
                if element.tag == qn("w:r"):
                    new_p.remove(element)

            # 直接插入到上一个 w:p 后面
            source_p.addnext(new_p)

            # 当前新增段落作为下一次的插入基准
            source_p = new_p

            # 创建 Paragraph
            new_paragraph = Paragraph(
                new_p,
                body,
            )

            # 添加文字
            new_run = new_paragraph.add_run(text)

            # 恢复源 Run 格式
            if source_rPr is not None:
                new_run._r.insert(
                    0,
                    deepcopy(source_rPr),
                )

        doc.save(docx_path)

    @staticmethod
    def delete_paragraph(docx_path, paragraph_index):
        """
        删除指定行段落，从1开始
        """

        from docx import Document

        doc = Document(docx_path)

        if paragraph_index < 1 or paragraph_index > len(doc.paragraphs):
            raise IndexError(
                f"段落序号超出范围：{paragraph_index}，"
                f"当前共有 {len(doc.paragraphs)} 个段落"
            )

        paragraph = doc.paragraphs[paragraph_index - 1]

        # 从 XML 中删除
        paragraph._element.getparent().remove(paragraph._element)

        doc.save(docx_path)

    @staticmethod
    def docx_to_pdf(docx_path, pdf_path=None):
        import win32com.client

        docx_path = os.path.abspath(docx_path)

        if pdf_path is None:
            pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

        pdf_path = os.path.abspath(pdf_path)

        wps = win32com.client.DispatchEx("Kwps.Application")
        wps.Visible = False
        wps.DisplayAlerts = False

        try:
            doc = wps.Documents.Open(docx_path)

            # 强制分页
            doc.Repaginate()

            # 2 = wdStatisticPages
            page_count = doc.ComputeStatistics(2)

            # 转 PDF
            doc.ExportAsFixedFormat(
                pdf_path,
                17,
            )

            doc.Close(False)

        finally:
            wps.Quit()

        return pdf_path, page_count

    # region 页眉页码处理

    @staticmethod
    def set_header_text(docx_path, text, output_path=None):
        """
        设置页眉文字
        """

        from docx import Document

        doc = Document(docx_path)

        for section in doc.sections:
            header = section.header

            for paragraph in header.paragraphs:
                # 找到最后一个普通文本 run
                for run in reversed(paragraph.runs):
                    if run.text.strip():
                        run.text = text
                        break

        if output_path is None:
            output_path = docx_path

        doc.save(output_path)

    @staticmethod
    def add_header(docx_path, text, align="center", style=None):
        """
        给 Word 文档添加页眉。
        """
        from docx import Document

        style = style or {}

        align_map = WordUtils._get_align_map()

        if align not in align_map:
            raise ValueError(
                f"不支持的页眉对齐方式: {align}，"
                f"可选值: {', '.join(align_map.keys())}"
            )

        doc = Document(docx_path)

        for section in doc.sections:
            header = section.header

            # 获取第一个段落
            if header.paragraphs:
                paragraph = header.paragraphs[0]
            else:
                paragraph = header.add_paragraph()

            # 清空段落内容
            for child in list(paragraph._p):
                if child.tag.endswith("}pPr"):
                    continue
                paragraph._p.remove(child)

            # 添加页眉文字
            run = paragraph.add_run(str(text))

            # 设置对齐
            paragraph.alignment = align_map[align]

            # 设置样式
            WordUtils._apply_run_style(run, style)

        doc.save(docx_path)

        return docx_path

    @staticmethod
    def add_page_number(
        docx_path,
        position="footer",
        align="center",
        prefix="",
        suffix="",
        style=None,
    ):
        """
        添加自动页码。

        Args:
            docx_path: Word 文件路径
            position: header / footer
            align: left / center / right
            prefix: 页码前缀，默认为空
            suffix: 页码后缀，默认为空
            style: 页码样式
        """
        from docx import Document
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        if position not in ("header", "footer"):
            raise ValueError("position 必须是 'header' 或 'footer'")

        if align not in ("left", "center", "right"):
            raise ValueError("align 必须是 'left'、'center' 或 'right'")

        style = style or {}

        doc = Document(docx_path)

        align_map = WordUtils._get_align_map()

        for section in doc.sections:

            # =====================================================
            # 获取页眉 / 页脚
            # =====================================================

            container = section.header if position == "header" else section.footer

            if container.paragraphs:
                paragraph = container.paragraphs[0]
            else:
                paragraph = container.add_paragraph()

            # =====================================================
            # 页码在页眉
            # =====================================================

            if position == "header":

                # -------------------------------------------------
                # 页眉不能修改 Paragraph 对齐
                # 使用右对齐 Tab Stop 定位页码
                # -------------------------------------------------

                pPr = paragraph._p.get_or_add_pPr()

                tabs = pPr.find(qn("w:tabs"))

                if tabs is None:
                    tabs = OxmlElement("w:tabs")
                    pPr.append(tabs)

                # 页面右边界
                section_width = section.page_width
                right_margin = section.right_margin

                # 计算右侧位置
                right_position = int(
                    (section.page_width - section.left_margin - section.right_margin)
                    / 635
                )

                # 检查是否已经存在右对齐 Tab
                right_tab_exists = False

                for tab in tabs.findall(qn("w:tab")):
                    if tab.get(qn("w:val")) == "right" and tab.get(qn("w:pos")) == str(
                        right_position
                    ):
                        right_tab_exists = True
                        break

                if not right_tab_exists:
                    tab = OxmlElement("w:tab")

                    tab.set(qn("w:val"), "right")
                    tab.set(qn("w:pos"), str(right_position))

                    tabs.append(tab)

                # -------------------------------------------------
                # 添加 Tab
                # -------------------------------------------------

                tab_run = paragraph.add_run()
                tab_run.add_tab()

                # -------------------------------------------------
                # 添加前缀
                # -------------------------------------------------

                if prefix:
                    prefix_run = paragraph.add_run(str(prefix))
                    WordUtils._apply_run_style(prefix_run, style)

                # -------------------------------------------------
                # 添加 PAGE 域
                # -------------------------------------------------

                run = paragraph.add_run()

                fld_begin = OxmlElement("w:fldChar")
                fld_begin.set(qn("w:fldCharType"), "begin")

                instr_text = OxmlElement("w:instrText")
                instr_text.set(qn("xml:space"), "preserve")
                instr_text.text = " PAGE "

                fld_end = OxmlElement("w:fldChar")
                fld_end.set(qn("w:fldCharType"), "end")

                run._r.append(fld_begin)
                run._r.append(instr_text)
                run._r.append(fld_end)

                WordUtils._apply_run_style(run, style)

                # -------------------------------------------------
                # 添加后缀
                # -------------------------------------------------

                if suffix:
                    suffix_run = paragraph.add_run(str(suffix))
                    WordUtils._apply_run_style(suffix_run, style)

            # =====================================================
            # 页码在页脚
            # =====================================================

            else:

                paragraph.alignment = align_map[align]

                # -------------------------------------------------
                # 前缀
                # -------------------------------------------------

                if prefix:
                    prefix_run = paragraph.add_run(str(prefix))
                    WordUtils._apply_run_style(prefix_run, style)

                # -------------------------------------------------
                # PAGE 域
                # -------------------------------------------------

                run = paragraph.add_run()

                fld_begin = OxmlElement("w:fldChar")
                fld_begin.set(qn("w:fldCharType"), "begin")

                instr_text = OxmlElement("w:instrText")
                instr_text.set(qn("xml:space"), "preserve")
                instr_text.text = " PAGE "

                fld_end = OxmlElement("w:fldChar")
                fld_end.set(qn("w:fldCharType"), "end")

                run._r.append(fld_begin)
                run._r.append(instr_text)
                run._r.append(fld_end)

                WordUtils._apply_run_style(run, style)

                # -------------------------------------------------
                # 后缀
                # -------------------------------------------------

                if suffix:
                    suffix_run = paragraph.add_run(str(suffix))
                    WordUtils._apply_run_style(suffix_run, style)

        doc.save(docx_path)

        return docx_path

    @staticmethod
    def _apply_run_style(run, style):
        """
        对 Run 应用增量样式。

        style 中未指定的属性不会修改。

        支持：
            font_name
            font_size
            font_color
            bold
            italic
            underline
            strike
            double_strike
            subscript
            superscript
        """
        if not style:
            return

        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt, RGBColor

        rPr = run._r.get_or_add_rPr()

        # =========================================================
        # 字体
        # =========================================================
        if "font_name" in style:
            font_name = str(style["font_name"])

            rFonts = rPr.find(qn("w:rFonts"))

            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                rPr.insert(0, rFonts)

            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                rFonts.set(qn(f"w:{attr}"), font_name)

        # =========================================================
        # 字号
        # =========================================================
        if "font_size" in style:
            size = int(float(style["font_size"]) * 2)

            sz = rPr.find(qn("w:sz"))

            if sz is None:
                sz = OxmlElement("w:sz")
                rPr.append(sz)

            sz.set(qn("w:val"), str(size))

            # 中文/复杂脚本字号
            szCs = rPr.find(qn("w:szCs"))

            if szCs is None:
                szCs = OxmlElement("w:szCs")
                rPr.append(szCs)

            szCs.set(qn("w:val"), str(size))

        # =========================================================
        # 字体颜色
        # =========================================================
        if "font_color" in style:
            color = str(style["font_color"]).lstrip("#").upper()

            color_element = rPr.find(qn("w:color"))

            if color_element is None:
                color_element = OxmlElement("w:color")
                rPr.append(color_element)

            color_element.set(qn("w:val"), color)

        # =========================================================
        # 粗体
        # =========================================================
        if "bold" in style:
            element = rPr.find(qn("w:b"))

            if element is None:
                element = OxmlElement("w:b")
                rPr.append(element)

            element.set(qn("w:val"), "1" if style["bold"] else "0")

        # =========================================================
        # 斜体
        # =========================================================
        if "italic" in style:
            element = rPr.find(qn("w:i"))

            if element is None:
                element = OxmlElement("w:i")
                rPr.append(element)

            element.set(qn("w:val"), "1" if style["italic"] else "0")

        # =========================================================
        # 下划线
        # =========================================================
        if "underline" in style:
            element = rPr.find(qn("w:u"))

            if element is None:
                element = OxmlElement("w:u")
                rPr.append(element)

            element.set(qn("w:val"), "single" if style["underline"] else "none")

        # =========================================================
        # 删除线
        # =========================================================
        if "strike" in style:
            element = rPr.find(qn("w:strike"))

            if element is None:
                element = OxmlElement("w:strike")
                rPr.append(element)

            element.set(qn("w:val"), "1" if style["strike"] else "0")

        # =========================================================
        # 双删除线
        # =========================================================
        if "double_strike" in style:
            element = rPr.find(qn("w:dstrike"))

            if element is None:
                element = OxmlElement("w:dstrike")
                rPr.append(element)

            element.set(qn("w:val"), "1" if style["double_strike"] else "0")

        # =========================================================
        # 下标
        # =========================================================
        if "subscript" in style:
            run.font.subscript = style["subscript"]

        # =========================================================
        # 上标
        # =========================================================
        if "superscript" in style:
            run.font.superscript = style["superscript"]

    # endregion

    @staticmethod
    def _get_align_map():
        """
        获取对齐方式映射字典。
        Returns:
            dict: 对齐方式映射字典
        """
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
        }

        return align_map
