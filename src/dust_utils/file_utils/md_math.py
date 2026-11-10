"""Markdown 文档中的 LaTeX 公式识别工具。

为了让 ``MdToDocx`` 不依赖额外的 markdown-it 插件，这里在 Markdown 文本层面
把公式抽取成唯一的占位符，解析完 token 后再由占位符还原成图片。

- 块级公式：``$$ ... $$``、``\\[ ... \\]``，支持跨行与多行环境（如 aligned）
- 行内公式：``$ ... $``、``\\( ... \\)``

为避免误伤，抽取时会跳过围栏代码块、行内代码，以及转义后的 ``\\$``；
同时要求行内公式首尾不是空白、收尾 ``$`` 后面不紧跟数字（避免把金额当成公式）。
"""

import re
import secrets
from dataclasses import dataclass

__all__ = ["MathItem", "MathPreprocessor"]


@dataclass
class MathItem:
    """一段公式。"""

    tex: str
    display: bool = False


# 围栏代码块（``` 或 ~~~）
_FENCE_RE = re.compile(r"^ {0,3}(`{3,}|~{3,})")
# 行内代码：`code`
_CODE_SPAN_RE = re.compile(r"(`+)(?:(?!\1).)+?\1")
# 行内公式 $...$
_INLINE_DOLLAR_RE = re.compile(r"(?<![\\$])\$(?=\S)([^\n$]+?)(?<=\S)\$(?![\d$])")
# 行内公式 \(...\)
_INLINE_PAREN_RE = re.compile(r"(?<!\\)\\\((.+?)(?<!\\)\\\)")


class MathPreprocessor:
    """把 Markdown 里的公式替换为占位符，并保留占位符到公式的映射。"""

    def __init__(
        self,
        allow_inline: bool = True,
        allow_display: bool = True,
        allow_digits: bool = True,
    ):
        """
        Args:
            allow_inline: 是否识别行内公式 ``$...$`` / ``\\(...\\)``
            allow_display: 是否识别块级公式 ``$$...$$`` / ``\\[...\\]``
            allow_digits: 是否允许公式以数字开头（关闭后 ``$5$`` 不再算公式）
        """
        self.allow_inline = allow_inline
        self.allow_display = allow_display
        self.allow_digits = allow_digits

        self.nonce = secrets.token_hex(4).upper()
        self.items: dict[str, MathItem] = {}
        self._counter = 0

        self.pattern = re.compile(rf"DUSTMATH{self.nonce}_(\d+)_")

    # region 对外接口

    def extract(self, md_text: str) -> str:
        """把公式替换成占位符，返回新的 Markdown 文本。"""
        if not self.allow_inline and not self.allow_display:
            return md_text

        lines = md_text.split("\n")
        out = []
        fence = None
        i = 0

        while i < len(lines):
            line = lines[i]

            match = _FENCE_RE.match(line)
            if match:
                mark = match.group(1)
                if fence is None:
                    fence = mark[0] * 3
                elif line.strip().startswith(fence):
                    fence = None
                out.append(line)
                i += 1
                continue

            if fence is not None:
                out.append(line)
                i += 1
                continue

            # 块级公式优先
            if self.allow_display:
                block = self._block_math(lines, i)
                if block is not None:
                    tex, consumed = block
                    placeholder = self._add(tex, display=True)
                    # 前后补空行，保证占位符独占一个段落
                    if out and out[-1].strip():
                        out.append("")
                    out.append(placeholder)
                    out.append("")
                    i += consumed
                    continue

            if self.allow_inline:
                line = self._inline_math(line)

            out.append(line)
            i += 1

        return "\n".join(out)

    def split(self, text: str):
        """把文本按占位符切分成 [str | MathItem] 片段列表。"""
        if not text or not self.items:
            return [text]

        parts = []
        pos = 0
        for match in self.pattern.finditer(text):
            if match.start() > pos:
                parts.append(text[pos : match.start()])
            item = self.items.get(match.group(0))
            if item is not None:
                parts.append(item)
            else:
                parts.append(match.group(0))
            pos = match.end()
        if pos < len(text):
            parts.append(text[pos:])
        return parts

    def display_item(self, text: str):
        """整段文本就是一个块级公式占位符时返回该公式，否则返回 None。"""
        if not text:
            return None
        stripped = text.strip()
        match = self.pattern.fullmatch(stripped)
        if not match:
            return None
        item = self.items.get(stripped)
        if item is not None and item.display:
            return item
        return None

    # endregion

    # region 内部实现

    def _add(self, tex: str, display: bool):
        self._counter += 1
        placeholder = f"DUSTMATH{self.nonce}_{self._counter}_"
        self.items[placeholder] = MathItem(tex=tex.strip(), display=display)
        return placeholder

    def _block_math(self, lines, start):
        """尝试解析从 start 行开始的块级公式，返回 (tex, 占用行数) 或 None。"""
        text = lines[start].strip()

        if text.startswith("$$"):
            opener, closer = "$$", "$$"
        elif text.startswith("\\["):
            opener, closer = "\\[", "\\]"
        else:
            return None

        # 单行写法：$$x=1$$ 或 \[x=1\]
        if len(text) > len(opener) + len(closer) and text.endswith(closer):
            tex = text[len(opener) : -len(closer)].strip()
            return (tex, 1) if tex else None

        # 多行写法
        body = [text[len(opener) :]]
        j = start + 1
        while j < len(lines):
            current = lines[j].strip()
            if current.endswith(closer):
                body.append(current[: len(current) - len(closer)])
                tex = "\n".join(body).strip()
                return (tex, j - start + 1) if tex else None
            body.append(lines[j])
            j += 1

        return None

    def _inline_math(self, line: str) -> str:
        """抽取行内公式，行内代码会被暂时屏蔽以免误伤。"""
        if "$" not in line and "\\(" not in line:
            return line

        codes = []

        def mask_code(match):
            codes.append(match.group(0))
            return f"\x00CODE{len(codes) - 1}\x00"

        line = _CODE_SPAN_RE.sub(mask_code, line)

        def replace(match):
            tex = match.group(1).strip()
            if not tex:
                return match.group(0)
            if not self.allow_digits and tex[0].isdigit():
                return match.group(0)
            return self._add(tex, display=False)

        line = _INLINE_DOLLAR_RE.sub(replace, line)
        line = _INLINE_PAREN_RE.sub(replace, line)

        for index, code in enumerate(codes):
            line = line.replace(f"\x00CODE{index}\x00", code)

        return line

    # endregion
