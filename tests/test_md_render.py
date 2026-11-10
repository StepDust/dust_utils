"""MdToDocx 端到端测试：Mermaid 图表 + LaTeX 公式 + 代码块。"""

import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from dust_utils.file_utils.md_math import MathPreprocessor
from dust_utils.file_utils.md_to_docx import MdToDocx


def test_math_extract():
    """公式预处理：占位符、代码块/行内代码/金额不被误伤。"""
    pre = MathPreprocessor()
    md = (
        "价格 $100 与 $200 不变\n\n"
        "行内 $E=mc^2$ 结束\n\n"
        "```\n$code$\n```\n\n"
        "行内代码 `$x$` 保留\n\n"
        "$$\n\\begin{aligned}\na &= b \\\\\nc &= d\n\\end{aligned}\n$$\n"
    )
    out = pre.extract(md)

    assert "DUSTMATH" in out, "行内公式未替换"
    assert "$100 与 $200" in out, "金额被误判为公式"
    assert "$code$" in out, "围栏代码块内的 $ 被误判"
    assert "`$x$`" in out, "行内代码内的 $ 被误判"
    assert len(pre.items) == 2, f"公式数量应为 2，实际 {len(pre.items)}"

    display = [i for i in pre.items.values() if i.display]
    assert len(display) == 1, "块级公式识别失败"
    assert "aligned" in display[0].tex


def test_convert(tmp_dir):
    md_path = os.path.join(ROOT, "test_config", "md_to_config", "test_render.md")
    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    style_path = os.path.join(ROOT, "test_config", "md_to_config", "style.json")
    with open(style_path, "r", encoding="utf-8") as f:
        styles = json.load(f)["docx_theme"]["default"]

    output = os.path.join(tmp_dir, "test_render.docx")

    # 离线模式下绝不允许发起网络请求
    import requests

    original = (requests.get, requests.post)

    def blocked(*args, **kwargs):
        raise AssertionError(f"离线模式不应该发起网络请求：{args[:1]}")

    requests.get = requests.post = blocked
    try:
        converter = MdToDocx()
        converter.convert(md_text, output, styles)
    finally:
        requests.get, requests.post = original

    assert os.path.exists(output), "文档未生成"

    from docx import Document

    doc = Document(output)
    pictures = len(doc.inline_shapes)
    text = "\n".join(p.text for p in doc.paragraphs)

    print(f"段落数：{len(doc.paragraphs)}，图片数：{pictures}，大小：{os.path.getsize(output)}")
    assert pictures == 10, f"图片数量异常（8 个公式 + 2 张图）：{pictures}"
    assert "def hello" in text, "普通代码块内容丢失"
    assert "你好" in text, "普通代码块中文丢失"
    assert "a^2 + b^2 = c^2" not in text, "行内公式未被渲染"
    assert "aligned" not in text, "块级公式未被渲染"

    # 输出结构，便于人工核对顺序与尺寸
    for index, paragraph in enumerate(doc.paragraphs):
        label = paragraph.style.name
        content = paragraph.text.strip().replace("\n", " ")
        if len(content) > 46:
            content = content[:46] + "..."
        print(f"  {index:02d} [{label}] {content}")


def main():
    import tempfile

    test_math_extract()
    print("[OK] MathPreprocessor")

    with tempfile.TemporaryDirectory() as tmp_dir:
        test_convert(tmp_dir)
        keep = os.path.join(ROOT, "test_config", "md_to_config", "test_render.docx")
        src = os.path.join(tmp_dir, "test_render.docx")
        with open(src, "rb") as f:
            data = f.read()
        with open(keep, "wb") as f:
            f.write(data)
        print(f"[OK] MdToDocx -> {keep}")


if __name__ == "__main__":
    main()
