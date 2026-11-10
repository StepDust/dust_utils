"""diagram_render 测试：离线渲染 Mermaid 与 LaTeX 为 PNG。"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from dust_utils.file_utils.diagram_render import DiagramRenderer, split_math_rows


def test_split_rows():
    """多行公式环境应能拆成网格。"""
    rows = split_math_rows(r"\begin{aligned} a^2 &= b^2 \\ c &= d \end{aligned}")
    assert rows == [["a^2", "= b^2"], ["c", "= d"]], rows

    rows = split_math_rows(r"\begin{cases} x>0 & 正 \\ x<0 & 负 \end{cases}")
    assert len(rows) == 2, rows

    assert split_math_rows(r"E=mc^2") is None
    print("[OK] split_math_rows")


def test_mermaid(renderer):
    """Mermaid 走本地浏览器渲染，不联网。"""
    result = renderer.render_mermaid("flowchart TD\n A[开始] --> B{已登录?}\n B -->|是| C[首页]")
    assert result.engine in ("browser", "mmdc"), result.engine
    assert result.width_px > 50 and result.height_px > 50
    assert os.path.exists(result.path)
    print(f"[OK] mermaid engine={result.engine} {result.width_px}x{result.height_px}")


def test_latex(renderer):
    """公式走本地 MathJax，支持 aligned / cases。"""
    text = renderer.render_latex("E = mc^2", font_size=11)
    assert text.engine in ("mathjax", "matplotlib"), text.engine
    assert 0.1 < text.height_inch < 0.5, text.height_inch

    rows = renderer.render_latex(
        r"\begin{aligned} a^2+b^2 &= c^2 \\ \int_0^1 x^n dx &= \frac{1}{n+1} \end{aligned}",
        display=True,
        font_size=11,
    )
    assert rows.height_inch > text.height_inch, "多行公式应比单行公式高"
    print(f"[OK] latex engine={text.engine} inline={round(text.height_inch, 3)}in "
          f"aligned={round(rows.height_inch, 2)}in")


def test_matplotlib_fallback(renderer):
    """没有 MathJax 时，matplotlib 也要能渲染多行公式。"""
    result = renderer.render_latex(
        r"\begin{aligned} a &= b \\ c &= d \end{aligned}",
        display=True,
        engine="matplotlib",
        font_size=11,
    )
    assert result.engine == "matplotlib"
    assert result.height_inch > 0.2, result.height_inch
    print(f"[OK] matplotlib rows {result.width_px}x{result.height_px}")


def test_bad_mermaid(renderer):
    """语法错误的图表要报错，而不是静默插入空白图。"""
    try:
        renderer.render_mermaid("flowchart TD\n A -->", engine="browser")
    except Exception as exc:
        print(f"[OK] bad mermaid -> {type(exc).__name__}")
        return
    raise AssertionError("错误语法未被识别")


def test_cache(renderer):
    """同样的内容第二次应命中缓存。"""
    first = renderer.render_latex("a+b=c", font_size=11)
    second = renderer.render_latex("a+b=c", font_size=11)
    assert first.path == second.path
    print("[OK] cache hit")


def main():
    test_split_rows()

    with tempfile.TemporaryDirectory() as cache_dir:
        renderer = DiagramRenderer(cache_dir=cache_dir, timeout=60)
        assert renderer.allow_online is False, "默认必须离线"

        test_mermaid(renderer)
        test_latex(renderer)
        test_matplotlib_fallback(renderer)
        test_bad_mermaid(renderer)
        test_cache(renderer)


if __name__ == "__main__":
    main()
