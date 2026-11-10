"""Mermaid 图表与 LaTeX 公式渲染为 PNG 的工具（默认完全离线）。

python-docx 无法直接写入 Mermaid 与公式，因此统一先渲染成 PNG 图片，
再由 :class:`dust_utils.file_utils.md_to_docx.MdToDocx` 以图片方式插入 Word。

默认只使用本地引擎，不会把文档内容发送到任何服务器：

- Mermaid：``mmdc``（本地 mermaid-cli） →
  ``browser``（本机 Edge/Chrome + mermaid.min.js）
- LaTeX  ：``mathjax``（本机 Edge/Chrome + 本地 tex-svg.js，支持完整 LaTeX）
           → ``matplotlib``（离线 MathText，支持常用写法）

``browser`` 与 ``mathjax`` 需要两个 JS 资源文件，首次使用前执行一次
:func:`setup_assets` 下载到 ``~/.dust_utils/assets``，之后永久离线可用。

确实需要在线服务时，显式传 ``allow_online=True`` 才会追加 ``kroki`` /
``mermaid.ink`` / ``codegogs`` 作为兜底引擎。

渲染结果按内容哈希缓存到磁盘，重复转换同一篇文档时不会重复渲染。
"""

import base64
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import zlib
from dataclasses import dataclass, field
from pathlib import Path

from loguru import logger

__all__ = [
    "RenderResult",
    "DiagramRenderError",
    "DiagramRenderer",
    "render_mermaid",
    "render_latex",
    "get_renderer",
    "setup_assets",
    "download_mermaid_js",
    "download_mathjax_js",
    "asset_dir",
]

# ---------- 在线服务地址（仅在 allow_online=True 时使用） ----------
KROKI_MERMAID_URL = "https://kroki.io/mermaid/png"
MERMAID_INK_URL = "https://mermaid.ink/img/"
CODECOGS_URL = "https://latex.codecogs.com/png.latex"

# ---------- 本地资源 ----------
MERMAID_JS_URL = "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"
MATHJAX_JS_URL = "https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js"

DEFAULT_CACHE_DIR = os.path.join(tempfile.gettempdir(), "dust_utils", "diagram_cache")

# Windows 常见的 Chromium 内核浏览器位置
BROWSER_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"
_USER_AGENT = "Mozilla/5.0 (compatible; dust_utils)"

# 页面模板：渲染成功时截图后裁边，失败时把背景标红以便调用方识别
_PAGE = """<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body {{ margin: 0; padding: 16px; background: #fff; display: inline-block; }}
  #c {{ display: inline-block; line-height: 1; }}
  mjx-container {{ margin: 0 !important; padding: 0 !important; }}
</style>
</head><body><div id="c"></div>
{script}
</body></html>"""

_MERMAID_PAGE_SCRIPT = """<script src="{js}"></script>
<script>
mermaid.initialize({{ startOnLoad: false, theme: "{theme}", securityLevel: "loose",
                      flowchart: {{ useMaxWidth: false }},
                      sequence: {{ useMaxWidth: false }} }});
mermaid.render("dustGraph", {code}).then(function (r) {{
  document.getElementById("c").innerHTML = r.svg;
}}).catch(function (e) {{
  document.body.style.background = "#ff0000";
  document.getElementById("c").textContent = String(e);
}});
</script>"""

_MATHJAX_PAGE_SCRIPT = """<div style="font-size:{font_size}px;line-height:1">{tex}</div>
<script>window.MathJax = {{
  tex: {{ displayMath: [["$$", "$$"]], inlineMath: [["$", "$"]] }},
  svg: {{ fontCache: "none" }},
  startup: {{ typeset: true }}
}};</script>
<script src="{js}"></script>
<script>
MathJax.startup.promise.then(function () {{ return MathJax.typesetPromise(); }})
  .then(function () {{
    if (document.querySelector("[data-mjx-error]")) {{
      document.body.style.background = "#ff0000";
    }}
  }}).catch(function (e) {{
    document.body.style.background = "#ff0000";
    document.getElementById("c").textContent = String(e);
  }});
</script>"""


class DiagramRenderError(RuntimeError):
    """所有渲染引擎都失败时抛出，errors 中保存每个引擎的失败原因。"""

    def __init__(self, kind: str, errors: dict):
        self.kind = kind
        self.errors = errors
        detail = "; ".join(f"{k}: {v}" for k, v in errors.items()) or "未知错误"
        super().__init__(f"{kind} 渲染失败，尝试的引擎均未成功 -> {detail}")


@dataclass
class RenderResult:
    """一次渲染的结果。"""

    path: str = ""
    width_px: int = 0
    height_px: int = 0
    dpi: float = 96.0
    engine: str = ""
    meta: dict = field(default_factory=dict)

    @property
    def width_inch(self) -> float:
        """图片逻辑宽度（英寸）。"""
        return self.width_px / (self.dpi or 96.0)

    @property
    def height_inch(self) -> float:
        """图片逻辑高度（英寸）。"""
        return self.height_px / (self.dpi or 96.0)


# region 通用工具函数


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:32]


def _image_size(data: bytes):
    """读取图片像素尺寸，优先 Pillow，退化为解析 PNG 头。"""
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as img:
            return img.size
    except Exception:
        pass

    if data[:8] == _PNG_MAGIC:
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")

    raise ValueError("无法识别图片尺寸（既不是 PNG 也不是 Pillow 支持的格式）")


def _ensure_png(data: bytes) -> bytes:
    """保证返回 PNG 字节；非 PNG（如 jpeg）在有 Pillow 时转码。"""
    if data[:8] == _PNG_MAGIC:
        return data
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as img:
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format="PNG")
            return buf.getvalue()
    except Exception as exc:
        raise ValueError(f"返回内容不是有效图片：{exc}") from exc


def _b64url(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def _norm_color(color: str) -> str:
    """把 white / #fff / FFFFFF 统一成 FFFFFF（6 位十六进制，无 #）。"""
    named = {"white": "FFFFFF", "black": "000000", "transparent": "FFFFFF"}
    text = (color or "white").strip().lower()
    if text in named:
        return named[text]
    text = text.lstrip("#")
    if len(text) == 3:
        text = "".join(ch * 2 for ch in text)
    if len(text) != 6 or any(ch not in "0123456789abcdef" for ch in text):
        return "FFFFFF"
    return text.upper()


def mermaid_ink_url(
    code: str, theme: str = "default", background: str = "FFFFFF"
) -> str:
    """生成 mermaid.ink 的 PNG 链接（pako = zlib 压缩 + base64url）。"""
    payload = json.dumps({"code": code, "mermaid": {"theme": theme}}).encode("utf-8")
    token = _b64url(zlib.compress(payload, 9))
    return f"{MERMAID_INK_URL}pako:{token}?type=png&bgColor={background}"


def _split_top_level(text: str, sep: str):
    """按顶层分隔符切分，忽略花括号内与转义字符中的分隔符。"""
    parts, buf, depth, i = [], [], 0, 0
    while i < len(text):
        if depth == 0 and text.startswith(sep, i):
            parts.append("".join(buf))
            buf = []
            i += len(sep)
            continue
        ch = text[i]
        if ch == "\\" and i + 1 < len(text) and not text.startswith(sep, i):
            buf.append(text[i : i + 2])
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth = max(depth - 1, 0)
        buf.append(ch)
        i += 1
    parts.append("".join(buf))
    return parts


_MATH_ENV_RE = re.compile(
    r"\\begin\{(aligned|align\*?|alignedat|gather\*?|gathered|split|eqnarray\*?"
    r"|array|cases|matrix|pmatrix|bmatrix|Bmatrix|vmatrix|Vmatrix)\}"
    r"(?:\{[^}]*\})?(.*?)\\end\{\1\}",
    re.S,
)


def split_math_rows(tex: str):
    """
    把 aligned / cases / matrix 等多行环境拆成二维文本网格。

    matplotlib 的 MathText 不支持 \\begin{...} 环境，拆开后逐格渲染再拼合，
    可以在没有 MathJax 的情况下也能输出可读的多行公式。

    Returns:
        list[list[str]] | None: 每行由若干单元格组成；不是多行结构时返回 None
    """
    match = _MATH_ENV_RE.search(tex)
    if match:
        body = match.group(2)
    elif "\\\\" in tex:
        body = tex
    else:
        return None

    rows = []
    for raw_row in _split_top_level(body, "\\\\"):
        cells = [
            re.sub(r"\\tag\{[^}]*\}", "", cell).strip()
            for cell in _split_top_level(raw_row, "&")
        ]
        if any(cells):
            rows.append(cells)
    return rows or None


# endregion


# region 本地资源与浏览器


def asset_dir() -> Path:
    """本地资源目录：``~/.dust_utils/assets``，可用环境变量 DUST_ASSET_DIR 覆盖。"""
    override = os.environ.get("DUST_ASSET_DIR")
    path = (
        Path(override)
        if override
        else Path(os.path.expanduser("~")) / ".dust_utils" / "assets"
    )
    path.mkdir(parents=True, exist_ok=True)
    return path


def _resolve_asset(explicit, env_name: str, filename: str):
    """按 显式参数 -> 环境变量 -> 默认目录 的顺序查找本地资源文件。"""
    for candidate in (explicit, os.environ.get(env_name)):
        if candidate and os.path.exists(candidate):
            return candidate
    default = asset_dir() / filename
    return str(default) if default.exists() else None


def _download(url: str, dest: str, timeout: float = 120, force: bool = False) -> str:
    if os.path.exists(dest) and not force:
        logger.debug(f"资源已存在，跳过下载：{dest}")
        return dest

    import requests

    logger.info(f"正在下载离线渲染资源：{url}")
    resp = requests.get(url, timeout=timeout, headers={"User-Agent": _USER_AGENT})
    resp.raise_for_status()

    os.makedirs(os.path.dirname(os.path.abspath(dest)), exist_ok=True)
    tmp = f"{dest}.tmp"
    with open(tmp, "wb") as f:
        f.write(resp.content)
    os.replace(tmp, dest)

    logger.success(f"已保存：{dest}（{len(resp.content) / 1024 / 1024:.1f} MB）")
    return dest


def download_mermaid_js(
    dest: str = None,
    url: str = MERMAID_JS_URL,
    timeout: float = 120,
    force: bool = False,
) -> str:
    """下载 mermaid.min.js（仅需联网一次，之后 Mermaid 渲染可完全离线）。"""
    dest = dest or str(asset_dir() / "mermaid.min.js")
    return _download(url, dest, timeout, force)


def download_mathjax_js(
    dest: str = None,
    url: str = MATHJAX_JS_URL,
    timeout: float = 120,
    force: bool = False,
) -> str:
    """下载 MathJax 的 tex-svg.js（仅需联网一次，之后公式渲染可完全离线）。"""
    dest = dest or str(asset_dir() / "tex-svg.js")
    return _download(url, dest, timeout, force)


def setup_assets(dest: str = None, timeout: float = 120, force: bool = False) -> dict:
    """
    一次性下载离线渲染所需的两个 JS 资源，之后渲染不再需要网络。

    Args:
        dest: 保存目录，默认 ``~/.dust_utils/assets``
        timeout: 单个文件的下载超时（秒）
        force: 为 True 时即使文件存在也重新下载

    Returns:
        dict: {"mermaid": path, "mathjax": path}
    """
    folder = Path(dest) if dest else asset_dir()
    folder.mkdir(parents=True, exist_ok=True)
    return {
        "mermaid": download_mermaid_js(
            str(folder / "mermaid.min.js"), timeout=timeout, force=force
        ),
        "mathjax": download_mathjax_js(
            str(folder / "tex-svg.js"), timeout=timeout, force=force
        ),
    }


def find_browser(explicit: str = None) -> str:
    """查找本机 Chromium 内核浏览器（优先 Edge，其次 Chrome）。"""
    for candidate in (explicit, os.environ.get("DUST_BROWSER"), *BROWSER_CANDIDATES):
        if candidate and os.path.exists(candidate):
            return candidate

    for name in ("msedge", "chrome", "chromium", "google-chrome"):
        path = shutil.which(name)
        if path:
            return path

    raise FileNotFoundError(
        "未找到 Edge/Chrome，可通过 browser_path 参数或 DUST_BROWSER 环境变量指定"
    )


# endregion


class DiagramRenderer:
    """Mermaid / LaTeX 渲染器（默认全离线，带内存 + 磁盘缓存）。"""

    def __init__(
        self,
        cache_dir: str = None,
        timeout: float = 60,
        use_cache: bool = True,
        allow_online: bool = False,
        mermaid_js: str = None,
        mathjax_js: str = None,
        browser_path: str = None,
        mermaid_scale: float = 2.0,
        math_scale: float = 2.0,
        web_scale: float = 1.0,
        browser_size=(1800, 1300),
    ):
        """
        Args:
            cache_dir: 缓存目录，默认系统临时目录下的 dust_utils/diagram_cache
            timeout: 单个渲染任务（含浏览器 / 网络）的超时时间（秒）
            use_cache: 是否启用缓存
            allow_online: 是否允许使用在线服务兜底，默认 False（完全离线）
            mermaid_js: 本地 mermaid.min.js 路径，默认找 ~/.dust_utils/assets
            mathjax_js: 本地 tex-svg.js 路径，默认找 ~/.dust_utils/assets
            browser_path: 浏览器可执行文件路径，默认自动查找 Edge / Chrome
            mermaid_scale: 图表截图的清晰度倍数，dpi = 96 * scale
            math_scale: 公式截图的清晰度倍数
            web_scale: 在线服务位图相对 CSS 像素的倍数（仅在线引擎使用）
            browser_size: 无头浏览器窗口大小 (宽, 高)，内容更大时会自动放大重试
        """
        self.cache_dir = (
            cache_dir or os.environ.get("DUST_DIAGRAM_CACHE") or DEFAULT_CACHE_DIR
        )
        self.timeout = timeout
        self.use_cache = use_cache
        self.allow_online = allow_online
        self.mermaid_js = mermaid_js
        self.mathjax_js = mathjax_js
        self.browser_path = browser_path
        self.mermaid_scale = mermaid_scale
        self.math_scale = math_scale
        self.web_scale = web_scale
        self.browser_size = browser_size
        self._memory = {}

    # region 缓存

    def _cache_path(self, kind: str, key: str) -> str:
        folder = os.path.join(self.cache_dir, kind)
        os.makedirs(folder, exist_ok=True)
        return os.path.join(folder, f"{key}.png")

    def _read_cache(self, kind: str, key: str):
        if not self.use_cache:
            return None
        path = self._cache_path(kind, key)
        if not os.path.exists(path):
            return None
        try:
            with open(path, "rb") as f:
                size = _image_size(f.read())
        except Exception:
            return None
        return RenderResult(
            path=path, width_px=size[0], height_px=size[1], engine="cache"
        )

    def _write_cache(self, kind: str, key: str, data: bytes) -> str:
        path = self._cache_path(kind, key)
        tmp = f"{path}.tmp"
        with open(tmp, "wb") as f:
            f.write(data)
        os.replace(tmp, path)
        return path

    # endregion

    # region 通用渲染流程

    def _render(self, kind: str, key: str, engines, builder, dpi: float):
        """依次尝试 engines，第一个成功的引擎结果会被缓存并返回。"""
        if self.use_cache and key in self._memory:
            return self._memory[key]

        cached = self._read_cache(kind, key)
        if cached is not None:
            cached.dpi = dpi
            self._memory[key] = cached
            return cached

        errors = {}
        result = None
        for engine in engines:
            try:
                data, engine_dpi = builder(engine)
                data = _ensure_png(data)
                width, height = _image_size(data)
                path = self._write_cache(kind, key, data)
                result = RenderResult(
                    path=path,
                    width_px=width,
                    height_px=height,
                    dpi=engine_dpi or dpi,
                    engine=engine,
                )
                logger.debug(f"{kind} 渲染成功：engine={engine}, {width}x{height}px")
                break
            except Exception as exc:
                errors[engine] = f"{type(exc).__name__}: {exc}"
                logger.debug(f"{kind} 渲染失败：engine={engine} -> {exc}")

        if result is None:
            raise DiagramRenderError(kind, errors)

        self._memory[key] = result
        return result

    # endregion

    # region 浏览器截图

    def _screenshot(
        self, script: str, scale: float, tag: str, pad_scale: float = 8
    ) -> bytes:
        """用本机浏览器的无头模式渲染页面并截图，自动裁掉四周空白。"""
        exe = find_browser(self.browser_path)

        with tempfile.TemporaryDirectory() as tmp:
            page_path = os.path.join(tmp, "page.html")
            with open(page_path, "w", encoding="utf-8") as f:
                f.write(_PAGE.format(script=script))
            url = Path(page_path).as_uri()

            width, height = self.browser_size
            last_error = ""
            for factor in (1.0, 1.6):
                size = (int(width * factor), int(height * factor))
                shot = os.path.join(tmp, f"shot{factor}.png")

                ok, error = self._run_browser(exe, url, size, scale, shot)
                if not ok:
                    last_error = error
                    continue

                with open(shot, "rb") as f:
                    png, clipped = self._crop(f.read(), scale, tag, pad_scale)

                if not clipped:
                    return png
                logger.debug(f"{tag} 内容超出截图窗口，放大窗口重试")

            raise RuntimeError(last_error or f"{tag} 截图失败")

    def _run_browser(self, exe: str, url: str, size, scale: float, shot: str):
        errors = []
        for headless in ("--headless=new", "--headless"):
            if os.path.exists(shot):
                os.remove(shot)

            cmd = [
                exe,
                headless,
                "--disable-gpu",
                "--hide-scrollbars",
                "--allow-file-access-from-files",
                "--disable-extensions",
                f"--virtual-time-budget={int(min(self.timeout, 60) * 1000)}",
                f"--force-device-scale-factor={scale}",
                f"--window-size={size[0]},{size[1]}",
                f"--screenshot={shot}",
                url,
            ]
            try:
                proc = subprocess.run(
                    cmd, capture_output=True, timeout=max(self.timeout * 2, 120)
                )
            except subprocess.TimeoutExpired:
                errors.append(f"{headless}: 超时")
                continue

            if os.path.exists(shot) and os.path.getsize(shot) > 0:
                return True, ""

            stderr = (proc.stderr or b"").decode("utf-8", "ignore")[:300]
            errors.append(f"{headless}: {stderr}")

        return False, " | ".join(errors)

    def _crop(self, raw: bytes, scale: float, tag: str, pad_scale: float = 8):
        """裁掉截图四周空白；页面被标记为错误（红色背景）时抛异常。"""
        from PIL import Image, ImageChops

        img = Image.open(io.BytesIO(raw)).convert("RGB")
        width, height = img.size
        if width < 8 or height < 8:
            raise RuntimeError(f"{tag} 截图为空")

        corner = img.getpixel((width - 2, height - 2))
        if corner[0] > 150 and corner[1] < 120 and corner[2] < 120:
            raise RuntimeError(f"{tag} 渲染失败（页面已标记错误，请检查语法）")

        white = Image.new("RGB", img.size, (255, 255, 255))
        box = ImageChops.difference(img, white).getbbox()
        if not box:
            raise RuntimeError(f"{tag} 渲染结果为空")

        clipped = box[2] >= width - 2 or box[3] >= height - 2

        pad = int(pad_scale * scale)
        box = (
            max(box[0] - pad, 0),
            max(box[1] - pad, 0),
            min(box[2] + pad, width),
            min(box[3] + pad, height),
        )

        buf = io.BytesIO()
        img.crop(box).save(buf, format="PNG")
        return buf.getvalue(), clipped

    # endregion

    # region Mermaid

    def render_mermaid(
        self,
        code: str,
        engine: str = "auto",
        theme: str = "default",
        background: str = "white",
        scale: float = None,
    ) -> RenderResult:
        """
        渲染 Mermaid 图表为 PNG。

        Args:
            code: Mermaid 代码（不含 ``` 围栏）
            engine: auto / mmdc / browser（离线）；
                kroki / mermaid_ink 需 allow_online=True
            theme: Mermaid 主题，default / neutral / dark / forest
            background: 背景色，仅本地 mmdc 与在线引擎使用
            scale: 清晰度倍数，默认取实例的 mermaid_scale

        Returns:
            RenderResult: 渲染结果，``path`` 为 PNG 缓存文件
        """
        code = (code or "").strip()
        if not code:
            raise ValueError("Mermaid 代码为空")

        scale = float(scale or self.mermaid_scale)
        bg = _norm_color(background)
        engines = self._mermaid_engines(engine)
        key = _sha256(f"mermaid|{theme}|{bg}|{scale}|{code}")

        result = self._render(
            "mermaid",
            key,
            engines,
            lambda name: self._mermaid_builder(name, code, theme, bg, scale),
            96 * scale,
        )
        result.meta.setdefault("code", code)
        return result

    def _mermaid_engines(self, engine: str):
        if engine and engine != "auto":
            return [engine]

        names = []
        if shutil.which("mmdc"):
            names.append("mmdc")
        names.append("browser")
        if self.allow_online:
            names.extend(["kroki", "mermaid_ink"])
        return names

    def _mermaid_builder(
        self, engine: str, code: str, theme: str, bg: str, scale: float
    ):
        if engine == "browser":
            js = _resolve_asset(self.mermaid_js, "DUST_MERMAID_JS", "mermaid.min.js")
            if not js and self.allow_online:
                js = download_mermaid_js()
            if not js:
                raise FileNotFoundError(
                    "未找到 mermaid.min.js，请先执行 setup_assets()"
                    " 或 download_mermaid_js()（仅需联网一次）"
                )
            script = _MERMAID_PAGE_SCRIPT.format(
                js=Path(js).as_uri(), theme=theme, code=json.dumps(code)
            )
            return self._screenshot(script, scale, "Mermaid"), 96 * scale

        if engine == "mmdc":
            return self._mmdc(code, theme, bg, scale), 96 * scale

        if engine == "kroki":
            return self._http_post(KROKI_MERMAID_URL, code), 96 * self.web_scale

        if engine in ("mermaid_ink", "mermaid.ink"):
            return self._http_get(mermaid_ink_url(code, theme, bg)), 96 * self.web_scale

        raise ValueError(f"不支持的 Mermaid 引擎：{engine}")

    def _mmdc(self, code: str, theme: str, bg: str, scale: float) -> bytes:
        """调用本地 mermaid-cli（需要 npm i -g @mermaid-js/mermaid-cli）。"""
        exe = shutil.which("mmdc")
        if not exe:
            raise FileNotFoundError(
                "未找到 mmdc，可执行 npm i -g @mermaid-js/mermaid-cli"
            )

        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "in.mmd")
            out = os.path.join(tmp, "out.png")
            with open(src, "w", encoding="utf-8") as f:
                f.write(code)

            args = [
                exe,
                "-i",
                src,
                "-o",
                out,
                "-b",
                f"#{bg}",
                "-t",
                theme,
                "-s",
                str(scale),
            ]
            if os.name == "nt" and exe.lower().endswith((".cmd", ".bat")):
                args = [os.environ.get("COMSPEC", "cmd.exe"), "/c", *args]

            proc = subprocess.run(
                args, capture_output=True, timeout=max(self.timeout * 3, 120)
            )
            if proc.returncode != 0 or not os.path.exists(out):
                stderr = (proc.stderr or b"").decode("utf-8", "ignore")[:500]
                raise RuntimeError(f"mmdc 退出码 {proc.returncode}: {stderr}")

            with open(out, "rb") as f:
                return f.read()

    # endregion

    # region LaTeX

    def render_latex(
        self,
        tex: str,
        engine: str = "auto",
        display: bool = False,
        font_size: float = 12,
        dpi: float = 240,
        color: str = "#000000",
    ) -> RenderResult:
        """
        渲染 LaTeX 公式为 PNG。

        Args:
            tex: LaTeX 公式源码（不含 $ / \\[ \\] 定界符）
            engine: auto / mathjax / matplotlib（离线），codecogs（需 allow_online）
            display: 是否为独立成行的公式
            font_size: 字号（pt）
            dpi: matplotlib 引擎的渲染清晰度
            color: 字体颜色（仅 matplotlib 生效）

        Returns:
            RenderResult: 渲染结果
        """
        tex = (tex or "").strip()
        if not tex:
            raise ValueError("LaTeX 公式为空")

        engines = self._latex_engines(engine)
        key = _sha256(f"latex|{display}|{font_size}|{dpi}|{color}|{tex}")

        result = self._render(
            "latex",
            key,
            engines,
            lambda name: self._latex_builder(name, tex, display, font_size, dpi, color),
            dpi,
        )
        result.meta.setdefault("tex", tex)
        return result

    def _latex_engines(self, engine: str):
        if engine and engine != "auto":
            return [engine]

        names = ["mathjax", "matplotlib"]
        if self.allow_online:
            names.append("codecogs")
        return names

    def _latex_builder(
        self,
        engine: str,
        tex: str,
        display: bool,
        font_size: float,
        dpi: float,
        color: str,
    ):
        if re.search(r"[\u4e00-\u9fff]", tex):
            logger.warning("公式中包含中文，渲染可能缺字，建议把中文移到公式外面")

        if engine == "mathjax":
            return self._mathjax(tex, display, font_size), 96 * self.math_scale

        if engine == "matplotlib":
            return self._latex_matplotlib(tex, font_size, dpi, color), dpi

        if engine == "codecogs":
            return self._latex_codecogs(tex, dpi), dpi

        raise ValueError(f"不支持的 LaTeX 引擎：{engine}")

    def _mathjax(self, tex: str, display: bool, font_size: float) -> bytes:
        """本地 MathJax 渲染，支持 aligned / cases / matrix 等完整 LaTeX。"""
        js = _resolve_asset(self.mathjax_js, "DUST_MATHJAX_JS", "tex-svg.js")
        if not js and self.allow_online:
            js = download_mathjax_js()
        if not js:
            raise FileNotFoundError(
                "未找到 tex-svg.js，请先执行 setup_assets()"
                " 或 download_mathjax_js()（仅需联网一次）"
            )

        # 公式字号（pt）换算成 CSS 像素，截图时再按 math_scale 放大
        size = round(font_size * 96 / 72, 2)
        body = f"$${tex}$$" if display else f"${tex}$"
        script = _MATHJAX_PAGE_SCRIPT.format(
            js=Path(js).as_uri(), tex=body, font_size=size
        )
        return self._screenshot(script, self.math_scale, "LaTeX", pad_scale=2)

    def _latex_matplotlib(
        self, tex: str, font_size: float, dpi: float, color: str
    ) -> bytes:
        """离线 MathText 渲染；多行环境拆行后逐格渲染再拼合。"""
        from matplotlib import mathtext
        from matplotlib.font_manager import FontProperties

        prop = FontProperties(size=font_size)

        def render_one(expr: str):
            buf = io.BytesIO()
            depth = mathtext.math_to_image(
                expr if expr.startswith("$") else f"${expr}$",
                buf,
                prop=prop,
                dpi=dpi,
                format="png",
                color=color,
            )
            return buf.getvalue(), depth

        try:
            data, _ = render_one(tex)
            return data
        except Exception as exc:
            rows = split_math_rows(tex)
            if not rows:
                raise
            logger.debug(f"公式含多行环境，改为逐行渲染：{exc}")
            return self._stack_rows(rows, font_size, dpi, color)

    def _stack_rows(self, rows, font_size: float, dpi: float, color: str) -> bytes:
        """把多行公式逐格渲染后按基线对齐拼成一张 PNG。"""
        from matplotlib import mathtext
        from matplotlib.font_manager import FontProperties
        from PIL import Image

        prop = FontProperties(size=font_size)
        unit = font_size * dpi / 72.0
        gap = max(int(unit * 0.6), 2)
        row_gap = max(int(unit * 0.35), 2)

        # 1️⃣ 逐格渲染，记录每格的图片与基线深度
        grid = []
        for cells in rows:
            row = []
            for cell in cells:
                if not cell.strip():
                    row.append(None)
                    continue
                buf = io.BytesIO()
                depth = mathtext.math_to_image(
                    f"${cell}$",
                    buf,
                    prop=prop,
                    dpi=dpi,
                    format="png",
                    color=color,
                )
                buf.seek(0)
                row.append((Image.open(buf).convert("RGB"), depth / 72.0 * dpi))
            grid.append(row)

        columns = max(len(row) for row in grid)
        col_widths = [0] * columns
        for row in grid:
            for index, cell in enumerate(row):
                if cell:
                    col_widths[index] = max(col_widths[index], cell[0].width)

        # 2️⃣ 每行的上/下沿（用于基线对齐）
        row_metrics = []
        for row in grid:
            cells = [cell for cell in row if cell]
            asc = max((img.height - depth for img, depth in cells), default=0)
            desc = max((depth for _img, depth in cells), default=0)
            row_metrics.append((int(asc), int(desc)))

        total_width = sum(col_widths) + gap * max(columns - 1, 0)
        total_height = sum(asc + desc for asc, desc in row_metrics) + row_gap * (
            len(grid) - 1
        )
        canvas = Image.new(
            "RGB", (max(total_width, 1), max(total_height, 1)), (255, 255, 255)
        )

        # 3️⃣ 按基线拼合
        top = 0
        for row, (asc, desc) in zip(grid, row_metrics, strict=False):
            left = 0
            for index, cell in enumerate(row):
                if cell:
                    img, depth = cell
                    y = top + asc - (img.height - depth)
                    canvas.paste(img, (int(left), int(max(y, 0))))
                left += col_widths[index] + gap
            top += asc + desc + row_gap

        buf = io.BytesIO()
        canvas.save(buf, format="PNG")
        return buf.getvalue()

    def _latex_codecogs(self, tex: str, dpi: float) -> bytes:
        """在线渲染（需 allow_online=True），支持完整 LaTeX。"""
        from urllib.parse import quote

        expr = f"\\dpi{{{int(dpi)}}}\\bg{{white}}{tex}"
        return self._http_get(f"{CODECOGS_URL}?{quote(expr, safe='')}")

    # endregion

    # region 网络请求（仅 allow_online=True 时使用）

    def _http_get(self, url: str) -> bytes:
        import requests

        resp = requests.get(
            url, timeout=self.timeout, headers={"User-Agent": _USER_AGENT}
        )
        resp.raise_for_status()
        return resp.content

    def _http_post(self, url: str, body: str) -> bytes:
        import requests

        resp = requests.post(
            url,
            data=body.encode("utf-8"),
            timeout=self.timeout,
            headers={"Content-Type": "text/plain", "User-Agent": _USER_AGENT},
        )
        resp.raise_for_status()
        return resp.content

    # endregion


_default_renderer = None


def get_renderer(**kwargs) -> DiagramRenderer:
    """获取全局默认渲染器（首次调用时创建）。"""
    global _default_renderer
    if _default_renderer is None:
        _default_renderer = DiagramRenderer(**kwargs)
    return _default_renderer


def render_mermaid(code: str, **kwargs) -> RenderResult:
    """便捷函数：用全局渲染器把 Mermaid 渲染为 PNG。"""
    return get_renderer().render_mermaid(code, **kwargs)


def render_latex(tex: str, **kwargs) -> RenderResult:
    """便捷函数：用全局渲染器把 LaTeX 公式渲染为 PNG。"""
    return get_renderer().render_latex(tex, **kwargs)


if __name__ == "__main__":
    # 用法：python -m dust_utils.file_utils.diagram_render [--force]
    import sys

    setup_assets(force="--force" in sys.argv)
