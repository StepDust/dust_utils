# 个人公共工具库 dust_utils

## 虚拟环境推荐使用方式（强烈建议）
每个项目应使用独立的虚拟环境，以避免依赖冲突和版本问题。
- 创建虚拟环境
    ```bash
    uv venv --python (Get-Command python).Source --clear
    ```
- 初始化环境
   ```bash
   uv init
   ```
- 激活虚拟环境
    - Windows
        ```bash
        .venv\Scripts\activate
        ```
- 退出虚拟环境
    ```bash
    deactivate
    ```

## 安装 dust_utils 库
- git安装
    - [core]删除后表示安装基础工具，可改为指定模块的名称
    - 模块名称在**pyproject.toml**下的**project.optional-dependencies**
    ```bash
    uv add "dust_utils[core] @ git+https://github.com/StepDust/dust_utils.git"
    ```
- 本地安装
    - "../dust_utils[core]"是本地路径
    - [core]删除后表示安装基础工具，可改为指定模块的名称
    - 模块名称在**pyproject.toml**下的**project.optional-dependencies**
    ```bash
    uv add --editable "../dust_utils[core]"
    ```
- dust_utils项目下调试模块
    ```bash
    uv pip install -e ".[core]"
    ```

## 输出解释器路径
```bash
python -c "import sys; print(sys.executable)"
```
## 输出python版本&位数
```bash
python -c "import sys; print(sys.version)"
```

## Markdown 转 Word（支持 Mermaid 与 LaTeX，默认全部离线）
`MdToDocx` 会先把 Markdown 里的 Mermaid 图表与 LaTeX 公式渲染成 PNG，再以图片形式插入 Word。
**默认只使用本地引擎，文档内容不会发送到任何服务器。**

- 依赖安装（`[md]` 之外再补 `[diagram]`）
    ```bash
    uv add "dust_utils[md,diagram]"
    ```
- 首次准备（只需联网一次，把 JS 资源下载到 `~/.dust_utils/assets`）
    ```bash
    python -m dust_utils.file_utils.diagram_render
    ```
    - Mermaid：本机 Edge/Chrome 无头渲染 + 本地 `mermaid.min.js`
    - 公式：本机 Edge/Chrome 无头渲染 + 本地 MathJax `tex-svg.js`（支持完整 LaTeX）
    - 已装 Node 的也可以 `npm i -g @mermaid-js/mermaid-cli`，走 `mmdc` 引擎
- Markdown 写法
    - 行内公式：`$E=mc^2$` 或 `\(E=mc^2\)`
    - 块级公式：`$$ ... $$` 或 `\[ ... \]`，支持跨行与 `aligned` 等环境
    - 图表：```` ```mermaid ```` 代码块（也支持时序图、类图、甘特图等）
    - 普通代码块按等宽字体逐行写入，不再丢内容
- 使用示例
    ```python
    from dust_utils.file_utils import MdToDocx

    MdToDocx().convert(
        md_text,
        "out.docx",
        styles,
        render={
            "mermaid_engine": "auto",  # auto / mmdc / browser
            "latex_engine": "auto",  # auto / mathjax / matplotlib
            "allow_online": False,  # 保持 False 即完全离线
            "mermaid_width": None,  # 插入宽度（英寸），None 用图表自然尺寸
            "cache_dir": None,  # 渲染缓存目录，默认系统临时目录
        },
    )
    ```
- 引擎与降级顺序
    | 类型 | 引擎顺序 | 说明 |
    | --- | --- | --- |
    | Mermaid | `mmdc` → `browser` | 都离线；browser 用本机 Edge/Chrome 截图 |
    | 公式 | `mathjax` → `matplotlib` | mathjax 完整 LaTeX；matplotlib 兜底并自动拆行 |
    | 在线（可选） | `kroki` / `mermaid_ink` / `codecogs` | 仅 `allow_online=True` 时才会使用 |
- 注意事项
    - 渲染结果按内容哈希缓存，重复转换同一文档不会重复渲染
    - 渲染失败自动降级：图表写入代码文本、公式写入原文，不会丢内容
    - 公式里尽量避免中文，中文请放到公式外面（缺字形）
    - 金额 `$100`、行内代码 `` `$x$` ``、代码块里的 `$` 不会被误判为公式
    - 也可单独调用：`from dust_utils.file_utils import render_mermaid, render_latex`

## 注意事项
- 修改 dust_utils 源码后，所有项目自动使用最新代码（得益于 editable 模式）。
- 只有 pyproject.toml 的依赖发生变化时，才需在相关项目中重新运行安装命令。
- 建议按需导入模块，例如：
```python
from dust_utils import setup_logger
from dust_utils.ai_utils.ai_chat import AIChat
```
这样可以避免加载不必要的第三方库。


## 常用模块，仅作记录
- **pipreqs**
   - 用于自动生成 requirements.txt 文件，包含项目所有依赖（包括直接和间接依赖）。
   - 安装：`pip install pipreqs`
   - 使用：在项目根目录下运行 `pipreqs . --encoding=utf-8 --ignore=__pycache__,venv --force`
   - 使用utf8编码，忽略 __pycache__ 和 venv 目录，强制覆盖已有的 requirements.txt 文件。
- **pyinstaller**
   - 用于将 Python 脚本打包成独立的可执行文件，方便在无 Python 环境的机器上运行。
   - 安装：`pip install pyinstaller`
   - 使用：在项目根目录下运行 `pyinstaller --onefile your_script.py`（将生成 dist 目录，包含可执行文件）
- **requests**
   - 用于发送 HTTP 请求，处理响应等。
   - 安装：`pip install requests`
   - 使用：在 Python 脚本中导入 `import requests` 即可开始使用。
- **colorama**
   - 用于在终端中输出彩色文本，增强可读性。
   - 安装：`pip install colorama`
   - 使用：在 Python 脚本中导入 `import colorama` 即可开始使用。
- **wcwidth**
   - 用于处理终端中显示的宽字符宽度，避免显示异常。
   - 安装：`pip install wcwidth`
   - 使用：在 Python 脚本中导入 `import wcwidth` 即可开始使用。
- **wxPython**
   - 用于创建图形用户界面（GUI）应用程序。
   - 安装：`pip install wxPython`
   - 使用：在 Python 脚本中导入 `import wx` 即可开始使用。
- **playwright**
   - 用于自动化测试和浏览器操作。
   - 安装：`pip install playwright`
   - 使用：在 Python 脚本中导入 `import playwright` 即可开始使用。

## spec移除打包依赖，仅作记录
- pywin32
- Pillow
- mermaid-py
- esprima
- openai
- wxPython
- pymysql
- playwright



# 常用uv命令

### 安装 pyproject.toml 中的依赖
```
uv sync
```