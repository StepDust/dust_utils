# 个人公共工具库 dust_utils

## 虚拟环境推荐使用方式（强烈建议）
每个项目应使用独立的虚拟环境，以避免依赖冲突和版本问题。
- 创建虚拟环境
    ```bash
    uv venv
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
    - 模块名称在**pyproject.toml**下的**project.
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

## 注意事项
- 修改 dust_utils 源码后，所有项目自动使用最新代码（得益于 editable 模式）。
- 只有 pyproject.toml 的依赖发生变化时，才需在相关项目中重新运行安装命令。
- 建议按需导入模块，例如：
```python
from dust_utils.logger_setup import setup_logger
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