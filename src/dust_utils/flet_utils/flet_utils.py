import flet as ft
from flet import Colors
import copy
import os
from loguru import logger
import asyncio
import functools
import threading
import platform
import subprocess


class FletUtils:
    def __init__(self, page: ft.Page, spacing=10, font_family="微软雅黑"):
        self.page = page
        self.page.padding = 0
        self.spacing = spacing
        # 设置字体
        self.page.fonts = {
            "Noto Sans SC": "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanssc/NotoSansSC-Regular.ttf",
        }
        self.page.theme = ft.Theme(
            font_family=font_family,
            scrollbar_theme=ft.ScrollbarTheme(
                thickness=7,
                radius=7,
                thumb_color="#808080",  # 可见的滑块颜色
                track_color="transparent",  # 透明轨道
                thumb_visibility=True,
            ),
        )

        # 设置窗口默认居中打开
        self.page.run_task(self.page.window.center)

    # region

    @staticmethod
    def run_in_thread(func):
        """
        装饰器：让函数自动在子线程中运行
        """

        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            thread = threading.Thread(
                target=functools.partial(func, *args, **kwargs), daemon=True
            )
            thread.start()
            return thread

        return wrapper

    def set_theme_radius(self, radius: int = 7):
        """
        动态设置全局主题圆角（不重建整个 Theme，只原地修改）
        """
        if not self.page.theme:
            self.page.theme = ft.Theme()

        rounded = ft.RoundedRectangleBorder(radius=radius)

        # ==================== 辅助函数 ====================
        def _apply_shape(theme_attr, theme_class, style_class=None):
            """辅助函数：统一为按钮类主题应用圆角"""
            if hasattr(self.page.theme, theme_attr) and getattr(
                self.page.theme, theme_attr
            ):
                target = getattr(self.page.theme, theme_attr)
                if hasattr(target, "style") and target.style:
                    target.style.shape = rounded
                else:
                    target.style = ft.ButtonStyle(shape=rounded)
            else:
                # 如果不存在则创建
                if style_class:
                    setattr(
                        self.page.theme,
                        theme_attr,
                        theme_class(style=style_class(shape=rounded)),
                    )
                else:
                    setattr(self.page.theme, theme_attr, theme_class(shape=rounded))

        # ==================== 应用圆角 ====================
        # 按钮类控件
        _apply_shape("button_theme", ft.ButtonTheme, ft.ButtonStyle)
        _apply_shape("filled_button_theme", ft.FilledButtonTheme, ft.ButtonStyle)
        _apply_shape("outlined_button_theme", ft.OutlinedButtonTheme, ft.ButtonStyle)
        _apply_shape("icon_button_theme", ft.IconButtonTheme, ft.ButtonStyle)

        # Card 和 Dialog
        _apply_shape("card_theme", ft.CardTheme)
        _apply_shape("dialog_theme", ft.DialogTheme)

        # 强制刷新页面
        self.page.update()

    # endregion

    # region 拖拽分屏

    def get_horizontal_drag_handle(
        self,
        left_panel,
        drag_width=4,
        panel_min_width=200,
        panel_max_width=800,
        bg_color: Colors | str = ft.Colors.TRANSPARENT,
        hover_color: Colors | str = ft.Colors.BLUE_400,
    ):
        """
        创建一个可拖拽的垂直分割手柄（支持 hover 高亮 + 左右拖拽）
        :param left_panel: 需要调整宽度的左侧面板（Container）
        :return: GestureDetector 控件
        """

        dot_01 = ft.Container(
            width=4,
            height=4,
            border_radius=4,
            bgcolor=ft.Colors.GREY,
        )

        dot_02 = copy.deepcopy(dot_01)
        dot_03 = copy.deepcopy(dot_01)

        # ==================== Hover 处理 ====================
        def on_handle_hover(e: ft.ControlEvent):
            # e.data 在 Container.on_hover 中是布尔值：True=进入，False=离开
            is_hover = bool(e.data) if e.data is not None else False

            # 修改背景色（视觉反馈）
            handle_container.bgcolor = hover_color if is_hover else bg_color
            # 同时让圆点也变色
            dot_01.bgcolor = ft.Colors.WHITE if is_hover else ft.Colors.GREY
            dot_02.bgcolor = ft.Colors.WHITE if is_hover else ft.Colors.GREY
            dot_03.bgcolor = ft.Colors.WHITE if is_hover else ft.Colors.GREY

            handle_container.update()

        # ==================== 拖拽处理 ====================
        def on_pan_update(e: ft.DragUpdateEvent):
            delta = (
                e.local_delta.x if e.local_delta and e.local_delta.x is not None else 0
            )
            if abs(delta) < 0.5:  # 过滤微小抖动
                return

            new_width = max(
                panel_min_width, min(panel_max_width, left_panel.width + delta)
            )
            if abs(left_panel.width - new_width) > 0.5:
                left_panel.width = new_width
                left_panel.update()

        # ==================== 创建手柄 ====================

        handle_container = ft.Container(
            width=drag_width,
            # expand=True,  # 垂直方向撑满高度
            bgcolor=ft.Colors.TRANSPARENT,
            content=ft.Column(  # 用 Column 垂直排列圆点
                controls=[
                    # 第一列：3 个小圆点
                    ft.Column(
                        controls=[
                            dot_01,
                            dot_02,
                            dot_03,
                        ],
                        spacing=4,
                        alignment=ft.MainAxisAlignment.CENTER,
                    ),
                ],
                # spacing=6,  # 两行圆点之间的垂直间距
                alignment=ft.MainAxisAlignment.CENTER,  # 整体垂直居中
            ),
            alignment=ft.alignment.Alignment.CENTER,  # 整个水平居中
            on_hover=on_handle_hover,
        )

        # 只负责拖拽和鼠标样式
        drag_handle = ft.GestureDetector(
            content=handle_container,
            on_pan_update=on_pan_update,
            mouse_cursor=ft.MouseCursor.RESIZE_LEFT_RIGHT,
            drag_interval=5,
        )

        return drag_handle

    # endregion


# region 日志捕获


class LogConsole(ft.Container):

    def __init__(self, page: ft.Page):
        super().__init__()

        self.logs = []
        self.ft_page = page
        self.line_index = 0

        self.list_view = ft.ListView(
            expand=True,
            auto_scroll=True,  # 自动滚动
            spacing=7,
            build_controls_on_demand=True,  # 按需构建
            cache_extent=100,
        )

        self.expand = True
        self.bgcolor = "#1e1e1e"
        self.padding = 10

        self.content = ft.Column(
            controls=[
                self._build_toolbar(),
                self.list_view,
            ],
            expand=True,
        )

        logger.add(
            self._sink,
            # enqueue=True,
        )

    # -------------------------
    # toolbar
    # -------------------------
    def _build_toolbar(self):

        btn_style = ft.ButtonStyle(
            bgcolor={
                ft.ControlState.DEFAULT: "#292929",  # 基础深灰
                ft.ControlState.HOVERED: "#2e2e2e",  # 明显提亮一档
                ft.ControlState.PRESSED: "#404040",  # 明显压暗（接近黑）
            },
            color="#ffffff",
            side=None,
            text_style=ft.TextStyle(
                size=12,
                weight=ft.FontWeight.W_500,
                font_family="微软雅黑",
            ),
        )

        return ft.Row(
            controls=[
                # 左侧标题
                ft.Text(
                    "",
                    color="#ffffff",
                ),
                # 撑开空间（关键）
                ft.Container(expand=True),
                # 右侧按钮组
                ft.Row(
                    spacing=10,
                    controls=[
                        ft.IconButton(
                            tooltip="打开日志",
                            icon_size=16,
                            icon=ft.Icons.FILE_OPEN_OUTLINED,
                            on_click=self.on_open_log,
                            style=btn_style,
                        ),
                        ft.IconButton(
                            tooltip="清空日志",
                            icon_size=16,
                            icon=ft.Icons.DELETE_OUTLINE,
                            on_click=self.on_clear,
                            style=btn_style,
                        ),
                        # ft.Icons.KEYBOARD_DOUBLE_ARROW_UP_ROUNDED
                        ft.IconButton(
                            tooltip="向上翻页",
                            icon_size=16,
                            icon=ft.Icons.KEYBOARD_DOUBLE_ARROW_UP_ROUNDED,
                            on_click=self.on_page_up,
                            style=btn_style,
                        ),
                        ft.IconButton(
                            tooltip="向下翻页",
                            icon_size=16,
                            icon=ft.Icons.KEYBOARD_DOUBLE_ARROW_DOWN_ROUNDED,
                            on_click=self.on_page_down,
                            style=btn_style,
                        ),
                    ],
                ),
            ],
        )

    # -------------------------
    # loguru sink
    # -------------------------
    def _sink(self, message):

        record = message.record
        extra = record.get("extra", {})

        color = extra.get("color", "")

        level = record["level"].name

        color_map = {
            "TRACE": "#8c8c8c",
            "DEBUG": "#00c2ff",
            "INFO": "#d9d9d9",
            "SUCCESS": "#23D18B",
            "WARNING": "#fadb14",
            "ERROR": "#ff4d4f",
            "CRITICAL": "#ff0000",
            "DIVIDER": "#ffffff",
        }

        # level颜色
        level_color = color_map.get(level, "#ffffff")
        if color:
            level_color = color

        # 时间
        time_text = record["time"].strftime("%H:%M:%S")

        # message
        message_text = record["message"]
        self.line_index += 1
        font_size = 12

        rich = ft.Container(
            content=ft.Row(
                vertical_alignment=ft.CrossAxisAlignment.START,
                spacing=0,
                controls=[
                    # 行号
                    ft.Container(
                        width=40,
                        alignment=ft.alignment.Alignment.CENTER_LEFT,
                        content=ft.Text(
                            f"{self.line_index:>4} ",
                            selectable=True,
                            color="#6E7681",
                            size=font_size,
                            font_family="Consolas",
                        ),
                    ),
                    # 时间
                    ft.Container(
                        content=ft.Text(
                            time_text,
                            selectable=True,
                            color="#52c41a",
                            size=font_size,
                            font_family="Consolas",
                        ),
                    ),
                    # 分割线
                    ft.Text(
                        " │ ",
                        color="#666666",
                        size=font_size,
                        font_family="Consolas",
                        selectable=True,
                    ),
                    # LEVEL
                    ft.Container(
                        content=ft.Text(
                            f"{level:<8}",
                            selectable=True,
                            color=level_color,
                            size=font_size,
                            weight=ft.FontWeight.BOLD,
                            font_family="Consolas",
                        ),
                    ),
                    # 分割线
                    ft.Text(
                        " │ ",
                        color="#666666",
                        size=font_size,
                        font_family="Consolas",
                        selectable=True,
                    ),
                    # message
                    ft.Container(
                        expand=True,
                        content=ft.Text(
                            message_text,
                            selectable=True,
                            color=level_color,
                            size=font_size,
                            font_family="微软雅黑",
                        ),
                    ),
                ],
            ),
            on_hover=self.on_row_hover,
            border_radius=2,
        )

        item = ft.GestureDetector(
            content=rich,
            on_double_tap=lambda e: self.ft_page.run_task(
                self.on_copy_msg, text=message_text
            ),
        )

        self.logs.append(item)

        self.list_view.controls.append(item)

        # 最大日志限制
        if len(self.list_view.controls) > 5000:
            self.list_view.controls.pop(0)
            self.logs.pop(0)

        self.list_view.update()

        if self.ft_page:
            self.ft_page.run_task(self._auto_scroll, offset=-1)

    # region 事件函数

    async def _auto_scroll(self, offset: float = -1, delta: float | None = None):
        """
        统一的滚动方法
        :param offset: 绝对像素偏移（用于滚到指定位置）
        :param delta: 相对滚动距离（用于上翻/下翻）
        :param key: 目标控件的 key，优先使用
        """
        await asyncio.sleep(0.05)  # 等待渲染完成
        try:
            if delta is not None:
                await self.list_view.scroll_to(delta=delta, duration=0)
            elif offset is not None:
                # 如果指定了 offset，就直接用（注意 offset=0 是顶部）
                await self.list_view.scroll_to(offset=offset, duration=0)
            else:
                # 兜底：滚动到底部
                await self.list_view.scroll_to(offset=999999, duration=0)
        except Exception:
            pass

    def on_row_hover(self, e):
        row = e.control

        if e.data:
            row.bgcolor = "#2a2a2a"  # 悬浮高亮
        else:
            row.bgcolor = "transparent"  # 恢复默认

        row.update()

    async def on_copy_msg(self, text: str):
        await ft.Clipboard().set(text)
        snack = ft.SnackBar(
            content=ft.Text("日志复制成功", color="white"),
            bgcolor="#16b777",
            behavior=ft.SnackBarBehavior.FLOATING,  # 浮动样式
            width=200,
            duration=800,  # 持续时间
        )

        # 关键：使用 show_snack_bar，它本身就不阻塞
        self.ft_page.show_dialog(snack)

    def on_clear(self, e):
        self.logs.clear()
        self.line_index = 0
        self.list_view.controls.clear()
        self.update()

    def on_open_log(self, e):
        """用系统默认程序打开文件（跨平台）"""
        log_path = logger.get_log_path()

        if not os.path.exists(log_path):
            logger.error(f"文件不存在: {log_path}")
            return

        system = platform.system()
        try:
            if system == "Windows":
                os.startfile(log_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", log_path])
            else:  # Linux 或其它 Unix
                subprocess.run(["xdg-open", log_path])
        except Exception as e:
            logger.error(f"打开文件失败: {e}")

    def on_page_up(self, e):
        """上翻一页（向上滚动一个列表可见高度）"""
        h = self._get_list_view_height()
        delta = -h if h else -150  # 取不到高度时用默认值
        self.ft_page.run_task(self._auto_scroll, delta=delta)

    def on_page_down(self, e):
        """下翻一页（向下滚动一个列表可见高度）"""
        h = self._get_list_view_height()
        delta = h if h else 150
        self.ft_page.run_task(self._auto_scroll, delta=delta)

    def _get_list_view_height(self):
        """估算 ListView 当前可见区域的高度（像素）"""
        toolbar_height = 50  # 与 _build_toolbar 中一致
        extra_padding = 20  # 顶部/底部的边距（根据你设置的 padding 调整）
        window_height = self.ft_page.height or 600  # 后备值
        return max(150, window_height - toolbar_height - extra_padding)

    # endregion


# endregion
