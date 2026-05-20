import wx
import re
import os
import threading
import functools

from .mini_alert import MiniAlert
import random
import json
import logging

# 配置日志
from loguru import logger


class WxUtils:
    def __init__(
        self,
        wx_frame,
        config=None,
        row_height=30,
        text_width=250,
        label_width=120,
        btn_width=60,
        gap=20,
    ):

        self.wx_frame = wx_frame
        self.config = config
        self.is_test = False
        # 初始化布局参数
        self.row_height = row_height
        self.text_width = text_width
        self.label_width = label_width
        self.btn_width = btn_width
        self.gap = gap
        # 初始化文本控件列表
        self.text_ctrls = []
        self.btn_ctrls = []
        self.choice_ctrls = []
        self.multicheck_ctrls = []

        # 设置光标
        self.normal_cursor = wx.Cursor(wx.CURSOR_HAND)
        self.disabled_cursor = wx.Cursor(wx.CURSOR_NO_ENTRY)

    # region 控件创建

    def create_folder_ctrls(
        self,
        sizer,
        rows,
        parent=None,
        row_height=None,
        text_width=None,
        label_width=None,
        btn_width=None,
        gap=None,
    ):
        """
        创建多个文件夹选择控件。

        :param panel: 父面板
        :param main_sizer: 主布局器
        :param rows: 行配置列表，每个元素是一个字典，包含 'config_key', 'title' 等
        :param row_height: 行高
        :param text_width: 文本框宽度
        :param label_width: 标签宽度
        :param btn_width: 按钮宽度
        :param btn_text: 按钮文本
        :param btn_event: 按钮点击事件处理函数
        :param gap: 元素间距
        """

        # 初始化布局参数
        if row_height is None:
            row_height = self.row_height
        if text_width is None:
            text_width = self.text_width
        if label_width is None:
            label_width = self.label_width
        if btn_width is None:
            btn_width = self.btn_width
        if gap is None:
            gap = self.gap

        if parent is None:
            parent = sizer.GetContainingWindow()  # 获取 sizer 所在面板

        for row in rows:
            config_key = row.get("config_key", "").strip()
            title = row.get("title", "")
            text_ctrl, btn_ctrl = self.create_text_ctrl(
                sizer=sizer,
                row=row,
                parent=parent,
                btn_text="打开",
                btn_event=None,
                row_height=row_height,
                text_width=text_width,
                label_width=label_width,
                btn_width=btn_width,
                gap=gap,
            )

            text_ctrl.Bind(
                wx.EVT_LEFT_DCLICK,
                lambda event, ctrl=text_ctrl, title=f"选择{title}", config_key=config_key: self.on_select_folder(
                    event=event, text_ctrl=ctrl, title=title, config_key=config_key
                ),
            )

            text_ctrl.SetToolTip("双击选择文件夹")

            btn_ctrl.Bind(
                wx.EVT_BUTTON,
                lambda event, ctrl=text_ctrl: self.on_open_folder(
                    event=event, text_ctrl=ctrl
                ),
            )

    def create_file_ctrls(
        self,
        sizer,
        rows,
        parent=None,
        row_height=None,
        text_width=None,
        label_width=None,
        btn_width=None,
        btn_text="",
        btn_event=None,
        gap=None,
    ):
        """
        创建多个文件选择控件。

        :param panel: 父面板
        :param main_sizer: 主布局器
        :param rows: 行配置列表，每个元素是一个字典，包含 'config_key', 'title' 等
        :param row_height: 行高
        :param text_width: 文本框宽度
        :param label_width: 标签宽度
        :param btn_width: 按钮宽度
        :param btn_text: 按钮文本
        :param btn_event: 按钮点击事件处理函数
        :param gap: 元素间距
        """

        # 初始化布局参数
        if row_height is None:
            row_height = self.row_height
        if text_width is None:
            text_width = self.text_width
        if label_width is None:
            label_width = self.label_width
        if btn_width is None:
            btn_width = self.btn_width
        if gap is None:
            gap = self.gap

        if parent is None:
            parent = sizer.GetContainingWindow()  # 获取 sizer 所在面板

        for row in rows:
            config_key = row.get("config_key", "").strip()
            title = row.get("title", "")
            text_ctrl, browse_btn = self.create_text_ctrl(
                sizer=sizer,
                row=row,
                btn_text=btn_text,
                btn_event=btn_event,
                parent=parent,
                row_height=row_height,
                text_width=text_width,
                label_width=label_width,
                btn_width=btn_width,
                gap=gap,
            )

            text_ctrl.Bind(
                wx.EVT_LEFT_DCLICK,
                lambda event, ctrl=text_ctrl, title=f"选择{title}", config_key=config_key, suffixs=row.get(
                    "suffixs", []
                ): self.on_select_file(
                    event=event,
                    text_ctrl=ctrl,
                    title=title,
                    config_key=config_key,
                    suffixs=suffixs,
                ),
            )
            text_ctrl.SetToolTip("双击选择文件")

    def create_text_ctrl(
        self,
        sizer,
        row,
        parent=None,
        btn_text="",
        btn_event=None,
        row_height=None,
        text_width=None,
        label_width=None,
        btn_width=None,
        gap=None,
    ):
        """
        创建一个标签+文本框+按钮的行布局。

        :param panel: 父面板
        :param main_sizer: 主布局器
        :param row: 行配置字典，包含 'config_key', 'title' 等
        :param btn_text: 按钮文本
        :param btn_event: 按钮点击事件处理函数
        :param row_height: 行高
        :param text_width: 文本框宽度
        :param label_width: 标签宽度
        :param btn_width: 按钮宽度
        :param gap: 元素间距
        :return: 文本框和按钮的元组
        """

        # 初始化布局参数
        # 初始化参数
        if row_height is None:
            row_height = self.row_height
        if text_width is None:
            text_width = self.text_width
        if label_width is None:
            label_width = self.label_width
        if btn_width is None:
            btn_width = self.btn_width
        if gap is None:
            gap = self.gap
        if parent is None:
            parent = sizer.GetContainingWindow()  # 获取主面板

        name = row.get("name", "").strip()
        config_key = row.get("config_key", "").strip()
        title = row.get("title", "")
        if not name:
            name = config_key

        if not config_key and not name:
            wx.MessageBox(
                f"渲染文本框控件时【{title}】的config_key或name不能为空", "配置错误"
            )
            return

        # 创建行容器 Panel 并设置背景色
        row_panel = wx.Panel(parent, size=(-1, row_height))
        row_panel.SetBackgroundColour(self.get_test_color(parent))

        # 行水平布局
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_panel.SetSizer(row_sizer)

        # 标签
        label = wx.StaticText(row_panel, label=title, style=wx.ALIGN_RIGHT)
        label.SetMinSize((label_width, -1))
        row_sizer.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, int(gap / 2))

        # 文本框
        text_ctrl = wx.TextCtrl(
            row_panel,
            size=(text_width, row_height),
            style=wx.TE_MULTILINE,
        )
        # 如果有配置文件且 config_key 不为空，从配置文件获取默认值
        if config_key.strip() and self.config:
            text_ctrl.SetValue(self.config.Read(config_key, ""))
        text_ctrl.Bind(
            wx.EVT_TEXT, lambda event: self.on_text_change(event, text_ctrl, config_key)
        )
        row_sizer.Add(text_ctrl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, gap)
        self.text_ctrls.append({"ctrl": text_ctrl, "name": name})

        # 按钮
        btn = None
        if btn_text:
            btn = wx.Button(row_panel, label=btn_text, size=(btn_width, row_height))
            btn.SetCursor(self.normal_cursor)
            if btn_event:
                btn.Bind(wx.EVT_BUTTON, btn_event)
            row_sizer.Add(btn, 0, wx.ALIGN_CENTER_VERTICAL)
            self.btn_ctrls.append({"ctrl": btn, "name": name})

        # 添加行到主布局
        sizer.Add(row_panel, 0, wx.TOP | wx.EXPAND, gap)

        return text_ctrl, btn

    def create_choice_ctrl(
        self,
        sizer,
        row,
        parent=None,
        btn_text="",
        btn_event=None,
        row_height=None,
        text_width=None,
        label_width=None,
        btn_width=None,
        gap=None,
        choice_event=None,
    ):
        """
        创建一个标签+下拉框+按钮的行布局。

        :param parent: 父面板
        :param sizer: 主布局器
        :param row: 行配置字典，包含 'config_key', 'title', 'options' 等
        :param btn_text: 按钮文本
        :param btn_event: 按钮点击事件处理函数
        :param row_height: 行高
        :param text_width: 下拉框宽度
        :param label_width: 标签宽度
        :param btn_width: 按钮宽度
        :param gap: 元素间距
        :param choice_event: 自定义选择变更事件处理函数
        :return: 下拉框和按钮的元组
        """

        if row_height is None:
            row_height = self.row_height
        if text_width is None:
            text_width = self.text_width
        if label_width is None:
            label_width = self.label_width
        if btn_width is None:
            btn_width = self.btn_width
        if gap is None:
            gap = self.gap
        if parent is None:
            parent = sizer.GetContainingWindow()

        name = row.get("name", "").strip()
        config_key = row.get("config_key", "").strip()
        title = row.get("title", "")
        options = row.get("options", []) or []
        if not name:
            name = config_key

        if not config_key and not name:
            wx.MessageBox(
                f"渲染下拉框控件时【{title}】的config_key或name不能为空", "配置错误"
            )
            return

        row_panel = wx.Panel(parent, size=(-1, row_height))
        row_panel.SetMinSize((-1, row_height))
        row_panel.SetBackgroundColour(self.get_test_color(parent))

        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_panel.SetSizer(row_sizer)

        label = wx.StaticText(row_panel, label=title, style=wx.ALIGN_RIGHT)
        label.SetMinSize((label_width, -1))
        row_sizer.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, int(gap / 2))

        choice_box = wx.Panel(row_panel, size=(-1, row_height))
        choice_box.SetMinSize((-1, row_height))
        vsizer = wx.BoxSizer(wx.VERTICAL)
        choice_ctrl = wx.ComboBox(
            choice_box,
            choices=[str(x) for x in options],
            style=wx.CB_READONLY,
        )
        choice_ctrl.SetMinSize((text_width, -1))
        vsizer.AddStretchSpacer(1)
        vsizer.Add(choice_ctrl, 0, wx.ALIGN_CENTER_HORIZONTAL)
        vsizer.AddStretchSpacer(1)
        choice_box.SetSizer(vsizer)

        if config_key.strip() and self.config:
            value = self.config.Read(config_key, "")
            idx = -1
            if value:
                try:
                    idx = options.index(value)
                except ValueError:
                    idx = -1
            if idx >= 0:
                choice_ctrl.SetSelection(idx)
            elif options:
                choice_ctrl.SetSelection(0)
        elif options:
            choice_ctrl.SetSelection(0)

        def _choice_handler(event):
            self.on_choice_change(event, choice_ctrl, config_key)

            if choice_event:
                choice_event(event)

        choice_ctrl.Bind(wx.EVT_COMBOBOX, _choice_handler)
        row_sizer.Add(choice_box, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, gap)
        self.choice_ctrls.append({"ctrl": choice_ctrl, "name": name})

        btn = None
        if btn_text:
            btn = wx.Button(row_panel, label=btn_text, size=(btn_width, row_height))
            btn.SetCursor(self.normal_cursor)
            if btn_event:
                btn.Bind(wx.EVT_BUTTON, btn_event)
            row_sizer.Add(btn, 0, wx.ALIGN_CENTER_VERTICAL)
            self.btn_ctrls.append({"ctrl": btn, "name": name})

        sizer.Add(row_panel, 0, wx.TOP | wx.EXPAND, gap)

        return choice_ctrl, btn

    def create_multicheck_ctrl(
        self,
        sizer,
        row,
        parent=None,
        btn_text="",
        btn_event=None,
        row_height=None,
        text_width=None,
        label_width=None,
        btn_width=None,
        gap=None,
        check_event=None,
    ):
        """
        创建一个标签 + 多选列表框 + （可选）按钮 的行布局。

        保存格式：JSON 数组，例如 ["日志记录", "暗色主题"]

        :param sizer: 主布局器
        :param row: 行配置字典，包含 'config_key', 'title', 'options', 'name' 等
        :param parent: 父面板（默认从 sizer 获取）
        :param btn_text: 右侧按钮文字（可选）
        :param btn_event: 按钮点击回调（可选）
        :param row_height: 行高度（多选建议较高）
        :param text_width: 多选框宽度
        :param label_width: 标签宽度
        :param btn_width: 按钮宽度
        :param gap: 元素间距
        :param check_event: 勾选变化时的额外回调（可选）
        :return: (wx.CheckListBox, wx.Button or None)
        """

        # 参数默认值
        if row_height is None:
            row_height = (
                int(self.row_height * 3.2) if hasattr(self, "row_height") else 120
            )
        if text_width is None:
            text_width = self.text_width if hasattr(self, "text_width") else 220
        if label_width is None:
            label_width = self.label_width if hasattr(self, "label_width") else 120
        if btn_width is None:
            btn_width = self.btn_width if hasattr(self, "btn_width") else 80
        if gap is None:
            gap = self.gap if hasattr(self, "gap") else 8
        if parent is None:
            parent = sizer.GetContainingWindow()

        # 提取 row 配置
        name = row.get("name", "").strip()
        config_key = row.get("config_key", "").strip()
        title = row.get("title", "")
        options = row.get("options", []) or []

        if not name:
            name = config_key

        if not config_key and not name:
            wx.MessageBox(
                f"渲染多选框时【{title}】的 config_key 或 name 不能为空",
                "配置错误",
                wx.OK | wx.ICON_ERROR,
            )
            return None, None

        # 创建行容器
        row_panel = wx.Panel(parent, size=(-1, row_height))
        row_panel.SetMinSize((-1, row_height))
        if hasattr(self, "get_test_color"):
            row_panel.SetBackgroundColour(self.get_test_color(parent))

        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_panel.SetSizer(row_sizer)

        # 标签
        label = wx.StaticText(row_panel, label=title, style=wx.ALIGN_RIGHT)
        label.SetMinSize((label_width, -1))
        row_sizer.Add(label, 0, wx.ALIGN_TOP | wx.RIGHT | wx.TOP, int(gap / 2))

        # 多选框容器 + CheckListBox
        check_container = wx.Panel(row_panel, size=(text_width, row_height))
        check_container.SetMinSize((text_width, row_height))

        vsizer = wx.BoxSizer(wx.VERTICAL)
        checklist = wx.CheckListBox(
            check_container,
            choices=[str(x) for x in options],
            style=wx.LB_NEEDED_SB,  # 显示滚动条
        )

        vsizer.AddStretchSpacer(1)
        vsizer.Add(checklist, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 4)
        vsizer.AddStretchSpacer(1)
        check_container.SetSizer(vsizer)

        row_sizer.Add(check_container, 0, wx.ALIGN_TOP | wx.RIGHT, gap)

        # ─────────────── 恢复已保存的勾选状态 ───────────────
        checked_indices = []
        if config_key and hasattr(self, "config") and self.config:
            saved_json = self.config.Read(config_key, "[]")
            try:
                saved_list = json.loads(saved_json)
                if not isinstance(saved_list, list):
                    saved_list = []
            except (json.JSONDecodeError, TypeError):
                saved_list = []

            # 匹配选项（字符串比较，更稳健）
            for i, opt in enumerate(options):
                if str(opt) in [str(item) for item in saved_list]:
                    checked_indices.append(i)

        # 设置初始勾选
        for idx in checked_indices:
            checklist.Check(idx)

        # ─────────────── 勾选变化时保存 ───────────────
        def _on_checklist_changed(event):
            # 方法1：使用列表推导式（推荐，简洁高效）
            checked_items = [
                checklist.GetString(i)
                for i in range(checklist.GetCount())
                if checklist.IsChecked(i)
            ]

            # 方法2：使用循环（更显式，容易加调试或额外逻辑）
            # checked_items = []
            # for i in range(checklist.GetCount()):
            #     if checklist.IsChecked(i):
            #         checked_items.append(checklist.GetString(i))

            # 转为 JSON 字符串保存（ensure_ascii=False 防止中文乱码）
            value_json = json.dumps(checked_items, ensure_ascii=False)

            # 保存到配置
            if hasattr(self, "config") and self.config and config_key:
                self.config.Write(config_key, value_json)
                self.config.Flush()  # 立即写入磁盘

            # 可选：打印调试信息（开发时用，正式可删除）
            # print(f"保存 {config_key}: {checked_items}")

            # 如果你有通用处理函数
            if hasattr(self, "on_checklist_change"):
                self.on_checklist_change(
                    event, checklist, config_key, checked_items=checked_items
                )

            # 如果传入了自定义回调
            if check_event:
                check_event(event)

        # ─────────────── 点击选项任意位置都 toggle 勾选 ───────────────
        def _on_left_down(event):
            x, y = event.GetPosition()

            item = checklist.HitTest((x, y))  # ← 只接收一个返回值

            if item != wx.NOT_FOUND:
                checked = checklist.IsChecked(item)
                checklist.Check(item, not checked)

                # 手动触发事件（保持保存逻辑正常执行）
                fake_event = wx.CommandEvent(wx.wxEVT_CHECKLISTBOX, checklist.GetId())
                fake_event.SetInt(item)
                fake_event.SetEventObject(checklist)
                checklist.GetEventHandler().ProcessEvent(fake_event)

                # 可选：高亮该行
                checklist.SetSelection(item)

        checklist.Bind(wx.EVT_LEFT_DOWN, _on_left_down)

        checklist.Bind(wx.EVT_CHECKLISTBOX, _on_checklist_changed)

        # 记录控件（方便后续查找/统一管理）
        self.multicheck_ctrls.append({"ctrl": checklist, "name": name})

        # 可选右侧按钮
        btn = None
        if btn_text:
            btn = wx.Button(row_panel, label=btn_text, size=(btn_width, -1))
            if hasattr(self, "normal_cursor"):
                btn.SetCursor(self.normal_cursor)
            if btn_event:
                btn.Bind(wx.EVT_BUTTON, btn_event)
            row_sizer.Add(btn, 0, wx.ALIGN_TOP)

            if not hasattr(self, "btn_ctrls"):
                self.btn_ctrls = []
            self.btn_ctrls.append({"ctrl": btn, "name": name})

        # 把整行加到外面传进来的 sizer
        sizer.Add(row_panel, 0, wx.TOP | wx.EXPAND, gap)

        return checklist, btn

    def create_hr(
        self, sizer, parent=None, gap=None, color=wx.Colour(47, 54, 60), border=0
    ):
        """
        在 sizer 中添加一条水平分割线，占据指定行高，并可设置背景色。
        :param sizer: 目标 sizer
        :param row_height: 分割线行高（控件高度）
        :param color: 分割线颜色
        :param bg_color: 背景颜色
        :param border: sizer 边距
        """
        if not gap:
            gap = self.gap
        if parent is None:
            parent = sizer.GetContainingWindow()  # 获取 sizer 所在面板

        # 创建行 Panel
        line_panel = wx.Panel(parent, size=(-1, gap))
        line_panel.SetBackgroundColour(self.get_test_color(parent))

        # 绘制细线
        def on_paint(event):
            dc = wx.PaintDC(line_panel)
            width, height = line_panel.GetSize()
            y = height - 1
            dc.SetPen(wx.Pen(color, 1))
            dc.DrawLine(0, y, width, y)

        line_panel.Bind(wx.EVT_PAINT, on_paint)
        sizer.Add(line_panel, 0, wx.EXPAND | wx.ALL, border)

    def create_run_btns(self, sizer, btn_group, parent=None, row_height=None, gap=None):
        """
        创建按钮，宽度略微拉伸以填满容器，固定高度，按钮之间间距固定
        """
        if not gap:
            gap = self.gap
        if not row_height:
            row_height = self.row_height
        if parent is None:
            parent = sizer.GetContainingWindow()

        btn_panel = wx.Panel(parent)
        btn_panel.SetBackgroundColour(wx.Colour(221, 221, 221))

        # 水平 BoxSizer
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_list = []

        for idx, item in enumerate(btn_group):
            name = item.get("name", "").strip()
            width = item.get("width", -1)
            if not name:
                name = item["title"]
            btn = wx.Button(btn_panel, label=item["title"], size=(width, row_height))
            btn.SetCursor(self.normal_cursor)
            # 去掉边框
            btn.SetWindowStyleFlag(wx.BORDER_NONE)
            # 设置圆角
            btn.SetWindowStyle(wx.NO_BORDER)
            btn.Bind(wx.EVT_BUTTON, item["event"])
            self.btn_ctrls.append({"ctrl": btn, "name": name})

            # proportion=1 表示按钮会均分剩余空间，wx.EXPAND 保证填满
            # 每个按钮四周留 gap/2
            if idx < len(btn_group) - 1:
                if width > -1:
                    btn_sizer.Add(btn, 0, wx.EXPAND | wx.RIGHT, gap // 2)
                else:
                    btn_sizer.Add(btn, 1, wx.EXPAND | wx.RIGHT, gap // 2)
            else:
                btn_sizer.Add(btn, 1, wx.EXPAND | wx.RIGHT, 0)

            btn_list.append(btn)

        # 给按钮面板四周留 gap/2 内边距
        outer_sizer = wx.BoxSizer(wx.VERTICAL)
        outer_sizer.Add(btn_sizer, 1, wx.EXPAND | wx.ALL, gap // 2)

        btn_panel.SetSizer(outer_sizer)

        # 添加到主 sizer 底部
        sizer.Add(btn_panel, 0, wx.EXPAND)

        return btn_list

    # region 创建日志控件
    def create_log_ctrls(
        self,
        panel,
        main_sizer,
        row_height=None,
        text_width=None,
        label_width=None,
        btn_width=None,
        gap=None,
    ):
        """
        创建日志控件。

        :param panel: 父面板
        :param main_sizer: 主布局器
        :param row_height: 行高
        :param text_width: 文本框宽度
        :param label_width: 标签宽度
        :param btn_width: 按钮宽度
        :param gap: 元素间距
        """
        # 初始化布局参数
        if row_height is None:
            row_height = self.row_height
        if text_width is None:
            text_width = self.text_width
        if label_width is None:
            label_width = self.label_width
        if btn_width is None:
            btn_width = self.btn_width
        if gap is None:
            gap = self.gap

        # 第一行按钮
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.AddStretchSpacer(1)

        btn_open = wx.Button(panel, label="打开日志", size=(btn_width, row_height))
        btn_clear = wx.Button(panel, label="清空日志", size=(btn_width, row_height))
        btn_up = wx.Button(panel, label="向上翻页", size=(btn_width, row_height))
        btn_down = wx.Button(panel, label="向下翻页", size=(btn_width, row_height))

        btn_group = [btn_open, btn_clear, btn_up, btn_down]
        # 设置按钮样式
        for btn in btn_group:
            btn.SetBackgroundColour(wx.Colour(35, 41, 46))  # #23292e
            btn.SetForegroundColour(wx.Colour(247, 247, 247))
            # 去掉边框
            btn.SetWindowStyleFlag(wx.BORDER_NONE)
            # 设置圆角
            btn.SetWindowStyle(wx.NO_BORDER)
            # 添加左右边距
            btn.SetWindowVariant(wx.WINDOW_VARIANT_NORMAL)
            btn.SetMinSize((btn.GetSize().width + gap / 2, btn.GetSize().height))

        # 设置按钮间距
        for btn in btn_group:
            btn_sizer.Add(btn, 0, wx.TOP | wx.RIGHT, gap)
        main_sizer.Add(btn_sizer, 0, wx.EXPAND)

        # 第二行 日志文本框
        self.log_text = wx.TextCtrl(
            panel,
            style=wx.TE_MULTILINE
            | wx.TE_READONLY
            | wx.TE_RICH2
            | wx.NO_BORDER
            | wx.TE_NO_VSCROLL,
            size=(text_width, row_height * 10),  # 日志框高度设为按钮高度的10倍
        )
        # 设置背景色和边框色为#2f363c
        self.log_text.SetBackgroundColour(wx.Colour(47, 54, 60))  # #2f363c
        main_sizer.Add(self.log_text, 1, wx.EXPAND | wx.ALL, gap)
        # 去掉内边距
        self.log_text.SetMargins(0, 0)

        self.btn_hovering = False  # 鼠标悬浮标志
        # 鼠标进入/离开按钮
        self.log_text.Bind(wx.EVT_ENTER_WINDOW, self.on_btn_enter)
        self.log_text.Bind(wx.EVT_LEAVE_WINDOW, self.on_btn_leave)
        # 滚轮事件绑定到父窗口或者 frame
        self.log_text.Bind(wx.EVT_MOUSEWHEEL, self.on_mouse_wheel)
        # 绑定按钮事件
        btn_clear.Bind(wx.EVT_BUTTON, self.on_clear_log)
        btn_open.Bind(wx.EVT_BUTTON, self.on_open_log)
        btn_up.Bind(wx.EVT_BUTTON, self.on_up_log)
        btn_down.Bind(wx.EVT_BUTTON, self.on_down_log)

        # 绑定 logger
        wx_handler = WxLogHandler(self.log_text)
        wx_handler.setFormatter(ColorFormatter("%(asctime)s %(message)s"))

        logger = logging.getLogger()
        logger.addHandler(wx_handler)

    def on_btn_enter(self, event):
        self.btn_hovering = True
        event.Skip()

    def on_btn_leave(self, event):
        self.btn_hovering = False
        event.Skip()

    def on_mouse_wheel(self, event):
        if self.btn_hovering:
            rotation = event.GetWheelRotation()
            lines = rotation // event.GetWheelDelta() * event.GetLinesPerAction()
            # 滚动 txt 文本控件
            current_pos = self.log_text.GetScrollPos(wx.VERTICAL)
            self.log_text.ScrollLines(-lines)  # wx TextCtrl 中向上滚动是负数
        event.Skip()

    def on_clear_log(self, event):
        self.log_text.Clear()

    def on_open_log(self, event):
        log_file = logger.log_path()
        if log_file:
            os.startfile(log_file)

    def on_up_log(self, event):
        self.log_text.ScrollPages(-1)  # 向上一页

    def on_down_log(self, event):
        # 向下滚动一页
        self.log_text.ScrollPages(1)  # 向下一页

    # endregion

    # endregion

    # region 事件处理

    def on_text_change(self, event, text_ctrl, config_key):
        """
        文本框内容改变事件处理函数。

        :param event: 事件对象
        :param text_ctrl: 文本框控件
        :param config_key: 配置键，用于保存文本框内容
        """
        if config_key.strip() and self.config:
            self.config.Write(config_key, text_ctrl.GetValue())
            self.config.Flush()

    def on_choice_change(self, event, choice_ctrl, config_key):
        """
        下拉框选择改变事件处理函数。

        :param event: 事件对象
        :param choice_ctrl: 下拉框控件
        :param config_key: 配置键，用于保存选择值
        """
        if config_key.strip() and self.config:
            self.config.Write(config_key, choice_ctrl.GetStringSelection())
            self.config.Flush()

    def on_select_folder(self, event, text_ctrl, title, config_key):
        """
        浏览文件夹事件处理函数。

        :param event: 事件对象
        :param text_ctrl: 文本框控件
        :param title: 对话框标题
        :param config_key: 配置键，用于保存选择的路径
        """
        dlg = wx.DirDialog(self.wx_frame, title, style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            text_ctrl.SetValue(dlg.GetPath())
            # 保存到配置文件
            if config_key.strip() and self.config:
                self.config.Write(config_key, dlg.GetPath())
                self.config.Flush()  # 使用 Flush() 保存，不要调用 Save()
        dlg.Destroy()

    def on_select_file(self, event, suffixs, text_ctrl, title, config_key):
        """
        浏览文件事件处理函数。

        :param event: 事件对象
        :param text_ctrl: 文本框控件
        :param suffixs: 文件后缀
        :param title: 对话框标题
        :param config_key: 配置键，用于保存选择的路径
        """
        # 构建文件类型过滤器
        wildcard = "所有文件 (*.*)|*.*"
        if suffixs and isinstance(suffixs, str):
            suffixs = [suffixs]
        if suffixs and isinstance(suffixs, list):
            filters = []
            for suffix in suffixs:
                suffix = suffix.strip()
                if suffix:
                    # 去掉前面的点号(如果有)
                    suffix = suffix.lstrip(".")
                    filters.append(f"*.{suffix}")
            if filters:
                wildcard = f"支持的文件 ({', '.join(filters)})|{';'.join(filters)}"

        dlg = wx.FileDialog(
            self.wx_frame,
            title,
            wildcard=wildcard,
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        )

        if dlg.ShowModal() == wx.ID_OK:
            text_ctrl.SetValue(dlg.GetPath())
            # 保存到配置文件
            if config_key.strip() and self.config:
                self.config.Write(config_key, dlg.GetPath())
                self.config.Flush()  # 使用 Flush() 保存，不要调用 Save()
        dlg.Destroy()

    def on_open_folder(self, event, text_ctrl):
        """
        打开文件夹按钮的点击事件处理函数
        """
        dir_path = text_ctrl.GetValue()
        if not os.path.exists(dir_path):
            logger.error(f"目录不存在：{dir_path}")
            return

        logger.info(f"尝试打开目录：{dir_path}")

        if os.name == "nt":
            os.startfile(dir_path)
        else:
            subprocess.run(["open" if os.name == "darwin" else "xdg-open", dir_path])

    # endregion

    # region 工具函数

    @staticmethod
    def get_size(max_width=None, max_height=None, min_width=400, min_height=300):
        """
        获取适合当前屏幕的窗口大小
        默认：宽度70%，高度80%
        并限制在最小/最大范围内
        """

        # 获取屏幕尺寸
        screen_width, screen_height = wx.GetDisplaySize()

        # 按比例计算
        width = int(screen_width * 0.7)
        height = int(screen_height * 0.8)

        # 限制最大值（如果传了的话）
        if max_width is not None:
            width = min(width, max_width)

        if max_height is not None:
            height = min(height, max_height)

        # 限制最小值
        width = max(width, min_width)
        height = max(height, min_height)

        return wx.Size(width, height)

    def get_test_color(self, panel, is_test=None):
        """返回随机颜色,非测试模式返回透明色"""
        if not is_test:
            is_test = self.is_test
        if is_test:
            # 随机生成RGB三个通道的值
            r = random.randint(0, 255)
            g = random.randint(0, 255)
            b = random.randint(0, 255)
            return wx.Colour(r, g, b)
        # 非测试模式返回默认颜色
        return panel.GetBackgroundColour()

    def get_text_ctrl(self, name):
        """
        获取指定配置键的文本框控件。

        :param config_key: 配置键
        :return: 文本框控件
        """
        for item in self.text_ctrls:
            if item["name"] == name:
                return item["ctrl"]
        return None

    def get_choice_ctrl(self, name):
        """
        获取指定配置键的下拉框控件。

        :param config_key: 配置键
        :return: 下拉框控件
        """
        for item in self.choice_ctrls:
            if item["name"] == name:
                return item["ctrl"]
        return None

    def get_multicheck_ctrl(self, name):
        """
        获取指定配置键的多选框控件。

        :param config_key: 配置键
        :return: 多选框控件
        """
        for item in self.multicheck_ctrls:
            if item["name"] == name:
                return item["ctrl"]
        return None

    def get_btn_ctrl(self, name):
        """
        获取指定配置键的按钮控件。

        :param config_key: 配置键
        :return: 按钮控件
        """
        for item in self.btn_ctrls:
            if item["name"] == name:
                return item["ctrl"]
        return None

    def toggle_btn(self, label, name=None, btn=None):
        """
        切换按钮状态
        :param btn: 按钮控件
        :param name: 按钮名字
        """
        if btn is None and (name is None or name.strip() == ""):
            logger.error("切换按钮状态失败，name和btn至少有一个有值")
            return

        if btn is None:
            btn = self.get_btn_ctrl(name=name)

        if btn is None:
            logger.error(f"未找到 name = {name} 的按钮")
            return

        # 获取当前状态
        is_enabled = getattr(btn, "_disabled_simulate", False)

        # 设置标签
        btn.SetLabel(label)

        btn.SetCursor(self.disabled_cursor if not is_enabled else self.normal_cursor)
        btn._disabled_simulate = not is_enabled
        btn.Refresh()

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

    @staticmethod
    def copy_to_clipboard(content):
        """
        复制内容到剪贴板
        :param content: 要复制的内容
        """
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(str(content)))
            wx.TheClipboard.Close()
            return True
        return False

    @staticmethod
    def custom_alert(
        title,
        msg,
        close_time=0,
        position_x=0,
        position_y=0,
        btn_group=None,
        initial="左上",
    ):
        """
        自定义弹窗：支持初始点（initial）四象限参考，position_left/position_top为对应角的相对偏移
        :param title: 弹窗标题 str
        :param msg: 弹窗内容 str
        :param close_time: 自动关闭，单位秒
        :param position_x: X轴偏移量（int，正负均可，和initial结合）
        :param position_y: Y轴偏移量（int，正负均可，和initial结合）
        :param btn_group: 按钮文本 list
        :param initial: 定位参考点，可选值["左上", "右上", "左下", "右下"]，默认为"左上"
        :return: 按钮文本或"自动关闭"
        """

        # region 使用示例
        # result = WxUtils.custom_alert(
        #         "提示111",
        #         "这是一个基本的梵蒂冈\n变化幅度和提示框",
        #         0,
        #         "auto",
        #         "100",
        #         [],
        #         "右上",
        #     )
        #     print(f"基础用法结果: {result}")
        # endregion

        dlg = MiniAlert(
            title, msg, close_time, position_x, position_y, btn_group, initial
        )
        dlg.ShowModal()
        dlg.Destroy()  # 必须销毁！
        return getattr(dlg, "result", None)

    # endregion


# region 日志捕获


class ColorFormatter(logging.Formatter):
    """支持 16 进制颜色代码的自定义 Formatter"""

    COLOR_CODES = {
        "DIVIDER": "#eeeeee",
        "DEBUG": "#2196F3",
        "INFO": "#999999",
        "SUCCESS": "#4CAF50",
        "WARNING": "#FF9800",
        "ERROR": "#F44336",
        "CRITICAL": "#B71C1C",
    }

    def _hex_to_ansi(self, hex_color):
        hex_color = hex_color.lstrip("#")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return f"\033[38;2;{r};{g};{b}m"

    def format(self, record):
        default_color = self.COLOR_CODES.get(record.levelname, "#FFFFFF")
        color = getattr(record, "color", default_color)
        ansi_color = self._hex_to_ansi(color)
        message = super().format(record)
        return f"{ansi_color}{message}\033[0m"


class WxLogHandler(logging.Handler):
    """把带 ANSI 颜色码的日志输出到 wx.TextCtrl"""

    ANSI_TRUECOLOR_FG = re.compile(r"\x1b\[38;2;(\d+);(\d+);(\d+)m")
    ANSI_RESET = "\x1b[0m"

    def __init__(self, text_ctrl: wx.TextCtrl):
        super().__init__()
        self.text_ctrl = text_ctrl
        self.default_attr = wx.TextAttr(wx.BLACK, wx.NullColour)
        # 设置字体
        font_list = [
            "Maple Mono NF CN",
            "Menlo",
            "Consolas",
            "Maple UI",
            "PingFang",
            "Microsoft YaHei",
            "monospace",
        ]
        self.font = self._get_first_available_font(font_list)
        self.text_ctrl.SetFont(self.font)
        self.line_count = 0

    def _get_first_available_font(self, font_names):
        """按顺序返回第一个可用字体"""
        for name in font_names:
            font = wx.Font(wx.FontInfo().FaceName(name))
            if font.IsOk():
                return font
        return wx.Font(wx.FontInfo())  # 默认字体

    def emit(self, record):
        msg = self.format(record)
        wx.CallAfter(self._append, msg)
        self.flush()  # ✅ 强制刷新

    def flush(self):
        pass  # 这里不用真的写，因为 wx.TextCtrl 是立即写入的

    def _append_text(self, text, color_tuple):
        if not text:
            return

        # 自动换行
        if not text.endswith("\n"):
            text += "\n"

        # 拆分多行
        lines = text.splitlines()
        for line in lines:
            self.line_count += 1
            line_number_str = f"{self.line_count:4d} "  # 行号 + 空格
            start = self.text_ctrl.GetLastPosition()
            # 先插入行号
            self.text_ctrl.AppendText(line_number_str)
            end = self.text_ctrl.GetLastPosition()
            self.text_ctrl.SetStyle(
                start, end, wx.TextAttr(wx.Colour(255, 255, 255))
            )  # 白色

            # 再插入日志内容
            start = self.text_ctrl.GetLastPosition()
            self.text_ctrl.AppendText(line + "\n")
            end = self.text_ctrl.GetLastPosition()
            self.text_ctrl.SetStyle(start, end, wx.TextAttr(wx.Colour(*color_tuple)))

    def _append(self, msg):
        pos = 0
        text_len = len(msg)
        while pos < text_len:
            match = self.ANSI_TRUECOLOR_FG.search(msg, pos)
            if match:
                start, end = match.span()
                # 获取 RGB
                r, g, b = map(int, match.groups())
                fg_color = (r, g, b)
                # 找重置位置
                reset_pos = msg.find(self.ANSI_RESET, end)
                if reset_pos == -1:
                    reset_pos = text_len
                self._append_text(msg[end:reset_pos], fg_color)
                pos = reset_pos + len(self.ANSI_RESET)
            else:
                # 没有 ANSI，按级别颜色
                self._append_text(msg[pos:], (0, 0, 0))
                break

    def _append_with_attr(self, text, attr):
        if not text:
            return
        start = self.text_ctrl.GetLastPosition()
        self.text_ctrl.AppendText(text)
        end = self.text_ctrl.GetLastPosition()
        self.text_ctrl.SetStyle(start, end, attr)
        self.text_ctrl.Update()

    def _apply_ansi_codes(self, attr, codes):
        new_attr = wx.TextAttr(
            attr.GetTextColour(), attr.GetBackgroundColour(), attr.GetFont()
        )
        it = iter(codes)

        for code in it:
            if not code.isdigit():
                continue
            c = int(code)

            if c == 0:  # reset
                new_attr = wx.TextAttr(
                    self.default_attr.GetTextColour(),
                    self.default_attr.GetBackgroundColour(),
                    self.default_attr.GetFont(),
                )
            elif c == 1:  # bold
                font = new_attr.GetFont() or wx.Font(
                    10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD
                )
                font.SetWeight(wx.FONTWEIGHT_BOLD)
                new_attr.SetFont(font)
            elif 30 <= c <= 37:
                new_attr.SetTextColour(self._ansi_16_color(c - 30))
            elif 40 <= c <= 47:
                new_attr.SetBackgroundColour(self._ansi_16_color(c - 40))
            elif c in (38, 48):  # TrueColor / 256色
                try:
                    mode = next(it)
                    if mode == "2":  # TrueColor
                        r, g, b = int(next(it)), int(next(it)), int(next(it))
                        color = wx.Colour(r, g, b)
                    elif mode == "5":  # 256色
                        idx = int(next(it))
                        color = self._ansi_256_color(idx)
                    else:
                        continue

                    if c == 38:
                        new_attr.SetTextColour(color)
                    else:
                        new_attr.SetBackgroundColour(color)
                except StopIteration:
                    pass
        return new_attr

    def _ansi_16_color(self, idx):
        """简单的 16 色映射"""
        table = [
            wx.BLACK,
            wx.RED,
            wx.GREEN,
            wx.YELLOW,
            wx.BLUE,
            wx.CYAN,
            wx.LIGHT_GREY,
            wx.WHITE,
        ]
        return table[idx % len(table)]

    def _ansi_256_color(self, idx):
        """简单的 256 色映射，按灰度 fallback"""
        return wx.Colour(idx, idx, idx)

    def _apply_ansi_codes(self, attr, codes):
        """解析 ANSI 颜色码，返回新的 wx.TextAttr"""
        new_attr = wx.TextAttr(
            attr.GetTextColour(), attr.GetBackgroundColour(), attr.GetFont()
        )
        it = iter(codes)

        for code in it:
            if not code.isdigit():
                continue
            c = int(code)

            if c == 0:  # reset
                new_attr = wx.TextAttr(
                    self.default_attr.GetTextColour(),
                    self.default_attr.GetBackgroundColour(),
                    self.default_attr.GetFont(),
                )
            elif c == 1:  # bold
                font = new_attr.GetFont() or wx.Font(
                    10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD
                )
                font.SetWeight(wx.FONTWEIGHT_BOLD)
                new_attr.SetFont(font)
            elif 30 <= c <= 37:
                new_attr.SetTextColour(self._ansi_16_color(c - 30))
            elif 40 <= c <= 47:
                new_attr.SetBackgroundColour(self._ansi_16_color(c - 40))
            elif c in (38, 48):  # TrueColor / 256 色
                try:
                    mode = next(it)
                    if mode == "2":  # TrueColor
                        r, g, b = int(next(it)), int(next(it)), int(next(it))
                        color = wx.Colour(r, g, b)
                    elif mode == "5":  # 256 色
                        idx = int(next(it))
                        color = self._ansi_256_color(idx)
                    else:
                        continue

                    if c == 38:
                        new_attr.SetTextColour(color)
                    else:
                        new_attr.SetBackgroundColour(color)
                except StopIteration:
                    pass
        return new_attr

    # -------- 颜色映射辅助 --------

    def _ansi_256_color(self, idx):
        if idx < 16:
            return self._ansi_16_color(idx % 8)
        elif 16 <= idx <= 231:
            idx -= 16
            r = (idx // 36) % 6 * 51
            g = (idx // 6) % 6 * 51
            b = idx % 6 * 51
            return wx.Colour(r, g, b)
        elif 232 <= idx <= 255:
            gray = (idx - 232) * 10 + 8
            return wx.Colour(gray, gray, gray)
        return wx.Colour(255, 255, 255)

    def _ansi_16_color(self, idx):
        table = [
            wx.Colour(0, 0, 0),  # 黑
            wx.Colour(128, 0, 0),  # 红
            wx.Colour(0, 128, 0),  # 绿
            wx.Colour(128, 128, 0),  # 黄
            wx.Colour(0, 0, 128),  # 蓝
            wx.Colour(128, 0, 128),  # 品红
            wx.Colour(0, 128, 128),  # 青
            wx.Colour(192, 192, 192),  # 白（灰）
        ]
        return table[idx % 8]


# endregion
