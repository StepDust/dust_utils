import wx

# 配置日志
from loguru import logger


class MiniAlert(wx.Dialog):
    def __init__(self, title, msg, close_time, pos_x, pos_y, btn_group, initial):
        super().__init__(
            None, title=title, style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP
        )
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        top_sizer = wx.BoxSizer(wx.HORIZONTAL)
        bitmap = wx.ArtProvider.GetBitmap(wx.ART_INFORMATION, wx.ART_OTHER, (32, 32))
        icon = wx.StaticBitmap(panel, bitmap=bitmap)
        top_sizer.Add(icon, 0, wx.LEFT | wx.TOP, 20)
        self.SetIcon(wx.ArtProvider.GetIcon(wx.ART_FILE_SAVE, wx.ART_FRAME_ICON))
        message_text = wx.StaticText(panel, label=msg)
        font = message_text.GetFont()
        font.SetPointSize(10)
        message_text.SetFont(font)
        top_sizer.Add(message_text, 1, wx.LEFT | wx.TOP | wx.RIGHT, 20)
        main_sizer.Add(top_sizer, 1, wx.EXPAND)
        line = wx.StaticLine(panel, style=wx.LI_HORIZONTAL)
        main_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.countdown_label = wx.StaticText(panel, label="")
        button_sizer.Add(
            self.countdown_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 20
        )
        button_sizer.AddStretchSpacer()
        self.btn_map = {}
        for idx, label in enumerate(btn_group or ["确定"]):
            btn = wx.Button(panel, id=10000 + idx, label=label)
            btn.SetBackgroundColour(wx.Colour(230, 230, 230))
            btn.Bind(wx.EVT_BUTTON, self.on_button)
            button_sizer.Add(btn, 0, wx.LEFT, 10)
            self.btn_map[btn.GetId()] = label
        main_sizer.Add(button_sizer, 0, wx.EXPAND | wx.ALL, 20)
        panel.SetSizer(main_sizer)
        self.SetMinSize(wx.Size(600, 225))
        self.SetSize(600, 225)
        display_size = wx.DisplaySize()
        frame_size = self.GetSize()
        # 处理 auto 情况，实现自动居中并计算偏移
        frame_w, frame_h = frame_size[0], frame_size[1]
        screen_w, screen_h = display_size[0], display_size[1]
        # 解析 left/top
        if isinstance(pos_x, str) and pos_x == "auto":
            pl = 0
            left_auto = True
        else:
            pl = int(pos_x)
            left_auto = False
        if isinstance(pos_y, str) and pos_y == "auto":
            pt = 0
            top_auto = True
        else:
            pt = int(pos_y)
            top_auto = False
        # 计算起始点居中和偏移
        if initial == "左上":
            base_x = (screen_w - frame_w) // 2 if left_auto else 0
            base_y = (screen_h - frame_h) // 2 if top_auto else 0
            new_x = base_x + pl
            new_y = base_y + pt
        elif initial == "右上":
            base_x = (screen_w - frame_w) // 2 if left_auto else (screen_w - frame_w)
            base_y = (screen_h - frame_h) // 2 if top_auto else 0
            new_x = base_x - pl if left_auto else base_x - pl
            new_y = base_y + pt
        elif initial == "左下":
            base_x = (screen_w - frame_w) // 2 if left_auto else 0
            base_y = (screen_h - frame_h) // 2 if top_auto else (screen_h - frame_h)
            new_x = base_x + pl
            new_y = base_y - pt if top_auto else base_y - pt
        elif initial == "右下":
            base_x = (screen_w - frame_w) // 2 if left_auto else (screen_w - frame_w)
            base_y = (screen_h - frame_h) // 2 if top_auto else (screen_h - frame_h)
            new_x = base_x - pl if left_auto else base_x - pl
            new_y = base_y - pt if top_auto else base_y - pt
        else:
            base_x = (screen_w - frame_w) // 2 if left_auto else 0
            base_y = (screen_h - frame_h) // 2 if top_auto else 0
            new_x = base_x + pl
            new_y = base_y + pt
        self.SetPosition((new_x, new_y))
        if wx.Platform == "__WXMSW__":
            self.SetWindowStyle(self.GetWindowStyle() | wx.BORDER_NONE)
            self.SetTransparent(254)
        self.remaining_time = close_time
        self.result = None
        if self.remaining_time > 0:
            self.timer = wx.Timer(self)
            self.Bind(wx.EVT_TIMER, self.on_timer, self.timer)
            self.timer.Start(1000)
            self.update_countdown()

    def on_button(self, event):
        self.result = self.btn_map.get(event.GetEventObject().GetId(), None)
        self.EndModal(0)

    def on_timer(self, event):
        self.remaining_time -= 1
        if self.remaining_time <= 0:
            self.result = "自动关闭"
            self.EndModal(0)
        else:
            self.update_countdown()

    def update_countdown(self):
        self.countdown_label.SetLabel(f"倒计时: {self.remaining_time}秒")
