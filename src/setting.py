import wx
import json
import os

def convert_value_type(key, value):

    bool_keys = ["autoUpdate", "debug"]
    # 定义需要转换为整数的键
    int_keys = ["decimalPlaceNum", "initItemsRow", "initNamesCol", "newGridRow", "newGridCol"]
    
    # 转换为布尔值
    if key in bool_keys:
        if isinstance(value, str):
            return value.lower() in ("true", "1", "yes", "on")
        return bool(value)
    
    # 转换为整数
    if key in int_keys:
        try:
            return int(value)
        except (ValueError, TypeError):
            return 0  # 默认值
            
    # 其他情况保持原值
    return value

class OptionsFrame(wx.Frame):

    def __init__(self, *args, **kw):
        super(OptionsFrame, self ).__init__(*args, **kw)
        # 显式设置最小窗口大小
        self.SetSizeHints(500, 400, -1, -1)
        self.panel = wx.Panel(self)
        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # 读取_json/options.json文件
        self.options_file = '_json/options.json'
        if os.path.exists(self.options_file):
            with open(self.options_file, 'r', encoding='utf-8') as f:
                self.options = json.load(f)
        else:
            self.options = {}

        # 创建滚动窗口
        self.scrolled_window = wx.ScrolledWindow(self.panel, -1, style=wx.VSCROLL | wx.HSCROLL)
        self.scrolled_window.SetScrollRate(5, 5)
        self.scrolled_sizer = wx.BoxSizer(wx.VERTICAL)

        # 创建重置和保存配置按钮
        self.reset_button = wx.Button(self.panel, label='重置')
        self.save_button = wx.Button(self.panel, label='保存配置')
        self.reset_button.Bind(wx.EVT_BUTTON, self.on_reset)
        self.save_button.Bind(wx.EVT_BUTTON, self.on_save)

        # 创建静态文本和输入框
        self.input_boxes = {}
        for key, value in self.options.items():
            box_sizer = wx.BoxSizer(wx.HORIZONTAL)
            static_text = wx.StaticText(self.scrolled_window, label=key)
            # 如果是版本号字段，设置为只读
            if key == "version":
                input_box = wx.TextCtrl(self.scrolled_window, style=wx.TE_READONLY)
            else:
                input_box = wx.TextCtrl(self.scrolled_window)
            input_box.SetValue(str(value))
            self.input_boxes[key] = input_box

            box_sizer.Add(static_text, 0, wx.ALL | wx.CENTER, 5)
            box_sizer.Add(input_box, 1, wx.ALL | wx.EXPAND, 5)
            self.scrolled_sizer.Add(box_sizer, 0, wx.EXPAND | wx.ALL, 5)

        self.scrolled_window.SetSizer(self.scrolled_sizer)
        self.scrolled_window.Layout()
        self.scrolled_sizer.Fit(self.scrolled_window)

        # 将滚动窗口和按钮添加到主面板的布局中
        self.sizer.Add(self.scrolled_window, 1, wx.EXPAND | wx.ALL, 5)
        self.sizer.Add(self.reset_button, 0, wx.ALL | wx.CENTER, 5)
        self.sizer.Add(self.save_button, 0, wx.ALL | wx.CENTER, 5)

        self.panel.SetSizer(self.sizer)
        self.panel.Layout()
        self.sizer.Fit(self)

    def on_reset(self, event):
        # 重置时从options.json重新加载版本号
        if os.path.exists(self.options_file):
            with open(self.options_file, 'r', encoding='utf-8') as f:
                self.options = json.load(f)
        for key, input_box in self.input_boxes.items():
            input_box.SetValue(str(self.options[key]))

    def on_save(self, event):
        for key, input_box in self.input_boxes.items():
            # 版本号字段不保存，保持原值
            if key != "version":
                value = input_box.GetValue()
                # 转换值的类型
                self.options[key] = convert_value_type(key, value)
        with open(self.options_file, 'w', encoding='utf-8') as f:
            json.dump(self.options, f, ensure_ascii=False, indent=4)
        wx.MessageBox('配置已保存', '信息', wx.OK | wx.ICON_INFORMATION)

if __name__ == '__main__':
    app = wx.App(False)
    frame = OptionsFrame(None, title='配置选项', size=(500, 400))
    frame.Show()
    app.MainLoop()
