# -*- coding: utf-8 -*-

#--------------------------------------------------------------------------
# Python code generated with wxFormBuilder (version 3.9.0 Jun 14 2020)
# http://www.wxformbuilder.org/
#
# PLEASE DO *NOT* EDIT THIS FILE!
#--------------------------------------------------------------------------

import wx
import wx.xrc
import wx.grid
import gettext
import excel as e
import pandas as pd
import os
import requests
import json
import threading
import re
import sys
import os
from setting import OptionsFrame

_ = gettext.gettext

# --------------------------------------------------------------------------
#  AI 聊天窗口 
# --------------------------------------------------------------------------
class AIChatFrame(wx.Frame):
	def __init__(self, parent):
		wx.Frame.__init__(self, parent, id=wx.ID_ANY, title="AI 助手", pos=wx.DefaultPosition, size=wx.Size(600, 500), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)
		self.parent = parent
		self.conversation_history = []
		self.is_destroyed = False # 窗口销毁标志

		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour(255, 255, 255))

		main_sizer = wx.BoxSizer(wx.VERTICAL)

		self.history_ctrl = wx.TextCtrl(self, wx.ID_ANY, "", wx.DefaultPosition, wx.DefaultSize, wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2)
		main_sizer.Add(self.history_ctrl, 1, wx.ALL | wx.EXPAND, 5)

		input_sizer = wx.BoxSizer(wx.HORIZONTAL)
		self.input_ctrl = wx.TextCtrl(self, wx.ID_ANY, "", wx.DefaultPosition, wx.DefaultSize, wx.TE_PROCESS_ENTER)
		input_sizer.Add(self.input_ctrl, 1, wx.ALL | wx.EXPAND, 5)
		self.send_button = wx.Button(self, wx.ID_ANY, "发送", wx.DefaultPosition, wx.DefaultSize, 0)
		input_sizer.Add(self.send_button, 0, wx.ALL, 5)
		main_sizer.Add(input_sizer, 0, wx.EXPAND, 5)

		self.SetSizer(main_sizer)
		self.Layout()
		self.Centre(wx.BOTH)
		
		self.send_button.Bind(wx.EVT_BUTTON, self.on_send)
		self.input_ctrl.Bind(wx.EVT_TEXT_ENTER, self.on_send)
		self.Bind(wx.EVT_CLOSE, self.on_close)

	def on_close(self, event):
		self.is_destroyed = True # 设置销毁标志
		if self.parent:
			self.parent.ai_chat_frame = None
		self.Destroy()

	def on_send(self, event):
		user_message = self.input_ctrl.GetValue().strip()
		if not user_message:
			return

		self.history_ctrl.SetDefaultStyle(wx.TextAttr(wx.BLUE))
		self.history_ctrl.AppendText(f"You:\n{user_message}\n\n")

		self.input_ctrl.Clear()
		self.send_button.Disable()
		self.input_ctrl.Disable()

		self.conversation_history.append({"role": "user", "content": user_message})
		threading.Thread(target=self.get_ai_response, args=(user_message,), daemon=True).start()

	def get_ai_response(self, user_message):
		if self.is_destroyed: return
		wx.CallAfter(self.history_ctrl.SetDefaultStyle, wx.TextAttr(wx.Colour(0, 128, 0)))
		# 检查是否是第一次AI回答
		if len(self.conversation_history) == 1:  # 只有系统消息和第一条用户消息
			wx.CallAfter(self.history_ctrl.AppendText, "AI:\n[提示] AI回答可能需要排队一段时间，请耐心等待\n")

		api_url = "https://api.suanli.cn/v1/chat/completions"
		api_key = "sk-OTi0r196VHjX2iMgNaPevYrXSP4VKO4s2coOjIyPdXq02okY"
		headers = {'Authorization': f'Bearer {api_key}', 'Content-Type': 'application/json'}

		grid_data = self.parent.get_data_for_saving()
		xml_data_string = e.data_to_xml_string(grid_data)

		system_prompt = f"""你是一个强大的表格处理助手，也能闲聊。
这是当前表格的XML数据：
<data>
{xml_data_string}
</data>
请根据我的要求进行对话或操作。
如果需要修改表格，请在你的回答中包含一个用```xml ... ```包围的、完整的、新的表格XML代码块。
XML的格式必须是 <root><row><col>...</col></row>...</root>，不要有<data>标签，不要被现有表格大小所拘束。
如果没有修改表格，就正常聊天，不要输出XML。
"""
		
		messages_to_send = [{"role": "system", "content": system_prompt}] + self.conversation_history
		payload = {"model": "free:Qwen3-30B-A3B", "messages": messages_to_send, "stream": True}

		full_response_content = ""
		try:
			with requests.post(api_url, headers=headers, json=payload, stream=True, timeout=300) as response:
				response.raise_for_status()
				for chunk in response.iter_lines():
					if self.is_destroyed: break
					if chunk:
						decoded_chunk = chunk.decode('utf-8')
						if decoded_chunk.startswith('data:'):
							json_data_str = decoded_chunk[len('data:'):].strip()
							if json_data_str and json_data_str != '[DONE]':
								try:
									json_data = json.loads(json_data_str)
									content_chunk = json_data.get('choices', [{}])[0].get('delta', {}).get('content', '')
									if content_chunk:
										full_response_content += content_chunk
										wx.CallAfter(self.update_history_text, content_chunk)
								except json.JSONDecodeError:
									continue
		except requests.exceptions.RequestException as err:
			if not self.is_destroyed:
				wx.CallAfter(self.update_history_text, f"\n[网络错误]: {err}", wx.RED)
		finally:
			if not self.is_destroyed:
				wx.CallAfter(self.finalize_response, full_response_content)

	def update_history_text(self, text, color=None):
		if self.is_destroyed: return
		original_style = self.history_ctrl.GetDefaultStyle()
		if color:
			self.history_ctrl.SetDefaultStyle(wx.TextAttr(color))
		self.history_ctrl.AppendText(text)
		if color:
			self.history_ctrl.SetDefaultStyle(original_style)

	def finalize_response(self, full_response):
		if self.is_destroyed: return
		
		self.conversation_history.append({"role": "assistant", "content": full_response})
		self.update_history_text("\n\n")

		xml_string = None
		# 查找所有```xml的位置
		xml_start_matches = list(re.finditer(r'```xml', full_response))
		if xml_start_matches:
			# 使用最后一个```xml标记
			last_match = xml_start_matches[-1]
			start_pos = last_match.end()  # 从```xml之后开始
			
			# 从该位置向后查找```
			end_pos = full_response.find('```', start_pos)
			if end_pos != -1:
				# 提取中间的内容
				xml_string = full_response[start_pos:end_pos].strip()
		
		# 如果没有找到```xml ... ```格式，检查整个响应是否是XML
		if not xml_string:
			stripped_response = full_response.strip()
			if stripped_response.startswith('<root>') and stripped_response.endswith('</root>'):
				xml_string = stripped_response
		
		if xml_string:
			new_data = e.xml_string_to_data(xml_string)
			if new_data is not None:
				# 显示预览并询问用户是否同意修改
				wx.CallAfter(self.show_preview_and_confirm, new_data)

		self.send_button.Enable()
		self.input_ctrl.Enable()
		self.input_ctrl.SetFocus()

	def show_preview_and_confirm(self, new_data):
		# 创建预览窗口
		preview_frame = PreviewFrame(self, new_data)
		preview_frame.Show()
		
	def apply_ai_changes(self, new_data):
		# 应用AI的更改
		wx.CallAfter(self.parent.update_grid_with_data, new_data)
		self.update_history_text("[提示]: 已根据AI的回复更新表格内容。\n\n", wx.RED)

#--------------------------------------------------------------------------
#  Class excelPage
#---------------------------------------------------------------------------
class excelPage_ ( wx.Frame ):
	def __init__(self, parent,path):
		wx.Frame.__init__ (self, parent,
						   id = wx.ID_ANY, title = _(f"123Excel II - {path}"),
						   pos = wx.DefaultPosition, size = wx.Size( 1000,600 ),
						   style = wx.DEFAULT_FRAME_STYLE|wx.BORDER_NONE|wx.TAB_TRAVERSAL
							)
		self.path = path
		self.file_type = os.path.splitext(path)[1].lower()
		self.ai_chat_frame = None
		self.Maximize()
		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour( 255, 255, 255 ))

		main = wx.BoxSizer(wx.VERTICAL)
		toolBar = wx.BoxSizer(wx.HORIZONTAL)
		mainTool = wx.BoxSizer(wx.VERTICAL)
		tags = wx.BoxSizer(wx.HORIZONTAL)

		self.tagFile = wx.Button(self, wx.ID_ANY, _(u"文件"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.tagFile.SetToolTip(_(u"对于文件的更改，选择"))
		tags.Add(self.tagFile, 0, 0, 5)

		self.tagJiSuan = wx.Button(self, wx.ID_ANY, _(u"运算"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.tagJiSuan.SetToolTip(_(u"对于表格项目的运算"))
		tags.Add(self.tagJiSuan, 0, 0, 5)

		self.tagAutom = wx.Button(self, wx.ID_ANY, _(u"自动化"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.tagAutom.SetToolTip(_(u"使用python自定义自动化项目"))
		tags.Add(self.tagAutom, 0, 0, 5)

		self.tagAI = wx.Button(self, wx.ID_ANY, _(u"AI"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.tagAI.SetToolTip(_(u"打开AI助手进行聊天或操作表格"))
		tags.Add(self.tagAI, 0, 0, 5)

		self.tagNone1 = wx.Button(self, wx.ID_ANY, _(u"None"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.tagNone1, 0, 0, 5)

		self.tagNone2 = wx.Button(self, wx.ID_ANY, _(u"None"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.tagNone2, 0, 0, 5)

		self.tagNone3 = wx.Button(self, wx.ID_ANY, _(u"None"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.tagNone3, 0, 0, 5)

		self.tagNone4 = wx.Button(self, wx.ID_ANY, _(u"None"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.tagNone4, 0, 0, 5)

		self.tagNone5 = wx.Button(self, wx.ID_ANY, _(u"None"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.tagNone5, 0, 0, 5)

		self.tagNone6 = wx.Button(self, wx.ID_ANY, _(u"None"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.tagNone6, 0, 0, 5)
		mainTool.Add(tags, 0, 0, 5)

		tools = wx.BoxSizer(wx.HORIZONTAL)
		self.save = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/save.png", wx.BITMAP_TYPE_ANY ))
		self.save.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/save1.png", wx.BITMAP_TYPE_ANY ))
		self.save.SetToolTip(_(u"保存文件"))
		tools.Add(self.save, 0, 0, 5)

		self.saveas = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/saveas.png", wx.BITMAP_TYPE_ANY ))
		self.saveas.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/saveas1.png", wx.BITMAP_TYPE_ANY ))
		self.saveas.SetToolTip(_(u"另存为"))
		tools.Add(self.saveas, 0, 0, 5)

		self.getMax = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/max.png", wx.BITMAP_TYPE_ANY ))
		self.getMax.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/max1.png", wx.BITMAP_TYPE_ANY ))
		self.getMax.Hide()
		self.getMax.SetToolTip(_(u"最大值"))
		tools.Add(self.getMax, 0, 0, 5)

		self.getMin = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/min.png", wx.BITMAP_TYPE_ANY ))
		self.getMin.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/min1.png", wx.BITMAP_TYPE_ANY ))
		self.getMin.Hide()
		self.getMin.SetToolTip(_(u"最小值"))
		tools.Add(self.getMin, 0, 0, 5)

		self.getAvg = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/avg.png", wx.BITMAP_TYPE_ANY ))
		self.getAvg.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/avg1.png", wx.BITMAP_TYPE_ANY ))
		self.getAvg.Hide()
		self.getAvg.SetToolTip(_(u"平均值"))
		tools.Add(self.getAvg, 0, 0, 5)

		self.getCustomize = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/customize.png", wx.BITMAP_TYPE_ANY ))
		self.getCustomize.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/customize1.png", wx.BITMAP_TYPE_ANY ))
		self.getCustomize.Hide()
		self.getCustomize.SetToolTip(_(u"自定义准则，挑选符合准则的项"))
		tools.Add(self.getCustomize, 0, 0, 5)

		self.Paiming = wx.Button(self, wx.ID_ANY, _(u"自动标注排名"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.Paiming.Hide()
		self.Paiming.SetToolTip(_(u"自动排名"))
		tools.Add(self.Paiming, 0, 0, 5)

		self.RankSubSort = wx.Button(self, wx.ID_ANY, _(u"自动按值排列"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.RankSubSort.Hide()
		self.RankSubSort.SetToolTip(_(u"自动按值排列"))
		tools.Add(self.RankSubSort, 0, 0, 5)

		self.customizeInput = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
		self.customizeInput.Hide()
		self.customizeInput.SetToolTip(
"""自定义准则填写
x表示项内所有值
支持的符号:
== 等于  	>= 大于等于	<= 小于等于
>    大于	<    小于	!=   不等于
以及各种数学运算符
+-*/加减乘除	**次方	//整除	%取余
逻辑运算:(前后要空格)
and 和，表示同时满足两个准则，可重复使用，
    如：x>=3 and x<10 and x!=6
or  或，表示满足两个准则,可重复使用
not  非，表示不满足某种条件
如：not x>10 等价于 x<=10"
在最后写上"#"再写上 x - 72之类的算术式可以同时输出算数结果，如：
张三|100
李四|103
x>=100 # x-100
输出:
x>=100 # x-100:
张三:100|0
李四:103|3
""")
		tools.Add(self.customizeInput, 1, 0, 5)
		mainTool.Add(tools, 1, wx.EXPAND, 5)

		values = wx.BoxSizer(wx.HORIZONTAL)
		self.funcText = wx.StaticText(self, wx.ID_ANY, _(u"f(x)"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.funcText.Wrap(-1)
		values.Add(self.funcText, 0, wx.ALL, 5)

		self.inputFanc = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
		self.inputFanc.SetToolTip(_(u"函数"))
		values.Add(self.inputFanc, 1, wx.EXPAND, 5)

		sheetChoiceChoices = []
		self.sheetChoice = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, sheetChoiceChoices, 0)
		if self.file_type == '.xlsx':
			try:
				excel_file = pd.ExcelFile(path)
				sheets = excel_file.sheet_names
				for sheet in sheets:
					self.sheetChoice.Append(sheet)
				self.sheetChoice.SetSelection(0)
			except Exception as err:
				wx.MessageBox(f"无法加载Excel工作表: {err}", "错误", wx.OK | wx.ICON_ERROR)
		else:
			self.sheetChoice.Append("默认")
			self.sheetChoice.SetSelection(0)
			self.sheetChoice.Disable()
		self.sheetChoice.SetToolTip(_(u"sheet"))
		values.Add(self.sheetChoice, 0, 0, 5)

		self.nameCol = wx.StaticText(self, wx.ID_ANY, _(u"名称列"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.nameCol.Wrap(-1)
		self.nameCol.Hide()
		values.Add(self.nameCol, 0, wx.TOP|wx.LEFT, 5)

		self.nameColInput = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
		self.nameColInput.Hide()
		self.nameColInput.SetToolTip(_(u"所需运算的值对应的名称\n例如:学生名所在行"))
		values.Add(self.nameColInput, 0, wx.RIGHT, 5)

		self.itemRow = wx.StaticText(self, wx.ID_ANY, _(u"项目行"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.itemRow.Wrap(-1)
		self.itemRow.Hide()
		values.Add(self.itemRow, 0, wx.TOP|wx.LEFT, 5)

		self.itemRowInput1 = wx.TextCtrl(self, wx.ID_ANY, "1", wx.DefaultPosition, wx.DefaultSize, 0)
		self.itemRowInput1.Hide()
		self.itemRowInput1.SetToolTip(_(u"填写项目的列\n例如:写科目的列"))
		values.Add(self.itemRowInput1, 0, 0, 5)

		itemChoiceChoices = []
		self.itemChoice = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, itemChoiceChoices, 0)
		self.itemChoice.Hide()
		self.itemChoice.SetToolTip(_(u"项目"))
		values.Add(self.itemChoice, 0, wx.LEFT, 5)
		mainTool.Add(values, 1, wx.EXPAND, 5)
		toolBar.Add(mainTool, 1, wx.EXPAND, 5)

		others = wx.BoxSizer(wx.HORIZONTAL)
		self.setting = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/setting.png", wx.BITMAP_TYPE_ANY ))
		self.setting.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/setting1.png", wx.BITMAP_TYPE_ANY ))
		self.setting.SetToolTip(_(u"设置"))
		others.Add(self.setting, 0, wx.ALL, 5)

		self.exit = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/exit.png", wx.BITMAP_TYPE_ANY ))
		self.exit.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/exit1.png", wx.BITMAP_TYPE_ANY ))
		self.exit.SetToolTip(_(u"退出"))
		others.Add(self.exit, 0, wx.ALL, 5)
		toolBar.Add(others, 0, 0, 5)
		toolBar.AddStretchSpacer()
		main.Add(toolBar, 0, wx.EXPAND, 5)

		grid_sizer = wx.BoxSizer(wx.HORIZONTAL)
		self.mainGrid = wx.grid.Grid(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)
		
		self.mainGrid.CreateGrid(0, 0)
		data = []
		try:
			if self.file_type == '.xlsx':
				if self.sheetChoice.GetCount() > 0:
					sheet_name = self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					data = e.readExcel(path, sheet_name)
			elif self.file_type == '.xml':
				data = e.readXml(path)
			elif self.file_type == '.json':
				data = e.readJson(path)
			else:
				wx.MessageBox("不支持的文件类型！", "错误", wx.OK | wx.ICON_ERROR)
			self.update_grid_with_data(data)
		except Exception as err:
			wx.MessageBox(f"打开文件时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
			self.update_grid_with_data([])

		self.mainGrid.EnableEditing(True)
		self.mainGrid.EnableGridLines(True)
		self.mainGrid.EnableDragGridSize(False)
		self.mainGrid.SetMargins(0, 0)
		self.mainGrid.EnableDragColMove(False)
		self.mainGrid.EnableDragColSize(True)
		self.mainGrid.SetColLabelSize(30)
		self.mainGrid.SetColLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
		self.mainGrid.EnableDragRowSize(True)
		self.mainGrid.SetRowLabelSize(60)
		self.mainGrid.SetRowLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
		self.mainGrid.SetSelectionMode(wx.grid.Grid.SelectCells)
		self.mainGrid.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_TOP)
		grid_sizer.Add(self.mainGrid, 1, wx.EXPAND, 5)
		main.Add(grid_sizer, 1, wx.EXPAND, 5)

		self.undersideText = wx.StaticText(self, wx.ID_ANY, _(u"\t最大值:\t最小值:\t平均值:\t总和:"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.undersideText.Wrap(-1)
		main.Add(self.undersideText, 0, 0, 5)
		self.SetSizer( main )
		self.Layout()
		self.Centre(wx.BOTH)

		self.tagFile.Bind(wx.EVT_BUTTON, self.toFileTag)
		self.tagJiSuan.Bind(wx.EVT_BUTTON, self.toJiSuanTag)
		self.tagAutom.Bind(wx.EVT_BUTTON, self.toAutomTag)
		self.tagAI.Bind(wx.EVT_BUTTON, self.onAI)
		self.save.Bind(wx.EVT_BUTTON, self.save_)
		self.saveas.Bind(wx.EVT_BUTTON, self.saveas_)
		self.getMax.Bind(wx.EVT_BUTTON, self.getMax_)
		self.getMin.Bind(wx.EVT_BUTTON, self.getMin_)
		self.getAvg.Bind(wx.EVT_BUTTON, self.getAvg_)
		self.getCustomize.Bind(wx.EVT_BUTTON, self.getCus_)
		self.setting.Bind(wx.EVT_BUTTON, self.toSetting)
		self.exit.Bind(wx.EVT_BUTTON, self.exit_)
		self.itemRowInput1.Bind(wx.EVT_TEXT, self.getItemChoice)
		self.mainGrid.Bind(wx.grid.EVT_GRID_RANGE_SELECTED, self.setUndersideText)
		self.sheetChoice.Bind(wx.EVT_CHOICE, self.onSheetChange)
		self.Paiming.Bind(wx.EVT_BUTTON, self.getPaiming_)
		self.RankSubSort.Bind(wx.EVT_BUTTON, self.getRankSubSort)

		self.Tags = [self.tagFile, self.tagJiSuan, self.tagAutom, self.tagAI, self.tagNone1, self.tagNone2, self.tagNone3, self.tagNone4, self.tagNone5, self.tagNone6]
		self.fileTagControl = [self.funcText, self.inputFanc, self.sheetChoice, self.save, self.saveas]
		self.toolTagControl = [self.RankSubSort,self.Paiming,self.getMax, self.getMin, self.getAvg, self.getCustomize, self.customizeInput, self.nameCol, self.nameColInput, self.itemRow, self.itemRowInput1, self.itemChoice]
		self.toFileTag(None)

	def onAI(self, event):
		if not self.ai_chat_frame:
			self.ai_chat_frame = AIChatFrame(self)
			self.ai_chat_frame.Show()
		else:
			self.ai_chat_frame.Raise()

	def update_grid_with_data(self, data):
		self.mainGrid.BeginBatch()
		try:
			data_rows = len(data) if data else 0
			data_cols = max(len(row) for row in data) if data_rows > 0 else 0
			
			target_rows = max(data_rows + 50, 100)
			target_cols = max(data_cols + 20, 50)
			
			current_rows = self.mainGrid.GetNumberRows()
			current_cols = self.mainGrid.GetNumberCols()
			
			row_diff = target_rows - current_rows
			if row_diff > 0: self.mainGrid.AppendRows(row_diff)
			elif row_diff < 0: self.mainGrid.DeleteRows(target_rows, -row_diff)

			col_diff = target_cols - current_cols
			if col_diff > 0: self.mainGrid.AppendCols(col_diff)
			elif col_diff < 0: self.mainGrid.DeleteCols(target_cols, -col_diff)

			self.mainGrid.ClearGrid()
			if data_rows > 0:
				for i, row_data in enumerate(data):
					for j, cell_data in enumerate(row_data):
						self.mainGrid.SetCellValue(i, j, str(cell_data) if cell_data is not None and not pd.isna(cell_data) else "")
		finally:
			self.mainGrid.EndBatch()
			self.mainGrid.ForceRefresh()
	
	def get_data_for_saving(self):
		max_row = -1
		max_col = -1
		for r in range(self.mainGrid.GetNumberRows()):
			for c in range(self.mainGrid.GetNumberCols()):
				if self.mainGrid.GetCellValue(r, c):
					if r > max_row: max_row = r
					if c > max_col: max_col = c
		
		if max_row == -1:
			return []

		data = []
		for r in range(max_row + 1):
			row_data = []
			for c in range(max_col + 1):
				row_data.append(self.mainGrid.GetCellValue(r, c))
			data.append(row_data)
		return data

	def save_( self, event ):
		try:
			data = self.get_data_for_saving()
			if self.file_type == '.xlsx':
				e.write_to_excel(data, self.path)
			elif self.file_type == '.xml':
				e.write_to_xml(data, self.path)
			elif self.file_type == '.json':
				e.write_to_json(data, self.path)
			else:
				wx.MessageBox("不支持的文件类型，无法保存。", "错误", wx.OK | wx.ICON_ERROR)
				return
			wx.MessageBox("保存成功！", "提示", wx.OK | wx.ICON_INFORMATION)
		except Exception as err:
			wx.MessageBox(f"保存文件时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		if event: event.Skip()

	def saveas_( self, event ):
		wildcard = "Excel 文件 (*.xlsx)|*.xlsx|XML 文件 (*.xml)|*.xml|JSON 文件 (*.json)|*.json"
		dialog = wx.FileDialog(self, "另存为", wildcard=wildcard, style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
		if dialog.ShowModal() == wx.ID_OK:
			new_path = dialog.GetPath()
			filter_index = dialog.GetFilterIndex()
			if filter_index == 0 and not new_path.lower().endswith('.xlsx'):
				new_path += '.xlsx'
			elif filter_index == 1 and not new_path.lower().endswith('.xml'):
				new_path += '.xml'
			elif filter_index == 2 and not new_path.lower().endswith('.json'):
				new_path += '.json'

			new_file_type = os.path.splitext(new_path)[1].lower()
			try:
				data = self.get_data_for_saving()
				if new_file_type == '.xlsx':
					e.write_to_excel(data, new_path)
				elif new_file_type == '.xml':
					e.write_to_xml(data, new_path)
				elif new_file_type == '.json':
					e.write_to_json(data, new_path)
				else:
					wx.MessageBox("不支持的文件类型。", "错误", wx.OK | wx.ICON_ERROR)
					return
				self.path = new_path
				self.file_type = new_file_type
				self.SetTitle(_(f"123Excel II - {self.path}"))
				if self.file_type == '.xlsx':
					self.sheetChoice.Enable()
					self.sheetChoice.Clear()
					self.sheetChoice.Append("Sheet")
					self.sheetChoice.SetSelection(0)
				else:
					self.sheetChoice.Clear()
					self.sheetChoice.Append("默认")
					self.sheetChoice.SetSelection(0)
					self.sheetChoice.Disable()
				wx.MessageBox("另存为成功！", "提示", wx.OK | wx.ICON_INFORMATION)
			except Exception as err:
				wx.MessageBox(f"另存为文件时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		dialog.Destroy()
		event.Skip()

	def get_current_data(self):
		try:
			if self.file_type == '.xlsx':
				if self.sheetChoice.GetCount() > 0:
					return e.readExcel(self.path, self.sheetChoice.GetStringSelection())
				return []
			elif self.file_type == '.xml':
				return e.readXml(self.path)
			elif self.file_type == '.json':
				return e.readJson(self.path)
			return []
		except Exception as err:
			wx.MessageBox(f"读取数据时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
			return []

	def toFileTag( self, event ):
		for i in self.fileTagControl: i.Show()
		for i in self.toolTagControl: i.Hide()
		self.Layout()
		if event: event.Skip()

	def toJiSuanTag( self, event ):
		for i in self.fileTagControl: i.Hide()
		for i in self.toolTagControl: i.Show()
		self.getItemChoice(None)
		self.Layout()
		if event: event.Skip()

	def toAutomTag( self, event ):
		wx.MessageBox("此功能尚未实现。", "提示", wx.OK | wx.ICON_INFORMATION)
		event.Skip()

	def toSetting( self, event ):
		options_frame = OptionsFrame(self, title='配置选项', size=(500, 400))
		options_frame.Show()
		event.Skip()

	def exit_( self, event ):
		self.Close()
		event.Skip()
		
	def onSheetChange(self, event):
		if self.file_type != '.xlsx':
			event.Skip()
			return
		try:
			sheet_name = self.sheetChoice.GetString(self.sheetChoice.GetSelection())
			data = e.readExcel(self.path, sheet_name)
			self.update_grid_with_data(data)
		except Exception as err:
			wx.MessageBox(f"加载工作表时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		event.Skip()

	def setUndersideText( self, event ):
		selected_cells = []
		if self.mainGrid.GetSelectionBlockTopLeft():
			for top_left, bottom_right in zip(self.mainGrid.GetSelectionBlockTopLeft(), self.mainGrid.GetSelectionBlockBottomRight()):
				for row in range(top_left.Row, bottom_right.Row + 1):
					for col in range(top_left.Col, bottom_right.Col + 1):
						cell_value = self.mainGrid.GetCellValue(row, col)
						if cell_value:
							try:
								selected_cells.append(float(cell_value))
							except ValueError:
								pass
		if selected_cells:
			max_val, min_val, sum_val = max(selected_cells), min(selected_cells), sum(selected_cells)
			avg_val = sum_val / len(selected_cells)
			self.undersideText.SetLabel(f"最大值:{max_val:.2f} 最小值:{min_val:.2f} 平均值:{avg_val:.2f} 总和:{sum_val:.2f}")
		else:
			self.undersideText.SetLabel("最大值:无 最小值:无 平均值:无 总和:无")
		event.Skip()

	def getItemChoice( self, event ):
		try:
			item_row_val = self.itemRowInput1.GetValue()
			if not item_row_val.isdigit() or int(item_row_val) <= 0:
				return
			data = self.get_current_data()
			if not data:
				return
			items = e.getItems(data, int(item_row_val))
			current_selection = self.itemChoice.GetStringSelection()
			self.itemChoice.Clear()
			if items:
				self.itemChoice.AppendItems([str(i) for i in items if i])
				if current_selection in self.itemChoice.GetStrings():
					self.itemChoice.SetStringSelection(current_selection)
				else:
					self.itemChoice.SetSelection(0)
		except Exception as err:
			print(f"更新项目列表时出错: {err}")
		if event: event.Skip()

	def _get_common_calculation_data(self):
		data = self.get_current_data()
		if not all([data, self.itemRowInput1.GetValue().isdigit(), self.nameColInput.GetValue().isdigit(), self.itemChoice.GetStringSelection()]):
			wx.MessageBox("请确保'项目行'和'名称列'已正确填写，并已选择一个项目。", "输入错误", wx.OK | wx.ICON_ERROR)
			return None
		item_row = int(self.itemRowInput1.GetValue())
		name_col = int(self.nameColInput.GetValue())
		item_name = self.itemChoice.GetStringSelection()
		names = e.getNames(data, item_row, name_col)
		values = e.getValue(data, item_row, name_col, item_name)
		return e.getNameValueDict(names, values)

	def getMax_( self, event ):
		try:
			NameValueDict = self._get_common_calculation_data()
			if NameValueDict is None: return
			max_val = e.getMaxValue(NameValueDict)
			max_names = e.getMaxNames(NameValueDict, max_val)
			outputWindow(parent=self, Message=max_names).Show()
		except Exception as err:
			wx.MessageBox(f"计算最大值时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		event.Skip()

	def getMin_( self, event ):
		try:
			NameValueDict = self._get_common_calculation_data()
			if NameValueDict is None: return
			min_val = e.getMinValue(NameValueDict)
			min_names = e.getMinNames(NameValueDict, min_val)
			outputWindow(parent=self, Message=min_names).Show()
		except Exception as err:
			wx.MessageBox(f"计算最小值时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		event.Skip()

	def getAvg_( self, event ):
		try:
			NameValueDict = self._get_common_calculation_data()
			if NameValueDict is None: return
			avg = e.getAverageValue(NameValueDict)
			outputWindow(parent=self, Message=avg).Show()
		except Exception as err:
			wx.MessageBox(f"计算平均值时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		event.Skip()

	def getCus_( self, event ):
		try:
			if not self.customizeInput.GetValue():
				wx.MessageBox("自定义准则不能为空。", "输入错误", wx.OK | wx.ICON_ERROR)
				return
			NameValueDict = self._get_common_calculation_data()
			if NameValueDict is None: return
			cus = e.getCustomizeValue(NameValueDict, self.customizeInput.GetValue())
			outputWindow(parent=self, Message=cus).Show()
		except Exception as err:
			wx.MessageBox(f"执行自定义计算时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		event.Skip()
	
	def getPaiming_( self, event ):
		try:
			file = self.path
			sub = self.itemChoice.GetStringSelection()
			paiming = e.AddPaiMing(file,sub)
			wx.MessageBox(f"排名完毕，保存在同一路径下的{str(file)}_{str(sub)}排名.xlsx 文件","提示", wx.OK | wx.ICON_INFORMATION)
		except Exception as err:
			wx.MessageBox(f"计算排名时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		event.Skip()
	
	def getRankSubSort(self, event ):
		try:
			file = self.path
			sub = self.itemChoice.GetStringSelection()
			paiming = e.RankSubSort(file,sub)
			wx.MessageBox(f"排名完毕，保存在同一路径下的{str(file)}_{str(sub)}排名排序.xlsx 文件","提示", wx.OK | wx.ICON_INFORMATION)
		except Exception as err:
			wx.MessageBox(f"计算排名时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		event.Skip()

class outputWindow ( wx.Frame ):
	def __init__(self, parent ,Message):
		wx.Frame.__init__ (self, parent, id = wx.ID_ANY, title = "运算结果", pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)
		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour( 239, 235, 235 ))
		main = wx.BoxSizer(wx.VERTICAL)
		self.m_listBox1 = wx.ListBox(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, Message, wx.LB_MULTIPLE)
		main.Add(self.m_listBox1, 1, wx.ALL|wx.EXPAND, 5)
		self.copy = wx.Button(self, wx.ID_ANY, _(u"复制选中项"), wx.DefaultPosition, wx.DefaultSize, 0)
		main.Add(self.copy, 0, wx.ALL|wx.EXPAND, 5)
		self.SetSizer( main )
		self.Layout()
		self.Centre(wx.BOTH)
		self.copy.Bind(wx.EVT_BUTTON, self.copy_)

	def copy_( self, event ):
		selections = self.m_listBox1.GetSelections()
		if not selections:
			wx.MessageBox("没有选中任何项。", "提示")
			return
		items = [self.m_listBox1.GetString(i) for i in selections]
		all_items_text = '\n'.join(items)
		if wx.TheClipboard.Open():
			wx.TheClipboard.SetData(wx.TextDataObject(all_items_text))
			wx.TheClipboard.Close()
			wx.MessageBox(f"{len(items)} 项已复制到剪贴板！", "提示", wx.OK | wx.ICON_INFORMATION)
		else:
			wx.MessageBox("无法打开剪贴板", "错误", wx.OK | wx.ICON_ERROR)
		event.Skip()

# --------------------------------------------------------------------------
#  表格预览窗口
# --------------------------------------------------------------------------
class PreviewFrame(wx.Frame):
	def __init__(self, parent, data):
		wx.Frame.__init__(self, parent, id=wx.ID_ANY, title="AI修改预览", pos=wx.DefaultPosition, size=wx.Size(800, 600), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)
		self.parent = parent
		self.data = data
		
		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour(255, 255, 255))

		main_sizer = wx.BoxSizer(wx.VERTICAL)
		
		# 添加说明文本
		self.info_text = wx.StaticText(self, wx.ID_ANY, "以下是AI建议的修改预览，请确认是否应用这些更改：", wx.DefaultPosition, wx.DefaultSize, 0)
		self.info_text.Wrap(-1)
		main_sizer.Add(self.info_text, 0, wx.ALL | wx.EXPAND, 5)
		
		# 创建预览表格
		self.preview_grid = wx.grid.Grid(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)
		self.preview_grid.CreateGrid(0, 0)
		self.update_grid_with_data(data)
		
		# 设置表格属性
		self.preview_grid.EnableEditing(False)
		self.preview_grid.EnableGridLines(True)
		self.preview_grid.EnableDragGridSize(False)
		self.preview_grid.SetMargins(0, 0)
		self.preview_grid.EnableDragColMove(False)
		self.preview_grid.EnableDragColSize(True)
		self.preview_grid.SetColLabelSize(30)
		self.preview_grid.SetColLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
		self.preview_grid.EnableDragRowSize(True)
		self.preview_grid.SetRowLabelSize(60)
		self.preview_grid.SetRowLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
		self.preview_grid.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_TOP)
		
		main_sizer.Add(self.preview_grid, 1, wx.ALL | wx.EXPAND, 5)
		
		# 添加按钮
		button_sizer = wx.BoxSizer(wx.HORIZONTAL)
		self.cancel_button = wx.Button(self, wx.ID_ANY, "取消", wx.DefaultPosition, wx.DefaultSize, 0)
		self.apply_button = wx.Button(self, wx.ID_ANY, "应用更改", wx.DefaultPosition, wx.DefaultSize, 0)
		button_sizer.Add(self.cancel_button, 0, wx.ALL, 5)
		button_sizer.AddStretchSpacer()
		button_sizer.Add(self.apply_button, 0, wx.ALL, 5)
		main_sizer.Add(button_sizer, 0, wx.ALL | wx.EXPAND, 5)
		
		self.SetSizer(main_sizer)
		self.Layout()
		self.Centre(wx.BOTH)
		
		# 绑定事件
		self.cancel_button.Bind(wx.EVT_BUTTON, self.on_cancel)
		self.apply_button.Bind(wx.EVT_BUTTON, self.on_apply)
		
	def update_grid_with_data(self, data):
		self.preview_grid.BeginBatch()
		try:
			data_rows = len(data) if data else 0
			data_cols = max(len(row) for row in data) if data_rows > 0 else 0
			
			target_rows = max(data_rows + 10, 20)
			target_cols = max(data_cols + 5, 10)
			
			current_rows = self.preview_grid.GetNumberRows()
			current_cols = self.preview_grid.GetNumberCols()
			
			row_diff = target_rows - current_rows
			if row_diff > 0: self.preview_grid.AppendRows(row_diff)
			elif row_diff < 0: self.preview_grid.DeleteRows(target_rows, -row_diff)

			col_diff = target_cols - current_cols
			if col_diff > 0: self.preview_grid.AppendCols(col_diff)
			elif col_diff < 0: self.preview_grid.DeleteCols(target_cols, -col_diff)

			self.preview_grid.ClearGrid()
			if data_rows > 0:
				for i, row_data in enumerate(data):
					for j, cell_data in enumerate(row_data):
						self.preview_grid.SetCellValue(i, j, str(cell_data) if cell_data is not None and not pd.isna(cell_data) else "")
		finally:
			self.preview_grid.EndBatch()
			self.preview_grid.ForceRefresh()
			
	def on_cancel(self, event):
		self.Close()
		
	def on_apply(self, event):
		# 调用父窗口的apply_ai_changes方法应用更改
		self.parent.apply_ai_changes(self.data)
		self.Close()

if __name__ == '__main__':
	app = wx.App(False)
	test_file = "test.xlsx"
	if not os.path.exists(test_file):
		pd.DataFrame([["姓名", "分数"], ["张三", 95], ["李四", 88]]).to_excel(test_file, index=False, header=False)
	excelPage = excelPage_(None, test_file)
	excelPage.Show()
	app.MainLoop()
