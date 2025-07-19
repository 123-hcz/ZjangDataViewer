# excelPage.py
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

_ = gettext.gettext

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

		data = []
		try:
			if self.file_type == '.xlsx':
				if self.sheetChoice.GetCount() > 0:
					sheet_name = self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					data = e.readExcel(path, sheet_name)
			elif self.file_type == '.xml':
				data = e.readXml(path)
			else:
				wx.MessageBox("不支持的文件类型！", "错误", wx.OK | wx.ICON_ERROR)

			if data:
				rows, cols = len(data), max(len(row) for row in data) if data else 0
				self.mainGrid.CreateGrid(rows + 100, cols + 100)
				for i, row in enumerate(data):
					for j, cell in enumerate(row):
						self.mainGrid.SetCellValue(i, j, str(cell) if not pd.isna(cell) else "")
			else:
				self.mainGrid.CreateGrid(100, 100)
		except Exception as err:
			self.mainGrid.CreateGrid(100, 100)
			wx.MessageBox(f"打开文件时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)

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

		self.Tags = [self.tagFile, self.tagJiSuan, self.tagAutom, self.tagNone1, self.tagNone2, self.tagNone3, self.tagNone4, self.tagNone5, self.tagNone6]
		self.fileTagControl = [self.funcText, self.inputFanc, self.sheetChoice, self.save, self.saveas]
		self.toolTagControl = [self.getMax, self.getMin, self.getAvg, self.getCustomize, self.customizeInput, self.nameCol, self.nameColInput, self.itemRow, self.itemRowInput1, self.itemChoice]
		self.toFileTag(None)

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
			else:
				wx.MessageBox("不支持的文件类型，无法保存。", "错误", wx.OK | wx.ICON_ERROR)
				return
			wx.MessageBox("保存成功！", "提示", wx.OK | wx.ICON_INFORMATION)
		except Exception as err:
			wx.MessageBox(f"保存文件时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)
		if event: event.Skip()

	def saveas_( self, event ):
		wildcard = "Excel 文件 (*.xlsx)|*.xlsx|XML 文件 (*.xml)|*.xml"
		dialog = wx.FileDialog(self, "另存为", wildcard=wildcard, style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
		if dialog.ShowModal() == wx.ID_OK:
			new_path = dialog.GetPath()
			filter_index = dialog.GetFilterIndex()
			if filter_index == 0 and not new_path.lower().endswith('.xlsx'):
				new_path += '.xlsx'
			elif filter_index == 1 and not new_path.lower().endswith('.xml'):
				new_path += '.xml'
			new_file_type = os.path.splitext(new_path)[1].lower()
			try:
				data = self.get_data_for_saving()
				if new_file_type == '.xlsx':
					e.write_to_excel(data, new_path)
				elif new_file_type == '.xml':
					e.write_to_xml(data, new_path)
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
		wx.MessageBox("此功能尚未实现。", "提示", wx.OK | wx.ICON_INFORMATION)
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
			self.mainGrid.ClearGrid()
			if self.mainGrid.GetNumberRows() > 0:
				self.mainGrid.DeleteRows(0, self.mainGrid.GetNumberRows())
			if self.mainGrid.GetNumberCols() > 0:
				self.mainGrid.DeleteCols(0, self.mainGrid.GetNumberCols())

			if data:
				rows, cols = len(data), max(len(row) for row in data) if data else 0
				self.mainGrid.AppendRows(rows + 100)
				self.mainGrid.AppendCols(cols + 100)
				for i, row in enumerate(data):
					for j, cell in enumerate(row):
						self.mainGrid.SetCellValue(i, j, str(cell) if not pd.isna(cell) else '')
			else:
				self.mainGrid.AppendRows(100)
				self.mainGrid.AppendCols(100)
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

if __name__ == '__main__':
	app = wx.App(False)
	excelPage = excelPage_(None, "D:\\下载\\8年级录分(1).xlsx")
	excelPage.Show()
	app.MainLoop()