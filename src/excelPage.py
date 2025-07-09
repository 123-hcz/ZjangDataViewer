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
# import openFile
import gettext

_ = gettext.gettext
import excel as e
import pandas as pd
#--------------------------------------------------------------------------
#  Class excelPage
#---------------------------------------------------------------------------


class excelPage_ ( wx.Frame ):

	def __init__(self, parent,path):
		global fileTagControl,toolTagControl,Tags

		wx.Frame.__init__ (self, parent,
						   id = wx.ID_ANY, title = _(f"123Excel II - {path}"),
						   pos = wx.DefaultPosition, size = wx.Size( 1000,600 ),
						   style = wx.DEFAULT_FRAME_STYLE|wx.BORDER_NONE|wx.TAB_TRAVERSAL
							)
		self.path = path
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

		self.save = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.save.SetBitmap(wx.Bitmap( u"../src/image/icon_packs/classic/save.png", wx.BITMAP_TYPE_ANY ))
		self.save.SetBitmapPressed(wx.Bitmap( u"../src/image/icon_packs/classic/save1.png", wx.BITMAP_TYPE_ANY ))
		self.save.SetToolTip(_(u"保存文件"))

		tools.Add(self.save, 0, 0, 5)

		self.saveas = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.saveas.SetBitmap(wx.Bitmap( u"../src/image/icon_packs/classic/saveas.png", wx.BITMAP_TYPE_ANY ))
		self.saveas.SetBitmapPressed(wx.Bitmap( u"../src/image/icon_packs/classic/saveas1.png", wx.BITMAP_TYPE_ANY ))
		self.saveas.SetToolTip(_(u"另存为"))

		tools.Add(self.saveas, 0, 0, 5)

		self.getMax = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.getMax.SetBitmap(wx.Bitmap( u"../src/image/icon_packs/classic/max.png", wx.BITMAP_TYPE_ANY ))
		self.getMax.SetBitmapPressed(wx.Bitmap( u"../src/image/icon_packs/classic/max1.png", wx.BITMAP_TYPE_ANY ))
		self.getMax.Hide()
		self.getMax.SetToolTip(_(u"最大值"))

		tools.Add(self.getMax, 0, 0, 5)

		self.getMin = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.getMin.SetBitmap(wx.Bitmap( u"../src/image/icon_packs/classic/min.png", wx.BITMAP_TYPE_ANY ))
		self.getMin.SetBitmapPressed(wx.Bitmap( u"../src/image/icon_packs/classic/min1.png", wx.BITMAP_TYPE_ANY ))
		self.getMin.Hide()
		self.getMin.SetToolTip(_(u"最小值"))

		tools.Add(self.getMin, 0, 0, 5)

		self.getAvg = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.getAvg.SetBitmap(wx.Bitmap( u"../src/image/icon_packs/classic/avg.png", wx.BITMAP_TYPE_ANY ))
		self.getAvg.SetBitmapPressed(wx.Bitmap( u"../src/image/icon_packs/classic/avg1.png", wx.BITMAP_TYPE_ANY ))
		self.getAvg.Hide()
		self.getAvg.SetToolTip(_(u"平均值"))

		tools.Add(self.getAvg, 0, 0, 5)

		self.getCustomize = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.getCustomize.SetBitmap(wx.Bitmap( u"../src/image/icon_packs/classic/customize.png", wx.BITMAP_TYPE_ANY ))
		self.getCustomize.SetBitmapPressed(wx.Bitmap( u"../src/image/icon_packs/classic/customize1.png", wx.BITMAP_TYPE_ANY ))
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
		excel_file = pd.ExcelFile(path)
		sheets = excel_file.sheet_names
		for sheet in sheets:
			self.sheetChoice.Append(sheet)
		self.sheetChoice.SetSelection(0)
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

		self.itemRowInput1 = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
		self.itemRowInput1.Hide()
		self.itemRowInput1.SetToolTip(_(u"填写项目的列\n例如:写科目的列"))
		self.itemRowInput1.SetValue(str(1))

		values.Add(self.itemRowInput1, 0, 0, 5)
		try:
			itemChoiceChoices = e.getItems(
				e.readExcel(path,self.sheetChoice.GetStringSelection()),
				int(self.itemRowInput1.GetValue())
			)
			itemChoiceChoices = [str(item) for item in itemChoiceChoices]
		except:
			itemChoiceChoices = []

		self.itemChoice = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, itemChoiceChoices, 0)
		self.itemChoice.SetSelection(0)
		self.itemChoice.Hide()
		self.itemChoice.SetToolTip(_(u"项目"))

		values.Add(self.itemChoice, 0, wx.LEFT, 5)


		mainTool.Add(values, 1, wx.EXPAND, 5)


		toolBar.Add(mainTool, 1, wx.EXPAND, 5)

		others = wx.BoxSizer(wx.HORIZONTAL)

		self.setting = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.setting.SetBitmap(wx.Bitmap( u"../src/image/icon_packs/classic/setting.png", wx.BITMAP_TYPE_ANY ))
		self.setting.SetBitmapPressed(wx.Bitmap( u"../src/image/icon_packs/classic/setting1.png", wx.BITMAP_TYPE_ANY ))
		self.setting.SetToolTip(_(u"设置"))

		others.Add(self.setting, 0, wx.ALL, 5)

		self.exit = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.exit.SetBitmap(wx.Bitmap( u"../src/image/icon_packs/classic/exit.png", wx.BITMAP_TYPE_ANY ))
		self.exit.SetBitmapPressed(wx.Bitmap( u"../src/image/icon_packs/classic/exit1.png", wx.BITMAP_TYPE_ANY ))
		self.exit.SetToolTip(_(u"退出"))

		others.Add(self.exit, 0, wx.ALL, 5)


		toolBar.Add(others, 0, 0, 5)


		toolBar.AddStretchSpacer() #.Add((0, 30), 1, wx.EXPAND, 5)


		main.Add(toolBar, 0, 0, 5)

		grid = wx.BoxSizer(wx.HORIZONTAL)

		self.mainGrid = wx.grid.Grid(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)


		# Grid
		try:
			data = e.readExcel(path, self.sheetChoice.GetString(self.sheetChoice.GetSelection()))
			self.mainGrid.CreateGrid(len(data) + 100, len(data[0]) + 100)
			for i in range(len(data)):
				for j in range(len(data[0])):
					if str(data[i][j]) != "nan":
						self.mainGrid.SetCellValue(i, j, str(data[i][j]))
					else:
						self.mainGrid.SetCellValue(i, j,str(""))
		except:
			self.mainGrid.CreateGrid(100, 100)
			self.mainGrid.EnableEditing(True)
			self.mainGrid.EnableGridLines(True)
			self.mainGrid.EnableDragGridSize(False)
			self.mainGrid.SetMargins(0, 0)
		self.Refresh()
		# Columns
		self.mainGrid.EnableDragColMove(False)
		self.mainGrid.EnableDragColSize(True)
		self.mainGrid.SetColLabelSize(30)
		self.mainGrid.SetColLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)

		# Rows
		self.mainGrid.EnableDragRowSize(True)
		self.mainGrid.SetRowLabelSize(60)
		self.mainGrid.SetRowLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)

		# Label Appearance
		self.mainGrid.SetSelectionMode(wx.grid.Grid.SelectCells)  # 支持单元格选择
		# Cell Defaults
		self.mainGrid.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_TOP)
		grid.Add(self.mainGrid, 1, wx.EXPAND, 5)


		main.Add(grid, 1, wx.EXPAND, 5)

		self.undersideText = wx.StaticText(self, wx.ID_ANY, _(u"\t最大值:\t最小值:\t平均值:\t总和:"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.undersideText.Wrap(-1)

		main.Add(self.undersideText, 0, 0, 5)


		self.SetSizer( main )
		self.Layout()

		self.Centre(wx.BOTH)

		# Connect Events
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

		#Tags
		Tags = [
			self.tagFile,
			self.tagJiSuan,
			self.tagAutom,
			self.tagNone1,
			self.tagNone2,
			self.tagNone3,
			self.tagNone4,
			self.tagNone5,
			self.tagNone6,
		]
		#fileTag
		fileTagControl = [
			self.funcText,
			self.inputFanc,
			self.sheetChoice,
			self.funcText,
			self.inputFanc,
			self.sheetChoice,
			self.save,
			self.saveas

		]


		#toolTag
		toolTagControl = [
			self.getMax,
			self.getMin,
			self.getAvg,
			self.getCustomize,
			self.customizeInput,
			self.nameCol,
			self.nameColInput,
			self.itemRow,
			self.itemRowInput1,
			self.itemChoice,

		]

		for i in Tags:
			i.Show()
		for i in fileTagControl:
			i.Show()
		for i in toolTagControl:
			i.Hide()


		# Connect Events
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

	# Virtual event handlers, overide them in your derived class

	def toFileTag( self, event ):
		self.Hide()
		for i in Tags:
			i.Show()
		for i in fileTagControl:
			i.Show()
		for i in toolTagControl:
			i.Hide()
		self.Show()
		self.Refresh()
		event.Skip()

	def toJiSuanTag( self, event ):
		self.Hide()
		for i in Tags:
			i.Show()
		for i in fileTagControl:
			i.Hide()
		for i in toolTagControl:
			i.Show()
		self.Show()

		self.Refresh()
		event.Skip()

	def toAutomTag( self, event ):
		event.Skip()

	def save_( self, event ):
		try:

			data = []
			for i in range(self.mainGrid.GetNumberRows()):
				row = []
				for j in range(self.mainGrid.GetNumberCols()):
					cell_value = self.mainGrid.GetCellValue(i, j)
					row.append(cell_value)
				data.append(row)

			e.write_to_excel(data, self.path)
			wx.MessageBox("保存成功！", "提示", wx.OK | wx.ICON_INFORMATION)


		except Exception as e_:
			raise e_
		event.Skip()

	def saveas_( self, event ):
		event.Skip()

	def setUndersideText( self, event ):
		selected_cells = []
		for top_left, bottom_right in zip(self.mainGrid.GetSelectionBlockTopLeft(),
										  self.mainGrid.GetSelectionBlockBottomRight()):
			rows = range(top_left.Row, bottom_right.Row + 1)
			cols = range(top_left.Col, bottom_right.Col + 1)
			print(rows,cols)

			for row in rows:
				for col in cols:
					cell_value = self.mainGrid.GetCellValue(row, col)
					if cell_value:
						try:
							selected_cells.append(float(cell_value))
						except ValueError:
							pass  # 忽略非数值内容

		if selected_cells:
			max_val = max(selected_cells)
			min_val = min(selected_cells)
			avg_val = sum(selected_cells) / len(selected_cells)
			sum_val = sum(selected_cells)
			self.undersideText.SetLabel(f"最大值:{max_val} 最小值:{min_val} 平均值:{avg_val} 总和:{sum_val}")
		else:
			self.undersideText.SetLabel(f"最大值:无 最小值:无 平均值:无 总和:无")
			print(selected_cells)
		self.Refresh()


		event.Skip()
	def getItemChoice( self, event ):
		self.itemChoice.Clear()
		for i in e.getItems(
			e.readExcel(
				self.path,
				self.sheetChoice.GetString(self.sheetChoice.GetSelection())
			),
			int(self.itemRowInput1.GetValue())
		):
			try:
				self.itemChoice.Append(str(i))
			except TypeError as e_:
				raise e_
		self.itemChoice.SetSelection(0)
		self.Refresh()
		event.Skip()
	def getMax_( self, event):

		NameValueDict = e.getNameValueDict(
				e.getNames(
					e.readExcel(
						self.path,
						self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					),
					int(self.itemRowInput1.GetValue()),
					int(self.nameColInput.GetValue())
				),
				e.getValue(
					e.readExcel(
						self.path,
						self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					),
					int(self.itemRowInput1.GetValue()),
					int(self.nameColInput.GetValue()),
					self.itemChoice.GetString(self.itemChoice.GetSelection())
				)
			)
		max = e.getMaxValue(
			NameValueDict
		)
		max = e.getMaxNames(NameValueDict,max)
		MessageWindow = outputWindow(parent=None ,Message=max)
		MessageWindow.Show()

		event.Skip()

	def getMin_( self, event ):
		NameValueDict = e.getNameValueDict(
				e.getNames(
					e.readExcel(
						self.path,
						self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					),
					int(self.itemRowInput1.GetValue()),
					int(self.nameColInput.GetValue())
				),
				e.getValue(
					e.readExcel(
						self.path,
						self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					),
					int(self.itemRowInput1.GetValue()),
					int(self.nameColInput.GetValue()),
					self.itemChoice.GetString(self.itemChoice.GetSelection())
				)
			)
		min = e.getMinValue(
			NameValueDict
		)
		min = e.getMinNames(
			NameValueDict,
			min
		)
		MessageWindow = outputWindow(parent=None ,Message=min)
		MessageWindow.Show()

		event.Skip()

	def getAvg_( self, event ):
		NameValueDict = e.getNameValueDict(
				e.getNames(
					e.readExcel(
						self.path,
						self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					),
					int(self.itemRowInput1.GetValue()),
					int(self.nameColInput.GetValue())
				),
				e.getValue(
					e.readExcel(
						self.path,
						self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					),
					int(self.itemRowInput1.GetValue()),
					int(self.nameColInput.GetValue()),
					self.itemChoice.GetString(self.itemChoice.GetSelection())
				)
			)
		avg = e.getAverageValue(
			NameValueDict
		)
		MessageWindow = outputWindow(parent=None ,Message=avg)
		MessageWindow.Show()
		event.Skip()

	def getCus_( self, event ):
		NameValueDict = e.getNameValueDict(
				e.getNames(
					e.readExcel(
						self.path,
						self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					),
					int(self.itemRowInput1.GetValue()),
					int(self.nameColInput.GetValue())
				),
				e.getValue(
					e.readExcel(
						self.path,
						self.sheetChoice.GetString(self.sheetChoice.GetSelection())
					),
					int(self.itemRowInput1.GetValue()),
					int(self.nameColInput.GetValue()),
					self.itemChoice.GetString(self.itemChoice.GetSelection())
				))
		cus = e.getCustomizeValue(
			NameValueDict,
			e.getCustomizeEquation(
				self.customizeInput.GetValue(),
				"float(Values[i])"
			)
		)
		MessageWindow = outputWindow(parent=None ,Message=cus)
		MessageWindow.Show()

		event.Skip()

	def toSetting( self, event ):
		event.Skip()

	def exit_( self, event ):
		self.Close()
		event.Skip()


#--------------------------------------------------------------------------
#  Class outputWindow
#---------------------------------------------------------------------------

class outputWindow ( wx.Frame ):

	def __init__(self, parent ,Message):
		wx.Frame.__init__ (self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 800,460 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)

		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour( 239, 235, 235 ))

		main = wx.BoxSizer(wx.VERTICAL)

		m_listBox1Choices = Message
		self.m_listBox1 = wx.ListBox(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_listBox1Choices, 0)
		main.Add(self.m_listBox1, 1, wx.ALL|wx.EXPAND, 5)

		self.copy = wx.Button(self, wx.ID_ANY, _(u"复制"), wx.DefaultPosition, wx.DefaultSize, 0)
		main.Add(self.copy, 0, wx.ALL|wx.EXPAND, 5)


		self.SetSizer( main )
		self.Layout()

		self.Centre(wx.BOTH)

		# Connect Events
		self.copy.Bind(wx.EVT_BUTTON, self.copy_)

	def __del__( self ):
		# Disconnect Events
		self.copy.Unbind(wx.EVT_BUTTON, None)


	# Virtual event handlers, overide them in your derived class
	def copy_( self, event ):
		# 获取所有选项
		items = self.m_listBox1.GetStrings()
		if not items:
			wx.MessageBox("列表为空，无法复制。", "提示")
			return

		# 转换为换行分隔的字符串
		all_items_text = '\n'.join(items)

		# 复制到剪贴板
		if wx.TheClipboard.Open():
			wx.TheClipboard.SetData(wx.TextDataObject(all_items_text))
			wx.TheClipboard.Close()
			wx.MessageBox("所有项已复制到剪贴板！", "提示")
		else:
			wx.MessageBox("无法打开剪贴板", "错误")

		event.Skip()


if __name__ == '__main__':
	app = wx.App()
	excelPage = excelPage_(None,"D:\下载\8年级录分(1).xlsx")
	excelPage.Show()
	app.MainLoop()


