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
_ = gettext.gettext

#--------------------------------------------------------------------------
#  Class MyFrame1
#---------------------------------------------------------------------------

class MyFrame1 ( wx.Frame ):

	def __init__(self, parent):
		wx.Frame.__init__ (self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 430,460 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)

		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour( 239, 235, 235 ))

		main = wx.BoxSizer(wx.HORIZONTAL)

		self.m_Table1 = wx.grid.Grid(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

		# Grid
		self.m_Table1.CreateGrid(50, 3)
		self.m_Table1.EnableEditing(True)
		self.m_Table1.EnableGridLines(True)
		self.m_Table1.EnableDragGridSize(False)
		self.m_Table1.SetMargins(0, 0)

		# Columns
		self.m_Table1.EnableDragColMove(False)
		self.m_Table1.EnableDragColSize(True)
		self.m_Table1.SetColLabelSize(30)
		self.m_Table1.SetColLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)

		# Rows
		self.m_Table1.EnableDragRowSize(True)
		self.m_Table1.SetRowLabelSize(60)
		self.m_Table1.SetRowLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)

		# Label Appearance

		# Cell Defaults
		self.m_Table1.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_TOP)
		main.Add(self.m_Table1, 0, wx.ALL|wx.EXPAND, 5)

		bar = wx.BoxSizer(wx.VERTICAL)

		self.reset = wx.Button(self, wx.ID_ANY, _(u"恢复默认"), wx.DefaultPosition, wx.DefaultSize, 0)
		bar.Add(self.reset, 0, wx.ALL, 5)

		self.check = wx.Button(self, wx.ID_ANY, _(u"类型检查"), wx.DefaultPosition, wx.DefaultSize, 0)
		bar.Add(self.check, 0, wx.ALL, 5)

		self.open = wx.Button(self, wx.ID_ANY, _(u"打开json文件"), wx.DefaultPosition, wx.DefaultSize, 0)
		bar.Add(self.open, 0, wx.ALL, 5)


		main.Add(bar, 1, wx.EXPAND, 5)


		self.SetSizer( main )
		self.Layout()

		self.Centre(wx.BOTH)

	def __del__( self ):
		pass


if __name__ == "__main__":
	app = wx.App(False)
	frame = MyFrame1(None)
	frame.Show(True)
	app.MainLoop()