# -*- coding: utf-8 -*-

#--------------------------------------------------------------------------
# Python code generated with wxFormBuilder (version 3.9.0 Jun 14 2020)
# http://www.wxformbuilder.org/
#
# PLEASE DO *NOT* EDIT THIS FILE!
#--------------------------------------------------------------------------

import wx
import wx.xrc
import wx.dataview

import gettext
_ = gettext.gettext

#--------------------------------------------------------------------------
#  Class MyFrame1
#---------------------------------------------------------------------------

class MyFrame1 ( wx.Frame ):

	def __init__(self, parent):
		wx.Frame.__init__ (self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 723,579 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)

		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour( 239, 235, 235 ))

		main = wx.BoxSizer(wx.VERTICAL)

		headBar = wx.BoxSizer(wx.VERTICAL)

		tags = wx.BoxSizer(wx.HORIZONTAL)

		self.file = wx.Button(self, wx.ID_ANY, _(u"文件"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.file, 0, 0, 5)

		self.find = wx.Button(self, wx.ID_ANY, _(u"查找"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.find, 0, 0, 5)

		self.NoneTag = wx.Button(self, wx.ID_ANY, _(u"None"), wx.DefaultPosition, wx.DefaultSize, 0)
		tags.Add(self.NoneTag, 0, 0, 5)


		headBar.Add(tags, 0, wx.EXPAND, 5)

		tools = wx.BoxSizer(wx.HORIZONTAL)

		self.save = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.save.SetBitmap(wx.Bitmap( u"image/icon_packs/classic/save.png", wx.BITMAP_TYPE_ANY ))
		self.save.SetBitmapPressed(wx.Bitmap( u"image/icon_packs/classic/save1.png", wx.BITMAP_TYPE_ANY ))
		tools.Add(self.save, 0, 0, 5)

		self.saveas = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.saveas.SetBitmap(wx.Bitmap( u"image/icon_packs/classic/saveas.png", wx.BITMAP_TYPE_ANY ))
		self.saveas.SetBitmapPressed(wx.Bitmap( u"image/icon_packs/classic/saveas1.png", wx.BITMAP_TYPE_ANY ))
		tools.Add(self.saveas, 0, 0, 5)


		headBar.Add(tools, 0, wx.EXPAND, 5)

		self.tree = wx.dataview.DataViewTreeCtrl(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)
		headBar.Add(self.tree, 1, wx.ALL|wx.EXPAND, 5)

		self.underBar = wx.StaticText(self, wx.ID_ANY, _(u"MyLabel"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.underBar.Wrap(-1)

		headBar.Add(self.underBar, 0, wx.ALL, 5)


		main.Add(headBar, 1, wx.EXPAND, 5)


		self.SetSizer( main )
		self.Layout()

		self.Centre(wx.BOTH)

	def __del__( self ):
		pass

if __name__ == "__main__":
	app = wx.App(False)
	frame = MyFrame1(None)
	frame.Show()
	app.MainLoop()
