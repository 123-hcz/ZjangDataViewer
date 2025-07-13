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
import json

import gettext
_ = gettext.gettext

#--------------------------------------------------------------------------
#  Class MyFrame1
#---------------------------------------------------------------------------
def load_json(file_path):
	with open(file_path, 'r', encoding='utf-8') as file:
		return json.load(file)

class MyFrame1 ( wx.Frame ):


	def add_json_to_tree(self, tree, parent_item, data):
		if isinstance(data, dict):
			for key, value in data.items():
				item = tree.AppendItem(parent_item, str(key))
				self.add_json_to_tree(tree, item, value)
		elif isinstance(data, list):
			for index, value in enumerate(data):
				item = tree.AppendItem(parent_item, f"[{index}]")
				self.add_json_to_tree(tree, item, value)
		else:
			tree.AppendItem(parent_item, str(data))

	def __init__(self, parent ,path = '_json/options.json'):
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

		self.save.SetBitmap(wx.Bitmap( u"./image/icon_packs/classic/save.png", wx.BITMAP_TYPE_ANY ))
		self.save.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/save1.png", wx.BITMAP_TYPE_ANY ))
		tools.Add(self.save, 0, 0, 5)

		self.saveas = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.saveas.SetBitmap(wx.Bitmap( u"./image/icon_packs/classic/saveas.png", wx.BITMAP_TYPE_ANY ))
		self.saveas.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/saveas1.png", wx.BITMAP_TYPE_ANY ))
		tools.Add(self.saveas, 0, 0, 5)


		headBar.Add(tools, 0, wx.EXPAND, 5)

		self.tree = wx.dataview.DataViewTreeCtrl(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)
		headBar.Add(self.tree, 1, wx.ALL|wx.EXPAND, 5)

		self.underBar = wx.StaticText(self, wx.ID_ANY, _(u"UB"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.underBar.Wrap(-1)

		headBar.Add(self.underBar, 0, wx.ALL, 5)


		main.Add(headBar, 1, wx.EXPAND, 5)


		self.SetSizer( main )
		self.Layout()

		self.Centre(wx.BOTH)

		data = load_json(path)
		root = self.tree.AppendItem(wx.dataview.NullDataViewItem, "Root")
		self.add_json_to_tree(self.tree, root, data)

	def __del__( self ):
		pass

if __name__ == "__main__":
	app = wx.App(False)
	frame = MyFrame1(None)
	frame.Show()
	app.MainLoop()
