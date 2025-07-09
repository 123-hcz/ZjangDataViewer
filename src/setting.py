# -*- coding: utf-8 -*-

#--------------------------------------------------------------------------
# Python code generated with wxFormBuilder (version 3.9.0 Jun 14 2020)
# http://www.wxformbuilder.org/
#
# PLEASE DO *NOT* EDIT THIS FILE!
#--------------------------------------------------------------------------

import wx
import wx.xrc

import gettext
_ = gettext.gettext

#--------------------------------------------------------------------------
#  Class main
#---------------------------------------------------------------------------

class main ( wx.Frame ):

	def __init__(self, parent):
		wx.Frame.__init__ (self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 496,677 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)

		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour( 239, 235, 235 ))

		mainsizer = wx.BoxSizer(wx.HORIZONTAL)

		commonSetting = wx.BoxSizer(wx.VERTICAL)

		self.commonSettingText = wx.StaticText(self, wx.ID_ANY, _(u"常用设置"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.commonSettingText.Wrap(-1)

		commonSetting.Add(self.commonSettingText, 0, wx.ALL, 5)

		self.version = wx.StaticText(self, wx.ID_ANY, _(u"版本"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.version.Wrap(-1)

		commonSetting.Add(self.version, 0, wx.ALL, 5)

		self.settingInput = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.HSCROLL)
		self.settingInput.SetToolTip(_(u"输入配置"))

		commonSetting.Add(self.settingInput, 1, wx.ALL|wx.EXPAND, 5)


		mainsizer.Add(commonSetting, 1, wx.EXPAND, 5)

		self.line1 = wx.StaticLine(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_VERTICAL)
		mainsizer.Add(self.line1, 0, wx.EXPAND, 5)

		advancedSetting = wx.BoxSizer(wx.VERTICAL)

		self.advSettingText = wx.StaticText(self, wx.ID_ANY, _(u"高级设置"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.advSettingText.Wrap(-1)

		advancedSetting.Add(self.advSettingText, 0, wx.ALL, 5)

		nullValueReplace = wx.BoxSizer(wx.HORIZONTAL)

		self.nvrText = wx.StaticText(self, wx.ID_ANY, _(u"空值替换"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.nvrText.Wrap(-1)

		nullValueReplace.Add(self.nvrText, 0, wx.ALL, 5)

		self.nvr = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
		nullValueReplace.Add(self.nvr, 0, 0, 5)


		advancedSetting.Add(nullValueReplace, 1, wx.EXPAND, 5)

		debug = wx.BoxSizer(wx.HORIZONTAL)

		self.debugText = wx.StaticText(self, wx.ID_ANY, _(u"调试模式"), wx.DefaultPosition, wx.DefaultSize, 0)
		self.debugText.Wrap(-1)

		debug.Add(self.debugText, 0, wx.ALL, 5)

		self.d = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
		debug.Add(self.d, 0, 0, 5)


		advancedSetting.Add(debug, 1, wx.EXPAND, 5)


		mainsizer.Add(advancedSetting, 1, wx.EXPAND, 5)


		self.SetSizer( mainsizer )
		self.Layout()

		self.Centre(wx.BOTH)

	def __del__( self ):
		pass


if __name__ == '__main__':
	app = wx.App()
	main = main(None)
	main.Show()
	app.MainLoop()
