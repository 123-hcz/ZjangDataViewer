import wx
import openFile
import wx.xrc
import gettext
from excelPage import excelPage_ # 下一个类
import os
import random
import xml.etree.ElementTree as et
import zipfile
import _json
import decimal
import pandas as pd
import openpyxl
import webbrowser
import init

_ = gettext.gettext
import excel as e



if __name__ == '__main__':
	# 创建 downloads 文件夹用于保存下载的文件
	downloads_path = os.path.join(os.path.dirname(__file__), '..', 'downloads')
	os.makedirs(downloads_path, exist_ok=True)
	
	# 先正常启动应用以执行检查更新逻辑
	app = wx.App()
	frame = openFile.openFilePage(None)
	frame.Show()  # 显示文件选择窗口
	
	# 在检查更新之后，创建选择框让用户选择核心
	def show_core_selection():
		choices = [ "pandas 核心：稳定，兼容性强，有最新功能，推荐","web 核心：页面优美，暂却部分功能"]
		dialog = wx.SingleChoiceDialog(frame, "请选择核心：", "选择核心", choices)
		dialog.SetSelection(0)  
		
		if dialog.ShowModal() == wx.ID_OK:
			selection = dialog.GetSelection()
			dialog.Destroy()
			
			if selection == 1:  # 选择 web 核心
				# 使用 webbrowser 打开指定网页
				webbrowser.open("https://123xls.fucku.top")
				# 关闭文件选择窗口
				frame.Close()
			else:  # 选择 pandas 核心
				# 保持文件选择窗口打开，正常运行应用
				pass
		else:
			dialog.Destroy()
	
	# 使用 wx.CallAfter 确保在检查更新之后执行核心选择逻辑
	wx.CallAfter(show_core_selection)
	
	app.MainLoop()



