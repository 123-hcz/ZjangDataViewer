# openFile.py

# -*- coding: utf-8 -*-

#--------------------------------------------------------------------------
# Python code generated with wxFormBuilder (version 3.9.0 Jun 14 2020)
# http://www.wxformbuilder.org/
#
# PLEASE DO *NOT* EDIT THIS FILE!
#--------------------------------------------------------------------------

import wx
import wx.xrc
import webbrowser
import os
import requests
import threading

import gettext
from excelPage import excelPage_ # 下一个类
import pandas as pd
import excel as e # 导入excel模块以便创建新XML文件

def read_version_from_file(filepath="version.txt"):
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            version = f.read().strip()
            print(f"[日志] 本地版本: {version}")
            return version
    except FileNotFoundError:
        print("[日志] 未找到本地版本文件")
        return "版本信息未找到"
    except Exception as exc:
        print(f"[日志] 读取版本信息出错: {exc}")
        return f"读取版本信息出错: {exc}"

def fetch_latest_lts_version():
    url = "https://api.github.com/repos/123-hcz/ZjangDataViewer/releases"
    try:
        print("[日志] 正在请求 GitHub Releases API...")
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            releases = resp.json()
            if releases:
                latest = releases[0]
                print(f"[日志] 获取到最新版本: {latest['tag_name']}")
                return latest["tag_name"]
            print("[日志] 未找到任何版本")
            return None
        else:
            print(f"[日志] GitHub API 请求失败: {resp.status_code}")
            return None
    except Exception as exc:
        print(f"[日志] 获取版本出错: {exc}")
        return None

def check_update_and_prompt(frame, local_version):
    def _check():
        latest_lts = fetch_latest_lts_version()
        if latest_lts is None:
            print("[日志] 未能获取 LTS 版本，跳过自动升级提示")
            return
        if local_version != latest_lts:
            print(f"[日志] 检测到新版本: {latest_lts}，本地为: {local_version}，准备弹窗提示用户")
            def ask_update():
                dlg = wx.MessageDialog(
                    frame,
                    f"检测到新版本：{latest_lts}\n当前版本：{local_version}\n是否前往下载页面？",
                    "发现新版本",
                    wx.YES_NO | wx.ICON_QUESTION
                )
                result = dlg.ShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    print("[日志] 用户选择升级，打开下载页面")
                    webbrowser.open("https://gitee.com/hcz-123/ZjangDataViewer/releases")
                else:
                    print("[日志] 用户选择暂不升级")
            wx.CallAfter(ask_update)
        else:
            print("[日志] 当前已是最新版本")
    threading.Thread(target=_check, daemon=True).start()

_ = gettext.gettext

#--------------------------------------------------------------------------
#  Class openFilePage
#---------------------------------------------------------------------------

class openFilePage ( wx.Frame ):

	def __init__(self, parent ):
		wx.Frame.__init__ (self, parent, id = wx.ID_ANY, title = "123Excel II" , pos = wx.DefaultPosition, size = wx.Size( 950,663 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)

		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour( 255,255,255 ))

		main = wx.BoxSizer(wx.HORIZONTAL)

		path = wx.BoxSizer(wx.VERTICAL)

		self.m_dirPicker3 = wx.DirPickerCtrl(self, wx.ID_ANY, wx.EmptyString, _(u"选择文件夹"), wx.DefaultPosition, wx.DefaultSize, wx.DIRP_DEFAULT_STYLE|wx.DIRP_SMALL)
		path.Add(self.m_dirPicker3, 0, wx.EXPAND|wx.ALL, 5) # 使用EXPAND以确保宽度一致

		self.pathBox = wx.GenericDirCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.DIRCTRL_3D_INTERNAL | wx.DIRCTRL_SHOW_FILTERS | wx.SUNKEN_BORDER, "Excel & XML & JSON 文件 (*.xlsx;*.xml;*.json)|*.xlsx;*.xml;*.json", 0)
		self.pathBox.ShowHidden(False)
		self.pathBox.Bind(wx.EVT_DIRCTRL_SELECTIONCHANGED, self.getPath)
		self.pathBox.Bind(wx.EVT_DIRCTRL_FILEACTIVATED, self.onFileActivated)
		path.Add(self.pathBox, 1, wx.EXPAND | wx.ALL, 5)


		# 修正：将path（路径框）的比例设为1，使其可以自由伸展
		main.Add(path, 1, wx.EXPAND, 5)

		open_sizer = wx.BoxSizer(wx.VERTICAL)

		self.readMode = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, [_(u"以二维表格形式阅读"), _(u"以树状结构阅读[正在开发]")], 0)
		self.readMode.SetSelection(0)
		open_sizer.Add(self.readMode, 1, wx.ALL | wx.EXPAND, 5)

		self.open = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/open.png", wx.BITMAP_TYPE_ANY ))
		self.open.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/open1.png", wx.BITMAP_TYPE_ANY ))
		self.open.Bind(wx.EVT_BUTTON, self.onOpenClick)
		open_sizer.Add(self.open, 1, wx.ALL|wx.EXPAND, 5)

		self.new = wx.BitmapButton(self, wx.ID_ANY, wx.Bitmap( u"./image/icon_packs/classic/new.png", wx.BITMAP_TYPE_ANY ))
		self.new.SetDefault()
		self.new.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/new1.png", wx.BITMAP_TYPE_ANY ))
		self.new.Bind(wx.EVT_BUTTON, self.onNewClick)
		open_sizer.Add(self.new, 1, wx.ALL|wx.EXPAND, 5)

		version = read_version_from_file()
		self.versionText = wx.StaticText(self, wx.ID_ANY, _(version), wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL)
		open_sizer.Add(self.versionText, 0, wx.ALL|wx.EXPAND, 5)

		# 修正：将open_sizer（按钮区）的比例设为0，使其保持最小宽度，不再被拉伸
		main.Add(open_sizer, 0, wx.EXPAND, 5)

		self.SetSizer( main )
		self.Layout()

		self.Centre(wx.BOTH)

		check_update_and_prompt(self, version)

	def getPath(self, event):
		self.path = self.pathBox.GetPath()
		if os.path.isdir(self.path):
			self.m_dirPicker3.SetPath(self.path)
		else:
			self.m_dirPicker3.SetPath(os.path.dirname(self.path))
		event.Skip()

	def onFileActivated(self, event):
		self.onOpenClick(event)
		event.Skip()

	def onOpenClick(self, event):
		selected_path = self.pathBox.GetPath()

		if not selected_path or not os.path.isfile(selected_path):
			wx.MessageBox("请选择一个有效的文件。", "提示", wx.OK | wx.ICON_INFORMATION)
			return

		file_ext = os.path.splitext(selected_path)[1].lower()
		if file_ext not in ['.xlsx', '.xml' ,'.json']:
			wx.MessageBox("不支持的文件类型。请选择一个 .xlsx 或 .xml 或 .json 文件。", "错误", wx.OK | wx.ICON_ERROR)
			return

		print("用户选择了路径:", selected_path)
		nextPage = excelPage_(parent=None, path=selected_path)
		nextPage.Show()
		self.Close()
		event.Skip()

	def onNewClick(self, event):
		dialog = wx.FileDialog(
			self,
			message="创建新文件",
			defaultDir=wx.GetHomeDir(),
			defaultFile="新建文件",
			wildcard="Excel 文件 (*.xlsx)|*.xlsx|XML 文件 (*.xml)|*.xml|JSON 文件 (*.json)|*.json",
			style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
		)

		if dialog.ShowModal() == wx.ID_OK:
			file_path = dialog.GetPath()
			filter_index = dialog.GetFilterIndex()

			if filter_index == 0 and not file_path.lower().endswith('.xlsx'):
				file_path += '.xlsx'
			elif filter_index == 1 and not file_path.lower().endswith('.xml'):
				file_path += '.xml'

			file_ext = os.path.splitext(file_path)[1].lower()

			try:
				if file_ext == ".xlsx":
					pd.DataFrame().to_excel(file_path, index=False)
				elif file_ext == ".xml":
					e.write_to_xml([], file_path)
				else:
					wx.MessageBox("未知的文件类型。", "错误", wx.OK | wx.ICON_ERROR)
					dialog.Destroy()
					return

				nextPage = excelPage_(parent=None, path=file_path)
				nextPage.Show()
				self.Close()
			except Exception as err:
				wx.MessageBox(f"创建新文件时出错: {err}", "错误", wx.OK | wx.ICON_ERROR)

		dialog.Destroy()
		event.Skip()

	def __del__( self ):
		pass

if __name__ == '__main__':
	app = wx.App()
	frame = openFilePage(None)
	frame.Show()
	app.MainLoop()