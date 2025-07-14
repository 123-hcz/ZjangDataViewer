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

def read_version_from_file(filepath="version.txt"):
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            version = f.read().strip()
            print(f"[日志] 本地版本: {version}")
            return version
    except FileNotFoundError:
        print("[日志] 未找到本地版本文件")
        return "版本信息未找到"
    except Exception as e:
        print(f"[日志] 读取版本信息出错: {e}")
        return f"读取版本信息出错: {e}"

def fetch_latest_lts_version():
    url = "https://api.github.com/repos/123-hcz/ZjangDataViewer/releases"
    try:
        print("[日志] 正在请求 GitHub Releases API...")
        resp = requests.get(url, timeout=5, verify=False)
        print("[警告] SSL 校验证书被忽略，建议修复本地 Python 证书环境！")
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
    except Exception as e:
        print(f"[日志] 获取版本出错: {e}")
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
                    webbrowser.open("https://github.com/123-hcz/ZjangDataViewer/releases")
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
		wx.Frame.__init__ (self, parent, id = wx.ID_ANY, title = "123Excel II" , pos = wx.DefaultPosition, size = wx.Size( 750,663 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)

		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
		self.SetBackgroundColour(wx.Colour( 255,255,255 ))

		main = wx.BoxSizer(wx.HORIZONTAL)

		path = wx.BoxSizer(wx.VERTICAL)

		self.m_dirPicker3 = wx.DirPickerCtrl(self, wx.ID_ANY, wx.EmptyString, _(u"选择文件夹"), wx.DefaultPosition, wx.DefaultSize, wx.DIRP_DEFAULT_STYLE|wx.DIRP_SMALL)
		path.Add(self.m_dirPicker3, 0, 0, 5)

		self.pathBox = wx.GenericDirCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.DIRCTRL_3D_INTERNAL | wx.SUNKEN_BORDER, wx.EmptyString, 0)

		self.pathBox.ShowHidden(False)
		self.pathBox.Bind(wx.EVT_DIRCTRL_SELECTIONCHANGED, self.getPath)
		path.Add(self.pathBox, 1, wx.EXPAND | wx.BOTTOM, 5)


		main.Add(path, 0, wx.EXPAND, 5)

		open = wx.BoxSizer(wx.VERTICAL)

		self.open = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.open.SetBitmap(wx.Bitmap( u"./image/icon_packs/classic/open.png", wx.BITMAP_TYPE_ANY ))
		self.open.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/open1.png", wx.BITMAP_TYPE_ANY ))
		open.Add(self.open, 1, wx.ALL|wx.EXPAND, 5)
		self.open.Bind(wx.EVT_BUTTON, self.onOpenClick)

		self.new = wx.BitmapButton(self, wx.ID_ANY, wx.NullBitmap, wx.DefaultPosition, wx.DefaultSize, wx.BU_AUTODRAW|0)

		self.new.SetDefault()
		self.new.SetBitmap(wx.Bitmap( u"./image/icon_packs/classic/new.png", wx.BITMAP_TYPE_ANY ))
		self.new.SetBitmapPressed(wx.Bitmap( u"./image/icon_packs/classic/new1.png", wx.BITMAP_TYPE_ANY ))
		self.new.Bind(wx.EVT_BUTTON, self.onNewClick)
		open.Add(self.new, 1, wx.ALL|wx.EXPAND, 5)

		version = read_version_from_file()
		self.versionText = wx.StaticText(self, wx.ID_ANY, _(version), wx.DefaultPosition, wx.DefaultSize, 0)
		open.Add(self.versionText, 0, wx.ALL, 5)

		main.Add(open, 1, wx.EXPAND, 5)

		self.SetSizer( main )
		self.Layout()

		self.Centre(wx.BOTH)

		# 自动检查 LTS 版本
		check_update_and_prompt(self, version)

	def getSelectedPath(self):
		return self.pathBox.GetPath()

	def getPath(self, event):
		self.path = self.pathBox.GetPath()
		self.m_dirPicker3.SetPath(self.path)
		print(f'当前路径: {self.path}')
	def onOpenClick(self, event):
		selected_path = self.m_dirPicker3.GetPath()
		print("用户选择了路径:", selected_path)
		nextPage = excelPage_(parent=None, path=selected_path)
		nextPage.Show()
		self.Close()

	def onNewClick(self, event):
		# 显示“另存为”对话框，让用户选择路径和输入文件名
		dialog = wx.FileDialog(
			self,
			message="保存新文件",
			defaultDir=wx.GetHomeDir(),
			defaultFile="新建表格.xlsx",
			wildcard="Excel 文件 (*.xlsx)|*.xlsx",
			style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
		)

		if dialog.ShowModal() == wx.ID_OK:
			file_path = dialog.GetPath()

			# 如果没有 .xlsx 后缀，自动添加
			if not file_path.endswith(".xlsx"):
				file_path += ".xlsx"

			# 创建一个空的 Excel 文件（使用 pandas）
			pd.DataFrame().to_excel(file_path, index=False)

			# 打开 excel 编辑页面
			nextPage = excelPage_(parent=None, path=file_path)
			nextPage.Show()
			self.Close()

		dialog.Destroy()

	def __del__( self ):
		pass




if __name__ == '__main__':
	app = wx.App()
	frame = openFilePage(None)
	frame.Show()
	app.MainLoop()
