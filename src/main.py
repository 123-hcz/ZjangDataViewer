import wx
import openFile
import wx.xrc
import gettext
from excelPage import excelPage_ # 下一个类
import os
import random
import xml.etree.ElementTree as et
import zipfile
import json
import decimal
import pandas as pd
import openpyxl

_ = gettext.gettext
import excel as e



if __name__ == '__main__':
	app = wx.App()
	openFile.openFilePage(None).Show()
	app.MainLoop()