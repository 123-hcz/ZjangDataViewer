import wx
import wx.xrc
import webbrowser
import requests
import threading

def checkFile(fileName,fillText):
    try:
        with open(fileName, "r", encoding="utf-8") as f:
           A = f.read()
    except:
        with open(fileName, "w",encoding="utf-8") as f:
            f.write(fillText)

checkFile("_json/options.json",'''{
    "autoUpdate": true,
    "None_value_show": "",
    "language": "中文:中国",
    "debug": true,
    "usingRoundWay": "ROUND_HALF_UP",
    "decimalPlaceNum": 3,
    "initItemsRow": 1,
    "initNamesCol": 3,
    "newGridRow": 500,
    "newGridCol": 100,
    "showColAdd": "*1.5",
    "showRowAdd": "*1.5",
    "showNoneValue": "",
    "theme": "theme_Classic",
    "version": "Unknown"
}'''
)
