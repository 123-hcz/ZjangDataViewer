
import wx
import wx.dataview as dv
import json

def load_json(filePath):
    with open(filePath, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data


class JsonViewerFrame(wx.Frame):
    def __init__(self, parent, title ,filePath = None):
        super(JsonViewerFrame, self).__init__(parent, title=title)

        data = load_json(filePath)
        print(data)
        self.tree = dv.DataViewTreeCtrl(self)
        root = self.tree.AppendItem(wx.dataview.NullDataViewItem, filePath)
        # self.build_tree(data, self.root)
        self.Show()

if __name__ == "__main__":
    app = wx.App(False)
    frame = JsonViewerFrame(parent = None, title = "JSON Viewer", filePath = '_json/options.json')
    app.MainLoop()
