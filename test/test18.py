#coding=utf-8

import wx

class MyDialog(wx.Dialog):
    def __init__(self, parent, id, title):
        wx.Dialog.__init__(self, parent, id, title, size=(360, 370))
        
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        heading = wx.StaticText(self, -1, '中 欧', (130, 15))
        heading.SetFont(font)
        
        wx.StaticLine(self, -1, (25, 50), (300, 1))
        
        wx.StaticText(self, -1, '斯洛伐克', (25, 80))
        wx.StaticText(self, -1, '匈牙利', (25, 100))
        wx.StaticText(self, -1, '波 兰', (25, 120))
        wx.StaticText(self, -1, '捷 克', (25, 140))
        wx.StaticText(self, -1, '德 国', (25, 160))
        wx.StaticText(self, -1, '斯洛文尼亚', (25, 180))
        wx.StaticText(self, -1, '奥地利', (25, 200))
        wx.StaticText(self, -1, '瑞 典', (25, 220))
        
        wx.StaticText(self, -1, '5 379 000', (250, 80))
        wx.StaticText(self, -1, '10 084 000', (250, 100))
        wx.StaticText(self, -1, '38 635 000', (250, 120))
        wx.StaticText(self, -1, '10 240 000', (250, 140))
        wx.StaticText(self, -1, '82 443 000', (250, 160))
        wx.StaticText(self, -1, '2 001 000', (250, 180))
        wx.StaticText(self, -1, '8 032 000', (250, 200))
        wx.StaticText(self, -1, '7 288 000', (250, 220))

        wx.StaticLine(self, -1, (25, 260), (300, 1))
        sum = wx.StaticText(self, -1, '164 102 000', (240, 280))
        sum_font = sum.GetFont()
        sum_font.SetWeight(wx.BOLD)
        sum.SetFont(sum_font)
        
        wx.Button(self, 1, 'Ok', (240, 310), (60, 30))
        
        self.Bind(wx.EVT_BUTTON, self.OnOk, id=1)
        self.Center()
        
    def OnOk(self, event):
        self.Close()
        
class MyApp(wx.App):
    def OnInit(self):
        dia = MyDialog(None, -1, 'centraleurope.py')
        dia.ShowModal()
        dia.Destroy()
        return True
    
app = MyApp()
app.MainLoop()