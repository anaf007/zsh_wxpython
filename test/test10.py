#coding=utf-8
import  wx
class ComboxFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,'conbo box example')
        panel = wx.Panel(self,-1)
        samplelist = ['a','b','c','d']
        wx.StaticText(panel,-1,'select',(15,15))
        wx.ComboBox(panel,-1,'default value',(15,30),wx.DefaultSize,samplelist,wx.CB_DROPDOWN)
        wx.ComboBox(panel,-1,'default value',(150,30),wx.DefaultSize,samplelist,wx.CB_SIMPLE)


if __name__=="__main__":
    app = wx.PySimpleApp()
    ComboxFrame().Show()
    app.MainLoop()

