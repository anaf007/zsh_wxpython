#coding=utf8

import wx

class MsgWindow(wx.Frame):
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title, pos=(0,0))
        
        #重要的就下边两句
        self.scroller = wx.ScrolledWindow(self, -1)
        self.scroller.SetScrollbars(1,1, 1000,800)

        self.pnl =  wx.Panel(self.scroller)
        self.ms  = wx.BoxSizer(wx.VERTICAL)
        self.ms.Add(wx.Button(self.pnl,-1,u'121'))
        self.SetMinSize((500,550))
        self.pnl.SetSizer(self.ms)
        self.ms.Fit(self)

if __name__ == '__main__':
    app = wx.App(redirect=False)
    msg_win = MsgWindow(None, -1, u'消息')
    msg_win.Show(True)
    app.MainLoop()