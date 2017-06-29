#coding=utf-8
import wx
import images as img
import wx.aui
from views import *
#png图片报错 iCCP: known incorrect sRGB profile  去掉警告
wx.Log.SetLogLevel(0)
class MainPage(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1,size=(980,660),title=u'中石化仓储信息系统',style=wx.MINIMIZE_BOX)
        self.Center()
        wx.Panel(self,size=(980,0))
        self.pal = wx.Panel(self,size=(980,660))
        self.pal.Bind(wx.EVT_ERASE_BACKGROUND, self.OnEraseBackground) 
        self.pal.Bind(wx.EVT_PAINT, self.OnPaint)
        #notebook
        self.nb = wx.aui.AuiNotebook(self.pal,size=(980,516),pos=(0,115))
        self.nb.AddPage(mainPage(self.pal), u"  首  页  ")
        sizer = wx.BoxSizer()
        sizer.Add(self.nb, 1, wx.EXPAND)
        self.pal.SetSizer(sizer)
        wx.CallAfter(self.nb.SendSizeEvent)
        self.pal.Bind(wx.aui.EVT_AUINOTEBOOK_PAGE_CLOSE,self.onClosePage)
        
        self.pal.Bind(wx.EVT_LEFT_DOWN,self.OnLeftDown)

    def OnPaint (self, e):
        self.dc = wx.PaintDC(self.pal)
        self.dc.DrawBitmap(img.tb1.GetBitmap(), 15,25, 0)
        self.dc.DrawBitmap(img.tb2.GetBitmap(), 100,25, 0)
        self.dc.DrawBitmap(img.tb3.GetBitmap(), 185,25, 0)
        self.dc.DrawBitmap(img.tb4.GetBitmap(), 270,25, 0)
        self.dc.DrawBitmap(img.tb5.GetBitmap(), 355,25, 0)
        self.dc.DrawBitmap(img.tb6.GetBitmap(), 440,25, 0)
        self.dc.DrawBitmap(img.tb7.GetBitmap(), 525,25, 0)
        self.dc.DrawBitmap(img.tb8.GetBitmap(), 610,25, 0)
        self.dc.DrawBitmap(img.tb9.GetBitmap(), 695,25, 0)
        self.dc.DrawBitmap(img.tb10.GetBitmap(), 780,25, 0)
        self.dc.DrawBitmap(img.tb11.GetBitmap(), 865,25, 0)
        self.dc.DrawBitmap(img.close.GetBitmap(), 960,5, 0)
        self.dc.DrawBitmap(img.zuixiaohua.GetBitmap(), 935,0, 0)


    def onClosePage(self,evt):
    	if self.nb.GetPageCount()<=1:
    		self.nb.AddPage(mainPage(self), u"  首  页  ")

    def OnEraseBackground(self, evt):
        dc = evt.GetDC()
        dc.DrawBitmap(img.mainFrame.GetBitmap(), 0, 0)

        
            
    def OnLeftDown(self,evt):
    	pos =  evt.GetPosition()
    	if pos[0]>960 and pos[0]<980 and pos[1]>0 and pos[1]<25:
    		dlg =  wx.MessageDialog(self,u'确认关闭程序吗？',u'提示',wx.OK|wx.CANCEL)
    		if dlg.ShowModal()==wx.ID_OK:
    		    self.Close()
    		dlg.Destroy()
    	if pos[0]>935 and pos[0]<960 and pos[1]>0 and pos[1]<25:
    		self.Iconize(True)  

        if pos[0]>20 and pos[0]<70 and pos[1]>30 and pos[1]<85:
            self.nb.AddPage(jinhuodan(self), u"  进货单  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>105 and pos[0]<155 and pos[1]>30 and pos[1]<85:
            self.nb.AddPage(chuhuodan(self), u"  出货单  ") 
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>190 and pos[0]<240 and pos[1]>30 and pos[1]<85: 
            self.nb.AddPage(kucun(self), u"  商品库存  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>275 and pos[0]<325 and pos[1]>30 and pos[1]<85: 
            self.nb.AddPage(chukushuju(self), u"  出库数据  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>360 and pos[0]<410 and pos[1]>30 and pos[1]<85: 
            self.nb.AddPage(facheguanli(self), u"  发车管理  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>445 and pos[0]<495 and pos[1]>30 and pos[1]<85: 
            self.nb.AddPage(chailingfuhe(self), u"  拆零复核  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>530 and pos[0]<580 and pos[1]>30 and pos[1]<85: 
            self.nb.AddPage(zhengjianfuhe(self), u"  整件复核  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>615 and pos[0]<665 and pos[1]>30 and pos[1]<85: 
            self.nb.AddPage(dabao(self), u"  打包管理  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>700 and pos[0]<750 and pos[1]>30 and pos[1]<85: 
            self.nb.AddPage(baseSku(self), u"  商品数据  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>785 and pos[0]<835 and pos[1]>30 and pos[1]<85: pass
            # self.nb.AddPage(chuhuohuo(self), u"  货箱管理  ")
            # self.nb.SetSelection(self.nb.GetSelection()+1)
        if pos[0]>870 and pos[0]<920 and pos[1]>30 and pos[1]<85: 
            self.nb.AddPage(allocation(self), u"  货位管理  ")
            self.nb.SetSelection(self.nb.GetSelection()+1)
    
if __name__ == '__main__':
    app = wx.PySimpleApp()
    MainPage().Show()
    app.MainLoop()
