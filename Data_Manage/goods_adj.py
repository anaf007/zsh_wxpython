#coding=utf-8
import wx,time,xlwt,webbrowser
from SQLet import *
from Public.GridData import GridData
"""补货"""
class buhuo(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'补货信息发送',size=(300,800))
        pnl = wx.Panel(self)
        pnl.SetBackgroundColour('White')
        body_sizer = wx.BoxSizer()
        self.text_fahuowei = wx.TextCtrl(pnl,-1,'')
        self.text_sku = wx.TextCtrl(pnl,-1,'')
        self.text_count = wx.TextCtrl(pnl,-1,'')
        send_btn = wx.Button(pnl,-1,u'发送')
        body_sizer.Add(wx.StaticText(pnl,-1,u'发货位'),1,wx.ALL|wx.EXPAND,5)
        body_sizer.Add(self.text_fahuowei,1,wx.ALL|wx.EXPAND,5)
        body_sizer.Add(wx.StaticText(pnl,-1,u'商品编码'),1,wx.ALL|wx.EXPAND,5)
        body_sizer.Add(self.text_sku,1,wx.ALL|wx.EXPAND,5)
        body_sizer.Add(wx.StaticText(pnl,-1,u'数量'),1,wx.ALL|wx.EXPAND,5)
        body_sizer.Add(self.text_count,1,wx.ALL|wx.EXPAND,5)
        body_sizer.Add(send_btn,1,wx.ALL|wx.EXPAND,5)
        pnl.SetSizer(body_sizer)
        body_sizer.SetSizeHints(self)
        body_sizer.Fit(self)

        send_btn.Bind(wx.EVT_BUTTON,self.Tosend_btn)

    def Tosend_btn(self,evt):
        fahuowei = self.text_fahuowei.GetValue()
        sku = self.text_sku.GetValue()
        count = self.text_count.GetValue()
        result_shipmnet = Connect().get_one('pid','Shipment',order="id desc")[0]
        result_scatterred = Connect().get_one('*','send_scattered',where='pid="%s" and fahuowei="%s" and sku="%s"'%(result_shipmnet,fahuowei,sku))
        if not result_scatterred:
            wx.MessageBox(u'该发货位没有该商品，请检查是否有误。',u'提示',wx.ICON_ERROR);return
        insert_data = {'count':count,'static':1,'ean':u'补货—'+result_scatterred[2],'name':result_scatterred[3],'persons':result_scatterred[4],'allocation':result_scatterred[5],'pid':result_scatterred[8]}
        Connect().insert(insert_data,'send_scattered')
        wx.MessageBox(u'发送完成',u'提示',wx.ICON_HAND)