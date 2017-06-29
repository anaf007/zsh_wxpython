#coding=utf-8
__author__ = 'anngle'
import wx
from SQLet import *
from Public.GridData import GridData
class set_message_main(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'拣选信息')
        panel = wx.Panel(self)
        panel.SetBackgroundColour('White')
        self.mdbm_textCtrl = wx.TextCtrl(panel,-1,u'')
        goBtn = wx.Button(panel,-1,u'发送信息')
        backBtn = wx.Button(panel,-1,u'撤销发送')
        _cols = (u"ID",u'拣选人员',u"商品编码",u'商品条码',u'外箱码',u'位置',u'数量(件数)',u'门店编码',u'状态',u'拣选完成时间')
        self.data =[]
        self.data = GridData(self.data,_cols)
        self.grid = wx.grid.Grid(panel)
        self.grid.SetTable(self.data)
        self.grid.AutoSize()

        topSizer = wx.BoxSizer()
        topSizer.Add(wx.StaticText(panel,-1,u'门店编码：'),1,wx.TOP|wx.LEFT,5)
        topSizer.Add(self.mdbm_textCtrl,1,wx.ALL,5)
        topSizer.Add(goBtn,1,wx.ALL,5)
        topSizer.Add(backBtn,1,wx.ALL,5)
        BodySizer = wx.BoxSizer(wx.VERTICAL) #垂直

        BodySizer.Add(topSizer)
        BodySizer.Add(self.grid,1,wx.ALL,5)
        panel.SetSizer(BodySizer)
        BodySizer.Fit(self)
        BodySizer.SetSizeHints(self)

        # self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup) #在grid行头右键点击
        goBtn.Bind(wx.EVT_BUTTON, self.OnGoBtn) #点击提交按钮


    def OnGoBtn(self,evt):
        data_mdbm = self.mdbm_textCtrl.GetValue()
        data_pid = Connect().get_one('pid','Shipment',where='',order='id Desc')[0]
        res_shipment = Connect().select('*','Shipment',where='mdbm="%s" and pid="%s"'%(data_mdbm,data_pid))
        shipment_list = []
        for x in res_shipment:
            res_base_sku = Connect().get_one('*','base_sku',where='sku="%s"'%x[2])
            if res_base_sku:
                res_now_goods = Connect().select('*','now_goods',where='sku="%s"'%(x[2]),order="allocation DESC")
                zheng_count = int(x[5])//int(res_base_sku[5])*int(res_base_sku[5]) #单据 取模 数量
                if zheng_count>0:
                    for n in res_now_goods:
                        res_quyu = Connect().get_one('quyu','huowei',where='title="%s"'%n[6])[0]
                        res_bangding = Connect().get_one('persons','bangding_user',where='district="%s"'%res_quyu)[0]
                        if zheng_count<=n[5]:
                            #商品编码，条码，外箱码，名称，拣选人，位置,数量，件数,状态，流水，门店编码
                            shipment_list.append([res_base_sku[2],res_base_sku[3],res_base_sku[4],\
                                                      res_base_sku[1],res_bangding,n[6],\
                                                      zheng_count,int(zheng_count)//int(res_base_sku[5]),0,x[10],x[7]])
                            break
                        else:
                            shipment_list.append([res_base_sku[2],res_base_sku[3],res_base_sku[4],\
                                                  res_base_sku[1],res_bangding,n[6],\
                                                  n[5],int(n[5])//int(res_base_sku[5]),0,x[10],x[7]])
                            zheng_count = zheng_count-n[5]


        for x in shipment_list:
            print x[3]








