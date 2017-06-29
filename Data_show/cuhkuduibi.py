#coding=utf-8
__author__ = 'anngle'
"""出库数据对比是否够出库存不够的商品生产补货任务"""
import wx,xlwt,time,threading
from SQLet import *
from itertools import groupby
from operator import itemgetter
from Public.GridData import GridData
class Outbound_comparison(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'出库数据对比')
        panel = wx.Panel(self)
        panel.SetBackgroundColour('White')
        self.countBtn = wx.Button(panel,-1,u'出库数据对比计算')
        _cols = (u"商品编码",u'商品条码',u"外箱码",u'名称',u'数量')
        self.data =[]
        self.data = GridData(self.data,_cols)
        self.grid = wx.grid.Grid(panel)
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        self.gauge = wx.Gauge(panel,-1,100,size=(300,25),style = wx.GA_PROGRESSBAR)
        self.logmessage = wx.StaticText(panel,-1,u'')

        BodySizer = wx.BoxSizer(wx.VERTICAL) #垂直
        BodySizer.Add(self.countBtn,0,wx.ALL,5)
        BodySizer.Add(self.gauge,0,wx.ALL,5)
        BodySizer.Add(self.logmessage,0,wx.ALL,5)

        BodySizer.Add(self.grid,0,wx.ALL,5)
        panel.SetSizer(BodySizer)
        BodySizer.Fit(self)
        BodySizer.SetSizeHints(self)

        self.countBtn.Bind(wx.EVT_BUTTON, self.OnCountBtn)


    def OnCountBtn(self,evt):
        out_threader = duibi_Thread(self)
        out_threader.start()
        self.countBtn.Enable(False)
    def LogMessage(self,msg):
        self.logmessage.SetLabel(msg)
    def LogGauge(self,count):
        self.gauge.SetValue(count)


class duibi_Thread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()

    def stop(self):
        self.timeToQuit.set()

    def run(self):
        try:
            Connect().delete('replenishment_task',where='1=1')
            data_pid = Connect().get_one('pid','Shipment',where='',order='id Desc')[0]


            wx.CallAfter(self.window.LogMessage,u'出库商品数与库存商品数对比中,请等待。')
            result_shipment = Connect().select('sku,sum(count) as count,name','Shipment',where='pid="%s"'%data_pid,group="sku",order='sku')
            for x,i in enumerate(result_shipment):
                wx.CallAfter(self.window.LogGauge,float(x)/(len(result_shipment))*100)
                now_goods_count = Connect().select('sum(count)','now_goods','sku=%s'%i[0])[0][0]
                if now_goods_count:
                    if i[1]-int(now_goods_count)>0:
                        Connect().insert({'sku':i[0],'count':i[1]-int(now_goods_count),'name':i[2]},'replenishment_task')
                else:
                    Connect().insert({'sku':i[0],'count':i[1],'name':i[2]},'replenishment_task')
            wx.CallAfter(self.window.LogMessage,u'计算完毕正在导出表格,请等待。')
            #写入excel
            file = xlwt.Workbook(encoding='utf8')
            table = file.add_sheet(u'补货表')
            title_list = [u'商品编码',u'数量',u'货位号',u'商品名称']
            for i,x in enumerate(title_list):
                table.write(0,i,x)
            all_replenishment_task = Connect().select('sku,count,name','replenishment_task')
            wx.CallAfter(self.window.LogGauge,0)
            wx.CallAfter(self.window.LogMessage,u'计算完成.')
            
            if len(all_replenishment_task)==0:
                wx.MessageBox(u'没有缺货商品，请继续下一步操作。',u'提示',wx.OK);return
            wx.CallAfter(self.window.LogMessage,u'计算完成，请打开表格查看缺货商品')
            
            for index,x in enumerate(all_replenishment_task,1):
                wx.CallAfter(self.window.LogGauge,float(index)/(len(all_replenishment_task))*10+90)
                table.write(index,0,x[0])
                table.write(index,1,x[1])
                table.write(index,2,'')
                table.write(index,3,x[2])
            select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
            name_time = time.localtime(time.time())
            if select_dialog.ShowModal() == wx.ID_OK:
                file.save(select_dialog.GetPath()+u"/缺货表"+time.strftime('%Y%m%d%H%M%S',name_time)+".xls")
                select_dialog.Destroy()
                wx.MessageBox(u'操作完成',u'提示',wx.OK)
        except Exception, e:
            wx.MessageBox(u'程序发生错误，错误信息是：%s.'%e,u'警告',wx.ICON_ERROR)