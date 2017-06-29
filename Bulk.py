#coding=utf-8
import wx,xlwt,time,xlrd,threading,webbrowser
from Public.GridData import *
from SQLet import *
from operator import itemgetter
class bulk(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'选择拣选车辆',pos=(0,0))
        pnl = wx.Panel(self)
        _cols = (u'车辆',u'拣选总行数',u'完成行数',u'进度')
        self._data =[]
        result_goods = Connect()._sql_query('*','select c.carnum,\
            count(s.id),\
        sum(CASE when s.static=2 then 1 else 0 end) yiwancheng,\
        sum(CASE when s.static=2 then 1 else 0 end)/count(s.id) wanchenglv\
        from send_car c \
        join send_scattered s \
        where  s.mdbm=c.mdbm AND\
        c.pid = s.pid and \
        c.pid = (select pid from shipment ORDER BY pid desc limit 1) AND\
        s.send_type =0 \
        GROUP BY c.carNum\
        ORDER BY wanchenglv,c.mdbm;\
        ')

        for x in result_goods:
            self._data.append([x[0],x[1],x[2],str(float(x[3]*100))+"%"])
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(pnl,size=(600,800))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        body_sizer = wx.BoxSizer()
        body_sizer.Add(self.grid,0,wx.ALL,5)
        body_sizer.Fit(self)
        body_sizer.SetSizeHints(self)
        pnl.SetSizer(body_sizer)
        self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup) 

    def OnShowPopup(self,evt):
        evt_value = evt.GetRow()
        if evt_value==-1:return
        mark_data_list =  []
        mark_data_list.append([self._data[evt_value][0]])
        popupmenu1=wx.Menu()
        delete_pop_menu = popupmenu1.Append(-1,u'打开')
        self.Bind(wx.EVT_MENU,lambda evt,mark=mark_data_list : self.Onopenpopmenu(evt,mark) ,delete_pop_menu )
        self.grid.PopupMenu(popupmenu1)
    def Onopenpopmenu(self,evt,mark):
    	select_mdbm(mark[0][0]).Show();


class select_mdbm(wx.Frame):
    def __init__(self,mark):
        wx.Frame.__init__(self,None,-1,u'车辆：%s'%mark,size=(500,500))
        pnl = wx.Panel(self)
        _cols = (u'门店编码',u"散货拣选总行数",u'完成行数',u'进度')
        self._data =[]

        result_goods = Connect()._sql_query('*','\
        select c.mdbm,count(s.id),\
        sum(CASE when s.static=2 then 1 else 0 end) yiwancheng,\
        sum(CASE when s.static=2 then 1 else 0 end)/count(s.id) wanchenglv\
        from send_car c \
        join send_scattered s \
        where  s.mdbm=c.mdbm AND\
        c.pid = s.pid and \
        c.pid = (select pid from send_car ORDER BY pid desc limit 1) AND\
        s.send_type =0  AND\
        c.carNum = "%s"\
        GROUP BY c.mdbm\
        ORDER BY wanchenglv,c.mdbm;\
            '%mark)
        for x in result_goods:
            self._data.append([x[0],x[1],x[2],str(float(x[3]*100))+"%"])
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(pnl,size=(600,800))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        body_sizer = wx.BoxSizer()
        main_sizer = wx.BoxSizer()
        main_sizer.Add(self.grid,1,wx.ALL|wx.EXPAND,5)
        body_sizer.Add(main_sizer,0,wx.ALL,5)
        pnl.SetSizer(body_sizer)
        body_sizer.Fit(self)
        self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup) 

    def OnShowPopup(self,evt):
        evt_value = evt.GetRow()
        if evt_value==-1:return
        mark_data_list =  []
        mark_data_list.append([self._data[evt_value][0],self._data[evt_value][1]])
        popupmenu1=wx.Menu()
        delete_pop_menu = popupmenu1.Append(-1,u'打开')
        self.Bind(wx.EVT_MENU,lambda evt,mark=mark_data_list : self.Onopenpopmenu(evt,mark) ,delete_pop_menu )
        self.grid.PopupMenu(popupmenu1)
    def Onopenpopmenu(self,evt,mark):
        url = 'http://192.168.2.150/zsh/index.php/Storeroom/send_scattered/mdbm/'+str(mark[0][0])+'/'
        webbrowser.open(url)
        







