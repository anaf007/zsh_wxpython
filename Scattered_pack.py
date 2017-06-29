#coding=utf-8
import wx,time,xlwt,webbrowser
from SQLet import *
from Public.GridData import GridData
"""散货打包"""
class sanhuo_pack(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'打包管理',size=(550,800))
        pnl = wx.Panel(self)
        pnl.SetBackgroundColour('White')
        btn = wx.Button(pnl,-1,u'导入表格',size=(500,25))
        _cols = [u'发货位',u'门店编码',u'装箱号',u'打包时间']
        result_pack = Connect().select('fahuowei,mdbm,time','scattered_pack',order="time desc",group='time',limit='1000')
        self._data = []
        for i in result_pack:
            format_time = time.localtime(float(str(i[2][0:10])))
            pack_time = time.strftime('%Y-%m-%d %H:%M:%S',format_time)
            self._data.append([i[0],i[1],i[2],pack_time])

        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(pnl,size=(400,700))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        body_sizer = wx.BoxSizer(wx.VERTICAL)
        body_sizer.Add(btn,0,wx.ALL|wx.EXPAND,5)
        body_sizer.Add(self.grid,0,wx.ALL|wx.EXPAND,5)
        pnl.SetSizer(body_sizer)
        body_sizer.SetSizeHints(self)
        body_sizer.Fit(self)

        self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup) #在grid行头右键点击

        btn.Bind(wx.EVT_BUTTON,self.ToExcel)


    def OnShowPopup(self,evt):
        evt_value = evt.GetRow()
        if evt_value==-1:return
        mark_data_list =  []
        mark_data_list.append([self._data[evt_value][2]])
        popupmenu1=wx.Menu()
        delete_pop_menu = popupmenu1.Append(-1,u'打开')
        self.Bind(wx.EVT_MENU,lambda evt,mark=mark_data_list : self.Onopenpopmenu(evt,mark) ,delete_pop_menu )
        self.grid.PopupMenu(popupmenu1)

    def Onopenpopmenu(self,evt,mark):
        pack_info(mark[0][0]).Show()

    def ToExcel(self,evt):
        #写入excel
        file = xlwt.Workbook(encoding='utf8')
        table = file.add_sheet(u'散货打包总表')
        title = [u'发货位',u'门店编码',u'装箱号',u'打包时间']
        result_pack = Connect().select('fahuowei,mdbm,time','scattered_pack',order="time desc ",group='time',limit='1000')

        for i,x in enumerate(title):
            table.write(0,i,x)
        for index,x in enumerate(result_pack,1):
            table.write(index,0,x[0])
            table.write(index,1,x[1])
            table.write(index,2,x[2])
            format_time = time.localtime(float(str(x[2][0:10])))
            pack_time = time.strftime('%Y-%m-%d %H:%M:%S',format_time)
            table.write(index,3,pack_time)
        select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
        if select_dialog.ShowModal() == wx.ID_OK:
            file.save(select_dialog.GetPath()+u"/散货打包总表"+str(int(time.time()))+".xls")
            select_dialog.Destroy()
            wx.MessageBox(u'操作完成',u'提示',wx.ICON_HAND)


class pack_info(wx.Frame):
    def __init__(self,mark):
        wx.Frame.__init__(self,None,-1,mark)
        pnl = wx.Panel(self)
        self.pack_num = mark
        self.sku_text = wx.TextCtrl(pnl,-1,'')
        self.count_text = wx.TextCtrl(pnl,-1,'')
        save_btn = wx.Button(pnl,-1,u'添加')
        print_btn = wx.Button(pnl,-1,u'打印装箱清单')
        _cols = [u'ID',u'发货位',u'门店编码',u'商品编码',u'商品条码',u'数量',u'货位号',u'商品名称']
        self._data = Connect().select('id,fahuowei,mdbm,sku,ean,count,allocation,name','scattered_pack',where='time="%s"'%mark)
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(pnl)
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        top_sizer= wx.BoxSizer()
        top_sizer.Add(wx.StaticText(pnl,-1,u'商品编码:'))
        top_sizer.Add(self.sku_text,1,wx.ALL,5)
        top_sizer.Add(wx.StaticText(pnl,-1,u'数量:'))
        top_sizer.Add(self.count_text,1,wx.ALL,5)
        top_sizer.Add(save_btn,1,wx.ALL,5)
        top_sizer.Add(print_btn,1,wx.ALL,5)
        body_sizer = wx.BoxSizer(wx.VERTICAL)
        body_sizer.Add(top_sizer)
        body_sizer.Add(self.grid,2,wx.ALL|wx.EXPAND,5)
        pnl.SetSizer(body_sizer)
        body_sizer.SetSizeHints(self)
        body_sizer.Fit(self)

        save_btn.Bind(wx.EVT_BUTTON,self.ToSaveBtn)
        print_btn.Bind(wx.EVT_BUTTON,self.ToPrint)
        self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup) #在grid行头右键点击



    def OnShowPopup(self,evt):
        evt_value = evt.GetRow()
        if evt_value==-1:return
        mark_data_list =  []
        mark_data_list.append([self._data[evt_value][0]])
        popupmenu1=wx.Menu()
        delete_pop_menu = popupmenu1.Append(-1,u'删除')
        self.Bind(wx.EVT_MENU,lambda evt,mark=mark_data_list : self.Onopenpopmenu(evt,mark) ,delete_pop_menu )
        self.grid.PopupMenu(popupmenu1)

    def Onopenpopmenu(self,evt,mark):
        if Connect().delete('scattered_pack',where='id="%s"'%(mark[0][0])):
            wx.MessageBox(u'删除成功，请重新关闭在打开窗口查看',u'提示',wx.ICON_ASTERISK)
        else:
            wx.MessageBox(u'删除失败，请重新关闭在打开窗口查看',u'提示',wx.ICON_ERROR);return


    def ToSaveBtn(self,evt):
        data_sku =  self.sku_text.GetValue()
        data_count =  self.count_text.GetValue()
        if not data_sku.isdigit() or not data_count.isdigit():
            wx.MessageBox(u'商品编码或数量只能为数字',u'提示',wx.ICON_ERROR);return
        result_pack = Connect().get_one('*','scattered_pack',where='time="%s"'%self.pack_num)
        if not result_pack:return
        result_scattered = Connect().get_one('*','send_scattered',where='pid="%s" and mdbm="%s" and sku="%s"'%(result_pack[7],result_pack[8],data_sku))
        if not result_scattered:
            wx.MessageBox(u'该门店没有该商品编码',u'提示',wx.ICON_ERROR);return

        insert_data = {'sku':data_sku,'ean':int(str(result_scattered[2])),'name':result_scattered[3],\
                       'allocation':result_scattered[5],'count':data_count,'pid':result_scattered[8],\
                       'mdbm':result_scattered[9],'fahuowei':result_scattered[10],'time':self.pack_num}

        if Connect().insert(insert_data,'scattered_pack'):
            wx.MessageBox(u'添加成功',u'提示',wx.ICON_HAND)
        else:
            wx.MessageBox(u'添加失败',u'提示',wx.ICON_ERROR)

    def ToPrint(self,evt):
        url = 'http://192.168.2.150/zsh/index.php/Storeroom/print_pack_info/pack_num/'+self.pack_num
        webbrowser.open(url)

