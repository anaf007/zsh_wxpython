#coding=utf-8
__author__ = 'anngle'
"""入库单据的管理，添加删除等"""
import wx,xlrd,traceback,time,threading
from Public.GridData import GridData
from operator import itemgetter
from SQLet import *
class InData_main(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'入库单据管理',pos=(0,0),size=(1320,900))
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour('White')
        proress_gauge = wx.ProgressDialog(u"正在打开窗口", u"窗口正在打开请稍等",100)
        proress_gauge.SetSize((400,120))
        self.pre_text = wx.TextCtrl(self.panel,-1,'',size=(150,20))
        self.sku_text = wx.TextCtrl(self.panel,-1,'',size=(50,20))
        self.ean_text = wx.TextCtrl(self.panel,-1,'',size=(110,20))
        self.name_text = wx.TextCtrl(self.panel,-1,'')
        self.gysname_text = wx.TextCtrl(self.panel,-1,'')
        searchButton = wx.Button(self.panel, -1,u"查询")
        ToExcelButton = wx.Button(self.panel, -1,u"导出表格")
        proress_gauge.Update(10)
        search_sizer = wx.BoxSizer(wx.HORIZONTAL)#水平
        search_sizer.Add(wx.StaticText(self.panel,-1,u'订单号：'),0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(self.pre_text,0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品编码：'),0,wx.wx.TOP|wx.LEFT,5)
        search_sizer.Add(self.sku_text,0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品条码：'),0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(self.ean_text,0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品名称：'),0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(self.name_text,0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'供应商：'),0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(self.gysname_text,0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(searchButton,0,wx.TOP|wx.LEFT,5)
        search_sizer.Add(ToExcelButton,0,wx.TOP|wx.LEFT,5)

        self.add_pre_text = wx.TextCtrl(self.panel,-1,'')
        self.add_sku_text = wx.TextCtrl(self.panel,-1,'')
        self.add_ean_text = wx.TextCtrl(self.panel,-1,'')
        self.add_xm_text = wx.TextCtrl(self.panel,-1,'')
        self.add_name_text = wx.TextCtrl(self.panel,-1,'')
        self.add_gysname_text = wx.TextCtrl(self.panel,-1,'')
        saveButton = wx.Button(self.panel, -1,u"添加")
        proress_gauge.Update(30)
        save_sizer = wx.BoxSizer(wx.VERTICAL)#垂直
        save_sizer.Add(wx.StaticText(self.panel,-1,u'订单号：'))
        save_sizer.Add(self.add_pre_text,0,wx.ALL|wx.EXPAND,10)
        save_sizer.Add(wx.StaticText(self.panel,-1,u'商品编码：'))
        save_sizer.Add(self.add_sku_text,0,wx.ALL|wx.EXPAND,10)
        save_sizer.Add(wx.StaticText(self.panel,-1,u'商品条码：'))
        save_sizer.Add(self.add_ean_text,0,wx.ALL|wx.EXPAND,10)
        save_sizer.Add(wx.StaticText(self.panel,-1,u'外箱码：'))
        save_sizer.Add(self.add_xm_text,0,wx.ALL|wx.EXPAND,10)
        save_sizer.Add(wx.StaticText(self.panel,-1,u'商品名称：'))
        save_sizer.Add(self.add_name_text,0,wx.ALL|wx.EXPAND,10)
        save_sizer.Add(wx.StaticText(self.panel,-1,u'供应商：'))
        save_sizer.Add(self.add_gysname_text,0,wx.ALL|wx.EXPAND,10)
        self.gauge = wx.Gauge(self.panel,-1,100,size=(170,25),style = wx.GA_PROGRESSBAR)
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))
        save_sizer.Add(saveButton)
        self.file_select=wx.FilePickerCtrl(self.panel,wx.ID_ANY,u"",u"选择文件",u"*.*",\
                                           wx.DefaultPosition,wx.DefaultSize,wx.FLP_DEFAULT_STYLE,\
                                           validator=wx.DefaultValidator)
        self.file_select.GetPickerCtrl().SetLabel(u'选择文件')
        self.file_select_Button = wx.Button(self.panel, -1,u"入库单导入")
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))
        save_sizer.Add(self.file_select)
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))
        save_sizer.Add(self.file_select_Button)
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))
        save_sizer.Add(self.gauge)
        self.logmessage = wx.StaticText(self.panel,-1,u'')
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))
        save_sizer.Add(self.logmessage)
        proress_gauge.Update(40)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)#水平
        BodySizer = wx.BoxSizer(wx.VERTICAL) #垂直

        self._data = []
        _cols = (u"ID",u'订单号',u"商品编码",u"商品条码",u"商品名称",u"数量",u"供应商名称",u'流水号',u'状态')
        try:
            pid = Connect().get_one('pid','pre','',"id DESC")
            res_dic = Connect().select('id,pre,sku,ean,name,count,gysname,pid,static','pre',where='pid="%s"'%pid[0])
            for x in res_dic:
                self._data.append([x[0],x[1],x[2],x[3],x[4],x[5],x[6],x[7],x[8]])
        except:
            pid = [];res_dic = []
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(self.panel,size=(1050,800))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()

        main_sizer.Add(save_sizer)
        main_sizer.Add(self.grid,1,wx.EXPAND,wx.ALL,10)
        BodySizer.Add(search_sizer,0,wx.ALL,10)
        BodySizer.Add(main_sizer,0,wx.ALL|wx.EXPAND,10)

        self.panel.SetSizer(BodySizer)
        BodySizer.Fit(self)
        BodySizer.SetSizeHints(self)
        proress_gauge.Update(90)
        self.Bind(wx.EVT_BUTTON,self.OnFile_Select_save,self.file_select_Button)
        self.Bind(wx.EVT_BUTTON,self.OnSearchButton,searchButton)
        self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup)
        proress_gauge.Destroy()

    def LogMessage(self,msg):
        self.logmessage.SetLabel(msg)
    def LogGauge(self,count):
        self.count = count
        self.gauge.SetValue(count)


    def OnSearchButton(self,evt):
        data_pre =  self.pre_text.GetValue()
        data_sku =  self.sku_text.GetValue()
        data_ean =  self.ean_text.GetValue()
        # data_xm =  self.xm_text.GetValue()
        data_name =  self.name_text.GetValue()
        gysname =  self.gysname_text.GetValue()
        select_data = ''
        if data_pre!='':
            select_data = 'pre like "%'+data_pre+'%"'
        if data_sku!='':
            select_data = 'sku="'+data_sku+'"'
        if data_ean!='':
            select_data = 'ean="'+data_ean+'"'
        # if data_xm!='':
        #     select_data = 'xiangma="'+data_xm+'"'
        if data_name!='':
            select_data = 'name like "%'+data_name+'%"'
        if gysname!='':
            select_data = 'gysname like "%'+gysname+'%"'

        self.grid.DeleteRows(0,self.grid.GetNumberRows())
        res_dic = Connect().select('*','pre',where=select_data)
        # res_dic.sort(key=itemgetter(1),reverse=False)
        for i,x in enumerate(res_dic):
            self.data.InsertRows(i+1)
            self.data.set_value(i, 0, x[0])
            self.data.set_value(i, 1, x[6])
            self.data.set_value(i, 2, x[2])
            self.data.set_value(i, 3, x[3])
            self.data.set_value(i, 4, x[1])
            self.data.set_value(i, 5, x[5])
            self.data.set_value(i, 6, x[9])
            self.data.set_value(i, 7, x[7])
            self.data.set_value(i, 8, x[8])
        # self.grid.EnableCellEditControl(False)#不可编辑
        self.grid.Refresh()
        self.pre_text.SetValue('')
        self.sku_text.SetValue('')
        self.ean_text.SetValue('')
        # self.xm_text.SetValue('')
        self.name_text.SetValue('')
        self.gysname_text.SetValue('')



    def OnFile_Select_save(self,evt):
        try:
            file_path = self.file_select.GetPath()
            if file_path=='':
                wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
            file_types = file_path.split('.')
            file_type = file_types[len(file_types) - 1];
            if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
                wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
            wx.MessageBox(u'后台正在处理单据数据,请稍等...',u'提示',wx.OK)

            out_threader = out_Thread(self,file_path)
            self.file_select.SetPath('')
            self.file_select_Button.Enable(False)
            self.file_select.Enable(False)
            out_threader.start()
        except Exception, e:
            wx.MessageBox(u'程序发生错误，错误信息是：%s'%e,u'警告',wx.ICON_ERROR)





    def OnShowPopup(self,evt):
        evt_value = evt.GetRow()
        if evt_value==-1:
            return
        mark_data_list =  []
        mark_data_list.append(self._data[evt_value][0])
        popupmenu1=wx.Menu()#创建一个菜单
        delete_pop_menu = popupmenu1.Append(-1,u'入库')
        self.Bind(wx.EVT_MENU,lambda evt,mark=mark_data_list : self.Onopenpopmenu(evt,mark) ,delete_pop_menu )
        self.grid.PopupMenu(popupmenu1)

    def Onopenpopmenu(self,evt,mark):add_pre(mark).Show()



class add_pre(wx.Frame):
    def __init__(self,mark):
        wx.Frame.__init__(self,None,-1,u'单品入库',size=(500,300))
        result_pre = Connect().get_one('*','pre',where='id=%s'%mark[0])
        pnl = wx.Panel(self)
        wx.StaticText(pnl,-1,u'商品编码:',pos=(20,20))
        wx.StaticText(pnl,-1,str(result_pre[2]),pos=(120,20))
        wx.StaticText(pnl,-1,u'商品条码:',pos=(20,50))
        wx.StaticText(pnl,-1,str(result_pre[3]),pos=(120,50))
        wx.StaticText(pnl,-1,u'商品名称:',pos=(20,80))
        wx.StaticText(pnl,-1,result_pre[1],pos=(120,80))
        wx.StaticText(pnl,-1,u'数量:',pos=(20,110))
        self.add_count  = wx.TextCtrl(pnl,-1,u'',pos=(120,110))
        wx.StaticText(pnl,-1,u'货位号:',pos=(20,140))
        self.add_allocation  = wx.TextCtrl(pnl,-1,u'',pos=(120,140))
        savebtn = wx.Button(pnl,-1,u'提交入库', pos=(110,170))

        savebtn.Bind(wx.EVT_BUTTON,self.OnSaveBton)
        self.mark = mark
        self.add_result = result_pre

    def OnSaveBton(self,evt):
        data_add_count = self.add_count.GetValue()
        data_add_allocation = self.add_allocation.GetValue().upper()
        if data_add_count=='' or data_add_allocation=='':
            wx.MessageBox(u'信息必须输入',u'警告',wx.ICON_ERROR);return
        if not data_add_count.isdigit():
            wx.MessageBox(u'数量必须是数字',u'警告',wx.ICON_ERROR);return
        if not Connect().get_one('*','huowei',where='title="%s"'%data_add_allocation):
            wx.MessageBox(u'没有这个货位',u'警告',wx.ICON_ERROR);return

        result_base_sku = Connect().get_one('*','base_sku','sku="%s"'%self.add_result[2])
        add_js = round(float(str(data_add_count))/float(str(result_base_sku[5])),2)
        Connect().insert({'name':result_base_sku[1],'sku':result_base_sku[2],'ean':result_base_sku[3],\
                'xiangma':result_base_sku[4],'count':data_add_count,'allocation':data_add_allocation,\
                'unit':result_base_sku[5],'js':add_js,'username':'','time':time.time(),'pre':self.add_result[6]},'add_goods')
        result_now_goods =Connect().get_one('*','now_goods',where="sku='%s' and allocation='%s'"%(self.add_result[2],data_add_allocation) )
        if result_now_goods:
            Connect().update({'count':int(str(result_now_goods[5]))+int(data_add_count),'js':float(str(result_now_goods[8]))+add_js},'now_goods',where='id=%s'%result_now_goods[0])
        else:
            Connect().insert({'name':result_base_sku[1],'sku':result_base_sku[2],'ean':result_base_sku[3],\
                        'xiangma':result_base_sku[4],'count':data_add_count,'allocation':data_add_allocation,'unit':result_base_sku[5],'js':add_js},'now_goods')
        wx.MessageBox(u'添加完成',u'提示',wx.OK)
        self.Close();



class out_Thread(threading.Thread):
    def __init__(self,windows,path):
        threading.Thread.__init__(self)
        self.window = windows
        self.path = path
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()

    def stop(self):
        self.timeToQuit.set()

    def run(self):
        filedata = xlrd.open_workbook(self.path,encoding_override='utf-8')
        table = filedata.sheets()[0]
        message = ""
        try:
            try:
                if table.col(0)[0].value.strip() != u'单据号':
                    message = u"第一行名称必须叫‘单据号’，请返回修改"
                if table.col(4)[0].value.strip() != u'供应商名称':
                    message = u"第五行名称必须叫‘供应商名称’，请返回修改"
                if table.col(5)[0].value.strip() != u'商品编码':
                    message = u"第六行名称必须叫‘商品编码’，请返回修改"
                if table.col(6)[0].value.strip() != u'商品名称':
                    message = u"第七行名称必须叫‘商品名称’，请返回修改"
                if table.col(7)[0].value.strip() != u'商品条码':
                    message = u"第八行名称必须叫‘商品条码’，请返回修改"
                if table.col(9)[0].value.strip() != u'采购数量':
                    message = u"第十行名称必须叫‘采购数量’，请返回修改"
            except Exception,f:
                wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
            if message !="":
                wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
            nrows = table.nrows #行数
            table_data_list =[]
            for rownum in range(1,nrows):
                if table.row_values(rownum):
                    table_data_list.append(table.row_values(rownum))
                wx.CallAfter(self.window.LogMessage,u'正在载入数据到内存中,当前第%s行。'%str(rownum+1))
                wx.CallAfter(self.window.LogGauge,float(rownum)/(table.nrows)*25)
            insert_data_list = []
            table_data_list.sort(key=itemgetter(0))
            wx.CallAfter(self.window.LogMessage,u'正在对比是否有重复单号。')
            for i,x in enumerate(table_data_list):
                wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*25+25)
                if Connect().select('*','pre',where='pre="%s"'%x[0]):
                    wx.MessageBox(u'系统已经存在该单号，请勿重复导入',u'警告',wx.ICON_ERROR);return

            for i,x in enumerate(table_data_list):
                wx.CallAfter(self.window.LogMessage,u'正在检查基础数据,当前第%s行。'%str(i+1))
                wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*25+50)
                if x[0]=='' or x[4]=='' or x[5]=='' or x[9]=='':
                    raise Exception,u'在第%s列数据有空，请填写数据后在提交\n'%str(i)
                int(float(str(x[9]))) #采购数量
                int(float(str(x[3]))) #门店编码
                int(float(str(x[5]))) #商品编码
                if not Connect().select('*','base_sku','sku="%s"'%int(float(str(x[5])))):
                    Connect().insert({'name':x[6],'sku':x[5],'ean':x[7],'xiangma':x[7],'unit':1,\
                                      'remark':u'自动添加','type':u'自动添加','is_zheng':'0'},'base_sku')
            pid = "pre"+str(int(time.time()))
            install_time=int(time.time())
            for i,x in enumerate(table_data_list):
                wx.CallAfter(self.window.LogMessage,u'正在保存数据,当前第%s行。'%str(i+1))
                wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*25+75)
                insert_data_list = {'name':x[6],'sku':int(float(str(x[5]))),'ean':x[7],\
                                    'xiangma':u'暂不获取','count':x[9],'pre':x[0],\
                                    'pid':pid,'static':0,'gysname':x[4],'js':u'暂不计算',\
                                    'unit':u'暂不计算','time':install_time}
                Connect().insert(insert_data_list,'pre')
            wx.MessageBox(u'导入订单完成',u'提示',wx.ICON_HAND)
        except Exception, e:
            wx.MessageBox(u'数据校验错误：%s。提示：采购数量、门店编码、商品编码只能是数字!'%e,u'警告',wx.ICON_ERROR);return

