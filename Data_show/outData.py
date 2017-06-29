#coding=utf-8
__author__ = 'anngle'
"""出库单据的管理，添加删除等"""
import wx,xlrd,traceback,time,threading
from Public.GridData import GridData
from operator import itemgetter
from SQLet import *
class outData_main(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'出库单据管理',size=(1320,900),pos=(0,0))
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour('White')
        proress_gauge = wx.ProgressDialog(u"正在打开窗口", u"窗口正在打开请稍等",100)
        self.mdbm_text = wx.TextCtrl(self.panel,-1,'',size=(60,20))
        self.sku_text = wx.TextCtrl(self.panel,-1,'',size=(50,20))
        self.ean_text = wx.TextCtrl(self.panel,-1,'',size=(110,20))
        self.name_text = wx.TextCtrl(self.panel,-1,'',size=(200,20))
        searchButton = wx.Button(self.panel, -1,u"查询")
        ToExcelButton = wx.Button(self.panel, -1,u"导出表格")
        proress_gauge.SetSize((400,120))
        search_sizer = wx.BoxSizer(wx.HORIZONTAL)#水平
        search_sizer.Add(wx.StaticText(self.panel,-1,u'门店编码：'),0,wx.ALL,5)
        search_sizer.Add(self.mdbm_text,0,wx.ALL,3)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品编码：'),0,wx.ALL,5)
        search_sizer.Add(self.sku_text,0,wx.ALL,3)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品条码：'),0,wx.ALL,5)
        search_sizer.Add(self.ean_text,0,wx.ALL,3)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品名称：'),0,wx.ALL,5)
        search_sizer.Add(self.name_text,0,wx.ALL,3)
        search_sizer.Add(searchButton,0,wx.ALL,3)
        search_sizer.Add(ToExcelButton,0,wx.ALL,3)

        self.add_pre_text = wx.TextCtrl(self.panel,-1,'')
        self.add_sku_text = wx.TextCtrl(self.panel,-1,'')
        self.add_ean_text = wx.TextCtrl(self.panel,-1,'')
        self.add_xm_text = wx.TextCtrl(self.panel,-1,'')
        self.add_name_text = wx.TextCtrl(self.panel,-1,'')
        self.add_gysname_text = wx.TextCtrl(self.panel,-1,'')
        saveButton = wx.Button(self.panel, -1,u"添加")
        proress_gauge.Update(10)
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
        save_sizer.Add(saveButton)
        self.gauge = wx.Gauge(self.panel,-1,100,size=(170,25),style = wx.GA_PROGRESSBAR)
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))
        self.file_select=wx.FilePickerCtrl(self.panel,wx.ID_ANY,u"",u"选择文件",u"*.*",\
                                           wx.DefaultPosition,wx.DefaultSize,wx.FLP_DEFAULT_STYLE,\
                                           validator=wx.DefaultValidator)
        self.file_select.GetPickerCtrl().SetLabel(u'选择文件')
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))

        self.file_select_Button = wx.Button(self.panel, -1,u"出库单导入")
        save_sizer.Add(self.file_select)
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))
        save_sizer.Add(self.file_select_Button)
        save_sizer.Add(wx.StaticText(self.panel,-1,u''))
        save_sizer.Add(self.gauge)
        self.logmessage = wx.StaticText(self.panel,-1,u'')
        save_sizer.Add(self.logmessage)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)#水平
        BodySizer = wx.BoxSizer(wx.VERTICAL) #垂直
        proress_gauge.Update(20)
        self.gauge.SetValue(0.1)
        self._data = []
        _cols = (u"ID",u'订单号',u"商品编码",u"商品条码",u"商品名称",u"数量",u"门店编码",u'门店名称')
        try:
            pid = Connect().get_one('pid','Shipment','',"id DESC")
            res_dic = Connect().select('*','Shipment',where='pid="%s"'%pid[0])
            for x in res_dic:
                self._data.append([x[0],x[6],x[2],x[3],x[1],x[5],x[7],x[8]])
        except:
            pid = [];res_dic = []
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(self.panel,size=(1100,800))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        proress_gauge.Update(60)
        main_sizer.Add(save_sizer,0,wx.ALL,5)
        main_sizer.Add(self.grid,0,wx.ALL|wx.EXPAND,5)
        BodySizer.Add(search_sizer,0,wx.ALL,5)
        BodySizer.Add(main_sizer,0,wx.ALL|wx.EXPAND,5)
        self.panel.SetSizer(BodySizer)
        proress_gauge.Update(90)
        self.Bind(wx.EVT_BUTTON,self.OnFile_Select_save,self.file_select_Button)
        self.Bind(wx.EVT_BUTTON,self.OnSearchButton,searchButton)
        proress_gauge.Destroy()


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
            out_threader.start()
            self.file_select.SetPath('')
            self.file_select_Button.Enable(False)
            self.file_select.Enable(False)
        except Exception, e:
            wx.MessageBox(u'程序发生错误，错误信息是：%s'%e,u'警告',wx.ICON_ERROR)


    def OnSearchButton(self,evt):
        data_mdbm =  self.mdbm_text.GetValue()
        data_sku =  self.sku_text.GetValue()
        data_ean =  self.ean_text.GetValue()
        data_name =  self.name_text.GetValue()
        select_data = ''
        if data_mdbm!='':
            select_data = 'mdbm = "'+data_mdbm+'"'
        if data_sku!='':
            select_data = 'sku="'+data_sku+'"'
        if data_ean!='':
            select_data = 'ean="'+data_ean+'"'
        if data_name!='':
            select_data = 'name like "%'+data_name+'%"'

        self.grid.DeleteRows(0,self.grid.GetNumberRows())
        res_dic = Connect().select('*','Shipment',where=select_data)
        res_dic.sort(key=itemgetter(7),reverse=False)
        for i,x in enumerate(res_dic):
            self.data.InsertRows(i+1)
            self.data.set_value(i, 0, x[0])
            self.data.set_value(i, 1, x[6])
            self.data.set_value(i, 2, x[2])
            self.data.set_value(i, 3, x[3])
            self.data.set_value(i, 4, x[1])
            self.data.set_value(i, 5, x[5])
            self.data.set_value(i, 6, x[7])
            self.data.set_value(i, 7, x[8])
        # self.grid.EnableCellEditControl(False)#不可编辑
        self.grid.Refresh()
        self.mdbm_text.SetValue('')
        self.sku_text.SetValue('')
        self.ean_text.SetValue('')
        self.name_text.SetValue('')

    def LogMessage(self,msg):
        self.logmessage.SetLabel(msg)
    def LogGauge(self,count):
        self.gauge.SetValue(count)



class out_Thread(threading.Thread):
    def __init__(self,windows,path):
        threading.Thread.__init__(self)
        self.window = windows
        self.path =  path
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
                if table.col(3)[0].value.strip() != u'门店编码':
                    message = u"第四行名称必须叫‘门店编码’，请返回修改"
                if table.col(4)[0].value.strip() != u'门店名称':
                    message = u"第五行名称必须叫‘门店名称’，请返回修改"
                if table.col(5)[0].value.strip() != u'商品编码':
                    message = u"第六行名称必须叫‘商品编码’，请返回修改"
                if table.col(6)[0].value.strip() != u'商品编码':
                    message = u"第七行名称必须叫‘商品编码’，请返回修改"
                if table.col(7)[0].value.strip() != u'商品名称':
                    message = u"第八行名称必须叫‘商品名称’，请返回修改"
                if table.col(8)[0].value.strip() != u'商品条码':
                    message = u"第九行名称必须叫‘商品条码’，请返回修改"
                if table.col(10)[0].value.strip() != u'配送数量':
                    message = u"第十一行名称必须叫‘配送数量’，请返回修改"
            except Exception,f:
                wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
            if message !="":
                wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
            table_data_list =[]
            for rownum in range(1,table.nrows):
                if table.row_values(rownum):
                    table_data_list.append(table.row_values(rownum))
                wx.CallAfter(self.window.LogMessage,u'正在载入数据到内存中,当前第%s行。'%str(rownum+1))
                wx.CallAfter(self.window.LogGauge,float(rownum)/(table.nrows)*33)
            insert_data_list = []
            for i,x in enumerate(table_data_list):
                wx.CallAfter(self.window.LogMessage,u'正在对比数据,当前第%s行'%str(i+1))
                wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*33+33)
                if x[0]=='' or x[3]=='' or x[4]=='' or x[5]=='' or x[6]=='' or x[7]=='' or x[8]=='' or x[10]=='':
                    raise Exception,u'在第%s列数据有空，请填写数据后在提交\n'%str(i)
                int(float(str(x[10]))) #采购数量
                int(float(str(x[3]))) #门店编码
                int(float(str(x[5]))) #商品编码
                if not Connect().select('*','base_sku','sku="%s"'%int(float(str(x[5])))):
                    Connect().insert({'name':x[7],'sku':x[5],'ean':x[6],'xiangma':x[5],'unit':1,\
                                      'remark':u'自动添加','type':u'自动添加','is_zheng':'1'},'base_sku')
            table_data_list.sort(key=itemgetter(3))
            pid = "order"+str(int(time.time()))
            install_time=int(time.time())
            for i,x in enumerate(table_data_list):
                wx.CallAfter(self.window.LogMessage,u'正在保存数据,当前第%s行'%str(i+1))
                wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*33+68)
                insert_data_list = {'`name`':x[7],'sku':int(float(str(x[5]))),'ean':int(float(str(x[8]))),\
                                    'count':int(float(str(x[10]))),'pre':x[0],\
                                    'pid':pid,'static':0,'mdbm':int(float(str(x[3]))),'mdname':x[4],\
                                    '`time`':install_time,'out_count':0}
                Connect().insert(insert_data_list,'`Shipment`')
            wx.CallAfter(self.window.LogMessage,u'出库单导入完成')
            wx.MessageBox(u'导入订单完成',u'提示',wx.OK)

        except Exception, e:
            wx.MessageBox(u'数据校验错误：%s。提示：采购数量、门店编码、商品编码等只能是数字!'%e,u'警告',wx.ICON_ERROR);return



