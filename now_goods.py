#coding=utf-8
__author__ = 'anngle'
import wx,xlwt,time,xlrd,threading
from Public.GridData import *
from SQLet import *
from operator import itemgetter
class goods_main(wx.MDIChildFrame):
    def __init__(self,parent):
        proress_gauge = wx.ProgressDialog(u"正在打开窗口", u"窗口正在打开请稍等",100)
        proress_gauge.SetSize((400,120))
        proress_gauge.Update(10)
        wx.MDIChildFrame.__init__(self,parent,title=u'库存管理',pos=(5,5))
        pnl = wx.Panel(self)
        pnl.SetBackgroundColour('White')
        self.sku_textCtrl = wx.TextCtrl(pnl,-1,u'',size=(60,20))
        self.ean_textCtrl = wx.TextCtrl(pnl,-1,u'',size=(110,20))
        self.xm_textCtrl = wx.TextCtrl(pnl,-1,u'',size=(120,20))
        self.allocation_textCtrl = wx.TextCtrl(pnl,-1,u'',size=(60,20))
        self.name_textCtrl = wx.TextCtrl(pnl,-1,u'',size=(200,20))
        selectBtn = wx.Button(pnl,-1,u'查询')
        ToExcelBtn = wx.Button(pnl,-1,u'导出库存表')

        top_sizer = wx.BoxSizer()
        proress_gauge.Update(50)
        top_sizer.Add(wx.StaticText(pnl,-1,u'商品编码:'),0,wx.ALL,5)
        top_sizer.Add(self.sku_textCtrl,0,wx.ALL,5)
        top_sizer.Add(wx.StaticText(pnl,-1,u'商品条码:'),0,wx.ALL,5)
        top_sizer.Add(self.ean_textCtrl,0,wx.ALL,5)
        top_sizer.Add(wx.StaticText(pnl,-1,u'外箱码:'),0,wx.ALL,5)
        top_sizer.Add(self.xm_textCtrl,0,wx.ALL,5)
        top_sizer.Add(wx.StaticText(pnl,-1,u'货位:'),0,wx.ALL,5)
        top_sizer.Add(self.allocation_textCtrl,0,wx.ALL,5)
        top_sizer.Add(wx.StaticText(pnl,-1,u'名称:'),0,wx.ALL,5)
        top_sizer.Add(self.name_textCtrl,0,wx.ALL,5)
        top_sizer.Add(selectBtn,1,wx.ALL,5)

        top_sizer.Add(ToExcelBtn,0,wx.ALL,5)

        _cols = (u"id",u'商品编码',u"商品条码",u'外箱码',u'名称',u'货位',u'数量',u'件数',u'规格')
        self.goods_data =[]
        result_goods = Connect().select('*','now_goods',order="sku")
        for x in result_goods:
            self.goods_data.append([x[0],x[2],x[3],x[4],x[1],x[6],x[5],0,x[7]])
        self.data = GridData(self.goods_data,_cols)
        self.grid = wx.grid.Grid(pnl,size=(1000,600))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()

        #-----------左侧添加栏------
        self.add_sku = wx.TextCtrl(pnl,-1,'',size=(150,20))
        self.add_count = wx.TextCtrl(pnl,-1,'',size=(150,20))
        self.add_allocation = wx.TextCtrl(pnl,-1,'',size=(150,20))
        self.add_btn = wx.Button(pnl,-1,u'添加')
        main_sizer = wx.BoxSizer()
        left_sizer = wx.BoxSizer(wx.VERTICAL)
        left_sizer.Add(wx.StaticText(pnl,-1,u'商品编码：'),0,wx.ALL,5)
        left_sizer.Add(self.add_sku,0,wx.ALL,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'数量：'),0,wx.ALL,5)
        left_sizer.Add(self.add_count,0,wx.ALL,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'货位号：'),0,wx.ALL,5)
        left_sizer.Add(self.add_allocation,0,wx.ALL,5)
        left_sizer.Add(self.add_btn,0,wx.ALL,5)

        self.file_select=wx.FilePickerCtrl(pnl,wx.ID_ANY,u"",u"选择文件",u"*.*",\
                           validator=wx.DefaultValidator)
        self.file_select.GetPickerCtrl().SetLabel(u'选择文件')
        self.InExcelBtn = wx.Button(pnl,-1,u"导入库存表")

        left_sizer.Add(self.file_select,0,wx.ALL,5)
        left_sizer.Add(self.InExcelBtn,0,wx.ALL,5)
        self.gauge = wx.Gauge(pnl,-1,100,size=(170,25),style = wx.GA_PROGRESSBAR)
        self.logmessage = wx.StaticText(pnl,-1,u'')
        left_sizer.Add(self.gauge,0,wx.ALL,5)
        left_sizer.Add(self.logmessage,0,wx.ALL,5)
        main_sizer.Add(left_sizer,0,wx.ALL,5)
        main_sizer.Add(self.grid,0,wx.ALL|wx.EXPAND,5)
        body_sizer = wx.BoxSizer(wx.VERTICAL)
        body_sizer.Add(top_sizer,0,wx.ALL,5)
        body_sizer.Add(main_sizer,0,wx.ALL,5)

        pnl.SetSizer(body_sizer)
        body_sizer.Fit(self)
        body_sizer.SetSizeHints(self)
        proress_gauge.Update(90)
        ToExcelBtn.Bind(wx.EVT_BUTTON,self.OnToExcel)
        self.InExcelBtn.Bind(wx.EVT_BUTTON,self.OnInExcel)
        self.add_btn.Bind(wx.EVT_BUTTON,self.OnAdd_btn)
        selectBtn.Bind(wx.EVT_BUTTON,self.OnSearchBtn)
        proress_gauge.Destroy()


    def OnToExcel(self,evt):
        ExcelFile = xlwt.Workbook(encoding='utf-8')
        result_goods = Connect().select('*','now_goods')
        table = ExcelFile.add_sheet(u'库存商品')
        table_title = [u'商品编码',u'商品条码',u'外箱码',u'名称',u'数量',u'件数',u'货位号',u'规格']
        for i,x in enumerate(table_title):
            table.write(0,i,x)
        for i,x in enumerate(result_goods,1):
            table.write(i,0,x[2])
            table.write(i,1,x[3])
            table.write(i,2,x[4])
            table.write(i,3,x[1])
            table.write(i,4,x[5])
            table.write(i,5,x[8])
            table.write(i,6,x[6])
            table.write(i,7,x[7])
        select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
        name_time = time.localtime(time.time())
        if select_dialog.ShowModal() == wx.ID_OK:
            ExcelFile.save(select_dialog.GetPath()+u"/库存表"+time.strftime('%Y%m%d%H%M%S',name_time)+".xls")
            select_dialog.Destroy()
            wx.MessageBox(u'操作完成',u'提示',wx.ICON_HAND)

    def OnInExcel(self,evt):
        try:
            file_path = self.file_select.GetPath()
            if file_path=='':
                wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
            file_types = file_path.split('.')
            file_type = file_types[len(file_types) - 1];
            if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
                wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
            
            out_threader = goods_Thread(self,file_path)
            out_threader.start()
            self.InExcelBtn.Enable(False)
            self.file_select.Enable(False)
            self.file_select.SetPath('')
        except Exception, e:
            wx.MessageBox(u'程序发生错误，错误信息是：%s.(提示：可以检查下基础数据)'%e,u'警告',wx.ICON_ERROR)

    def OnSearchBtn(self,evt):
        data_sku =  self.sku_textCtrl.GetValue()
        data_ean =  self.ean_textCtrl.GetValue()
        data_allocation =  self.allocation_textCtrl.GetValue()
        data_name =  self.name_textCtrl.GetValue()
        select_data = ''
        if data_sku!='':
            select_data = 'sku="'+data_sku+'"'
        if data_ean!='':
            select_data = 'ean="'+data_ean+'"'
        if data_name!='':
            select_data = 'name like "%'+data_name+'%"'
        if data_allocation!='':
            select_data = 'allocation like "%'+data_allocation+'%"'

        self.grid.DeleteRows(0,self.grid.GetNumberRows())
        res_dic = Connect().select('*','now_goods',where=select_data)
        res_dic.sort(key=itemgetter(2),reverse=False)
        for i,x in enumerate(res_dic):
            self.data.InsertRows(i+1)
            self.data.set_value(i, 0, x[0])
            self.data.set_value(i, 1, x[2])
            self.data.set_value(i, 2, x[3])
            self.data.set_value(i, 3, x[4])
            self.data.set_value(i, 4, x[1])
            self.data.set_value(i, 5, x[6])
            self.data.set_value(i, 6, x[5])
            self.data.set_value(i, 7, 0)
            self.data.set_value(i, 8, x[7])
        # self.grid.EnableCellEditControl(False)#不可编辑
        self.grid.Refresh()
        self.sku_textCtrl.SetValue('')
        self.ean_textCtrl.SetValue('')
        self.allocation_textCtrl.SetValue('')
        self.name_textCtrl.SetValue('')

    def OnAdd_btn(self,evt):
        save_data_sku=self.add_sku.GetValue()
        save_data_count=self.add_count.GetValue()
        save_data_allocation=self.add_allocation.GetValue().upper()

        self.add_sku.SetValue('')
        self.add_count.SetValue('')
        self.add_allocation.SetValue('')

        if save_data_allocation=='':
            wx.MessageBox(u'货位号不能为空！',u'警告',wx.ICON_ERROR);return
        if not save_data_sku.isdigit() or save_data_sku<=0 or save_data_sku=='':
            wx.MessageBox(u'商品编码输入不正确，请重新输入！',u'警告',wx.ICON_ERROR);return
        if not save_data_count.isdigit() or save_data_count<=0 or save_data_count=='':
            wx.MessageBox(u'数量输入不正确，请重新输入！',u'警告',wx.ICON_ERROR);return
        if not Connect().get_one('*','huowei',where='title="%s"'%save_data_allocation):
            wx.MessageBox(u'没有该货位号，请检查是否输入错误',u'提示',wx.ICON_ERROR);return

        result_base_sku = Connect().get_one('*','base_sku','sku="%s" '%int(save_data_sku))

        if result_base_sku:
            result_now_goods = Connect().get_one('*','now_goods','sku="%s" and allocation="%s"'%(save_data_sku,save_data_allocation))
            if result_now_goods:
                Connect().update({'count':int(float(str(result_now_goods[5])))+int(float(save_data_count))},'now_goods',where="id=%s"%result_now_goods[0])

            else:
                Connect().insert({'sku':int(float(str(result_base_sku[2]))),'ean':result_base_sku[3],'xiangma':result_base_sku[4],\
                                  'name':result_base_sku[1],'count':int(float(save_data_count)),'allocation':save_data_allocation,'unit':int(float(str(result_base_sku[5]))),\
                                  'js':0},'now_goods')
            wx.MessageBox(u'商品：%s 添加完成。'%result_base_sku[1],u'提示',wx.OK)
        else:
            wx.MessageBox(u'没有基础数据，请先添加基础数据',u'提示',wx.ICON_ERROR);return

    def LogMessage(self,msg):
        self.logmessage.SetLabel(msg)
    def LogGauge(self,count):
        self.gauge.SetValue(count)



class goods_Thread(threading.Thread):
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
        if table.col(0)[0].value.strip().lstrip().rstrip(',') != u'商品编码':  #去除空格在对比
            message = u"第一行名称必须叫‘商品编码’，请返回修改"
        if table.col(1)[0].value.strip().lstrip().rstrip(',') != u'数量':  #去除空格在对比
            message = u"第二行名称必须叫‘数量’，请返回修改"
        if table.col(2)[0].value.strip().lstrip().rstrip(',') != u'货位号':  #去除空格在对比
            message = u"第三行名称必须叫‘货位号’，请返回修改"
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return

        nrows = table.nrows #行数
        table_data_list =[]
        wx.CallAfter(self.window.LogMessage,u'(1/3)正在载入数据到内存中。')
        for rownum in range(1,nrows):
            row = table.row_values(rownum)
            if row:
                table_data_list.append(row)
                wx.CallAfter(self.window.LogGauge,float(rownum)/(table.nrows)*20)
        wx.CallAfter(self.window.LogMessage,u'(2/3)正在检查数据。')
        for i,x in enumerate(table_data_list):
            if not Connect().get_one('*','huowei','title="%s"'%str(x[2])):
                wx.MessageBox(u'单据检测有误，行数：%s，没有该货位%s。'%(str(i),x[2]),u'警告',wx.ICON_ERROR);return
            if not Connect().get_one('*','base_sku','sku="%d"'%int(float(str(x[0])))):
                wx.MessageBox(u'单据检测有误，行数：%s，没有基础数据，编码：%s。'%(str(i),x[0]),u'警告',wx.ICON_ERROR);return
            wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*20+20)
        
        wx.CallAfter(self.window.LogMessage,u'(3/3)正在保存数据。')
        for i,x in enumerate(table_data_list):
            wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*60+40)
            result_base_sku = Connect().get_one('*','base_sku','sku="%s" '%int(float(str(x[0]))))

            if result_base_sku:
                result_now_goods = Connect().get_one('*','now_goods','sku="%s" and allocation="%s"'%(int(float(str(x[0]))),str(x[2])))
                if result_now_goods:
                    Connect().update({'count':int(float(str(result_now_goods[5])))+int(float(str(x[1])))},'now_goods',where="id=%s"%result_now_goods[0])
                else:
                    Connect().insert({'sku':int(float(str(result_base_sku[2]))),'ean':result_base_sku[3],'xiangma':result_base_sku[4],\
                                      'name':result_base_sku[1],'count':int(float(str(x[1]))),'allocation':x[2],'unit':int(float(str(result_base_sku[5]))),\
                                      'js':0},'now_goods')
            else:
                wx.MessageBox(u'没有基础数据：编码为：%s'%x[0],wx.ICON_ERROR);
        wx.CallAfter(self.window.LogMessage,u'添加商品入库完成。')
        wx.CallAfter(self.window.LogGauge,0)