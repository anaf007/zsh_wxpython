#coding=utf-8
__author__ = 'anngle'
"""基础数据管理 添加删除更新等"""
import wx,xlrd,xlwt,time
from Public.GridData import GridData
from operator import itemgetter
from SQLet import *
class sku_main(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'基础数据',pos=(0,0),size=(1200,900))
        self.panel = wx.Panel(self)
        proress_gauge = wx.ProgressDialog(u"正在打开窗口", u"窗口正在打开请稍等",100)
        proress_gauge.SetSize((400,120))
        self.panel.SetBackgroundColour('White')
        self.sku_textCtrl = wx.TextCtrl(self.panel,-1,u'')
        self.ean_textCtrl = wx.TextCtrl(self.panel,-1,u'')
        self.xm_textCtrl = wx.TextCtrl(self.panel,-1,u'')
        self.name_textCtrl = wx.TextCtrl(self.panel,-1,u'')
        searchButton = wx.Button(self.panel, -1,u"查询")
        ToExcelButton = wx.Button(self.panel, -1,u"导出表格")

        proress_gauge.Update(10)
        BodySizer = wx.BoxSizer(wx.VERTICAL) #垂直
        search_sizer=wx.BoxSizer()
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品编码：'),0,wx.ALL,5)
        search_sizer.Add(self.sku_textCtrl,0,wx.ALL,5)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品条码：'),0,wx.ALL,5)
        search_sizer.Add(self.ean_textCtrl,0,wx.ALL,5)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'外箱码：'),0,wx.ALL,5)
        search_sizer.Add(self.xm_textCtrl,0,wx.ALL,5)
        search_sizer.Add(wx.StaticText(self.panel,-1,u'商品名称：'),0,wx.ALL,5)
        search_sizer.Add(self.name_textCtrl,0,wx.ALL,5)
        search_sizer.Add(searchButton,0,wx.ALL,5)
        search_sizer.Add(ToExcelButton,0,wx.ALL,5)

        #-----------左侧添加栏------------------------
        self.add_sku_textCtrl = wx.TextCtrl(self.panel,-1,u'',size=(150,20))
        self.add_ean_textCtrl = wx.TextCtrl(self.panel,-1,u'',size=(150,20))
        self.add_xm_textCtrl = wx.TextCtrl(self.panel,-1,u'',size=(150,20))
        self.add_name_textCtrl = wx.TextCtrl(self.panel,-1,u'',size=(150,20))
        self.add_unit_textCtrl = wx.TextCtrl(self.panel,-1,u'',size=(150,20))
        self.add_note_textCtrl = wx.TextCtrl(self.panel,-1,u'',size=(150,20))
        radio_list = (u'拆零',u'整件')
        self.radio_is_zheng = wx.RadioBox(self.panel,-1,u'是否整件',(10,10),wx.DefaultSize,radio_list,2,wx.RA_SPECIFY_COLS)
        choice_list = [u'食品',u'酒水',u'面品',u'油品',u'日化']
        self.choice_type = wx.Choice(self.panel,-1,size=(150,25),choices=choice_list)
        InExcelButton = wx.Button(self.panel, -1,u"导入表格")
        self.file_select=wx.FilePickerCtrl(self.panel,wx.ID_ANY,u"",u"选择文件",u"*.*",\
                           validator=wx.DefaultValidator)
        self.file_select.GetPickerCtrl().SetLabel(u'选择文件')
        self.add_btn = wx.Button(self.panel,-1,u'添加')

        main_sizer = wx.BoxSizer()
        add_sizer = wx.BoxSizer(wx.VERTICAL)
        add_sizer.Add(wx.StaticText(self.panel,-1,u'商品编码：'),0,wx.ALL,5)
        add_sizer.Add(self.add_sku_textCtrl,0,wx.ALL,5)
        add_sizer.Add(wx.StaticText(self.panel,-1,u'商品条码：'),0,wx.ALL,5)
        add_sizer.Add(self.add_ean_textCtrl,0,wx.ALL,5)
        add_sizer.Add(wx.StaticText(self.panel,-1,u'外箱码：'),0,wx.ALL,5)
        add_sizer.Add(self.add_xm_textCtrl,0,wx.ALL,5)
        add_sizer.Add(wx.StaticText(self.panel,-1,u'商品名称：'),0,wx.ALL,5)
        add_sizer.Add(self.add_name_textCtrl,0,wx.ALL,5)
        add_sizer.Add(wx.StaticText(self.panel,-1,u'商品规格：'),0,wx.ALL,5)
        add_sizer.Add(self.add_unit_textCtrl,0,wx.ALL,5)
        add_sizer.Add(wx.StaticText(self.panel,-1,u'备注：'),0,wx.ALL,5)
        add_sizer.Add(self.add_note_textCtrl,0,wx.ALL,5)
        add_sizer.Add(self.radio_is_zheng,0,wx.ALL,5)
        add_sizer.Add(self.choice_type,0,wx.ALL,5)

        add_sizer.Add(self.add_btn,0,wx.ALL,5)
        add_sizer.Add(wx.StaticText(self.panel,-1,u''),0,wx.ALL,5)
        add_sizer.Add(self.file_select,0,wx.ALL,5)
        add_sizer.Add(InExcelButton,0,wx.ALL,5)
        add_sizer.Add(wx.StaticText(self.panel,-1,u''),0,wx.ALL,5)
        self.gauge = wx.Gauge(self.panel,-1,100,size=(170,25),style = wx.GA_PROGRESSBAR)
        add_sizer.Add(self.gauge)
        self.logmessage = wx.StaticText(self.panel,-1,u'')
        add_sizer.Add(self.logmessage)

        proress_gauge.Update(30)
        self._data = []
        _cols = (u"ID",u"商品编码",u"商品条码",u"外箱码",u"商品名称",u"类型",u"是否整件",u"每箱个数",u"备注(按件按包出)")
        res_dic = Connect().select('*','base_sku','',limit='')
        res_dic.sort(key=itemgetter(2),reverse=False)
        for x in res_dic:
            self._data.append([x[0],x[2],x[3],x[4],x[1],x[7],x[8],x[5],x[6]])
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(self.panel,size=(1200,800))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        proress_gauge.Update(60)
        self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup) #在grid行头右键点击
        searchButton.Bind(wx.EVT_BUTTON, self.OnSearchButton) #点击搜索按钮搜索
        ToExcelButton.Bind(wx.EVT_BUTTON,self.ToExcel) #导出Excel
        InExcelButton.Bind(wx.EVT_BUTTON,self.OnFile_Select_save)
        self.add_btn.Bind(wx.EVT_BUTTON,self.OnSaveButton)
        proress_gauge.Update(90)
        BodySizer.Add(search_sizer)
        main_sizer.Add(add_sizer)
        main_sizer.Add(self.grid,0,wx.ALL|wx.EXPAND,5)
        BodySizer.Add(main_sizer,0,wx.ALL|wx.EXPAND,5)
        self.panel.SetSizer(BodySizer)
        BodySizer.Fit(self)
        BodySizer.SetSizeHints(self)
        proress_gauge.Destroy()


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
        result =  Connect().delete('base_sku',where='id="%s"'%(mark[0][0]))
        if result !=0:
            wx.MessageBox(u'删除成功，请重新关闭在打开窗口查看',u'提示',wx.ICON_ASTERISK)
        else:
            wx.MessageBox(u'删除失败，请重新关闭在打开窗口查看',u'提示',wx.ICON_ERROR);return
    def OnSearchButton(self,evt):
        data_sku =  self.sku_textCtrl.GetValue()
        data_ean =  self.ean_textCtrl.GetValue()
        data_xm =  self.xm_textCtrl.GetValue()
        data_name =  self.name_textCtrl.GetValue()
        select_data = ''
        if data_sku!='':
            select_data = 'sku="'+data_sku+'"'
        if data_ean!='':
            select_data = 'ean="'+data_ean+'"'
        if data_xm!='':
            select_data = 'xiangma="'+data_xm+'"'
        if data_name!='':
            select_data = 'name like "%'+data_name+'%"'

        self.grid.DeleteRows(0,self.grid.GetNumberRows())
        res_dic = Connect().select('*','base_sku',where=select_data)
        res_dic.sort(key=itemgetter(2),reverse=False)
        for i,x in enumerate(res_dic):
            self.data.InsertRows(i+1)
            self.data.set_value(i, 0, x[0])
            self.data.set_value(i, 1, x[2])
            self.data.set_value(i, 2, x[3])
            self.data.set_value(i, 3, x[4])
            self.data.set_value(i, 4, x[1])
            self.data.set_value(i, 5, x[7])
            self.data.set_value(i, 6, x[8])
            self.data.set_value(i, 7, x[5])
            self.data.set_value(i, 8, x[6])

        self.grid.Refresh()
        self.sku_textCtrl.SetValue('')
        self.ean_textCtrl.SetValue('')
        self.xm_textCtrl.SetValue('')
        self.name_textCtrl.SetValue('')

    def OnSaveButton(self,evt):
        save_data_sku=self.add_sku_textCtrl.GetValue()
        save_data_ean=self.add_ean_textCtrl.GetValue()
        save_data_xm=self.add_xm_textCtrl.GetValue()
        save_data_name=self.add_name_textCtrl.GetValue()
        save_data_unit=self.add_unit_textCtrl.GetValue()
        save_data_note=self.add_note_textCtrl.GetValue()
        save_data_type=self.choice_type.GetString(self.choice_type.GetSelection())
        save_data_radio=self.radio_is_zheng.GetSelection()
        self.add_btn.Enable(False)
        #---------------
        self.add_sku_textCtrl.SetValue('')
        self.add_ean_textCtrl.SetValue('')
        self.add_xm_textCtrl.SetValue('')
        self.add_name_textCtrl.SetValue('')
        self.add_unit_textCtrl.SetValue('')
        self.add_note_textCtrl.SetValue('')
        if save_data_sku=='':
            wx.MessageBox(u'商品编码不能为空！',u'警告',wx.ICON_ERROR);return
        if save_data_ean=='':
            wx.MessageBox(u'商品条码不能为空！',u'警告',wx.ICON_ERROR);return
        if save_data_xm=='':
            wx.MessageBox(u'外箱码不能为空！',u'警告',wx.ICON_ERROR);return
        if not save_data_sku.isdigit():
            wx.MessageBox(u'商品编码输入的必须是数字！',u'警告',wx.ICON_ERROR);return

        if not save_data_unit.isdigit():
            wx.MessageBox(u'规格输入的必须是数字！',u'警告',wx.ICON_ERROR);return
        if save_data_name=='':
            wx.MessageBox(u'商品名称不能为空！',u'警告',wx.ICON_ERROR);return
        if save_data_unit=='' or save_data_unit<=0:
            wx.MessageBox(u'规格不能为空且不能小于1！',u'警告',wx.ICON_ERROR);return
        if save_data_type=='':
            wx.MessageBox(u'类型不能为空！',u'警告',wx.ICON_ERROR);return
        res_dic = Connect().select('*','base_sku',where="sku='%s' and ean='%s' and xiangma='%s'"%(save_data_sku,save_data_ean,save_data_xm),limit='')
        if res_dic:
            wx.MessageBox(u'该条商品记录已经存在,请重新输入编码、条码、外箱码。',u'警告',wx.ICON_ERROR);return
        else:
            insert_data_list={'sku':save_data_sku,'ean':save_data_ean,'xiangma':save_data_xm,\
                              'name':save_data_name,'unit':save_data_unit,'type':save_data_type,\
                              'is_zheng':save_data_radio,'remark':save_data_note}
            try:
                Connect().insert(insert_data_list,'base_sku')
                wx.MessageBox(u'添加完成，请重新打开窗口查看该记录',u'警告',wx.ICON_ASTERISK)
            except Exception, e:
                wx.MessageBox(u'程序发生错误，错误信息是：%s'%e,u'警告',wx.ICON_ERROR)
        self.add_btn.Enable(True)
    #表格文件更新基础数据
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
            filedata = xlrd.open_workbook(file_path,encoding_override='utf-8')
            table = filedata.sheets()[0]
            message = ""
            try:
                if table.col(0)[0].value.strip().lstrip().rstrip(',') != u'商品编码':  #去除空格在对比
                    message = u"第一行名称必须叫‘商品编码’，请返回修改"
                if table.col(1)[0].value.strip().lstrip().rstrip(',') != u'商品条码':  #去除空格在对比
                    message = u"第二行名称必须叫‘商品条码’，请返回修改"
                if table.col(2)[0].value.strip().lstrip().rstrip(',') != u'外箱码':  #去除空格在对比
                    message = u"第三行名称必须叫‘外箱码’，请返回修改"
                if table.col(3)[0].value.strip().lstrip().rstrip(',') != u'商品名称':  #去除空格在对比
                    message = u"第四行名称必须叫‘商品名称’，请返回修改"
                if table.col(4)[0].value.strip().lstrip().rstrip(',') != u'类型':  #去除空格在对比
                    message = u"第五行名称必须叫‘类型’，请返回修改"
                if table.col(5)[0].value.strip().lstrip().rstrip(',') != u'酒水件数':  #去除空格在对比
                    message = u"第六行名称必须叫‘酒水件数’，请返回修改"
                if table.col(6)[0].value.strip().lstrip().rstrip(',') != u'规格':  #去除空格在对比
                    message = u"第七行名称必须叫‘规格’，请返回修改"
                if table.col(7)[0].value.strip().lstrip().rstrip(',') != u'备注':  #去除空格在对比
                    message = u"第八行名称必须叫‘备注’，请返回修改"
                if message !="":
                    wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
            except Exception, hangtou:
                wx.MessageBox(u'程序发生错误，检查行头有误：%s'%hangtou,u'警告',wx.ICON_ERROR)
            nrows = table.nrows #行数
            table_data_list =[]
            true_data = []
            for rownum in range(1,nrows):
                 row = table.row_values(rownum)
                 if row:
                     table_data_list.append(row)
            insert_data_list = []

            for x in table_data_list:
                if Connect().get_one('sku','base_sku','sku="%s" and ean="%s" and xiangma="%s"'\
                %(int(float(str(x[0]))),int(float(str(x[1]))),int(float(str(x[2]))))):
                    true_data.append(str(x[0])+","+str(x[1])+","+str(x[2]))
                else:
                    insert_data = {'sku':int(float(str(x[0]))),'name':x[3],'ean':int(float(str(x[1]))),\
                                   'xiangma':int(float(str(x[2]))),'remark':x[7],'type':x[4],'is_zheng':int(float(str(x[5]))),\
                                   'unit':int(float(str(x[6])))}
                    insert_data_list.append(insert_data)
            self.file_select.SetPath('')
            if len(true_data) >0:
                str_true_data = str(true_data)
                wx.MessageBox(u'已经存在的数据(编码，条码，外箱码)：'+str_true_data,u'已经存在的数据',wx.ICON_ERROR);return
            if len(insert_data_list)!=0:
                Connect().insert_many(insert_data_list,'base_sku')
            wx.MessageBox(u'数据添加完成.',u'警告',wx.ICON_QUESTION)
        except Exception, e:
            wx.MessageBox(u'程序发生错误，错误信息是：%s'%e,u'警告',wx.ICON_ERROR)


    #导出表格
    def ToExcel(self,evt):
        ExcelFile = xlwt.Workbook(encoding='utf-8')
        result_goods = Connect().select('*','base_sku')
        table = ExcelFile.add_sheet(u'基础数据')
        table_title = [u'商品编码',u'商品条码',u'外箱码',u'名称',u'类型',u'酒水整件',u'规格']
        for i,x in enumerate(table_title):
            table.write(0,i,x)
        for i,x in enumerate(result_goods,1):
            table.write(i,0,x[2])
            table.write(i,1,x[3])
            table.write(i,2,x[4])
            table.write(i,3,x[1])
            table.write(i,4,x[7])
            table.write(i,5,x[8])
            table.write(i,6,x[5])
        select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
        if select_dialog.ShowModal() == wx.ID_OK:
            ExcelFile.save(select_dialog.GetPath()+u"/基础数据表"+str(int(time.time()))+".xls")
            select_dialog.Destroy()
            wx.MessageBox(u'操作完成',u'提示',wx.ICON_HAND);





