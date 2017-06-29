#coding=utf-8
import wx,xlwt,time,xlrd,threading
from SQLet import *
from Public.GridData import *
class zhuzhuanxiang_main(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'周转箱',pos=(0,0))
        pnl = wx.Panel(self)
        pnl.SetBackgroundColour('White')
        top_sizer  = wx.StaticBoxSizer(wx.StaticBox(pnl,-1,u'添加周转箱'))
        self.file_select=wx.FilePickerCtrl(pnl,wx.ID_ANY,u"",u"选择文件",u"*.*",\
                           validator=wx.DefaultValidator)
        self.file_select.GetPickerCtrl().SetLabel(u'选择文件')
        self.ExcelBtn = wx.Button(pnl,-1,u"导入表格")
        save_sizer = wx.BoxSizer()
        self.zzxtext = wx.TextCtrl(pnl,-1,'')
        save_sizer.Add(wx.StaticText(pnl,-1,u'周转箱名称:'),0,wx.ALL,5)
        save_sizer.Add(self.zzxtext,0,wx.ALL,5)
        self.saveBtn = wx.Button(pnl,-1,u'添加')
        save_sizer.Add(self.saveBtn,0,wx.ALL,5)
        top_sizer.Add(self.file_select,0,wx.ALL,5)
        top_sizer.Add(self.ExcelBtn,0,wx.ALL,5)
        top2_sizer  = wx.StaticBoxSizer(wx.StaticBox(pnl,-1,u'回收周转箱'))
        self.huishou_select=wx.FilePickerCtrl(pnl,wx.ID_ANY,u"",u"选择文件",u"*.*",\
                           validator=wx.DefaultValidator)
        self.huishou_select.GetPickerCtrl().SetLabel(u'选择文件')
        self.huishouBtn = wx.Button(pnl,-1,u"导入表格")
        top2_sizer.Add(self.huishou_select,0,wx.ALL,5)
        top2_sizer.Add(self.huishouBtn,0,wx.ALL,5)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        _cols = (u"id",u'名称',u"状态")
        self.zzx_data =[]
        result_zhouzhuanxiang = Connect().select('*','zhouzhuanxiang',order="title")
        for x in result_zhouzhuanxiang:
            if x[2]==0:
                self.zzx_data.append([x[0],x[1],u'未使用'])
            else:
                self.zzx_data.append([x[0],x[1],u'已使用'])
        self.data = GridData(self.zzx_data,_cols)
        self.grid = wx.grid.Grid(pnl,size=(500,600))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        main_sizer.Add(top_sizer,0,wx.ALL,5)
        main_sizer.Add(top2_sizer,0,wx.ALL,5)
        main_sizer.Add(save_sizer,0,wx.ALL,5)
        main_sizer.Add(self.grid,0,wx.ALL,5)
        pnl.SetSizer(main_sizer)
        main_sizer.Fit(self)
        self.saveBtn.Bind(wx.EVT_BUTTON,self.OnAdd_btn)
        self.ExcelBtn.Bind(wx.EVT_BUTTON,self.OnExcelBtn)
        self.huishouBtn.Bind(wx.EVT_BUTTON,self.OnHuishouBtn)
        self.zzxtext.SetFocus()

    def OnAdd_btn(self,evt):
    	title = self.zzxtext.GetValue()
    	if Connect().get_one('*','zhouzhuanxiang','title="%s"'%title):
    		wx.MessageBox(u'该周转箱名称已经存在，请勿重复添加',u'提示',wx.ICON_ERROR);return
        if Connect().insert({'title':title},'zhouzhuanxiang'):
        	wx.MessageBox(u'周转箱添加完成',u'提示',wx.OK);
        self.zzxtext.SetValue('')
        self.zzxtext.SetFocus()

    def OnExcelBtn(self,evtt):
        file_path = self.file_select.GetPath()
        if file_path=='':
            wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
        file_types = file_path.split('.')
        file_type = file_types[len(file_types) - 1];
        if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
            wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
        filedata = xlrd.open_workbook(file_path,encoding_override='utf-8')
        table = filedata.sheets()[0]
        message = ""
        try:
            if table.col(0)[0].value.strip() != u'周转箱号':
                message = u"第一行名称必须叫‘单据号’，请返回修改"
        except Exception,f:
            wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
        nrows = table.nrows #行数
        table_data_list =[]
        for rownum in range(1,nrows):
            if table.row_values(rownum):
                table_data_list.append(table.row_values(rownum))
        for i in table_data_list:
            if Connect().get_one('id','zhouzhuanxiang','title="%s"'%int(float(str(i[0])))):
                wx.MessageBox(u'周转箱：%s已经存在。请勿重复添加'%i[0],u'警告',wx.ICON_ERROR);return
        for i in table_data_list:
            data = {'title':i[0],'static':0}
            Connect().insert(data,'zhouzhuanxiang')
        self.file_select.SetPath('')
        wx.MessageBox(u'添加完成',u'提示',wx.OK)

           
            

        
    def OnHuishouBtn(self,evtt):
        file_path = self.huishou_select.GetPath()
        if file_path=='':
            wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
        file_types = file_path.split('.')
        file_type = file_types[len(file_types) - 1];
        if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
            wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
        filedata = xlrd.open_workbook(file_path,encoding_override='utf-8')
        table = filedata.sheets()[0]
        message = ""
        try:
            if table.col(0)[0].value.strip() != u'周转箱号':
                message = u"第一行名称必须叫‘单据号’，请返回修改"
        except Exception,f:
            wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
        nrows = table.nrows #行数
        table_data_list =[]
        for rownum in range(1,nrows):
            if table.row_values(rownum):
                table_data_list.append(table.row_values(rownum))
        for i in table_data_list:
            if Connect().get_one('id','zhouzhuanxiang','title="%s"'%int(float(str(i[0])))):
                wx.MessageBox(u'周转箱：%s已经存在。请勿重复添加'%i[0],u'警告',wx.ICON_ERROR);return
        for i in table_data_list:
            data = {'title':i[0],'static':0}
            Connect().insert(data,'zhouzhuanxiang')
        self.huishou_select.SetPath('')
        wx.MessageBox(u'添加完成',u'提示',wx.OK)



