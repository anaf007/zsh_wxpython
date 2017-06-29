#coding=utf-8
__author__ = 'Administrator'
import wx,threading,xlrd
from SQLet import *
from Public.GridData import GridData
class huowei(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'货位管理',size=(800,850))
        pnl = wx.Panel(self)
        self.title = wx.TextCtrl(pnl,-1,'')
        self.quyu = wx.TextCtrl(pnl,-1,'')
        self.type = wx.TextCtrl(pnl,-1,u'发货区')
        self.sort = wx.TextCtrl(pnl,-1,'1000')
        self.saveBtn = wx.Button(pnl,-1,u'保存')
        left_sizer = wx.BoxSizer(wx.VERTICAL)
        left_sizer.Add(wx.StaticText(pnl,-1,u'货位名称：'),0,wx.LEFT|wx.TOP,5)
        left_sizer.Add(self.title,0,wx.LEFT,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'货位区域：(如发货A区)'),0,wx.LEFT|wx.TOP,5)
        left_sizer.Add(self.quyu,0,wx.LEFT,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'货位类型:'),0,wx.LEFT|wx.TOP,5)
        left_sizer.Add(self.type,0,wx.LEFT,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'货位排序：'),0,wx.LEFT|wx.TOP,5)
        left_sizer.Add(self.sort,0,wx.LEFT,5)
        left_sizer.Add(self.saveBtn,0,wx.LEFT,5)
        
        self.file_select=wx.FilePickerCtrl(pnl,wx.ID_ANY,u"",u"选择文件",u"*.*",\
                           validator=wx.DefaultValidator)
        self.file_select.GetPickerCtrl().SetLabel(u'选择文件')
        self.file_saveBtn = wx.Button(pnl,-1,u"导入库存表")
        
        left_sizer.Add(self.file_select,0,wx.ALL,5)
        left_sizer.Add(self.file_saveBtn,0,wx.ALL,5)
        self.gauge = wx.Gauge(pnl,-1,100,size=(170,25),style = wx.GA_PROGRESSBAR)
        self.logmessage = wx.StaticText(pnl,-1,u'')
        left_sizer.Add(self.gauge,0,wx.ALL,5)
        left_sizer.Add(self.logmessage,0,wx.ALL,5)
        
        
        body_sizer = wx.BoxSizer()
        body_sizer.Add(left_sizer,0,wx.ALL,5)
        
        _cols = (u"id",u'货位名称',u"货位类型",u'货位排序',u'货位区域')
        self.goods_data =[]
        result_goods = Connect().select('*','huowei',order="title")
        for x in result_goods:
            self.goods_data.append([x[0],x[1],x[2],x[3],x[4]])
        self.data = GridData(self.goods_data,_cols)
        self.grid = wx.grid.Grid(pnl,size=(600,700))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        
        body_sizer.Add(self.grid,0,wx.ALL|wx.EXPAND,5)
        
        pnl.SetSizer(body_sizer)
        body_sizer.Fit(self)
        body_sizer.SetSizeHints(self)
        
        self.file_saveBtn.Bind(wx.EVT_BUTTON,self.OnFileSaveBtn)
        self.saveBtn.Bind(wx.EVT_BUTTON,self.OnSaveBtn)
        
    def OnFileSaveBtn(self,evt):
        try:
            file_path = self.file_select.GetPath()
            if file_path=='':
                wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
            file_types = file_path.split('.')
            file_type = file_types[len(file_types) - 1];
            if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
                wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
            wx.MessageBox(u'后台正在处理单据数据,请稍等...',u'提示',wx.OK)
            out_threader = allocation_Thread(self,file_path)
            out_threader.start()
        except Exception, e:
            wx.MessageBox(u'程序发生错误，错误信息是：%s.'%e,u'警告',wx.ICON_ERROR)

    def OnSaveBtn(self,evt):
        data_title =  self.title.GetValue()
        data_type =  self.type.GetValue()
        data_quyu =  self.quyu.GetValue()
        data_sort =  self.sort.GetValue()
        if data_title=='' or data_type=='' or data_quyu =='' or data_sort=='':
            wx.MessageBox(u'请输入完整的信息.',u'警告',wx.ICON_ERROR);return
        if not Connect().get_one('*', 'huowei', 'title="%s"'%data_title):
            ins_data = {'title':data_title,'quyu':data_quyu,'type':data_type,'sort':data_sort}
            Connect().insert(ins_data, 'huowei')
            wx.MessageBox(u'添加完成.',u'警告',wx.OK)
        else:
            wx.MessageBox(u'该货位名称已存在.',u'警告',wx.ICON_ERROR);return
    def LogMessage(self,msg):
        self.logmessage.SetLabel(msg)
    def LogGauge(self,count):
        self.gauge.SetValue(count)
           
           
class allocation_Thread(threading.Thread): 
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
        if table.col(0)[0].value.strip().lstrip().rstrip(',') != u'货位名称':  #去除空格在对比
            message = u"第一行名称必须叫‘货位名称’，请返回修改"
        if table.col(1)[0].value.strip().lstrip().rstrip(',') != u'货位类型':  #去除空格在对比
            message = u"第三行名称必须叫‘货位类型’，请返回修改"
        if table.col(2)[0].value.strip().lstrip().rstrip(',') != u'排序':  #去除空格在对比
            message = u"第三行名称必须叫‘排序’，请返回修改"
        if table.col(3)[0].value.strip().lstrip().rstrip(',') != u'货位区域':  #去除空格在对比
            message = u"第二行名称必须叫‘货位区域’，请返回修改"
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return

        nrows = table.nrows #行数
        table_data_list =[]
        for rownum in range(1,nrows):
            row = table.row_values(rownum)
            if row:
                table_data_list.append(row)
                wx.CallAfter(self.window.LogMessage,u'正在载入数据到内存中,当前第%s行。'%str(rownum+1))
                wx.CallAfter(self.window.LogGauge,float(rownum)/(table.nrows)*10)
        for i,x in enumerate(table_data_list):
            if Connect().get_one('*','huowei','title="%s"'%str(x[0])):
                wx.MessageBox(u'已存在的货位名称%s，行数：%s，'%(x[0],str(i)),u'警告',wx.ICON_ERROR);return
            wx.CallAfter(self.window.LogMessage,u'正在检查数据,当前第%s行。'%str(i+1))
            wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*50+10)
        for i,x in enumerate(table_data_list):
            wx.CallAfter(self.window.LogMessage,u'正在保存数据,当前第%s行。'%str(i+1))
            wx.CallAfter(self.window.LogGauge,float(i)/(len(table_data_list))*40+60)
            ins_data = {'title':x[0],'quyu':x[3],'type':x[1],'sort':x[2]}
            Connect().insert(ins_data, 'huowei')
        wx.CallAfter(self.window.LogMessage,u'货位添加完成。')
        wx.MessageBox(u'货位添加完成.',u'提示',wx.OK)
        wx.CallAfter(self.window.LogGauge,0)
        