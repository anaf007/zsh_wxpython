#coding=utf-8
import wx,time,xlwt,webbrowser,threading,xlrd
from SQLet import *
from operator import itemgetter
from Public.GridData import GridData
class send_car(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'发车管理',size=(500,800))
        pnl = wx.Panel(self)
        
        self.car_num = wx.TextCtrl(pnl,-1,'')
        self.mdbm = wx.TextCtrl(pnl,-1,'')
        self.btn = wx.Button(pnl,-1,u'查询')
        
        top_sizer = wx.BoxSizer()
        top_sizer.Add(wx.StaticText(pnl,-1,u'车牌号'),0,wx.ALL,5)
        top_sizer.Add(self.car_num,0,wx.ALL,5)
        top_sizer.Add(wx.StaticText(pnl,-1,u'门店编码'),0,wx.ALL,5)
        top_sizer.Add(self.mdbm,0,wx.ALL,5)
        top_sizer.Add(self.btn,0,wx.ALL,5)

        self.left_carNum = wx.TextCtrl(pnl,-1,'')
        self.left_mdbm = wx.TextCtrl(pnl,-1,'')
        self.left_person = wx.TextCtrl(pnl,-1,'')
        self.left_load = wx.TextCtrl(pnl,-1,'')
        self.left_carLen = wx.TextCtrl(pnl,-1,'')
        self.left_phone = wx.TextCtrl(pnl,-1,'')
        self.left_btn = wx.Button(pnl,-1,u'保存')
        
        left_sizer = wx.BoxSizer(wx.VERTICAL)
        left_sizer.Add(wx.StaticText(pnl,-1,u'车牌号'),0,wx.ALL,5)
        left_sizer.Add(self.left_carNum,0,wx.ALL,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'门店编码'),0,wx.ALL,5)
        left_sizer.Add(self.left_mdbm,0,wx.ALL,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'联系人'),0,wx.ALL,5)
        left_sizer.Add(self.left_person,0,wx.ALL,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'载重'),0,wx.ALL,5)
        left_sizer.Add(self.left_load,0,wx.ALL,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'车长'),0,wx.ALL,5)
        left_sizer.Add(self.left_carLen,0,wx.ALL,5)
        left_sizer.Add(wx.StaticText(pnl,-1,u'联系电话'),0,wx.ALL,5)
        left_sizer.Add(self.left_phone,0,wx.ALL,5)
        left_sizer.Add(self.left_btn,0,wx.ALL,5)
        
        self.file_select=wx.FilePickerCtrl(pnl,wx.ID_ANY,u"",u"选择文件",u"*.*",\
                                           wx.DefaultPosition,wx.DefaultSize,wx.FLP_DEFAULT_STYLE,\
                                           validator=wx.DefaultValidator)
        self.file_select.GetPickerCtrl().SetLabel(u'选择文件')
        self.file_Button = wx.Button(pnl, -1,u"表格导入")
        left_sizer.Add(self.file_select,0,wx.ALL,5)
        left_sizer.Add(self.file_Button,0,wx.ALL,5)
        
        main_sizer = wx.BoxSizer()
        _cols = (u"ID",u'车牌号',u"门店编码",u"联系人",u"载重",u"车长",u"联系电话")
        shipmnet_pid = Connect().get_one('pid','Shipment',order="id desc")[0]
        car_num = Connect().select('*','send_car','pid="%s"'%shipmnet_pid)
        self._data = car_num
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(pnl,size=(800,500))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        main_sizer.Add(left_sizer,0,wx.ALL,5)
        main_sizer.Add(self.grid)
        body_sizer = wx.BoxSizer(wx.VERTICAL)
        body_sizer.Add(top_sizer,0,wx.ALL,5)
        body_sizer.Add(main_sizer,0,wx.ALL,5)
        
        pnl.SetSizer(body_sizer)
        body_sizer.SetSizeHints(self)
        body_sizer.Fit(self)
        
        self.file_Button.Bind(wx.EVT_BUTTON,self.OnFileButton)
        
        
    def OnFileButton(self,evt):
        try:
            file_path = self.file_select.GetPath()
            if file_path=='':
                wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
            self.file_Button.Enable(False)
            file_types = file_path.split('.')
            file_type = file_types[len(file_types) - 1];
            if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
                wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
            wx.MessageBox(u'后台正在处理单据数据,请稍等...',u'提示',wx.OK)
            out_threader = car_Thread(self,file_path)
            out_threader.start()
        except Exception, e:
            wx.MessageBox(u'程序发生错误，错误信息是：%s'%e,u'警告',wx.ICON_ERROR)

    def LogGauge(self,count):
        if count>100:
            self.file_Button.SetLabel(u'操作完成')
        else:
            self.file_Button.SetLabel(str(count)+"%")
        
class car_Thread(threading.Thread):
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
                if table.col(0)[0].value.strip() != u'车牌号':
                    message = u"第一行名称必须叫‘车牌号’，请返回修改"
                if table.col(1)[0].value.strip() != u'门店编码':
                    message = u"第四行名称必须叫‘门店编码’，请返回修改"
                if table.col(2)[0].value.strip() != u'联系人':
                    message = u"第五行名称必须叫‘联系人’，请返回修改"
                if table.col(3)[0].value.strip() != u'载重':
                    message = u"第六行名称必须叫‘载重’，请返回修改"
                if table.col(4)[0].value.strip() != u'车长':
                    message = u"第七行名称必须叫‘车长’，请返回修改"
                if table.col(5)[0].value.strip() != u'联系电话':
                    message = u"第八行名称必须叫‘联系电话’，请返回修改"
            except Exception,f:
                wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
            if message !="":
                wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
            table_data_list =[]
            for rownum in range(1,table.nrows):
                if table.row_values(rownum):
                    table_data_list.append(table.row_values(rownum))
            
            shipmnet_pid = Connect().get_one('pid','Shipment',order="id desc")[0]
            
            mdbm_not_list = []
            for i,x in enumerate(table_data_list):
                wx.CallAfter(self.window.LogGauge,round(float(rownum)/(table.nrows)*30,2))
                mdbm_list = x[1].split(u'、')
                for j in mdbm_list:
                    if j.replace('\t','').replace('\n','').replace(' ',''):
                        if not Connect().get_one('id','shipment',where='pid="%s" and mdbm="%s"'%(shipmnet_pid,j.replace('\t','').replace('\n','').replace(' ',''))):
                            mdbm_not_list.append(j)
            if  mdbm_not_list:
                wx.MessageBox(u'以下门店没有出库数据，请检查门店编码是否输入正确。%s'%mdbm_not_list,u'警告',wx.ICON_ERROR);
            for i,x in enumerate(table_data_list):
                wx.CallAfter(self.window.LogGauge,round(float(str(i))/(table.nrows)*40+30,2))
                mdbm_list = x[1].split(u'、')
                for j in mdbm_list:
                    if j.replace('\t','').replace('\n','').replace(' ',''):
                        
                        install_data = {'`carNum`':x[0],'`person`':x[2],'`load`':x[3],\
                                    '`carLen`':x[4],'`phone`':x[5],'`pid`':shipmnet_pid,\
                                    '`mdbm`':int(j.replace('\t','').replace('\n','').replace(' ',''))}
                
                        Connect().insert(install_data,'`send_car`')
                        
            shipment_mdbm_list = Connect().select('`mdbm`','`shipment`',where="pid='%s'"%shipmnet_pid,group='pid,mdbm')  
            mdbm_list = []
            not_mdbm = []
            wx.CallAfter(self.window.LogGauge,85)
            for i in shipment_mdbm_list:
                mdbm_list.append(i[0].replace(' ', ''))
            for i in mdbm_list:
                if not Connect().get_one('id', '`send_car`',where='pid="%s" and mdbm="%s"'%(shipmnet_pid,i)):
                    not_mdbm.append(i)
            wx.CallAfter(self.window.LogGauge,96)
            if not_mdbm:
                ExcelFile = xlwt.Workbook(encoding='utf-8')
                table = ExcelFile.add_sheet(u'未分配车辆的门店')
                table.write(0,0,u'门店编码')
                for i,x in enumerate(not_mdbm,1):
                    table.write(i,0,x)
                select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
                name_time = time.localtime(time.time())

                if select_dialog.ShowModal() == wx.ID_OK:
                    ExcelFile.save(select_dialog.GetPath()+u"/未分配车辆的门店"+time.strftime('%Y%m%d%H%M%S',name_time)+".xls")
                    select_dialog.Destroy()
            wx.CallAfter(self.window.LogGauge,102)
            wx.MessageBox(u'操作完成',u'提示',wx.OK)
            
        except Exception, e:
            wx.MessageBox(u'发车计算运算错误：%s。'%e,u'警告',wx.ICON_ERROR);return




