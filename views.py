#coding=utf-8
import wx,wx.grid,time,xlwt,sys,webbrowser,os,winsound,copy,operator
from wx.lib.itemspicker import ItemsPicker,IP_REMOVE_FROM_CHOICES,EVT_IP_SELECTION_CHANGED
import  wx.lib.popupctl as  pop
#统一处理panel线程事件
from Controller import *
reload(sys)
sys.setdefaultencoding("utf-8")
#主页
class mainPage(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self,parent)
        t = wx.StaticText(self, -1, u"欢迎光临")
#进货单
class jinhuodan(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		self.SetBackgroundColour('white')
		wx.StaticText(self, -1, u"进货单",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.bianhao = wx.StaticText(self, -1, u"单据编号：——",pos=(600,20))
		self.time = wx.StaticText(self, -1, u"导入时间：——",pos=(800,20))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		_cols = [u'商品编码',u'商品条码',u'商品名称',u'数量',u'货位号',u'供应商']
		self.grid  = wx.grid.Grid(self,size=(980,370),pos=(0,50))
		self.grid.CreateGrid(20,6)
		width_Col = [60,110,325,60,60,260]
		for i,x in enumerate(width_Col):
			self.grid.SetColSize(i,x)
		self.grid.SetBackgroundColour('white')
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		self.grid.SetDefaultCellOverflow(False)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		self.hostBtn = wx.Button(self,-1,u'查看历史单据',pos=(10,440))
		self.fileBtn = wx.Button(self,-1,u'导入新单据',pos=(120,440))
		tishi = wx.StaticText(self,-1,u'提示：导入的Excel文件字段名和顺序必须与“进货单模板”相符合',pos=(450,450))
		downStatic = wx.StaticText(self,-1,u'点击下载模板',pos=(880,450))
		downStatic.SetForegroundColour('blue')
		font = wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD)
		downStatic.SetFont(font)
		tishi.SetFont(font)
		self.gau = wx.Gauge(self,-1, 100, size=(980, 3), pos=(0,483))
		
		downStatic.Bind(wx.EVT_LEFT_DOWN,self.DownStatic)
		self.fileBtn.Bind(wx.EVT_BUTTON,self.ExcelInput)
		self.hostBtn.Bind(wx.EVT_BUTTON,self.OnhostBtn)
		try:
			jinhuo_Thread(self).start()
		except Exception, e:
			wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR);return


	def DownStatic(self,evt):
		ExcelFile = xlwt.Workbook(encoding='utf-8')
		table = ExcelFile.add_sheet(u'0')
		table.write(0,0,u'商品编码')
		table.write(0,1,u'商品条码')
		table.write(0,2,u'商品名称')
		table.write(0,3,u'数量')
		table.write(0,4,u'货位号')
		table.write(0,5,u'供应商')
		select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
		if select_dialog.ShowModal() == wx.ID_OK:
			ExcelFile.save(select_dialog.GetPath()+u"/进货单模板.xls")
		select_dialog.Destroy()

	def OnhostBtn(self,evt):
		pass
	def SetTextStringValue(self,msg):
		self.bianhao.SetLabel(msg[0])
		self.time.SetLabel(msg[1])
	def SetGridValue(self,gridData):
		if gridData:
			self.grid.BeginBatch()
			for i in range(self.grid.GetNumberRows()-1):
				self.grid.DeleteRows(1)
			for i,x in enumerate(gridData):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,str(x[2]))
				self.grid.SetCellValue(i,1,str(x[3]))
				self.grid.SetCellValue(i,2,str(x[1]))
				self.grid.SetCellValue(i,3,str(x[4]))
				self.grid.SetCellValue(i,4,x[8])
				self.grid.SetCellValue(i,5,x[7])
				self.grid.ForceRefresh()
			self.grid.DeleteRows(self.grid.GetNumberRows()-1)
			self.grid.EndBatch()

	def ExcelInput(self,evt):
		dlg = wx.FileDialog(self, message=u"请选择要更新的表单文件",\
			defaultDir=os.getcwd(),defaultFile="",wildcard=u"Excel2003文件 (*.xls)|*.xls|\nExcel2007文件\
			(*.xlsx)|*.xlsx",style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			try:
				file_path = dlg.GetPaths()[0]
				file_types = file_path.split('.')
				file_type = file_types[len(file_types) - 1]
				if file_path=='':
					wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
				if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
					wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
				self.fileBtn.Enable(False)
				InExcel_Thread(self,file_path).start()
			except Exception, e:
				wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR);return


	def SetBtnText(self,text):
		self.fileBtn.SetLabel(text)
	def SetGauge(self,count):
		self.gau.SetValue(count)
	def SetInFileOk(self):
		wx.MessageBox(u'导入完成',u'提示',wx.ICON_INFORMATION)
		for i in range(self.grid.GetNumberRows()-1):
			self.grid.DeleteRows(1)
		jinhuo_Thread(self).start()
#出货单
class chuhuodan(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		self.SetBackgroundColour('white')
		wx.StaticText(self, -1, u"出货单",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.bianhao = wx.StaticText(self, -1, u"单据编号：——",pos=(600,20))
		self.time = wx.StaticText(self, -1, u"时间：——",pos=(800,20))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		_cols = [u'门店编码',u'商品编码',u'商品条码',u'商品名称',u'数量',u'门店名称']
		self.grid  = wx.grid.Grid(self,size=(980,370),pos=(0,50))
		self.grid.CreateGrid(20,6)
		width_Col = [70,60,110,315,70,250]
		for i,x in enumerate(width_Col):
			self.grid.SetColSize(i,x)
		self.grid.SetBackgroundColour('white')
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		self.grid.SetDefaultCellOverflow(False)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		self.hostBtn = wx.Button(self,-1,u'查看历史单据',pos=(10,440))
		self.fileBtn = wx.Button(self,-1,u'导入新单据',pos=(120,440))
		# self.file_select =wx.FilePickerCtrl(self,wx.ID_ANY,u"",u"选择文件",u"*.*",validator=wx.DefaultValidator,pos=(120,440))
		# self.file_select.GetPickerCtrl().SetLabel(u'浏览')
		tishi = wx.StaticText(self,-1,u'提示：导入的Excel文件字段名和顺序必须与“出货单模板”相符合',pos=(450,450))
		downStatic = wx.StaticText(self,-1,u'点击下载模板',pos=(880,450))
		downStatic.SetForegroundColour('blue')
		font = wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD)
		downStatic.SetFont(font)
		tishi.SetFont(font)
		self.gau = wx.Gauge(self,-1, 100, size=(980, 3), pos=(0,483))
		
		downStatic.Bind(wx.EVT_LEFT_DOWN,self.DownStatic)
		self.fileBtn.Bind(wx.EVT_BUTTON,self.ExcelInput)
		try:
			chuhuo_Thread(self).start()
		except Exception, e:
			wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR);return
		

	def SetTextStringValue(self,msg):
		self.bianhao.SetLabel(msg[0])
		self.time.SetLabel(msg[1])
	def SetInFileOk(self):
		self.fileBtn.SetLabel(u'导入完成')
		wx.MessageBox(u'导入完成',u'提示',wx.ICON_INFORMATION)
		for i in range(self.grid.GetNumberRows()-1):
			self.grid.DeleteRows(1)
		chuhuo_Thread(self).start()

	def SetGridValue(self,gridData):
		if gridData:
			self.grid.BeginBatch()
			for i in range(self.grid.GetNumberRows()-1):
				self.grid.DeleteRows(1)
			for i,x in enumerate(gridData):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,x[5])
				self.grid.SetCellValue(i,1,str(x[2]))
				self.grid.SetCellValue(i,2,x[3])
				self.grid.SetCellValue(i,3,x[1])
				self.grid.SetCellValue(i,4,str(x[4]))
				self.grid.SetCellValue(i,5,x[6])
				self.grid.ForceRefresh()
			self.grid.DeleteRows(self.grid.GetNumberRows()-1)
			self.grid.EndBatch()


	def DownStatic(self,evt):
		ExcelFile = xlwt.Workbook(encoding='utf-8')
		table = ExcelFile.add_sheet(u'0')
		table.write(0,0,u'门店编码')
		table.write(0,1,u'商品编码')
		table.write(0,2,u'商品条码')
		table.write(0,3,u'商品名称')
		table.write(0,4,u'数量')
		table.write(0,5,u'门店名称')
		select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
		if select_dialog.ShowModal() == wx.ID_OK:
			ExcelFile.save(select_dialog.GetPath()+u"/出货单模板.xls")
			select_dialog.Destroy()

	def SetBtnText(self,text):
		self.fileBtn.SetLabel(text)
	def SetGauge(self,count):
		self.gau.SetValue(count)

	def ExcelInput(self,evt):
		dlg = wx.FileDialog(self, message=u"请选择要更新的表单文件",\
			defaultDir=os.getcwd(),defaultFile="",wildcard=u"Excel2003文件 (*.xls)|*.xls|\nExcel2007文件\
			(*.xlsx)|*.xlsx",style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			try:
				file_path = dlg.GetPaths()[0]
				file_types = file_path.split('.')
				file_type = file_types[len(file_types) - 1]
				if file_path=='':
					wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
				if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
					wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
				self.fileBtn.Enable(False)
				OutExcel_Thread(self,file_path).start()
			except Exception, e:
				wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR);return
#库存
class kucun(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		self.SetBackgroundColour('white')
		textKucun = wx.StaticText(self, -1, u"商品库存",pos=(20,10))
		textKucun.SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.search = wx.SearchCtrl(self,size=(200,-1), pos=(760,10),style=wx.TE_PROCESS_ENTER)
		self.search.SetDescriptiveText(u'请输入要搜索的内容')
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		_cols = [u'商品编码',u'商品条码',u'外箱条码',u'商品名称',u'数量',u'件数',u'规格',u'货位号']
		self.grid  = wx.grid.Grid(self,size=(980,370),pos=(0,50))
		self.grid.CreateGrid(20,8)
		width_Col = [60,120,120,310,65,75,65,60]
		for i,x in enumerate(width_Col):
			self.grid.SetColSize(i,x)
		self.grid.SetBackgroundColour('white')
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		self.grid.SetDefaultCellOverflow(False)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		self.OutBtn = wx.Button(self,-1,u'导出库存',pos=(10,440))
		tishi = wx.StaticText(self,-1,u'提示：导入的Excel文件字段名和顺序必须与“货品单模板”相符合',pos=(450,450))
		downStatic = wx.StaticText(self,-1,u'点击下载模板',pos=(880,450))
		downStatic.SetForegroundColour('blue')
		font = wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD)
		downStatic.SetFont(font)
		tishi.SetFont(font)
		self.gau = wx.Gauge(self,-1, 100, size=(980, 3), pos=(0,483))
		downStatic.Bind(wx.EVT_LEFT_DOWN,self.DownStatic)
		# self.fileBtn.Bind(wx.EVT_BUTTON,self.ExcelInput)
		self.OutBtn.Bind(wx.EVT_BUTTON,self.OutExcelThreadStar)
		kucun_Thread(self).start()


	def OutExcelThreadStar(self,evt):
		select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
		if select_dialog.ShowModal() == wx.ID_OK:
			ExcelFile = xlwt.Workbook(encoding='utf-8')
			table = ExcelFile.add_sheet(u'库存商品')
			title = ['商品编码',u'商品条码',u'外箱条码',u'商品名称',u'数量',u'规格',u'货位号']
			for i,x in enumerate(title):
				table.write(0,i,x)
			for i,x in enumerate(self.gridData):
				for i_index,x_v in enumerate(x):
					table.write(i+1,i_index,x_v)
			ExcelFile.save(select_dialog.GetPath()+u"/所有库存商品"+str(int(time.time()))+".xls")
		select_dialog.Destroy()

	def DownStatic(self,evt):
		ExcelFile = xlwt.Workbook(encoding='utf-8')
		table = ExcelFile.add_sheet(u'货品单模板')
		table.write(0,0,u'商品编码')
		table.write(0,1,u'商品名称')
		table.write(0,2,u'数量')
		table.write(0,3,u'货位号')
		select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
		if select_dialog.ShowModal() == wx.ID_OK:
			ExcelFile.save(select_dialog.GetPath()+u"/库存货品单模板.xls")
			select_dialog.Destroy()

	def ExcelInput(self,evt):pass
	def SetGridValue(self,gridData):
		if gridData:
			self.gridData = gridData
			self.grid.BeginBatch()
			for i in range(self.grid.GetNumberRows()-1):
				self.grid.DeleteRows(1)
			for i,x in enumerate(gridData):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,str(x[0]))
				self.grid.SetCellValue(i,1,x[1])
				self.grid.SetCellValue(i,2,x[2])
				self.grid.SetCellValue(i,3,x[3])
				self.grid.SetCellValue(i,4,str(x[4]))
				js = round(float(str(x[4]))/float(str(x[5])),1)
				self.grid.SetCellValue(i,5,str(js))
				self.grid.SetCellValue(i,6,str(x[5]))
				self.grid.SetCellValue(i,7,x[6])
				self.grid.ForceRefresh()
			self.grid.DeleteRows(self.grid.GetNumberRows()-1)
			self.grid.EndBatch()		
#出库数据
class chukushuju(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		self.SetBackgroundColour('white')
		self.CreateView()
		self.CalcBtn.Bind(wx.EVT_BUTTON,self.CalcBtnClick)
		self.printBtn.Bind(wx.EVT_BUTTON,self.OnPrintBtn)
		self.UpdateZhengShipinBtn.Bind(wx.EVT_BUTTON,self.OnUpdateZhengShipinBtn)
		self.ToExcelBtn.Bind(wx.EVT_BUTTON,self.OnToExcelBtn)
		self.gau = wx.Gauge(self,-1, 100, size=(980, 3), pos=(0,483))
		self.CalcBtn.Enable(False)
		chukushuju_Thread(self).start()

	def CreateView(self):
		wx.StaticText(self, -1, u"出库数据",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.bianhao = wx.StaticText(self, -1, u"单据编号：——",pos=(600,20))
		self.time = wx.StaticText(self, -1, u"时间：——",pos=(820,20))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		self.AllBtn = wx.Button(self,-1,u'所有数据',pos=(10,440))
		self.CalcBtn = wx.Button(self,-1,u'计算出库数据',pos=(105,440),size=(150,30))
		self.UpdateZhengShipinBtn = wx.Button(self,-1,u'更改整件食品位置',pos=(260,440),size=(-1,30))
		self.UpdateShipinBtn = wx.Button(self,-1,u'更改零头食品位置',pos=(380,440),size=(-1,30))
		self.UpdateJiushuiBtn = wx.Button(self,-1,u'更改酒水位置',pos=(500,440),size=(-1,30))
		self.ToExcelBtn = wx.Button(self,-1,u'导出拣选表格',pos=(600,440),size=(-1,30))
		self.printBtn = wx.Button(self,-1,u'标签打印',pos=(815,440),size=(-1,30))
		_colsSan = [u'门店编码',u'商品编码',u'商品条码',u'商品名称',u'个数',u'货位号',u'拣选人',u'完成时间']
		_colsZheng = [u'门店编码',u'商品编码',u'商品条码',u'商品名称',u'件数',u'规格',u'货位号',u'拣选人',u'完成时间']
		self.nb = wx.Notebook(self,size=(950,375),pos=(10,50))
		self.grid_zheng  = wx.grid.Grid(self.nb,size=(950,335))
		self.grid_san  = wx.grid.Grid(self.nb,size=(950,335))
		self.grid_dabao  = wx.grid.Grid(self.nb,size=(950,335))
		self.grid_zheng.CreateGrid(20,9)
		self.grid_san.CreateGrid(20,8)
		self.grid_dabao.CreateGrid(20,9)
		self.nb.AddPage(self.grid_san, u"零头食品")
		self.nb.AddPage(self.grid_zheng, u"整件酒水")
		self.nb.AddPage(self.grid_dabao, u"零头食品打包")
		width_ColSan = [70,70,120,300,70,70,70,70]
		width_ColZheng = [70,70,120,230,70,70,70,70,70]
		for i,x in enumerate(width_ColSan):
			self.grid_san.SetColSize(i,x)
		for i,x in enumerate(width_ColZheng):
			self.grid_zheng.SetColSize(i,x)
		for i,x in enumerate(_colsSan):
			self.grid_san.SetColLabelValue(i,x)
		for i,x in enumerate(_colsZheng):
			self.grid_zheng.SetColLabelValue(i,x)
		self.grid_zheng.SetDefaultCellOverflow(False)
		
		
	def SetEnableCalc(self,result):
		self.CalcBtn.Enable(False)
		self.CalcBtn.SetLabel(u'已经计算过出库数据了')
		san_index = 0
		zheng_index = 0
		zhengF_index = 0
		sandabao_index = 0
		for i in range(self.grid_san.GetNumberRows()-1):
			self.grid_san.DeleteRows(1)
		for i in range(self.grid_zheng.GetNumberRows()-1):
			self.grid_zheng.DeleteRows(1)
		for i,x in enumerate(result):
			if str(x[14])=='0':
				self.grid_san.AppendRows(1)
				self.grid_san.SetCellValue(san_index,0,str(x[9]))
				self.grid_san.SetCellValue(san_index,1,str(x[1]))
				self.grid_san.SetCellValue(san_index,2,x[2])
				self.grid_san.SetCellValue(san_index,3,x[3])
				self.grid_san.SetCellValue(san_index,4,str(x[6]))
				self.grid_san.SetCellValue(san_index,5,x[5])
				if x[4]:
					self.grid_san.SetCellValue(san_index,6,x[4])
				else:
					self.grid_san.SetCellValue(san_index,6,'')
				if x[12]:
					self.grid_san.SetCellValue(san_index,7,str(x[12]))
				else:
					self.grid_san.SetCellValue(san_index,7,'')
				san_index = san_index+1
			else:
				self.grid_zheng.AppendRows(1)
				self.grid_zheng.SetCellValue(zheng_index,0,str(x[9]))
				self.grid_zheng.SetCellValue(zheng_index,1,str(x[1]))
				self.grid_zheng.SetCellValue(zheng_index,2,x[2])
				self.grid_zheng.SetCellValue(zheng_index,3,x[3])
				js = float(str(x[6]))/float(str(x[13]))
				self.grid_zheng.SetCellValue(zheng_index,4,str(round(js,1)))
				self.grid_zheng.SetCellValue(zheng_index,5,str(x[13]))
				self.grid_zheng.SetCellValue(zheng_index,6,x[5])
				if x[4]:
					self.grid_zheng.SetCellValue(zheng_index,7,x[4])
				else:
					self.grid_zheng.SetCellValue(zheng_index,7,'')
				if x[12]:
					self.grid_zheng.SetCellValue(zheng_index,8,str(x[12]))
				else:
					self.grid_zheng.SetCellValue(zheng_index,8,'')
				zheng_index = zheng_index +1
		self.grid_zheng.DeleteRows(self.grid_zheng.GetNumberRows()-1)
		self.grid_san.DeleteRows(self.grid_san.GetNumberRows()-1)


	def OnUpdateZhengShipinBtn(self,evt):
		dlg = wx.FileDialog(self, message=u"请选择要更新的表单文件",\
			defaultDir=os.getcwd(),defaultFile="",wildcard=u"所有文件 (*.*)|*.*|Excel2003文件 (*.xls)|*.xls|\nExcel2007文件\
			(*.xlsx)|*.xlsx",style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			try:
				file_path = dlg.GetPaths()[0]
				file_types = file_path.split('.')
				file_type = file_types[len(file_types) - 1]
				if file_path=='':
					wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
				if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
					wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
				UpdateZhengShipinThread(self,file_path,self.pid).start()
			except Exception, e:
				wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR);return

	def SetUpdateZhengShipinBtn(self,title):
		self.UpdateZhengShipinBtn.SetLabel(title)
	def OnToExcelBtn(self,evt):
		JianxuanToExcelThread(self,self.pid).start()
	def SetExcelFile(self,result):
		if result:
			ExcelFile = xlwt.Workbook(encoding='utf-8')
			table = ExcelFile.add_sheet(u'0')
			_cols = [u'门店编码',u'商品编码',u'商品条码',u'商品名称',u'拣选人',u'拣选货位',u'数量',u'规格',u'状态',u'完成时间',u'类型']
			for i,x in enumerate(_cols):
				table.write(0,i,x)
			for i,x in enumerate(result):
				table.write(i+1,0,str(x[9])) #门店编码
				table.write(i+1,1,str(x[1])) #商品编码
				table.write(i+1,2,str(x[2])) #商品条码
				table.write(i+1,3,str(x[3])) #商品名称
				table.write(i+1,4,str(x[4])) #拣选人
				table.write(i+1,5,str(x[5])) #拣选货位
				table.write(i+1,6,str(x[6])) #数量
				table.write(i+1,7,str(x[13])) #规格 
				if str(x[7])=='0':
					table.write(i+1,8,u'未拣选')
				elif str(x[7])=='1':
					table.write(i+1,8,u'拣选中')
				elif str(x[7])=='2':
					table.write(i+1,8,u'已完成')
				if x[12]:
					table.write(i+1,9,time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(x[12])))
				else:
					table.write(i+1,9,'')

				if str(x[14])=='0':
					table.write(i+1,10,u'零头食品') #类型
				if str(x[14])=='1':
					table.write(i+1,10,u'整件酒水')
				if str(x[14])=='2':
					table.write(i+1,10,u'整件食品')
				if str(x[14])=='3':
					table.write(i+1,10,u'零头食品打包')
			select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
			if select_dialog.ShowModal() == wx.ID_OK:
				ExcelFile.save(select_dialog.GetPath()+u"/所有拣选数据"+time.strftime(u"%Y%m%d%H%M%S",time.localtime())+".xls")
				select_dialog.Destroy()


	def CalcBtnClick(self,evt):
		dlg = wx.MessageDialog(self, u"确保数据已无误，计算数据前请保存库存。", u"提示", wx.YES_NO)
		if dlg.ShowModal() == wx.ID_YES:
			self.CalcBtn.Enable(False)
			chukuCalcThread(self).start()
	def SetEnableTrue(self):
		self.CalcBtn.Enable(True)
	def LogMessage(self,message):
		self.CalcBtn.SetLabel(message)
	def SetGauge(self,num):
		self.gau.SetValue(num)
	def SetTitle(self,title):
		self.bianhao.SetLabel('单据编号：'+title[0][1])
		self.pid = title[0][1]
		self.time.SetLabel(u'时间：'+str(title[0][3])[0:16])
	def OnPrintBtn(self,evt):
		self.Printdlg = wx.Dialog(self,-1,u'标签打印内容框',style=wx.DEFAULT_DIALOG_STYLE)
		wx.StaticText(self.Printdlg,-1,u'标签内容:',pos=(10,10))
		wx.StaticText(self.Printdlg,-1,u'打印次数:',pos=(10,60))
		self.PrintNote = wx.TextCtrl(self.Printdlg,-1,u'',pos=(100,10))
		self.PrintCount = wx.TextCtrl(self.Printdlg,-1,u'1',pos=(100,60))
		printBtn = wx.Button(self.Printdlg,-1,u'打印',pos=(10,100))
		closeBtn = wx.Button(self.Printdlg,-1,u'关闭',pos=(120,100))
		printBtn.Bind(wx.EVT_BUTTON,self.Onprint)
		closeBtn.Bind(wx.EVT_BUTTON,self.OnDClose)
		self.Printdlg.ShowModal()
			
	def Onprint(self,evt):
		note = self.PrintNote.GetValue()
		count = self.PrintCount.GetValue()
		url = 'http://192.168.2.150/zsh/index.php/Index/PrintBiaoqian/note/%s/count/%s'%(note,count)
		webbrowser.open(url)
	def OnDClose(self,evt):
		self.Printdlg.Destroy()
#发车管理
class facheguanli(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		self.SetBackgroundColour('white')
		wx.StaticText(self, -1, u"发车管理",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		# self.weifache = wx.StaticText(self, -1, u"未发车门店：-个",pos=(560,20))
		# self.yifache = wx.StaticText(self, -1, u"已发车门店：-个，",pos=(440,20))
		self.bianhao = wx.StaticText(self, -1, u"单据编号：-",pos=(680,20))
		self.time = wx.StaticText(self, -1, u"时间：-",pos=(860,20))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		_cols = [u'车辆名称',u'车牌号',u'门店总数',u'联系人',u'联系电话',u'车身长度',u'备注',u'出库单号',u'状态']
		self.grid  = wx.grid.Grid(self,size=(980,368),pos=(0,50))
		self.grid.CreateGrid(20,9)
		self.grid.SetBackgroundColour('white')
		self.grid.SetColSize(0, 210)
		self.grid.SetColSize(1, 70)
		self.grid.SetColSize(2, 60)
		self.grid.SetColSize(3, 70)
		self.grid.SetColSize(4, 85)
		self.grid.SetColSize(5, 70)
		self.grid.SetColSize(6, 120)
		self.grid.SetColSize(7, 130)
		self.grid.SetColSize(8, 65)
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		self.HistoryBtn = wx.Button(self,-1,u'历史发车',pos=(10,440))
		self.AddBtn = wx.Button(self,-1,u'添加车辆',pos=(120,440))
		tishi = wx.StaticText(self,-1,u'提示：在行头数字双击查看车辆发运门店信息及添加发运门店。',pos=(300,450))
		font = wx.Font(12, wx.SWISS, wx.NORMAL, wx.BOLD)
		tishi.SetFont(font)
		self.AddBtn.Bind(wx.EVT_BUTTON,self.AddBtnPanel)
		# self.grid.Bind(wx.grid.EVT_GRID_LABEL_LEFT_DCLICK,self.OnCellLeftDClick) #左双击
		# self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_DCLICK,self.OnCellRightDClick)#右双击
		self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup) #行头右键点击
		facheSetGrid(self).start()
	def AddBtnPanel(self,evt):
		facheguanliFrame().Show()
	def OnShowPopup(self,evt):
		if evt.GetRow()==-1:return
		carNum =  self.grid.GetCellValue(evt.GetRow(),0)
		carNum =  self.grid.GetCellValue(evt.GetRow(),0)
		pid =  self.grid.GetCellValue(evt.GetRow(),7)
		popupmenu1=wx.Menu()
		update_pop_menu = popupmenu1.Append(-1,u'修改车辆信息')
		yunxu_pop_menu = popupmenu1.Append(-1,u'允许发车')
		wancheng_pop_menu = popupmenu1.Append(-1,u'完成发车')
		self.Bind(wx.EVT_MENU,lambda evt,mark=[pid,carNum]: self.OnUpdatePopMenu(evt,mark),update_pop_menu)
		self.Bind(wx.EVT_MENU,lambda evt,mark=[pid,carNum]: self.OnYunxuPopMenu(evt,mark),yunxu_pop_menu)
		self.Bind(wx.EVT_MENU,lambda evt,mark=[pid,carNum]: self.OnWanchengPopMenu(evt,mark),wancheng_pop_menu)
		self.grid.PopupMenu(popupmenu1)
	def OnUpdatePopMenu(self,evt,mark):
		pass

	def OnYunxuPopMenu(self,evt,mark):
		if wx.MessageDialog(None, u"确认发送车辆名称为“%s”的车次吗？"%mark[1], u"提示", wx.YES_NO | wx.ICON_QUESTION).ShowModal()==wx.ID_YES:
			updateSend_car(self,[mark[1],mark[0]]).start()
	def OnWanchengPopMenu(self,evt,mark):
		if wx.MessageDialog(None, u"确认完成车辆名称为“%s”的车次吗？"%mark[1], u"提示", wx.YES_NO | wx.ICON_QUESTION).ShowModal()==wx.ID_YES:
			if wx.MessageDialog(None, u"完成车辆名称为“%s”的车次后，该车次不在列表中显示，请确认已经完成发车？"%mark[1], u"提示", wx.YES_NO | wx.ICON_QUESTION).ShowModal()==wx.ID_YES:
				wancheng_car(self,mark).start()

	def OnCellLeftDClick(self,evt):
		if evt.GetRow()==-1:return
		carNum =  self.grid.GetCellValue(evt.GetRow(),0)
		pid =  self.grid.GetCellValue(evt.GetRow(),7)
		# facheAddMdbm([carNum,pid]).Show()
	def SetViewGrid(self,result):
		if result:
			for i,x in enumerate(result):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,x[1])
				self.grid.SetCellValue(i,1,x[0])
				self.grid.SetCellValue(i,2,str(x[9]))
				self.grid.SetCellValue(i,3,x[2])
				self.grid.SetCellValue(i,4,x[5])
				self.grid.SetCellValue(i,5,x[4])
				self.grid.SetCellValue(i,6,x[7])
				self.grid.SetCellValue(i,7,x[6])
				if x[8]==0:
					self.grid.SetCellValue(i,8,u'未发车')
				if x[8]==1:
					self.grid.SetCellValue(i,8,u'已发车')
				if x[8]==2:
					self.grid.SetCellValue(i,8,u'已完成')		
#发车管理添加车辆窗口
class facheguanliFrame(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self,None,-1,u'添加车辆',size=(800,600),style=wx.STAY_ON_TOP)
		self.SetBackgroundColour('white')
		self.Center()
		self.pidCom_list = []
		self.pidCom =  wx.ComboBox(self,500,u'',(300,50),(160,-1),[],wx.CB_DROPDOWN)
		fachePid(self).start()
		self.CreateView()
		self.closeBtn.Bind(wx.EVT_BUTTON,self.OnCloseBtn)
		self.AddBtn.Bind(wx.EVT_BUTTON,self.OnAddCar)
		self.wxpop = pop.PopupControl(self,-1,pos=(300,150),size=(200,30))
		self.win = wx.Window(self,-1,pos=(0,0),style=0,size=(1200,550))
		self.winBtn = wx.Button(self.win,-1,u'完成',pos=(650,10))
		self.winText = wx.TextCtrl(self.win,-1,pos=(10,10),size=(600,30))
		self.winBtn.Bind(wx.EVT_BUTTON,self.OnOKBtn)
		
		

	def CreateView(self):
		self.pid = wx.StaticText(self, -1, u"*出库单号:",pos=(200,50))
		self.name = wx.StaticText(self, -1, u"*车辆名称:",pos=(200,100))
		self.mdbm = wx.StaticText(self, -1, u"*关联门店：",pos=(200,150))
		self.num = wx.StaticText(self, -1, u"车 牌 号：",pos=(200,200))
		self.len_car = wx.StaticText(self, -1, u"车身长度：",pos=(200,250))
		self.person = wx.StaticText(self, -1, u"联 系 人：",pos=(200,300))
		self.phone = wx.StaticText(self, -1, u"联系电话：",pos=(200,350))
		self.note = wx.StaticText(self, -1, u"备    注：",pos=(200,400))
		self.AddBtn = wx.Button(self,-1,u'添加车辆',pos=(300,500))
		self.closeBtn = wx.Button(self,-1,u'关闭',pos=(450,500))
		wx.StaticText(self,-1,u'带"*"号为必填项！',pos=(300,470))

		self.nameText = wx.TextCtrl(self,-1,pos=(300,100),size=(160,30))
		self.numText = wx.TextCtrl(self,-1,pos=(300,200),size=(160,30))
		LenCom_list = [u"4.2米",u'6.2米',u'7.2米',u'9.6米',u'12.5米',u'17.5米']
		self.LenCom =  wx.ComboBox(self, 500, u"",(300,250),(160,-1),LenCom_list,wx.CB_DROPDOWN)
		self.personText = wx.TextCtrl(self,-1,pos=(300,300),size=(160,30))
		self.phoneText = wx.TextCtrl(self,-1,pos=(300,350),size=(160,30))
		self.noteText = wx.TextCtrl(self,-1,pos=(300,400),size=(200,50))
		

	def OnOKBtn(self,evt):
		self.wxpop.PopDown()
		self.pidCom.GetLabel()
		self.wxpop.SetValue(self.winText.GetValue()[0:-1])
		evt.Skip()
	def SetCheckList(self,result):
		pos_x = 20
		pos_y = 50
		for x in result:
			wx.CheckBox(self.win,-1,str(x[0]),pos=(pos_x,pos_y))
			pos_x += 100
			if pos_x>1200:
				pos_y += 20
				pos_x = 20
		self.win.Bind(wx.EVT_CHECKBOX,self.ONCheck)
		self.wxpop.SetPopupContent(self.win)
	def ONCheck(self,evt):
		if evt.GetEventObject().GetValue():
			self.winText.SetValue(str(evt.GetEventObject().GetLabel()+","+self.winText.GetValue()))


	def SetPid(self,result):
		for i in result:
			self.pidCom_list.append(i[0])
			self.pidCom.Append(i[0])
		self.pidCom.SetLabel(self.pidCom_list[0])
		self.nameText.SetFocus()
		facheAddMdbmThread(self,self.pidCom.GetLabel()).start()
	def OnAddCar(self,evt):
		pidCom = self.pidCom.GetLabel()
		nameText = self.nameText.GetValue()
		numText = self.numText.GetValue()
		LenCom = self.LenCom.GetString(self.LenCom.GetSelection());
		personText = self.personText.GetValue()
		phoneText = self.phoneText.GetValue()
		noteText = self.noteText.GetValue()
		mdbm = self.wxpop.GetValue()
		if nameText.strip()=='' or mdbm.strip()=='':
			wx.MessageDialog(self, u"车辆名称不能为空请重新输入！", u"提示", wx.OK| wx.ICON_INFORMATION).ShowModal();return;
		facheAddCar_Thread(self,[pidCom,nameText,numText,LenCom,personText,phoneText,noteText,mdbm]).start()
		

	def OnCloseBtn(self,evt):
		self.Close()
	def SetStatic(self,result):
		if result==1:
			self.nameText.SetValue('')
			self.nameText.SetFocus()
			wx.MessageDialog(self, u"添加车辆完成", u"提示", wx.OK| wx.ICON_INFORMATION).ShowModal();
		else:
			wx.MessageDialog(self, u"添加失败", u"提示", wx.OK| wx.ICON_INFORMATION).ShowModal();			
#发车添加门店
class facheAddMdbm(wx.Frame):
	def __init__(self,mark):
		wx.Frame.__init__(self,None,-1,mark[0],size=(800,500),style=wx.STAY_ON_TOP)
		self.SetBackgroundColour('white')
		self.Center()
		info = u'车辆名称：%s，单据编号：%s'%(mark[0],mark[1])
		info_text = wx.StaticText(self,-1,info)
		Subbtn = wx.Button(self,-1,u'提交',pos=(250,460))
		
		Closebtn = wx.Button(self,-1,u'关闭',pos=(350,460))
		sizer =wx.BoxSizer(wx.VERTICAL)
		sizer.Add(info_text, 0, wx.ALL, 10)
		self.ip = ItemsPicker(self,-1,[],u'未选门店:', u'已选门店:',ipStyle =IP_REMOVE_FROM_CHOICES)
		self.ip.Bind(EVT_IP_SELECTION_CHANGED, self.OnSelectionChange)
		Closebtn.Bind(wx.EVT_BUTTON,self.OnClose)
		self.ip._source.SetMinSize((250,350))
		sizer.Add(self.ip, 0, wx.ALL, 10)
		sizer.Add(wx.StaticText(self,-1,u''), 0, wx.ALL, 40)

		# sizer.Add(Subbtn, 0, wx.ALL, 5)
		self.SetSizer(sizer)
		self.itemCount = 3
		self.Fit()
		facheAddMdbmThread(self,mark).start()

	def OnAdd(self,e):
		items = self.ip.GetItems()
		self.itemCount += 1
		newItem = "item%d" % self.itemCount
		self.ip.SetItems(items + [newItem])
	def OnSelectionChange(self, e):
		print  e.GetItems()
	def OnClose(self,evt):
		self.Close();
	def SetItem(self,item):
		self.ip.SetItems(self.ip.GetItems() + [item])
#拆零复核
class chailingfuhe(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		wx.StaticText(self, -1, u"拆零复核",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		_cols = [u'门店编码',u'拣选总行数',u'已拣行数',u'商品总个数',u'已复核个数',u'复核完成率',u'车辆名称',u'出库单据号']
		self.grid  = wx.grid.Grid(self,size=(980,368),pos=(0,50))
		self.grid.CreateGrid(20,8)
		self.grid.SetBackgroundColour('white')
		self.grid.SetColSize(0, 100)
		self.grid.SetColSize(1, 100)
		self.grid.SetColSize(2, 100)
		self.grid.SetColSize(3, 100)
		self.grid.SetColSize(4, 90)
		self.grid.SetColSize(5, 90)
		self.grid.SetColSize(6, 150)
		self.grid.SetColSize(7, 150)
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		chailingfuheSetGrid(self).start()
		self.grid.Bind(wx.grid.EVT_GRID_LABEL_LEFT_DCLICK,self.OnCellLeftDClick) #左双击

	def SetViewGrid(self,result):
		if result:
			for i,x in enumerate(result):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,x[0])
				self.grid.SetCellValue(i,1,str(x[1]))
				self.grid.SetCellValue(i,2,str(x[2]))
				self.grid.SetCellValue(i,3,str(x[3]))
				self.grid.SetCellValue(i,4,str(x[4]))
				self.grid.SetCellValue(i,5,str(round(float(str(x[5])),2)))
				self.grid.SetCellValue(i,6,x[7])
				self.grid.SetCellValue(i,7,x[6])

	def OnCellLeftDClick(self,evt):
		if evt.GetRow()==-1:return
		mdbm =  self.grid.GetCellValue(evt.GetRow(),0)
		pid =  self.grid.GetCellValue(evt.GetRow(),7)
		url = 'http://192.168.2.150/zsh/index.php/Storeroom/send_scattered/mdbm/%s/pid/%s'%(mdbm,pid)
		webbrowser.open(url)
		# chailingfuheFrame([mdbm,pid]).Show()  #客户端复核，还未完成
#拆零复核
class chailingfuheFrame(wx.Frame):
	def __init__(self,info):
		wx.Frame.__init__(self,None,-1,u'拆零复核,出库单号：%s,门店编码：%s'%(info[1],str(info[1])),size=(1000,700),style=wx.STAY_ON_TOP)
		self.Center()
		self.SetBackgroundColour('white')
		self.info=info
		self.CreateView()
		chailingfuheFrameThread(self,info).start()

	def CreateView(self):
		wx.StaticText(self,-1,u'出库单号:%s'%self.info[1],pos=(20,20)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		wx.StaticText(self,-1,u'门店编码:%s'%self.info[0],pos=(20,60)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.sku = wx.StaticText(self,-1,u'商品编码:-',pos=(20,100))
		self.name= wx.StaticText(self,-1,u'商品名称:-',pos=(20,140))
		self.eanText= wx.TextCtrl(self,-1,u'',pos=(350,30),style=wx.BORDER_NONE|wx.TE_PROCESS_ENTER,size=(600,85))
		self.eanText.SetForegroundColour('blue')
		self.eanText.SetFont(font=wx.Font(56, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.jindu = wx.StaticText(self,-1,u'复核进度:(0/0)',pos=(20,180))
		self.jindugau = wx.Gauge(self,-1,100,pos=(250,180),size=(700,25))
		
		self.jindu.SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.sku.SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.name.SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))

		self.nb = wx.Notebook(self,size=(980,400),pos=(10,220))
		_cols = [u'商品编码',u'商品条码',u'商品名称',u'拣选人',u'货位号',u'出库个数',u'已打包数',u'相差个数']
		saomiao_cols = [u'商品编码',u'商品条码',u'商品名称',u'拣选人',u'货位号',u'相差个数',u'打包个数']
		self.grid  = wx.grid.Grid(self.nb,size=(940,360))
		self.grid_weidabao  = wx.grid.Grid(self.nb,size=(940,360))
		self.grid.CreateGrid(20,7)
		self.grid_weidabao.CreateGrid(20,8)
		self.grid.EnableDragRowSize(False)
		self.grid.EnableDragGridSize(False)
		self.grid_weidabao.EnableDragRowSize(False)
		self.grid_weidabao.EnableDragGridSize(False)
		self.grid.EnableEditing(False)
		self.grid_weidabao.EnableEditing(False)
		self.nb.AddPage(self.grid, u"已扫描商品")
		self.nb.AddPage(self.grid_weidabao, u"所有复核商品")
		self.grid_weidabao.SetBackgroundColour('white')
		grid_weidabao_width = [100,100,270,80,80,80,80,80]
		for i,x in enumerate(grid_weidabao_width):
			self.grid_weidabao.SetColSize(i, x)
		grid_width = [100,100,350,80,80,80,80]
		for i,x in enumerate(grid_width):
			self.grid.SetColSize(i, x)
		for i,x in enumerate(_cols):
			self.grid_weidabao.SetColLabelValue(i,x)
		for i,x in enumerate(saomiao_cols):
			self.grid.SetColLabelValue(i,x)
		wx.StaticLine(self,pos=(0,640),size=(1000,2))
		self.dabaoBtn = wx.Button(self,-1,u'打包',pos=(10,650))
		self.OKBtn = wx.Button(self,-1,u'完成复核',pos=(120,650))
		self.messageText = wx.TextCtrl(self,-1,pos=(240,645),size=(760,50),style=wx.TE_MULTILINE|wx.TE_RICH2| wx.TE_READONLY | wx.TE_MULTILINE | wx.BORDER_NONE)
		self.dabaoBtn.Bind(wx.EVT_BUTTON,self.OnDabaoBtn)
		self.OKBtn.Bind(wx.EVT_BUTTON,self.OnOKBtn)
		self.eanText.Bind(wx.EVT_TEXT_ENTER,lambda evt,mark=[self.info[1],self.info[0]]:self.OnInput(evt,mark))
		self.eanText.SetFocus()

	def ThreadSetStatic(self,count,sku,name):
		self.gauCount = self.gauCount+count
		self.jindugau.SetValue(self.gauCount)
		self.jindu.SetLabel(u'复核进度:('+str(self.gauCount)+'/'+str(self.zcount)+")")
		self.sku.SetLabel(u'商品编码：%s'%sku)
		self.name.SetLabel(u'商品名称：%s'%name)

	def scanViewGrid(self,result,Addresult):
		self.grid.BeginBatch()
		self.grid_weidabao.BeginBatch()
		for x in range(self.grid.GetNumberRows()-1):
			self.grid.DeleteRows(1)
		for x in range(self.grid_weidabao.GetNumberRows()-1):
			self.grid_weidabao.DeleteRows(1)
		if Addresult:
			for i,[key,x] in enumerate(self.Addresult.items()):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,str(int(float(str(x[0])))))
				self.grid.SetCellValue(i,1,x[1])
				self.grid.SetCellValue(i,2,x[2])
				self.grid.SetCellValue(i,3,x[3])
				self.grid.SetCellValue(i,4,x[4])
				self.grid.SetCellValue(i,5,str(x[5]))
				self.grid.SetCellValue(i,6,str(x[6]))
		if result:
			for i,[key,x] in enumerate(self.result.items()):
				self.grid_weidabao.AppendRows(1)
				self.grid_weidabao.SetCellValue(i,0,str(int(float(str(x[0])))))
				self.grid_weidabao.SetCellValue(i,1,x[1])
				self.grid_weidabao.SetCellValue(i,2,x[2])
				if x[3]:
					self.grid_weidabao.SetCellValue(i,3,x[3])
				else:
					self.grid_weidabao.SetCellValue(i,3,'')
				self.grid_weidabao.SetCellValue(i,4,x[4])
				self.grid_weidabao.SetCellValue(i,5,str(int(float(str(x[5])))))
				self.grid_weidabao.SetCellValue(i,6,str(int(float(str(x[6])))))
				self.grid_weidabao.SetCellValue(i,7,str(int(float(str(x[7])))))
				self.xiangchaCount = self.xiangchaCount+int(float(str(x[7])))
		self.grid.EndBatch()
		self.grid_weidabao.EndBatch()
		

	def SetViewGrid(self,result,Addresult ={}):
		self.Addresult = {}
		if result:
			self.result = {}
			for x in result:
				self.result[str(x[1])] = list(x)
			self.zcount = 0
			self.xiangchaCount = 0
			for i,[key,x] in enumerate(self.result.items()):
				self.grid_weidabao.AppendRows(1)
				self.grid_weidabao.SetCellValue(i,0,str(int(float(str(x[0])))))
				self.grid_weidabao.SetCellValue(i,1,x[1])
				self.grid_weidabao.SetCellValue(i,2,x[2])
				if x[3]:
					self.grid_weidabao.SetCellValue(i,3,x[3])
				else:
					self.grid_weidabao.SetCellValue(i,3,'')
				self.grid_weidabao.SetCellValue(i,4,x[4])
				self.grid_weidabao.SetCellValue(i,5,str(int(float(str(x[5])))))
				self.grid_weidabao.SetCellValue(i,6,str(int(float(str(x[6])))))
				self.grid_weidabao.SetCellValue(i,7,str(int(float(str(x[7])))))
				self.zcount = int(float(str(x[5])))+self.zcount
				self.xiangchaCount = self.xiangchaCount+int(float(str(x[7])))
			self.jindu.SetLabel(u'复核进度:('+str(self.zcount-self.xiangchaCount)+'/'+str(self.zcount)+")")
			self.jindugau.SetRange(int(self.zcount))
			self.jindugau.SetValue(int(self.zcount-self.xiangchaCount))
			self.mendianzcount = self.zcount-self.xiangchaCount
			self.gauCount = self.zcount-self.xiangchaCount
			

	def UpdateOk(self):
		wx.MessageDialog(self,u'发送完成',u'提示',wx.OK | wx.ICON_INFORMATION).ShowModal();
		self.messageText.AppendText(time.strftime(u"%Y-%m-%d %H:%M:%S",time.localtime())+u' 开始复核,发送拣选信息.')
	def OnDabaoBtn(self,evt):
		pass


	def OnInput(self,evt,mark):
		ean = self.eanText.GetValue().strip()
		if ean:
			try:
				self.SetEanInputNull()
				chailingfuheInput(self,self.result,self.Addresult,ean,1).start()
			except Exception, e:
				wx.MessageDialog(self,u'系统发生错误：%s，请联系管理员。'%str(e),u'提示',wx.OK | wx.ICON_INFORMATION).ShowModal();
			

	def ThreadMessage(self,message):
		self.messageText.AppendText("\n"+time.strftime(u"%Y-%m-%d %H:%M:%S",time.localtime())+u' %s'%message)
		self.SetEanInputNull()
	def SuccessMessage(self,data,ean):
		self.SetEanInputNull()
		self.mendianzjs = self.mendianzjs+1
		self.jindugau.SetValue(self.mendianzjs)
		self.jindu.SetLabel(u'复核进度:('+str(self.mendianzjs)+'/'+str(self.zjs)+")")
		self.sku.SetLabel(u'商品编码:%s'%str(data[0][1]))
		self.unit.SetLabel(u'规格:%s'%str(data[0][13]))
		self.name.SetLabel(u'商品名称:%s'%data[0][3])
		zjs = round(float(str(data[0][6]))/float(str(data[0][13])),2)
		js = round(float(str(data[0][15]))/float(str(data[0][13])),2)+1
		self.messageText.AppendText(time.strftime("\n"+u"%Y-%m-%d %H:%M:%S",time.localtime())+u'键入 "%s" 成功，商品:%s 当前复核第%s件,共计%s件.'%(ean,data[0][3],js,zjs))

	def OnOKBtn(self,evt):
		if wx.MessageDialog(self,u'完成门店复核，并关闭当前窗口？',u'提示',wx.YES_NO ).ShowModal()==wx.ID_YES:
			self.Close()
	def SetEanInputNull(self):
		self.eanText.SetValue('')
		self.eanText.SetFocus()
#整件复核
class zhengjianfuhe(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		wx.StaticText(self, -1, u"整件复核",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		_cols = [u'门店编码',u'拣选总行数',u'已拣行数',u'商品总件数',u'已复核件数',u'复核完成率',u'车辆名称',u'出库单据号']
		self.grid  = wx.grid.Grid(self,size=(980,368),pos=(0,50))
		self.grid.CreateGrid(20,8)
		self.grid.SetBackgroundColour('white')
		self.grid.SetColSize(0, 100)
		self.grid.SetColSize(1, 100)
		self.grid.SetColSize(2, 100)
		self.grid.SetColSize(3, 100)
		self.grid.SetColSize(4, 90)
		self.grid.SetColSize(5, 90)
		self.grid.SetColSize(6, 150)
		self.grid.SetColSize(7, 150)
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		zhengjianfuheSetGrid(self).start()
		self.grid.Bind(wx.grid.EVT_GRID_LABEL_LEFT_DCLICK,self.OnCellLeftDClick) #左双击

	def SetViewGrid(self,result):
		if result:
			for i,x in enumerate(result):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,x[0])
				self.grid.SetCellValue(i,1,str(x[1]))
				self.grid.SetCellValue(i,2,str(x[2]))
				self.grid.SetCellValue(i,3,str(round(float(str(x[3])),2)))
				self.grid.SetCellValue(i,4,str(round(float(str(x[4])),2)))
				self.grid.SetCellValue(i,5,str(round(float(str(x[5])),2)))
				self.grid.SetCellValue(i,6,x[7])
				self.grid.SetCellValue(i,7,x[6])
	def OnCellLeftDClick(self,evt):
		if evt.GetRow()==-1:return
		mdbm =  self.grid.GetCellValue(evt.GetRow(),0)
		pid =  self.grid.GetCellValue(evt.GetRow(),7)
		zhengjianfuheFrame([mdbm,pid]).Show()
class zhengjianfuheFrame(wx.Frame):
	def __init__(self,info):
		wx.Frame.__init__(self,None,-1,u'出库单号：%s,门店编码：%s'%(info[1],str(info[1])),size=(1000,700),style=0)
		self.Center()
		self.SetBackgroundColour('white')
		self.info=info
		self.CreateView()
		zhengjianfuheFrameThread(self,info).start()

		
	def CreateView(self):
		wx.StaticText(self,-1,u'出库单号:%s'%self.info[1],pos=(20,20)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		wx.StaticText(self,-1,u'门店编码:%s'%self.info[0],pos=(20,60)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.sku = wx.StaticText(self,-1,u'商品编码:-',pos=(20,100))
		self.unit= wx.StaticText(self,-1,u'规格:-',pos=(210,100))
		self.name= wx.StaticText(self,-1,u'商品名称:-',pos=(20,140))
		self.eanText= wx.TextCtrl(self,-1,u'',pos=(350,30),style=wx.BORDER_NONE|wx.TE_PROCESS_ENTER,size=(600,85))
		self.eanText.SetForegroundColour('blue')
		self.eanText.SetFont(font=wx.Font(56, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.jindu = wx.StaticText(self,-1,u'复核进度:(0/0)',pos=(20,180))
		self.jindugau = wx.Gauge(self,-1,100,pos=(250,180),size=(700,25))
		
		self.jindu.SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.sku.SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.unit.SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		self.name.SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))

		_cols = [u'商品编码',u'商品条码',u'商品名称',u'拣选人',u'货位号',u'件数',u'规格',u'相差件数']
		self.grid  = wx.grid.Grid(self,size=(1000,410),pos=(0,220))
		self.grid.CreateGrid(20,8)
		self.grid.SetBackgroundColour('white')
		self.grid.SetColSize(0, 100)
		self.grid.SetColSize(1, 100)
		self.grid.SetColSize(2, 290)
		self.grid.SetColSize(3, 80)
		self.grid.SetColSize(4, 80)
		self.grid.SetColSize(5, 80)
		self.grid.SetColSize(6, 80)
		self.grid.SetColSize(7, 80)
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		wx.StaticLine(self,pos=(0,640),size=(1000,2))
		self.sendBtn = wx.Button(self,-1,u'开始复核',pos=(10,650))
		self.OKBtn = wx.Button(self,-1,u'完成复核',pos=(120,650))
		self.messageText = wx.TextCtrl(self,-1,pos=(240,645),size=(760,50),style=wx.TE_MULTILINE|wx.TE_RICH2| wx.TE_READONLY | wx.TE_MULTILINE | wx.BORDER_NONE)
		self.sendBtn.Bind(wx.EVT_BUTTON,self.OnBeginBtn)
		self.OKBtn.Bind(wx.EVT_BUTTON,self.OnOKBtn)
		self.eanText.Bind(wx.EVT_TEXT_ENTER,lambda evt,mark=[self.info[1],self.info[0]]:self.OnInput(evt,mark))
		self.messageText.SetFocus()
		self.grid.EnableEditing(False)
		self.grid.EnableDragRowSize(False)
		self.grid.EnableDragGridSize(False)

	def SetViewGrid(self,result,baseSku,ean=''):
		self.grid.BeginBatch()
		self.baseSku = baseSku
		self.result = copy.deepcopy(result)
		sorted(self.result.iteritems(),key=operator.itemgetter(1))  
		for x in range(self.grid.GetNumberRows()-1):
			self.grid.DeleteRows(1)
		if result:
			self.zjs = 0
			self.xiangchajs = 0
			for i,[key,x] in enumerate(self.result.items()):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,str(int(float(str(x[0][0])))))
				self.grid.SetCellValue(i,1,x[0][1])
				self.grid.SetCellValue(i,2,x[0][2])
				if x[0][3]:
					self.grid.SetCellValue(i,3,x[0][3])
				else:
					self.grid.SetCellValue(i,3,'')
				self.grid.SetCellValue(i,4,x[0][4])
				self.grid.SetCellValue(i,5,str(int(float(str(x[0][5])))))
				self.grid.SetCellValue(i,6,str(int(float(str(x[0][6])))))
				self.grid.SetCellValue(i,7,str(int(float(str(x[0][7])))))
				self.zjs = int(float(str(x[0][5])))+self.zjs
				self.xiangchajs = self.xiangchajs+int(float(str(x[0][7])))
			self.jindu.SetLabel(u'复核进度:('+str(self.zjs-self.xiangchajs)+'/'+str(self.zjs)+")")
			self.jindugau.SetRange(int(self.zjs))
			self.jindugau.SetValue(int(self.zjs-self.xiangchajs))
			self.mendianzjs = self.zjs-self.xiangchajs
		self.grid.EndBatch()
			

	def OnBeginBtn(self,evt):
		self.eanText.SetFocus()
		self.sendBtn.Enable(False)
		zhengjianfuheSendBtn(self,self.info).start()
	def UpdateOk(self):
		wx.MessageDialog(self,u'发送完成',u'提示',wx.OK | wx.ICON_INFORMATION).ShowModal();
		self.messageText.AppendText(time.strftime(u"%Y-%m-%d %H:%M:%S",time.localtime())+u' 开始复核,发送拣选信息.')

	def OnInput(self,evt,mark):
		ean = self.eanText.GetValue().strip()
		if ean:
			try:
				# zhengjianfuheInput(self,[mark[0],mark[1]],ean,self.baseSku).start()
				zhengjianfuheInput(self,mark,ean,self.result,self.baseSku).start()
			except Exception, e:
				wx.MessageDialog(self,u'系统发生错误：%s，请联系管理员。'%str(e),u'提示',wx.OK | wx.ICON_INFORMATION).ShowModal();
			

	def ErrorMessage(self,message):
		self.messageText.AppendText("\n"+time.strftime(u"%Y-%m-%d %H:%M:%S",time.localtime())+u' %s'%message)
		self.SetEanInputNull()
	def SuccessMessage(self,data,ean):
		self.SetEanInputNull()
		self.mendianzjs = self.mendianzjs+1
		self.jindugau.SetValue(self.mendianzjs)
		self.jindu.SetLabel(u'复核进度:('+str(self.mendianzjs)+'/'+str(self.zjs)+")")
		self.sku.SetLabel(u'商品编码:%s'%str(data[0]))
		self.unit.SetLabel(u'规格:%s'%str(data[6]))
		self.name.SetLabel(u'商品名称:%s'%data[2])
		zjs = round(data[5],1)
		js = round(float(str(data[5]))-float(str(data[7])),2)
		self.messageText.AppendText(time.strftime("\n"+u"%Y-%m-%d %H:%M:%S",time.localtime())+u'键入 "%s(%s)" 成功,共计%s件,当前复核第%s件.'%(ean,data[2],str(zjs),str(js)))

	def OnOKBtn(self,evt):
		if wx.MessageDialog(self,u'完成门店复核，并关闭当前窗口？',u'提示',wx.YES_NO ).ShowModal()==wx.ID_YES:
			self.Close()
	def SetEanInputNull(self):
		self.eanText.SetValue('')
		self.eanText.SetFocus()
#基础数据
class baseSku(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		wx.StaticText(self, -1, u"商品数据管理",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		self.search = wx.SearchCtrl(self,size=(200,-1), pos=(760,10),style=wx.TE_PROCESS_ENTER)
		self.search.SetDescriptiveText(u'请输入要搜索的内容')
		_cols = [u'商品编码',u'商品条码',u'外箱条码',u'商品名称',u'规格',u'类型']
		self.grid  = wx.grid.Grid(self,size=(980,368),pos=(0,50))
		self.grid.CreateGrid(20,6)
		self.grid.SetBackgroundColour('white')
		self.grid.SetColSize(0, 100)
		self.grid.SetColSize(1, 100)
		self.grid.SetColSize(2, 100)
		self.grid.SetColSize(3, 380)
		self.grid.SetColSize(4, 100)
		self.grid.SetColSize(5, 100)
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		self.InFileBtn = wx.Button(self,-1,u'导入数据表格',pos=(10,440))
		self.ToFileBtn = wx.Button(self,-1,u'导出数据表格',pos=(110,440))
		self.AddBtn = wx.Button(self,-1,u'添加商品数据',pos=(210,440))
		self.InFileBtn.Bind(wx.EVT_BUTTON,self.OnInFileBtn)
		self.AddBtn.Bind(wx.EVT_BUTTON,self.OnAddBtn)
		self.ToFileBtn.Bind(wx.EVT_BUTTON,self.OnOutFileBtn)
		baseSkuViewGridThread(self).start()
		self.grid.Bind(wx.grid.EVT_GRID_LABEL_LEFT_DCLICK,self.OnCellLeftDClick) #左双击
		self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_DCLICK,self.OnCellRightDClick)#右双击
		self.search.Bind(wx.EVT_TEXT_ENTER, self.OnDoSearch)
		self.grid.EnableEditing(False) #是否可编辑
		self.grid.EnableDragColSize(enable=False)#控制列宽是否可以拉动  
		self.grid.EnableDragRowSize(enable=False)#控制行高是否可以拉动


	def OnDoSearch(self,evt):
		print self.search.GetValue()
	def OnAddBtn(self,evt):
		self.Adddlg = wx.Dialog(self,-1,u'添加商品数据',style=wx.DEFAULT_DIALOG_STYLE)
		wx.StaticText(self.Adddlg,-1,u'商品编码:',pos=(10,10))
		wx.StaticText(self.Adddlg,-1,u'商品条码:',pos=(10,40))
		wx.StaticText(self.Adddlg,-1,u'外箱条码:',pos=(10,70))
		wx.StaticText(self.Adddlg,-1,u'商品名称:',pos=(10,100))
		wx.StaticText(self.Adddlg,-1,u'商品规格:',pos=(10,130))
		wx.StaticText(self.Adddlg,-1,u'商品类型:',pos=(10,160))
		self.AddSku = wx.TextCtrl(self.Adddlg,-1,pos=(70,8))
		self.AddEan = wx.TextCtrl(self.Adddlg,-1,pos=(70,38))
		self.AddXm = wx.TextCtrl(self.Adddlg,-1,pos=(70,68))
		self.AddName = wx.TextCtrl(self.Adddlg,-1,pos=(70,98),size=(300,25))
		self.AddUnit = wx.SpinCtrl(self.Adddlg, -1, "", pos=(70,128),style=wx.BORDER_NONE)
		self.AddUnit.SetRange(1,10000)
		self.AddUnit.SetValue(1)
		self.AddType =  wx.Choice(self.Adddlg, -1, pos=(70, 158),choices = [u'食品',u'酒水'])
		saveBtn = wx.Button(self.Adddlg,-1,u'保存',pos=(150,188))
		saveBtn.Bind(wx.EVT_BUTTON,self.OnSaveBtn)
		self.AddType.SetSelection(0)
		self.Adddlg.ShowModal()
	def OnSaveBtn(self,evt):
		sku = self.AddSku.GetValue().strip()
		ean = self.AddEan.GetValue().strip()
		xm = self.AddXm.GetValue().strip()
		name = self.AddName.GetValue().strip()
		unit = self.AddUnit.GetValue()
		skuType = self.AddType.GetSelection()
		if sku=='' or ean=='' or xm=='' or name=='' or unit =='' or skuType=='':
			wx.MessageBox(u'请输入完整的信息。',u'提示',wx.ICON_ERROR);return
		Base_skuSaveThread(self,[sku,ean,xm,name,unit,skuType]).start()
	def AddBaseSkuMessage(self,message):
		wx.MessageBox(message,u'提示',wx.ICON_INFORMATION);
		self.Adddlg.Destroy()
	def OnOutFileBtn(self,evt):
		ExcelFile = xlwt.Workbook(encoding='utf-8')
		table = ExcelFile.add_sheet(u'0')
		_cols = [u'商品编码',u'商品条码',u'外箱条码',u'商品名称',u'规格',u'类型']
		for i,x in enumerate(_cols):
			table.write(0,i,x)
		for i,x in enumerate(self.Outresult):
			table.write(i+1,0,str(x[2]))
			table.write(i+1,1,str(x[3]))
			table.write(i+1,2,str(x[4]))
			table.write(i+1,3,str(x[1]))
			table.write(i+1,4,str(x[5]))
			table.write(i+1,5,str(x[8]))
		
		select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
		if select_dialog.ShowModal() == wx.ID_OK:
			ExcelFile.save(select_dialog.GetPath()+u"/商品数据导出.xls")
			select_dialog.Destroy()

	def OnCellLeftDClick(self,evt):
		self.Printdlg = wx.Dialog(self,-1,u'修改商品数据',style=wx.DEFAULT_DIALOG_STYLE)
		if evt.GetRow()==-1:return
		sku =  self.grid.GetCellValue(evt.GetRow(),0)
		ean =  self.grid.GetCellValue(evt.GetRow(),1)
		xiangma =  self.grid.GetCellValue(evt.GetRow(),2)
		name =  self.grid.GetCellValue(evt.GetRow(),3)
		unit =  self.grid.GetCellValue(evt.GetRow(),4)
		skutype =  self.grid.GetCellValue(evt.GetRow(),5)
		wx.StaticText(self.Printdlg,-1,u'商品编码:',pos=(10,10))
		wx.StaticText(self.Printdlg,-1,u'商品条码:',pos=(10,40))
		wx.StaticText(self.Printdlg,-1,u'外箱条码:',pos=(10,70))
		wx.StaticText(self.Printdlg,-1,u'商品名称:',pos=(10,100))
		wx.StaticText(self.Printdlg,-1,u'商品规格:',pos=(10,130))
		wx.StaticText(self.Printdlg,-1,u'商品类型:',pos=(10,160))
		self.updateSku = wx.TextCtrl(self.Printdlg,-1,sku,pos=(70,8),style=wx.BORDER_NONE)
		self.updateEan = wx.TextCtrl(self.Printdlg,-1,ean,pos=(70,38))
		self.updateXm = wx.TextCtrl(self.Printdlg,-1,xiangma,pos=(70,68))
		self.updateName = wx.TextCtrl(self.Printdlg,-1,name,pos=(70,98),size=(300,25))
		self.updateSku.Enable(False)
		self.updateUnit = wx.SpinCtrl(self.Printdlg, -1, "", pos=(70,128),style=wx.BORDER_NONE)
		self.updateUnit.SetRange(1,10000)
		self.updateUnit.SetValue(int(unit))
		self.updateType =  wx.Choice(self.Printdlg, -1, pos=(70, 158),choices = [u'食品',u'酒水'])
		saveBtn = wx.Button(self.Printdlg,-1,u'更新',pos=(150,188))
		oldData = [sku,ean,xiangma,name,unit,skutype]
		saveBtn.Bind(wx.EVT_BUTTON,lambda evt,mark=oldData:self.OnUpdateRow(evt,mark))
		if skutype==u'酒水':
			self.updateType.SetSelection(1)
		else:
			self.updateType.SetSelection(0)
		
		self.Printdlg.ShowModal()

	def OnUpdateRow(self,evt,data):
		newData = [self.updateSku.GetValue(),self.updateEan.GetValue(),self.updateXm.GetValue(),self.updateName.GetValue(),self.updateUnit.GetValue(),self.updateType.GetSelection()]
		BaskuUpdateBtnThread(self,[data,newData]).start()


	def OnCellRightDClick(self,evt):
		dlg = wx.MessageDialog(None, u"确认删除该行数据吗？%s"%self.grid.GetCellValue(evt.GetRow(),3), u"提示", wx.YES_NO | wx.ICON_QUESTION)
		if dlg.ShowModal() == wx.ID_YES:
			if evt.GetRow()==-1:return
			sku =  self.grid.GetCellValue(evt.GetRow(),0)
			ean =  self.grid.GetCellValue(evt.GetRow(),1)
			xiangma =  self.grid.GetCellValue(evt.GetRow(),2)
			unit =  self.grid.GetCellValue(evt.GetRow(),4)
			baseSkuDeleteRow(self,[sku,ean,xiangma,unit]).start()
	def DeleteRowOk(self,result):
		wx.MessageBox(u'已删除%s行数据'%str(result),u'提示',wx.ICON_INFORMATION)
		# baseSkuViewGridThread(self).start()
	def UpdateRowOk(self,result):
		wx.MessageBox(u'已更新%s行数据'%str(result),u'提示',wx.ICON_INFORMATION)
		self.Printdlg.Destroy()
		# baseSkuViewGridThread(self).start()


	def OnInFileBtn(self,evt):
		dlg = wx.FileDialog(self, message=u"请选择要更新的表单文件",\
			defaultDir=os.getcwd(),defaultFile="",wildcard=u"Excel2003文件 (*.xls)|*.xls|\nExcel2007文件\
			(*.xlsx)|*.xlsx",style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			try:
				file_path = dlg.GetPaths()[0]
				file_types = file_path.split('.')
				file_type = file_types[len(file_types) - 1]
				if file_path=='':
					wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
				if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
					wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
				BaskuFileBtnThread(self,file_path).start()
			except Exception, e:
				wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR);return
	def SetOkText(self):
		self.InFileBtn.SetLabel(u'导入完成')
	def SetViewGrid(self,result):
		self.Outresult = result
		if result:
			for i in range(19):
				self.grid.DeleteRows(1)
			for i,x in enumerate(result):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,str(x[2]))
				self.grid.SetCellValue(i,1,x[3])
				self.grid.SetCellValue(i,2,str(x[4]))
				self.grid.SetCellValue(i,3,x[1])
				self.grid.SetCellValue(i,4,str(x[5]))
				if str(x[8])=='0':
					self.grid.SetCellValue(i,5,u'食品')
				else:
					self.grid.SetCellValue(i,5,u'酒水')
			self.grid.DeleteRows(len(result))
#货位
class allocation(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		wx.StaticText(self, -1, u"货位管理",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		self.search = wx.SearchCtrl(self,size=(200,-1), pos=(760,10),style=wx.TE_PROCESS_ENTER)
		self.search.SetDescriptiveText(u'请输入要搜索的货位名称')
		_cols = [u'货位名称',u'货位类型',u'货位排序',u'货位区域']
		self.grid  = wx.grid.Grid(self,size=(980,368),pos=(0,50))
		self.grid.CreateGrid(20,4)
		self.grid.SetBackgroundColour('white')
		self.grid.SetColSize(0, 245)
		self.grid.SetColSize(1, 235)
		self.grid.SetColSize(2, 200)
		self.grid.SetColSize(3, 200)
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		self.InFileBtn = wx.Button(self,-1,u'导入货位表格',pos=(10,440))
		self.OutFileBtn = wx.Button(self,-1,u'导出货位表格',pos=(110,440))
		allocationGridThread(self).start()
		self.OutFileBtn.Bind(wx.EVT_BUTTON,self.OnOutFileBtn)
		self.InFileBtn.Bind(wx.EVT_BUTTON,self.OnInFileBtn)

	def SetViewGrid(self,result):
		if result:
			self.OutResult = result
			for i,x in enumerate(result):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,x[1])
				self.grid.SetCellValue(i,1,x[2])
				self.grid.SetCellValue(i,2,str(x[3]))
				self.grid.SetCellValue(i,3,x[4])

	def OnOutFileBtn(self,evt):
		ExcelFile = xlwt.Workbook(encoding='utf-8')
		table = ExcelFile.add_sheet(u'0')
		_cols = [u'货位名称',u'货位类型',u'货位排序',u'货位区域']
		for i,x in enumerate(_cols):
			table.write(0,i,x)
		for i,x in enumerate(self.OutResult):
			table.write(i+1,0,x[1])
			table.write(i+1,1,x[2])
			table.write(i+1,2,str(x[3]))
			table.write(i+1,3,x[4])
		
		select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
		if select_dialog.ShowModal() == wx.ID_OK:
			ExcelFile.save(select_dialog.GetPath()+u"/货位数据导出.xls")
			select_dialog.Destroy()

	def OnInFileBtn(self,evt):
		dlg = wx.FileDialog(self, message=u"请选择要更新的表单文件",\
			defaultDir=os.getcwd(),defaultFile="",wildcard=u"Excel2003文件 (*.xls)|*.xls|\nExcel2007文件\
			(*.xlsx)|*.xlsx",style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			try:
				file_path = dlg.GetPaths()[0]
				file_types = file_path.split('.')
				file_type = file_types[len(file_types) - 1]
				if file_path=='':
					wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
				if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
					wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
				AllocationFileAddThread(self,file_path).start()
			except Exception, e:
				wx.MessageBox(u'系统错误，错误信息：%s'%e,u'提示',wx.ICON_ERROR);return
	def SetOkText(self):
		wx.MessageBox(u'添加完成',u'提示',wx.ICON_INFORMATION)
#打包
class dabao(wx.Panel):
	def __init__(self,parent):
		wx.Panel.__init__(self,parent)
		wx.StaticText(self, -1, u"打包管理",pos=(20,10)).SetFont(font=wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD))
		wx.StaticLine(self,pos=(20,40),size=(940,2))
		self.search = wx.SearchCtrl(self,size=(200,-1), pos=(760,10),style=wx.TE_PROCESS_ENTER)
		self.search.SetDescriptiveText(u'请输入要搜索的内容')
		_cols = [u'商品编码',u'商品条码',u'商品名称',u'打包人员',u'打包数量',u'地堆位置',u'门店编码',u'打包时间',u'出库单号']
		self.grid  = wx.grid.Grid(self,size=(980,368),pos=(0,50))
		self.grid.CreateGrid(20,9)
		self.grid.SetBackgroundColour('white')
		width_list = [70,110,190,70,70,70,70,100,130] 
		for i,x in enumerate(width_list):
			self.grid.SetColSize(i, x)
		for i,x in enumerate(_cols):
			self.grid.SetColLabelValue(i,x)
		wx.StaticLine(self,pos=(0,430),size=(980,2))
		self.OutFileBtn = wx.Button(self,-1,u'导出表格',pos=(10,440))
		# self.OutFileBtn = wx.Button(self,-1,u'导出货位表格',pos=(110,440))
		DabaoGridThread(self).start()
		self.OutFileBtn.Bind(wx.EVT_BUTTON,self.OnOutFileBtn)

	def SetViewGrid(self,result):
		if result:
			self.result = result
			self.grid.BeginBatch()
			for i,x in enumerate(self.result):
				self.grid.AppendRows(1)
				self.grid.SetCellValue(i,0,str(x[1]))
				if x[2]:
					self.grid.SetCellValue(i,1,x[2])
				else:
					self.grid.SetCellValue(i,1,'')
				self.grid.SetCellValue(i,2,x[3])
				self.grid.SetCellValue(i,3,x[4])
				self.grid.SetCellValue(i,4,str(x[6]))
				self.grid.SetCellValue(i,5,str(x[9]))
				self.grid.SetCellValue(i,6,str(x[8]))
				self.grid.SetCellValue(i,7,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(float(x[10][0:10])))[2:16])
				self.grid.SetCellValue(i,8,x[7])
			self.grid.EndBatch()
	def OnOutFileBtn(self,evt):
		ExcelFile = xlwt.Workbook(encoding='utf-8')
		table = ExcelFile.add_sheet(u'0')
		_cols = [u'商品编码',u'商品条码',u'商品名称',u'打包人员',u'打包数量',u'地堆位置',u'门店编码',u'打包时间',u'出库单号']
		for i,x in enumerate(_cols):
			table.write(0,i,x)
		for i,x in enumerate(self.result):
			self.grid.AppendRows(1)
			table.write(i+1,0,str(x[1]))
			if x[2]:
				table.write(i+1,1,str(x[2]))
			else:
				table.write(i+1,1,'')
			table.write(i+1,2,str(x[3]))
			table.write(i+1,3,str(x[4]))
			table.write(i+1,4,str(x[6]))
			table.write(i+1,5,str(x[9]))
			table.write(i+1,6,str(x[8]))
			table.write(i+1,7,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(float(x[10][0:10])))[2:16])
			table.write(i+1,8,str(x[7]))

			
		select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
		if select_dialog.ShowModal() == wx.ID_OK:
			ExcelFile.save(select_dialog.GetPath()+u"/打包数据"+time.strftime(u"%Y%m%d%H%M%S",time.localtime())+".xls")
			select_dialog.Destroy()



				
				


		

		

