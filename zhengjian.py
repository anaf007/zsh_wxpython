#coding=utf-8
import wx,xlwt,time,xlrd,threading,webbrowser,winsound
from Public.GridData import *
from SQLet import *
from operator import itemgetter
import sys
sys.getdefaultencoding()
class Zhengjian(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'整件复核选择车辆',pos=(0,0))
        pnl = wx.Panel(self)
        _cols = (u'车辆',u'拣选总行数',u'完成行数',u'进度')
        self._data =[]

        #执行sql 查询 车辆 总行数  进度
        result_goods = Connect()._sql_query('*','select c.carnum,count(s.id),\
		sum(CASE when s.static=2 then 1 else 0 end) yiwancheng,\
		sum(CASE when s.static=2 then 1 else 0 end)/count(s.id) wanchenglv\
		from send_car c \
		join send_scattered s \
		where  s.mdbm=c.mdbm AND\
		c.pid = s.pid and \
		c.pid = (select pid from send_car  ORDER BY pid desc limit 1) AND\
		s.send_type in (1,2,3) \
		GROUP BY c.carNum\
        ORDER BY wanchenglv,s.mdbm ;\
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

        #根据车辆 查询门店 总行数完成数 比例
        result_goods = Connect()._sql_query('*','\
		select c.mdbm,count(s.id),\
		sum(CASE when s.static=2 then 1 else 0 end) yiwancheng,\
		sum(CASE when s.static=2 then 1 else 0 end)/count(s.id) wanchenglv\
		from send_car c \
		join send_scattered s \
		where  s.mdbm=c.mdbm AND\
		c.pid = s.pid and \
		c.pid = (select pid from send_car  ORDER BY pid desc limit 1) AND\
		s.send_type in (1,2,3)  AND\
		c.carNum = "%s"\
		GROUP BY c.mdbm\
        ORDER BY wanchenglv,s.mdbm ;\
        	'%mark)
        _cols = (u'门店编码',u"商品总行数",u'完成行数',u'进度')
        self._data =[]
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
        mark_data_list.append([self._data[evt_value][0]])
        popupmenu1=wx.Menu()
        delete_pop_menu = popupmenu1.Append(-1,u'打开')
        self.Bind(wx.EVT_MENU,lambda evt,mark=mark_data_list : self.Onopenpopmenu(evt,mark) ,delete_pop_menu )
        self.grid.PopupMenu(popupmenu1)
    def Onopenpopmenu(self,evt,mark):
        fuhe_mdbm(mark).Show()
    
        


#整件复核界面
class fuhe_mdbm(wx.Frame):
    def __init__(self,mark):
        wx.Frame.__init__(self,None,-1,u'当前复核门店：%s'%mark[0][0],pos=(0,0))
        pnl = wx.Panel(self)
        display=wx.DisplaySize()
        #设计跟昆明一鸟样界面把
        """发送按钮 输入框
        门店编码  门店地址  联系电话
        总件数   地址
        商品编码 规格 
        商品条码 当前累计件数+1
        名称
        单品进度（）
        门店进度（）
        """
        #左侧顶端输入框
        mdbm_part =  mark[0][0]
        self.sendBtn = wx.Button(pnl,-1,u'发送拣选')
        self.inputText = wx.TextCtrl(pnl,-1,style=wx.TE_PROCESS_ENTER)#输入框 回车
        input_sizer = wx.GridBagSizer()
        input_sizer.Add(self.sendBtn,pos=(0,0),span=(1,1), flag=wx.ALL,border=5)
        input_sizer.Add(self.inputText,pos=(0,1),span=(1,5),flag=wx.EXPAND|wx.ALL,border=5)
        self.inputText.SetFocus()
        self.mdbm = wx.StaticText(pnl,-1,u'')
        self.mdbm.SetLabel(str(mdbm_part))
        self.mdname = wx.StaticText(pnl,-1,u'',size=(200,30))
        self.mdPhone = wx.StaticText(pnl,-1,u'')        
        input_sizer.Add(wx.StaticText(pnl,-1,u'门店编码:'),pos=(1,0),span=(1,1), flag=wx.ALL,border=5)
        input_sizer.Add(self.mdbm,pos=(1,1),span=(1,1), flag=wx.ALL,border=5)
        input_sizer.Add(wx.StaticText(pnl,-1,u'门店名称:'),pos=(1,2),span=(1,1), flag=wx.ALL,border=5)
        input_sizer.Add(self.mdname,pos=(1,3),span=(1,1), flag=wx.ALL,border=5)
        input_sizer.Add(wx.StaticText(pnl,-1,u'联系电话:'),pos=(1,4),span=(1,1), flag=wx.ALL,border=5)
        input_sizer.Add(self.mdPhone,pos=(1,5),span=(1,1), flag=wx.ALL,border=5)


        self.zongjianshu = wx.StaticText(pnl,-1,u'')
        self.address = wx.StaticText(pnl,-1,u'')
        input_sizer.Add(wx.StaticText(pnl,-1,u'总件数:'),pos=(2,0),span=(1,1), flag=wx.EXPAND)
        input_sizer.Add(self.zongjianshu,pos=(2,1),span=(1,1), flag=wx.ALL,border=5)
        input_sizer.Add(wx.StaticText(pnl,-1,u'地址:'),pos=(2,2),span=(1,1), flag=wx.ALL)
        input_sizer.Add(self.address,pos=(2,3),span=(1,1), flag=wx.ALL)

        self.skuText = wx.StaticText(pnl,-1,'')
        self.unitText = wx.StaticText(pnl,-1,'')
        input_sizer.Add(wx.StaticText(pnl,-1,u'商品编码:'),pos=(3,0),span=(1,1), flag=wx.ALL)
        input_sizer.Add(self.skuText,pos=(3,1),span=(1,1), flag=wx.ALL)
        input_sizer.Add(wx.StaticText(pnl,-1,u'商品规格:'),pos=(3,2),span=(1,1), flag=wx.ALL)
        input_sizer.Add(self.unitText,pos=(3,3),span=(1,1), flag=wx.ALL)

        self.eanText = wx.StaticText(pnl,-1,u' ')
        self.zjsText = wx.StaticText(pnl,-1,u' ')
        self.eanText.SetFont(wx.Font(42,wx.SWISS,wx.NORMAL,wx.BOLD))
        self.zjsText.SetFont(wx.Font(42,wx.SWISS,wx.NORMAL,wx.BOLD))
        self.eanText.SetForegroundColour('blue')
        self.zjsText.SetForegroundColour('#ff6a6a')
        input_sizer.Add(self.eanText,pos=(4,0),span=(1,5), flag=wx.EXPAND)
        input_sizer.Add(self.zjsText,pos=(4,5),span=(1,1), flag=wx.ALL)

        self.spname = wx.StaticText(pnl,-1,u'')
        input_sizer.Add(self.spname,pos=(5,0),span=(1,5), flag=wx.ALL,border=5)

        self.danpinjindu = wx.StaticText(pnl,-1,u'单品进度:')
        self.dapinjindutiao = wx.Gauge(pnl,-1,100)
        input_sizer.Add(self.danpinjindu,pos=(6,0),span=(1,1), flag=wx.ALL)
        input_sizer.Add(self.dapinjindutiao,pos=(6,1),span=(1,5), flag=wx.EXPAND|wx.ALL,border=5)
        self.mendianjindu = wx.StaticText(pnl,-1,u'门店进度:')
        self.mendianjindutiao = wx.Gauge(pnl,-1,100)
        input_sizer.Add(self.mendianjindu,pos=(7,0),span=(1,1), flag=wx.ALL)
        input_sizer.Add(self.mendianjindutiao,pos=(7,1),span=(1,5), flag=wx.EXPAND|wx.ALL,border=5)
        
        left_top_sizer = wx.StaticBoxSizer(wx.StaticBox(pnl,-1,u'复核信息'),wx.VERTICAL)
        left_top_sizer.Add(input_sizer,0,wx.ALL,5)

        left_bottom_sizer = wx.StaticBoxSizer(wx.StaticBox(pnl,-1,u'错误信息'))
        
        error_cols = (u"id",u'商品编码',u"条码(键入的)",u'名称',u'件数',u'操作人',u'错误类型')
        self.error_data =[]
        result_error = Connect().select('*','error_fuhe',order="id desc",limit='200')
        for x in result_error:
            self.error_data.append([x[0],x[1],x[2],x[3],x[4],x[5],x[6]])
        self.egdata = GridData(self.error_data,error_cols)
        self.errorgrid = wx.grid.Grid(pnl,size=(800,600))
        self.errorgrid.SetTable(self.egdata)
        self.errorgrid.AutoSize()
        

        left_bottom_sizer.Add(self.errorgrid,0,wx.ALL|wx.EXPAND,5)
        left_bsizer = wx.BoxSizer(wx.VERTICAL)
        left_bsizer.Add(left_top_sizer,0,wx.ALL|wx.EXPAND,5)
        left_bsizer.Add(left_bottom_sizer,0,wx.ALL|wx.EXPAND,5)
        
        self.pid = Connect().get_one('pid','Shipment',where='',order='id Desc')[0]
        right_top_sizer = wx.StaticBoxSizer(wx.StaticBox(pnl,-1,u'复核列表'))
        
        scattered_cols = (u"id",u'商品编码',u"条码/外箱码",u'拣选人',u'件数/数量',u'规格',u'拣选货位',u'已复核数',u'名称')
        self.scattered_data =[]
        result_scattered = Connect().select('*','send_scattered',order="sku",where='pid="%s" and mdbm="%s" and send_type in (1,2,3)'%(self.pid,mdbm_part))
        for x in result_scattered:
            self.scattered_data.append([x[0],x[1],x[2],x[4],str(float(x[6])/float(x[13]))+"/"+str(x[6]),x[13],x[5],x[15],x[3]])
        self.sgdata = GridData(self.scattered_data,scattered_cols)
        self.scatteredgrid = wx.grid.Grid(pnl,size=(800,600))
        self.scatteredgrid.SetTable(self.sgdata)
        self.scatteredgrid.AutoSize()
        right_top_sizer.Add(self.scatteredgrid,0,wx.ALL,5)

        right_bottom_sizer = wx.StaticBoxSizer(wx.StaticBox(pnl,-1,u'系统日志'))
        self.messageText = wx.TextCtrl(pnl,-1,u'',style=wx.TE_MULTILINE|wx.TE_RICH2,\
            size=(display[0]*0.41,display[1]*0.20))
        
        right_bottom_sizer.Add(self.messageText,0,wx.ALL|wx.EXPAND,5)
        right_sizer = wx.BoxSizer(wx.VERTICAL)
        right_sizer.Add(right_top_sizer,0,wx.ALL|wx.EXPAND,5)
        right_sizer.Add(right_bottom_sizer,0,wx.ALL|wx.EXPAND,5)
        body_sizer  = wx.BoxSizer()
        body_sizer.Add(left_bsizer,0,wx.EXPAND,5)
        body_sizer.Add(right_sizer,0,wx.EXPAND,5)
        pnl.SetSizer(body_sizer)

        self.sendBtn.Bind(wx.EVT_BUTTON,self.OnSendBtn)
        mark_data_list = [mdbm_part,self.pid]
        self.inputText.Bind(wx.EVT_TEXT_ENTER,lambda evt,mark=mark_data_list:self.OnInput(evt,mark))
        self.zongjianshuText = 0
        self.zongjianshu_index = 0
        for i in result_scattered:
        	self.zongjianshuText += float(str(i[6]))/float(str(i[13]))
        	self.zongjianshu_index += float(str(i[15]))/float(str(i[13]))
        self.zongjianshu.SetLabel(str(self.zongjianshuText))
        self.mendianjindu.SetLabel(u'门店进度:%s/%s'%(str(int(self.zongjianshu_index)),str(int(self.zongjianshuText))))
        self.mendianjindutiao.SetRange(self.zongjianshuText)
        self.mendianjindutiao.SetValue(self.zongjianshu_index)



    def OnSendBtn(self,evt):
        mdbm = self.mdbm.GetLabel()
        pid = self.pid
        if Connect().update({'static':1,'start_time':str(int(time.time()))},'send_scattered',where='pid="%s" and mdbm="%s" and send_type in (1,2,3) and static=0'%(pid,mdbm)):
            self.messageText.SetValue(str(self.messageText.GetValue())+(time.strftime(u"%Y-%m-%d %H:%M:%S",time.localtime())+u' 发送拣选信息\n'))
            wx.MessageBox(u'发送完成',u'提示',wx.OK)
            self.sendBtn.Enable(False)
        else:
            wx.MessageBox(u'发送失败',u'提示',wx.ICON_ERROR)
        self.inputText.SetFocus()
        

    def OnInput(self,evt,mark):
        xm = self.inputText.GetValue()
        mdbm = mark[0]
        pid = mark[1]
    	self.inputText.SetValue('')
        self.inputText.SetFocus()
        count_threader = zhengjianfuhe_Thread(self,xm,mdbm,pid)
        count_threader.start()

    def LogThreadMessage(self,msg):
        self.eanText.SetLabel(msg)
    def LogThreadMessageText(self,msg):
        self.messageText.AppendText(msg)
    def ThreadClearText(self):
    	self.dapinjindutiao.SetValue(0)
    	self.dapinjindutiao.SetRange(0)
    	self.skuText.SetLabel('')
    	self.unitText.SetLabel('')
    	self.zjsText.SetLabel('')
    	self.spname.SetLabel('')
    	self.danpinjindu.SetLabel(u'单品进度:')
    def ThreadSetgauge(self,args):
    	self.dapinjindutiao.SetValue(args[1])
    	self.dapinjindutiao.SetRange(args[0])
    	self.zongjianshu_index +=1
    	self.mendianjindutiao.SetValue(int(self.zongjianshu_index))

    	self.skuText.SetLabel(str(int(float(args[2]))))
    	self.unitText.SetLabel(str(int(float(args[4]))))
    	self.zjsText.SetLabel(str(int(float(args[1]))))
    	self.spname.SetLabel(args[3])
    	self.danpinjindu.SetLabel(u'单品进度:'+str(int(float(args[1])))+"/"+str(int(float(args[0]))))
    	self.mendianjindu.SetLabel(u'门店进度:%s/%s'%(str(int(self.zongjianshu_index)),str(int(self.zongjianshuText))))
        

                   



class zhengjianfuhe_Thread(threading.Thread):
    def __init__(self,windows,xm,mdbm,pid):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
        self.xm = xm
        self.mdbm = mdbm
        self.pid = pid

    def stop(self):
        self.timeToQuit.set()

    def run(self):
    	wx.CallAfter(self.window.LogThreadMessage,self.xm)
        result = Connect().select('*','send_scattered',where='pid="%s" and mdbm="%s" and ean="%s"'%(self.pid,self.mdbm,self.xm))

        if not result:
            wx.CallAfter(self.window.LogThreadMessageText,str(time.strftime(u"%Y-%m-%d %H:%M:%S",time.localtime()))+u'  键入：'+str(self.xm)+u' 复核失败,没有该条码\n')
            wx.CallAfter(self.window.ThreadClearText)
            winsound.PlaySound("error.wav", winsound.SND_FILENAME)
            return
        else:
            for i in result:
            	if int(float(str(i[6])))-int(float(str(i[15])))==0:continue;
                if int(float(str(i[6])))-(int(float(str(i[15])))+int(float(str(i[13]))))>=0:
                    Connect().update({'out_count':int(float(str(i[15])))+int(float(str(i[13])))},'send_scattered',where='id="%s"'%i[0])
                    wx.CallAfter(self.window.LogThreadMessageText,str(time.strftime(u"%Y-%m-%d %H:%M:%S",time.localtime()))+u'  键入：'+str(self.xm)+u' 复核成功。\n')
                    js = float(str(i[6]))/float(str(i[13]))
                    chujs = float(str(i[15]))/float(str(i[13]))+1.0
                    wx.CallAfter(self.window.ThreadSetgauge,[js,chujs,i[1],i[3],i[13]])
                    winsound.PlaySound("ok.wav", winsound.SND_FILENAME)
                    Connect().insert({'sku':i[0],'count':i[13],'pid':self.pid,'mdbm':self.mdbm,'time':time.time()},'out_fuhe')
                    return
            wx.CallAfter(self.window.LogThreadMessageText,str(time.strftime(u"%Y-%m-%d %H:%M:%S",time.localtime()))+u'  键入：'+str(self.xm)+u' 复核失败,数量超出\n')
            wx.CallAfter(self.window.ThreadClearText)
            winsound.PlaySound("error.wav", winsound.SND_FILENAME)
                    

                # if int(float(str(i[6])))-(int(float(str(i[15])))+int(float(str(i[13]))))<0:

                




















