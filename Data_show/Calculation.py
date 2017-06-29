#coding=utf-8
__author__ = 'anngle'
"""门店发货位计算,总数表，分拨表 （复核表计算？）散货拣选信息,拣选人员选定 """
import wx,xlwt,time,math,threading,xlrd
# from xpinyin import Pinyin
from SQLet import *
from itertools import groupby
from operator import itemgetter
from Public.GridData import GridData
class calculation(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'发货位')
        panel = wx.Panel(self)
        panel.SetBackgroundColour('White')
        self.countBtn = wx.Button(panel,-1,u'计算发货位',pos=(20,20))
        self.RepeatBtn = wx.Button(panel,-1,u'再次计算',pos=(120,5))
        self.toExcelBtn = wx.Button(panel,-1,u'导出所有门店件数',pos=(220,5))
#         _cols = (u"门店编码",u'总件数',u"发货位")
#         self.data =[]
#         self.data = GridData(self.data,_cols)
#         self.grid = wx.grid.Grid(panel)
#         self.grid.SetTable(self.data)
#         self.grid.AutoSize()
        static_text = wx.StaticText(panel,-1,u'计算发货位时间会比较长，期间系统会按照门店要货量排列生成发货位，\
计算出整件数量以及散货拣选数量,并扣除库存，请确保单据正确，基础数据正确，及导出保存库存。',pos=(10,150),size=(500,100))
        static_text.SetBackgroundColour('#EEE9BF')
        static_text.SetFont(wx.Font(18,wx.SWISS,wx.NORMAL,wx.BOLD))
        BodySizer = wx.BoxSizer(wx.VERTICAL) #垂直
        BodySizer.Add(self.countBtn,0,wx.ALL,5)
        self.gauge = wx.Gauge(panel,-1,100,size=(500,25),style = wx.GA_PROGRESSBAR)
        self.logmessage = wx.StaticText(panel,-1,u'')

        BodySizer.Add(self.gauge,0,wx.ALL,5)
        BodySizer.Add(self.logmessage,0,wx.ALL,5)
        BodySizer.Add(static_text,0,wx.ALL,5)
        
        panel.SetSizer(BodySizer)
        BodySizer.Fit(self)
        BodySizer.SetSizeHints(self)

        self.countBtn.Bind(wx.EVT_BUTTON, self.OnCountBtn)
        self.RepeatBtn.Bind(wx.EVT_BUTTON, self.OnRepeatBtn)
        self.toExcelBtn.Bind(wx.EVT_BUTTON, self.OntoExcelBtn)
        
        #写入数据库前判定
#         data_pid = Connect().get_one('pid','Shipment',where='',order='id Desc')[0]
#         """  测试用  按钮 取消 禁用 """
#         if not Connect().get_one('id','distribution',where="pid='%s'"%data_pid):
#             self.countBtn.Enable(False)
#             self.countBtn.SetLabel(u'已经计算过了')


    def LogMessage(self,msg):
        self.logmessage.SetLabel(msg)
    def LogGauge(self,count):
        self.gauge.SetValue(count)

    #计算发货位、总数拣选表 分拨表
    def OnCountBtn(self,evt):
        count_threader = Calculation_Thread(self)
        count_threader.start()
        self.countBtn.Enable(False)
        self.RepeatBtn.Enable(False)
        self.toExcelBtn.Enable(False)

    def OnRepeatBtn(self,evt):
        try:
            #写入excel
            data_pid = Connect().get_one('pid','Shipment',where='',order='id Desc')[0]
            file = xlwt.Workbook(encoding='utf8')
            result_send_position_excel = Connect().select('*', 'send_position',where='pid="%s"'%data_pid)
            table = file.add_sheet(u'发货位')
            table.write(0,0,u'门店编码')
            table.write(0,1,u'门店地址')
            table.write(0,2,u'数量')
            table.write(0,3,u'件数')
            table.write(0,4,u'发货位')
            for index,x in enumerate(result_send_position_excel):
                table.write(index+1,0,x[1])
                table.write(index+1,1,x[4])
                table.write(index+1,2,x[2])
                table.write(index+1,3,x[3])
                table.write(index+1,4,x[5])
            result_total_excel = Connect().select('*', 'total',where='pid="%s"'%data_pid)
            table_zongshu = file.add_sheet(u'总数打印表')
            table_zongshu.write(0,0,u'拣选')
            table_zongshu.write(0,1,u'商品编码')
            table_zongshu.write(0,2,u'商品名称')
            table_zongshu.write(0,3,u'配送数量')
            table_zongshu.write(0,4,u'件数')
            table_zongshu.write(0,5,u'规格')
            table_zongshu.write(0,6,u'位置')
            for index,x in enumerate(result_total_excel):
                table_zongshu.write(index+1,0,x[1])
                table_zongshu.write(index+1,1,x[2])
                table_zongshu.write(index+1,2,x[3])
                table_zongshu.write(index+1,3,x[4])
                table_zongshu.write(index+1,4,round(float(str(x[4]))/float(str(x[5])),2))
                table_zongshu.write(index+1,5,x[5])
                table_zongshu.write(index+1,6,x[6])
            result_total_excel = Connect().select('*', 'distribution',where='pid="%s"'%data_pid)
            table_fenbo = file.add_sheet(u'拣选分拨单')
            table_fenbo.write(0,0,u'拣选')
            table_fenbo.write(0,1,u'单据号')
            table_fenbo.write(0,2,u'门店编码')
            table_fenbo.write(0,3,u'门店名称')
            table_fenbo.write(0,4,u'商品编码')
            table_fenbo.write(0,5,u'商品名称')
            table_fenbo.write(0,6,u'配送数量')
            table_fenbo.write(0,7,u'件数')
            table_fenbo.write(0,8,u'规格')
            table_fenbo.write(0,9,u'发货码头')
            for index,x in enumerate(result_total_excel):
                table_fenbo.write(index+1,0,x[1])
                table_fenbo.write(index+1,1,x[10])
                table_fenbo.write(index+1,2,x[2])
                table_fenbo.write(index+1,3,x[3])
                table_fenbo.write(index+1,4,x[4])
                table_fenbo.write(index+1,5,x[5])
                table_fenbo.write(index+1,6,x[6])
                table_fenbo.write(index+1,7,round(float(str(x[6]))/float(str(x[7])),2))
                table_fenbo.write(index+1,8,x[7])
                table_fenbo.write(index+1,9,x[8])
            result_fuhe_excel = Connect().select('*', 'distribution',where='pid="%s"'%data_pid,order="fahuowei,jianxuanhao")
            table_fuhe = file.add_sheet(u'复核单据')
            table_fuhe.write(0,0,u'拣选')
            table_fuhe.write(0,1,u'单据号')
            table_fuhe.write(0,2,u'门店编码')
            table_fuhe.write(0,3,u'门店名称')
            table_fuhe.write(0,4,u'商品编码')
            table_fuhe.write(0,5,u'商品名称')
            table_fuhe.write(0,6,u'配送数量')
            table_fuhe.write(0,7,u'件数')
            table_fuhe.write(0,8,u'规格')
            table_fuhe.write(0,9,u'发货码头')
            for index,x in enumerate(result_fuhe_excel,1):
                table_fuhe.write(index,0,x[1])
                table_fuhe.write(index,1,x[10])
                table_fuhe.write(index,2,x[2])
                table_fuhe.write(index,3,x[3])
                table_fuhe.write(index,4,x[4])
                table_fuhe.write(index,5,x[5])
                table_fuhe.write(index,6,x[6])
                table_fuhe.write(index,7,round(float(str(x[6]))/float(str(x[7])),2))
                table_fuhe.write(index,8,x[7])
                table_fuhe.write(index,9,x[8])
            dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
            if dialog.ShowModal() == wx.ID_OK:
                file.save(dialog.GetPath()+u"/系统导出的表格"+str(int(time.time()))+".xls")
                dialog.Destroy()
                wx.MessageBox(u'操作完成',u'提示',wx.ICON_HAND);
        except Exception, e:
            wx.MessageBox(u'%s'%e,u'警告',wx.ICON_ERROR);return


    def OntoExcelBtn(self,evt):
        result_shipment = Connect().get_one('*','Shipment',order='id desc')
        result_send_postition = Connect().select('*','send_position',where='pid="%s"'%result_shipment[10])
        data = []
        result_pack = []
        result_distribution = []
        for x in result_send_postition:
            result_pack = Connect().select('id','scattered_pack',where='pid="%s" and mdbm="%s"'%(result_shipment[10],x[1]),group="time")
            result_distribution = Connect().select('sum(count/unit)','distribution',where='pid="%s" and mdbm="%s"'%(result_shipment[10],x[1]))
            if result_distribution:
                data.append([x[1],len(result_pack),result_distribution[0][0],x[5]])

            else:
                data.append([x[1],len(result_pack),0,x[5]])

        file = xlwt.Workbook(encoding='utf8')
        table = file.add_sheet(u'总件数')

        table.write(0,0,u'发货位')
        table.write(0,1,u'门店编码')
        table.write(0,2,u'散货打包件数')
        table.write(0,3,u'整件分拨件数')
        for i,x in enumerate(data):
            table.write(i+1,0,x[3])
            table.write(i+1,1,x[0])
            table.write(i+1,2,x[1])
            table.write(i+1,3,x[2])
        select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
        if select_dialog.ShowModal() == wx.ID_OK:
            file.save(select_dialog.GetPath()+u"/所有门店总件数"+str(int(time.time()))+".xls")
            select_dialog.Destroy()
            wx.MessageBox(u'操作完成',u'提示',wx.ICON_HAND)


class Calculation_Thread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()

    def stop(self):
        self.timeToQuit.set()

    def run(self):
        try:
#             wx.MessageBox(u'数据计算开始，请耐心等候',u'提示',wx.OK);
            if Connect().get_one('*','replenishment_task'):
                wx.MessageBox(u'补货任务未完成，请完成补货任务后再继续',u'提示',wx.ICON_ERROR);return
            data_pid = Connect().get_one('pid','Shipment',where='',order='id Desc')[0]
            #获取出库
            wx.CallAfter(self.window.LogMessage,u'(1/6)正在获取出库数据及规格。')
            wx.CallAfter(self.window.LogGauge,10)
                
            res_shipment = Connect()._sql_query("*", 'select * from (select mdbm,mdname,pid,SUM(count) as count,sku from shipment where pid="%s" GROUP BY mdbm,sku ) as s inner join (select * from base_sku GROUP BY sku) as b on s.sku = b.sku  ORDER BY s.mdbm,s.sku'%data_pid)
            file = xlwt.Workbook(encoding='utf8')
            #得到出库表及规格之后按照字段区分整件和酒水
            zhengjian = []
            shipin = []
            try:
                for x in res_shipment:
                    if int(str(x[13]))==1:
                        zhengjian.append(x)
                    else:
                        shipin.append(x)
            except Exception, e:
                wx.MessageBox(u'判断商品类型有误，请检查基础数据"是否整件"列是否有空，只允许0、1。%s'%e,u'',wx.ICON_ERROR);return
            wx.CallAfter(self.window.LogMessage,u'(2/6)正在计算整件酒水油品等出库数据。')
            wx.CallAfter(self.window.LogGauge,25)

            #要读取库存扣数  并添加到整件出库表
            task_list = [] #任务列表
            for j in zhengjian:
                result_goods = Connect().select('*','now_goods',where='sku="%s"'%j[4],order='allocation desc')
                
                x_count = int(float(str(j[3]))) #单据数
                for i in result_goods:
                    result_huoweibangdings = Connect()._sql_query('*','select h.title,b.district,b.persons from huowei as h inner join bangding_user as b on h.quyu=b.district where h.title="%s"'%i[6])
                    result_huoweibangding = []
                    if  result_huoweibangdings:
                        result_huoweibangding = result_huoweibangdings[0]
                    else:
                        result_huoweibangding.append(i[6])
                        result_huoweibangding.append('')
                        result_huoweibangding.append(u'无人拣选')
                    if x_count-i[5]>0:
                    	task_list.append([j[0],j[1],j[2],int(str(j[3])),j[4],j[6],j[8],j[9],result_huoweibangding[2],i[6],i[5],j[10]])
                        Connect().delete('now_goods',where="id=%s"%i[0])
                        x_count = x_count-i[5]
                    elif x_count-i[5]==0:
                    	task_list.append([j[0],j[1],j[2],int(str(j[3])),j[4],j[6],j[8],j[9],result_huoweibangding[2],i[6],i[5],j[10]])
                        Connect().delete('now_goods',where="id=%s"%i[0])
                        break
                    else:
                    	task_list.append([j[0],j[1],j[2],int(str(j[3])),j[4],j[6],j[8],j[9],result_huoweibangding[2],i[6],x_count,j[10]])
                        Connect().update({'count':str(int(i[5])-int(x_count))},'now_goods',where="id=%s"%i[0])
                        break
                    result_huoweibangding=[]
            
            wx.CallAfter(self.window.LogMessage,u'(3/6)正在计算食品整件数。')
            wx.CallAfter(self.window.LogGauge,45)
            shipin_zheng = []
            shipin_san = []
            for i in shipin:
            	if float(int(str(i[3])))//float(int(str(i[10])))*float(int(str(i[10])))>0:
            		shipin_zheng.append([i,round(float(int(str(i[3])))//float(int(str(i[10])))*float(int(str(i[10]))),2)])
            	if float(int(str(i[3])))%float(int(str(i[10])))>0:
                	shipin_san.append([i,float(int(str(i[3])))%float(int(str(i[10])))])
            
            shipin_zheng_task = []
            for j in shipin_zheng:
            	result_goods = Connect().select('*','now_goods',where='sku="%s"'%j[0][4],order='allocation desc')
                
                x_count = int(float(str(j[1]))) #单据数
                for i in result_goods:
                    if x_count-i[5]>0:
                    	shipin_zheng_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],i[6],i[5],j[0][10]])
                        Connect ().delete('now_goods',where="id=%s"%i[0])
                        x_count = x_count-i[5]
                    elif x_count-i[5]==0:
                    	shipin_zheng_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],i[6],i[5],j[0][10]])
                        Connect().delete('now_goods',where="id=%s"%i[0])
                        break
                    else:
                    	shipin_zheng_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],i[6],x_count,j[0][10]])
                        Connect().update({'count':str(int(i[5])-int(x_count))},'now_goods',where="id=%s"%i[0])
                        break


            #合并商品总数汇总数量e商品编码i货位j数量
            #合并整件食品 为一行记录a门店i货位j数量
            Merge_SZheng = {} #计算门店商品单品总数
            Merge = {} #计算商品总数
            shipin_zheng_task.sort(key=itemgetter(8))
            for a,b,c,d,e,f,g,h,i,j,k in shipin_zheng_task:
            	merge_sku_title = str(e)+"-"+str(i)
            	if Merge.get(merge_sku_title):
            		Merge[merge_sku_title] = [a,b,c,d,e,f,g,h,i,int(Merge.get(merge_sku_title)[9])+int(str(j)),k]
            	else:
            		Merge[merge_sku_title] = [a,b,c,d,e,f,g,h,i,j,k]
                merge_title = str(a)+"-"+str(e)
                if Merge_SZheng.get(str(merge_title)):
                    Merge_SZheng[str(merge_title)] = \
                    [a,b,c,d,e,f,g,h,\
                    Merge_SZheng.get(str(merge_title))[8]+',('+str(j)+')'+i,\
                    int(str(Merge_SZheng.get(str(merge_title))[9]))+int(str(j)),k]
                else:
                    Merge_SZheng[str(merge_title)] = [a,b,c,d,e,f,g,h,i,j,k]
            Merge_SZheng_list = [(Merge_SZheng[k], k) for k in Merge_SZheng]
            Merge_huizong_list = [(Merge[k], k) for k in Merge]
            wx.CallAfter(self.window.LogMessage,u'(4/6)正在计算食品散货数')
            wx.CallAfter(self.window.LogGauge,70)
            shipin_san_task = [] #任务列表
            for j in shipin_san:
                result_goods = Connect().select('*','now_goods',where='sku="%s"'%j[0][4],order='allocation desc')
                x_count = int(float(str(j[1]))) #单据数
                for i in result_goods:
                    result_huoweibangding = Connect()._sql_query('*','select h.title,b.district,b.persons from huowei as h inner join bangding_user as b on h.quyu=b.district where h.title="%s"'%i[6])[0]
                    
                    if  not result_huoweibangding:
                    	result_huoweibangding.append([i[6],'',u'无人拣选'])

                    if x_count-i[5]>0:
                    	shipin_san_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],result_huoweibangding[2],i[6],i[5],j[0][10]])
                        Connect().delete('now_goods',where="id=%s"%i[0])
                        x_count = x_count-i[5]
                    elif x_count-i[5]==0:
                    	shipin_san_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],result_huoweibangding[2],i[6],i[5],j[0][10]])
                        Connect().delete('now_goods',where="id=%s"%i[0])
                        break
                    else:
                    	shipin_san_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],result_huoweibangding[2],i[6],x_count,j[0][10]])
                        Connect().update({'count':str(int(i[5])-int(x_count))},'now_goods',where="id=%s"%i[0])
                        break
                    result_huoweibangding=[]

            wx.CallAfter(self.window.LogMessage,u'(5/6)正在保存数据')
            wx.CallAfter(self.window.LogGauge,90)

            for i in task_list:
            	insert_data = {'sku':int(float(int(i[4]))),'ean':i[7],'name':i[5],'send_type':1,'persons':i[8],'allocation':i[9],'count':int(float(int(i[10]))),\
                    'static':0,'pid':i[2],'mdbm':int(float(int(i[0]))),'unit':int(float(int(i[11])))}
            	Connect().insert(insert_data,'send_scattered')
            
            #0门店编码1地址2pid3单品数4编码5名称6条码7箱码8货位9数量10规格
            for k,v in Merge_SZheng_list:
            	insert_data = {'sku':int(float(int(k[4]))),'ean':k[7],'name':k[5],\
            	'send_type':2,'allocation':k[8],'count':int(float(int(k[9]))),\
                    'static':0,'pid':k[2],'mdbm':int(float(int(k[0]))),'unit':int(float(int(k[10])))}
            	
            	Connect().insert(insert_data,'send_scattered')
            
            for i in shipin_san_task:
            	insert_data = {'sku':int(float(int(i[4]))),'ean':i[6],'name':i[5],\
            	'send_type':0,'persons':i[8],'allocation':i[9],'count':int(float(int(i[10]))),\
                    'static':0,'pid':i[2],'mdbm':int(float(int(i[0]))),'unit':int(float(int(i[11])))}
            	Connect().insert(insert_data,'send_scattered')
            
            wx.CallAfter(self.window.LogMessage,u'(7/7)生成表格中')
            wx.CallAfter(self.window.LogGauge,100)
            
            table_shipment = file.add_sheet(u'出库表数据')
            title_shipment = [u'门店编码',u'地址',u'pid',u'单品总数量',u'商品编码',u'id',u'商品名称',u'商品编码',u'商品条码',u'外箱码',u'规格',u'忽略',u'类型',u'是否整件']
            for i,x in enumerate(title_shipment):
                table_shipment.write(0,i,x)
            for i,x in enumerate(res_shipment):
                for i_index,x_v in enumerate(x):
                    table_shipment.write(i+1,i_index,x_v)
            
            write_send_scattered = Connect().select('*','send_scattered','pid="%s"'%data_pid)
            table_write_scattered = file.add_sheet(u'所有出库数据')
            title_write_scattered = ['id',u'商品编码',u'商品条码(外箱码)',u'商品名称',u'拣选人员',u'拣选货位',u'拣选数量',u'状态',u'流水号',u'门店编码',u'忽略',u'忽略',u'忽略',u'规格',u'0散货1整件酒水2整件食品']
            for i,x in enumerate(title_write_scattered):
                table_write_scattered.write(0,i,x)
            for i,x in enumerate(write_send_scattered):
                for i_index,x_v in enumerate(x):
                    table_write_scattered.write(i+1,i_index,x_v)
            
            table_task = file.add_sheet(u'整件酒水数据')
            title_task_list = [u'门店编码',u'地址',u'pid',u'单品总数量',u'商品编码',u'商品名称',u'商品条码',u'外箱码',u'拣选人',u'拣选货位',u'拣选数量',u'规格']
            for i,x in enumerate(title_task_list):
                table_task.write(0,i,x)
            for i,x in enumerate(task_list):
                for i_index,x_v in enumerate(x):
                    table_task.write(i+1,i_index,x_v)

            table_shipin_task = file.add_sheet(u'食品整件数据')
            title_shipin_task = [u'门店编码',u'地址',u'pid',u'单品总数量',u'商品编码',u'商品名称',u'商品条码',u'外箱码',u'原货位',u'重定位数量',u'规格']
            for i,x in enumerate(title_shipin_task):
                table_shipin_task.write(0,i,x)
            Merge_SZheng_list_index = 1
            for i,x in Merge_SZheng_list:
                for i_index,x_v in enumerate(i):
                    table_shipin_task.write(Merge_SZheng_list_index,i_index,x_v)
                Merge_SZheng_list_index = Merge_SZheng_list_index+1
            
            table_shipin_san_task = file.add_sheet(u'食品散货数据')
            title_shipin_san_task = [u'门店编码',u'地址',u'pid',u'单品总数量',u'商品编码',u'商品名称',u'商品条码',u'外箱码',u'拣选人',u'拣选货位',u'拣选数量',u'规格']
            for i,x in enumerate(title_shipin_san_task):
                table_shipin_san_task.write(0,i,x)
            for i,x in enumerate(shipin_san_task):
                for i_index,x_v in enumerate(x):
                    table_shipin_san_task.write(i+1,i_index,x_v)
            
            title_shipin_zheng = [u'商品编码',u'外箱码',u'商品名称',u'汇总数量',u'规格',u'原货位',u'件数',u'重定货位']
            #合并数组链接货位，相加拣选数量 e编码i货位j数量
            Merge = {}
            for k,v in Merge_huizong_list:
                if Merge.get(str(k[4])):
                    Merge[str(k[4])] = [k[0],k[1],k[2],k[3],k[4],k[5],k[6],k[7],\
                    Merge.get(str(k[4]))[8]+',('+str(k[9])+')'+k[8],\
                    int(float(str(Merge.get(str(k[4]))[9])))+int(float(str(k[9]))),k[10]]
                else:
                    Merge[str(k[4])] = [k[0],k[1],k[2],k[3],k[4],k[5],k[6],k[7],k[8],k[9],k[10]]
            aa = [(Merge[k], k) for k in Merge]
            
            table_shipin_zheng = file.add_sheet(u'整件食品汇总定位数据')
            title_shipin_zheng_de = [u'商品编码',u'外箱码',u'名称',u'数量',u'规格',u'原货位(货位数量)',u'件数',u'重定货位']
            for i,x in enumerate(title_shipin_zheng_de):
                table_shipin_zheng.write(0,i,x)
            index = 1
            for k,v in aa:
            	table_shipin_zheng.write(index,0,k[4])
            	table_shipin_zheng.write(index,1,k[7])
            	table_shipin_zheng.write(index,2,k[5])
            	table_shipin_zheng.write(index,3,k[9])
            	table_shipin_zheng.write(index,4,k[10])
            	table_shipin_zheng.write(index,5,k[8])
            	table_shipin_zheng.write(index,6,round(float(str(k[9]))/float(str(k[10])),2))
            	index = index +1
            wx.CallAfter(self.window.LogMessage,u'操作结束')
            wx.CallAfter(self.window.LogGauge,0)
            select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
            if select_dialog.ShowModal() == wx.ID_OK:
                file.save(select_dialog.GetPath()+u"/出库数据"+str(int(time.time()))+".xls")
                select_dialog.Destroy()
                wx.MessageBox(u'操作完成',u'提示',wx.OK)
            

        except Exception, e:
            wx.MessageBox(u'程序运行出错：%s'%e,u'警告',wx.ICON_ERROR);return






#散货拣选信息
class sanhuo_jianxuan(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'散货拣选',pos=(20,20))
        pnl = wx.Panel(self)
        pnl.SetBackgroundColour('White')
        self.fahuowei_text = wx.TextCtrl(pnl,-1,u'')
        persons_btn = wx.Button(pnl,-1,u'选择捡货人')
        send_btn = wx.Button(pnl,-1,u'允许拣选')
        back_btn = wx.Button(pnl,-1,u'撤回拣选')
        toExcel_btn = wx.Button(pnl,-1,u'导出表格')
        r_list = Connect().select('username','users',where='type=0')
        choice_list = []
        for i in  r_list:
            choice_list.append(i[0])
        self.choice_type = wx.Choice(pnl,-1,size=(150,25),choices=choice_list)
        self.huowei = wx.TextCtrl(pnl,-1,u'')
        self.update_btn = wx.Button(pnl,-1,u'更新')
        top_sizer = wx.BoxSizer()
        top_sizer.Add(wx.StaticText(pnl,-1,u'发货位:'),0,wx.ALL,5)
        top_sizer.Add(self.fahuowei_text,0,wx.ALL,5)
        top_sizer.Add(persons_btn,0,wx.ALL,5)
        top_sizer.Add(send_btn,0,wx.ALL,5)
        top_sizer.Add(back_btn,0,wx.ALL,5)
        top_sizer.Add(toExcel_btn,0,wx.ALL,5)
        top_sizer.Add(self.choice_type,0,wx.ALL,5)
        top_sizer.Add(wx.StaticText(pnl,-1,u'货位号'),0,wx.ALL,5)
        top_sizer.Add(self.huowei,0,wx.ALL,5)
        top_sizer.Add(self.update_btn,0,wx.ALL,5)
        result_shipmnet = Connect().get_one('pid','Shipment',order="id desc")[0]
        scattered_data = Connect().select('fahuowei,mdbm,persons,sku,ean,name,allocation,count,static','send_scattered',where="pid='%s' and send_type = 0"%result_shipmnet,order='static asc')

        _cols = (u"发货位",u'门店编码',u'拣选人员',u"商品编码",u'商品条码',u'商品名称',u'货位号',u'数量',u'状态')
        self.data =[]
        self.data = GridData(scattered_data,_cols)
        self.grid = wx.grid.Grid(pnl)
        self.grid.SetTable(self.data)
        self.grid.AutoSize()

        body_sizer = wx.BoxSizer(wx.VERTICAL)
        body_sizer.Add(top_sizer)
        body_sizer.Add(self.grid,0,wx.ALL|wx.EXPAND,5)
        pnl.SetSizer(body_sizer)
        body_sizer.Fit(self)
        body_sizer.SetSizeHints(self)
        persons_btn.Bind(wx.EVT_BUTTON,self.OnSelect_person)
        # send_btn.Bind(wx.EVT_BUTTON,self.OnSend_btn)
        back_btn.Bind(wx.EVT_BUTTON,self.Onback_btn)
        toExcel_btn.Bind(wx.EVT_BUTTON,self.OnToExcel_btn)
        self.update_btn.Bind(wx.EVT_BUTTON,self.OnUpdatebtn)

    #获得散货拣货信息
    def OnSend_btn(self,evt):
        fahuowei = self.fahuowei_text.GetValue()
        self.fahuowei_text.SetValue("")
        if not fahuowei.isdigit():
            wx.MessageBox(u'输入的必须是数字',u'提示',wx.ICON_ERROR);return
        result_shipmnet = Connect().get_one('pid','Shipment',order="id desc")[0]
        result_mendian = Connect().get_one('*','send_position',where="pid='%s' and fahuowei=%s"%(result_shipmnet,fahuowei))
        if not result_mendian:wx.MessageBox(u'没有该发货位',u'提示',wx.ICON_ERROR);return
        result_shipmnet_jianxuan = Connect().select('*','Shipment',where='pid="%s" and mdbm=%s'%(result_shipmnet,result_mendian[1]))
        if Connect().get_one('id','send_scattered',where='mdbm="%s" and pid="%s" and send_type = 0'%(result_mendian[1],result_shipmnet)):return
        list_jianxuan_shipment= []
        for i in result_shipmnet_jianxuan:
            result_base_sku_unit = Connect().get_one('unit','base_sku',where='sku=%s'%i[2])[0]
            if float(i[5])%float(result_base_sku_unit)>0:
                #0名称1编码2条码3散货数量4订单号5门店编码6pid
                list_jianxuan_shipment.append([i[1],i[2],i[3],float(i[5])%float(result_base_sku_unit),\
                    i[6],i[7],i[10]])
        if len(list_jianxuan_shipment)==0:
            wx.MessageBox(u'该发货位没有散货拣选，请继续下一个发货位',u'提示',wx.ICON_ERROR);return
        list_goods_jianhuo = []
        result_users = Connect().select('*','users',where='static="1"')
        if len(result_users)==0:wx.MessageBox(u'没有人拣货，请返回选择拣货员',u'提示',wx.ICON_ERROR);return
        for x in list_jianxuan_shipment:
            result_goods = Connect().select('*','now_goods',where='sku=%s'%x[1])
            pre_count = int(x[3])
            #0名称1编码2条码3拣选数量4订单号5门店编码6pid7货位
            for i in result_goods:
                if pre_count-int(i[5])>0:
                    pre_count = pre_count-int(i[5])
                    list_goods_jianhuo.append([x[0],x[1],x[2],int(i[5]),x[4],x[5],x[6],i[6]])
                    Connect().delete('now_goods',where="id=%s"%i[0])
                elif pre_count-int(i[5])==0:
                    list_goods_jianhuo.append([x[0],x[1],x[2],int(i[5]),x[4],x[5],x[6],i[6]])
                    Connect().delete('now_goods',where="id=%s"%i[0])
                    break
                elif pre_count-int(i[5])<0:
                    list_goods_jianhuo.append([x[0],x[1],x[2],pre_count,x[4],x[5],x[6],i[6]])
                    Connect().update({'count':int(int(i[5])-pre_count)},'now_goods',where='id=%s'%i[0])
                    break
        list_goods_jianhuo.sort(key=itemgetter(7))

        person_jianxuan_list = []
        for u in range(len(result_users),0,-1):
            for l in range(len(list_goods_jianhuo),len(list_goods_jianhuo)-\
                    int(math.ceil(len(list_goods_jianhuo)/float(u))),-1):
                list_goods_jianhuo[l-1].append(result_users[u-1][3])
                person_jianxuan_list.append(list_goods_jianhuo[l-1])
                list_goods_jianhuo.pop()
        for x in person_jianxuan_list:
            Connect().insert({'sku':int(x[1]),'ean':x[2],'name':x[0],'persons':x[8],'allocation':x[7],'count':int(x[3]),\
                    'static':1,'pid':x[6],'mdbm':int(x[5]),'fahuowei':fahuowei,'start_time':int(time.time())},'send_scattered')
        wx.MessageBox(u'发送完成',u'提示',wx.OK)

    def OnSelect_person(self,evt):select_jianxuanren().Show()

    def Onback_btn(self,evt):
        fahuowei = self.fahuowei_text.GetValue()
        if fahuowei=="":
            wx.MessageBox(u'请输入发货位',u'警告',wx.ICON_ERROR);return

    def OnToExcel_btn(self,evt):
        result_shipmnet = Connect().get_one('pid','Shipment',order="id desc")[0]
        result_send_scattered = Connect().select('*','send_scattered','pid="%s"'%result_shipmnet,order="fahuowei")
        #写入excel
        file = xlwt.Workbook(encoding='utf8')
        table = file.add_sheet(u'散货拣选表')
        title_list = [u'发货位',u'商品编码',u'商品条码',u'商品名称',u'拣选人',u'拣选数量',u'拣选货位',u'门店编码',u'允许拣选时间',u'完成拣选时间']
        for i,x in enumerate(title_list):
            table.write(0,i,x)
        for index,x in enumerate(result_send_scattered,1):
            table.write(index,0,x[10])
            table.write(index,1,x[1])
            table.write(index,2,x[2])
            table.write(index,3,x[3])
            table.write(index,4,x[4])
            table.write(index,5,x[6])
            table.write(index,6,x[5])
            table.write(index,7,x[9])
            start_time = time.localtime(x[11])
            end_time = time.localtime(x[12])
            table.write(index,8,time.strftime('%Y-%m-%d %H:%M:%S',start_time))
            table.write(index,9,time.strftime('%Y-%m-%d %H:%M:%S',end_time))
        select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
        name_time = time.localtime(time.time())
        if select_dialog.ShowModal() == wx.ID_OK:
            file.save(select_dialog.GetPath()+u"/散货拣选表"+time.strftime('%Y%m%d%H%M%S',name_time)+".xls")
            select_dialog.Destroy()
            wx.MessageBox(u'操作完成',u'提示',wx.OK)

    def OnUpdatebtn(self,evt):
        huowei = self.huowei.GetValue().upper()
        persons = self.choice_type.GetString(self.choice_type.GetSelection())
        
        if persons =='':
            wx.MessageBox(u'请选择拣选人',u'警告',wx.ICON_ERROR);return
        if huowei =='':
            wx.MessageBox(u'请输入货位',u'警告',wx.ICON_ERROR);return
        if Connect().update({'persons':persons},'send_scattered',where='allocation="%s" and static=0 '%huowei):
            wx.MessageBox(u'更新完成。',u'警告',wx.OK);
            self.huowei.SetValue('')
            self.huowei.SetFocus()
        else:
            wx.MessageBox(u'更新失败。',u'警告',wx.ICON_ERROR);





#拣选人员选定
class select_jianxuanren(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,u'选取拣货人员',size=(500,500))
        pnl = wx.Panel(self)
        wx.StaticText(pnl,-1,u'请勾选拣货人员：',pos=(20,20))
        users = Connect().select('*','users',where="type=0")
        pos_x = 20
        pos_y = 50
        for x in users:
            if x[5]==1:
                wx.CheckBox(pnl,-1,x[3],pos=(pos_x,pos_y)).SetValue(True)
            elif x[5]==0:
                wx.CheckBox(pnl,-1,x[3],pos=(pos_x,pos_y))
            pos_x += 120
            if pos_x>400:
                pos_y += 50
                pos_x = 20

        self.Bind(wx.EVT_CHECKBOX,self.ONCheck)

    def ONCheck(self,evt):
        if evt.GetEventObject().GetValue():
            Connect().update({'static':1},'users',where='username="%s"'%evt.GetEventObject().GetLabel())
        else:
            Connect().update({'static':0},'users',where='username="%s"'%evt.GetEventObject().GetLabel())




class fenpei(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'整件食品数据分配',pos=(0,0),size=(1000,900))
        pnl  =  wx.Panel(self)
        top_sizer = wx.BoxSizer()
        self.upload_btn = wx.Button(pnl,-1,u'更新')
        self.file_select=wx.FilePickerCtrl(pnl,wx.ID_ANY,u"",u"选择文件",u"*.*",\
               wx.DefaultPosition,wx.DefaultSize,wx.FLP_DEFAULT_STYLE,\
               validator=wx.DefaultValidator)
        self.file_select.GetPickerCtrl().SetLabel(u'选择文件')
        top_sizer.Add(self.file_select,0,wx.ALL,5)
        top_sizer.Add(self.upload_btn,0,wx.ALL,5)
        self._data = []
        _cols = (u"ID",u'门店编码',u"拣选人",u"拣选货位",u'拣选数量',u"商品编码",u"外箱码",u"商品名称")
        try:
            pid = Connect().get_one('pid','Shipment','',"id DESC")
            res_dic = Connect().select('*','send_scattered',where='pid="%s" and send_type=2'%pid[0],order='allocation')
            for x in res_dic:
            	if x[4] == None:
            		self._data.append([x[0],x[9],'',x[5],x[6],x[1],x[2],x[3]])
            	else:
	            	self._data.append([x[0],x[9],x[4],x[5],x[6],x[1],x[2],x[3]])
        except:
            pid = [];res_dic = []
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(pnl,size=(900,800))
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        body_sizer = wx.BoxSizer(wx.VERTICAL)
        body_sizer.Add(top_sizer)
        body_sizer.Add(self.grid)
        pnl.SetSizer(body_sizer)
        # self.upload_btn.SetLabel(u'999b')
        self.upload_btn.Bind(wx.EVT_BUTTON,self.OnUploadBtn)

    def OnUploadBtn(self,evt):
    	try:
            file_path = self.file_select.GetPath()
            if file_path=='':
                wx.MessageBox(u'请选择文件',u'提示',wx.ICON_ERROR);return
            self.upload_btn.Enable(False)
            self.file_select.SetPath('')
            file_types = file_path.split('.')
            file_type = file_types[len(file_types) - 1];
            if file_type.lower()  != "xlsx" and file_type.lower()!= "xls":
                wx.MessageBox(u'文件格式不正确，请选择正确的表单文件',u'提示',wx.ICON_ERROR);return
            # wx.MessageBox(u'后台正在处理单据数据,请稍等...',u'提示',wx.OK)
            out_threader = fenpei_Thread(self,file_path)
            out_threader.start()
            
            
            
        except Exception, e:
            wx.MessageBox(u'程序发生错误，错误信息是：%s'%e,u'警告',wx.ICON_ERROR)

    def LogGauge(self,count):
    	if count>100:
    		self.upload_btn.SetLabel(u'更新完成')
    	else:
        	self.upload_btn.SetLabel(str(count)+"%")

class fenpei_Thread(threading.Thread):
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
            if table.col(0)[0].value.strip() != u'商品编码':
                message = u"第一行名称必须叫‘商品编码’，请返回修改"
            if table.col(3)[0].value.strip() != u'数量':
                message = u"第七行名称必须叫‘数量’，请返回修改"
            if table.col(7)[0].value.strip() != u'重定货位':
                message = u"第十一行名称必须叫‘重定货位’，请返回修改"
        except Exception,f:
            wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
        table_data_list = []
        for rownum in range(1,table.nrows):
            if table.row_values(rownum):
                table_data_list.append(table.row_values(rownum))
            wx.CallAfter(self.window.LogGauge,round(float(rownum)/(table.nrows)*10,2))
        for i,x in enumerate(table_data_list):
        	wx.CallAfter(self.window.LogGauge,round(float(i)/len(table_data_list)*10+10,2))
        	if not x[7]:
        		wx.MessageBox(u'数据错误，请检查是否重定货位是否输入正确。',u'警告',wx.ICON_ERROR);return
        	if not Connect().get_one('id','huowei','title="%s"'%x[7]):
        		wx.MessageBox(u'数据错误，第%s行货位：%s在系统中不存在。'%(i,x[7]),u'警告',wx.ICON_ERROR);return
        pid = Connect().get_one('pid','Shipment','',"id DESC")[0]
        
        for i,x in enumerate(table_data_list):
        	wx.CallAfter(self.window.LogGauge,round(float(i)/len(table_data_list)*50+20,2))
        	
        	jianxuanren = Connect()._sql_query('*','select b.persons from huowei as h  join bangding_user as b on h.quyu=b.district where h.title="%s"'%x[7])[0]
        	
        	if jianxuanren:
        		Connect().update({'persons':jianxuanren[0],'allocation':x[7]},'send_scattered','pid="%s" and send_type=2 and sku="%s" '%(pid,int(float(str(x[0])))))
        	else:
        		Connect().update({'persons':'无人拣选','allocation':x[7]},'send_scattered','pid="%s" and send_type=2 and sku="%s" '%(pid,int(float(str(x[0])))))
        wx.CallAfter(self.window.LogGauge,101)




























