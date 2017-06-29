#coding=utf-8
import threading,wx,time,xlrd,sys,xlwt,winsound,copy
from SQLet import *
from operator import itemgetter
from collections import defaultdict
reload(sys)
sys.setdefaultencoding("utf-8")

#进货单
class jinhuo_Thread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        threading.Event().clear()
        self.window = windows
    def run(self):
        pre_top = Connect()._sql_query('*','select * from pre_top ORDER BY id desc limit 1')
        if pre_top:
            pre = Connect()._sql_query('*','select * from pre where pid="%s" order by gysname,sku'%pre_top[0][1])
        else: 
            pre=[]
    	if pre:
            wx.CallAfter(self.window.SetGridValue,pre)
            wx.CallAfter(self.window.SetTextStringValue,[u'单据编号：%s'\
    			%pre_top[0][1].upper(),u'单据时间：%s'%str(pre_top[0][3])[0:16]])
        Connect().close()
#进货单导入
class InExcel_Thread(threading.Thread):
    def __init__(self,windows,path):
        threading.Thread.__init__(self)
        threading.Event().clear()
        self.window = windows
        self.path = path
    def run(self):
        filedata = xlrd.open_workbook(self.path,encoding_override='utf-8')
        table = filedata.sheets()[0]
        message =""
        try:
            if table.col(0)[0].value.strip() != u'商品编码':
                message = u"第一行名称必须叫‘商品编码’，请返回修改"
            if table.col(1)[0].value.strip() != u'商品条码':
                message = u"第二行名称必须叫‘商品条码’，请返回修改"
            if table.col(2)[0].value.strip() != u'商品名称':
                message = u"第三行名称必须叫‘商品名称’，请返回修改"
            if table.col(3)[0].value.strip() != u'数量':
                message = u"第四行名称必须叫‘数量’，请返回修改"
            if table.col(4)[0].value.strip() != u'货位号':
                message = u"第五行名称必须叫‘货位号’，请返回修改"
            if table.col(5)[0].value.strip() != u'供应商':
                message = u"第六行名称必须叫‘供应商’，请返回修改"
        except Exception,f:
            wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
        nrows = table.nrows #行数
        table_data_list =[]
        wx.CallAfter(self.window.SetBtnText,u'加载数据中')
        for rownum in range(1,nrows):
            if table.row_values(rownum):
                table_data_list.append(table.row_values(rownum))
            wx.CallAfter(self.window.SetGauge,float(rownum)/(table.nrows)*25)
        wx.CallAfter(self.window.SetBtnText,u'检查数据中')
        #得到库存字典 
        goods_list = Connect().select('*','goods')
        goods_dic = defaultdict(list)
        for goods in goods_list:
            goods_dic[str(goods[2])].append(goods)
        #得到基础数据字典 
        baseSku_list = Connect().select('*','base_sku')
        baseSku_dic = defaultdict(list)
        for b in baseSku_list:
            baseSku_dic[str(b[2])].append(b)
        #得到货位字典 
        huowei_list = Connect().select('*','huowei')
        huowei_dic = defaultdict(list)
        for huowei in huowei_list:
            huowei_dic[str(huowei[1])].append(huowei)
        error_text = ''
        try:
            for i,x in enumerate(table_data_list):
                wx.CallAfter(self.window.SetGauge,i/(float(len(table_data_list)))*25+25)
                if not baseSku_dic.has_key(str(int(str(x[0])))):
                	error_text = error_text+u'商品编码：%s（%s）有误，没有该基础数据。\n'%(str(int(float(str(x[0])))),x[2])
                if not huowei_dic.has_key(str(x[4])):
                	error_text = error_text+u'货位号：%s有误，没有该货位数据。\n'%x[4]
            if error_text:
            	raise Exception,error_text
        except Exception, e:
            wx.MessageBox(u'系统错误，错误信息：\n'%e,u'提示',wx.ICON_ERROR);return
             
        wx.CallAfter(self.window.SetBtnText,u'保存数据中')
        pid = "PRE"+str(int(time.time()))
        install_time=int(time.time())
        insert_data_pre_marge=[]
        insert_data_goods_marge = []
        update_data_goods_marge = []
        table_dic  = {}
        top_count = 0
        for i in table_data_list:
            top_count = top_count+int(float(str(i[3])))
            if table_dic.has_key(str(i[0])+"--"+i[4]):
                table_dic[str(i[0])+"--"+i[4]] = [i[0],i[1],i[2],int(float(str(i[3])))+table_dic[str(i[0])+"--"+i[4]][3],i[4],i[5]]
            else:
                table_dic[str(i[0])+"--"+i[4]] = i
        insert_data = []
        for x,y in table_dic.iteritems():
            insert_data.append(y)
        insert_top_pre= {'pid':pid,'count':top_count,'static':0}
        
        for i,x in enumerate(insert_data):
            wx.CallAfter(self.window.SetGauge,i/(float(len(table_data_list)))*50+50)
            insert_data_list = {'name':x[2],'sku':int(float(str(x[0]))),'ean':x[1],\
                            'count':x[3],'pid':pid,'static':0,'gysname':x[5],'allocation':x[4]}
            insert_data_pre_marge.append(insert_data_list)
            
            if goods_dic.has_key(str(int(str(x[0])))):
                for g in goods_dic[str(int(str(x[0])))]:
                    if x[4] == g[5]:
                        update_data_goods_marge.append([g[0],g[4]+x[3]])
                    else:
                        insert_data_goods_marge.append({'sku':int(float(str(x[0]))),'ean':x[1],\
                    'name':x[2],'count':int(float(str(x[3]))),'allocation':x[4]})
            else:
                insert_data_goods_marge.append({'sku':int(float(str(x[0]))),'ean':x[1],\
                    'name':x[2],'count':int(float(str(x[3]))),'allocation':x[4]})

        if insert_top_pre:
            Connect().insert(insert_top_pre,'pre_top')
        if insert_data_pre_marge:
            Connect().insert_many(insert_data_pre_marge,'pre')
        if insert_data_goods_marge:
            Connect().insert_many(insert_data_goods_marge,'goods')
        update_sql = 'begin; UPDATE goods set count = CASE id'
        jonId = '';joinV = ''
        for i in update_data_goods_marge:
            jonId = jonId+' WHEN '+str(i[0])+' THEN '+str(i[1])
            joinV = str(i[0])+' ,'+joinV
        joinV = joinV[0:-1]
        updateSql = update_sql+jonId+' END WHERE id IN ('+ str(joinV)+") ;COMMIT;"
        try:
            if update_data_goods_marge:
                Connect()._sql_query('*',updateSql)
        except Exception, e:
            wx.MessageBox(u'更新库存错误：%s'%e,u'警告',wx.ICON_ERROR);return
        
        wx.CallAfter(self.window.SetGauge,100)
        wx.CallAfter(self.window.SetBtnText,u'导入完成')
        wx.CallAfter(self.window.SetInFileOk)
        Connect().close()
#出货单
class chuhuo_Thread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        threading.Event().clear()
        self.window = windows
    def run(self):
    	shipment_top = Connect()._sql_query('*','select * from shipment_top ORDER BY id desc limit 1')
        if shipment_top:
            shipment = Connect()._sql_query('*','select * from shipment where pid="%s" order by mdbm,sku'%shipment_top[0][1])
        else: 
            shipment=[]
        if shipment:
            wx.CallAfter(self.window.SetGridValue,shipment)
            wx.CallAfter(self.window.SetTextStringValue,[u'单据编号：%s'\
                %shipment_top[0][1].upper(),u'单据时间：%s'%str(shipment_top[0][3])[0:16]])
        Connect().close()
#出货单导入
class OutExcel_Thread(threading.Thread):
    def __init__(self,windows,path):
        threading.Thread.__init__(self)
        threading.Event().clear()
        self.window = windows
        self.path = path
    def run(self):
        filedata = xlrd.open_workbook(self.path,encoding_override='utf-8')
        table = filedata.sheets()[0]
        message =""
        try:
            if table.col(0)[0].value.strip() != u'门店编码':
                message = u"第一行名称必须叫‘门店编码’，请返回修改"
            if table.col(1)[0].value.strip() != u'商品编码':
                message = u"第二行名称必须叫‘商品编码’，请返回修改"
            if table.col(2)[0].value.strip() != u'商品条码':
                message = u"第三行名称必须叫‘商品条码’，请返回修改"
            if table.col(3)[0].value.strip() != u'商品名称':
                message = u"第四行名称必须叫‘商品名称’，请返回修改"
            if table.col(4)[0].value.strip() != u'数量':
                message = u"第五行名称必须叫‘数量’，请返回修改"
            if table.col(5)[0].value.strip() != u'门店名称':
                message = u"第六行名称必须叫‘门店名称’，请返回修改"
            
        except Exception,f:
            wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
        nrows = table.nrows #行数
        table_data_list =[]
        wx.CallAfter(self.window.SetBtnText,u'加载数据中')
        for rownum in range(1,nrows):
            if table.row_values(rownum):
                table_data_list.append(table.row_values(rownum))
            wx.CallAfter(self.window.SetGauge,float(rownum)/(table.nrows)*15)
        wx.CallAfter(self.window.SetBtnText,u'检查数据中')
        table_dic = {}
        for i in table_data_list:
            if table_dic.has_key(str(int(float(str(i[0]))))+"--"+str(int(float(str(i[1]))))):
                table_dic[str(int(float(str(i[0]))))+"--"+str(int(float(str(i[1]))))] = [i[0],i[1],i[2],i[3],\
                int(float(str(i[4])))+table_dic[str(int(float(str(i[0]))))+"--"+str(int(float(str(i[1]))))][4],i[5]]
            else:
                table_dic[str(int(float(str(i[0]))))+"--"+str(int(float(str(i[1]))))] = i
        insert_data = []
        for x,y in table_dic.iteritems():
            insert_data.append(y)
        #得到基础数据字典 
        baseSku_list = Connect().select('*','base_sku')
        baseSku_dic = defaultdict(list)
        for b in baseSku_list:
            baseSku_dic[str(b[2])].append(b)
        error_text = ''
        try:
            for i,x in enumerate(insert_data):
                wx.CallAfter(self.window.SetGauge,i/(float(len(insert_data)))*35+15)
                if not baseSku_dic.has_key(str(int(float(str(x[1]))))):
                	error_text = error_text+u'商品编码：%s（%s）有误，没有该基础数据。\n'%(str(int(float(str(x[1])))),x[3])
            if error_text:
            	raise Exception,error_text
                
        except Exception, e:
            wx.MessageBox(u'系统错误，错误信息：\n%s'%e,u'提示',wx.ICON_ERROR);return
        wx.CallAfter(self.window.SetBtnText,u'保存数据中')
        pid = "ORDER"+str(int(time.time()))
        insert_data_list_marge = []
        top_count = 0
        for i,x in enumerate(insert_data):
            top_count = top_count+int(float(str(x[4])))
            wx.CallAfter(self.window.SetGauge,i/(float(len(insert_data)))*50+50)
            insert_data_list = {'name':x[3],'sku':int(float(str(x[1]))),'ean':x[2],'count':int(float(str(x[4]))),\
                                    'pid':pid,'mdbm':int(float(str(x[0]))),'mdname':x[5]}
            insert_data_list_marge.append(insert_data_list)

        insert_top_shipment= {'pid':pid,'count':top_count,'static':0}
        Connect().insert(insert_top_shipment,'shipment_top')
        if insert_data_list_marge:
            Connect().insert_many(insert_data_list_marge,'shipment')
            
        wx.CallAfter(self.window.SetGauge,100)
        wx.CallAfter(self.window.SetInFileOk)
        Connect().close()
#库存
class kucun_Thread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        threading.Event().clear()
    def run(self):
        result_goods = Connect()._sql_query('*','select b.sku,b.ean,b.xiangma,b.name,SUM(n.count),b.unit,n.allocation \
            from goods n join (select * from base_sku group by sku) b where b.sku=n.sku group by b.sku ORDER by b.sku ')
        if result_goods:
            wx.CallAfter(self.window.SetGridValue,result_goods)
        Connect().close()
#出库数据
class chukushuju_Thread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        threading.Event().clear()
    def run(self):
        shipment_top = Connect()._sql_query('*','select * from shipment_top ORDER BY id desc limit 1')
        wx.CallAfter(self.window.SetTitle,shipment_top)
        result = Connect()._sql_query('*','select * from send_scattered where pid="%s"'%shipment_top[0][1])
        if result:
            wx.CallAfter(self.window.SetEnableCalc,result)
        else:
            wx.CallAfter(self.window.SetEnableTrue)
        Connect().close()
#出库数据计算
class chukuCalcThread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        threading.Event().clear()
    def run(self):
        try:
            pid = Connect()._sql_query('*','select * from shipment_top ORDER BY pid desc limit 1')[0][1]
            
            wx.CallAfter(self.window.LogMessage,u'获取出库数据及规格')
            wx.CallAfter(self.window.SetGauge,10)
            shipment = Connect()._sql_query("*", 'select * from (select mdbm,mdname,pid,SUM(count) as count,sku from \
                shipment where pid="%s" GROUP BY mdbm,sku ) as s inner join (select * from base_sku GROUP BY sku) as b \
            on s.sku = b.sku  ORDER BY s.mdbm,s.sku'%pid)
            #得到库存字典
            goods_list = []
            goods_tup = Connect().select('*','goods')
            for i in goods_tup:
            	goods_list.append([i[0],i[1],i[2],i[3],i[4],i[5]])
            goods_dic = defaultdict(list)
            for goods in goods_list:
                goods_dic[str(goods[2])].append(goods)

            #得到出库表及规格之后按照字段区分整件和酒水
            zhengjian = []
            shipin = []
            
            try:
                for x in shipment:
                    if int(float(str(x[13])))==1:
                        zhengjian.append(x)
                    else:
                        shipin.append(x)
            except Exception, e:
                wx.MessageBox(u'商品类型有误，请检查基础数据"是否整件"列是否有空，只允许0、1。%s'%e,u'',wx.ICON_ERROR);return
            wx.CallAfter(self.window.SetGauge,25)
            wx.CallAfter(self.window.LogMessage,u'正计算整件出库数据')
            delete_id = []
            update_id = []
            task_list = [] #任务列表
            in_goods = []
            for j in zhengjian:
                x_count = int(float(str(j[3]))) #单据数
                if goods_dic.has_key(str(j[4])):
                    for x,i in enumerate(goods_dic[str(j[4])]):
                        if x_count-i[4]>0:
                            task_list.append([j[0],j[1],j[2],int(str(j[3])),j[4],j[6],j[8],j[9],'',i[5],i[4],j[10]])
                            delete_id.append(i[0])
                            x_count = x_count-i[4]
                            del goods_dic[str(j[4])][x]
                        elif x_count-i[4]==0:
                            task_list.append([j[0],j[1],j[2],int(str(j[3])),j[4],j[6],j[8],j[9],'',i[5],i[4],j[10]])
                            delete_id.append(i[0])
                            del goods_dic[str(j[4])][x]
                            x_count = 0
                            break
                        else:
                            task_list.append([j[0],j[1],j[2],int(str(j[3])),j[4],j[6],j[8],j[9],'',i[5],x_count,j[10]])
                            update_id.append([i[0],str(int(i[4])-int(x_count))])
                            goods_dic[str(j[4])][x][4] = int(str(goods_dic[str(j[4])][x][4]))-int(x_count)
                            x_count = 0
                            break
                    if x_count>0:
                        task_list.append([j[0],j[1],j[2],int(str(j[3])),j[4],j[6],j[8],j[9],u'',u'商品缺货',x_count,j[10]])

                else:
                    task_list.append([j[0],j[1],j[2],int(str(j[3])),j[4],j[6],j[8],j[9],u'',u'商品缺货',int(str(j[3])),j[10]])
            
            insert_data_merge = []
            for i in task_list:
                insert_data = {'sku':int(float(int(i[4]))),'ean':i[7],'name':i[5],'send_type':1,'persons':i[8],'allocation':i[9],'count':int(float(int(i[10]))),\
                    'static':0,'pid':i[2],'mdbm':int(float(int(i[0]))),'unit':int(float(int(i[11])))}
                insert_data_merge.append(insert_data)
            if insert_data_merge:
            	Connect().insert_many(insert_data_merge,'send_scattered')
            
            wx.CallAfter(self.window.SetGauge,40)
            wx.CallAfter(self.window.LogMessage,u'计算整件食品数。')
            shipin_zheng = []
            shipin_san = []

            for i in shipin:
                if float(int(str(i[3])))//float(int(str(i[10])))>0:
                    shipin_zheng.append([i,round(float(int(str(i[3])))//float(int(str(i[10])))*float(int(str(i[10]))),2)])
                if float(int(str(i[3])))%float(int(str(i[10])))>0:
                    shipin_san.append([i,float(int(str(i[3])))%float(int(str(i[10])))])
            shipin_zheng_task = []
            
            for j in shipin_zheng:
                x_count = int(float(str(j[1]))) #单据数
                if goods_dic.has_key(str(j[0][4])):
                    for x,i in enumerate(goods_dic[str(j[0][4])]):
                        if x_count-i[4]>0:
                            shipin_zheng_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],i[5],int(float(str(i[4]))),j[0][10]])
                            delete_id.append(i[0])
                            x_count = x_count-i[4]
                            del goods_dic[str(j[0][4])][x]
                        elif x_count-i[4]==0:
                            shipin_zheng_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],i[5],x_count,j[0][10]])
                            delete_id.append(i[0])
                            del goods_dic[str(j[0][4])][x]
                            x_count=0
                            break
                        else:
                            shipin_zheng_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],i[5],x_count,j[0][10]])
                            update_id.append([i[0],str(int(i[4])-int(x_count))])
                            goods_dic[str(j[0][4])][x][4] = int(str(goods_dic[str(j[0][4])][x][4]))-int(x_count)
                            x_count=0
                            break
                    if x_count>0:
                        shipin_zheng_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],u'商品缺货',x_count,j[0][10]])
                else:
                    shipin_zheng_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],u'商品缺货',int(float(str(j[1]))),j[0][10]])

            

            #合并商品总数汇总数量e商品编码i货位j数量
            #合并整件食品 为一行记录a门店i货位j数量
            Merge_SZheng = {} #计算门店商品单品总数
            Merge = {} #计算商品总数
            # if shipin_zheng_task:
            #     shipin_zheng_task.sort(key=itemgetter(8))
            for a,b,c,d,e,f,g,h,i,j,k in shipin_zheng_task:
                # if str(a)=='4508030':
                #     print a,c,c,d,e,f,g,h,i,j,k
                merge_sku_title = str(e)+u"-"+str(i)
                if Merge.get(merge_sku_title):
                    Merge[merge_sku_title] = [a,b,c,d,e,f,g,h,i,int(Merge.get(merge_sku_title)[9])+int(str(j)),k]
                else:
                    Merge[merge_sku_title] = [a,b,c,d,e,f,g,h,i,j,k]
                merge_title = str(a)+"-"+str(e)
                if Merge_SZheng.get(str(merge_title)):
                    Merge_SZheng[str(merge_title)] = \
                    [a,b,c,d,e,f,g,h,\
                    Merge_SZheng.get(str(merge_title))[8]+'('+str(j)+'),'+i,\
                    int(str(Merge_SZheng.get(str(merge_title))[9]))+int(str(j)),k]
                else:
                    Merge_SZheng[str(merge_title)] = [a,b,c,d,e,f,g,h,i,j,k]
            if Merge_SZheng:
                Merge_SZheng_list = [(Merge_SZheng[k], k) for k in Merge_SZheng]
            else:
                Merge_SZheng_list=[]
            if Merge:
                Merge_huizong_list = [(Merge[k], k) for k in Merge]
            else:
                Merge_huizong_list = []
            insert_data_merge = []

            #0门店编码1地址2pid3单品数4编码5名称6条码7箱码8货位9数量10规格
            for k,v in Merge_SZheng_list:
                insert_data = {'sku':int(float(int(k[4]))),'ean':k[7],'name':k[5],\
                'send_type':2,'allocation':k[8],'count':int(float(int(k[9]))),\
                    'static':0,'pid':k[2],'mdbm':int(float(int(k[0]))),'unit':int(float(int(k[10])))}
                insert_data_merge.append(insert_data)
            if insert_data_merge:
                Connect().insert_many(insert_data_merge,'send_scattered')

            wx.CallAfter(self.window.LogMessage,u'计算散货食品数')
            wx.CallAfter(self.window.SetGauge,65)

            shipin_san_task = [] #任务列表
            for j in shipin_san:
                x_count = int(float(str(j[1]))) #单据数
                if goods_dic.has_key(str(j[0][4])):
                    for x,i in enumerate(goods_dic[str(j[0][4])]):
                        
                        if x_count-i[4]>0:
                            shipin_san_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],'',i[5],int(float(str(i[4]))),j[0][10]])
                            delete_id.append(i[0])
                            x_count = x_count-i[4]
                            del goods_dic[str(j[0][4])][x]
                        elif x_count-i[4]==0:
                            shipin_san_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],'',i[5],x_count,j[0][10]])
                            delete_id.append(i[0])
                            del goods_dic[str(j[0][4])][x]
                            x_count=0
                            break
                        else:
                            shipin_san_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],'',i[5],x_count,j[0][10]])
                            update_id.append([i[0],str(int(i[4])-int(x_count))])
                            goods_dic[str(j[0][4])][x][4] = int(str(goods_dic[str(j[0][4])][x][4]))-int(x_count)
                            x_count=0
                            break
                    if x_count>0:
                        shipin_san_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],u'',u'商品缺货',x_count,j[0][10]])
                else:
                    shipin_san_task.append([j[0][0],j[0][1],j[0][2],int(str(j[0][3])),j[0][4],j[0][6],j[0][8],j[0][9],u'',u'商品缺货',int(float(str(j[1]))),j[0][10]])
            wx.CallAfter(self.window.LogMessage,u'保存数据中')
            wx.CallAfter(self.window.SetGauge,85)
            insert_data_merge = []
            for i in shipin_san_task:
                insert_data = {'sku':int(float(int(i[4]))),'ean':i[6],'name':i[5],\
                'send_type':0,'persons':i[8],'allocation':i[9],'count':int(float(int(i[10]))),\
                    'static':0,'pid':i[2],'mdbm':int(float(int(i[0]))),'unit':int(float(int(i[11])))}
                
                insert_data_merge.append(insert_data)
            Connect().insert_many(insert_data_merge,'send_scattered')
            
            t_sql=''
            #清除重复的
            delete_dic = {}
            for i in delete_id:
                if not delete_dic.has_key(str(i)):
                    delete_dic[str(i)] = i
            delete_list= []
            for i in delete_dic.iteritems():
                delete_list.append(i[0])
            update_dic = {}
            for i in update_id:
                if update_dic.has_key(str(i[0])):
                    if int(str(update_dic[str(i[0])][1]))>int(str(i[1])):
                        update_dic[str(i[0])][1] = int(str(i[1]))
                else:
                    update_dic[str(i[0])] = i
                    
            update_list = []
            for i,x in update_dic.iteritems():
                update_list.append(x)
            for i in delete_list:
            	t_sql = str(i)+','+t_sql
            delete_sql = 'delete from goods where id in(%s);'%t_sql[0:-1]
            
            if delete_list:
                Connect().delete('goods',where='id in(%s)'%t_sql[0:-1])
                
            update_sql = 'begin; UPDATE goods set count = CASE id'
            jonId = '';joinV = ''
            for i in update_list:
            	jonId = jonId+' WHEN '+str(i[0])+' THEN '+str(i[1])
            	joinV = str(i[0])+' ,'+joinV
            updateSql = update_sql+jonId+' END WHERE id IN ('+ str(joinV[0:-1])+") ;COMMIT;"
            
            if update_list:
                Connect()._sql_query('*',updateSql)
            
            wx.CallAfter(self.window.LogMessage,u'生成表格中')
            wx.CallAfter(self.window.SetGauge,95)
            file = xlwt.Workbook(encoding='utf8')

            table_shipment = file.add_sheet(u'出库表数据')
            title_shipment = [u'门店编码',u'地址',u'pid',u'单品总数量',u'商品编码',u'id',u'商品名称',u'商品编码',u'商品条码',u'外箱码',u'规格',u'忽略',u'类型',u'是否整件']
            for i,x in enumerate(title_shipment):
                table_shipment.write(0,i,x)
            for i,x in enumerate(shipment):
                for i_index,x_v in enumerate(x):
                    table_shipment.write(i+1,i_index,x_v)
            
            write_send_scattered = Connect().select('*','send_scattered','pid="%s"'%pid)
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
            wx.CallAfter(self.window.LogMessage,u'计算完成，操作结束')
            wx.CallAfter(self.window.SetGauge,100)
            try:
                Connect().close()
            except Exception, e:
                wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)
            wx.MessageBox(u'操作完成,请保存表格',u'提示',wx.OK)
            select_dialog = wx.DirDialog(None, u"选择保存的路径",style=wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
            if select_dialog.ShowModal() == wx.ID_OK:
                file.save(select_dialog.GetPath()+u"/出库数据"+time.strftime(u"%Y%m%d%H%M%S",time.localtime())+".xls")
                select_dialog.Destroy()
        except Exception, e:
            wx.MessageBox(u'系统错误，错误信息：%s'%str(e),u'提示',wx.OK)

#重定位 整件食品
class UpdateZhengShipinThread(threading.Thread):
    def __init__(self,windows,path,pid):
        threading.Thread.__init__(self)
        self.window = windows
        self.path=path
        self.pid = pid
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        filedata = xlrd.open_workbook(self.path,encoding_override='utf-8')
        table = filedata.sheets()[0]
        message =""
        try:
            if table.col(0)[0].value.strip() != u'商品编码':
                message = message+u"第一行名称必须叫‘商品编码’，请返回修改\n"
            if table.col(7)[0].value.strip() != u'重定货位':
                message = message+u"第八行名称必须叫‘重定货位’，请返回修改\n"
        except Exception,f:
            wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。\n%s'%f,u'警告',wx.ICON_ERROR);return
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
        nrows = table.nrows #行数
        table_data_list =[]
        wx.CallAfter(self.window.SetUpdateZhengShipinBtn,u'加载数据中')
        for rownum in range(1,nrows):
            if table.row_values(rownum):
                table_data_list.append(table.row_values(rownum))
        wx.CallAfter(self.window.SetUpdateZhengShipinBtn,u'检查数据中')
        #得到基础数据字典 
        baseSku_list = Connect().select('*','base_sku')
        baseSku_dic = defaultdict(list)
        for b in baseSku_list:
            baseSku_dic[str(b[2])].append(b)
        #得到货位字典 
        huowei_list = Connect().select('*','huowei')
        huowei_dic = defaultdict(list)
        for huowei in huowei_list:
            huowei_dic[str(huowei[1])].append(huowei)
        
        try:
            error_text = ''
            for i,x in enumerate(table_data_list):
                if not baseSku_dic.has_key(str(int(float(str(x[0]))))):
                    error_text = error_text+u'商品编码：%s（%s）有误，没有该基础数据。\n'%(str(int(float(str(x[0])))),x[2])
                if not huowei_dic.has_key(str(x[7])):
                    error_text = error_text+u'货位号：%s有误，没有该货位数据。\n'%x[7]
            if error_text:
                raise Exception,error_text
        except Exception, e:
            wx.MessageBox(u'系统错误，错误信息：\n%s'%e,u'提示',wx.ICON_ERROR);return

        update_sql = 'begin; UPDATE send_scattered set allocation = CASE sku'
        jonId = '';joinV = ''
        for i in table_data_list:
            jonId = jonId+' WHEN '+str(int(float(str(i[0]))))+' THEN "'+str(i[7])+'"'
            joinV = str(int(float(str(i[0]))))+' ,'+joinV
        updateSql = update_sql+jonId+' END WHERE pid = "'+self.pid+'" and \
        send_type=2 and sku IN ('+ str(joinV[0:-1])+") ;COMMIT;"
        if table_data_list:
            Connect()._sql_query('*',updateSql)
            wx.CallAfter(self.window.SetUpdateZhengShipinBtn,u'更新完成')
        else:
            wx.CallAfter(self.window.SetUpdateZhengShipinBtn,u'更新失败')
        
        try:
            Connect().close()
        except Exception, e:
            wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)
#发车表格
class facheSetGrid(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
    def stop(self):
        self.timeToQuit.set()
    def run(self):
    	result = Connect()._sql_query('*',"select c.carNum,c.carName,c.person,c.`load`,c.carLen,c.phone,c.pid,c.note,\
    		c.static,count(m.id) from send_car c join send_mdbm m  WHERE c.pid=m.pid \
		and c.carName=m.carName and c.static in (0,1) GROUP BY c.carName order by c.static,c.id desc")
    	wx.CallAfter(self.window.SetViewGrid,result)
        Connect().close()
#添加车辆pid
class fachePid(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        threading.Event().clear()
    def run(self):
    	result = Connect().select('pid','shipment_top',limit='5',order="id desc")
    	wx.CallAfter(self.window.SetPid,result)
        Connect().close()
#更新车辆发车状态允许发车
class updateSend_car(threading.Thread):
    def __init__(self,windows,data):
        threading.Thread.__init__(self)
        self.window = windows
        self.data = data
        threading.Event().clear()
    def run(self):
    	Connect().update({'static':1},'send_car','pid="%s" and carName="%s"'%(self.data[1],self.data[0]))
        Connect().close()
#更新车辆发车状态完成发车
class wancheng_car(threading.Thread):
    def __init__(self,windows,data):
        threading.Thread.__init__(self)
        self.window = windows
        self.data = data
        threading.Event().clear()
    def run(self):
    	Connect().update({'static':2},'send_car','pid="%s" and carName="%s"'%(self.data[0],self.data[1]))
        Connect().close()
#添加车辆添加提交
class facheAddCar_Thread(threading.Thread):
    def __init__(self,windows,data):
        threading.Thread.__init__(self)
        self.window = windows
        self.data = data
        threading.Event().clear()
        
    def run(self):
        if Connect().get_one('*','shipment_top','pid="%s"'%self.data[0]):
            mdbm_list =  self.data[7].split(',')
            inster_mdbm_mrge = []
            for i in mdbm_list:
            	inster_mdbm_mrge.append({'pid':self.data[0],'mdbm':i,'carName':self.data[1]})
            Connect().insert_many(inster_mdbm_mrge,'send_mdbm')
            insert = {'carNum':self.data[2],'carName':self.data[1],'person':self.data[4],'carLen':self.data[3],\
    		'phone':self.data[5],'pid':self.data[0],'note':self.data[6]}
            Connect().insert(insert,'send_car')
            wx.CallAfter(self.window.SetStatic,1)
        else:
        	wx.CallAfter(self.window.SetStatic,0)
        try:
            Connect().close()
        except Exception, e:
            wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)
class facheAddMdbmThread(threading.Thread):
    def __init__(self,windows,data):
        threading.Thread.__init__(self)
        self.window = windows
        self.data = data
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
    def stop(self):
        self.timeToQuit.set()
    def run(self):
    	result_mdbm = Connect().select('mdbm','shipment',where='pid="%s"'%self.data,group='mdbm')
    	wx.CallAfter(self.window.SetCheckList,result_mdbm)
        try:
            Connect().close()
        except Exception, e:
            wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)
#拆零复核 设置grid
class chailingfuheSetGrid(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        result = Connect()._sql_query('*',"SELECT m.mdbm m ,count(s.id) zong,\
        sum(CASE when s.static=2 then 1 else 0 end) yiwancheng,sum(s.count),\
        sum(s.out_count),sum(s.out_count)/sum(s.count)*100 wanchenglv,\
        c.pid,c.carName\
        from send_scattered s join \
        send_car c join send_mdbm m \
        WHERE  c.static=1 and c.pid = m.pid and c.carName=m.carName and s.pid=m.pid and s.mdbm=m.mdbm \
        and s.send_type=0 \
        GROUP BY m.mdbm ORDER BY wanchenglv,c.carName,zong")
        wx.CallAfter(self.window.SetViewGrid,result)
        try:
            Connect().close()
        except Exception, e:
            wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)
class chailingfuheFrameThread(threading.Thread):
    def __init__(self,windows,info):
        threading.Thread.__init__(self)
        self.window = windows
        self.info = info
        threading.Event().clear()
    def run(self):
        result = Connect()._sql_query('*',"select sku,ean,name,persons,allocation,sum(count),\
            sum(out_count),sum(count)-sum(out_count) as xiangcha  from send_scattered where pid='%s' \
            AND mdbm='%s' and send_type IN(0) group by sku order by xiangcha desc "%(self.info[1],self.info[0]))
        
        wx.CallAfter(self.window.SetViewGrid,result)
        Connect().close()
class chailingfuheInput(threading.Thread):
    def __init__(self,windows,result,addResult,ean,count):
        threading.Thread.__init__(self)
        self.window = windows
        self.result = result #字典
        self.Addresult = addResult #字典
        self.ean = ean
        self.count = count
        threading.Event().clear()
    def run(self):
    	if self.result.has_key(str(self.ean)):
    		if self.Addresult.has_key(str(self.ean)):
    			addCount = self.Addresult[str(self.ean)][6]
    			if addCount+self.count<=self.Addresult[str(self.ean)][7]:
    				self.Addresult[str(self.ean)][6] = self.Addresult[str(self.ean)][6]+self.count
    				self.result[str(self.ean)][7] = self.result[str(self.ean)][7] -self.count
    				self.result[str(self.ean)][6] =self.result[str(self.ean)][6] +self.count
    				wx.CallAfter(self.window.scanViewGrid,self.result.items(),self.Addresult)
    				r_dic = {'ean':self.ean,'name':self.Addresult[str(self.ean)][2],'zong':self.Addresult[str(self.ean)][5],'yidabao':self.result[str(self.ean)][6],'scan':self.Addresult[str(self.ean)][6]}
    				wx.CallAfter(self.window.ThreadMessage,u' 键入%(ean)s(%(name)s)成功,共%(zong)s个,已打包%(yidabao)s,当前已扫描第%(scan)s个。'%r_dic)
    				wx.CallAfter(self.window.ThreadSetStatic,self.count,self.result[str(self.ean)][0],self.result[str(self.ean)][2])
    				winsound.PlaySound("ok.wav", winsound.SND_FILENAME)
    			else:
    				wx.CallAfter(self.window.ThreadMessage,u' 键入%(ean)s(%(name)s)失败,数量超出。'%{'name':self.Addresult[str(self.ean)][2],'ean':self.Addresult[str(self.ean)][1]})
    				winsound.PlaySound("error.wav", winsound.SND_FILENAME)
    		else:
    			if self.count+self.result[str(self.ean)][6]<self.result[str(self.ean)][7]:
    				self.Addresult[str(self.ean)] = copy.deepcopy(self.result[str(self.ean)])#深拷贝 不能直接用等于号
    				self.Addresult[str(self.ean)][6] = 1#6out_count 做为待打包数量
    				self.result[str(self.ean)][6] =self.result[str(self.ean)][6] +self.count
    				self.result[str(self.ean)][7] =self.result[str(self.ean)][7] -self.count
    				wx.CallAfter(self.window.scanViewGrid,self.result.items(),self.Addresult)
    				r_dic = {'ean':self.ean,'name':self.Addresult[str(self.ean)][2],'zong':self.Addresult[str(self.ean)][5],'yidabao':self.result[str(self.ean)][6],'scan':self.Addresult[str(self.ean)][6]}
    				wx.CallAfter(self.window.ThreadMessage,u' 键入%(ean)s(%(name)s)成功,共%(zong)s个,已打包%(yidabao)s,当前已扫描第%(scan)s个。'%r_dic)
    				wx.CallAfter(self.window.ThreadSetStatic,self.count,self.result[str(self.ean)][0],self.result[str(self.ean)][2])
    				winsound.PlaySound("ok.wav", winsound.SND_FILENAME)
    			else:
    				wx.CallAfter(self.window.ThreadMessage,u' 键入%(ean)s(%(name)s)失败,数量超出。'%{'ean':self.ean,'name':self.result[str(self.ean)][2]})
    				winsound.PlaySound("error.wav", winsound.SND_FILENAME)

    	else:
    		wx.CallAfter(self.window.ThreadMessage,u' 键入%s失败,没有改商品记录。'%self.ean)
    		winsound.PlaySound("error.wav", winsound.SND_FILENAME)
#整件复核显示门店
class zhengjianfuheSetGrid(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        result = Connect()._sql_query('*',"SELECT m.mdbm m ,count(s.id) zong,\
        sum(CASE when s.static=2 then 1 else 0 end) yiwancheng,sum(s.count/s.unit),\
        sum(s.out_count/s.unit),sum(s.out_count/s.unit)/sum(s.count/s.unit)*100 wanchenglv,\
        c.pid,c.carName\
        from send_scattered s join \
        send_car c join send_mdbm m \
        WHERE  c.static=1 and c.pid = m.pid and c.carName=m.carName and s.pid=m.pid and s.mdbm=m.mdbm \
        and s.send_type in(1,2,3) \
        GROUP BY m.mdbm ORDER BY wanchenglv,c.carName,zong")
        wx.CallAfter(self.window.SetViewGrid,result)
        try:
            Connect().close()
        except Exception, e:
            wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)
class zhengjianfuheFrameThread(threading.Thread):
    def __init__(self,windows,info):
        threading.Thread.__init__(self)
        self.window = windows
        self.info = info
        threading.Event().clear()
    def run(self):
        result = Connect()._sql_query('*',"select sku,ean,name,persons,allocation,sum(count/unit) as js,unit,\
            sum(count/unit)-sum(out_count/unit) out_js from send_scattered where pid='%s' \
            AND mdbm='%s' and send_type IN(1,2,3) group by sku order by js desc"%(self.info[1],self.info[0]))
        baseSku_list = Connect().select('*','base_sku')

        result_dic = defaultdict(list)
        for x in result:
            result_dic[str(x[1])].append(list(x))  #外箱条码作为key
        baseSku_dic = defaultdict(list)
        for b in baseSku_list:
            baseSku_dic[str(b[4])].append(b) #w外箱条码作为key
        wx.CallAfter(self.window.SetViewGrid,result_dic,baseSku_dic,'')
        Connect().close()
class zhengjianfuheSendBtn(threading.Thread):
    def __init__(self,windows,info):
        threading.Thread.__init__(self)
        self.window = windows
        self.info = info
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        Connect().update({'static':1},'send_scattered','pid="%s" and mdbm="%s" and static=0 and send_type in (1,2,3)'%(self.info[1],self.info[0]))
        wx.CallAfter(self.window.UpdateOk)
        try:
            Connect().close()
        except Exception, e:
            wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)
class zhengjianfuheInput(threading.Thread):
    def __init__(self,windows,mark,ean,result,baseSku):
        threading.Thread.__init__(self)
        self.window = windows
        self.pid = mark[0]
        self.mdbm = mark[1]
        self.ean = ean
        self.baseSku = baseSku
        self.result = result
        threading.Event().clear()
    def run(self):
        resultS = []
        wx.CallAfter(self.window.SetEanInputNull)
        if self.result.has_key(self.ean):
        	resultS = self.result[self.ean][0]
        elif self.baseSku.has_key(self.ean):
            baseSku_sku = self.baseSku[self.ean][0][2]
            for k,v in self.result.items():
                if str(v[0][0])==str(baseSku_sku):
                    resultS = v[0]
        else:
        	wx.CallAfter(self.window.ErrorMessage,u'键入 "%s" 失败,没有该商品的出库数据'%self.ean)
        	winsound.PlaySound("error.wav", winsound.SND_FILENAME);return
        
        if not resultS:
            wx.CallAfter(self.window.ErrorMessage,u'键入 "%s" 失败,没有该商品的出库数据'%self.ean)
            winsound.PlaySound("error.wav", winsound.SND_FILENAME);return

        if int(float(str(resultS[7])))-1>=0:
            self.result[resultS[1]][0][7] = self.result[resultS[1]][0][7]-1
            wx.CallAfter(self.window.SuccessMessage,resultS,self.ean)
            wx.CallAfter(self.window.SetViewGrid,self.result,self.baseSku,self.ean)
            winsound.PlaySound("ok.wav", winsound.SND_FILENAME)

            update_data = {'out_count':(int(float(str(resultS[5])))-int(float(str(resultS[7]))))*float(str(resultS[6]))}
            update_result = Connect().select('*','send_scattered',where='pid="%s" and mdbm="%s" and sku="%s" '%(self.pid,self.mdbm,resultS[0]))
            for i in update_result:
                if int(float(str(i[15])))+int(float(str(i[13])))<=int(float(str(i[6]))):
                    Connect().update(update_data,'send_scattered','id="%s" '%i[0])
                    Connect().insert({'sku':str(resultS[0]),'count':str(resultS[6]),'pid':self.pid,'mdbm':self.mdbm,'time':time.time()},'out_fuhe')
                    Connect().close()
                    return
            Connect().close()
        else:
            wx.CallAfter(self.window.ErrorMessage,u'键入 "%(ean)s"(%(name)s) 失败,数量超出。'%{'ean':self.ean,'name':resultS[2]})
            winsound.PlaySound("error.wav", winsound.SND_FILENAME);return

class BaskuFileBtnThread(threading.Thread):
    def __init__(self,windows,path):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
        self.path = path
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        filedata = xlrd.open_workbook(self.path,encoding_override='utf-8')
        table = filedata.sheets()[0]
        message =""
        try:
            if table.col(0)[0].value.strip() != u'商品编码':
                message = u"第一行名称必须叫‘商品编码’，请返回修改"
            if table.col(1)[0].value.strip() != u'商品条码':
                message = u"第二行名称必须叫‘商品条码’，请返回修改"
            if table.col(2)[0].value.strip() != u'外箱条码':
                message = u"第三行名称必须叫‘外箱条码’，请返回修改"
            if table.col(3)[0].value.strip() != u'商品名称':
                message = u"第四行名称必须叫‘商品名称’，请返回修改"
            if table.col(4)[0].value.strip() != u'规格':
                message = u"第五行名称必须叫‘规格’，请返回修改"
            if table.col(5)[0].value.strip() != u'类型':
                message = u"第六行名称必须叫‘类型’，请返回修改"
            
        except Exception,f:
            wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
        nrows = table.nrows #行数
        table_data_list =[]
        for rownum in range(1,nrows):
            if table.row_values(rownum):
                table_data_list.append(table.row_values(rownum))
        insert_data_list_marge = []
        try:
            for i,x in enumerate(table_data_list):
                insert_data_list = {'name':x[3],'sku':int(float(str(x[0]))),'ean':x[1],'xiangma':x[2],'unit':int(float(str(x[4]))),\
                    'is_zheng':int(float(str(x[5])))}
                insert_data_list_marge.append(insert_data_list)
            
        except Exception, e:
            wx.MessageBox(u'转换类型错误，请检查数据是否有误%s'%e,u'警告',wx.ICON_ERROR);return

        if insert_data_list_marge:
            Connect().insert_many(insert_data_list_marge,'base_sku')
            wx.CallAfter(self.window.SetOkText)
            try:
                Connect().close()
            except Exception, e:
                wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)
class baseSkuViewGridThread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        result = Connect().select('*','base_sku',order='sku')
        wx.CallAfter(self.window.SetViewGrid,result)
        try:
            Connect().close()
        except Exception, e:
            wx.MessageBox(u'关闭数据库连接错误：%s'%e,u'警告',wx.ICON_ERROR)        
class baseSkuDeleteRow(threading.Thread):
    def __init__(self,windows,data):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
        self.data = data
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        if self.data:
            map_data = (str(self.data[0]),str(self.data[1]),str(self.data[2]),str(self.data[3]))
            result = Connect().delete('base_sku','sku="%s" and ean="%s" and xiangma="%s" and unit="%s"'%map_data)
            wx.CallAfter(self.window.DeleteRowOk,result)
            Connect().close()
class BaskuUpdateBtnThread(threading.Thread):
    def __init__(self,windows,data):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
        self.data = data
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        update_data = {'ean':self.data[1][1],'xiangma':self.data[1][2],'name':self.data[1][3],'unit':self.data[1][4],'is_zheng':self.data[1][5]}
        result = Connect().update(update_data,'base_sku','sku="%s" and ean="%s" and xiangma="%s" '%(self.data[0][0],self.data[0][1],self.data[0][2]))
        wx.CallAfter(self.window.UpdateRowOk,result) 
        Connect().close()
class Base_skuSaveThread(threading.Thread):
    def __init__(self,windows,data):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
        self.data = data
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        if Connect().select('*','base_sku','sku="%s" and ean="%s" and xiangma="%s"'%(str(self.data[0]),str(self.data[1]),str(self.data[2]))):
            wx.CallAfter(self.window.AddBaseSkuMessage,u'添加失败，符合该编码/条码/外箱条码的记录已经存在.')
        else:
            insert_data ={'sku':str(self.data[0]),'ean':str(self.data[1]),'xiangma':str(self.data[2]),'name':str(self.data[3]),'unit':str(self.data[4]),'is_zheng':str(self.data[5])}
            if Connect().insert(insert_data,'base_sku'):
                wx.CallAfter(self.window.AddBaseSkuMessage,u'添加成功1条记录')
            else:
                wx.CallAfter(self.window.AddBaseSkuMessage,u'添加失败')
        Connect().close()
class allocationGridThread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()
    def stop(self):
        self.timeToQuit.set()
    def run(self):
        result = Connect().select('*','huowei',order='sort')
        wx.CallAfter(self.window.SetViewGrid,result)
        Connect().close()
class AllocationFileAddThread(threading.Thread):
    def __init__(self,windows,path):
        threading.Thread.__init__(self)
        self.window = windows
        threading.Event().clear()
        self.path = path
    def run(self):
        filedata = xlrd.open_workbook(self.path,encoding_override='utf-8')
        table = filedata.sheets()[0]
        message =""
        try:
            if table.col(0)[0].value.strip() != u'货位名称':
                message = u"第一行名称必须叫‘货位名称’，请返回修改"
            if table.col(1)[0].value.strip() != u'货位类型':
                message = u"第二行名称必须叫‘货位类型’，请返回修改"
            if table.col(2)[0].value.strip() != u'货位排序':
                message = u"第三行名称必须叫‘货位排序’，请返回修改"
            if table.col(3)[0].value.strip() != u'货位区域':
                message = u"第四行名称必须叫‘货位区域’，请返回修改"
            
        except Exception,f:
            wx.MessageBox(u'表单列数读取错误，请检查表格列数是否正确。%s'%f,u'警告',wx.ICON_ERROR);return
        if message !="":
            wx.MessageBox(message,u'警告',wx.ICON_ERROR);return
        nrows = table.nrows #行数
        table_data_list =[]
        for rownum in range(1,nrows):
            if table.row_values(rownum):
                table_data_list.append(table.row_values(rownum))
        insert_data_list_marge = []
        try:
            for i,x in enumerate(table_data_list):
            	if x[0]=='' or x[0]=='' or x[0]=='' or x[0]=='':
            		wx.MessageBox(u'读取行错误，有空数据在：%s行'%e%str(i),u'警告',wx.ICON_ERROR);return
                insert_data_list = {'title':x[0],'type':x[1],'sort':int(float(str(x[2]))),'quyu':x[3]}
                insert_data_list_marge.append(insert_data_list)
            
        except Exception, e:
            wx.MessageBox(u'转换类型错误，请检查数据是否有误%s'%e,u'警告',wx.ICON_ERROR);return

        if insert_data_list_marge:
            Connect().insert_many(insert_data_list_marge,'huowei')
            wx.CallAfter(self.window.SetOkText)
            Connect().close()
class DabaoGridThread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        threading.Event().clear()
    def run(self):
    	result = Connect().select('*','scattered_pack',limit='5000',order='id desc')
    	wx.CallAfter(self.window.SetViewGrid,result)
    	Connect().close()

class JianxuanToExcelThread(threading.Thread):
    def __init__(self,windows,pid):
        threading.Thread.__init__(self)
        self.window = windows
        self.pid = pid
        threading.Event().clear()
    def run(self):
        result = Connect().select('*','send_scattered',where='pid="%s"'%self.pid)
        wx.CallAfter(self.window.SetExcelFile,result)
        Connect().close()



        








        
    	


