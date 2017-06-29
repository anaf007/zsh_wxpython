#coding=utf-8
__author__ = 'anngle'
import wx,webbrowser
from Data_Manage.sku import *
from Data_show.InData import *
from Data_show.outData import *
from Data_Manage.bangding import *
from Data_Manage.set_message import *
from Data_show.Calculation import *
from Data_show.cuhkuduibi import *
from now_goods import *
from allocation import *
from Data_Manage.goods_adj import buhuo
from Scattered_pack import sanhuo_pack
from Data_Manage.send_car import send_car
from Bulk import bulk
from zhouzhuanxiang import zhuzhuanxiang_main as zzxm
from zhengjian import Zhengjian as zj

class Main(wx.MDIParentFrame):
    def __init__(self):
        wx.MDIParentFrame.__init__(self,None,-1,u'中石化辅助系统',size=(500,500))
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        self.createMenuBar()

    def menuData(self): # 菜单数据
        return ((u'&仓库操作',
            (u'&拆零复核','', self.OnFuhe),
            (u'&出库数量对比', '', self.Onchukuduibi),
            (u'&计算导出表格','', self.OnFahuowei),
            (u'&整件食品分配','', self.OnFenpei),
            (u'&整件复核','', self.OnZhengjianfuhe),
            ),
                (u'&数据管理',
            (u'&基础数据','', self.OnSku),
            (u'&区域绑定','', self.OnBandding),
            (u'&整件拣选信息','', self.OnSetMessage),
            (u'&库存查询', '', self.OnNowGoods),
            (u'&发车管理', '', self.OnSend_Car),
            (u'&补货任务', '', self.OnGoods_Adj)),
                (u'&数据查看',
            (u'&入库单据','', self.OnInData),
            (u'&出库单据', '', self.OnOutData),
            (u'&货位管理', '', self.OnHuowei),
            (u'&周转箱', '', self.Onzhouzhuanxiang),
            (u'&散货拣选', '', self.Onsanhuo_jianxuan),
            (u'&打包管理','', self.OnPack),),

        )
    # 创建菜单
    def createMenuBar(self):
        menuBar = wx.MenuBar()
        for eachMenuData in self.menuData():
            menuLabel = eachMenuData[0]
            menuItems = eachMenuData[1:]
            menuBar.Append(self.createMenu(menuItems), menuLabel)
        self.SetMenuBar(menuBar)

    def createMenu(self, menuData):
        menu = wx.Menu()
        for eachLabel, eachStatus, eachHandler in menuData:
            if not eachLabel:
                menu.AppendSeparator()
                continue
            menuItem = menu.Append(-1, eachLabel, eachStatus)
            self.Bind(wx.EVT_MENU, eachHandler, menuItem)
        return menu

    def OnFuhe(self, event):bulk(self).Show()
        


    def OnFahuowei(self,evt):calculation(self).Show()
    def OnSku(self, event):sku_main(self).Show()
    def OnInData(self, event):InData_main(self).Show()
    def OnOutData(self, event):outData_main(self).Show()
    def OnNowGoods(self, event): goods_main(self).Show()
    def OnBandding(self, event): bangding_main(self).Show()
    def Onchukuduibi(self, event): Outbound_comparison(self).Show()
    def OnSetMessage(self, event): set_message_main(self).Show()
    def Onsanhuo_jianxuan(self, event): sanhuo_jianxuan(self).Show()
    def OnPack(self, event): sanhuo_pack(self).Show()
    def OnGoods_Adj(self, event): buhuo(self).Show()
    def OnHuowei(self,evt):huowei(self).Show()
    def OnSend_Car(self,evt):send_car(self).Show()
    def OnFenpei(self,evt):fenpei(self).Show()
    def Onzhouzhuanxiang(self,evt):zzxm(self).Show()
    def OnZhengjianfuhe(self,evt):zj(self).Show()
    
    def OnCloseWindow(self, event):
        self.Destroy()


if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = Main()
    frame.Show()
    app.MainLoop()