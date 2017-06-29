#coding=utf-8
__author__ = 'anngle'
import wx
from SQLet import *
from Public.GridData import GridData
class bangding_main(wx.MDIChildFrame):
    def __init__(self,parent):
        wx.MDIChildFrame.__init__(self,parent,title=u'区域人员绑定')
        panel = wx.Panel(self)
        panel.SetBackgroundColour('White')
        self.persons_data = []
        # self.type_data = [u'整件区',u'拆零区',u'库存区']
        self.quyu_data = []
        self._data = []
        res_quyu = Connect().select('quyu','huowei',group='quyu')
        for x in res_quyu:
            self.quyu_data.append(x[0])
        _cols = (u"ID",u'人员',u"区域")
        res_user = Connect().select('username','users')
        for x in res_user:
            self.persons_data.append(x[0])
        res_dic = Connect().select('*','bangding_user')
        for x in res_dic:
            self._data.append([x[0],x[2],x[1]])
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(panel)
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        self.persons_choice=wx.Choice(panel,-1,choices=self.persons_data)
        self.quyu_choice=wx.Choice(panel,-1,choices=self.quyu_data)
        # self.type_choice=wx.Choice(panel,-1,choices=self.type_data) #类型
        saveBtn = wx.Button(panel,-1,u'添加')
        topSizer = wx.BoxSizer()
        topSizer.Add(self.persons_choice,1,wx.ALL,5)
        # topSizer.Add(self.type_choice,1,wx.ALL,5)
        topSizer.Add(self.quyu_choice,1,wx.ALL,5)
        topSizer.Add(saveBtn,1,wx.ALL,5)
        BodySizer = wx.BoxSizer(wx.VERTICAL) #垂直

        BodySizer.Add(topSizer)
        BodySizer.Add(self.grid,1,wx.ALL,5)
        panel.SetSizer(BodySizer)
        BodySizer.Fit(self)
        BodySizer.SetSizeHints(self)

        self.grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK,self.OnShowPopup) #在grid行头右键点击
        saveBtn.Bind(wx.EVT_BUTTON, self.OnSaveBtn) #点击保存按钮


    def OnSaveBtn(self,evt):
        data_persons = self.persons_choice.GetString(self.persons_choice.GetSelection())
        data_quyu = self.quyu_choice.GetString(self.quyu_choice.GetSelection())
        if not Connect().get_one('*','bangding_user',where='persons="%s" and district="%s"'%(data_persons,data_quyu)):
            Connect().insert({'persons':data_persons,'district':data_quyu},'bangding_user')
            wx.MessageBox(u'添加完成',u'提示',wx.OK)
        else:
            wx.MessageBox(u'人员“%s”已绑定在该区域，请勿重复绑定'%data_persons,u'提示',wx.ICON_ERROR)

    def OnShowPopup(self,evt):
        evt_value = evt.GetRow()
        if evt_value==-1:
            return
        data_list =  []
        data_list.append([self._data[evt_value][0]])
        popMenu=wx.Menu()#创建一个菜单
        delete_pop_menu = popMenu.Append(-1,u'删除')
        self.Bind(wx.EVT_MENU,lambda evt,mark=data_list : self.OnOpenPopMenu(evt,mark) ,delete_pop_menu )
        self.grid.PopupMenu(popMenu)

    def OnOpenPopMenu(self,evt,mark):
        result =  Connect().delete('bangding_user',where='id="%s"'%mark[0][0])
        if result:
            wx.MessageBox(u'删除成功，请重新关闭在打开窗口查看',u'提示',wx.ICON_ASTERISK)
        else:
            wx.MessageBox(u'删除失败，请重新关闭在打开窗口查看',u'提示',wx.ICON_ERROR);return




