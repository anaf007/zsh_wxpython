#coding=utf-8
import wx
import wx.grid as gridlib
from Public.GridData import *
#######################################################################
class MyForm(wx.Frame):
    """"""
 
    #----------------------------------------------------------------------
    def __init__(self):
        """Constructor"""
        wx.Frame.__init__(self, parent=None, title="A Simple Grid")
        panel = wx.Panel(self)

        self._data = []
        _cols = (u"ID",u"商品编码",u"商品条码",u"外箱码",u"商品名称",u"类型",u"是否整件",u"每箱个数",u"备注(按件按包出)")
        self._data.append(['1',2,3,4,5,6,7,8,9])
        self.data = GridData(self._data,_cols)
        self.grid = wx.grid.Grid(panel)
        self.grid.SetTable(self.data)
        self.grid.AutoSize()
        btn  = wx.Button(panel,-1,"btn_top")

        body_sizer = wx.BoxSizer(wx.VERTICAL)
        top_sizer = wx.BoxSizer()
        main_sizer = wx.BoxSizer()
        main_sizer.Add(self.grid,1,wx.EXPAND|wx.ALIGN_RIGHT)
        body_sizer.Add(top_sizer)
        body_sizer.Add(main_sizer)
        top_sizer.Add(btn)
        panel.SetSizer(body_sizer)
 
if __name__ == "__main__":
    app = wx.PySimpleApp()
    frame = MyForm().Show()
    app.MainLoop()