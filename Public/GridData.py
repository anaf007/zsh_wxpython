#coding=utf-8
# import wx
import wx.grid
"""
用于数据表格生成
"""

class GridData(wx.grid.PyGridTableBase):
    def __init__(self,data,cols):
        wx.grid.PyGridTableBase.__init__(self)
        self._data = data
        self._cols = cols
        
    _highlighted = set()
    
    def GetColLabelValue(self, col):
        return self._cols[col]

    def GetNumberRows(self):
        return len(self._data)

    def GetNumberCols(self):
        return len(self._cols)

    def GetValue(self, row, col):
        return self._data[row][col]

    def SetValue(self, row, col, val):
        self._data[row][col] = val

    def GetAttr(self, row, col, kind):
        attr = wx.grid.GridCellAttr()
#         attr.SetBackgroundColour(wx.WHITE if row in self._highlighted else wx.WHITE)
        return attr
    def AppendRows(self, numRows=1): 
        return (self.GetRowCount() + numRows)
    def set_value(self, row, col, val):
        self._highlighted.add(row)
        self.SetValue(row, col, val)
    def DeleteRows(self,pos=0,numRows=1):
        if self._data is None or len(self._data) == 0:
            return False
# 
#         for rowNum in range(0,numRows):
#             self._data.remove(self._data[pos+rowNum])
       
        gridView = self.GetView()
        gridView.BeginBatch()
        deleteMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_NOTIFY_ROWS_DELETED,pos,numRows)
        gridView.ProcessTableMessage(deleteMsg)
        gridView.EndBatch()
        getValueMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
        gridView.ProcessTableMessage(getValueMsg)
 
#         if self.onGridValueChanged:
#             self.onGridValueChanged()
           
        return True
    
    def InsertRows(self,pos=1,numRows=1):
        
#         for num in range(0,numRows):
#             newData={};
#             newData[u'lable'] = u''     
#             self._data.insert(pos,newData)
            
        gridView = self.GetView()
        gridView.BeginBatch()
        insertMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_NOTIFY_ROWS_INSERTED,pos,numRows)
        gridView.ProcessTableMessage(insertMsg)
        gridView.EndBatch()
        getValueMsg = wx.grid.GridTableMessage(self,wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
        gridView.ProcessTableMessage(getValueMsg)        
               
        return True