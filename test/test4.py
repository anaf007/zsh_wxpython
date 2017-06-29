__author__ = 'anngle'
#coding=utf-8
import wx
from wx.lib.delayedresult import startWorker
import threading

class TestWindow(wx.Frame):
    def __init__(self, title="Test Window"):
        self.app = wx.App(False)
        wx.Frame.__init__(self, None, -1, title)
        panel = wx.Panel(self)
        self.btnBegin = wx.Button(panel, -1, label='Begin')
        self.Bind(wx.EVT_BUTTON, self.handleButton, self.btnBegin)
        self.txtCtrl = wx.TextCtrl(panel, style=wx.TE_READONLY, size=(300, -1))
        vsizer = wx.BoxSizer(wx.VERTICAL)
        vsizer.Add(self.btnBegin, 0, wx.ALL, 5)
        vsizer.Add(self.txtCtrl, 0, wx.ALL, 5)
        panel.SetSizer(vsizer)
        vsizer.SetSizeHints(self)
        self.Show()

    #处理Begin按钮事件
    def handleButton(self, event):
        self.workFunction()

    #开始执行耗时处理，有继承类实现
    def workFunction(self, *args, **kwargs):
        print'In workFunction(), Thread=', threading.currentThread().name
        print '\t*args:', args
        print '\t**kwargs:', kwargs

        self.btnBegin.Enable(False)

    #耗时处理处理完成后，调用该函数执行画面更新显示，由继承类实现
    def consumer(self, delayedResult, *args, **kwargs):
        print 'In consumer(), Thread=', threading.currentThread().name
        print '\tdelayedResult:', delayedResult
        print '\t*args:', args
        print '\t**kwargs:', kwargs

        self.btnBegin.Enable(True)

    #模拟进行耗时处理并返回处理结果，给继承类使用
    def doSomeThing(self, *args, **kwargs):
        print'In doSomeThing(), Thread=', threading.currentThread().name
        print '\t*args:', args
        print '\t**kwargs:', kwargs

        count = 0
        while count < 10**8:
            count += 1

        return count



class DoSomeThingInSeperateThread(TestWindow):
     def __init__(self):
         TestWindow.__init__(self, 'Do Something In Seperate Thread')
         self.jobId = 100
         self.app.MainLoop()

     #调用wx.lib.delayedresult.startWorker，把函数处理放到单独的线程中去完成。
     #完成后会自动调用consumer进行画面更新处理
     #startWorker函数的各参数接下来分析
     def workFunction(self, *args, **kwargs):
         TestWindow.workFunction(self, args, kwargs)
         startWorker(self.consumer, self.doSomeThing, jobID=self.jobId)

     #第一参数要为DelayedResult类型，即包含处理结果或异常信息，调用get接口取得。
     def consumer(self, delayedResult, *args, **kwargs):
         TestWindow.consumer(self, delayedResult, args, kwargs)
         assert(self.jobId == delayedResult.getJobID())
         try:
             var = delayedResult.get()
         except Exception, e:
             print 'Result for job %s raised exception:%s' %(delayedResult.getJobID, e)

         self.txtCtrl.SetValue(str(var))

if __name__ == '__main__':
    win = DoSomeThingInSeperateThread()