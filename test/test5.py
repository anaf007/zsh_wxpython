__author__ = 'anngle'
#coding=utf-8
#coding=utf-8
from test.test4 import TestWindow
class DoSomeThingInGUI():
     def __init__(self):
         TestWindow.__init__(self, 'Do Something In GUI Thread')
         self.app.MainLoop()

     #执行doSomeThing()，并在执行完成后，主动调用consumer()更新画面显示
     def workFunction(self, *args, **kwargs):
         TestWindow.workFunction(self, args, kwargs)
         var = self.doSomeThing()
         self.consumer(var)

     def consumer(self, delayedResult, *args, **kwargs):
         TestWindow.consumer(self, delayedResult, args, kwargs)
         self.txtCtrl.SetValue(str(delayedResult))

if __name__ == '__main__':
    win = DoSomeThingInGUI()