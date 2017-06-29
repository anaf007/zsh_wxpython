import wx
 
class PyCalc(wx.App):
    def __init__(self, redirect=False, filename=None):
        wx.App.__init__(self, redirect, filename)
 
    def OnInit(self):
        # create frame here
        self.frame = wx.Frame(None, wx.ID_ANY, title="Calculator")
        panel = wx.Panel(self.frame, wx.ID_ANY)
        self.displayTxt = wx.TextCtrl(panel, wx.ID_ANY, "0", 
                                      size=(155,-1),
                                      style=wx.TE_RIGHT|wx.TE_READONLY)
        size=(35, 35)
        zeroBtn = wx.Button(panel, wx.ID_ANY, "0", size=size)
        oneBtn = wx.Button(panel, wx.ID_ANY, "1", size=size)
        twoBtn = wx.Button(panel, wx.ID_ANY, "2", size=size)
        threeBtn = wx.Button(panel, wx.ID_ANY, "3", size=size)
        fourBtn = wx.Button(panel, wx.ID_ANY, "4", size=size)
        fiveBtn = wx.Button(panel, wx.ID_ANY, "5", size=size)
        sixBtn = wx.Button(panel, wx.ID_ANY, "6", size=size)
        sevenBtn = wx.Button(panel, wx.ID_ANY, "7", size=size)
        eightBtn = wx.Button(panel, wx.ID_ANY, "8", size=size)
        nineBtn = wx.Button(panel, wx.ID_ANY, "9", size=size)
        zeroBtn.Bind(wx.EVT_BUTTON, self.method1)
        oneBtn.Bind(wx.EVT_BUTTON, self.method2)
        twoBtn.Bind(wx.EVT_BUTTON, self.method3)
        threeBtn.Bind(wx.EVT_BUTTON, self.method4)
        fourBtn.Bind(wx.EVT_BUTTON, self.method5)
        fiveBtn.Bind(wx.EVT_BUTTON, self.method6)
        sixBtn.Bind(wx.EVT_BUTTON, self.method7)
        sevenBtn.Bind(wx.EVT_BUTTON, self.method8)
        eightBtn.Bind(wx.EVT_BUTTON, self.method9)
        nineBtn.Bind(wx.EVT_BUTTON, self.method10)
        divBtn = wx.Button(panel, wx.ID_ANY, "/", size=size)
        multiBtn = wx.Button(panel, wx.ID_ANY, "*", size=size)
        subBtn = wx.Button(panel, wx.ID_ANY, "-", size=size)
        addBtn = wx.Button(panel, wx.ID_ANY, "+", size=(35,100))
        equalsBtn = wx.Button(panel, wx.ID_ANY, "Enter", size=(35,100))
        divBtn.Bind(wx.EVT_BUTTON, self.method11)
        multiBtn.Bind(wx.EVT_BUTTON, self.method12)
        addBtn.Bind(wx.EVT_BUTTON, self.method13)
        subBtn.Bind(wx.EVT_BUTTON, self.method14)
        equalsBtn.Bind(wx.EVT_BUTTON, self.method15)
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        masterBtnSizer = wx.BoxSizer(wx.HORIZONTAL)
        vBtnSizer = wx.BoxSizer(wx.VERTICAL)
        numSizer  = wx.GridBagSizer(hgap=5, vgap=5)
        numSizer.Add(divBtn, pos=(0,0), flag=wx.CENTER)
        numSizer.Add(multiBtn, pos=(0,1), flag=wx.CENTER)
        numSizer.Add(subBtn, pos=(0,2), flag=wx.CENTER)
        numSizer.Add(sevenBtn, pos=(1,0), flag=wx.CENTER)
        numSizer.Add(eightBtn, pos=(1,1), flag=wx.CENTER)
        numSizer.Add(nineBtn, pos=(1,2), flag=wx.CENTER)
        numSizer.Add(fourBtn, pos=(2,0), flag=wx.CENTER)
        numSizer.Add(fiveBtn, pos=(2,1), flag=wx.CENTER)
        numSizer.Add(sixBtn, pos=(2,2), flag=wx.CENTER)
        numSizer.Add(oneBtn, pos=(3,0), flag=wx.CENTER)
        numSizer.Add(twoBtn, pos=(3,1), flag=wx.CENTER)
        numSizer.Add(threeBtn, pos=(3,2), flag=wx.CENTER)
        numSizer.Add(zeroBtn, pos=(4,1), flag=wx.CENTER)        
        vBtnSizer.Add(addBtn, 0)
        vBtnSizer.Add(equalsBtn, 0)
        masterBtnSizer.Add(numSizer, 0, wx.ALL, 5)
        masterBtnSizer.Add(vBtnSizer, 0, wx.ALL, 5)
        mainSizer.Add(self.displayTxt, 0, wx.ALL, 5)
        mainSizer.Add(masterBtnSizer)
        panel.SetSizer(mainSizer)
        mainSizer.Fit(self.frame)
        self.frame.Show()
        return True
 
    def method1(self, event):
        pass
 
    def method2(self, event):
        pass
 
    def method3(self, event):
        pass
 
    def method4(self, event):
        pass
 
    def method5(self, event):
        pass
 
    def method6(self, event):
        pass
 
    def method7(self, event):
        pass
 
    def method8(self, event):
        pass
 
    def method9(self, event):
        pass
 
    def method10(self, event):
        pass
 
    def method13(self, event):
        pass
 
    def method14(self, event):
        pass
 
    def method12(self, event):
        pass
 
    def method11(self, event):
        pass
 
    def method15(self, event):
        pass
 
 
def main():
    app = PyCalc()
    app.MainLoop()
 
if __name__ == "__main__":
    main()