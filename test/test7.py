import wx
import wx.combo

class CustomPopup(wx.combo.ComboPopup):

    def Create(self, parent):
        # Create the popup with a bunch of radiobuttons
        self.panel = wx.Panel(parent)
        sizer = wx.GridSizer(cols=3)
        for x in range(10):
            # r = wx.RadioButton(self.panel, label="Element "+str(x))
            r = wx.CheckBox(self.panel, label="Element "+str(x))
            r.Bind(wx.EVT_RADIOBUTTON, self.on_selection)
            sizer.Add(r)
        self.panel.SetSizer(sizer)

        # Handle keyevents
        self.panel.Bind(wx.EVT_KEY_UP, self.on_key)

    def GetControl(self):
        return self.panel

    def GetAdjustedSize(self, minWidth, prefHeight, maxHeight):
        return wx.Size(280, 150)

    def on_key(self, evt):
        print "B"
        if evt.GetEventObject() is self.panel:
            # Trying to redirect the key event to the combo.. But this always returns false :(
            print self.GetCombo().GetTextCtrl().GetEventHandler().ProcessEvent(evt)
        print "a"
        evt.Skip()

    def on_selection(self, evt):
        self.Dismiss()
        # wx.MessageBox("Selection made")


class CustomFrame(wx.Frame):

    def __init__(self):
        # Toolbar-shaped frame with a ComboCtrl
        wx.Frame.__init__(self, None, -1, "Test", size=(550,550))
        # pnl  = wx.Panel(self,-1)
        combo = wx.combo.ComboCtrl(self)
        popup = CustomPopup()
        combo.SetPopupControl(popup)
        # self.btn = wx.Button(pnl,-1,u"button")
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(combo, 0)
        # sizer.Add(self.btn,1)
        # pnl.SetSizer(sizer)
        sizer.SetSizeHints(self)
        sizer.Fit(self)
        print popup.GetComboCtrl()


if __name__ == '__main__':
    app = wx.PySimpleApp()
    CustomFrame().Show()
    app.MainLoop()