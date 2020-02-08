import wxBaseGui as wbgui
import wx
import wife
class WifeUI(wbgui.MyFrame1):

    def __int__(self,parent):
        wbgui.MyFrame1.__init__(self,parent)
    def onConfirm(self, event):
        action=self.m_choice1.GetString(self.m_choice1.GetCurrentSelection())
        path=self.m_dirPicker1.GetPath()
        pathaction=self.m_comboBox1.GetValue()
        sheetnums=self.m_comboBox2.GetValue()
        resname=self.m_textCtrl1.GetValue()
        ###call interface
        res=wife.interface(action=action,path=path,pathaction=pathaction,sheetnums=sheetnums,resname=resname)
        wx.MessageBox('合并成功，文件保存于{}'.format(res))
    def onCancle(self, event):
        self.Close(force=True)


def main():
    app=wx.App()
    wife=WifeUI(None)
    wife.Show(True)
    app.MainLoop()

if __name__ == "__main__":
    main()

