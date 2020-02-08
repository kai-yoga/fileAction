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
        if action=='合并xlsx':
            check=wife.checkVersion(path,pathaction)
            if check:
                wx.MessageBox('xlsx，xls同时存在的文件有{},这部分文件会被删除!,若都要保留，请先修改这部分文件的文件名'.format(check))
        wx.MessageBox('点击“OK”开始合并!')
        print('开始合并')
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

