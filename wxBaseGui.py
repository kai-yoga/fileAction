# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc


###########################################################################
## Class MyFrame1
###########################################################################

class MyFrame1(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"小米粥", pos=wx.DefaultPosition, size=wx.Size(895, 595),
                          style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        gSizer1 = wx.GridSizer(0, 2, 0, 0)

        self.m_staticText1 = wx.StaticText(self, wx.ID_ANY, u"操作类型", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText1.Wrap(-1)
        gSizer1.Add(self.m_staticText1, 0, wx.ALL, 5)

        m_choice1Choices = [u"合并xlsx"]
        self.m_choice1 = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice1Choices, 0)
        self.m_choice1.SetSelection(0)
        gSizer1.Add(self.m_choice1, 0, wx.ALL, 5)

        self.m_staticText2 = wx.StaticText(self, wx.ID_ANY, u"操作路径", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText2.Wrap(-1)
        gSizer1.Add(self.m_staticText2, 0, wx.ALL, 5)

        self.m_dirPicker1 = wx.DirPickerCtrl(self, wx.ID_ANY, wx.EmptyString, u"Select a folder", wx.DefaultPosition,
                                             wx.DefaultSize, wx.DIRP_DEFAULT_STYLE)
        gSizer1.Add(self.m_dirPicker1, 0, wx.ALL, 5)

        self.m_staticText3 = wx.StaticText(self, wx.ID_ANY, u"路径方式", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText3.Wrap(-1)
        gSizer1.Add(self.m_staticText3, 0, wx.ALL, 5)

        m_comboBox1Choices = [u"不包含子文件夹", u"包含子文件夹"]
        self.m_comboBox1 = wx.ComboBox(self, wx.ID_ANY, u"请选择", wx.DefaultPosition, wx.DefaultSize, m_comboBox1Choices,
                                       0)
        gSizer1.Add(self.m_comboBox1, 0, wx.ALL, 5)

        self.m_staticText4 = wx.StaticText(self, wx.ID_ANY, u"结果方式", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText4.Wrap(-1)
        gSizer1.Add(self.m_staticText4, 0, wx.ALL, 5)

        m_comboBox2Choices = [u"一个工作表", u"多个工作表"]
        self.m_comboBox2 = wx.ComboBox(self, wx.ID_ANY, u"请选择", wx.DefaultPosition, wx.DefaultSize, m_comboBox2Choices,
                                       0)
        gSizer1.Add(self.m_comboBox2, 0, wx.ALL, 5)

        self.m_staticText5 = wx.StaticText(self, wx.ID_ANY, u"结果文件关键字", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText5.Wrap(-1)
        gSizer1.Add(self.m_staticText5, 0, wx.ALL, 5)

        self.m_textCtrl1 = wx.TextCtrl(self, wx.ID_ANY, u"res", wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer1.Add(self.m_textCtrl1, 0, wx.ALL, 5)

        self.m_button1 = wx.Button(self, wx.ID_ANY, u"确认", wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer1.Add(self.m_button1, 0, wx.ALL, 5)

        self.m_button2 = wx.Button(self, wx.ID_ANY, u"退出", wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer1.Add(self.m_button2, 0, wx.ALL, 5)

        self.SetSizer(gSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.m_button1.Bind(wx.EVT_BUTTON, self.onConfirm)
        self.m_button2.Bind(wx.EVT_BUTTON, self.onCancle)

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def onConfirm(self, event):
        event.Skip()

    def onCancle(self, event):
        event.Skip()


