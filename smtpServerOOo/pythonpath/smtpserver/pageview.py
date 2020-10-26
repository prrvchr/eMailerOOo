#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from .logger import logMessage

import traceback


class PageView(unohelper.Base):
    def __init__(self, ctx):
        self.ctx = ctx
        print("PageView.__init__()")

# PageView setter methods
    def initPage1(self, window, model):
        window.getControl('TextField1').Text = model.Email

    def activatePage2(self, window, model):
        self._setPageTitle(2, window, model, model.Email)
        self.updateProgress(window, model, 0)

    def updateProgress(self, window, model, value, offset=0):
        window.getControl('ProgressBar1').Value = value
        value += offset
        text = model.resolveString('PageWizard2.Label2.Label.%s' % value)
        window.getControl('Label2').Text = text

    def activatePage3(self, window, model):
        self._setPageTitle(3, window, model, model.Email)
        self.updatePage3(window, model)
        print("PageView.activatePage() 1")

    def updatePage3(self, window, model):
        print("PageView.updatePage3() 1")
        self._enablePrevious(window, model.isFirst())
        self._enableNext(window, model.isLast())
        mdl, text = model.getServerPage()
        window.getControl('Label2').Text = text
        self._loadPage3(window, mdl)
        print("PageView.updatePage3() 2")

    def enableUserName(self, window, enabled):
        window.getControl('Label9').Model.Enabled = enabled
        window.getControl('TextField2').Model.Enabled = enabled

    def enablePassword(self, window, enabled):
        window.getControl('Label10').Model.Enabled = enabled
        window.getControl('TextField3').Model.Enabled = enabled
        window.getControl('Label11').Model.Enabled = enabled
        window.getControl('TextField4').Model.Enabled = enabled

# PageView getter methods
    def getControlTag(self, control):
        return control.Model.Tag

    def getControlIndex(self, control):
        return control.getSelectedItemPos()


# PageView private methods
    def _setPageTitle(self, pageid, window, model, title):
        text = model.resolveString('PageWizard%s.Label1.Label' % pageid)
        window.getControl('Label1').Text = text % title

    def _loadPage3(self, window, model=None):
        server = '' if model is None else model.getServer()
        window.getControl('TextField1').Text = server
        port = '' if model is None else model.getPort()
        window.getControl('NumericField1').Text = port
        connection = 0 if model is None else model.getConnection()
        window.getControl('ListBox1').selectItemPos(connection, True)
        authentication = 0 if model is None else model.getAuthentication()
        window.getControl('ListBox2').selectItemPos(authentication, True)
        username = '' if model is None else model.getUserName()
        window.getControl('TextField2').Text = username

    def _enablePrevious(self, window, isfirst):
        window.getControl('CommandButton1').Model.Enabled = not isfirst

    def _enableNext(self, window, islast):
        window.getControl('CommandButton2').Model.Enabled = not islast

