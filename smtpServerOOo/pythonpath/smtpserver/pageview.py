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
        self._getEmail(window).Text = model.Email

    def activatePage2(self, window, model):
        self._setPageTitle(2, window, model, model.Email)
        self.updateProgress(window, model, 0)

    def updateProgress(self, window, model, value, offset=0):
        self._getProgressBar(window).Value = value
        text = model.resolveString(self._getProgressMessage(value + offset))
        self._getProgressLabel(window).Text = text

    def activatePage3(self, window, model):
        self._setPageTitle(3, window, model, model.Email)
        self.updatePage3(window, model)

    def updatePage3(self, window, model):
        self._enablePrevious(window, model.isFirst())
        self._enableNext(window, model.isLast())
        self._getServerPage(window).Text = model.getServerPage()
        self._loadPage3(window, model)

    def enableUserName(self, window, enabled):
        self._getLoginNameLabel(window).Model.Enabled = enabled
        self._getLoginName(window).Model.Enabled = enabled

    def enablePassword(self, window, enabled):
        self._getPasswordLabel(window).Model.Enabled = enabled
        self._getPassword(window).Model.Enabled = enabled
        self._getConfirmPwdLabel(window).Model.Enabled = enabled
        self._getConfirmPwd(window).Model.Enabled = enabled

# PageView getter methods
    def getControlTag(self, control):
        return control.Model.Tag

    def getControlIndex(self, control):
        return control.getSelectedItemPos()

    def isServerValid(self, window, validator):
        return validator(self._getServer(window).Text)

    def isPortValid(self, window, validator):
        return validator(int(self._getPort(window).Value))

    def isLoginNameValid(self, window, validator):
        control = self._getLoginName(window)
        if control.Model.Enabled:
            return validator(control.Text)
        return True

    def isPasswordValid(self, window, validator):
        control = self._getPassword(window)
        if control.Model.Enabled:
            if self._getConfirmPwd(window).Text != control.Text:
                return False
            return validator(control.Text)
        return True

# PageView private methods
    def _setPageTitle(self, pageid, window, model, title):
        text = model.resolveString(self._getPageTitleMessage(pageid))
        self._getPageTitle(window).Text = text % title

    def _loadPage3(self, window, model):
        self._getServer(window).Text = model.getServer()
        self._getPort(window).Text = model.getPort()
        self._getConnection(window).selectItemPos(model.getConnection(), True)
        self._getAuthentication(window).selectItemPos(model.getAuthentication(), True)
        self._getLoginName(window).Text = model.getLoginName()
        self._getPassword(window).Text = model.getPassword()
        self._getConfirmPwd(window).Text = model.getPassword()

    def _enablePrevious(self, window, isfirst):
        self._getPrevious(window).Model.Enabled = not isfirst

    def _enableNext(self, window, islast):
        self._getNext(window).Model.Enabled = not islast

# PageView private message methods
    def _getPageTitleMessage(self, pageid):
        return 'PageWizard%s.Label1.Label' % pageid

    def _getProgressMessage(self, value):
        return 'PageWizard2.Label2.Label.%s' % value

# PageView private control methods
    def _getPageTitle(self, window):
        return window.getControl('Label1')

    def _getEmail(self, window):
        return window.getControl('TextField1')

    def _getProgressBar(self, window):
        return window.getControl('ProgressBar1')

    def _getProgressLabel(self, window):
        return window.getControl('Label2')

    def _getServerPage(self, window):
        return window.getControl('Label2')

    def _getServer(self, window):
        return window.getControl('TextField1')

    def _getPort(self, window):
        return window.getControl('NumericField1')

    def _getConnection(self, window):
        return window.getControl('ListBox1')

    def _getAuthentication(self, window):
        return window.getControl('ListBox2')

    def _getLoginNameLabel(self, window):
        return window.getControl('Label9')

    def _getLoginName(self, window):
        return window.getControl('TextField2')

    def _getPasswordLabel(self, window):
        return window.getControl('Label10')

    def _getPassword(self, window):
        return window.getControl('TextField3')

    def _getConfirmPwdLabel(self, window):
        return window.getControl('Label11')

    def _getConfirmPwd(self, window):
        return window.getControl('TextField4')

    def _getPrevious(self, window):
        return window.getControl('CommandButton1')

    def _getNext(self, window):
        return window.getControl('CommandButton2')
