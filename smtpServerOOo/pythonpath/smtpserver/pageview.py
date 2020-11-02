#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontWeight import NORMAL

from unolib import createService

from .logger import logMessage

import traceback


class PageView(unohelper.Base):
    def __init__(self, ctx):
        self.ctx = ctx
        secure = {0: 3, 1: 4, 2: 4, 3: 5}
        unsecure = {0: 0, 1: 1, 2: 2, 3: 2}
        self._levels = {0: unsecure, 1: secure, 2: secure}
        print("PageView.__init__()")

# PageView setter methods
    def initPage1(self, window, model):
        self._getUser(window).Text = model.Email

    def activatePage2(self, window, model):
        self._setPageTitle(2, window, model, model.Email)
        self.updateProgress(window, model, 0)

    def updateProgress(self, window, model, value, offset=0):
        self._getProgressBar(window).Value = value
        text = model.resolveString(self._getProgressMessage(value + offset))
        self._getProgressLabel(window).Text = text

    def activatePage3(self, window, model):
        self._setPageTitle(3, window, model, model.Email)
        self._setConnectionMode(window, model.Online)
        self.updatePage3(window, model)

    def updatePage3(self, window, model):
        self._enablePrevious(window, model.isFirst())
        self._enableNext(window, model.isLast())
        self._getServerPage(window).Text = model.getServerPage()
        self._loadPage3(window, model)

    def enableLoginName(self, window, enabled):
        self._getLoginNameLabel(window).Model.Enabled = enabled
        self._getLoginName(window).Model.Enabled = enabled

    def enablePassword(self, window, enabled):
        self._getPasswordLabel(window).Model.Enabled = enabled
        self._getPassword(window).Model.Enabled = enabled
        self._getConfirmPwdLabel(window).Model.Enabled = enabled
        self._getConfirmPwd(window).Model.Enabled = enabled

    def enableConnect(self, window, enabled):
        self._getConnect(window).Model.Enabled = enabled

    def setConnectionSecurity(self, window, model, index):
        level = self._getConnectionLevel(window, index)
        self._setSecurityLevel(window, model, level)

    def setAuthenticationSecurity(self, window, model, index):
        level = self._getAuthenticationLevel(window, index)
        self._setSecurityLevel(window, model, level)

# PageView getter methods
    def getControlTag(self, control):
        return control.Model.Tag

    def getControlIndex(self, control):
        return control.getSelectedItemPos()

    def getPageStep(self, model, pageid):
        return model.resolveString(self._getRoadmapStep(pageid))

    def getPageTitle(self, model, pageid):
        return model.resolveString(self._getPageTitle(pageid))

    def isUserValid(self, window, validator):
        return validator(self._getUser(window).Text)

    def isServerValid(self, window, validator):
        return validator(self._getServer(window).Text)

    def isPortValid(self, window, validator):
        return validator(self._getPort(window).Value)

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

    def getUser(self, window):
        return self._getUser(window).Text

# PageView private methods
    def _setPageTitle(self, pageid, window, model, title):
        text = model.resolveString(self._getPageLabelMessage(pageid))
        self._getPageLabel(window).Text = text % title

    def _setConnectionMode(self, window, enabled):
        self._getConnectLabel(window).Model.Enabled = enabled

    def _loadPage3(self, window, model):
        self._getServer(window).Text = model.getServer()
        self._getPort(window).Value = model.getPort()
        self._getConnection(window).selectItemPos(model.getConnection(), True)
        self._getAuthentication(window).selectItemPos(model.getAuthentication(), True)
        self._getLoginName(window).Text = model.getLoginName()
        self._getPassword(window).Text = model.getPassword()
        self._getConfirmPwd(window).Text = model.getPassword()

    def _enablePrevious(self, window, isfirst):
        self._getPrevious(window).Model.Enabled = not isfirst

    def _enableNext(self, window, islast):
        self._getNext(window).Model.Enabled = not islast

    def _getConnectionLevel(self, window, index):
        i = self.getControlIndex(self._getAuthentication(window))
        return self._levels.get(index).get(i)

    def _getAuthenticationLevel(self, window, index):
        i = self.getControlIndex(self._getConnection(window))
        return self._levels.get(i).get(index)

    def _setSecurityLevel(self, window, model, level):
        control = self._getSecurityLabel(window)
        control.Model.FontWeight = BOLD if level < 3 else NORMAL
        control.Text = model.resolveString(self._getSecurityMessage(level))

# PageView private message methods
    def _getRoadmapStep(self, pageid):
        return 'PageWizard%s.Step' % pageid

    def _getPageTitle(self, pageid):
        return 'PageWizard%s.Title' % pageid

    def _getPageLabelMessage(self, pageid):
        return 'PageWizard%s.Label1.Label' % pageid

    def _getProgressMessage(self, value):
        return 'PageWizard2.Label2.Label.%s' % value

    def _getSecurityMessage(self, level):
        return 'PageWizard3.Label10.Label.%s' % level

# PageView private control methods
    def _getPageLabel(self, window):
        return window.getControl('Label1')

    def _getUser(self, window):
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
        return window.getControl('Label7')

    def _getLoginName(self, window):
        return window.getControl('TextField2')

    def _getPasswordLabel(self, window):
        return window.getControl('Label8')

    def _getPassword(self, window):
        return window.getControl('TextField3')

    def _getConfirmPwdLabel(self, window):
        return window.getControl('Label9')

    def _getConfirmPwd(self, window):
        return window.getControl('TextField4')

    def _getPrevious(self, window):
        return window.getControl('CommandButton1')

    def _getNext(self, window):
        return window.getControl('CommandButton2')

    def _getSecurityLabel(self, window):
        return window.getControl('Label10')

    def _getConnectLabel(self, window):
        return window.getControl('Label11')

    def _getConnect(self, window):
        return window.getControl('CommandButton3')

