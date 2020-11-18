#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.uno import XCurrentContext
from com.sun.star.mail import XAuthenticator

from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontWeight import NORMAL

from .logger import logMessage

import traceback


class PageView(unohelper.Base):
    def __init__(self, window):
        self.Window = window
        secure = {0: 3, 1: 4, 2: 4, 3: 5}
        unsecure = {0: 0, 1: 1, 2: 2, 3: 2}
        self._levels = {0: unsecure, 1: secure, 2: secure}
        print("PageView.__init__()")

# PageView setter methods
    def initPage1(self, model):
        self._getUser().Text = model.Email

    def activatePage2(self, model):
        self._setPageTitle(2, model, model.Email)
        self.updateProgress(model, 5)

    def updateProgress(self, model, value, offset=0):
        self._getProgressBar().Value = value
        text = model.resolveString(self._getProgressMessage(value + offset))
        self._getProgressLabel().Text = text

    def activatePage3(self, model):
        self._setPageTitle(3, model, model.Email)
        self._setConnectionMode(model.Online)
        self.updatePage3(model)

    def updatePage3(self, model):
        self._enablePrevious(model.isFirst())
        self._enableNext(model.isLast())
        self._getServerPage().Text = model.getServerPage()
        self._loadPage3(model)

    def enableLoginName(self, enabled):
        self._getLoginNameLabel().Model.Enabled = enabled
        self._getLoginName().Model.Enabled = enabled

    def enablePassword(self, enabled):
        self._getPasswordLabel().Model.Enabled = enabled
        self._getPassword().Model.Enabled = enabled
        self._getConfirmPwdLabel().Model.Enabled = enabled
        self._getConfirmPwd().Model.Enabled = enabled

    def enableConnect(self, enabled):
        self._getConnect().Model.Enabled = enabled

    def setConnectionSecurity(self, model, index):
        level = self._getConnectionLevel(index)
        self._setSecurityMessage(model, level)

    def setAuthenticationSecurity(self, model, index):
        level = self._getAuthenticationLevel(index)
        self._setSecurityMessage(model, level)

# PageView getter methods
    def getControlTag(self, control):
        return control.Model.Tag

    def getControlIndex(self, control):
        return control.getSelectedItemPos()

    def getPageTitle(self, model, pageid):
        return model.resolveString(self._getPageTitle(pageid))

    def isUserValid(self, validator):
        return validator(self._getUser().Text)

    def isServerValid(self, validator):
        return validator(self._getServer().Text)

    def isPortValid(self, validator):
        return validator(self._getPort().Value)

    def isLoginNameValid(self, validator):
        control = self._getLoginName()
        if control.Model.Enabled:
            return validator(control.Text)
        return True

    def isPasswordValid(self, validator):
        control = self._getPassword()
        if control.Model.Enabled:
            if self._getConfirmPwd().Text != control.Text:
                return False
            return validator(control.Text)
        return True

    def getUser(self):
        return self._getUser().Text

    def getConnectionContext(self, model):
        index = {0: 'Insecure', 1: 'Ssl', 2: 'Tls'}
        connection = index.get(self.getControlIndex(self._getConnection()))
        index = {0: 'None', 1: 'Login', 2: 'Login', 3: 'OAuth2'}
        authentication = index.get(self.getControlIndex(self._getAuthentication()))
        data = {'ServerName': self._getServer().Text,
                'Port': int(self._getPort().Value),
                'ConnectionType': connection,
                'AuthenticationType': authentication,
                'Timeout': model.Timeout}
        return CurrentContext(data)

    def getAuthenticator(self):
        user = self._getLoginName().Text
        password = self._getPassword().Text
        return Authenticator(user, password)

# PageView private methods
    def _setPageTitle(self, pageid, model, title):
        text = model.resolveString(self._getPageLabelMessage(pageid))
        self._getPageLabel().Text = text % title

    def _setConnectionMode(self, enabled):
        self._getConnectLabel().Model.Enabled = enabled

    def _loadPage3(self, model):
        self._getServer().Text = model.getServer()
        self._getPort().Value = model.getPort()
        self._getConnection().selectItemPos(model.getConnection(), True)
        self._getAuthentication().selectItemPos(model.getAuthentication(), True)
        self._getLoginName().Text = model.getLoginName()
        self._getPassword().Text = model.getPassword()
        self._getConfirmPwd().Text = model.getPassword()

    def _enablePrevious(self, isfirst):
        self._getPrevious().Model.Enabled = not isfirst

    def _enableNext(self, islast):
        self._getNext().Model.Enabled = not islast

    def _getConnectionLevel(self, index):
        i = self.getControlIndex(self._getAuthentication())
        return self._levels.get(index).get(i)

    def _getAuthenticationLevel(self, index):
        i = self.getControlIndex(self._getConnection())
        return self._levels.get(i).get(index)

    def _setSecurityMessage(self, model, level):
        control = self._getSecurityLabel()
        control.Model.FontWeight = BOLD if level < 3 else NORMAL
        control.Text = model.resolveString(self._getSecurityMessage(level))

# PageView private message methods
    def _getPageTitle(self, pageid):
        return 'PageWizard%s.Title' % pageid

    def _getPageLabelMessage(self, pageid):
        return 'PageWizard%s.Label1.Label' % pageid

    def _getProgressMessage(self, value):
        return 'PageWizard2.Label2.Label.%s' % value

    def _getSecurityMessage(self, level):
        return 'PageWizard3.Label10.Label.%s' % level

# PageView private control methods
    def _getPageLabel(self):
        return self.Window.getControl('Label1')

    def _getUser(self):
        return self.Window.getControl('TextField1')

    def _getProgressBar(self):
        return self.Window.getControl('ProgressBar1')

    def _getProgressLabel(self):
        return self.Window.getControl('Label2')

    def _getServerPage(self):
        return self.Window.getControl('Label2')

    def _getServer(self):
        return self.Window.getControl('TextField1')

    def _getPort(self):
        return self.Window.getControl('NumericField1')

    def _getConnection(self):
        return self.Window.getControl('ListBox1')

    def _getAuthentication(self):
        return self.Window.getControl('ListBox2')

    def _getLoginNameLabel(self):
        return self.Window.getControl('Label7')

    def _getLoginName(self):
        return self.Window.getControl('TextField2')

    def _getPasswordLabel(self):
        return self.Window.getControl('Label8')

    def _getPassword(self):
        return self.Window.getControl('TextField3')

    def _getConfirmPwdLabel(self):
        return self.Window.getControl('Label9')

    def _getConfirmPwd(self):
        return self.Window.getControl('TextField4')

    def _getPrevious(self):
        return self.Window.getControl('CommandButton1')

    def _getNext(self):
        return self.Window.getControl('CommandButton2')

    def _getSecurityLabel(self):
        return self.Window.getControl('Label10')

    def _getConnectLabel(self):
        return self.Window.getControl('Label11')

    def _getConnect(self):
        return self.Window.getControl('CommandButton3')


class CurrentContext(unohelper.Base,
                     XCurrentContext):
    def __init__(self, data):
        self._data = data

    def getValueByName(self, name):
        return self._data[name]


class Authenticator(unohelper.Base,
                    XAuthenticator):
    def __init__(self, user, password):
        self._user = user
        self._password = password

    def getUserName(self):
        return self._user

    def getPassword(self):
        return self._password
