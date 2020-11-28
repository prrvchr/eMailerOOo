#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.uno import XCurrentContext
from com.sun.star.mail import XAuthenticator

from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontWeight import NORMAL

from unolib import KeyMap
from unolib import getFileSequence

from .logger import clearLogger
from .logger import getLoggerUrl
from .logger import logMessage

import traceback


class PageView(unohelper.Base):
    def __init__(self, ctx, window):
        self.ctx = ctx
        self.Window = window
        secure = {0: 3, 1: 4, 2: 4, 3: 5}
        unsecure = {0: 0, 1: 1, 2: 2, 3: 2}
        self._levels = {0: unsecure, 1: secure, 2: secure}
        print("PageView.__init__()")

# PageView setter methods
    def initPage1(self, model):
        self._setEmail(model.Email)

    def activatePage2(self, model):
        self._setPageLabel(2, model, model.Email)

    def activatePage3(self, model):
        self._setPageLabel(3, model, model.Email)
        self.updatePage3(model)

    def activatePage4(self, model):
        self._setPageLabel(4, model, (model.getHost(), model.getPort()))
        self.setPage4Step(1)

    def updatePage2(self, model, value, offset=0):
        self._getProgressBar().Value = value
        text = model.resolveString(self._getPage2Message(value + offset))
        self._getProgressLabel().Text = text

    def updatePage3(self, model):
        self._enableNavigation(model)
        self._getServerPage().Text = model.getServerPage()
        self._loadPage3(model)

    def updatePage4(self, model, value):
        self._getProgressBar().Value = value
        if value == 0:
            clearLogger()
        self._updateLogger()

    def setPage4Step(self, step):
        self.Window.Model.Step = step

    def enableLogin(self, enabled):
        self._getLoginLabel().Model.Enabled = enabled
        self._getLogin().Model.Enabled = enabled

    def enablePassword(self, enabled):
        self._getPasswordLabel().Model.Enabled = enabled
        self._getPassword().Model.Enabled = enabled
        self._getConfirmPwdLabel().Model.Enabled = enabled
        self._getConfirmPwd().Model.Enabled = enabled

    def setConnectionSecurity(self, model, index):
        level = self._getConnectionLevel(index)
        self._setSecurityMessage(model, level)

    def setAuthenticationSecurity(self, model, index):
        level = self._getAuthenticationLevel(index)
        self._setSecurityMessage(model, level)

# PageView getter methods
    def getControlIndex(self, control):
        return control.getSelectedItemPos()

    def getPageTitle(self, model, pageid, offline):
        return model.resolveString(self._getPageTitle(pageid, offline))

    def isEmailValid(self, validator):
        return validator(self._getEmail())

    def isHostValid(self, validator):
        return validator(self._getHost())

    def isPortValid(self, validator):
        return validator(self._getPort())

    def isLoginValid(self, validator):
        control = self._getLogin()
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

    def getEmail(self):
        return self._getEmail()

    def getConnectionContext(self, model):
        connection, authentication = self._getConnectionMode(model)
        data = {'ServerName': self._getHost(),
                'Port': self._getPort(),
                'ConnectionType': connection,
                'AuthenticationType': authentication,
                'Timeout': model.Timeout}
        return CurrentContext(data)

    def getAuthenticator(self):
        user = self._getLogin().Text
        password = self._getPassword().Text
        return Authenticator(user, password)

    def getConfiguration(self, model):
        server = self.getServer()
        server.setValue('LoginMode', model.getLoginMode())
        user = KeyMap()
        user.setValue('Server', server.getValue('Server'))
        user.setValue('Port',  server.getValue('Port'))
        user.setValue('LoginName', self._getLogin().Text)
        user.setValue('Password', self._getPassword().Text)
        user.setValue('Domain', model.User.getValue('Domain'))
        return user, server

    def getServer(self):
        server = KeyMap()
        connection, authentication = self._getConnectionIndex()
        server.setValue('Server', self._getHost())
        server.setValue('Port', self._getPort())
        server.setValue('Connection', connection)
        server.setValue('Authentication', authentication)
        return server

# PageView private methods
    def _setPageLabel(self, pageid, model, format):
        text = model.resolveString(self._getPageLabelMessage(pageid))
        self._getPageLabel().Text = text % format

    def _loadPage3(self, model):
        self._setHost(model.getHost())
        self._setPort(model.getPort())
        self._setConnection(model.getConnection())
        self._setAuthentication(model.getAuthentication())
        self._getLogin().Text = model.getLoginName()
        self._getPassword().Text = model.getPassword()
        self._getConfirmPwd().Text = model.getPassword()

    def _enableNavigation(self, model):
        self._enablePrevious(model.isFirst())
        self._enableNext(model.isLast())

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

    def _updateLogger(self):
        url = getLoggerUrl(self.ctx)
        length, sequence = getFileSequence(self.ctx, url)
        control = self._getLogger()
        control.Text = sequence.value.decode('utf-8')
        selection = uno.createUnoStruct('com.sun.star.awt.Selection', length, length)
        control.setSelection(selection)

    def _getConnectionIndex(self):
        connection = self.getControlIndex(self._getConnection())
        authentication = self.getControlIndex(self._getAuthentication())
        return connection, authentication

    def _getConnectionMode(self, model):
        i, j = self._getConnectionIndex()
        connection = model.getConnectionMap(i)
        authentication = model.getAuthenticationMap(j)
        return connection, authentication

# PageView private message methods
    def _getPageTitle(self, pageid, offline):
        return 'PageWizard%s.Title.%s' % (pageid, offline)

    def _getPageLabelMessage(self, pageid):
        return 'PageWizard%s.Label1.Label' % pageid

    def _getPage2Message(self, value):
        return 'PageWizard2.Label2.Label.%s' % value

    def _getSecurityMessage(self, level):
        return 'PageWizard3.Label10.Label.%s' % level

# PageView private getter control methods
    def _getPageLabel(self):
        return self.Window.getControl('Label1')

    def _getEmail(self):
        return self.Window.getControl('TextField1').Text

    def _getProgressBar(self):
        return self.Window.getControl('ProgressBar1')

    def _getProgressLabel(self):
        return self.Window.getControl('Label2')

    def _getLogger(self):
        return self.Window.getControl('TextField1')

    def _getServerPage(self):
        return self.Window.getControl('Label2')

    def _getHost(self):
        return self.Window.getControl('TextField1').Text

    def _getPort(self):
        return int(self.Window.getControl('NumericField1').Value)

    def _getConnection(self):
        return self.Window.getControl('ListBox1')

    def _getAuthentication(self):
        return self.Window.getControl('ListBox2')

    def _getLoginLabel(self):
        return self.Window.getControl('Label7')

    def _getLogin(self):
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

# PageView private setter control methods
    def _setEmail(self, text):
        self.Window.getControl('TextField1').Text = text

    def _setHost(self, text):
        self.Window.getControl('TextField1').Text = text

    def _setPort(self, value):
        self.Window.getControl('NumericField1').Value = value

    def _setConnection(self, index):
        self.Window.getControl('ListBox1').selectItemPos(index, True)

    def _setAuthentication(self, index):
        self.Window.getControl('ListBox2').selectItemPos(index, True)


class CurrentContext(unohelper.Base,
                     XCurrentContext):
    def __init__(self, context):
        self._context = context

    def getValueByName(self, name):
        return self._context[name]


class Authenticator(unohelper.Base,
                    XAuthenticator):
    def __init__(self, user, password):
        self._user = user
        self._password = password

    def getUserName(self):
        return self._user

    def getPassword(self):
        return self._password
