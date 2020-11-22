#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.uno import Exception as UnoException
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .pagemodel import PageModel
from .pageview import PageView
from .dialogview import DialogView
from .pagehandler import DialogHandler

from unolib import createService
from unolib import getDialog

from .logger import setDebugMode
from .logger import logMessage
from .logger import getMessage
g_message = 'pagemanager'


import traceback


class PageManager(unohelper.Base):
    def __init__(self, ctx, wizard, model=None):
        self.ctx = ctx
        self._search = True
        self._connected = False
        self._wizard = wizard
        self._model = PageModel(self.ctx) if model is None else model
        self._views = {}
        self._dialog = None
        print("PageManager.__init__() %s" % self._model.Email)

    @property
    def View(self):
        pageid = self.Wizard.getCurrentPage().PageId
        return self.getView(pageid)
    @property
    def Model(self):
        return self._model
    @property
    def Wizard(self):
        return self._wizard

    def getView(self, pageid):
        if pageid in self._views:
            return self._views[pageid]
        print("PageManager.getView ERROR **************************")
        return None

    def getWizard(self):
        return self.Wizard

    def initPage(self, pageid, window):
        view = PageView(self.ctx, window)
        self._views[pageid] = view
        if pageid == 1:
            view.initPage1(self.Model)

    def activatePage(self, pageid):
        if pageid == 2:
            self.Wizard.enablePage(1, False)
            self.getView(pageid).activatePage2(self.Model)
            self.Model.getSmtpConfig(self.updatePage2, self.updateModel)
        elif pageid == 3:
            self.getView(pageid).activatePage3(self.Model)
        elif pageid == 4:
            self._connected = False
            setDebugMode(self.ctx, True)
            self.getView(pageid).activatePage4(self.Model)
            context = self.getView(3).getConnectionContext(self.Model)
            authenticator = self.getView(3).getAuthenticator()
            self.Model.smtpConnect(context, authenticator, self.updatePage4, self.callBack)

    def updatePage2(self, value, offset=0):
        if self.Wizard.DialogWindow is not None:
            self.getView(2).updatePage2(self.Model, value, offset)

    def updatePage4(self, value, offset=0, msg=None):
        #message = getMessage(self.ctx, g_message, value + offset)
        #if msg is not None:
        #    message = message % msg
        #logMessage(self.ctx, INFO, message, 'PageManager', 'updatePage4()')
        if self.Wizard.DialogWindow is not None:
            self.getView(4).updatePage4(self.Model, value)

    def updateModel(self, user, servers, offline):
        print("PageManager.updateModel() 1: %s" % self.Wizard.getCurrentPage().PageId)
        if self.Wizard.DialogWindow is not None:
            self.Model.User = user
            self.Model.Servers = servers
            self.Model.Online = not offline
            self._search = False
            print("PageManager.updateModel() 2")
            self.Wizard.enablePage(1, True)
            self.Wizard.updateTravelUI()
            if len(servers) > 0:
                print("PageManager.updateModel() 3")
                #self.Wizard._manager._setCurrentPage(3)
                #self.Wizard.travelNext()
                print("PageManager.updateModel() 4")

    def updateDialog(self, value, offset=0, msg=None):
        if self._dialog is not None:
            self._dialog.updateProgress(value, offset, msg)

    def callBack(self, state):
        if self.Wizard.DialogWindow is not None:
            self._connected = state
            self.Wizard.updateTravelUI()

    def canAdvancePage(self, pageid):
        advance = False
        if pageid == 1:
            advance = self._isUserValid(pageid)
        elif pageid == 2:
            advance = not self._search
        elif pageid == 3:
            advance = all((self._isServerValid(pageid),
                           self._isPortValid(pageid),
                           self._isLoginNameValid(pageid),
                           self._isPasswordValid(pageid)))
        elif pageid == 4:
            advance = self._connected
        return advance

    def commitPage(self, pageid, reason):
        if pageid == 1:
            self.Model.Email = self.View.getUser()
        elif pageid == 2:
            self._search = True
        elif pageid == 4:
            setDebugMode(self.ctx, False)
        return True

    def setPageTitle(self, pageid):
        self.Wizard.setTitle(self.getView(pageid).getPageTitle(self.Model, pageid))

    def updateTravelUI(self):
        self.Wizard.updateTravelUI()

    def changeConnection(self, control):
        index = self.View.getControlIndex(control)
        self.View.setConnectionSecurity(self.Model, index)

    def changeAuthentication(self, control):
        index = self.View.getControlIndex(control)
        if index == 0:
            self.View.enableLoginName(False)
            self.View.enablePassword(False)
        elif index == 3:
            self.View.enableLoginName(True)
            self.View.enablePassword(False)
        else:
            self.View.enableLoginName(True)
            self.View.enablePassword(True)
        self.View.setAuthenticationSecurity(self.Model, index)
        self.updateTravelUI()

    def previousServerPage(self):
        self.Model.previousServerPage()
        self.View.updatePage3(self.Model)

    def nextServerPage(self):
        self.Model.nextServerPage()
        self.View.updatePage3(self.Model)

    def showSmtpConnect(self):
        context = self.View.getConnectionContext(self.Model)
        authenticator = self.View.getAuthenticator()
        handler = DialogHandler(self)
        parent = self.Wizard.DialogWindow.getPeer()
        self._dialog = DialogView(self.ctx, 'SmtpDialog', handler, parent)
        self._dialog.setTitle(context)
        self.Model.smtpConnect(server, context, authenticator, self.updateDialog, self.connectServer, self.callBackDialog)
        if self._dialog.execute() == OK:
            print("PageManager.showSmtpConnect() OK")
        else:
            print("PageManager.showSmtpConnect() CANCEL")
        self._dialog.dispose()
        self._dialog = None

    def _isUserValid(self, pageid):
        return self.getView(pageid).isUserValid(self.Model.isUserValid)

    def _isServerValid(self, pageid):
        return self.getView(pageid).isServerValid(self.Model.isServerValid)

    def _isPortValid(self, pageid):
        return self.getView(pageid).isPortValid(self.Model.isPortValid)

    def _isLoginNameValid(self, pageid):
        return self.getView(pageid).isLoginNameValid(self.Model.isStringValid)

    def _isPasswordValid(self, pageid):
        return self.getView(pageid).isPasswordValid(self.Model.isStringValid)
