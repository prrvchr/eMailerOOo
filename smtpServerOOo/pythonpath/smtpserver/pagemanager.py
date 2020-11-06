#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.mail.MailServiceType import SMTP

from .pagemodel import PageModel
from .pageview import PageView
from .dialogview import DialogView
from .pagehandler import DialogHandler

from unolib import createService

from .logger import logMessage

import traceback


class PageManager(unohelper.Base):
    def __init__(self, ctx, wizard, model=None):
        self.ctx = ctx
        self._search = True
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
        self._views[pageid] = PageView(window)
        if pageid == 1:
            self.getView(pageid).initPage1(self.Model)

    def activatePage2(self):
        self.Wizard.enablePage(1, False)
        self.View.activatePage2(self.Model)
        self.Model.getSmtpConfig(self.updateProgress, self.updateModel)

    def activatePage3(self):
        self.View.activatePage3(self.Model)

    def updateProgress(self, value, offset=0):
        if self.Wizard.DialogWindow is not None:
            self.View.updateProgress(self.Model, value, offset)

    def updateModel(self, user, servers, offline):
        if self.Wizard.DialogWindow is not None:
            self.Model.User = user
            self.Model.Servers = servers
            self.Model.Online = not offline
            self._search = False
            self.Wizard.updateTravelUI()
            if len(servers) > 0:
                self.Wizard.travelNext()

    def updateDialog(self, value, offset=0, msg=None):
        if self._dialog is not None:
            self._dialog.updateProgress(value, offset, msg)

    def callBackDialog(self, state):
        if self._dialog is not None:
            self._dialog.callBack(state)

    def canAdvancePage(self, pageid):
        advance = False
        if pageid == 1:
            advance = self._isUserValid()
        elif pageid == 2:
            advance = not self._search
        elif pageid == 3:
            advance = all((self._isServerValid(),
                           self._isPortValid(),
                           self._isLoginNameValid(),
                           self._isPasswordValid()))
            self.getView(pageid).enableConnect(advance and self.Model.Online)
        return advance

    def commitPage1(self):
        self.Model.Email = self.View.getUser()

    def commitPage2(self):
        print("PageManager.commitPage2()")
        self.Wizard.enablePage(1, True)
        self._search = True

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
        self.Model.smtpConnect(context, authenticator, self.updateDialog, self.callBackDialog)
        if self._dialog.execute() == OK:
            print("PageManager.showSmtpConnect() OK")
        else:
            print("PageManager.showSmtpConnect() CANCEL")
        self._dialog.dispose()
        self._dialog = None

    def smtpConnect(self):
        self._dialog.enableButtonOk(False)
        self._dialog.enableButtonRetry(False)
        context = self.View.getConnectionContext(self.Model)
        authenticator = self.View.getAuthenticator()
        self.Model.smtpConnect(context, authenticator, self.updateDialog, self.callBackDialog)

    def _isUserValid(self):
        return self.View.isUserValid(self.Model.isUserValid)

    def _isServerValid(self):
        return self.View.isServerValid(self.Model.isServerValid)

    def _isPortValid(self):
        return self.View.isPortValid(self.Model.isPortValid)

    def _isLoginNameValid(self):
        return self.View.isLoginNameValid(self.Model.isStringValid)

    def _isPasswordValid(self):
        return self.View.isPasswordValid(self.Model.isStringValid)
