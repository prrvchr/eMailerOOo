#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from .pagemodel import PageModel
from .pageview import PageView

from .logger import logMessage

import traceback


class PageManager(unohelper.Base):
    def __init__(self, ctx, wizard, model=None):
        self.ctx = ctx
        self._search = True
        self._wizard = wizard
        self._model = PageModel(self.ctx) if model is None else model
        self._view = PageView(self.ctx)
        print("PageManager.__init__() %s" % self._model.Email)

    def getWizard(self):
        return self._wizard

    def initPage(self, pageid, window):
        if pageid == 1:
            self._view.initPage1(window, self._model)

    def activatePage2(self, window, progress):
        self._view.activatePage2(window, self._model)
        self._model.getSmtpConfig(progress, self.updateModel)
        self._search = True

    def activatePage3(self, window):
        self._view.activatePage3(window, self._model)

    def updateProgress(self, window, value, offset=0):
        self._view.updateProgress(window, self._model, value, offset)

    def updateModel(self, user, servers):
        self._model.User = user
        self._model.Servers = servers
        self._search = False
        self._wizard.updateTravelUI()
        if len(servers) > 0:
            self._wizard.travelNext()

    def canAdvancePage(self, pageid, window):
        advance = False
        if pageid == 1:
            advance = self._model.isEmailValid()
        elif pageid == 2:
            advance = not self._search
        elif pageid == 3:
            advance = all((self._isServerValid(window),
                           self._isPortValid(window),
                           self._isLoginNameValid(window),
                           self._isPasswordValid(window)))
            self._view.enableConnect(window, advance)
        return advance

    def commitPage2(self):
        self._search = True

    def getPageStep(self, pageid):
        return self._view.getPageStep(self._model, pageid)

    def setPageTitle(self, pageid):
        self._wizard.setTitle(self._view.getPageTitle(self._model, pageid))

    def updateTravelUI(self, window, control):
        try:
            tag = self._view.getControlTag(control)
            if tag == 'EmailAddress':
                self._model.Email = control.Text
            self._wizard.updateTravelUI()
        except Exception as e:
            print("PageManager.updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))

    def changeAuthentication(self, window, control):
        index = self._view.getControlIndex(control)
        if index == 0:
            self._view.enableUserName(window, False)
            self._view.enablePassword(window, False)
        elif index == 3:
            self._view.enableUserName(window, True)
            self._view.enablePassword(window, False)
        else:
            self._view.enableUserName(window, True)
            self._view.enablePassword(window, True)
        self._wizard.updateTravelUI()
        print("PageManager.changeAuthentication()")

    def previousServerPage(self, window):
        self._model.previousServerPage()
        self._view.updatePage3(window, self._model)
        print("PageManager.previousServerPage()")

    def nextServerPage(self, window):
        self._model.nextServerPage()
        self._view.updatePage3(window, self._model)
        print("PageManager.nextServerPage()")

    def smtpConnect(self, window, event):
        print("PageManager._smtpConnect()")

    def _isServerValid(self, window):
        return self._view.isServerValid(window, self._model.isServerValid)

    def _isPortValid(self, window):
        return self._view.isPortValid(window, self._model.isPortValid)

    def _isLoginNameValid(self, window):
        return self._view.isLoginNameValid(window, self._model.isStringValid)

    def _isPasswordValid(self, window):
        return self._view.isPasswordValid(window, self._model.isStringValid)
