#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
║                                                                                    ║
║   Permission is hereby granted, free of charge, to any person obtaining            ║
║   a copy of this software and associated documentation files (the "Software"),     ║
║   to deal in the Software without restriction, including without limitation        ║
║   the rights to use, copy, modify, merge, publish, distribute, sublicense,         ║
║   and/or sell copies of the Software, and to permit persons to whom the Software   ║
║   is furnished to do so, subject to the following conditions:                      ║
║                                                                                    ║
║   The above copyright notice and this permission notice shall be included in       ║
║   all copies or substantial portions of the Software.                              ║
║                                                                                    ║
║   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,                  ║
║   EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES                  ║
║   OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.        ║
║   IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY             ║
║   CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,             ║
║   TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE       ║
║   OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.                                    ║
║                                                                                    ║
╚════════════════════════════════════════════════════════════════════════════════════╝
"""

import uno
import unohelper

from com.sun.star.uno import Exception as UnoException

from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .ispdbhandler import WindowHandler

from .ispdbview import IspdbView

from ...unotool import getStringResource

from ...configuration import g_identifier

import traceback


class IspdbManager(unohelper.Base):
    def __init__(self, ctx, wizard, model, pageid, parent, service):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'IspdbPages')
        self._view = IspdbView(ctx, WindowHandler(self), parent, pageid)
        self._service = service.value
        self._version = 0

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        if self._model.refreshView(self._version):
            label = self._model.getPagesLabel(self._resolver, self._service)
            self._view.setPageLabel(label)
            config = self._model.getConfig(self._service)
            self._view.updatePage(config)
            self._version = self._model.getVersion()

    def commitPage(self, reason):
        server, user = self._view.getConfiguration(self._service)
        self._model.updateConfiguration(self._service, server, user)
        # If we are in offline mode, this page may be the last
        # and in this case we must save the configuration
        if reason == FINISH:
            self._model.saveConfiguration()
        return True

    def canAdvance(self):
        host, port, index = self._view.getData()
        valid = self._model.isConnectionValid(host, port)
        if valid and index > 0:
            login = self._view.getLogin()
            valid = self._model.isStringValid(login)
            if valid and index < 2:
                pwd, conf = self._view.getPasswords()
                valid = self._model.isStringValid(pwd) and pwd == conf
        return valid

# IspdbManager setter methods
    def updateTravelUI(self):
        self._wizard.updateTravelUI()

    def changeConnection(self, i):
        j = self._view.getAuthentication()
        message, level = self._model.getSecurity(self._resolver, i, j)
        self._view.setSecurityMessage(message, level)

    def changeAuthentication(self, j):
        self._view.enableLogin(j > 0)
        self._view.enablePassword(j == 1)
        i = self._view.getConnection()
        message, level = self._model.getSecurity(self._resolver, i, j)
        self._view.setSecurityMessage(message, level)
        self._wizard.updateTravelUI()

    def previousServerPage(self):
        server = self._view.getServerConfig(self._service)
        config = self._model.previousServerPage(self._service, server)
        config = self._model.getConfig(self._service)
        self._view.updatePage(config)

    def nextServerPage(self):
        server = self._view.getServerConfig(self._service)
        config = self._model.nextServerPage(self._service, server)
        config = self._model.getConfig(self._service)
        self._view.updatePage(config)
