#!
# -*- coding: utf_8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020 https://prrvchr.github.io                                     ║
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

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import createService
from smtpserver import getDialog
from smtpserver import getMessage
from smtpserver import logMessage


from .ispdbview import IspdbView

g_message = 'pagemanager'


import traceback


class IspdbManager(unohelper.Base):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._view = IspdbView(ctx, self, parent)

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        if self._model.isRefreshed():
            label = self._model.getPageLabel(self._pageid)
            self._view.setPageLabel(label % self._model.Email)
            config = self._model.getConfig()
            self._view.updatePage(config)

    def commitPage(self, reason):
        user, server = self._view.getConfiguration()
        self._model.updateConfiguration(user, server)
        if reason == FINISH:
            self._model.saveConfiguration()
        return True

    def canAdvance(self):
        host = self._view.getHost()
        port = self._view.getPort()
        valid = self._model.isConnectionValid(host, port)
        index = self._view.getAuthentication()
        print("IspdbManager.canAdvance() 1 %s" % valid)
        if valid and index > 0:
            login = self._view.getLogin()
            valid = self._model.isStringValid(login)
            print("IspdbManager.canAdvance() 2 %s" % valid)
            if valid and index < 3:
                pwd, conf = self._view.getPasswords()
                valid = self._model.isStringValid(pwd) and pwd == conf
                     
                print("IspdbManager.canAdvance() 3 %s" % valid)
        print("IspdbManager.canAdvance() 4 %s" % valid)
        return valid

    def _getAuthentication(self):
        return self._model.getAuthentication()

# IspdbManager setter methods
    def updateTravelUI(self):
        self._wizard.updateTravelUI()

    def changeConnection(self, i):
        j = self._view.getAuthentication()
        message, level = self._model.getSecurity(i, j)
        self._view.setSecurityMessage(message, level)

    def changeAuthentication(self, j):
        enabled = j > 0
        self._view.enableLogin(enabled)
        enabled = 0 < j < 3
        self._view.enablePassword(enabled)
        i = self._view.getConnection()
        message, level = self._model.getSecurity(i, j)
        self._view.setSecurityMessage(message, level)
        self._wizard.updateTravelUI()

    def previousServerPage(self):
        server = self._view.getServer()
        self._model.previousServerPage(server)
        config = self._model.getConfig()
        self._view.updatePage(config)

    def nextServerPage(self):
        server = self._view.getServer()
        self._model.nextServerPage(server)
        config = self._model.getConfig()
        self._view.updatePage(config)
