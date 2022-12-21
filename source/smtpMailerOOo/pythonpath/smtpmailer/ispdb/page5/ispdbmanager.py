#!
# -*- coding: utf-8 -*-

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

import unohelper

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from smtpmailer import Logger
from smtpmailer import createService
from smtpmailer import getDialog

from .ispdbview import IspdbView
from .send import SendView

import traceback


class IspdbManager(unohelper.Base):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._connected = False
        self._dialog = None
        self._view = IspdbView(ctx, self, parent)

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        self._connected = False
        self._view.setPageStep(1)
        self._model.connectServers(self.resetProgress, self.updateProgress, self.setLabel, self.setStep)

    def commitPage(self, reason):
        if reason == FINISH:
            self._model.saveConfiguration()
        return True

    def canAdvance(self):
        return self._connected

# IspdbManager setter methods
    def setLabel(self, *format):
        if not self._model.isDisposed():
            label = self._model.getPageLabel(self._pageid)
            self._view.setPageLabel(label % format)

    def updateProgress(self, value):
        if not self._model.isDisposed():
            self._view.updateProgress(value)

    def resetProgress(self, value):
        if not self._model.isDisposed():
            self._view.resetProgress(value)

    def setStep(self, step):
        if not self._model.isDisposed():
            self._connected = step > 3
            self._view.setPageStep(step)
            self._wizard.updateTravelUI()

    def sendMail(self):
        parent = self._wizard.DialogWindow.getPeer()
        title = self._model.getSendTitle()
        email = self._model.Email
        subject = self._model.getSendSubject()
        msg = self._model.getSendMessage()
        self._dialog = SendView(self._ctx, self, parent, title, email, subject, msg)
        if self._dialog.execute() == OK:
            self._view.setPageStep(1)
            recipient = self._dialog.getRecipient()
            subject = self._dialog.getSubject()
            msg = self._dialog.getMessage()
            self._model.sendMessage(recipient, subject, msg, self.resetProgress, self.updateProgress, self.setStep)
        self._dialog.dispose()
        self._dialog = None

    def updateDialog(self):
        recipient = self._dialog.getRecipient()
        subject = self._dialog.getSubject()
        msg = self._dialog.getMessage()
        enabled = self._model.validSend(recipient, subject, message)
        self._dialog.enableButtonSend(enabled)
