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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService

from .mailermodel import MailerModel
from .mailerview import MailerView

from ..logger import logMessage
from ..logger import getMessage
g_message = 'mailermanager'

import traceback


class MailerManager(unohelper.Base):
    def __init__(self, ctx, datasource, url, parent):
        self._ctx = ctx
        self._model = MailerModel(ctx, datasource, url)
        self._view = MailerView(ctx, self, parent)
        print("MailerManager.__init__()")

    @property
    def Model(self):
        return self._model

    def show(self):
        return self._view.execute()

    def dispose(self):
        self._view.dispose()

    def enableRemoveSender(self, enabled):
        self._view.enableRemoveSender(enabled)

    def addSender(self, sender):
        self._view.addSender(sender)

    def removeSender(self):
        # TODO: button 'RemoveSender' must be deactivated to avoid multiple calls  
        self._view.enableRemoveSender(False)
        sender, position = self._view.getSender()
        status = self.Model.removeSender(sender)
        if status == 1:
            self._view.removeSender(position)

    def editRecipient(self, email, exist):
        enabled = self._validateRecipient(email, exist)
        self._view.enableAddRecipient(enabled)
        self._view.enableRemoveRecipient(exist)

    def addRecipient(self):
        self._view.addToRecipient()

    def changeRecipient(self):
        self._view.enableRemoveRecipient(True)

    def removeRecipient(self):
        self._view.removeRecipient()

    def enterRecipient(self, control, email, exist):
        if self._validateRecipient(email, exist):
            self._view.addRecipient(control, email)

    def sendAsHtml(self):
        print("MailerManager.sendAsHtml()")
        self._view.setStep(1)

    def sendAsAttachment(self):
        print("MailerManager.sendAsAttachment()")
        self._view.setStep(2)

    def viewHtmlDocument(self):
        print("MailerManager.viewHtmlDocument()")

    def _validateRecipient(self, email, exist):
        return all((self.Model.isEmailValid(email), not exist))
