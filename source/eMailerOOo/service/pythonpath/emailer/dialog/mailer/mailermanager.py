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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from ..mail import MailManager
from ..mail import WindowHandler

from .mailermodel import MailerModel
from .mailerview import MailerView
from .mailerhandler import DialogHandler

from ...helper import getMailSpooler

import traceback


class MailerManager(MailManager):
    def __init__(self, ctx, model, parent, url):
        super().__init__(ctx, model)
        self._view = MailerView(ctx, DialogHandler(self), WindowHandler(ctx, self), parent, 1)
        self._view.setSenders(self._model.getSenders())
        self._model.loadDocument(url, self)

# XDispatchResultListener
    def _dispatchFinished(self, call, state, value):
        if call == 'Sender' and state == SUCCESS:
            self._view.addSender(value)
            self._updateUI()

# MailerManager setter methods
    def addRecipient(self):
        email = self._view.getRecipient()
        if self._model.isEmailValid(email):
            self._view.addRecipient(email)
            self._updateUI()

    def changeRecipient(self):
        self._view.enableRemoveRecipient(True)

    def editRecipient(self, email, exist):
        self._view.enableAddRecipient(not exist and self._model.isEmailValid(email))
        self._view.enableRemoveRecipient(exist)

    def enterRecipient(self, email):
        if self._model.isEmailValid(email):
            self._view.addRecipient(email)
            self._updateUI()

    def removeRecipient(self):
        self._view.removeRecipient()
        self._updateUI()

    def dispose(self):
        with self._lock:
            self._model.dispose()
            self._view.dispose()

    def execute(self):
        return self._view.execute()

    def sendDocument(self):
        try:
            subject, attachments = self._getSavedDocumentProperty()
            sender, recipients = self._view.getEmail()
            url = self._model.getUrl()
            spooler = getMailSpooler(self._ctx)
            id = spooler.addJob(sender, subject, url, recipients, attachments)
            spooler.dispose()
            self._view.endDialog()
        except Exception as e:
            msg = "Error: %s" % traceback.format_exc()
            print(msg)

# MailerManager private setter methods
    def _updateUI(self):
        enabled = self._canAdvance()
        self._view.enableButtonSend(enabled)

    def _notifyInit(self, document, title):
        if not self._model.isDisposed():
            self._initView(document)
            self._view.setTitle(title)
        self._model.closeDocument(document)

