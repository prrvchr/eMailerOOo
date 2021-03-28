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

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService

from smtpserver.mail import MailManager
from .mergerview import MergerView
from .mergerhandler import RecipientHandler

from smtpserver import logMessage
from smtpserver import getMessage

from threading import Condition
import traceback


class MergerManager(MailManager,
                    XWizardPage):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._disabled = False
        self._lock = Condition()
        self._view = MergerView(ctx, self, parent, 2)
        self._model.getSenders(self.initSenders)
        handler = RecipientHandler(self)
        self._model.initView(handler, self.initView, self.initRecipients)

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        pass

    def commitPage(self, reason):
        return True

    def canAdvance(self):
        return self._canAdvance()

# MergerManager setter methods
    def initRecipients(self, recipients):
        print("MergerManager.initRecipient()")
        self._view.setMergerRecipient(recipients)

    def changeRecipient(self):
        recipients = self._model.getRecipients()
        print("MergerManager.changeRecipient()")
        self._view.setMergerRecipient(recipients)

    def sendDocument(self):
        subject, attachments = self._getSavedDocumentProperty()
        sender = self._view.getSender()
        recipients = self._view.getRecipients()
        url = self._model.getUrl()
        service = 'com.sun.star.mail.MailServiceSpooler'
        spooler = createService(self._ctx, service)
        id = spooler.addJob(sender, subject, url, recipients, attachments)

# MergerManager private setter methods
    def _closeDocument(self, document):
        url = self._model.getUrl()
        if document.URL != url:
            document.close(True)

    def _updateUI(self):
        self._wizard.updateTravelUI()
