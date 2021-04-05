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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import MailManager

from smtpserver import createService
from smtpserver import getMessage
from smtpserver import logMessage

from .mailermodel import MailerModel
from .mailerview import MailerView

g_message = 'mailermanager'

from threading import Condition
import traceback


class MailerManager(MailManager):
    def __init__(self, ctx, sender, datasource, parent, path):
        self._ctx = ctx
        self._sender = sender
        self._disabled = False
        self._lock = Condition()
        self._model = MailerModel(ctx, datasource, path)
        self._view = MailerView(ctx, self, parent, 1)
        self._model.getSenders(self.initSenders)

# MailerManager setter methods
    def sendDocument(self):
        try:
            subject, attachments = self._getSavedDocumentProperty()
            print("MailerManager.sendDocument() *************************** %s - %s" % (subject, attachments))
            sender = self._view.getSender()
            recipients = self._view.getRecipients()
            url = self._model.getUrl()
            service = 'com.sun.star.mail.MailServiceSpooler'
            spooler = createService(self._ctx, service)
            id = spooler.addJob(sender, subject, url, recipients, attachments)
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)


# MailerManager private setter methods
    def _closeDocument(self, document):
        document.close(True)

    def _updateUI(self):
        cansend = self._canAdvance()
        self._sender.updateUI(cansend)
