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

from .sendermodel import SenderModel
from .senderview import SenderView

from smtpserver.mailer import MailerManager

from smtpserver import logMessage
from smtpserver import getMessage
g_message = 'sendermanager'

from threading import Condition
import traceback


class SenderManager(unohelper.Base):
    def __init__(self, ctx, path):
        self._ctx = ctx
        self._lock = Condition()
        self._model = SenderModel(ctx, path)
        self._view = SenderView(ctx)
        self._mailer = None
        print("SenderManager.__init__()")

    @property
    def Model(self):
        return self._model
    @property
    def Mailer(self):
        return self._mailer

    def getDocumentUrl(self):
        title = self._model.getFilePickerTitle()
        filters = self._model.getFilePickerFilters()
        path = self._model.Path
        url, path = self._view.getDocumentUrl(title, filters, path)
        self._model.Path = path
        return url

    def showDialog(self, datasource, parent, url):
        self._view.setDialog(self, parent)
        parent = self._view.getParent()
        path = self.Model.Path
        self._mailer = MailerManager(self._ctx, self, datasource, parent, path)
        self._model.getDocument(url, self.initMailer)
        return self._view.execute()

    def initMailer(self, document):
        with self._lock:
            if not self._view.isDisposed():
                # TODO: Document can be <None> if a lock or password exists !!!
                # TODO: It would be necessary to test a Handler on the descriptor...
                title = self._model.getDocumentTitle(document.URL)
                self._mailer.initView(document)
                self._view.setTitle(title)

    def updateUI(self, enabled):
        self._view.enableButtonSend(enabled)

    def sendDocument(self):
        self._mailer.sendDocument()
        self._view.endDialog()

    def dispose(self):
        with self._lock:
            self._view.dispose()
            self._mailer.dispose()
