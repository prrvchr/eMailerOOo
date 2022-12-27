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

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .sendermodel import SenderModel
from .senderview import SenderView

from ..mailer import MailerManager

from ..logger import getMessage
from ..logger import logMessage

g_message = 'sendermanager'

from threading import Condition
import traceback


class SenderManager(unohelper.Base):
    def __init__(self, ctx, model, parent, url):
        self._ctx = ctx
        self._lock = Condition()
        self._model = model
        self._view = SenderView(ctx, self, parent)
        datasource = self._model.DataSource
        parent = self._view.getParent()
        path = self._model.Path
        self._mailer = MailerManager(ctx, self, datasource, parent, path)
        self._model.getDocument(url, self.initMailer)
        print("SenderManager.__init__()")

    @property
    def Model(self):
        return self._model
    @property
    def Mailer(self):
        return self._mailer

    def initMailer(self, document, title):
        with self._lock:
            if not self._model.isDisposed():
                # TODO: Document can be <None> if a lock or password exists !!!
                # TODO: It would be necessary to test a Handler on the descriptor...
                self._mailer.initView(document)
                self._view.setTitle(title)

    def execute(self):
        return self._view.execute()

    def updateUI(self, enabled):
        self._view.enableButtonSend(enabled)

    def sendDocument(self):
        self._mailer.sendDocument()
        self._view.endDialog()

    def dispose(self):
        with self._lock:
            self._mailer.Model.dispose()
            self._model.dispose()
            self._view.dispose()
