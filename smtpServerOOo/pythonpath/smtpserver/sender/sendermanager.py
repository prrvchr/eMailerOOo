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

import traceback


class SenderManager(unohelper.Base):
    def __init__(self, ctx, path):
        self._ctx = ctx
        self._model = SenderModel(ctx, path)
        self._view = SenderView(ctx)
        self._mailer = None
        print("SenderManager.__init__()")

    def getPath(self):
        return self._model.Path if self._mailer is None else self._mailer.Model.Path

    def getDocumentUrlAndPath(self):
        resource = self._view.getFilePickerTitleResource()
        title = self._model.resolveString(resource)
        resource = self._view.getFilePickerFilterResource()
        writer = self._model.resolveString(resource)
        filter = (writer, '*.odt')
        url, path = self._view.getDocumentUrlAndPath(self._model.Path, title, filter)
        return url, path

    def showDialog(self, datasource, parent, url, path):
        self._view.setDialog(self, parent)
        parent = self._view.getParent()
        self._mailer = MailerManager(self._ctx, datasource, parent, path)
        self._model.getDocument(url, self.initMailer)
        return self._view.execute()

    def initMailer(self, document):
        if not self._view.isDisposed():
            resource = self._view.getTitleRessource()
            title = self._model.getDocumentTitle(document.URL, resource)
            self._view.setTitle(title)
            self._mailer.initView(document)

    def dispose(self):
        self._view.dispose()
        self._mailer.dispose()
