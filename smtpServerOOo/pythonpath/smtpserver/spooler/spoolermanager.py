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

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getUrlTransformer
from unolib import parseUrl

from .spoolermodel import SpoolerModel
from .spoolerview import SpoolerView

from ..mailer import MailerManager

from ..logger import logMessage
from ..logger import getMessage
g_message = 'spoolermanager'

import traceback


class SpoolerManager(unohelper.Base):
    def __init__(self, ctx):
        self._ctx = ctx
        self._model = SpoolerModel(ctx)
        self._view = None
        print("SpoolerManager.__init__()")

    @property
    def Model(self):
        return self._model

    def getRowSet(self):
        return self._model.getRowSet()

    def executeRowSet(self, rowset):
        self._model.executeRowSet(rowset)

    def viewSpooler(self, parent):
        self._view = SpoolerView(self._ctx, self, parent)
        self._view.execute()
        self._view.dispose()
        self._view = None
        print("SpoolerManager.viewSpooler() 1")

    def addDocument(self):
        transformer = getUrlTransformer(self._ctx)
        url = self._getDocumentUrl(transformer)
        if url is None:
            print("SpoolerManager.addDocument: %s" % 'No document')
            return
        print("SpoolerManager.addDocument: %s" % url.Name)
        self._showMailer(transformer, url)

    def removeDocument(self):
        pass

    def toogleRemove(self, enabled):
        self._view.enableButtonRemove(enabled)

    def _getDocumentUrl(self, transformer):
        url = None
        service = 'com.sun.star.util.PathSubstitution'
        directory = createService(self._ctx, service).substituteVariables('$(work)', True)
        service = 'com.sun.star.ui.dialogs.FilePicker'
        filepicker = createService(self._ctx, service)
        filepicker.setDisplayDirectory(directory)
        writer = self.Model.resolveString('FilePicker.Filter.Writer')
        filepicker.appendFilter(writer, '*.odt')
        filepicker.setCurrentFilter(writer)
        title = self.Model.resolveString('FilePicker.Title')
        filepicker.setTitle(title)
        if filepicker.execute() == OK:
            document = filepicker.getFiles()[0]
            url = parseUrl(transformer, document)
        return url

    def _showMailer(self, transformer, url):
        try:
            url = transformer.getPresentation(url, False)
            parent = self._view.getParent()
            mailer = MailerManager(self._ctx, self.Model.DataSource, parent, url)
            if mailer.show() == OK:
                self._addDocument(mailer, url)
            mailer.dispose()

        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _addDocument(self, mailer, url):
        print("SpoolerManager._addDocument: %s" % url.Name)
