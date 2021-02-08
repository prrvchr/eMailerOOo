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
from unolib import executeDispatch
from unolib import getPropertyValueSet

from .spoolermodel import SpoolerModel
from .spoolerview import SpoolerView
from .spoolerhandler import DispatchListener

from ..sender import SenderManager
#from ..mailer import MailerManager

from ..logger import logMessage
from ..logger import getMessage
g_message = 'spoolermanager'

import time
import traceback


class SpoolerManager(unohelper.Base):
    def __init__(self, ctx, datasource, parent):
        self._ctx = ctx
        self._model = SpoolerModel(ctx, datasource)
        rowset = self._model.getRowSet()
        self._view = SpoolerView(ctx, self, rowset, parent)
        service = 'com.sun.star.mail.MailServiceSpooler'
        self._spooler = createService(ctx, service)
        self.refreshSpoolerState()
        self._model.initRowSet()
        print("SpoolerManager.__init__()")

    @property
    def Model(self):
        return self._model

    def execute(self):
        return self._view.execute()

    def dispose(self):
        self._view.dispose()

    def addDocument1(self):
        try:
            sender = SenderManager(self._ctx, self._model.Path)
            url, self._model.Path = sender.getDocumentUrlAndPath()
            if url is not None:
                if sender.showDialog(self._model.DataSource, self._view.getParent(), url, self._model.Path) == OK:
                    self._model.Path = sender.getPath()
                    self._model.executeRowSet()
                sender.dispose()
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def addDocument(self):
        arguments = getPropertyValueSet({'Path': self._model.Path})
        listener = DispatchListener(self)
        executeDispatch(self._ctx, 'smtp:mailer', arguments, listener)

    def addJob(self, path):
        self._model.Path = path
        time.sleep(5)
        self._model.executeRowSet()

    def removeDocument(self):
        self._model.executeRowSet()

    def toogleRemove(self, enabled):
        self._view.enableButtonRemove(enabled)

    def toogleSpooler(self):
        if self._spooler.isStarted():
            self._spooler.stop()
        else:
            self._spooler.start()
        self.refreshSpoolerState()

    def refreshSpoolerState(self):
        state = int(self._spooler.isStarted())
        resource = self._view.getStateResource(state)
        label = self._model.resolveString(resource)
        self._view.setSpoolerState(label)
