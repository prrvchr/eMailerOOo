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

from smtpserver import createService
from smtpserver import executeDispatch
from smtpserver import getFileSequence
from smtpserver import getMessage
from smtpserver import getPropertyValueSet
from smtpserver import Logger
from smtpserver import LogHandler

from smtpserver import logMessage

from .spoolermodel import SpoolerModel
from .spoolerview import SpoolerView
from .spoolerhandler import DispatchListener


g_message = 'spoolermanager'

from threading import Condition
import traceback


class SpoolerManager(unohelper.Base):
    def __init__(self, ctx, datasource, parent):
        self._ctx = ctx
        self._lock = Condition()
        self._enabled = True
        self._model = SpoolerModel(ctx, datasource)
        self._view = SpoolerView(ctx, self, parent)
        service = 'com.sun.star.mail.MailServiceSpooler'
        self._spooler = createService(ctx, service)
        self._refreshSpoolerState()
        self._model.initSpooler(self.initView)
        #handler = LogHandler()
        self._logger = Logger(ctx, 'MailSpooler')

    @property
    def HandlerEnabled(self):
        return self._enabled

# SpoolerManager getter method
    def execute(self):
        return self._view.execute()

    def getTabPageTitle(self, tab):
        return self._model.getTabPageTitle(tab)

    def getDialogTitle(self):
        return self._model.getDialogTitle()

    def getGridModels(self, titles, width):
        return self._model.getGridModels(titles, width)

# SpoolerManager setter method
    def initView(self, titles, orders):
        with self._lock:
            if not self._model.isDisposed():
                # TODO: Attention: order is very important here
                # TODO: to have fonctionnal columns in the GridColumnModel
                self._view.initGrid(self, titles)
                self._model.executeRowSet()
                self._view.initColumnsList(titles)
                self._enabled = False
                self._view.initOrdersList(titles, orders)
                self._enabled = True
                self._view.initButtons()

    def dispose(self):
        with self._lock:
            self._model.dispose()
            self._view.dispose()

    def setGridColumnModel(self, titles, reset):
        self._model.setGridColumnModel(titles, reset)

    def changeOrder(self, orders):
        ascending = self._view.getSortDirection()
        self._model.setRowSetOrder(orders, ascending)

    def addDocument(self):
        arguments = getPropertyValueSet({'Path': self._model.Path,
                                         'Close': False})
        listener = DispatchListener(self)
        executeDispatch(self._ctx, 'smtp:mailer', arguments, listener)

    def addJob(self, path):
        self._model.Path = path
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
        self._refreshSpoolerState()

    def closeSpooler(self):
        self._model.save()
        self._view.endDialog()

    def refreshLog(self):
        print("SpoolerManager.refreshLog()")
        url = self._logger.getLoggerUrl()
        length, sequence = getFileSequence(self._ctx, url)
        text = sequence.value.decode('utf-8')
        self._view.setActivityLog(text)

    def clearLog(self):
        print("SpoolerManager.clearLog()")
        self._logger.clearLogger()
        self._view.setActivityLog('')

# SpoolerManager private methods
    def _refreshSpoolerState(self):
        state = int(self._spooler.isStarted())
        label = self._model.getSpoolerState(state)
        self._view.setSpoolerState(label)
