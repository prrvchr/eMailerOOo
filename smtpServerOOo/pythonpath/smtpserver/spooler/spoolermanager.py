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
from unolib import executeDispatch
from unolib import getPropertyValueSet

from .spoolermodel import SpoolerModel
from .spoolerview import SpoolerView
from .spoolerhandler import DispatchListener

from smtpserver import logMessage
from smtpserver import getMessage
g_message = 'spoolermanager'

import traceback


class SpoolerManager(unohelper.Base):
    def __init__(self, ctx, datasource, parent):
        self._ctx = ctx
        self._model = SpoolerModel(ctx, datasource)
        self._view = SpoolerView(ctx, self, parent)
        service = 'com.sun.star.mail.MailServiceSpooler'
        self._spooler = createService(ctx, service)
        self._refreshSpoolerState()
        self._model.initRowSet(self.initView)
        print("SpoolerManager.__init__()")

    @property
    def Model(self):
        return self._model

    def execute(self):
        return self._view.execute()

    def dispose(self):
        self._view.dispose()

    def initView(self):
        try:
            #query = self.Model.getQuery()
            self.Model.setRowSet()
            #composer = self.Model.getQueryComposer(query)
            #columns = composer.getTables().getByName('Spooler').getColumns().getElementNames()
            #enumeration = composer.getColumns().createEnumeration()
            #self._view.setColumnsList(columns, enumeration)
            #enumeration = composer.getOrderColumns().createEnumeration()
            #self._view.setOrdersList(columns, enumeration)
            #mri = createService(self._ctx, 'mytools.Mri')
            #mri.inspect(composer)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def changeColumn(self, columns):
        print("SpoolerManager.changeColumn() %s" % (columns, ))

    def changeOrder(self, orders):
        print("SpoolerManager.changeOrder() %s" % (orders, ))

    def addDocument(self):
        arguments = getPropertyValueSet({'Path': self._model.Path})
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

# SpoolerManager private methods
    def _refreshSpoolerState(self):
        state = int(self._spooler.isStarted())
        label = self._model.getSpoolerState(state)
        self._view.setSpoolerState(label)
