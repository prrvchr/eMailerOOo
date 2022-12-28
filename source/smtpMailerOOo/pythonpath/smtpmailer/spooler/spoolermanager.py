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

from .spoolermodel import SpoolerModel
from .spoolerview import SpoolerView
from .spoolerhandler import DispatchListener
from .spoolerhandler import SpoolerListener

from ..grid import GridListener
from ..grid import RowSetListener

from ..unotool import createService
from ..unotool import executeDispatch
from ..unotool import getFileSequence
from ..unotool import getPropertyValueSet

from ..logger import LogHandler
from ..logger import Pool

from ..logger import getMessage
from ..logger import logMessage


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
        service = 'com.sun.star.mail.SpoolerService'
        self._spooler = createService(ctx, service)
        self._spoolerlistener = SpoolerListener(self)
        self._spooler.addListener(self._spoolerlistener)
        self._refreshSpoolerState()
        parent = self._view.getGridParent()
        possize = self._view.getGridPosSize()
        self._model.initSpooler(possize, parent, GridListener(self), self.initView)
        self._logger = Pool(ctx).getLogger('SpoolerLogger')
        self._loghandler = LogHandler(ctx, self.refreshLog)
        self._logger.addLogHandler(self._loghandler)
        self.refreshLog()

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

# SpoolerManager setter method
    def initView(self, rowset):
        with self._lock:
            if not self._model.isDisposed():
                rowset.addRowSetListener(RowSetListener(self))
                self._view.initView()

    def setDataModel(self, rowset):
        with self._lock:
            if not self._model.isDisposed():
                self._model.setGridData(rowset)

    def changeGridSelection(self, index, grid):
        self._view.enableButtonRemove(index != -1)

    def started(self):
        self._refreshSpoolerView(1)

    def stopped(self):
        self._model.executeRowSet()
        # TODO: We dont have the last line of the Logger through the LogHandler
        self.refreshLog()
        self._refreshSpoolerView(0)

    def saveGrid(self):
        self._model.saveGrid()

    def dispose(self):
        with self._lock:
            print("SpoolerManager.dispose() 1 ***************************")
            self._spooler.removeListener(self._spoolerlistener)
            self._logger.removeLogHandler(self._loghandler)
            self._model.dispose()
            self._view.dispose()
            print("SpoolerManager.dispose() 2 ***************************")

    def addDocument(self):
        arguments = getPropertyValueSet({'Path': self._model.Path,
                                         'Close': False})
        listener = DispatchListener(self)
        executeDispatch(self._ctx, 'smtp:mailer', arguments, listener)

    def documentAdded(self, path):
        self._model.Path = path
        self._model.executeRowSet()

    def removeDocument(self):
        rows = self._model.getGridSelectedRows()
        self._model.removeRows(rows)

    def toogleSpooler(self):
        if self._spooler.isStarted():
            self._spooler.stop()
        else:
            self._spooler.start()

    def closeSpooler(self):
        self._view.endDialog()

    def refreshLog(self):
        url = self._logger.getLoggerUrl()
        length, sequence = getFileSequence(self._ctx, url)
        text = sequence.value.decode('utf-8')
        self._view.refreshLog(text, length)

    def clearLog(self):
        msg = self._logger.getMessage(201)
        self._logger.clearLogger(msg)
        self.refreshLog()

# SpoolerManager private methods
    def _refreshSpoolerState(self):
        state = int(self._spooler.isStarted())
        self._refreshSpoolerView(state)

    def _refreshSpoolerView(self, state):
        label = self._model.getSpoolerState(state)
        self._view.setSpoolerState(label)
