#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-24 https://prrvchr.github.io                                  ║
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

from .spoolerhandler import DialogHandler
from .spoolerhandler import RowSetListener
from .spoolerhandler import Tab1Handler
from .spoolerhandler import Tab2Handler

from ..listener import StreamListener

from ...grid import GridListener

from ...dispatchlistener import DispatchListener

from ...mailertool import getMailSpooler

from ...unotool import executeDispatch
from ...unotool import executeShell
from ...unotool import getPropertyValueSet
from ...unotool import getSimpleFile
from ...unotool import getTempFile

from ...logger import LogController
from ...logger import LoggerListener

from ...configuration import g_spoolerlog
from ...configuration import g_basename

from time import sleep
from threading import Condition
from threading import Thread
import traceback


class SpoolerManager(unohelper.Base):
    def __init__(self, ctx, datasource, parent):
        self._ctx = ctx
        self._lock = Condition()
        self._enabled = True
        self._model = SpoolerModel(ctx, datasource)
        titles = self._model.getDialogTitles()
        self._view = SpoolerView(ctx, DialogHandler(self), Tab1Handler(self), Tab2Handler(self), parent, *titles)
        self._spooler = getMailSpooler(ctx)
        self._spoolerlistener = StreamListener(self)
        self._spooler.addListener(self._spoolerlistener)
        self._refreshSpoolerState()
        window = self._view.getGridWindow()
        self._model.initSpooler(window, GridListener(self), self.initView)
        self._loggerlistener = LoggerListener(self)
        self._logger = LogController(ctx, g_spoolerlog, g_basename, self._loggerlistener)
        self.updateLogger()

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
        selected = index != -1
        sent = link = False
        if selected:
            sent, link = self._model.getRowClientInfo()
        self._view.enableButtons(selected, sent, link)

    def started(self):
        self._refreshSpoolerView(1)

    def closed(self):
        self._model.executeRowSet()
        self._refreshSpoolerView(0)

    def terminated(self):
        self._model.executeRowSet()
        self._refreshSpoolerView(0)

    def saveGrid(self):
        self._model.saveGrid()

    def dispose(self):
        with self._lock:
            self._spooler.removeListener(self._spoolerlistener)
            self._logger.removeModifyListener(self._loggerlistener)
            self._model.dispose()
            self._view.dispose()

    def addDocument(self):
        arguments = getPropertyValueSet({'Path': self._model.Path,
                                         'Close': False})
        listener = DispatchListener(self.documentAdded)
        executeDispatch(self._ctx, 'smtp:mailer', arguments, listener)

    def documentAdded(self, path):
        self._model.Path = path
        self._model.executeRowSet()

    def viewEml(self):
        self._view.disableButtons()
        job = self._model.getSelectedIdentifier('JobId')
        listener = DispatchListener(self._viewEml)
        arguments = getPropertyValueSet({'JobId': job})
        args = (listener, arguments)
        Thread(target=self._executeDispatch, args=args).start()

    def viewClient(self):
        self._view.disableButtons()
        self._viewClient(**self._model.getCommandArguments())

    def viewWeb(self):
        self._view.disableButtons()
        sender = self._model.getSelectedColumn('Sender')
        self._viewWeb(sender, **self._model.getCommandArguments())

    def _viewClient(self, **args):
        command, option = self._model.getClientCommand(args)
        print("SpoolerManager._viewLocal() Command: %s" % command)
        self._executeCommand(command, option)

    def _viewWeb(self, sender, **args):
        command = None
        client = self._model.getSenderClient(sender)
        if client:
            command, option = self._model.getWebLinkCommand(client, args)
        print("SpoolerManager._viewWeb() Command: %s" % command)
        self._executeCommand(command, option)

    def _executeCommand(self, command, option):
        if command:
            executeShell(self._ctx, command, option)
        self._enableButtonView(self._model.hasGridSelectedRows())

    def _executeDispatch(self, listener, arguments):
        executeDispatch(self._ctx, 'smtp:mail', arguments, listener)

    def _viewEml(self, mail):
        url = '%s/Email.eml' % getTempFile(self._ctx).Uri
        output = getSimpleFile(self._ctx).openFileWrite(url)
        output.writeBytes(uno.ByteSequence(mail.asBytes()))
        output.flush()
        output.closeOutput()
        executeShell(self._ctx, url)
        self._enableButtonView(self._model.hasGridSelectedRows())

    def _enableButtonView(self, selected):
        if selected:
            sent, link = self._model.getRowClientInfo()
        self._view.enableButtons(selected, sent, link)

    def removeDocument(self):
        rows = self._model.getGridSelectedRows()
        self._model.deselectAllRows()
        self._model.removeRows(rows)

    def toogleSpooler(self, state):
        if state:
            self._spooler.start()
        else:
            self._spooler.terminate()

    def closeSpooler(self):
        self._view.endDialog()

    def clearLogger(self):
        self._logger.clearLogger()

    def updateLogger(self):
        self._view.updateLog(*self._logger.getLogContent(True))

# SpoolerManager private methods
    def _refreshSpoolerState(self):
        state = int(self._spooler.isStarted())
        self._refreshSpoolerView(state)

    def _refreshSpoolerView(self, state):
        label = self._model.getSpoolerState(state)
        self._view.setSpoolerState(label, state)
