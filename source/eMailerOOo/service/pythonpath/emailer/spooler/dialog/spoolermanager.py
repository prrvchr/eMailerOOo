#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
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

from com.sun.star.awt.MessageBoxType import WARNINGBOX

from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.logging.LogLevel import ALL
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.util import CloseVetoException

from .spoolermodel import SpoolerModel

from .spoolerview import SpoolerView

from .spoolerhandler import CloseListener
from .spoolerhandler import WindowHandler
from .spoolerhandler import RowSetListener
from .spoolerhandler import TabHandler
from .spoolerhandler import TabPageListener
from .spoolerhandler import LoggerListener

from ...grid import GridSelectionListener

from ...listener import DispatchListener
from ...listener import StreamListener

from ...helper import getMailSender

from ...unotool import createMessageBox
from ...unotool import executeDesktopDispatch
from ...unotool import executeFrameDispatch
from ...unotool import executeShell
from ...unotool import getPropertyValueSet
from ...unotool import getSimpleFile
from ...unotool import getTempFile

from ...logger import LogController

from ...configuration import g_mailservicelog
from ...configuration import g_spoolerframe
from ...configuration import g_spoolerlog
from ...configuration import g_basename

from time import sleep
from threading import Lock
import traceback


class SpoolerManager(unohelper.Base):
    def __init__(self, ctx, datasource, notifier):
        self._ctx = ctx
        self._notifier = notifier
        self._lock = Lock()
        self._closing = False
        self._enabled = True
        self._initialize = True
        self._model = SpoolerModel(ctx, datasource)
        handler = WindowHandler(self)
        listener = CloseListener(self)
        handler1 = TabHandler(self)
        listener1 = TabPageListener(self)
        point = self._model.getDialogPosition()
        titles = self._model.getDialogTitles()
        self._view = SpoolerView(ctx, handler, listener, handler1, listener1, point, titles)
        self._sender = getMailSender(ctx)
        self._senderlistener = StreamListener(self)
        window = self._view.getGridWindow()
        self._model.initSpooler(window, GridSelectionListener(self), RowSetListener(self))
        self._loglistener1 = LoggerListener(self.updateLog1)
        self._log1 = LogController(ctx, g_spoolerlog, g_basename, self._loglistener1)
        self._log1.addRollerHandler()
        self._loglistener2 = LoggerListener(self.updateLog2)
        self._log2 = LogController(ctx, g_mailservicelog, g_basename, self._loglistener2)
        self._log2.addRollerHandler()
        self._closelistener = listener
        #self._updateLogger()

    @property
    def HandlerEnabled(self):
        return self._enabled

# XCallback
    def notify(self, rowset):
        with self._lock:
            if not self._model.isDisposed():
                self._view.initView()
                self._sender.addListener(self._senderlistener)

# XDispatchResultListener
    def dispatchFinished(self, notification):
        self._model.endDispatch()
        if self._closing:
            if not self._model.hasDispatch():
                self._close()
                self._view.close()
        else:
            if notification.State == SUCCESS:
                executeShell(self._ctx, notification.Result)
            else:
                parent = self._view.getWindow().Peer
                title = self._model.getMsgBoxTitle()
                dialog = createMessageBox(parent, WARNINGBOX, 1, title, notification.Result)
                dialog.execute()
                dialog.dispose()
            self._enableButtons(self._model.hasGridSelectedRows())
            self._view.enableStartSpooler(True)

# XCloseListener
    def queryClosing(self, source, ownership):
        if not ownership:
            if self._closing:
                raise CloseVetoException()
            self._closing = True
            self._view.disableButtons()
            if self._model.hasDispatch():
                if self._model.isSenderStarted():
                    self._sender.terminate()
                raise CloseVetoException()
            self._close()

    def notifyClosing(self, source):
        source.removeCloseListener(self._closelistener)
        self._dispose()

# XStreamListener
    def started(self):
        self._refreshSpoolerView(2)

    def closed(self):
        self._refreshSpoolerView(1)
        if self._closing:
            if not self._model.hasDispatch():
                self._close()
                self._view.close()
        else:
            self._enableButtons(self._model.hasGridSelectedRows())

    def terminated(self):
        self._refreshSpoolerView(0)

# XTabPageContainerListener
    def activateTab(self, tab):
        self._view.activateTab(tab)

# SpoolerManager getter method
    def getTabPageTitle(self, tab):
        return self._model.getTabPageTitle(tab)

    def getDialogTitle(self):
        return self._model.getDialogTitle()

# SpoolerManager setter method
    def setDataModel(self, rowset):
        with self._lock:
            if not self._model.isDisposed():
                self._model.setGridData(rowset)
                if self._initialize:
                    self._initView()

    def changeGridSelection(self, index, grid):
        selected = index != -1
        if selected:
            sent, link = self._model.getRowClientInfo()
        else:
            sent = link = False
        self._view.enableButtons(selected, sent, link)

    def saveGrid(self):
        self._model.saveGrid()

    def addDocument(self):
        self._model.addDocument()

    def viewEml(self):
        self._view.disableButtons()
        self._view.enableStartSpooler(False)
        self._model.startDispatch(DispatchListener(self))

    def viewClient(self):
        self._view.disableButtons()
        self._view.enableStartSpooler(False)
        self._viewClient(**self._model.getCommandArguments())

    def viewWeb(self):
        self._view.disableButtons()
        self._view.enableStartSpooler(False)
        sender = self._model.getSelectedColumn('Sender')
        self._viewWeb(sender, **self._model.getCommandArguments())

    def resubmitJobs(self):
        self._model.resubmitJobs('JobId')

    def _viewClient(self, **args):
        command, option = self._model.getClientCommand(args)
        self._executeCommand(command, option)

    def _viewWeb(self, sender, **args):
        command = None
        client = self._model.getSenderClient(sender)
        if client:
            command, option = self._model.getUrlCommand(client, args)
        self._executeCommand(command, option)

    def _executeCommand(self, command, option):
        if command:
            executeShell(self._ctx, command, option)
        self._enableButtons(self._model.hasGridSelectedRows())
        self._view.enableStartSpooler(True)

    def _enableButtons(self, selected):
        if selected:
            sent, link = self._model.getRowClientInfo()
        else:
            sent = link = False
        self._view.enableButtons(selected, sent, link)

    def removeDocument(self):
        rows = self._model.getGridSelectedRows()
        self._model.deselectAllRows()
        self._model.removeRows(rows)

    def toogleSpooler(self, state):
        if state:
            self._view.disableButtons()
            self._sender.start()
        else:
            self._sender.terminate()

    def clearLogger(self):
        tab = self._view.getActiveTab()
        if tab == 2:
            self._log1.clearLogger()
        elif tab == 3:
            self._log2.clearLogger()

    def closeSpooler(self):
        self._view.disableButtons()
        if not self._model.hasDispatch():
            self._close()
            self._view.close()
        else:
            self._closing = True
            if self._model.isSenderStarted():
                self._sender.terminate()

    def clearLog1(self):
        self._log1.clearLogger()

    def clearLog2(self):
        self._log2.clearLogger()

    def updateLog1(self):
        self._view.updateLog1(*self._log1.getLogContent(True))

    def updateLog2(self):
        self._view.updateLog2(*self._log2.getLogContent(True))

# SpoolerManager private methods
    def _close(self):
        with self._lock:
            self._model.save(self._view.getDialogPosition())
            self._sender.removeListener(self._senderlistener)
            self._log1.removeModifyListener(self._loglistener1)
            self._log2.removeModifyListener(self._loglistener2)
            self._log1.removeRollerHandler()
            self._log2.removeRollerHandler()

    def _dispose(self):
        with self._lock:
            self._model.dispose()

    def _updateLogger(self):
        self.updateLog1()
        self.updateLog2()

    def _refreshSpoolerView(self, status):
        self._view.setSpoolerState(*self._model.setSpoolerStatus(status))

    def _initView(self):
        self._initialize = False
        self._view.initView()
        self._sender.addListener(self._senderlistener)
        if self._notifier:
            struct = 'com.sun.star.frame.DispatchResultEvent'
            notification = uno.createUnoStruct(struct, self, SUCCESS, None)
            self._notifier.dispatchFinished(notification)

