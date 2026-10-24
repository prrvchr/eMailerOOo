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

from ...logger import LoggerListener

from ...grid import GridSelectionListener

from ...listener import DispatchListener
from ...listener import StreamListener

from ...runner import Cancel
from ...runner import Runner

from ...dbsetup import DBSetup

from ...logger import LogController

from ...helper import getMailSender

from ...unotool import createMessageBox
from ...unotool import executeDesktopDispatch
from ...unotool import executeShell
from ...unotool import notifyDispatch

from ...helper import showController

from ...configuration import g_mailservicelog
from ...configuration import g_spoolerframe
from ...configuration import g_spoolerlog
from ...configuration import g_basename

from time import sleep
import traceback


class SpoolerManager():
    def __init__(self, ctx, source, /, dispatch=None, **kwargs):
        self._ctx = ctx
        self._source = source
        self._dispatch = dispatch
        self._model = SpoolerModel(ctx)
        self._running = False
        self._notifier = None
        self._closing = False
        self._enabled = True
        self._initialize = True
        self._closelistener = CloseListener(self)
        self._view = SpoolerView(ctx, WindowHandler(self), self._closelistener, TabHandler(self), TabPageListener(self), *self._model.getViewData())
        #self._sender = None
        #self._senderlistener = StreamListener(self)
        self._loglistener1 = LoggerListener(self.updateLog1)
        self._log1 = LogController(ctx, g_spoolerlog, g_basename, self._loglistener1)
        self._log1.addRollerHandler()
        self._loglistener2 = LoggerListener(self.updateLog2)
        self._log2 = LogController(ctx, g_mailservicelog, g_basename, self._loglistener2)
        self._log2.addRollerHandler()
        setup = Runner(ctx, DBSetup(self._ctx).withCallback(self.notify), self._view.getIndicator(), Cancel())
        setup.start()

    def notify(self, success, created):
        self._view.resetIndicator()
        if success:
            self._model.initSpooler(RowSetListener(self))

    @property
    def HandlerEnabled(self):
        return self._enabled

# XDispatchResultListener
    def dispatchFinished(self, notification):
       print("SpoolerManager.dispatchFinished() 1")

    def dispatchFinished1(self, notification):
        self._model.endDispatch()
        if self._closing:
            if not self._model.hasDispatch():
                self._close()
                self._view.dispose()
        else:
            if notification.State == SUCCESS:
                executeShell(self._ctx, notification.Result)
            else:
                title = self._model.getMsgBoxTitle()
                dialog = createMessageBox(self._ctx, title, notification.Result, WARNINGBOX)
                dialog.execute()
                dialog.dispose()
            self._enableButtons(self._model.hasGridSelectedRows())

# XCloseListener
    def queryClosing(self, source, ownership):
        print("SpoolerManager.queryClosing() ownership: %s" % ownership)
        self._close()
        if self._running:
            raise CloseVetoException()
        if not ownership:
            self._dispose()

    def notifyClosing(self, source):
        print("SpoolerManager.notifyClosing()")
        source.removeCloseListener(self._closelistener)
        #self._dispose()

# XTabPageContainerListener
    def activateTab(self, tab):
        self._view.activateTab(self._ctx, tab)

# SpoolerManager getter method
    def getTabPageTitle(self, tab):
        return self._model.getTabPageTitle(tab)

    def getDialogTitle(self):
        return self._model.getDialogTitle()

# SpoolerManager setter method
    def setDataModel(self, rowset):
        print("SpoolerManager.setDataModel()")
        if not self._model.isDisposed():
            if not self._model.isGridInitialized():
                #self._sender = getMailSender(self._ctx)
                #self._senderlistener = StreamListener(self)
                self._model.initGrid(rowset, self._view.getGridWindow(), GridSelectionListener(self))
                self._view.initView()
            self._model.setGridData(rowset)
            if self._dispatch:
                notifyDispatch(self._source, self._dispatch)

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
        self._model.startDispatch(self._view.getFrame(), DispatchListener(self))

    def viewClient(self):
        self._view.disableButtons()
        self._viewClient(**self._model.getCommandArguments())

    def viewWeb(self):
        self._view.disableButtons()
        sender = self._model.getSelectedColumn('Sender')
        self._viewWeb(sender, **self._model.getCommandArguments())

    def resubmitJobs(self):
        column = 'JobId'
        self._model.resubmitJobs(column)

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

    def showController(self):
        showController(self._ctx)

    def clearLogger(self):
        tab = self._view.getActiveTab()
        if tab == 2:
            self._log1.clearLogger()
        elif tab == 3:
            self._log2.clearLogger()

    def closeSpooler(self):
        self._close()
        if not self._running:
            self._dispose()

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
        self._closing = True
        self._view.disableButtons()
        self._model.save(self._view.getDialogPosition())
        #self._sender.removeListener(self._senderlistener)
        self._log1.removeModifyListener(self._loglistener1)
        self._log2.removeModifyListener(self._loglistener2)
        self._log1.removeRollerHandler()
        self._log2.removeRollerHandler()
        self._model.close()

    def _dispose(self):
        self._model.dispose()
        self._view.dispose()

