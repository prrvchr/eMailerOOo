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

from com.sun.star.awt.MessageBoxType import WARNINGBOX

from com.sun.star.frame.DispatchResultState import SUCCESS

from .spoolermodel import SpoolerModel

from .spoolerview import SpoolerView

from .spoolerhandler import CloseListener
from .spoolerhandler import WindowHandler
from .spoolerhandler import RowSetListener
from .spoolerhandler import TabHandler
from .spoolerhandler import TabPageListener

from ..logger import LoggerListener

from ..grid import GridSelectionListener

from ..jobs import Mailer

from ..runner import Cancel
from ..runner import Runner

from ..dbsetup import DBSetup

from ..listener import SpoolerListener

from ..logger import LogController

from ..mailer import showMailer

from ..unotool import CallBack

from ..unotool import createMessageBox
from ..unotool import executeShell
from ..unotool import getCallBack
from ..unotool import notifyDispatch

from ..helper import showController

from ..configuration import g_mailservicelog
from ..configuration import g_spoolerlog
from ..configuration import g_basename


class SpoolerManager():
    def __init__(self, ctx, source, /, dispatch=None, **_):
        self._ctx = ctx
        self._source = source
        self._dispatch = dispatch
        self._model = SpoolerModel(ctx)
        self._running = False
        self._notifier = None
        self._disabled = False
        self._cancel = Cancel()
        self._enabled = True
        self._initialize = True
        self._rowlistener = None
        self._spoolerlistener = None
        self._closelistener = CloseListener(self)
        self._view = SpoolerView(ctx, WindowHandler(self), self._closelistener, TabHandler(self), TabPageListener(self), *self._model.getViewData())
        self._loglistener1 = LoggerListener(self.updateLog1)
        self._log1 = LogController(ctx, g_spoolerlog, g_basename, self._loglistener1)
        self._log1.addRollerHandler()
        self._loglistener2 = LoggerListener(self.updateLog2)
        self._log2 = LogController(ctx, g_mailservicelog, g_basename, self._loglistener2)
        self._log2.addRollerHandler()
        setup = Runner(ctx, DBSetup(self._ctx).withCallback(self.notify), self._cancel, self._view.getIndicator())
        setup.start()

    def notify(self, success, _created):
        self._view.resetIndicator()
        if success:
            self._rowlistener = RowSetListener(self)
            self._spoolerlistener = SpoolerListener(self)
            self._model.initSpooler(self._rowlistener, self._spoolerlistener)
            self._view.setToggleState(*self._model.initToggleState())
            self._view.enableToggleButton(True)

    @property
    def HandlerEnabled(self):
        return self._enabled

# XMailSpoolerListener
    def spoolerToggleState(self, spooler, _success):
        if self._cancel.isSet():
            return
        if spooler.isRunning():
            if not self._model.hasIndicator():
                self._model.addSpoolerIndicator(self._view.getIndicator())
            self._view.setToggleState(*self._model.getToggleState(1))
            self._view.enableEmlView(False)
        else:
            self._view.setToggleState(*self._model.getToggleState(0))
            if self._model.hasIndicator():
                self._model.resetIndicator()
                self._view.resetIndicator()
            self._view.enableEmlView(self._model.hasGridSelectedRows())

# XDispatchResultListener
    def dispatchFinished(self, _notification):
        print("SpoolerManager.dispatchFinished() 1")

    def dispatchFinished1(self, notification):
        self._model.endDispatch()
        if not self._cancel.isSet():
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
    def queryClosing(self, _source, ownership):
        print("SpoolerManager.queryClosing() ownership: %s" % ownership)
        if not ownership:
            self._cancel.set(disposed=True)
            self._close()

    def notifyClosing(self, source):
        print("SpoolerManager.notifyClosing()")
        source.removeCloseListener(self._closelistener)

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
        if self._cancel.isSet():
            return
        if not self._model.isGridInitialized():
            self._view.initView(self._model.initGrid(rowset, GridSelectionListener(self), *self._view.getGridData()))
        self._model.setGridData(rowset)
        if self._dispatch:
            notifyDispatch(self._source, self._dispatch)

    def changeGridSelection(self, index, _grid):
        self._enableButtons(index != -1)

    def _enableButtons(self, selected):
        self._view.enableButtons(*self._model.getRowClientInfo(selected))

    def saveGrid(self):
        self._model.saveGrid()

    def addDocument(self):
        showMailer(self._ctx)

    def toggleSpooler(self):
        if self._cancel.isSet():
            return
        if self._model.getSpooler().isRunning():
            self._view.setLabelState(self._model.getLabelState(2))
            self._model.stopSpooler()
        else:
            self._model.startSpooler(self._view.getIndicator())
        self._view.enableToggleButton(True)

    def viewEML(self):
        if self._cancel.isSet():
            return
        self._view.disableButtons()
        self._view.enableToggleButton(False)
        indicator = self._view.getIndicator()
        spooler = self._model.getSpooler()
        jobs = self._model.getSelectedJobs()
        job = Mailer(self._ctx, spooler, self._onViewEML, jobs)
        runner = Runner(self._ctx, job, self._cancel, indicator)
        runner.start()

    def _onViewEML(self, success, urls, error):
        if self._cancel.isSet():
            return
        self._view.resetIndicator()
        if success and error is None:
            for url in urls:
                executeShell(self._ctx, url)
        self._enableButtons(self._model.hasGridSelectedRows())
        self._view.enableToggleButton(True)

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
        self._cancel.set(disposed=True)
        self._close()

    def updateLog1(self):
        self._view.updateLog1(*self._log1.getLogContent(True))

    def updateLog2(self):
        self._view.updateLog2(*self._log2.getLogContent(True))

# SpoolerManager private methods
    def _close(self):
        self._view.disableButtons()
        self._model.save(self._view.getWindowPosSize())
        self._model.disposeSpooler(self._rowlistener, self._spoolerlistener)
        self._log1.removeModifyListener(self._loglistener1)
        self._log2.removeModifyListener(self._loglistener2)
        self._log1.removeRollerHandler()
        self._log2.removeRollerHandler()
        getCallBack(self._ctx).addCallback(CallBack(self._view.close), None)
