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

from com.sun.star.awt import XCallback

from com.sun.star.util import CloseVetoException

from .controllermodel import ControllerModel

from .controllerview import ControllerView

from .controllerhandler import CloseListener
from .controllerhandler import RowSetListener
from .controllerhandler import WindowHandler

from ..runner import Send

from ...dbsetup import DBSetup

from ...runner import Cancel
from ...runner import Runner


from ...dbinit import getDataSourceUrl

from ...unotool import createService
from ...unotool import executeDesktopDispatch
from ...unotool import notifyDispatch

from ...setup import SetupCheck
from ...setup import SetupManager

from ...helper import getExtensions
from ...helper import showSpooler

from ...logger import getLogger

from ...jdbcdriver import g_java
from ...jdbcdriver import g_script

from ...configuration import g_extension
from ...configuration import g_setup
from ...configuration import g_spoolerlog

import traceback


class ControllerManager():
    def __init__(self, ctx, source, /, dispatch=None, **kwargs):
        print("ControllerManager.__init__() 1")
        self._ctx = ctx
        self._source = source
        self._dispatch = dispatch
        self._model = ControllerModel(ctx)
        self._init = True
        self._closing = False
        self._running = False
        self._disabled = False
        self._cancel = Cancel()
        self._listener = CloseListener(self)
        self._view = ControllerView(ctx, WindowHandler(self), self._listener, *self._model.getViewData())
        self._logger = getLogger(ctx, g_spoolerlog)
        self._java = g_extension, g_java, g_script
        self._extensions = getExtensions()
        self._check = SetupCheck(ctx, self.notify, database=getDataSourceUrl(ctx), java=self._java, python=True, extensions=self._extensions)
        self._check.start()

    def notify(self, state):
        try:
            if not state.runnable:
                database = DBSetup(self._ctx)
                parent = self._view.getWindow()
                setup = SetupManager(self._ctx, self._source, database=database, java=self._java, python=True, extensions=self._extensions, parent=parent, name=g_setup, state=state)
                if not setup.isRunnable():
                    setup.dispose()
                    self._dispose()
                    return
                setup.dispose()
            self._model.initController(RowSetListener(self))
        except Exception as e:
            print("ControllerManager.notify() ERROR: %s" % traceback.format_exc())

# XCloseListener
    def queryClosing(self, source, ownership):
        print("ControllerManager.queryClosing() ownership: %s" % ownership)
        self._close()
        if self._running:
            raise CloseVetoException()
        if not ownership:
            self._dispose()

    def notifyClosing(self, source):
        print("ControllerManager.queryClosing() ownership: %s" % ownership)
        source.removeCloseListener(self._listener)
        #self._dispose()

# XRowSetListener
    def setDataModel(self, rowset):
        if self._init:
            self._init = False
            self._view.setLabelState(self._model.getLabelState(1))
            self._view.enableButtons(True)
            if self._dispatch:
                notifyDispatch(self._source, self._dispatch)

    def toggleSpooler(self, state):
        print("ControllerManager.toggle() state: %s" % state)
        if self._cancel.isSet() or self._disabled:
            return
        self._view.setToggleState(self._model.getToggleState(state))
        if state:
            spooler = self._model.getSpooler()
            self._view.setLabelState(self._model.getLabelState(2))
            try:
                #sender = Sender(self._ctx, spooler, self.finished, parent, progress, self._logger, self._cancel)
                progress = self._view.getIndicator()
                job = Send(self._ctx, spooler, self.notifyFinished)
                runner = Runner(self._ctx, job, progress, self._cancel)
                self._running = True
                runner.start()
                self._view.enableToggleButon(True)
            except Exception as e:
                print("ControllerManager.ControllerManager() ERROR: %s" % traceback.format_exc())
        else:
            self._view.setLabelState(self._model.getLabelState(0))
            self._cancel.set()

    def notifyFinished(self, success):
        print("ControllerManager.finished() 1")
        self._running = False
        if self._closing:
            print("ControllerManager.finished() 2")
            self._dispose()
        else:
            print("ControllerManager.finished() 3")
            self._cancel.reset()
            self._view.resetIndicator()
            self._disabled = True
            self._view.setToggleButton(0)
            self._disabled = False
            self._view.setLabelState(self._model.getLabelState(1))
            self._view.setToggleState(self._model.getToggleState(0))
            self._view.enableToggleButon(True)
        print("ControllerManager.finished() state: %s" % self._view.getToggleButton())

    def showSpooler(self):
        showSpooler(self._ctx)

    def closeController(self):
        self._close()
        if not self._running:
            self._dispose()

    def _close(self):
        self._cancel.set()
        self._view.enableButtons(False)
        self._model.savePosition(self._view.getWindowPosition())
        self._closing = True
        self._model.close()

    def _dispose(self):
        self._model.dispose()
        self._view.dispose()

