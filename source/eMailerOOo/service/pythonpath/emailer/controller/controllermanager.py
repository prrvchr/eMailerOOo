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
from .controllerhandler import WindowHandler

from ..dbsetup import DBSetup

from ..jobs import Sender

from ..listener import SpoolerListener

from ..unotool import CallBack

from ..dbinit import getDataSourceUrl

from ..unotool import createService
from ..unotool import executeDesktopDispatch
from ..unotool import getCallBack
from ..unotool import notifyDispatch

from ..setup import SetupCheck
from ..setup import SetupManager

from ..helper import getExtensions
from ..helper import showSpooler

from ..logger import getLogger

from ..jdbcdriver import g_java
from ..jdbcdriver import g_script

from ..configuration import g_extension
from ..configuration import g_setup
from ..configuration import g_spoolerlog

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
        self._disabled = False
        self._spoolerlistener = None
        self._closelistener = CloseListener(self)
        self._view = ControllerView(ctx, WindowHandler(self), self._closelistener, *self._model.getViewData())
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
                    self._closing = True
                    self._close()
                    return
                setup.dispose()
            if not self._closing:
                self._spoolerlistener = SpoolerListener(self)
                self._model.initSpooler(self._spoolerlistener)
                self._view.setToggleState(*self._model.initToggleState())
                self._view.enableButtons(True)
            print("ControllerManager.notify()")

        except Exception as e:
            print("ControllerManager.notify() ERROR: %s" % traceback.format_exc())

# XCloseListener
    def queryClosing(self, source, ownership):
        print("ControllerManager.queryClosing() ownership: %s" % ownership)
        if not ownership:
            self._closing = True
            self._close()

    def notifyClosing(self, source):
        print("ControllerManager.notifyClosing() 1")
        source.removeCloseListener(self._closelistener)
        print("ControllerManager.notifyClosing() 2")

# XMailSpoolerListener
    def spoolerToggleState(self, spooler, success):
        print("ControllerManager.spoolerStarted()")
        if self._closing:
            return
        if spooler.isRunning():
            if not self._model.hasIndicator():
                self._model.addSpoolerIndicator(self._view.getIndicator())
            self._view.setToggleState(*self._model.getToggleState(1))
        else:
            self._view.setToggleState(*self._model.getToggleState(0, reset=True))
            self._view.resetIndicator()
        print("ControllerManager.spoolerStarted()")

    def toggleSpooler(self):
        try:
            print("ControllerManager.toggle() toggleSpooler: %s" % self._disabled)
            #if self._disabled:
            #    self._disabled = False
            #    self._view.enableToggleButon(True)
            #    return
            if self._model.getSpooler().isRunning():
                self._view.setLabelState(self._model.getLabelState(2))
                self._model.stopSpooler()
            else:
                self._model.startSpooler(self._view.getIndicator())
            self._view.enableToggleButon(True)
        except Exception as e:
            print("ControllerManager.toggleSpooler() ERROR: %s" % traceback.format_exc())

    def showSpooler(self):
        showSpooler(self._ctx)

    def closeController(self):
        self._closing = True
        self._close()

    def _close(self):
        print("ControllerManager._close()")
        self._view.enableButtons(False)
        self._model.savePosition(self._view.getWindowPosition())
        self._model.disposeSpooler(self._spoolerlistener)
        getCallBack(self._ctx).addCallback(CallBack(self._view.close), None)

