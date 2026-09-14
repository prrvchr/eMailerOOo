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

from com.sun.star.logging.LogLevel import INFO

from .optionsmodel import OptionsModel

from .optionsview import OptionsView

from .optionshandler import OptionsListener

from ..listener import DispatchListener

from ..setup import SetupCheck

from .logger import LogManager

from ..unotool import executeDesktopDispatch
from ..unotool import getDesktop

from ..setup import showSetup

from ..helper import getExtensions
from ..helper import showController
from ..helper import showIspDB
from ..helper import showSpooler

from ..jdbcdriver import g_java
from ..jdbcdriver import g_script

from ..configuration import State
from ..configuration import g_defaultlog
from ..configuration import g_extension
from ..configuration import g_mailservicelog
from ..configuration import g_spoolerlog

import traceback


class OptionsManager():
    def __init__(self, ctx, logger, window):
        self._ctx = ctx
        self._disposed = False
        self._java = g_extension, g_java, g_script
        self._logger = logger
        self._model = OptionsModel(ctx)
        window.addEventListener(OptionsListener(self))
        self._view = OptionsView(window, *self._model.getViewData())
        self._logmanager = LogManager(ctx, window, 'requirements.txt', g_defaultlog, g_spoolerlog, g_mailservicelog)
        self._logmanager.initView()
        self._logger.logprb(INFO, 'OptionsManager', '__init__', 151)
        self._loadView(self.init)

    def init(self, state):
        if not self._disposed:
            self._view.initView(*self._model.getInitData(state))
            self._logmanager.loadSetting()
            self._logger.logprb(INFO, 'OptionsManager', 'init', 161)

    def refresh(self, state):
        if not self._disposed:
            self._view.refreshView(state)

    def dispose(self):
        self._disposed = True
        self._logmanager.dispose()
        self._view.dispose()

    def loadSetting(self):
        self._loadView(self.init)

    def saveSetting(self):
        option = self._model.saveData(*self._view.getData())
        log = self._logmanager.saveSetting()
        if log:
            State.restart = True
            self._loadView(self.refresh)
        self._logger.logprb(INFO, 'OptionsManager', 'saveSetting', 171, option, log)

    def changeTimeout(self, timeout):
        self._model.setTimeout(timeout)

    def showIspdb(self):
        showIspDB(self._ctx, Parent=self._view.getWindow())

    def showSpooler(self):
        showSpooler(self._ctx, Dispatch=DispatchListener(self))

    def showController(self):
        showController(self._ctx, Dispatch=DispatchListener(self))

    def showDataBase(self):
        url = self._model.getDataSourceUrl()
        getDesktop(self._ctx).loadComponentFromURL(url, '_default', 0, ())

    def showSetup(self):
        mode = self._view.getSetup()
        showSetup(self._ctx, Dispatch=DispatchListener(self), Mode=mode)

    def dispatchFinished(self, notification):
        if not self._disposed:
            self._loadView(self.refresh)

    def _loadView(self, callback):
        mode = self._view.getSetup()
        database = self._model.getDataSourceUrl()
        setup = SetupCheck(self._ctx, callback, database=database, java=self._java, python=True, extensions=getExtensions(), mode=mode)
        setup.start()
