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

import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .optionsmodel import OptionsModel

from .optionsview import OptionsView

from .optionshandler import OptionsListener

from ..logger import LogManager

from ..unotool import createService
from ..unotool import executeDispatch
from ..unotool import getDesktop

from ..configuration import g_defaultlog
from ..configuration import g_spoolerlog
from ..configuration import g_mailservicelog

import traceback


class OptionsManager(unohelper.Base):
    def __init__(self, ctx, logger, window):
        self._ctx = ctx
        self._logger = logger
        self._model = OptionsModel(ctx)
        window.addEventListener(OptionsListener(self))
        self._view = OptionsView(window)
        self._view.initView(OptionsManager._restart, *self._model.getViewData())
        self._logmanager = LogManager(ctx, window, 'requirements.txt', g_defaultlog, g_spoolerlog, g_mailservicelog)
        self._logmanager.initView()
        self._logger.logprb(INFO, 'OptionsManager', '__init__', 151)

    _restart = False

    def dispose(self):
        self._logmanager.dispose()
        self._view.dispose()

    def loadSetting(self):
        self._view.initView(OptionsManager._restart, *self._model.getViewData())
        self._logmanager.loadSetting()
        self._logger.logprb(INFO, 'OptionsManager', 'loadSetting', 161)

    def saveSetting(self):
        option = self._model.saveTimeout(self._view.getTimeout())
        log = self._logmanager.saveSetting()
        if log:
            OptionsManager._restart = True
            self._view.setRestart(OptionsManager._restart)
        self._logger.logprb(INFO, 'OptionsManager', 'saveSetting', 171, option, log)

    def changeTimeout(self, timeout):
        self._model.setTimeout(timeout)

    def showIspdb(self):
        executeDispatch(self._ctx, 'emailer:ShowIspdb')
        self._view.updateDataBase(self._model.getDataBaseStatus())

    def showSpooler(self):
        executeDispatch(self._ctx, 'emailer:ShowSpooler')
        self._view.updateDataBase(self._model.getDataBaseStatus())

    def showDataBase(self):
        url = self._model.getDataBaseUrl()
        getDesktop(self._ctx).loadComponentFromURL(url, '_default', 0, ())

