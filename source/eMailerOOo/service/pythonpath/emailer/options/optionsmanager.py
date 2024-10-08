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

from .optionsmodel import OptionsModel

from .optionsview import OptionsView

from .optionshandler import OptionsListener

from ..logger import LogManager

from ..spooler import StreamListener

from ..unotool import executeFrameDispatch
from ..unotool import getDesktop

from ..configuration import g_identifier
from ..configuration import g_defaultlog
from ..configuration import g_spoolerlog
from ..configuration import g_mailservicelog

import traceback


class OptionsManager(unohelper.Base):
    def __init__(self, ctx, window, logger):
        self._ctx = ctx
        self._logger = logger
        self._model = OptionsModel(ctx)
        window.addEventListener(OptionsListener(self))
        self._view = OptionsView(window)
        self._view.initView(*self._model.getViewData(OptionsManager._restart))
        self._logmanager = LogManager(ctx, window.getPeer(), 'requirements.txt', g_defaultlog, g_spoolerlog, g_mailservicelog)
        self._listener = StreamListener(self)
        self._model.addStreamListener(self._listener)
        self._logger.logprb(INFO, 'OptionsManager', '__init__()', 151)

    _restart = 0

    def started(self):
        self._view.setSpoolerStatus(*self._model.getSpoolerStatus(1))

    def closed(self):
        self._view.setSpoolerStatus(*self._model.getSpoolerStatus(0))

    def terminated(self):
        self._view.setSpoolerStatus(*self._model.getSpoolerStatus(0))

    def error(self, e):
        self._view.setOptionMessage(e.Message)

    def dispose(self):
        self._logmanager.dispose()
        self._model.removeStreamListener(self._listener)

    def loadSetting(self):
        self._view.initView(*self._model.getViewData(OptionsManager._restart))
        self._logmanager.loadSetting()
        self._logger.logprb(INFO, 'OptionsManager', 'loadSetting()', 161)

    def saveSetting(self):
        option = self._model.saveTimeout(self._view.getTimeout())
        log = self._logmanager.saveSetting()
        if log:
            OptionsManager._restart = 1
            self._view.setOptionMessage(self._model.getRestartMessage(1))
        self._logger.logprb(INFO, 'OptionsManager', 'saveSetting()', 171, option, log)

    def changeTimeout(self, timeout):
        self._model.setTimeout(timeout)

    def showIspdb(self):
        frame = getDesktop(self._ctx).getCurrentFrame()
        if frame is not None:
            executeFrameDispatch(self._ctx, frame, 'smtp:ispdb')
            self._view.updateDataBase(self._model.getDataBaseStatus())

    def toogleSpooler(self, state):
        self._model.toogleSpooler(state)
        if state:
            self._view.clearOptionMessage()

    def showSpooler(self):
        frame = getDesktop(self._ctx).getCurrentFrame()
        if frame is not None:
            executeFrameDispatch(self._ctx, frame, 'smtp:spooler')
            self._view.updateDataBase(self._model.getDataBaseStatus())

    def showDataBase(self):
        url = self._model.getDataBaseUrl()
        getDesktop(self._ctx).loadComponentFromURL(url, '_default', 0, ())

