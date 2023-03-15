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

from .optionsmodel import OptionsModel
from .optionsview import OptionsView
from .optionshandler import OptionsListener

from ..logger import LogManager

from ..spooler import SpoolerListener

from ..unotool import executeDispatch

from ..configuration import g_identifier

import os
import sys
import traceback


class OptionsManager(unohelper.Base):
    def __init__(self, ctx, window):
        self._ctx = ctx
        self._model = OptionsModel(ctx)
        self._view = OptionsView(ctx, window, *self._model.getViewData())
        version  = ' '.join(sys.version.split())
        path = os.pathsep.join(sys.path)
        infos = {111: version, 112: path}
        self._logger = LogManager(ctx, window.getPeer(), infos, g_identifier, 'Logger')
        window.addEventListener(OptionsListener(self))
        self._model.addSpoolerListener(SpoolerListener(self))

    def started(self):
        self._view.setSpoolerStatus(*self._model.getSpoolerStatus(1))

    def stopped(self):
        self._view.setSpoolerStatus(*self._model.getSpoolerStatus(0))

    def dispose(self):
        self._logger.dispose()

    def loadSetting(self):
        self._logger.loadSetting()
        self._view.initControl(*self._model.getViewData())

    def saveSetting(self):
        self._logger.saveSetting()
        self._model.saveTimeout()

    def changeTimeout(self, timeout):
        self._model.setTimeout(timeout)

    def showSmtpServer(self):
        try:
            executeDispatch(self._ctx, 'smtp:ispdb')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def toogleSpooler(self, state):
        self._model.toogleSpooler(state)

    def showSmtpSpooler(self):
        try:
            executeDispatch(self._ctx, 'smtp:spooler')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)
