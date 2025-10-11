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
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.ucb.ConnectionMode import ONLINE

from .worker import Dispatcher

from ...unotool import getConnectionMode

from ...dbtool import getObjectSequenceFromResult

from ...helper import getMailSpooler

from ...configuration import g_dns

from queue import Queue
import traceback


class Sender(Dispatcher):
    def __init__(self, ctx, cancel, progress, logger, listeners):
        super().__init__(ctx, cancel, progress, logger)
        self._cls = 'Sender'
        self._listeners = listeners
        self._queue = Queue()
        self._res = 1050
        self._jobs = []
        self.start()

    # Private getter methods
    def _getBatchs(self):
        spooler = getMailSpooler(self._ctx)
        result = spooler.getSendJobs()
        batchs = getObjectSequenceFromResult(result)
        result.close()
        spooler.dispose()
        return batchs

    def _isOffLine(self):
        return getConnectionMode(self._ctx, *g_dns) != ONLINE

    def _notifyClosed(self, value=100):
        super()._notifyClosed(value)
        for listener in self._listeners:
            listener.closed()
        self._cancel.set()

