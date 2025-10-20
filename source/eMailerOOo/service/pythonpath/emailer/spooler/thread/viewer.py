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

from com.sun.star.frame.DispatchResultState import FAILURE
from com.sun.star.frame.DispatchResultState import SUCCESS

from .worker import Master
from .worker import Executor
from .worker import WatchDog

from .task import Document

from ..export import PdfExport

from ...logger import getLogger

from ...unotool import getSimpleFile

from ...configuration import g_spoolerlog

from time import sleep
import traceback


class Viewer(Master):
    def __init__(self, ctx, cancel, progress, connection, result, datasource,
                 table, selection, url, merge, filter, notifier, source):
        super().__init__(cancel, progress)
        self._ctx = ctx
        sf = getSimpleFile(ctx)
        self._document = Document(ctx, cancel, sf, connection, result, datasource, table, selection, url, merge, filter)
        self._sf = sf
        self._notifier = notifier
        self._source = source
        if filter == 'pdf':
            self._export = PdfExport(ctx)
        else:
            self._export = None
        self._logger = getLogger(ctx, g_spoolerlog)
        self._resource = 900
        self._wait = 1.0
        self.start()

    def run(self):
        document = self._document
        result = uno.createUnoStruct('com.sun.star.beans.StringPair')
        result.First = document.Filter
        progress = self._getProgressRange()
        msg = self._logger.resolveString(self._resource + 1)
        self._startProgress(msg, progress)
        if not document.hasFile():
            document.close()
            msg = self._logger.resolveString(self._resource + 2, document.Url)
            self._setProgress(msg, progress)
            msg = self._logger.resolveString(self._resource + 3, document.Url)
            sleep(self._wait)
            status, result.Second = FAILURE, msg
        else:
            executor = Executor(self._logger, self._resource, self._cancel, self._progress, document)
            watchdog = WatchDog(executor)
            if watchdog.startExecutor(executor, document):
                executor.join()
            else:
                msg = self._logger.resolveString(self._resource + 21)
                self._setProgress(msg, -20)
            document.close()
            if self.isCanceled():
                msg = self._logger.resolveString(self._resource + 21)
                self._setProgress(msg, progress)
                status, result.Second = FAILURE, msg
            elif document.hasInvalidFields():
                msg = self._logger.resolveString(self._resource + 22)
                self._setProgress(msg, progress)
                msg = self._logger.resolveString(self._resource + 23, *document.getInvalidFields())
                sleep(self._wait)
                status, result.Second = FAILURE, msg
            else:
                msg = self._logger.resolveString(self._resource + 24)
                self._setProgress(msg, progress)
                sleep(self._wait)
                status, result.Second = SUCCESS, document.getUrl(0)
        self._endProgress()
        struct = 'com.sun.star.frame.DispatchResultEvent'
        notification = uno.createUnoStruct(struct, self._source, status, result)
        self._notifier.dispatchFinished(notification)
        print("Viewer.run() normal thread end")

    # Private getter methods
    def _getProgressRange(self):
        # XXX: Viewer + Executor
        return 20 + 20

