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

from .task import Document

from ...export import PdfExport

from ...unotool import getSimpleFile

from ...logger import getLogger

from ...configuration import g_spoolerlog

from time import sleep
import traceback


class Viewer(Master):
    def __init__(self, ctx, caller, progress, connection, result, datasource,
                 table, selection, url, merge, filter, notifier, source, callback):
        super().__init__(progress)
        self._ctx = ctx
        self._caller = caller
        self._callback = callback
        sf = getSimpleFile(ctx)
        self._document = Document(ctx, sf, connection, result, datasource, table, selection, url, merge, filter)
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

    def notify(self, data):
        document = self._document
        result = uno.createUnoStruct('com.sun.star.beans.StringPair')
        result.First = document.Filter
        progress = self._getProgressRange()
        msg = self._logger.resolveString(901)
        self._startProgress(msg, progress)
        if not document.hasFile():
            document.close()
            msg = self._logger.resolveString(902, document.Url)
            self._setProgress(msg, progress)
            msg = self._logger.resolveString(903, document.Url)
            sleep(self._wait)
            status, result.Second = FAILURE, msg
        elif document.hasInvalidFields():
            msg = self._logger.resolveString(922)
            self._setProgress(msg, progress)
            msg = self._logger.resolveString(923, *document.getInvalidFields())
            sleep(self._wait)
            status, result.Second = FAILURE, msg
        elif not self.isCanceled():
            self._caller.addCallback(self._getExecutor(document), None)
            msg = self._logger.resolveString(924)
            self._setProgress(msg, progress)
            sleep(self._wait)
            status, result.Second = SUCCESS, document.getUrl(0)
        else:
            msg = self._logger.resolveString(921)
            self._setProgress(msg, progress)
            status, result.Second = FAILURE, msg
        self._endProgress()
        callBack(status, result)

    # Private getter methods
    def _getExecutor(self, document):
        return Executor(self._logger, self._resource, self._progress, document)

    def _getProgressRange(self):
        # XXX: Viewer + Executor
        return 20 + 20

