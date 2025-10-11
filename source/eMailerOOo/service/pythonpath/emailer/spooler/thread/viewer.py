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
from .worker import Merger
from .worker import WatchDog

from .task import Document

from ..export import PdfExport

from ...unotool import executeShell
from ...unotool import getSimpleFile

import traceback


class Viewer(Master):
    def __init__(self, ctx, cancel, progress, connection, result, datasource,
                 table, selection, url, merge, filter, notifier, source):
        super().__init__(cancel, progress)
        self._ctx = ctx
        self._cls = 'Viewer'
        sf = getSimpleFile(ctx)
        self._document = Document(ctx, cancel, sf, connection, result, datasource, table, selection, url, merge, filter)
        self._sf = sf
        self._notifier = notifier
        self._source = source
        if filter == 'pdf':
            self._export = PdfExport(ctx)
        else:
            self._export = None
        self.start()

    def run(self):
        mtd = 'run'
        document = self._document
        result = uno.createUnoStruct('com.sun.star.beans.StringPair')
        result.First = document.Filter
        msg = "View document: %s" % document.Name
        self._startProgress(msg, self._getProgressRange())
        self._setProgressValue(-10)
        if not document.hasFile():
            document.close()
            self._setProgressText("Viewer file: %s don't exists" % document.Name)
            self._setProgressValue(-10)
            status, result.Second = FAILURE, "Document: %s don't exists" % document.Url
        else:
            merger = Merger(self._ctx, self._cancel, self._progress, document, self._export)
            watchdog = WatchDog(merger)
            if watchdog.startJob(merger, document):
                merger.join()
            document.close()
            if self.isCanceled():
                print("Viewer.run() viewer canceled")
                self._setProgressText("Viewer file: %s as been canceled" % document.Name)
                self._setProgressValue(-10)
                msg = "Viewer file: %s as been canceled" % document.Name
                status, result.Second = FAILURE, msg
            elif document.hasInvalidFields():
                print("Viewer.run() has invalid fields")
                self._setProgressText("Viewer file: %s as error" % document.Name)
                self._setProgressValue(-10)
                msg = "Document: %s - DataSource: %s -Field: %s - Value: %" % document.getInvalidFields()
                status, result.Second = FAILURE, msg
            else:
                self._setProgressText("Notifying viewing file: %s" % document.Name)
                self._setProgressValue(-10)
                status, result.Second = SUCCESS, document.getUrl(0)
        print("Viewer.run() notifier")
        self._endProgress()
        struct = 'com.sun.star.frame.DispatchResultEvent'
        notification = uno.createUnoStruct(struct, self._source, status, result)
        self._notifier.dispatchFinished(notification)
        print("Viewer.run() normal thread end")

    # Private getter methods
    def _getProgressRange(self):
        # XXX: Viewer + Merger
        return 20 + 20

