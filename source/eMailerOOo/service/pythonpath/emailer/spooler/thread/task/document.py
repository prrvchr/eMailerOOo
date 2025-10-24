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

from .job import Job

from ...export import PdfExport

from ....helper import executeExport
from ....helper import executeMerge
from ....helper import getJob

from ....unotool import getTempFile

import traceback


class Document(Job):
    def __init__(self, ctx, cancel, sf, connection, result, datasource, table, url, merge, filter, selection):
        super().__init__(sf, url, getTempFile(ctx).Uri, merge, filter)
        self._ctx = ctx
        self._exists = sf.exists(url)
        if self._exists:
            if filter == 'pdf':
                export = PdfExport(ctx)
            else:
                export = None
            self._export = export
            self._job = getJob(ctx, connection, datasource, table, result, export, selection)
        else:
            self._export = self._job = None
        self._rows = {0: (0,)}
        self._fields = None

    @property
    def JobCount(self):
        return 1

    def hasFile(self):
        return self._exists

    def hasInvalidFields(self):
        return self._fields is not None

    def getInvalidFields(self):
        return self._fields

    def getJob(self):
        return self._job

    def execute(self, cancel):
        if not cancel.isSet():
            if self._job:
                self._fields = executeMerge(self._job, self)
            elif self._filter:
                executeExport(self._ctx, self, self._export)

    def close(self):
        # XXX: If we want to avoid a memory dump when exiting LibreOffice,
        # XXX: it is imperative to close / dispose all these references.
        if self._job:
            result = self._job.getPropertyValue('ResultSet')
            if result:
                result.close()
            self._job.dispose()

    def getUrl(self, job):
        if self._rows and job in self._rows:
            url = self._url % self._rows[job][0]
        else:
            url = self._url
        return url

