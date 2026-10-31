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

from ..unotool import getUriFactory

from ..helper import parseUriFragment

import traceback


class Document(Job):
    def __init__(self, ctx, callback, /, connection=None, table=None, filter=None, document=None, export='', url=None):
        super().__init__(ctx, callback, connection, document)
        self._ownership = False
        self._table = table
        self._filter = filter
        self._export = export
        self._url = url

        self._steps = (self._getStepBuildStep())
        self.steps = iter(self._steps)
        self.total = len(self._steps)

    def skipSteps(self):
        return 0

    def callback(self, callback, success, error=None):
        url = self._urls[0] if success else None
        callback(success, url)

    def stepBuildStep(self):
        if self._url is None and self._document is None:
            return

        steps = []
        factory = None
        export = self._export
        merge = self._connection and self._table and self._filter
        if self._url:
            factory = getUriFactory(self._ctx)
            self._url, merge, export = parseUriFragment(factory, self._url, merge, export)
        if merge:
            steps.append(self._getStepInitRowSet(self._Connection.DataSource.Name, self._table))
            steps.append(self._getStepInitDispatcher())
        if export == 'pdf':
            steps.append(self._getStepInitPdfExport())
        if factory:
            steps.append(self._getStepInitDocument(self._url))
        if merge:
            steps.append(self._getStepValidateDocumentFields(1))
            steps.append(self._getStepMergeDocument(1, self._filter, export))
        steps.append(self._getStepExportDocument(1, export))

        if steps:
            self._buildSteps(steps)

