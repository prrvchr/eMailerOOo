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

from ..unotool import getLastNamedParts

from ..configuration import g_sep


class Document(Job):
    def __init__(self, ctx, callback, parent=None):
        super().__init__(ctx, callback, parent)

    def stepGetStepCount(self, *args):
        self.total += sum(1 for _ in self._getStepBuildDocument(*args))
        pass

    def _getJobSteps(self, ctx, url, merge, export, datasource, table, fltr, email):
        yield self._getStepInitJob(ctx, self.__class__.__name__)
        yield self._getStepGetStepCount(ctx, merge, export, datasource, table, fltr, url, 0, 0, 0, email, True, True, True)
        _, title = getLastNamedParts(url, g_sep)
        yield from self._getStepBuildDocument(ctx, merge, export, datasource, table, fltr, url, 0, 0, 0, email, True, True, True)
        yield self._getStepCloneDocument(title, 0, 0)
        yield self._getStepFinalize()

    def _getStepBuildDocument(self, ctx, merge, export, datasource, table, fltr,
                              url, num=0, batch=0, job=0, email=True, init=True, first=True, new=False):
        if init and merge:
            yield self._getStepInitDataSource(ctx, datasource, batch)
            yield self._getStepInitConnection(ctx, batch)
            yield self._getStepInitRowSet(ctx, datasource, table, batch)
            yield self._getStepInitDataDescriptor(datasource, table, batch)
        if first and export == 'pdf':
            yield self._getStepInitPdf(ctx, batch)

        _, title = getLastNamedParts(url, g_sep)

        if first:
            if email:
                yield self._getStepInitDocument(ctx, url, batch)
            else:
                yield self._getStepInitAttachement(ctx, url, num, batch)

        if first and merge:
            if email:
                yield self._getStepValidateDocumentFields(datasource, table, title, batch)
            else:
                yield self._getStepValidateAttachmentFields(datasource, table, num, title, batch)

        if merge:
            if new:
                yield self._getStepSetRowsetFilter(fltr, batch, job)
            if email:
                yield self._getStepMergeDocument(ctx, title, batch, job)
            else:
                yield self._getStepMergeAttachment(ctx, num, title, batch, job)
        if email:
            yield self._getStepExportDocument(export, title, batch, job)
        else:
            yield self._getStepExportAttachment(num, export, title, batch, job)
