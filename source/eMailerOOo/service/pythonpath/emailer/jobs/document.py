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

from ..unotool import Transferable

from ..unotool import getLastNamedParts
from ..unotool import getUriFactory

from ..configuration import g_sep


class Document(Job):
    def __init__(self, ctx, callback, parent=None):
        super().__init__(ctx, callback, parent=parent)
        self._factory = None
        self._transferable = None

    def callback(self, success=True, error=None):
        self._callback(success, self._urls[0], error)

    def stepInitJob(self, ctx):
        super().stepInitJob(ctx)
        self._factory = getUriFactory(ctx)
        self._transferable = Transferable(ctx)

    def stepGetStepCount(self, merge, export, jobs, email):
        self.total += self._stepGetStepCountDocument(merge, export, jobs, email)

    def _getJobSteps(self, ctx, merge, export, datasource, table, fltr, url, email):
        yield self._getStepInitJob(ctx, self.__class__.__name__)
        yield self._getStepGetStepCount(merge, export, 1, email)
        yield from self._getStepBuildDocument(ctx, merge, export, datasource, table, fltr, url, 0, 0, 0, email, True, True)
        self._tmps.clear()
        yield self._getStepFinalize()

    def _stepGetStepCountDocument(self, merge, export, jobs, email=True):
        if jobs <= 0:
            return 0
        total = 0

        if export == 'pdf':
            total += 1  # yield self._getStepInitPdf
        if email:
            total += 1  # yield self._getStepInitDocument
        else:
            total += 1  # yield self._getStepInitAttachement

        if merge:
            if email:
                total += 1  # yield self._getStepValidateDocumentFields
            else:
                total += 1  # yield self._getStepValidateAttachmentFields
            total += 1      # yield self._getStepInitDataDescriptor

        if merge:
            total += 1 * jobs  # yield self._getStepSetRowsetFilter
            if email:
                total += 1 * jobs  # yield self._getStepMergeDocument
            else:
                total += 1 * jobs  # yield self._getStepMergeAttachment

        total += 1 * jobs  # yield self._getStepExport[Document|Attachment]
        return total

    def _getStepBuildDocument(self, ctx, merge, export, datasource, table, fltr, url, num=0, batch=0, job=0, email=True, init=True, first=True):
        if init and merge:
            yield self._getStepInitDataSource(ctx, datasource, batch)
            yield self._getStepInitConnection(ctx, batch)
            yield self._getStepInitRowSet(ctx, datasource, table, batch)
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
            yield self._getStepInitDataDescriptor(datasource, table, batch)
        if merge:
            yield self._getStepSetRowsetFilter(fltr, batch, job)
            if email:
                yield self._getStepMergeDocument(ctx, title, batch, job)
            else:
                yield self._getStepMergeAttachment(ctx, num, title, batch, job)
        if email:
            yield self._getStepExportDocument(export, title, batch, job)
        else:
            yield self._getStepExportAttachment(num, export, title, batch, job)
