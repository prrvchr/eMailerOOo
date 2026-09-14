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

from .mail import Mail

from ..unotool import Resource

from ..dbtool import getObjectSequenceFromResult

from ..unotool import getLastNamedParts
from ..unotool import getTempFile

from ..helper import parseUriFragment

from ..configuration import g_sep

import traceback


class Mailer(Mail):
    def __init__(self, ctx, spooler, callback, jobs=(), parent=None):
        super().__init__(ctx, callback, spooler, parent=parent)
        self._results = []
        self._resources = {1: Resource(30101, lambda x: (x.total, )),
                           2: Resource(30102, lambda x: (x.total, )),
                           3: Resource(30103),
                           4: Resource(30104, lambda x: (len(x._results), )),
                           5: Resource(30105, lambda x: (len(x._results), )),
                           6: Resource(30106)}
        self.total = 5
        self.steps = self._getJobSteps(ctx, jobs)

    def callback(self, success=True, error=None):
        self._callback(success, self._results, error)

    def _getJobSteps(self, ctx, jobs):
        yield self._getStepInitJob(ctx, self.__class__.__name__)
        yield self._getStepGetJobs(jobs)
        yield self._getStepGetUsers(ctx)
        total = self.total
        yield self._getStepGetStepCount()
        if total < self.total:
            yield from self._getStepBuildStep(ctx)
        yield self._getStepFinalize()

    def stepGetJobs(self, jobs):
        result = self._spooler.getJobs(jobs)
        self._batchs = getObjectSequenceFromResult(result)
        result.close()

    def stepGetStepCount(self):
        total = 0
        for batch in self._batchs:
            total += 2
            merge = all((batch.Merge, batch.DataSource, batch.Table))
            if merge:
                total += 3
            jobs = len(batch.Jobs)
            if jobs > 0:
                total += 1
                if merge:
                    total += 2
                    total += 2 * jobs
                total += jobs
                total += jobs
            if batch.Attachments:
                for attachment in batch.Attachments:
                    url, merged, export = parseUriFragment(self._factory, attachment, merge)
                    if jobs > 0:
                        if export == 'pdf':
                            total += 1
                        total += 1
                        if merged:
                            total += 1
                            total += jobs
                        total += jobs
        self.total += total

    def _getStepBuildStep(self, ctx):
        for batch in self._batchs:
            self._batch = batch
            yield self._getStepInitBatch(batch.Batch)
            merge = all((batch.Merge, batch.DataSource, batch.Table))
            if merge:
                yield self._getStepInitDataSource(ctx, batch.Batch, batch.DataSource)
                yield self._getStepInitConnection(ctx, batch.Batch)
                yield self._getStepInitRowSet(ctx, batch.Batch, batch.DataSource, batch.Table)
            _, title = getLastNamedParts(batch.Document, g_sep)
            yield self._getStepInitDocument(ctx, batch.Batch, batch.Document)

            user = self._users[batch.Batch]

            first = True
            for i in range(len(batch.Jobs)):
                job = batch.Jobs[i]
                self._job = job
                filter = batch.Filters[i]
                recipient = batch.Recipients[i]

                if first and merge:
                    yield self._getStepValidateDocumentFields(batch.Batch, batch.Table, title)
                    yield self._getStepInitDataDescriptor(batch.Batch, batch.DataSource, batch.Table)

                if merge:
                    yield self._getStepSetRowsetFilter(batch.Batch, job, filter)
                    yield self._getStepMergeDocument(ctx, batch.Batch, job, title)
                yield self._getStepExportDocument(batch.Batch, job, 'html', title)

                num = 1
                for attachment in batch.Attachments:
                    url, merged, export = parseUriFragment(self._factory, attachment, merge)
                    _, name = getLastNamedParts(url, g_sep)
                    if first:
                        if export == 'pdf':
                            yield self._getStepInitPdf(ctx, batch.Batch)
                        yield self._getStepInitAttachement(ctx, batch.Batch, url, num)
                    if merged:
                        if first:
                            yield self._getStepValidateAttachmentFields(batch.Batch, batch.Table, num, name)
                        yield self._getStepMergeAttachment(ctx, batch.Batch, job, num, name)
                    yield self._getStepExportAttachment(batch.Batch, job, num, export, name)
                    num += 1

                yield self._getStepBuildMail(ctx, user, batch, job, recipient, merge, title)
                first = False

    def stepBuildMail(self, ctx, user, batch, job, recipient, merge, title):
        email = self._buildMailMessage(ctx, user, batch, job, recipient, merge)
        folder = getTempFile(ctx).Uri
        print("Mailer.stepBuildMail() url %s - exists: %s" % (folder, self._sf.exists(folder)))
        if not self._sf.exists(folder):
            self._sf.createFolder(folder)
        name, _ = getLastNamedParts(title)
        url = folder + g_sep + name + '.eml'
        stream = self._sf.openFileWrite(url)
        stream.writeBytes(email.asBytes())
        #stream.flush()
        stream.closeOutput()
        self._results.append(url)

