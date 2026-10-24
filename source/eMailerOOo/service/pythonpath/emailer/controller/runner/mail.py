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

from ...dbtool import getObjectSequenceFromResult

from ...unotool import getUriFactory

from ...helper import parseUriFragment

import traceback


class Mail(Job):
    def __init__(self, ctx, spooler, jobs, callback):
        super().__init__(ctx, callback)
        self._spooler = spooler
        self._jobs = jobs
        self._parent = parent

        self._steps = (self._getStepGetJobs(),
                       self._getStepBuildStep())
        self.steps = iter(self._steps)
        self.total = len(self._steps)

    def stepGetJobs(self):
        result = self._spooler.getJobs()
        self._batchs = getObjectSequenceFromResult(result)
        result.close()

    def stepBuildStep(self):
        steps = []
        factory = getUriFactory(self._ctx)
        for batch in self._batchs:

            steps.append(self._getStepInitBatch())
            merge = batch.Merge and batch.DataSource and batch.Table
            if merge:
                steps.append(self._getStepInitDataSource(batch.DataSource))
                steps.append(self._getStepInitConnection())
                steps.append(self._getStepInitRowSet(batch.DataSource, batch.Table))
                steps.append(self._getStepValidateMergeFields())
                steps.append(self._getStepInitDispatcher())
            steps.append(self._getStepInitDocument(batch.Document))

            first = True
            for i in range(len(batch.Jobs)):
                job = batch.Jobs[i]
                filter = batch.Filters[i]
                recipient = batch.Recipients[i]
                num = 1
                sends = [] 
                substeps = []

                if merge:
                    if first:
                        steps.append(self._getStepValidateDocumentFields(job))
                    steps.append(self._getStepInitMerge(job, filter))
                sends.append(self._getStepBuildMail(job, batch.Sender, recipient, batch.Subject))
                for attachment in batch.Attachments:
                    url, merged, export = parseUriFragment(factory, attachment, merge)
                    if first:
                        if export == 'pdf':
                            steps.append(self._getStepInitPdfExport())
                        steps.append(self._getStepInitAttachement(url, num))
                    if merged:
                        if first:
                            steps.append(self._getStepValidateAttachmentFields(job, num))
                            steps.append(self._getStepInitMergeAttachment(job, num, url, export))
                        substeps.append(self._getStepMergeAttachment(job, num, filter))
                    elif export:
                        substeps.append(self._getStepExportAttachment(job, num, export))
                    sends.append(self._getStepAddMailAttachement(job, num))
                    num += 1
                if merge:
                    steps.append(self._getStepMergeDocument(job, filter))
                steps.append(self._getStepExportDocument(job, 'html'))

                steps.extend(substeps)
                steps.extend(sends)
                first = False

        if steps:
            self._buildSteps(steps)
