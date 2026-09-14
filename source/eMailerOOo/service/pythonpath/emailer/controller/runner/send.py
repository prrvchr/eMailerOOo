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

from ...unotool import getLastNamedParts
from ...unotool import getResourceLocation
from ...unotool import getUriFactory

from ...helper import getMessageImage
from ...helper import parseUriFragment

from ...configuration import g_extension
from ...configuration import g_identifier
from ...configuration import g_img
from ...configuration import g_logo
from ...configuration import g_sep

import traceback


class Send(Job):
    def __init__(self, ctx, spooler, callback, parent=None):
        super().__init__(ctx, callback, parent=parent)
        self._spooler = spooler
        self._factory = getUriFactory(ctx)
        self._users = {}
        self._batch = None
        logo = g_img + g_sep + g_logo
        self._logo = getResourceLocation(ctx, g_identifier, logo)
        self.total = 4
        self.steps = self._getJobStep(ctx)

    def skipSteps(self):
        if self._batch is None:
            return 0
        for i in range(len(self._batch.Jobs)):
            job = self._batch.Jobs[i]
            self._spooler.setJobState(job, 2)
        skipped = 0
        while True:
            try:
                step = next(self.steps)
                resource, options, call, kwargs = step
                if call == self.stepInitBatch:
                    self.steps = self._pushBack(step)
                    return skipped
                skipped += 1
            except StopIteration:
                return 0

    def _pushBack(self, step):
        yield step
        yield from self.steps

    def _getJobStep(self, ctx):
        yield self._getStepGetJobs()
        yield self._getStepGetUsers()
        yield self._getStepGetStepCount()
        yield from self._getStepBuildStep(ctx)
        yield self._getStepFinalize()

    def stepGetJobs(self):
        result = self._spooler.getSendJobs()
        self._batchs = getObjectSequenceFromResult(result)
        result.close()

    def stepGetStepCount(self):
        total = 0
        for batch in self._batchs:
            total += 2
            merge = all((batch.Merge, batch.DataSource, batch.Table))
            if merge:
                total += 4
            imap = False
            user = self._users[batch.Batch]
            if user.supportIMAP() and merge:
                imap = True
            first = True
            for i in range(len(batch.Jobs)):
                if first and imap:
                    total += 1
                if merge:
                    if first:
                        total += 1
                    total += 1
                total += 1
                for attachment in batch.Attachments:
                    url, merged, export = parseUriFragment(self._factory, attachment, merge)
                    if first:
                        if export == 'pdf':
                            total += 1
                        total += 1
                    if merged:
                        if first:
                            total += 2
                        total += 1
                    elif export:
                        total += 1
                    total += 1
                total += 1
                if merge:
                    total += 1
                total += 1
                first = False
        self.total += total

    def stepBuildStep(self, ctx):
        for batch in self._batchs:
            self._batch = batch

            yield self._getStepInitBatch()
            merge = all((batch.Merge, batch.DataSource, batch.Table))
            if merge:
                yield self._getStepInitDataSource(ctx, batch.DataSource)
                yield self._getStepInitConnection(ctx)
                yield self._getStepInitRowSet(ctx, batch.DataSource, batch.Table)
                yield self._getStepInitDispatcher(ctx)
            yield self._getStepInitDocument(ctx, batch.Document)

            user = self._users[batch.Batch]
            imap = None
            if user.supportIMAP() and merge:
                imap = self._getStepBuildThreadMessage(ctx, batch)

            first = True
            for i in range(len(batch.Jobs)):
                job = batch.Jobs[i]
                filter = batch.Filters[i]
                recipient = batch.Recipients[i]
                num = 1
                sends = [] 
                substeps = []
                if first and imap:
                    sends.append(imap)

                if merge:
                    if first:
                        yield self._getStepValidateDocumentFields(job)
                    yield self._getStepInitMerge(job, filter)
                sends.append(self._getStepBuildMail(job, batch.Sender, recipient, batch.Subject))
                for attachment in batch.Attachments:
                    url, merged, export = parseUriFragment(self._factory, attachment, merge)
                    if first:
                        if export == 'pdf':
                            yield self._getStepInitPdfExport()
                        yield self._getStepInitAttachement(job, url, num)
                    if merged:
                        if first:
                            yield self._getStepValidateAttachmentFields(job, num)
                            yield self._getStepInitMergeAttachment(job, num, url, export)
                        substeps.append(self._getStepMergeAttachment(job, num, filter))
                    elif export:
                        substeps.append(self._getStepExportAttachment(job, num, export))
                    sends.append(self._getStepAddMailAttachement(job, num))
                    num += 1
                sends.append(self._getStepSendMail(job, recipient))
                if merge:
                    yield self._getStepMergeDocument(job, filter)
                yield self._getStepExportDocument(job, 'html')

                for step in substeps:
                    yield step
                for send in sends:
                    yield send
                first = False

    def stepBuildThreadMessage(self, ctx, batch):
        title = self.resolver.resolveString(12102, g_extension, batch.BatchId, batch.Query)
        message = self.resolver.resolveString(12103)
        label = self.resolver.resolveString(12104)
        files = self.resolver.resolveString(12105)
        if batch.Attachments:
            tag = '<a href="%s">%s</a>'
            sep = '</li><li>'
            tags = '<ol><li>%s</li></ol>' % self._getAttachments(batch.Attachments, tag, sep)
        else:
            tags = '<p>%s</p>' % self.resolver.resolveString(12106)
        path, name = getLastNamedParts(batch.Document, g_sep)
        logo = getMessageImage(ctx, self._logo)
        return '''\
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body>
    <img alt="%s Logo" src="data:image/png;charset=utf-8;base64,%s" src="%s" />
    <h3 style="display:inline;" >&nbsp;%s</h3>
    <p><b>%s:</b>&nbsp;%s</p>
    <p><b>%s:</b>&nbsp;<a href="%s">%s</a></p>
    <p><b>%s:</b></p>
    %s
  </body>
</html>
''' % (g_extension, logo, g_logourl, title, message, batch.Subject,
       label, batch.Document, name, files, tags)

    def _getAttachments(self, attachments, tag, sep):
        urls = []
        for attachment in attachments:
            urls.append(tag % (attachment.Url, attachment.Name))
        return sep.join(urls)

