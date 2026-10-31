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

from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.mail.MailServiceType import SMTP

from .mail import Mail

from ..unotool import Resource

from ..dbtool import getObjectSequenceFromResult

from ..unotool import getLastNamedParts
from ..unotool import getResourceLocation

from ..helper import getMailMessage
from ..helper import getMailServer
from ..helper import getMessageImage
from ..helper import parseUriFragment

from ..configuration import g_extension
from ..configuration import g_identifier
from ..configuration import g_img
from ..configuration import g_logo
from ..configuration import g_logourl
from ..configuration import g_sep

import traceback


class Sender(Mail):
    def __init__(self, ctx, spooler, callback, jobs=(), parent=None):
        super().__init__(ctx, callback, spooler, parent=parent)
        self._threadid = None
        self._count = 0
        self._users = {}
        self._server = None
        logo = g_img + g_sep + g_logo
        self._logo = getResourceLocation(ctx, g_identifier, logo)
        self._resources = {1: Resource(20101, lambda x: (x.total, )),
                           2: Resource(20102, lambda x: (x.total, )),
                           3: Resource(20103),
                           4: Resource(20104, lambda x: (x._count, )),
                           5: Resource(20105, lambda x: (x._count, )),
                           6: Resource(20106)}
        self.total = 5
        self.steps = self._getJobSteps(ctx, jobs)

    def skipSteps(self):
        if self._batch is None or self._job is None:
            return 0
        for i in range(len(self._batch.Jobs)):
            if self._job <= self._batch.Jobs[i]:
                self._spooler.setJobState(self._batch.Jobs[i], 2)
        skipped = 0
        while True:
            try:
                step = next(self.steps)
                resource, options, call, *args = step
                if call == self.stepInitBatch:
                    self.steps = self._pushBack(step)
                    return skipped
                skipped += 1
            except StopIteration:
                return 0

    def _pushBack(self, step):
        yield step
        yield from self.steps

    def _getJobSteps(self, ctx, jobs):
        yield self._getStepInitJob(ctx, self.__class__.__name__)
        yield self._getStepGetJobs(jobs)
        yield self._getStepGetUsers(ctx)
        total = self.total
        yield self._getStepGetStepCount()
        if total < self.total:
            yield from self._getStepBuildStep(ctx)
        yield self._getStepFinalize()

    def stepGetJobs(self, jobs=()):
        if jobs:
            result = self._spooler.getJobs(jobs)
        else:
            result = self._spooler.getSendJobs()
        self._batchs = getObjectSequenceFromResult(result)
        result.close()

    def stepGetStepCount(self):
        total = 0
        for batch in self._batchs:
            total += 2
            merge = all((batch.Merge, batch.DataSource, batch.Table))
            if merge:
                total += 3
            user = self._users[batch.Batch]
            imap = user.supportIMAP() and merge
            jobs = len(batch.Jobs)
            if jobs > 0:
                if imap:
                    total += 1
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

            imap = None
            user = self._users[batch.Batch]
            if user.supportIMAP() and merge:
                imap = self._getStepSendThreadMessage(ctx, user, batch)

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

                if first:
                    if imap:
                        yield imap
                    yield self._getStepConnectSmptServer(ctx, user, batch)
                yield self._getStepSendMail(ctx, user, batch, job, recipient, merge)
                first = False

    def stepSendThreadMessage(self, ctx, user, batch):
        threadid = None
        server = getMailServer(ctx, user, IMAP)
        folder = server.getSentFolder()
        if server.hasFolder(folder):
            message = self._getThreadMessage(ctx, batch)
            body = self._transferable.getByString(message)
            email = getMailMessage(ctx, batch.Sender, batch.Sender, batch.Subject, body)
            server.uploadMessage(folder, email)
            # XXX: Some providers use their own reference (ie: Microsoft with their ConversationId)
            # XXX: in order to group the mails and not the MessageId of the first email.
            # XXX: If so, the ThreadId is not empty.
            threadid = email.ForeignId if email.ForeignId else email.MessageId
        server.disconnect()
        self._threadid = threadid

    def stepConnectSmptServer(self, ctx, user):
        self._server = getMailServer(ctx, user, SMTP)

    def stepSendMail(self, ctx, user, batch, job, recipient, merge):
        email = self._buildMailMessage(ctx, user, batch, job, recipient, merge)
        if self._threadid:
            email.ThreadId = self._threadid
        try:
            self._server.sendMailMessage(email)
            self._count += 1
        except UnoException as e:
            self._spooler.setJobState(job, 2)
        else:
            self._spooler.updateJob(batch.Batch, job, recipient, email.ThreadId, email.MessageId, email.ForeignId, 1)

    def _getThreadMessage(self, ctx, batch):
        title = self._resolver.resolveString(12102).format(batch.Batch, batch.Query)
        message = self._resolver.resolveString(12103)
        label = self._resolver.resolveString(12104)
        files = self._resolver.resolveString(12105)
        if batch.Attachments:
            tag = '<a href="%s">%s</a>'
            sep = '</li><li>'
            tags = '<ol><li>%s</li></ol>' % self._getAttachments(batch.Attachments, tag, sep)
        else:
            tags = '<p>%s</p>' % self._resolver.resolveString(12106)
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
    <div style="margin-bottom: 15px;">
      <img alt="%s Logo" src="data:image/png;charset=utf-8;base64,%s" src="%s" style="float: left; margin-right: 10px;" />
      <div style="overflow: hidden;">
        <span style="font-size: 1.5em; font-weight: bold; line-height: 1.2;">%s</span><br />
        <span style="font-size: 1em; color: #333;">%s</span>
      </div>
    </div>
    <div style="clear: both;"></div>
    <p><b>%s:</b>&nbsp;%s</p>
    <p><b>%s:</b>&nbsp;<a href="%s">%s</a></p>
    <p><b>%s:</b></p>
    %s
  </body>
</html>
''' % (g_extension, logo, g_logourl, g_extension, title, message, batch.Subject,
       label, batch.Document, name, files, tags)

    def _getAttachments(self, attachments, tag, sep):
        urls = []
        for attachment in attachments:
            urls.append(tag % getLastNamedParts(attachment, g_sep))
        return sep.join(urls)

