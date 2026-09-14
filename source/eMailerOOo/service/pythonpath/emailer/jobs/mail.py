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

from .job import Job

from ..unotool import Transferable

from ..dbtool import getRowDict

from ..unotool import getLastNamedParts
from ..unotool import getUriFactory

from ..helper import getMailMessage

from ..configuration import g_sep

import traceback


class Mail(Job):
    def __init__(self, ctx, callback, spooler, parent=None):
        super().__init__(ctx, callback, parent=parent)
        self._spooler = spooler
        self._batch = None
        self._job = None
        self._factory = None
        self._transferable = None

    def stepInitJob(self, ctx):
        super().stepInitJob(ctx)
        self._factory = getUriFactory(ctx)
        self._transferable = Transferable(ctx)

    def _buildMailMessage(self, ctx, user, batch, job, recipient, merge):
        subject = self._getSubject(batch, merge)
        url = self._urls[0]
        print("Mail._buildMailMessage() url %s" % url)
        body = self._transferable.getByUrl(url)
        email = getMailMessage(ctx, batch.Sender, recipient, subject, body)
        for i in range(1, len(self._urls)):
            url = self._urls[i]
            path, name = getLastNamedParts(url, g_sep)
            attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
            attachment.Data = self._transferable.getByUrl(url)
            attachment.ReadableName = name
            email.addAttachment(attachment)
        if user.useReplyTo():
            email.ReplyToAddress = user.getReplyToAddress()
        return email

    def _getSubject(self, batch, merge):
        subject = batch.Subject
        if merge:
            resultset = self._rowset.createResultSet()
            if resultset.next():
                subject = subject.format(**getRowDict(resultset))
            resultset.close()
        return subject

