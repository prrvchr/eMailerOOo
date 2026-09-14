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

from .document import Document

from ..dbtool import getRowDict

from ..unotool import getLastNamedParts

from ..helper import getMailMessage
from ..helper import getMailUser

from ..configuration import g_sep


class Mail(Document):
    def __init__(self, ctx, callback, spooler, parent=None):
        super().__init__(ctx, callback, parent)
        self._spooler = spooler
        self._batchs = None
        self._users = None

    def stepGetUsers(self, ctx):
        users = {}
        for batch in self._batchs:
            users[batch.Batch] = getMailUser(ctx, batch.Sender)
        self._users = users

    def stepGetStepCount(self, *args):
        self.total += sum(1 for _ in self._getStepBuildStep(*args))

    def _getStepBuildStep(self, ctx):
        raise NotImplementedError('Need to be implemented!')

    def _buildMailMessage(self, ctx, user, recipient, merge, batch):
        subject = self._getSubject(merge, batch)
        url = self._urls[0]
        body = self._transferable.getByUrl(url)
        email = getMailMessage(ctx, batch.Sender, recipient, subject, body)
        for i in range(1, len(self._urls)):
            url = self._urls[i]
            _, name = getLastNamedParts(url, g_sep)
            attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
            attachment.Data = self._transferable.getByUrl(url)
            attachment.ReadableName = name
            email.addAttachment(attachment)
        if user.useReplyTo():
            email.ReplyToAddress = user.getReplyToAddress()
        return email

    def _getSubject(self, merge, batch):
        subject = batch.Subject
        if merge:
            resultset = self._rowset.createResultSet()
            if resultset.next():
                subject = subject.format(**getRowDict(resultset))
            resultset.close()
        return subject
