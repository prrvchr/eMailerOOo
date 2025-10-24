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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.uno import Exception as UnoException

from .type import Manager

from ....transferable import Transferable

from ....unotool import getResourceLocation

from ....helper import getMailMessage
from ....helper import getMailServer
from ....helper import getMailSpooler
from ....helper import getMailUser
from ....helper import getMessageImage

from ....configuration import g_extension
from ....configuration import g_identifier
from ....configuration import g_logourl
from ....configuration import g_logo

import traceback


class Sender(Manager):
    def __init__(self, ctx, logger, resource, cancel, progress, input):
        super().__init__(cancel, progress, input)
        self._ctx = ctx
        self._cls = 'Sender'
        self._logger = logger
        self._resource = resource
        self._transferable = Transferable(ctx, self._logger)
        logo = '%s/%s' % (g_extension, g_logo)
        self._logo = getResourceLocation(ctx, g_identifier, logo)

    def run(self):
        mtd = 'run'
        spooler = getMailSpooler(self._ctx)
        server = user = batchid = threadid = None
        total = 0
        while self._hasTasks():
            job, mail, email = self._input.get()
            if self.isCanceled():
                msg = self._logger.resolveString(self._resource + 31)
                self._logger.logp(SEVERE, self._cls, mtd, msg)
                break
            recipient = mail.getRecipient(job)
            msg = self._logger.resolveString(self._resource + 32, job, recipient)
            self._logger.logp(INFO, self._cls, mtd, msg)
            self._setProgress(msg, -10)
            if mail.BatchId != batchid:
                user = getMailUser(self._ctx, mail.Sender)
                if user is None:
                    self._input.task_done()
                    self._taskDone()
                    msg = self._logger.resolveString(self._resource + 33, mail.Sender)
                    self._logger.logp(SEVERE, self._cls, mtd, msg)
                    break
                batchid = mail.BatchId
                if server:
                    server.disconnect()
                server = getMailServer(self._ctx, user, SMTP)
                if user.supportIMAP() and mail.Merge:
                    threadid = self._createThread(user, mail)
                else:
                    threadid = None
            if threadid:
                email.ThreadId = threadid
            if user.useReplyTo():
                email.ReplyToAddress = user.getReplyToAddress()
            try:
                server.sendMailMessage(email)
            except UnoException as e:
                spooler.setJobState(job, 2)
                msg = self._logger.resolveString(self._resource + 34, e.Message, job, recipient)
                level = SEVERE
            else:
                total += 1
                spooler.updateJob(mail.BatchId, job, recipient, email.ThreadId, email.MessageId, email.ForeignId, 1)
                msg = self._logger.resolveString(self._resource + 35, job, recipient)
                level = INFO
            self._logger.logp(level, self._cls, mtd, msg)
            self._setProgress(msg, -10)
            self._input.task_done()
            self._taskDone()
            if self.isCanceled():
                msg = self._logger.resolveString(self._resource + 31)
                self._logger.logp(SEVERE, self._cls, mtd, msg)
                break
        if server:
            server.disconnect()
        spooler.dispose()
        msg = self._logger.resolveString(self._resource + 36, total)
        self._logger.logp(INFO, self._cls, mtd, msg)

    def _createThread(self, user, mail):
        threadid = None
        server = getMailServer(self._ctx, user, IMAP)
        folder = server.getSentFolder()
        if server.hasFolder(folder):
            subject = mail.Subject
            message = self._getThreadMessage(mail)
            body = self._transferable.getByString(message)
            email = getMailMessage(self._ctx, mail.Sender, mail.Sender, mail.Subject, body)
            server.uploadMessage(folder, email)
            # XXX: Some providers use their own reference (ie: Microsoft with their ConversationId)
            # XXX: in order to group the mails and not the MessageId of the first email.
            # XXX: If so, the ThreadId is not empty.
            threadid = email.ForeignId if email.ForeignId else email.MessageId
        server.disconnect()
        return threadid

    def _getThreadMessage(self, mail):
        title = self._logger.resolveString(1551, g_extension, mail.BatchId, mail.Query)
        message = self._logger.resolveString(1552)
        label = self._logger.resolveString(1553)
        files = self._logger.resolveString(1554)
        if mail.Attachments:
            tag = '<a href="%s">%s</a>'
            sep = '</li><li>'
            tags = '<ol><li>%s</li></ol>' % self._getAttachments(mail.Attachments, tag, sep)
        else:
            tags = '<p>%s</p>' % self._logger.resolveString(1555)
        logo = getMessageImage(self._ctx, self._logo)
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
''' % (g_extension, logo, g_logourl, title, message, mail.Subject,
       label, mail.Document, mail.Title, files, tags)

    def _getAttachments(self, attachments, tag, sep):
        urls = []
        for attachment in attachments:
            urls.append(tag % (attachment.Url, attachment.Name))
        return sep.join(urls)

