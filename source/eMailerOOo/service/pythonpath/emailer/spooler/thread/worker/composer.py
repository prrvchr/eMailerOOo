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

from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.uno import Exception as UnoException

from .type import Worker

from .sender import Sender

from ....unotool import Transferable

from ....helper import getMailMessage

import traceback


class Composer(Worker):
    def __init__(self, ctx, cls, caller, callback, logger, resource, progress, sf, input, output=None, result=None):
        super().__init__(progress, input, output)
        self._ctx = ctx
        self._cls = cls
        self._caller = caller
        self._callback = callback
        self._logger = logger
        self._resource = resource
        self._sf = sf
        self._result = result
        self._transferable = Transferable(ctx)
        self._sender = self._getSender(self.finished)

    def notify(self, data):
        mtd = 'run'
        if not self.isCanceled() and not self._input.empty():
            mail = self._input.get()
            mail.close()
            if mail.hasInvalidFields():
                self._input.task_done()
                msg = self._logger.resolveString(self._resource + 21, *mail.getInvalidFields())
                self._logger.logp(SEVERE, self._cls, mtd, msg)
                if self._result:
                    msg = self._logger.resolveString(self._resource + 22, mail.Url)
                    self._result.Result = msg
                else:
                    progress = 10 + 20 * mail.JobCount()
                    self._setProgressValue(-progress)
            elif not self.isCanceled():
                self._setProgressValue(-10)
                for job in mail.Jobs:
                    if self.isCanceled():
                        break
                    self._setProgressValue(-10)
                    recipient = mail.getRecipient(job)
                    if recipient is None:
                        msg = self._logger.resolveString(self._resource + 23, job)
                        self._logger.logp(SEVERE, self._cls, mtd, msg)
                        self._setProgressValue(-10)
                        if self._result:
                            msg = self._logger.resolveString(self._resource + 24, job)
                            self._result.Result = msg
                            break
                        else:
                            continue
                    subject = mail.getSubject(job)
                    url = mail.getUrl(job)
                    try:
                        body = self._transferable.getByUrl(url)
                        email = getMailMessage(self._ctx, mail.Sender, recipient, subject, body)
                        for task in mail.Attachments:
                            # XXX: the following tasks are the attachments if they exist
                            attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
                            url = task.getUrl(job)
                            attachment.Data = self._transferable.getByUrl(url)
                            attachment.ReadableName = task.Name
                            email.addAttachment(attachment)
                    except UnoException as e:
                        msg = self._logger.resolveString(self._resource + 25, e.Message, job, recipient)
                        self._logger.logp(SEVERE, self._cls, mtd, msg)
                    else:
                        if self._output is not None:
                            data = job, mail, email
                            self._output.put(data)
                        else:
                            stream = self._sf.openFileWrite(self._result.Result)
                            stream.writeBytes(uno.ByteSequence(email.asBytes()))
                            stream.flush()
                            stream.closeOutput()
                            self._result.State = SUCCESS
                    self._setProgressValue(-10)
            self._input.task_done()

        if self.isCanceled():
            msg = self._logger.resolveString(self._resource + 26)
            self._logger.logp(SEVERE, self._cls, mtd, msg)
        else:
            if self._output and not self._output.empty():
                caller.addCallback(sender, None)
        if not self._input.empty():
            caller.addCallback(self, None)

    def finished(self):
        self._callback()

    def _getSender(self, callback):
        sender = None
        if self._output:
            sender = Sender(self._ctx, self._caller, callback, self._logger, self._resource, self._progress, self._output)
        return sender

