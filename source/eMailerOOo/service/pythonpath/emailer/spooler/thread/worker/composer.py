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

from .type import Worker

from .sender import Sender

from ....transferable import Transferable

from ....helper import getMailMessage

import traceback


class Composer(Worker):
    def __init__(self, ctx, cancel, progress, logger, sf, input, output=None, url=None):
        super().__init__(cancel, progress, input, output)
        self._ctx = ctx
        self._logger = logger
        self._sf = sf
        self._url = url
        self._transferable = Transferable(ctx, logger)

    def run(self):
        sender = self._getSender()
        while self._hasTasks():
            print("Composer.run() 1 waiting for input")
            mail = self._input.get()
            mail.close()
            if mail.hasInvalidFields():
                self._input.task_done()
                self._taskDone()
                print("Composer.run() MissingFields Title: %s - DataSource: %s - Fields: %s - Value: %s" % mail.getInvalidFields())
                continue
            if self.isCanceled():
                self._input.task_done()
                self._taskDone()
                continue
            self._setProgressValue(-10)
            # We are sure that a task has at least one entry
            print("Composer.run() 2 process mail batchid: %s" % mail.BatchId)
            for job in mail.Jobs:
                print("Composer.run() 3 process mail jobid: %s" % job)
                if self.isCanceled():
                    continue
                self._setProgressValue(-10)
                recipient = mail.getRecipient(job)
                subject = mail.getSubject(job)
                url = mail.getUrl(job)
                print("Composer.run() 4 job: %s - url: %s" % (job, url))
                body = self._transferable.getByUrl(url)
                email = getMailMessage(self._ctx, mail.Sender, recipient, subject, body)
                if email:
                    for task in mail.Attachments:
                        # XXX: the following tasks are the attachments if they exist
                        attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
                        url = task.getUrl(job)
                        print("Composer.run() 5 job: %s - url: %s" % (job, url))
                        attachment.Data = self._transferable.getByUrl(url)
                        attachment.ReadableName = task.Name
                        email.addAttachment(attachment)
                    if sender:
                        data = job, mail, email
                        self._output.put(data)
                        sender.addTasks()
                        print("Composer.run() 6 sender.addTask: %s" % job)
                    elif self._url:
                        print("Composer.run() 7 notify")
                        stream = self._sf.openFileWrite(self._url)
                        stream.writeBytes(uno.ByteSequence(email.asBytes()))
                        stream.flush()
                        stream.closeOutput()
                self._setProgressValue(-10)
            self._input.task_done()
            self._taskDone()
        
        
        if self.isCanceled():
            print("Composer.run() aborted thread end")
        else:
            if sender:
                print("Composer.run() start Sender for %s tasks" % sender._tasks)
                sender.start()
                sender.join()
            print("Composer.run() normal thread end")

    def _getSender(self):
        sender = None
        if self._output:
            sender = Sender(self._ctx, self._cancel, self._progress, self._logger, self._output)
        return sender

