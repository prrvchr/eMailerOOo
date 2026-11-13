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

from com.sun.star.uno import Exception as UnoException

from ...runner import Task

from ...unotool import Resource
from ...unotool import Transferable

from ...helper import getMailService
from ...helper import getMailMessage

from ...unotool import hasInterface


class Send(Task):
    def __init__(self, ctx, callback, user, service, sender, recipient, subject, message, thread=None):
        super().__init__(ctx, 'Ispdb', callback)
        self._transferable = None
        self._server = None
        self._mail = None
        self._success = False
        self._resources = {1: Resource(30101, lambda x: (x.total, )),
                           2: Resource(30102, lambda x: (x.total, )),
                           3: Resource(30103),
                           4: Resource(30104),
                           5: Resource(30105),
                           6: Resource(30106)}
        self.step = 5 if service == SMTP and user.useIMAP() else 0
        self.total = 10 if user.useIMAP() else 5
        self.steps = self._getJobSteps(ctx, user, service, sender, recipient, subject, message, thread)

    def callback(self, success=True, error=None):
        self._callback(success and self._success, self._mail, error)

    def stepInitJob(self, ctx, service):
        self._transferable = Transferable(ctx)
        self._server = getMailService(ctx, service)

    def stepBuildMailMessage(self, ctx, sender, recipient, subject, message):
        body = self._transferable.getByString(message)
        self._mail = getMailMessage(ctx, sender, recipient, subject, body)

    def stepGetConnection(self, user, service):
        context = user.getConnectionContext(service.value)
        authenticator = user.getAuthenticator(service.value)
        self._server.connect(context, authenticator)

    def stepSendMailMessage(self, service, thread):
        if service == IMAP:
            self._sendImapMessage()
        else:
            self._sendSmtpMessage(thread)

    def stepFinalize(self):
        if self._server and self._server.isConnected():
            self._server.disconnect()
        self._server = None

    def _getStepInitJob(self, ctx, service, name):
        return 30201, (service.value, name), self.stepInitJob, ctx, service

    def _getStepBuildMailMessage(self, ctx, service, sender, recipient, subject, message):
        return 30301, (service.value, ), self.stepBuildMailMessage, ctx, sender, recipient, subject, message

    def _getStepGetConnection(self, user, service):
        return 30401, (service.value, ), self.stepGetConnection, user, service

    def _getStepSendMailMessage(self, service, thread):
        return 30501, (service.value, ), self.stepSendMailMessage, service, thread

    def _getStepFinalize(self, server):
        return 30601, (server, ), self.stepFinalize

    def _getJobSteps(self, ctx, user, service, sender, recipient, subject, message, thread):
        yield self._getStepInitJob(ctx, service, self.__class__.__name__)
        yield self._getStepBuildMailMessage(ctx, service, sender, recipient, subject, message)
        if self._mail:
            yield self._getStepGetConnection(user, service)
            if self._server.isConnected():
                yield self._getStepSendMailMessage(service, thread)
            else:
                yield self._getStepSkipStep(30401)
        else:
            yield self._getStepSkipStep(30301)
            yield self._getStepSkipStep(30301)
        yield self._getStepFinalize(service.value)

    def _sendImapMessage(self):
        success = False
        try:
            folder = self._server.getSentFolder()
            if self._server.hasFolder(folder):
                self._server.uploadMessage(folder, self._mail)
        except UnoException:
            # Exceptions are already logged in the MailServiceLogger log.
            pass
        else:
            success = True
        self._success = success

    def _sendSmtpMessage(self, thread):
        success = False
        try:
            # If IMAP is supported then we set the Message.ThreadId
            if thread and hasInterface(self._mail, 'com.sun.star.mail.XMailMessage2'):
                self._mail.ThreadId = thread
            self._server.sendMailMessage(self._mail)
        except UnoException:
            # Exceptions are already logged in the MailServiceLogger log.
            pass
        else:
            success = True
        self._mail = None
        self._success = success
