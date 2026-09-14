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

from com.sun.star.mail.MailServiceType import SMTP

from ...runner import Task

from ...unotool import Resource

from ...helper import getMailService


class Server(Task):
    def __init__(self, ctx, callback, user, service):
        super().__init__(ctx, 'Ispdb', callback)
        self._server = None
        self._connect = False
        self._imap = service == SMTP and user.useIMAP()
        self._resources = {1: Resource(20101, lambda x: (x.total, )),
                           2: Resource(20102, lambda x: (x.total, )),
                           3: Resource(20103),
                           4: Resource(20104),
                           5: Resource(20105),
                           6: Resource(20106)}
        self.step = 0 if service == SMTP else 3
        self.total = 6 if user.useIMAP() else 3
        self.steps = self._getJobSteps(ctx, user, service)

    def callback(self, success=True, error=None):
        self._callback(success, self._connect, self._imap, error)

    def stepInitJob(self, ctx, service):
        self._server = getMailService(ctx, service)

    def stepGetConnection(self, user, service):
        context = user.getConnectionContext(service.value)
        authenticator = user.getAuthenticator(service.value)
        self._server.connect(context, authenticator)
        if self._server.isConnected():
            self._connect = True

    def stepFinalize(self):
        if self._server and self._server.isConnected():
            self._server.disconnect()
        self._server = None

    def _getStepInitJob(self, ctx, service, name):
        return 20201, (service.value, name), self.stepInitJob, ctx, service

    def _getStepGetConnection(self, user, service):
        return 20301, (service.value, ), self.stepGetConnection, user, service

    def _getStepFinalize(self, server):
        return 20401, (server, ), self.stepFinalize

    def _getJobSteps(self, ctx, user, service):
        yield self._getStepInitJob(ctx, service, self.__class__.__name__)
        yield self._getStepGetConnection(user, service)
        yield self._getStepFinalize(service.value)
