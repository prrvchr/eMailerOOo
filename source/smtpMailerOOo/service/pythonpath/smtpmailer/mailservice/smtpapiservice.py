#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020 https://prrvchr.github.io                                     ║
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
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import MailException

from .smtpservice import SmtpService
from .apihelper import setDefaultFolder

from ..oauth2lib import getRequest

from ..unotool import getExceptionMessage
from ..unotool import hasInterface

import sys
import six
import base64
from collections import OrderedDict
import traceback


class SmtpApiService(SmtpService):

    def connect(self, context, authenticator):
        if self._logger.isDebugMode():
            self._logger.logprb(INFO, 'SmtpService', 'connect()', 221)
        if self.isConnected():
            raise AlreadyConnectedException()
        if not hasInterface(context, 'com.sun.star.uno.XCurrentContext'):
            raise IllegalArgumentException()
        if not hasInterface(authenticator, 'com.sun.star.mail.XAuthenticator'):
            raise IllegalArgumentException()
        server = context.getValueByName('ServerName')
        user = authenticator.getUserName()
        self._server = getRequest(self._ctx, server, user)
        self._context = context
        for listener in self._listeners:
            listener.connected(self._notify)
        if self._logger.isDebugMode():
            self._logger.logprb(INFO, 'SmtpService', 'connect()', 224)

    def isConnected(self):
        return self._server is not None

    def disconnect(self):
        if self.isConnected():
            if self._logger.isDebugMode():
                self._logger.logprb(INFO, 'SmtpService', 'disconnect()', 261)
            self._server = None
            self._context = None
            for listener in self._listeners:
                listener.disconnected(self._notify)
            if self._logger.isDebugMode():
                self._logger.logprb(INFO, 'SmtpService', 'disconnect()', 262)

    def sendMailMessage(self, message):
        url = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/'
        parameter = uno.createUnoStruct('com.sun.star.auth.RestRequestParameter')
        parameter.Method = 'POST'
        parameter.Url = url + 'send'
        parameter.Json = '{"threadId": "%s", "raw": "%s"}' % (message.ThreadId, message.asString(True))
        try:
            response = self._server.execute(parameter)
        except Exception as e:
            msg = self._logger.resolveString(253, message.Subject, getExceptionMessage(e))
            raise MailException(msg, self)
        if response.IsPresent:
            messageid = response.Value.getValue('id')
            labels = response.Value.getValue('labelIds')
            setDefaultFolder(self._server, url, messageid, labels)
            message.MessageId = messageid
