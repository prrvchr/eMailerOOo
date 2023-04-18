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

from com.sun.star.lang import IllegalArgumentException
from com.sun.star.io import AlreadyConnectedException
from com.sun.star.mail import MailException

from .imapservice import ImapService
from .apihelper import parseMessage
from .apihelper import setDefaultFolder

from ..unotool import getExceptionMessage
from ..unotool import hasInterface

from ..oauth2 import getOAuth2Token
from ..oauth2 import getRequest


import traceback

class ImapApiService(ImapService):

    def connect(self, context, authenticator):
        self._logger.logprb(INFO, 'ImapService', 'connect()', 321)
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
        self._logger.logprb(INFO, 'ImapService', 'connect()', 324)

    def disconnect(self):
        if self.isConnected():
            self._logger.logprb(INFO, 'ImapService', 'disconnect()', 361)
            self._server = None
            self._context = None
            for listener in self._listeners:
                listener.disconnected(self._notify)
            self._logger.logprb(INFO, 'ImapService', 'disconnect()', 362)

    def isConnected(self):
        return self._server is not None

    def getSentFolder(self):
        return 'SENT'

    def hasFolder(self, folder):
        find = False
        if folder == 'SENT':
            find = True
        return find

    def uploadMessage(self, folder, message):
        url = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/'
        parameter = self._server.getRequestParameter('uploadMessage')
        parameter.Method = 'POST'
        parameter.Url = url + 'send'
        parameter.Json = '{"raw": "%s"}' % message.asString(True)
        try:
            response = self._server.execute(parameter)
        except Exception as e:
            msg = self._logger.resolveString(253, message.Subject, getExceptionMessage(e))
            raise MailException(msg, self)
        else:
            if response.Ok:
                messageid, labels = parseMessage(response)
                setDefaultFolder(self._server, url, messageid, labels, folder)
                message.MessageId = messageid
            response.close()

