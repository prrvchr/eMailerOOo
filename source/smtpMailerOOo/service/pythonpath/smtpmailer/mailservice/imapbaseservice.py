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

from .imapservice import ImapService

from ..oauth2lib import getOAuth2Token

from ..unotool import hasInterface

from .. import imapclient

import traceback

class ImapBaseService(ImapService):

    def connect(self, context, authenticator):
        if self._logger.isDebugMode():
            self._logger.logprb(INFO, 'ImapService', 'connect()', 321)
        if self.isConnected():
            raise AlreadyConnectedException()
        if not hasInterface(context, 'com.sun.star.uno.XCurrentContext'):
            raise IllegalArgumentException()
        if not hasInterface(authenticator, 'com.sun.star.mail.XAuthenticator'):
            raise IllegalArgumentException()
        self._context = context
        server = context.getValueByName('ServerName')
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType').title()
        authentication = context.getValueByName('AuthenticationType').title()
        if connection == 'Ssl':
            self._server = imapclient.IMAPClient(server, port=port, ssl=True, timeout=timeout)
        else:
            self._server = imapclient.IMAPClient(server, port=port, ssl=False, timeout=timeout)
        if connection == 'Tls':
            self._server.starttls()
        if authentication == 'Login':
            user = authenticator.getUserName()
            password = authenticator.getPassword()
            code = self._server.login(user, password)
            print("ImapService.connect() 1: %s" % code)
        elif authentication == 'Oauth2':
            user = authenticator.getUserName()
            token = getOAuth2Token(self._ctx, self, server, user)
            code = self._server.oauth2_login(user, token)
            print("ImapService.connect() 2: %s" % code)
        for listener in self._listeners:
            listener.connected(self._notify)
        if self._logger.isDebugMode():
            self._logger.logprb(INFO, 'ImapService', 'connect()', 324)

    def disconnect(self):
        if self.isConnected():
            if self._logger.isDebugMode():
                self._logger.logprb(INFO, 'ImapService', 'disconnect()', 361)
            self._server.logout()
            self._server = None
            self._context = None
            for listener in self._listeners:
                listener.disconnected(self._notify)
            if self._logger.isDebugMode():
                self._logger.logprb(INFO, 'ImapService', 'disconnect()', 362)

    def isConnected(self):
        return self._server is not None

    def getSentFolder(self):
        data = self._server.find_special_folder(imapclient.SENT)
        return data if data is not None else ''

    def hasFolder(self, folder):
        find = False
        if folder:
            find = self._server.folder_exists(folder)
        return find

    def uploadMessage(self, folder, message):
        code = self._server.append(folder, message.asString(False))
        print("MailServiceProvider.uploadMessage() %s" % (code, ))
