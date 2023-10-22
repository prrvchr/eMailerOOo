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
from com.sun.star.lang import EventObject

from com.sun.star.io import AlreadyConnectedException

from com.sun.star.mail import MailException
from com.sun.star.mail import XImapService

from .apihelper import parseMessage
from .apihelper import setDefaultFolder

from ..oauth2 import getOAuth2Token
from ..oauth2 import getRequest

from ..unotool import getExceptionMessage
from ..unotool import hasInterface

from .. import imapclient

import base64
import traceback


class ImapService(unohelper.Base,
                  XImapService):
    def __init__(self, ctx, logger, domains, debug=False):
        self._ctx = ctx
        if debug:
            logger.logprb(INFO, 'ImapService', '__init__()', 311)
        self._listeners = []
        self._supportedconnection = ('Insecure', 'SSL', 'TLS')
        self._supportedauthentication = ('None', 'Login', 'OAuth2')
        self._server = None
        self._context = None
        self._imap = True
        self._url = ''
        self._domains = domains
        self._notify = EventObject(self)
        self._logger = logger
        self._debug = debug
        if debug:
            logger.logprb(INFO, 'ImapService', '__init__()', 312)

# XMailService interface implementation
    def addConnectionListener(self, listener):
        self._listeners.append(listener)

    def removeConnectionListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def getSupportedConnectionTypes(self):
        return self._supportedconnection

    def getSupportedAuthenticationTypes(self):
        return self._supportedauthentication

    def getCurrentConnectionContext(self):
        return self._context

# Interface not implemented
    def connect(self, context, authenticator):
        if self._debug:
            self._logger.logprb(INFO, 'ImapService', 'connect()', 321)
        if self.isConnected():
            raise AlreadyConnectedException()
        if not hasInterface(context, 'com.sun.star.uno.XCurrentContext'):
            raise IllegalArgumentException()
        if not hasInterface(authenticator, 'com.sun.star.mail.XAuthenticator'):
            raise IllegalArgumentException()
        self._server = self._getServer(context, authenticator)
        self._context = context
        for listener in self._listeners:
            listener.connected(self._notify)
        if self._debug:
            self._logger.logprb(INFO, 'ImapService', 'connect()', 324)

    def disconnect(self):
        if self.isConnected():
            if self._debug:
                self._logger.logprb(INFO, 'ImapService', 'disconnect()', 361)
            if self._imap:
                self._server.logout()
            self._server = None
            self._context = None
            for listener in self._listeners:
                listener.disconnected(self._notify)
            if self._debug:
                self._logger.logprb(INFO, 'ImapService', 'disconnect()', 362)

    def isConnected(self):
        return self._server is not None

    def getSentFolder(self):
        return self._getImapSentFolder() if self._imap else self._getApiSentFolder()

    def hasFolder(self, folder):
        return self._hasImapFolder(folder) if self._imap else self._hasApiFolder(folder)

    def uploadMessage(self, folder, message):
        if self._imap:
            self._uploadImapMessage(folder, message)
        else:
            self._uploadApiMessage(folder, message)

# Private methods implementation
    def _getServer(self, context, authenticator):
        servername = context.getValueByName('ServerName')
        username = authenticator.getUserName()
        password = authenticator.getPassword()
        host, sep, domain = servername.partition('.')
        if domain in self._domains:
            server = self._getHttpServer(servername, username, domain)
        else:
            server = self._getImapServer(context, servername, username, password)
        return server

    def _getHttpServer(self, servername, username, domain):
        self._imap = False
        self._url = self._domains[domain]
        server = getRequest(self._ctx, servername, username)
        return server

    def _getImapServer(self, context, servername, username, password):
        self._imap = True
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType')
        try:
            if connection.upper() == 'SSL':
                server = imapclient.IMAPClient(servername, port=port, ssl=True, timeout=timeout)
            else:
                server = imapclient.IMAPClient(servername, port=port, ssl=False, timeout=timeout)
        except imapclient.exceptions.IMAPClientError as e:
            msg = self._logger.resolveString(236, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, 'SmtpService', '_getImapServer()', msg)
            raise ConnectException(msg, self)
        if connection.upper() == 'TLS':
            self._doStartTls(server)
        self._doAuthentication(context, server, servername, username, password)
        return server

    def _doStartTls(self, server):
        try:
            server.starttls()
        except imapclient.exceptions.IMAPClientError as e:
            msg = self._logger.resolveString(236, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, 'SmtpService', '_doStartTls()', msg)
            raise ConnectException(msg, self)

    def _doAuthentication(self, context, server, servername, username, password):
        authentication = context.getValueByName('AuthenticationType')
        if authentication is None:
            if username:
                self._doLogin(server, username, password)
        elif authentication.upper() == 'LOGIN':
            self._doLogin(server, username, password)
        elif authentication.upper() == 'OAUTH2':
            self._doOAuth2(server, servername, username)

    def _doLogin(self, server, username, password):
        try:
            code = server.login(username, password)
        except imapclient.exceptions.IMAPClientError as e:
            msg = self._logger.resolveString(236, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, 'SmtpService', '_doLogin()', msg)
            raise ConnectException(msg, self)

    def _doOAuth2(self, server, servername, username):
        token = getOAuth2Token(self._ctx, self, servername, username)
        try:
            code = server.oauth2_login(username, token)
        except imapclient.exceptions.IMAPClientError as e:
            msg = self._logger.resolveString(236, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, 'SmtpService', '_doAuthentication()', msg)
            raise ConnectException(msg, self)

    def _getImapSentFolder(self):
        data = self._server.find_special_folder(imapclient.SENT)
        return data if data is not None else ''

    def _getApiSentFolder(self):
        return 'SENT'

    def _hasImapFolder(self, folder):
        find = False
        if folder:
            find = self._server.folder_exists(folder)
        return find

    def _hasApiFolder(self, folder):
        find = False
        if folder == 'SENT':
            find = True
        return find

    def _uploadImapMessage(self, folder, message):
        print("ImapBaseService.uploadMessage() 1")
        if hasInterface(message, 'com.sun.star.mail.XMailMessage2'):
            print("ImapBaseService.uploadMessage() 2 ********************************")
        print("ImapBaseService.uploadMessage() 3")
        code = self._server.append(folder, message.asString())
        print("MailServiceProvider.uploadMessage() %s" % (code, ))

    def _uploadApiMessage(self, folder, message):
        print("ImapApiService.uploadMessage() 1")
        if hasInterface(message, 'com.sun.star.mail.XMailMessage2'):
            print("ImapApiService.uploadMessage() 2 ********************************")
        print("ImapApiService.uploadMessage() 3")
        parameter = self._server.getRequestParameter('uploadMessage')
        parameter.Method = 'POST'
        parameter.Url = self._url + 'send'
        raw = base64.urlsafe_b64encode(message.asBytes().value)
        parameter.setJson('raw', raw)
        try:
            response = self._server.execute(parameter)
        except Exception as e:
            msg = self._logger.resolveString(253, message.Subject, getExceptionMessage(e))
            raise MailException(msg, self)
        else:
            if response.Ok:
                messageid, labels = parseMessage(response)
                setDefaultFolder(self._server, self._url, messageid, labels, folder)
                message.MessageId = messageid
            response.close()

