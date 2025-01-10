#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-24 https://prrvchr.github.io                                  ║
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

from com.sun.star.mail.MailServiceType import IMAP

from com.sun.star.auth import AuthenticationFailedException

from com.sun.star.lang import IllegalArgumentException
from com.sun.star.lang import EventObject

from com.sun.star.io import AlreadyConnectedException
from com.sun.star.io import NotConnectedException

from com.sun.star.mail import MailException
from com.sun.star.mail import XImapService

from .apihelper import getHttpProvider
from .apihelper import getHttpServer
from .apihelper import getHttpRequest
from .apihelper import getOAuth2TokenWithParameters
from .apihelper import executeHttpRequest

from ..oauth2 import getOAuth2

from ..unotool import getConfiguration
from ..unotool import getExceptionMessage
from ..unotool import hasInterface

from ..configuration import g_identifier

from imapclient import IMAPClient
from imapclient.imapclient import SENT
from imapclient.exceptions import IMAPClientError

from email.base64mime import body_encode as encode_base64
import traceback


class ImapService(unohelper.Base,
                  XImapService):
    def __init__(self, ctx, logger, providers, debug):
        mtd = '__init__'
        self._cls = 'ImapService'
        self._ctx = ctx
        if debug:
            logger.logprb(INFO, self._cls, mtd, 301)
        self._providers = providers
        self._listeners = []
        self._supportedconnection =     ('Insecure', 'SSL', 'TLS')
        self._supportedauthentication = ('None', 'Login', 'OAuth2')
        self._supportedmechanism =      ('OAUTHBEARER', 'XOAUTH2')
        self._mechanism = 'AUTH='
        self._server = None
        self._context = None
        self._tokens = None
        self._provider = None
        self._url = ''
        self._logger = logger
        self._debug = debug
        if debug:
            logger.logprb(INFO, self._cls, mtd, 302)

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
        if self._server is None:
            raise NotConnectedException()
        return self._context

    def connect(self, context, authenticator):
        mtd = 'connect'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 311)
        if self.isConnected():
            raise AlreadyConnectedException()
        if not hasInterface(context, 'com.sun.star.uno.XCurrentContext'):
            raise IllegalArgumentException()
        if not hasInterface(authenticator, 'com.sun.star.mail.XAuthenticator'):
            raise IllegalArgumentException()
        self._server = self._getServer(context, authenticator)
        self._context = context
        notify = EventObject(self)
        for listener in self._listeners:
            listener.connected(notify)
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 312)

    def disconnect(self):
        mtd = 'disconnect'
        if self.isConnected():
            if self._debug:
                self._logger.logprb(INFO, self._cls, mtd, 391)
            if self._isImapServer():
                self._server.logout()
            self._server = None
            self._context = None
            notify = EventObject(self)
            for listener in self._listeners:
                listener.disconnected(notify)
            if self._debug:
                self._logger.logprb(INFO, self._cls, mtd, 392)

    def isConnected(self):
        return self._server is not None

    def getSentFolder(self):
        if self._server is None:
            raise NotConnectedException()
        return self._getImapSentFolder() if self._isImapServer() else self._getApiSentFolder()

    def hasFolder(self, folder):
        if self._server is None:
            raise NotConnectedException()
        return self._hasImapFolder(folder) if self._isImapServer() else self._hasApiFolder(folder)

    def uploadMessage(self, folder, message):
        if self._server is None:
            raise NotConnectedException()
        if self._isImapServer():
            self._uploadImapMessage(folder, message)
        else:
            self._uploadHttpMessage(folder, message)

# Private methods implementation
    def _getServer(self, context, authenticator):
        host = context.getValueByName('ServerName')
        user = authenticator.getUserName()
        provider, url = getHttpProvider(self._providers, context.getValueByName('Provider'))
        if provider:
            server = self._getHttpServer(provider, url, host, user)
        else:
            server = self._getImapServer(context, host, user, authenticator.getPassword())
        return server

    def _isImapServer(self):
        return self._provider is None

    def _getHttpServer(self, provider, url, host, user):
        mtd = '_getHttpServer'
        self._provider = provider
        server = getHttpServer(self._ctx, url, host, user)
        if server is None:
            msg = self._logger.resolveString(321, user)
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        return server

    def _getImapServer(self, context, host, user, password):
        mtd = '_getImapServer'
        self._provider = None
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType')
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 331, connection)
        try:
            if connection.upper() == 'SSL':
                if self._debug:
                    self._logger.logprb(INFO, self._cls, mtd, 332, timeout)
                server = IMAPClient(host, port=port, ssl=True, timeout=timeout)
            else:
                if self._debug:
                    self._logger.logprb(INFO, self._cls, mtd, 333, timeout)
                server = IMAPClient(host, port=port, ssl=False, timeout=timeout)
        except IMAPClientError as e:
            msg = self._logger.resolveString(334, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise ConnectException(msg, self)
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 335, host, port)
        if connection.upper() == 'TLS':
            self._doStartTls(server)
        self._doAuthentication(context, server, host, port, user, password)
        return server

    def _doStartTls(self, server):
        mtd = '_doStartTls'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 341)
        try:
            server.starttls()
        except IMAPClientError as e:
            msg = self._logger.resolveString(342, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise ConnectException(msg, self)
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 343)

    def _doAuthentication(self, context, server, host, port, user, password):
        authentication = context.getValueByName('AuthenticationType')
        if authentication is None:
            if user:
                self._doLogin(server, user, password, 'Login')
        elif authentication.upper() == 'LOGIN':
            self._doLogin(server, user, password, authentication)
        elif authentication.upper() == 'OAUTH2':
            self._doOAuth2(server, host, port, user, authentication)

    def _doLogin(self, server, user, password, authentication):
        mtd = '_doLogin'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 351, authentication)
        try:
            code = server.login(user, password)
        except IMAPClientError as e:
            pwd = '*' * len(password)
            msg = self._logger.resolveString(352, user, pwd, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        if self._debug:
            pwd = '*' * len(password)
            self._logger.logprb(INFO, self._cls, mtd, 353, user, pwd, self._getReply(code))

    def _doOAuth2(self, server, host, port, user, authentication):
        mtd = '_doOAuth2'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 361, authentication)
        oauth2 = getOAuth2(self._ctx, host, user)
        if not oauth2:
            msg = self._logger.resolveString(362)
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        try:
            ex = None
            features = self._getServerFeatures(server)
            # FIXME: Due to the lack of established standards
            # FIXME: it is necessary to try most OAuth2 mechanisms
            for mechanism in self._supportedmechanism:
                if mechanism in features:
                    parameters = self._getTokenParameters(mechanism, host)
                    token = getOAuth2TokenWithParameters(oauth2, parameters, port)
                    callable = lambda x: encode_base64(token.encode(), eol='')
                    #callable = lambda x: token
                    try:
                        code = server.sasl_login(mechanism, callable)
                    except IMAPClientError as e:
                        ex = e
                        continue
                    if self._debug:
                        self._logger.logprb(INFO, self._cls, mtd, 363, mechanism, self._getReply(code))
                    return
            if ex:
                msg = self._logger.resolveString(364, mechanism, getExceptionMessage(ex))
            else:
                mechanisms = ', '.join(self._supportedmechanism)
                msg = self._logger.resolveString(365, mechanisms, ', '.join(features))
        except IMAPClientError as e:
            msg = self._logger.resolveString(364, mechanism, getExceptionMessage(e))
        if self._debug:
            self._logger.logp(SEVERE, self._cls, mtd, msg)
        raise AuthenticationFailedException(msg, self)

    def _getImapSentFolder(self):
        data = self._server.find_special_folder(SENT)
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

    def _uploadHttpMessage(self, folder, message):
        mtd = '_uploadHttpMessage'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 371, message.Subject)
        request = getHttpRequest(self._getProvider(), IMAP, 'Request')
        if request:
            executeHttpRequest(self, self._logger, self._debug, self._cls, 372, self._server, message, request)
            subrequest = getHttpRequest(self._getProvider(), IMAP, 'SubRequest')
            if subrequest:
                executeHttpRequest(self, self._logger, self._debug, self._cls, 372, self._server, message, subrequest)
        else:
            self._logger.logprb(SEVERE, self._cls, mtd, 374, message.Subject, self._provider)
            return
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 375, message.Subject)

    def _uploadImapMessage(self, folder, message):
        mtd = '_uploadImapMessage'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 381, message.Subject)
        try:
            code = self._server.append(folder, message.asString())
        except IMAPClientError as e:
            msg = self._logger.resolveString(382, message.Subject, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 383, message.Subject, self._getReply(code))

    def _getReply(self, reply):
        if isinstance(reply, bytes):
            reply = reply.decode('ascii', 'ignore')
        return reply

    def _getTokenParameters(self, mechanism, host):
        if self._tokens is None:
            self._tokens = getConfiguration(self._ctx, g_identifier).getByName('Tokens')
        return self._tokens.get(host, self._tokens.get(mechanism))

    def _getServerFeatures(self, server):
        return tuple(s.decode('ascii', 'ignore').replace(self._mechanism, '') for s in server.capabilities())

    def _getProvider(self):
        return self._providers.getByName(self._provider)

