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

from com.sun.star.logging.LogLevel import ALL
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.auth import AuthenticationFailedException

from com.sun.star.lang import IllegalArgumentException
from com.sun.star.lang import EventObject

from com.sun.star.io import AlreadyConnectedException
from com.sun.star.io import NotConnectedException
from com.sun.star.io import ConnectException
from com.sun.star.io import UnknownHostException

from com.sun.star.mail import MailException
from com.sun.star.mail import XSmtpService2

from .. import smtplib

from .apihelper import getHttpProvider
from .apihelper import getHttpServer
from .apihelper import getHttpRequest
from .apihelper import executeHTTPRequest

from ..oauth2 import getOAuth2Token

from ..unotool import getExceptionMessage
from ..unotool import hasInterface

import six
import base64
from collections import OrderedDict
import traceback


class SmtpService(unohelper.Base,
                  XSmtpService2):
    def __init__(self, ctx, logger, hosts, debug=False):
        mtd = '__init__'
        self._cls = 'SmtpService'
        self._ctx = ctx
        if debug:
            logger.logprb(INFO, self._cls, mtd, 201)
        self._listeners = []
        self._supportedconnection = ('Insecure', 'SSL', 'TLS')
        self._supportedauthentication = ('None', 'Login', 'OAuth2')
        self._server = None
        self._context = None
        self._provider = None
        self._url = ''
        self._hosts = hosts
        self._logger = logger
        self._debug = debug
        if debug:
            logger.logprb(INFO, self._cls, mtd, 202)

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
            self._logger.logprb(INFO, self._cls, mtd, 211)
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
            self._logger.logprb(INFO, self._cls, mtd, 212)

    def disconnect(self):
        mtd = 'disconnect'
        if self.isConnected():
            if self._debug:
                self._logger.logprb(INFO, self._cls, mtd, 291)
            if self._isSmtpServer():
                self._server.quit()
            self._server = None
            self._context = None
            notify = EventObject(self)
            for listener in self._listeners:
                listener.disconnected(notify)
            if self._debug:
                self._logger.logprb(INFO, self._cls, mtd, 292)

    def isConnected(self):
        return self._server is not None

    def sendMailMessage(self, message):
        if self._server is None:
            raise NotConnectedException()
        if self._isSmtpServer():
            self._sendSmtpMailMessage(message)
        else:
            self._sendHttpMailMessage(message)

# Private methods implementation
    def _getServer(self, context, authenticator):
        servername = context.getValueByName('ServerName')
        username = authenticator.getUserName()
        password = authenticator.getPassword()
        provider = getHttpProvider(self._hosts, servername)
        if provider:
            server = self._getHttpServer(provider, servername, username)
        else:
            server = self._getSmtpServer(context, servername, username, password)
        return server

    def _isSmtpServer(self):
        return self._provider is None

    def _getHttpServer(self, provider, servername, username):
        mtd = '_getHttpServer'
        self._provider = provider
        server = getHttpServer(self._ctx, provider, servername, username)
        if server is None:
            msg = self._logger.resolveString(221, username)
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        return server

    def _getSmtpServer(self, context, servername, username, password):
        mtd = '_getSmtpServer'
        self._provider = None
        server = None
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType')
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 231, connection)
        try:
            if connection.upper() == 'SSL':
                if self._debug:
                    self._logger.logprb(INFO, self._cls, mtd, 232, timeout)
                server = smtplib.SMTP_SSL(timeout=timeout)
            else:
                if self._debug:
                    self._logger.logprb(INFO, self._cls, mtd, 233, timeout)
                server = smtplib.SMTP(timeout=timeout)
            if self._debug:
                server.set_debuglevel(1)
            code, reply = server.connect(host=servername, port=port)
        except smtplib.SMTPConnectError as e:
            msg = self._logger.resolveString(234, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise ConnectException(msg, self)
        except smtplib.SMTPException as e:
            msg = self._logger.resolveString(234, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise UnknownHostException(msg, self)
        except Exception as e:
            msg = self._logger.resolveString(234, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise MailException(msg, self)
        if code != 220:
            msg = self._logger.resolveString(234, self._getReply(reply))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise ConnectException(msg, self)
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 235, servername, port, code, self._getReply(reply))
        if connection.upper() == 'TLS':
            self._doStartTls(server)
        self._doAuthentication(context, server, servername, username, password)
        return server

    def _doStartTls(self, server):
        mtd = '_doStartTls'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 241)
        try:
            code, reply = server.starttls()
        except smtplib.SMTPException as e:
            msg = self._logger.resolveString(242, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise ConnectException(msg, self)
        if code != 220:
            msg = self._logger.resolveString(242, self._getReply(reply))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise ConnectException(msg, self)
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 243, code, self._getReply(reply))

    def _doAuthentication(self, context, server, servername, username, password):
        authentication = context.getValueByName('AuthenticationType')
        if authentication is None:
            if username:
                self._doLogin(server, username, password, 'Login')
        elif authentication.upper() == 'LOGIN':
            self._doLogin(server, username, password, authentication)
        elif authentication.upper() == 'OAUTH2':
            self._doOAuth2(server, servername, username, password, authentication)

    def _doLogin(self, server, username, password, authentication):
        mtd = '_doLogin'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 251, authentication)
        if six.PY2: # fdo#59249 i#105669 Python 2 needs "ascii"
            username = username.encode('ascii')
            password = password.encode('ascii')
        try:
            code, reply = server.login(username, password)
        except smtplib.SMTPException as e:
            pwd = '*' * len(password)
            msg = self._logger.resolveString(252, username, pwd, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        print("SmtpService._doLogin() Code: %s - Reply: %s" % (code, self._getReply(reply)))
        if code != 235:
            pwd = '*' * len(password)
            msg = self._logger.resolveString(252, username, pwd, self._getReply(reply))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        if self._debug:
            pwd = '*' * len(password)
            self._logger.logprb(INFO, self._cls, mtd, 253, username, pwd, code, self._getReply(reply))

    def _doOAuth2(self, server, servername, username, password, authentication):
        mtd = '_doOAuth2'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 261, authentication)
        try:
            token = self._getToken(servername, username, True)
            server.ehlo_or_helo_if_needed()
            code, reply = server.docmd('AUTH', 'XOAUTH2 %s' % token)
        except smtplib.SMTPException as e:
            msg = self._logger.resolveString(262, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        if code != 235:
            msg = self._logger.resolveString(262, self._getReply(reply))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 263, code, self._getReply(reply))

    def _getToken(self, url, username, encode=False):
        token = getOAuth2Token(self._ctx, self, url, username)
        authstring = 'user=%s\1auth=Bearer %s\1\1' % (username, token)
        if encode:
            authstring = base64.b64encode(authstring.encode('ascii')).decode('ascii')
        return authstring

    def _sendHttpMailMessage(self, message):
        mtd = '_sendHttpMailMessage'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 271, message.Subject)
        request = getHttpRequest(self._ctx, self._provider, SMTP, 'Request')
        if request:
            executeHTTPRequest(self, self._logger, self._debug, self._cls, 272, self._server, message, request)
            subrequest = getHttpRequest(self._ctx, self._provider, SMTP, 'SubRequest')
            if subrequest:
                executeHTTPRequest(self, self._logger, self._debug, self._cls, 272, self._server, message, subrequest)
        else:
            self._logger.logprb(SEVERE, self._cls, mtd, 274, message.Subject, self._provider)
            return
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 275, message.Subject)

    def _sendSmtpMailMessage(self, message):
        mtd = '_sendSmtpMailMessage'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 281, message.Subject)
        recipients = self._getRecipients(message)
        try:
            refused = self._server.sendmail(message.SenderAddress, recipients, message.asString())
        except smtplib.SMTPSenderRefused as e:
            msg = self._logger.resolveString(282, message.Subject, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise MailException(msg, self)
        except smtplib.SMTPRecipientsRefused as e:
            msg = self._logger.resolveString(282, message.Subject, getExceptionMessage(e))
            # TODO: return SendMailMessageFailedException in place of MailException
            # TODO: error = SendMailMessageFailedException(msg, self)
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise MailException(msg, self)
        except smtplib.SMTPDataError as e:
            msg = self._logger.resolveString(282, message.Subject, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise MailException(msg, self)
        except Exception as e:
            msg = self._logger.resolveString(282, message.Subject, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise MailException(msg, self)
        else:
            if len(refused) > 0:
                for address, result in refused.items():
                    code, reply = result
                    self._logger.logprb(SEVERE, self._cls, mtd, 283, message.Subject, address, code, self._getReply(reply))
            elif self._debug:
                self._logger.logprb(INFO, self._cls, mtd, 284, message.Subject)

    def _getRecipients(self, message):
        recipients = OrderedDict()
        for recipient in message.getRecipients():
            recipients[recipient] = True
        if message.hasCcRecipients():
            for recipient in message.getCcRecipients():
                recipients[recipient] = True
        if message.hasBccRecipients():
            for recipient in message.getBccRecipients():
                recipients[recipient] = True
        return recipients.keys()

    def _getReply(self, reply):
        if isinstance(reply, six.binary_type):
            reply = reply.decode('ascii')
        return reply

