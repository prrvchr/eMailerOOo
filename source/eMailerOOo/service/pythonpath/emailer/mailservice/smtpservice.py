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

from smtplib import SMTP as SMTP_BASE
from smtplib import SMTP_SSL
from smtplib import SMTPConnectError
from smtplib import SMTPDataError
from smtplib import SMTPException
from smtplib import SMTPRecipientsRefused
from smtplib import SMTPSenderRefused

from collections import OrderedDict
import traceback


class SmtpService(unohelper.Base,
                  XSmtpService2):
    def __init__(self, ctx, logger, providers, debug):
        mtd = '__init__'
        self._cls = 'SmtpService'
        self._ctx = ctx
        if debug:
            logger.logprb(INFO, self._cls, mtd, 201)
        self._providers = providers
        self._listeners = []
        self._supportedconnection =     ('Insecure', 'SSL', 'TLS')
        self._supportedauthentication = ('None', 'Login', 'OAuth2')
        self._supportedmechanism =      ('OAUTHBEARER', 'XOAUTH2')
        self._server = None
        self._context = None
        self._tokens = None
        self._provider = None
        self._url = ''
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
        host = context.getValueByName('ServerName')
        user = authenticator.getUserName()
        provider, url = getHttpProvider(self._providers, context.getValueByName('Provider'))
        if provider:
            server = self._getHttpServer(provider, url, host, user)
        else:
            server = self._getSmtpServer(context, host, user, authenticator.getPassword())
        return server

    def _isSmtpServer(self):
        return self._provider is None

    def _getHttpServer(self, provider, url, host, user):
        mtd = '_getHttpServer'
        self._provider = provider
        server = getHttpServer(self._ctx, url, host, user)
        if server is None:
            msg = self._logger.resolveString(221, user)
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        return server

    def _getSmtpServer(self, context, host, user, password):
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
                server = SMTP_SSL(host=host, port=port, timeout=timeout)
            else:
                if self._debug:
                    self._logger.logprb(INFO, self._cls, mtd, 233, timeout)
                server = SMTP_BASE(host=host, port=port, timeout=timeout)
            if self._debug:
                server.set_debuglevel(1)
            code, reply = server.connect(host=host, port=port)
        except SMTPConnectError as e:
            msg = self._logger.resolveString(234, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise ConnectException(msg, self)
        except SMTPException as e:
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
            self._logger.logprb(INFO, self._cls, mtd, 235, host, port, code, self._getReply(reply))
        if connection.upper() == 'TLS':
            self._doStartTls(server)
        self._doAuthentication(context, server, host, port, user, password)
        return server

    def _doStartTls(self, server):
        mtd = '_doStartTls'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 241)
        try:
            code, reply = server.starttls()
        except SMTPException as e:
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

    def _doAuthentication(self, context, server, host, port, user, password):
        authentication = context.getValueByName('AuthenticationType')
        if authentication is None:
            if user:
                self._doLogin(server, user, password, 'Login')
        elif authentication.upper() == 'LOGIN':
            self._doLogin(server, user, password, authentication)
        elif authentication.upper() == 'OAUTH2':
            self._doOAuth2(server, host, port, user, password, authentication)

    def _doLogin(self, server, user, password, authentication):
        mtd = '_doLogin'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 251, authentication)
        try:
            code, reply = server.login(user, password)
        except SMTPException as e:
            pwd = '*' * len(password)
            msg = self._logger.resolveString(252, user, pwd, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        print("SmtpService._doLogin() Code: %s - Reply: %s" % (code, self._getReply(reply)))
        if code != 235:
            pwd = '*' * len(password)
            msg = self._logger.resolveString(252, user, pwd, self._getReply(reply))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        if self._debug:
            pwd = '*' * len(password)
            self._logger.logprb(INFO, self._cls, mtd, 253, user, pwd, code, self._getReply(reply))

    def _doOAuth2(self, server, host, port, user, password, authentication):
        mtd = '_doOAuth2'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 261, authentication)
        oauth2 = getOAuth2(self._ctx, host, user)
        if not oauth2:
            msg = self._logger.resolveString(262)
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise AuthenticationFailedException(msg, self)
        try:
            ex = None
            server.ehlo_or_helo_if_needed()
            features = self._getServerFeatures(server)
            # FIXME: Due to the lack of established standards
            # FIXME: it is necessary to try most OAuth2 mechanisms
            for mechanism in self._supportedmechanism:
                if mechanism in features:
                    parameters = self._getTokenParameters(mechanism, host)
                    callable = lambda x=None: getOAuth2TokenWithParameters(oauth2, parameters, port)
                    for ok in (True, False):
                        try:
                            code, reply = server.auth(mechanism, callable, initial_response_ok=ok)
                        except SMTPException as e:
                            ex = e
                            continue
                        if self._debug:
                            self._logger.logprb(INFO, self._cls, mtd, 263, mechanism, code, self._getReply(reply))
                        return
            if ex:
                msg = self._logger.resolveString(264, mechanism, getExceptionMessage(ex))
            else:
                mechanisms = ', '.join(self._supportedmechanism)
                msg = self._logger.resolveString(265, mechanisms, ', '.join(features))
        except SMTPException as e:
            msg = self._logger.resolveString(264, mechanism, getExceptionMessage(e))
        if self._debug:
            self._logger.logp(SEVERE, self._cls, mtd, msg)
        raise AuthenticationFailedException(msg, self)

    def _sendHttpMailMessage(self, message):
        mtd = '_sendHttpMailMessage'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 271, message.Subject)
        request = getHttpRequest(self._getProvider(), SMTP, 'Request')
        if request:
            executeHttpRequest(self, self._logger, self._debug, self._cls, 272, self._server, message, request)
            subrequest = getHttpRequest(self._getProvider(), SMTP, 'SubRequest')
            if subrequest:
                executeHttpRequest(self, self._logger, self._debug, self._cls, 272, self._server, message, subrequest)
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
        except SMTPSenderRefused as e:
            msg = self._logger.resolveString(282, message.Subject, getExceptionMessage(e))
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise MailException(msg, self)
        except SMTPRecipientsRefused as e:
            msg = self._logger.resolveString(282, message.Subject, getExceptionMessage(e))
            # TODO: return SendMailMessageFailedException in place of MailException
            # TODO: error = SendMailMessageFailedException(msg, self)
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, msg)
            raise MailException(msg, self)
        except SMTPDataError as e:
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
        if isinstance(reply, bytes):
            reply = reply.decode('ascii', 'ignore')
        return reply

    def _getTokenParameters(self, mechanism, host):
        if self._tokens is None:
            self._tokens = getConfiguration(self._ctx, g_identifier).getByName('Tokens')
        return self._tokens.get(host, self._tokens.get(mechanism))

    def _getServerFeatures(self, server):
        return server.esmtp_features.get('auth').strip().split() if server.has_extn('auth') else ()

    def _getProvider(self):
        return self._providers.getByName(self._provider)

