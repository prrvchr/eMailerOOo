#!
# -*- coding: utf-8 -*-

# *************************************************************
#  
#  Licensed to the Apache Software Foundation (ASF) under one
#  or more contributor license agreements.  See the NOTICE file
#  distributed with this work for additional information
#  regarding copyright ownership.  The ASF licenses this file
#  to you under the Apache License, Version 2.0 (the
#  "License"); you may not use this file except in compliance
#  with the License.  You may obtain a copy of the License at
#  
#    http://www.apache.org/licenses/LICENSE-2.0
#  
#  Unless required by applicable law or agreed to in writing,
#  software distributed under the License is distributed on an
#  "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
#  KIND, either express or implied.  See the License for the
#  specific language governing permissions and limitations
#  under the License.
#  
# *************************************************************

# Caolan McNamara caolanm@redhat.com
# a simple email mailmerge component

# manual installation for hackers, not necessary for users
# cp mailmerge.py /usr/lib/openoffice.org2.0/program
# cd /usr/lib/openoffice.org2.0/program
# ./unopkg add --shared mailmerge.py
# edit ~/.openoffice.org2/user/registry/data/org/openoffice/Office/Writer.xcu
# and change EMailSupported to as follows...
#  <prop oor:name="EMailSupported" oor:type="xs:boolean">
#   <value>true</value>
#  </prop>

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

from com.sun.star.logging.LogLevel import ALL
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import MailException

from .smtpservice import SmtpService

from .. import smtplib

from ..unotool import getExceptionMessage
from ..unotool import hasInterface

import sys
import six
import base64
from collections import OrderedDict
import traceback


class SmtpBaseService(SmtpService):

    def connect(self, context, authenticator):
        self._logger.logprb(INFO, 'SmtpService', 'connect()', 221)
        if self.isConnected():
            raise AlreadyConnectedException()
        if not hasInterface(context, 'com.sun.star.uno.XCurrentContext'):
            raise IllegalArgumentException()
        if not hasInterface(authenticator, 'com.sun.star.mail.XAuthenticator'):
            raise IllegalArgumentException()
        server = context.getValueByName('ServerName')
        error = self._setServer(context, server)
        if error is not None:
            self._logger.logprb(SEVERE, 'SmtpService', 'connect()', 222, error.Message)
            raise error
        authentication = context.getValueByName('AuthenticationType').title()
        if authentication != 'None':
            error = self._doLogin(authentication, authenticator, server)
            if error is not None:
                self._logger.logprb(SEVERE, 'SmtpService', 'connect()', 223, error.Message)
                raise error
        self._context = context
        for listener in self._listeners:
            listener.connected(self._notify)
        self._logger.logprb(INFO, 'SmtpService', 'connect()', 224)

    def _setServer(self, context, host):
        error = None
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType').title()
        self._logger.logprb(INFO, 'SmtpService', '_setServer()', 231, connection)
        try:
            if connection == 'Ssl':
                self._logger.logprb(INFO, 'SmtpService', '_setServer()', 232, timeout)
                server = smtplib.SMTP_SSL(timeout=timeout)
            else:
                self._logger.logprb(INFO, 'SmtpService', '_setServer()', 233, timeout)
                server = smtplib.SMTP(timeout=timeout)
            if self._logger.Level == ALL:
                server.set_debuglevel(1)
            code, reply = _getReply(*server.connect(host=host, port=port))
            self._logger.logprb(INFO, 'SmtpService', '_setServer()', 234, host, port, code, reply)
            if code != 220:
                msg = self.logger.resolveString(236, reply)
                error = ConnectException(msg, self)
            elif connection == 'Tls':
                code, reply = _getReply(*server.starttls())
                self._logger.logprb(INFO, 'SmtpService', '_setServer()', 235, code, reply)
                if code != 220:
                    msg = self.logger.resolveString(236, reply)
                    error = ConnectException(msg, self)
        except smtplib.SMTPConnectError as e:
            msg = self.logger.resolveString(236, getExceptionMessage(e))
            error = ConnectException(msg, self)
        except smtplib.SMTPException as e:
            msg = self.logger.resolveString(236, getExceptionMessage(e))
            error = UnknownHostException(msg, self)
        except Exception as e:
            msg = self.logger.resolveString(236, getExceptionMessage(e))
            error = MailException(msg, self)
        else:
            self._server = server
        if error is None:
            self._logger.logprb(INFO, 'SmtpService', '_setServer()', 237, connection, reply)
        return error

    def _doLogin(self, authentication, authenticator, server):
        error = None
        user = authenticator.getUserName()
        self._logger.logprb(INFO, 'SmtpService', '_doLogin()', 241, authentication)
        try:
            if authentication == 'Login':
                password = authenticator.getPassword()
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    password = password.encode('ascii')
                code, reply = _getReply(*self._server.login(user, password))
                pwd = '*' * len(password)
                self._logger.logprb(INFO, 'SmtpService', '_doLogin()', 242, user, pwd, code, reply)
            elif authentication == 'Oauth2':
                token = _getToken(self._ctx, self, server, user, True)
                self._server.ehlo_or_helo_if_needed()
                code, reply = _getReply(*self._server.docmd('AUTH', 'XOAUTH2 %s' % token))
                if code != 235:
                    msg = self._logger.resolveString(244, reply)
                    error = AuthenticationFailedException(msg, self)
                self._logger.logprb(INFO, 'SmtpService', '_doLogin()', 243, code, reply)
        except Exception as e:
            msg = self._logger.resolveString(244, getExceptionMessage(e))
            error = AuthenticationFailedException(msg, self)
        if error is None:
            self._logger.logprb(INFO, 'SmtpService', '_doLogin()', 245, authentication, reply)
        return error

    def isConnected(self):
        return self._server is not None

    def disconnect(self):
        if self.isConnected():
            self._logger.logprb(INFO, 'SmtpService', 'disconnect()', 261)
            self._server.quit()
            self._server = None
            self._context = None
            for listener in self._listeners:
                listener.disconnected(self._notify)
            self._logger.logprb(INFO, 'SmtpService', 'disconnect()', 262)

    def sendMailMessage(self, message):
        recipients = _getRecipients(message)
        error = None
        try:
            refused = self._server.sendmail(message.SenderAddress, recipients, message.asString(False))
        except smtplib.SMTPSenderRefused as e:
            msg = self._logger.resolveString(252, message.Subject, getExceptionMessage(e))
            error = MailException(msg, self)
        except smtplib.SMTPRecipientsRefused as e:
            msg = self._logger.resolveString(253, message.Subject, getExceptionMessage(e))
            # TODO: return SendMailMessageFailedException in place of MailException
            # TODO: error = SendMailMessageFailedException(msg, self)
            error = MailException(msg, self)
        except smtplib.SMTPDataError as e:
            msg = self._logger.resolveString(253, message.Subject, getExceptionMessage(e))
            error = MailException(msg, self)
        except Exception as e:
            msg = self._logger.resolveString(253, message.Subject, getExceptionMessage(e))
            error = MailException(msg, self)
        else:
            if len(refused) > 0:
                for address, result in refused.items():
                    code, reply = _getReply(*result)
                    self._logger.logprb(SEVERE, 'SmtpService', 'sendMailMessage()', 254, message.Subject, address, code, reply)
            else:
                self._logger.logprb(INFO, 'SmtpService', 'sendMailMessage()', 255, message.Subject)
        if error is not None:
            self._logger.logp(SEVERE, error.Message, 'SmtpService', 'sendMailMessage()')
            raise error

def _getToken(ctx, source, url, user, encode=False):
    token = getOAuth2Token(ctx, source, url, user)
    authstring = 'user=%s\1auth=Bearer %s\1\1' % (user, token)
    if encode:
        authstring = base64.b64encode(authstring.encode('ascii')).decode('ascii')
    return authstring

def _getReply(code, reply):
    if isinstance(reply, six.binary_type):
        reply = reply.decode('ascii')
    return code, reply

def _getRecipients(message):
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
