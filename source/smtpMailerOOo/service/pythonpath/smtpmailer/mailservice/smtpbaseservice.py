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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import MailException

from .smtpservice import SmtpService

from .. import smtplib

from ..unotool import getExceptionMessage
from ..unotool import hasInterface

from ..logger import getMessage
from ..logger import isDebugMode
from ..logger import logMessage

g_message = 'MailServiceProvider'

import sys
import six
import base64
from collections import OrderedDict
import traceback


class SmtpBaseService(SmtpService):

    def connect(self, context, authenticator):
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 221)
            logMessage(self._ctx, INFO, msg, 'SmtpService', 'connect()')
        if self.isConnected():
            raise AlreadyConnectedException()
        if not hasInterface(context, 'com.sun.star.uno.XCurrentContext'):
            raise IllegalArgumentException()
        if not hasInterface(authenticator, 'com.sun.star.mail.XAuthenticator'):
            raise IllegalArgumentException()
        server = context.getValueByName('ServerName')
        error = self._setServer(context, server)
        if error is not None:
            if isDebugMode():
                msg = getMessage(self._ctx, g_message, 222, error.Message)
                logMessage(self._ctx, SEVERE, msg, 'SmtpService', 'connect()')
            raise error
        authentication = context.getValueByName('AuthenticationType').title()
        if authentication != 'None':
            error = self._doLogin(authentication, authenticator, server)
            if error is not None:
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 223, error.Message)
                    logMessage(self._ctx, SEVERE, msg, 'SmtpService', 'connect()')
                raise error
        self._context = context
        for listener in self._listeners:
            listener.connected(self._notify)
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 224)
            logMessage(self._ctx, INFO, msg, 'SmtpService', 'connect()')

    def _setServer(self, context, host):
        error = None
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType').title()
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 231, connection)
            logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
        try:
            if connection == 'Ssl':
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 232, timeout)
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
                server = smtplib.SMTP_SSL(timeout=timeout)
            else:
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 233, timeout)
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
                server = smtplib.SMTP(timeout=timeout)
            if isDebugMode():
                server.set_debuglevel(1)
            code, reply = _getReply(*server.connect(host=host, port=port))
            if isDebugMode():
                msg = getMessage(self._ctx, g_message, 234, (host, port, code, reply))
                logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
            if code != 220:
                msg = getMessage(self._ctx, g_message, 236, reply)
                error = ConnectException(msg, self)
            elif connection == 'Tls':
                code, reply = _getReply(*server.starttls())
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 235, (code, reply))
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
                if code != 220:
                    msg = getMessage(self._ctx, g_message, 236, reply)
                    error = ConnectException(msg, self)
        except smtplib.SMTPConnectError as e:
            msg = getMessage(self._ctx, g_message, 236, getExceptionMessage(e))
            error = ConnectException(msg, self)
        except smtplib.SMTPException as e:
            msg = getMessage(self._ctx, g_message, 236, getExceptionMessage(e))
            error = UnknownHostException(msg, self)
        except Exception as e:
            msg = getMessage(self._ctx, g_message, 236, getExceptionMessage(e))
            error = MailException(msg, self)
        else:
            self._server = server
        if isDebugMode() and error is None: 
            msg = getMessage(self._ctx, g_message, 237, (connection, reply))
            logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
        return error

    def _doLogin(self, authentication, authenticator, server):
        error = None
        user = authenticator.getUserName()
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 241, authentication)
            logMessage(self._ctx, INFO, msg, 'SmtpService', '_doLogin()')
        try:
            if authentication == 'Login':
                password = authenticator.getPassword()
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    password = password.encode('ascii')
                code, reply = _getReply(*self._server.login(user, password))
                if isDebugMode():
                    pwd = '*' * len(password)
                    msg = getMessage(self._ctx, g_message, 242, (user, pwd, code, reply))
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_doLogin()')
            elif authentication == 'Oauth2':
                token = _getToken(self._ctx, self, server, user, True)
                self._server.ehlo_or_helo_if_needed()
                code, reply = _getReply(*self._server.docmd('AUTH', 'XOAUTH2 %s' % token))
                if code != 235:
                    msg = getMessage(self._ctx, g_message, 244, reply)
                    error = AuthenticationFailedException(msg, self)
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 243, (code, reply))
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_doLogin()')
        except Exception as e:
            msg = getMessage(self._ctx, g_message, 244, getExceptionMessage(e))
            error = AuthenticationFailedException(msg, self)
        if isDebugMode() and error is None:
            msg = getMessage(self._ctx, g_message, 245, (authentication, reply))
            logMessage(self._ctx, INFO, msg, 'SmtpService', '_doLogin()')
        return error

    def isConnected(self):
        return self._server is not None

    def disconnect(self):
        if self.isConnected():
            if isDebugMode():
                msg = getMessage(self._ctx, g_message, 261)
                logMessage(self._ctx, INFO, msg, 'SmtpService', 'disconnect()')
            self._server.quit()
            self._server = None
            self._context = None
            for listener in self._listeners:
                listener.disconnected(self._notify)
            if isDebugMode():
                msg = getMessage(self._ctx, g_message, 262)
                logMessage(self._ctx, INFO, msg, 'SmtpService', 'disconnect()')

    def sendMailMessage(self, message):
        recipients = _getRecipients(message)
        error = None
        try:
            refused = self._server.sendmail(message.SenderAddress, recipients, message.asString(False))
        except smtplib.SMTPSenderRefused as e:
            msg = getMessage(self._ctx, g_message, 252, (message.Subject, getExceptionMessage(e)))
            error = MailException(msg, self)
        except smtplib.SMTPRecipientsRefused as e:
            msg = getMessage(self._ctx, g_message, 253, (message.Subject, getExceptionMessage(e)))
            # TODO: return SendMailMessageFailedException in place of MailException
            # TODO: error = SendMailMessageFailedException(msg, self)
            error = MailException(msg, self)
        except smtplib.SMTPDataError as e:
            msg = getMessage(self._ctx, g_message, 253, (message.Subject, getExceptionMessage(e)))
            error = MailException(msg, self)
        except Exception as e:
            msg = getMessage(self._ctx, g_message, 253, (message.Subject, getExceptionMessage(e)))
            error = MailException(msg, self)
        else:
            if len(refused) > 0:
                for address, result in refused.items():
                    code, reply = _getReply(*result)
                    msg = getMessage(self._ctx, g_message, 254, (message.Subject, address, code, reply))
                    logMessage(self._ctx, SEVERE, msg, 'SmtpService', 'sendMailMessage()')
            elif isDebugMode():
                msg = getMessage(self._ctx, g_message, 255, message.Subject)
                logMessage(self._ctx, INFO, msg, 'SmtpService', 'sendMailMessage()')
        if error is not None:
            logMessage(self._ctx, SEVERE, error.Message, 'SmtpService', 'sendMailMessage()')
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
