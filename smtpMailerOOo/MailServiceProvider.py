#!
# -*- coding: utf_8 -*-

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

from com.sun.star.lang import XServiceInfo
from com.sun.star.mail import XMailServiceProvider2
from com.sun.star.mail import XMailService2
from com.sun.star.mail import XSmtpService2
from com.sun.star.mail import XMailMessage
from com.sun.star.lang import EventObject

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import POP3
from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import NoMailServiceProviderException
from com.sun.star.mail import SendMailMessageFailedException
from com.sun.star.mail import MailException
from com.sun.star.auth import AuthenticationFailedException 
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.io import AlreadyConnectedException
from com.sun.star.io import UnknownHostException
from com.sun.star.io import ConnectException

from email.mime.base import MIMEBase
from email.message import Message
from email.charset import Charset
from email.charset import QP
from email.encoders import encode_base64
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from email.utils import parseaddr

from smtpmailer import getConfiguration
from smtpmailer import getExceptionMessage
from smtpmailer import getMessage
from smtpmailer import getOAuth2
from smtpmailer import getOAuth2Token
from smtpmailer import hasInterface
from smtpmailer import isDebugMode
from smtpmailer import logMessage
from smtpmailer import smtplib

g_message = 'MailServiceProvider'

import re
import sys
import six
import imaplib
import poplib
import base64
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_providerImplName = 'com.sun.star.mail.MailServiceProvider2'
g_messageImplName = 'com.sun.star.mail.MailMessage2'
#g_providerImplName = 'org.openoffice.pyuno.MailServiceProvider2'
#g_messageImplName = 'org.openoffice.pyuno.MailMessage2'


class SmtpService(unohelper.Base,
                  XSmtpService2):
    def __init__(self, ctx):
        if isDebugMode():
            msg = getMessage(ctx, g_message, 121)
            logMessage(ctx, INFO, msg, 'SmtpService', '__init__()')
        self._ctx = ctx
        self._listeners = []
        self._supportedconnection = ('Insecure', 'Ssl', 'Tls')
        self._supportedauthentication = ('None', 'Login', 'OAuth2')
        self._server = None
        self._context = None
        self._sessions = {}
        self._notify = EventObject(self)
        if isDebugMode():
            msg = getMessage(ctx, g_message, 122)
            logMessage(ctx, INFO, msg, 'SmtpService', '__init__()')

    def addConnectionListener(self, listener):
        self._listeners.append(listener)

    def removeConnectionListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def getSupportedConnectionTypes(self):
        return self._supportedconnection

    def getSupportedAuthenticationTypes(self):
        return self._supportedauthentication

    def connect(self, context, authenticator):
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 131)
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
                msg = getMessage(self._ctx, g_message, 132, error.Message)
                logMessage(self._ctx, SEVERE, msg, 'SmtpService', 'connect()')
            raise error
        authentication = context.getValueByName('AuthenticationType').title()
        if authentication != 'None':
            error = self._doLogin(authentication, authenticator, server)
            if error is not None:
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 133, error.Message)
                    logMessage(self._ctx, SEVERE, msg, 'SmtpService', 'connect()')
                raise error
        self._context = context
        for listener in self._listeners:
            listener.connected(self._notify)
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 134)
            logMessage(self._ctx, INFO, msg, 'SmtpService', 'connect()')

    def _setServer(self, context, host):
        error = None
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType').title()
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 141, connection)
            logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
        try:
            if connection == 'Ssl':
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 142, timeout)
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
                server = smtplib.SMTP_SSL(timeout=timeout)
            else:
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 143, timeout)
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
                server = smtplib.SMTP(timeout=timeout)
            if isDebugMode():
                server.set_debuglevel(1)
            code, reply = getReply(*server.connect(host=host, port=port))
            if isDebugMode():
                msg = getMessage(self._ctx, g_message, 144, (host, port, code, reply))
                logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
            if code != 220:
                msg = getMessage(self._ctx, g_message, 146, reply)
                error = ConnectException(msg, self)
            elif connection == 'Tls':
                code, reply = getReply(*server.starttls())
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 145, (code, reply))
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
                if code != 220:
                    msg = getMessage(self._ctx, g_message, 146, reply)
                    error = ConnectException(msg, self)
        except smtplib.SMTPConnectError as e:
            msg = getMessage(self._ctx, g_message, 146, getExceptionMessage(e))
            error = ConnectException(msg, self)
        except smtplib.SMTPException as e:
            msg = getMessage(self._ctx, g_message, 146, getExceptionMessage(e))
            error = UnknownHostException(msg, self)
        except Exception as e:
            msg = getMessage(self._ctx, g_message, 146, getExceptionMessage(e))
            error = MailException(msg, self)
        else:
            self._server = server
        if isDebugMode() and error is None: 
            msg = getMessage(self._ctx, g_message, 147, (connection, reply))
            logMessage(self._ctx, INFO, msg, 'SmtpService', '_setServer()')
        return error

    def _doLogin(self, authentication, authenticator, server):
        error = None
        user = authenticator.getUserName()
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 151, authentication)
            logMessage(self._ctx, INFO, msg, 'SmtpService', '_doLogin()')
        try:
            if authentication == 'Login':
                password = authenticator.getPassword()
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    password = password.encode('ascii')
                code, reply = getReply(*self._server.login(user, password))
                if isDebugMode():
                    pwd = '*' * len(password)
                    msg = getMessage(self._ctx, g_message, 152, (user, pwd, code, reply))
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_doLogin()')
            elif authentication == 'Oauth2':
                token = getToken(self._ctx, self, server, user, True)
                self._server.ehlo_or_helo_if_needed()
                code, reply = getReply(*self._server.docmd('AUTH', 'XOAUTH2 %s' % token))
                if code != 235:
                    msg = getMessage(self._ctx, g_message, 154, reply)
                    error = AuthenticationFailedException(msg, self)
                if isDebugMode():
                    msg = getMessage(self._ctx, g_message, 153, (code, reply))
                    logMessage(self._ctx, INFO, msg, 'SmtpService', '_doLogin()')
        except Exception as e:
            msg = getMessage(self._ctx, g_message, 154, getExceptionMessage(e))
            error = AuthenticationFailedException(msg, self)
        if isDebugMode() and error is None:
            msg = getMessage(self._ctx, g_message, 155, (authentication, reply))
            logMessage(self._ctx, INFO, msg, 'SmtpService', '_doLogin()')
        return error

    def isConnected(self):
        return self._server is not None

    def disconnect(self):
        if self.isConnected():
            if isDebugMode():
                msg = getMessage(self._ctx, g_message, 171)
                logMessage(self._ctx, INFO, msg, 'SmtpService', 'disconnect()')
            self._server.quit()
            self._server = None
            self._context = None
            for listener in self._listeners:
                listener.disconnected(self._notify)
            if isDebugMode():
                msg = getMessage(self._ctx, g_message, 172)
                logMessage(self._ctx, INFO, msg, 'SmtpService', 'disconnect()')

    def getCurrentConnectionContext(self):
        return self._context

    def sendMailMessage(self, message):
        COMMASPACE = ', '
        recipients = message.getRecipients()
        sendermail = message.SenderAddress
        sendername = message.SenderName
        subject = message.Subject
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 161, subject)
            logMessage(self._ctx, INFO, msg, 'SmtpService', 'sendMailMessage()')
        ccrecipients = message.getCcRecipients()
        bccrecipients = message.getBccRecipients()
        attachments = message.getAttachments()
        textmsg = Message()
        content = message.Body
        flavors = content.getTransferDataFlavors()
        #Use first flavor that's sane for an email body
        for flavor in flavors:
            if flavor.MimeType.find('text/html') != -1 or flavor.MimeType.find('text/plain') != -1:
                textbody = content.getTransferData(flavor)
                if len(textbody):
                    mimeEncoding = re.sub('charset=.*', 'charset=UTF-8', flavor.MimeType)
                    if mimeEncoding.find('charset=UTF-8') == -1:
                        mimeEncoding = mimeEncoding + '; charset=UTF-8'
                    textmsg['Content-Type'] = mimeEncoding
                    textmsg['MIME-Version'] = '1.0'
                    try:
                        #it's a string, get it as utf-8 bytes
                        textbody = textbody.encode('utf-8')
                    except:
                        #it's a bytesequence, get raw bytes
                        textbody = textbody.value
                    if sys.version >= '3':
                        if sys.version_info.minor < 3 or (sys.version_info.minor == 3 and sys.version_info.micro <= 1):
                            #http://stackoverflow.com/questions/9403265/how-do-i-use-python-3-2-email-module-to-send-unicode-messages-encoded-in-utf-8-w
                            #see http://bugs.python.org/16564, etc. basically it now *seems* to be all ok
                            #in python 3.3.2 onwards, but a little busted in 3.3.0
                            textbody = textbody.decode('iso8859-1')
                        else:
                            textbody = textbody.decode('utf-8')
                        c = Charset('utf-8')
                        c.body_encoding = QP
                        textmsg.set_payload(textbody, c)
                    else:
                        textmsg.set_payload(textbody)
                break
        if len(attachments):
            msg = MIMEMultipart()
            msg.epilogue = ''
            msg.attach(textmsg)
        else:
            msg = textmsg
        header = Header(sendername, 'utf-8')
        header.append('<'+sendermail+'>','us-ascii')
        msg['Subject'] = subject
        msg['From'] = header
        msg['To'] = COMMASPACE.join(recipients)
        if len(ccrecipients):
            msg['Cc'] = COMMASPACE.join(ccrecipients)
        if message.ReplyToAddress != '':
            msg['Reply-To'] = message.ReplyToAddress
        mailerstring = "LibreOffice via smtpMailerOOo component"
        try:
            configuration = getConfiguration(self._ctx, '/org.openoffice.Setup/Product')
            name = configuration.getByName('ooName')
            version = configuration.getByName('ooSetupVersion')
            mailerstring = "%s %s via via smtpMailerOOo component" % (name, version)
        except:
            pass
        msg['X-Mailer'] = mailerstring
        msg['Date'] = formatdate(localtime=True)
        for attachment in attachments:
            content = attachment.Data
            flavors = content.getTransferDataFlavors()
            flavor = flavors[0]
            ctype = flavor.MimeType
            maintype, subtype = ctype.split('/', 1)
            msgattachment = MIMEBase(maintype, subtype)
            data = content.getTransferData(flavor)
            msgattachment.set_payload(data.value)
            encode_base64(msgattachment)
            fname = attachment.ReadableName
            try:
                msgattachment.add_header('Content-Disposition', 'attachment', \
                    filename=fname)
            except:
                msgattachment.add_header('Content-Disposition', 'attachment', \
                    filename=('utf-8','',fname))
            msg.attach(msgattachment)
        uniquer = {}
        for key in recipients:
            uniquer[key] = True
        if len(ccrecipients):
            for key in ccrecipients:
                uniquer[key] = True
        if len(bccrecipients):
            for key in bccrecipients:
                uniquer[key] = True
        truerecipients = uniquer.keys()
        error = None
        try:
            refused = self._server.sendmail(sendermail, truerecipients, msg.as_string())
        except smtplib.SMTPSenderRefused as e:
            msg = getMessage(self._ctx, g_message, 162, (subject, getExceptionMessage(e)))
            error = MailException(msg, self)
        except smtplib.SMTPRecipientsRefused as e:
            msg = getMessage(self._ctx, g_message, 163, (subject, getExceptionMessage(e)))
            # TODO: return SendMailMessageFailedException in place of MailException
            # TODO: error = SendMailMessageFailedException(msg, self)
            error = MailException(msg, self)
        except smtplib.SMTPDataError as e:
            msg = getMessage(self._ctx, g_message, 163, (subject, getExceptionMessage(e)))
            error = MailException(msg, self)
        except Exception as e:
            msg = getMessage(self._ctx, g_message, 163, (subject, getExceptionMessage(e)))
            error = MailException(msg, self)
        else:
            if len(refused) > 0:
                for address, result in refused.items():
                    code, reply = getReply(*result)
                    msg = getMessage(self._ctx, g_message, 164, (subject, address, code, reply))
                    logMessage(self._ctx, SEVERE, msg, 'SmtpService', 'sendMailMessage()')
            elif isDebugMode():
                msg = getMessage(self._ctx, g_message, 165, subject)
                logMessage(self._ctx, INFO, msg, 'SmtpService', 'sendMailMessage()')
        if error is not None:
            logMessage(self._ctx, SEVERE, error.Message, 'SmtpService', 'sendMailMessage()')
            raise error


class ImapService(unohelper.Base,
                  XMailService2):
    def __init__(self, ctx):
        self._ctx = ctx
        self._listeners = []
        self._supportedconnection = ('Insecure', 'Ssl', 'Tls')
        self._supportedauthentication = ('None', 'Login', 'OAuth2')
        self._server = None
        self._context = None
        self._sessions = {}
        self._notify = EventObject(self)

    def addConnectionListener(self, listener):
        self._listeners.append(listener)

    def removeConnectionListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def getSupportedConnectionTypes(self):
        return self._supportedconnection

    def getSupportedAuthenticationTypes(self):
        return self._supportedauthentication

    def connect(self, context, authenticator):
        #xConnectionContext = self._ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext", self._ctx)
        #xConnectionContext.setPropertyValue("MailServiceType", IMAP)
        #xAuthenticator = self._ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.Authenticator", self._ctx)
        self._context = context
        server = context.getValueByName('ServerName')
        port = context.getValueByName('Port')
        connection = context.getValueByName('ConnectionType').title()
        authentication = context.getValueByName('AuthenticationType').title()
        if connection == 'Ssl':
            self._server = imaplib.IMAP4_SSL(host=server, port=port)
        else:
            self._server = imaplib.IMAP4(host=server, port=port)
        if connection == 'Tls':
            self._server.starttls()
        if authentication == 'Login':
            user = authenticator.getUserName()
            password = authenticator.getPassword()
            if user != '':
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    if password != '':
                        password = password.encode('ascii')
                self._server.login(user, password)
        elif authentication == 'Oauth2':
            user = authenticator.getUserName()
            token = getToken(self._ctx, self, server, user)
            self._server.authenticate('XOAUTH2', lambda x: token)
        for listener in self._listeners:
            listener.connected(self._notify)

    def disconnect(self):
        if self.isConnected():
            self._server.logout()
            self._server = None
            self._context = None
        for listener in self._listeners:
            listener.disconnected(self._notify)

    def isConnected(self):
        return self._server is not None

    def getCurrentConnectionContext(self):
        return self._context


class Pop3Service(unohelper.Base,
                  XMailService2):
    def __init__(self, ctx):
        self._ctx = ctx
        self._listeners = []
        self._supportedconnection = ('Insecure', 'Ssl', 'Tls')
        self._supportedauthentication = ('None', 'Login')
        self._server = None
        self._context = None
        self._notify = EventObject(self)

    def addConnectionListener(self, listener):
        self.listeners.append(listener)

    def removeConnectionListener(self, listener):
        if listener in self._listeners:
            self.listeners.remove(listener)

    def getSupportedConnectionTypes(self):
        return self._supportedconnection

    def getSupportedAuthenticationTypes(self):
        return self._supportedauthentication

    def connect(self, context, authenticator):
        #xConnectionContext = self._ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext", self._ctx)
        #xConnectionContext.setPropertyValue("MailServiceType", POP3)
        #xAuthenticator = self._ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.Authenticator", self._ctx)
        self._context = context
        server = context.getValueByName('ServerName')
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType').title()
        authentication = context.getValueByName('AuthenticationType').title()
        if connection == 'Ssl':
            self._server = poplib.POP3_SSL(host=server, port=port, timeout=timeout)
        else:
            self._server = poplib.POP3(host=server, port=port, timeout=timeout)
        if connection == 'Tls':
            self._server.stls()
        if authentication == 'Login':
            user = authenticator.getUserName()
            password = authenticator.getPassword()
            if user != '':
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    if password != '':
                        password = password.encode('ascii')
                self._server.user(user)
                self._server.pass_(password)
        for listener in self._listeners:
            listener.connected(self._notify)

    def disconnect(self):
        if self.isConnected():
            self._server.quit()
            self._server = None
            self._context = None
        for listener in self._listeners:
            listener.disconnected(self._notify)

    def isConnected(self):
        return self._server is not None

    def getCurrentConnectionContext(self):
        return self._context


class MailServiceProvider(unohelper.Base,
                          XMailServiceProvider2,
                          XServiceInfo):
    def __init__(self, ctx):
        if isDebugMode():
            msg = getMessage(ctx, g_message, 101)
            logMessage(ctx, INFO, msg, 'MailServiceProvider', '__init__()')
        self._ctx = ctx
        if isDebugMode():
            msg = getMessage(ctx, g_message, 102)
            logMessage(ctx, INFO, msg, 'MailServiceProvider', '__init__()')

    def create(self, mailtype):
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 111, mailtype.value)
            logMessage(self._ctx, INFO, msg, 'MailServiceProvider', 'create()')
        if mailtype == SMTP:
            service = SmtpService(self._ctx)
        elif mailtype == POP3:
            service = Pop3Service(self._ctx)
        elif mailtype == IMAP:
            service = ImapService(self._ctx)
        else:
            e = self._getNoMailServiceProviderException(113, mailtype)
            logMessage(self._ctx, SEVERE, e.Message, 'MailServiceProvider', 'create()')
            raise e
        if isDebugMode():
            msg = getMessage(self._ctx, g_message, 112, mailtype.value)
            logMessage(self._ctx, INFO, msg, 'MailServiceProvider', 'create()')
        return service

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_providerImplName, service)
    def getImplementationName(self):
        return g_providerImplName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_providerImplName)

    def _getNoMailServiceProviderException(self, code, format):
        e = NoMailServiceProviderException()
        e.Message = getMessage(self._ctx, g_message, code, format)
        e.Context = self
        return e


class MailMessage(unohelper.Base,
                  XMailMessage):
    def __init__(self, ctx, to='', sender='', subject='', body=None, attachment=None):
        self._ctx = ctx
        self._recipients = [to]
        self._ccrecipients = []
        self._bccrecipients = []
        self._attachments = []
        if attachment is not None:
            self._attachments.append(attachment)
        self._sname, self._saddress = parseaddr(sender)
        self.ReplyToAddress = sender
        self.Subject = subject
        self.Body = body

    @property
    def SenderName(self):
        return self._sname

    @property
    def SenderAddress(self):
        return self._saddress

    def addRecipient(self, recipient):
        self._recipients.append(recipient)

    def addCcRecipient(self, ccrecipient):
        self._ccrecipients.append(ccrecipient)

    def addBccRecipient(self, bccrecipient):
        self._bccrecipients.append(bccrecipient)

    def getRecipients(self):
        return tuple(self._recipients)

    def getCcRecipients(self):
        return tuple(self._ccrecipients)

    def getBccRecipients(self):
        return tuple(self._bccrecipients)

    def addAttachment(self, attachment):
        self._attachments.append(attachment)

    def getAttachments(self):
        return tuple(self._attachments)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_messageImplName, service)
    def getImplementationName(self):
        return g_messageImplName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_messageImplName)


g_ImplementationHelper.addImplementation(MailServiceProvider,
                                         g_providerImplName,
                                         ('com.sun.star.mail.MailServiceProvider2', ), )

g_ImplementationHelper.addImplementation(MailMessage,
                                         g_messageImplName,
                                         ('com.sun.star.mail.MailMessage2', ), )


def getToken(ctx, source, url, user, encode=False):
    token = getOAuth2Token(ctx, source, url, user)
    authstring = 'user=%s\1auth=Bearer %s\1\1' % (user, token)
    if encode:
        authstring = base64.b64encode(authstring.encode('ascii')).decode('ascii')
    return authstring

def getReply(code, reply):
    if isinstance(reply, six.binary_type):
        reply = reply.decode('ascii')
    return code, reply
