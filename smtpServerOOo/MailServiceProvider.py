#!
# -*- coding: utf_8 -*-

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

from unolib import getOAuth2
from unolib import getConfiguration
from unolib import getInterfaceTypes
from unolib import getExceptionMessage

from smtpserver import logMessage
from smtpserver import getMessage
g_message = 'MailServiceProvider'

import sys
import smtplib
import imaplib
import poplib
import base64

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_providerImplName = 'org.openoffice.pyuno.MailServiceProvider2'
g_messageImplName = 'org.openoffice.pyuno.MailMessage2'


class SmtpService(unohelper.Base,
                  XSmtpService2):
    def __init__(self, ctx):
        self.ctx = ctx
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
        if self.isConnected():
            raise AlreadyConnectedException()
        unotype = uno.getTypeByName('com.sun.star.uno.XCurrentContext')
        if unotype not in getInterfaceTypes(context):
            raise IllegalArgumentException()
        unotype = uno.getTypeByName('com.sun.star.mail.XAuthenticator')
        if unotype not in getInterfaceTypes(authenticator):
            raise IllegalArgumentException()
        server = context.getValueByName('ServerName')
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType').upper()
        error = self._setServer(connection, server, port, timeout)
        if error is not None:
            raise error
        authentication = context.getValueByName('AuthenticationType').upper()
        error = self._doLogin(authentication, authenticator, server)
        if error is not None:
            raise error
        self._context = context
        for listener in self._listeners:
            listener.connected(self._notify)
        print("SmtpService.connect() 3")

    def _setServer(self, connection, host, port, timeout):
        error = None
        try:
            if connection == 'SSL':
                server = smtplib.SMTP_SSL(host=host, port=port, timeout=timeout)
            else:
                server = smtplib.SMTP(host=host, port=port, timeout=timeout)
            if connection == 'TLS':
                server.starttls()
        except smtplib.SMTPConnectError as e:
            msg = getMessage(self.ctx, g_message, 111, getExceptionMessage(e))
            error = ConnectException(msg, self)
        except smtplib.SMTPException as e:
            msg = getMessage(self.ctx, g_message, 111, getExceptionMessage(e))
            error = UnknownHostException(msg, self)
        except Exception as e:
            msg = getMessage(self.ctx, g_message, 111, getExceptionMessage(e))
            error = MailException(msg, self)
        else:
            self._server = server
        return error

    def _doLogin(self, authentication, authenticator, server):
        error = None
        if authentication == 'LOGIN':
            user = authenticator.getUserName()
            password = authenticator.getPassword()
            if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                user = user.encode('ascii')
                password = password.encode('ascii')
            try:
                self._server.login(user, password)
            except Exception as e:
                msg = getMessage(self.ctx, g_message, 121, getExceptionMessage(e))
                error = AuthenticationFailedException(msg, self)
        elif authentication == 'OAUTH2':
            user = authenticator.getUserName()
            print("SmtpService._doLogin() 2")
            token = getOAuth2Token(self.ctx, self._sessions, server, user, True)
            try:
                self._server.docmd('AUTH', 'XOAUTH2 %s' % token)
            except Exception as e:
                msg = getMessage(self.ctx, g_message, 121, getExceptionMessage(e))
                error = AuthenticationFailedException(msg, self)
        return error

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

    def sendMailMessage(self, message):
        COMMASPACE = ', '
        recipients = message.getRecipients()
        sendermail = message.SenderAddress
        sendername = message.SenderName
        subject = message.Subject
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
        if (len(attachments)):
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
        mailerstring = "LibreOffice via Caolan's mailmerge component"
        try:
            configuration = getConfiguration(self.ctx, '/org.openoffice.Setup/Product')
            name = configuration.getByName('ooName')
            version = configuration.getByName('ooSetupVersion')
            mailerstring = "%s %s via Caolan's mailmerge component" % (name, version)
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
        self.server.sendmail(sendermail, truerecipients, msg.as_string())


class ImapService(unohelper.Base,
                  XMailService2):
    def __init__(self, ctx):
        self.ctx = ctx
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
        #xConnectionContext = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext", self.ctx)
        #xConnectionContext.setPropertyValue("MailServiceType", IMAP)
        #xAuthenticator = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.Authenticator", self.ctx)
        self._context = context
        server = context.getValueByName('ServerName')
        port = context.getValueByName('Port')
        connection = context.getValueByName('ConnectionType')
        authentication = context.getValueByName('AuthenticationType')
        if connection.upper() == 'SSL':
            self._server = imaplib.IMAP4_SSL(host=server, port=port)
        else:
            self._server = imaplib.IMAP4(host=server, port=port)
        if connection.upper() == 'TLS':
            self._server.starttls()
        if authentication.upper() == 'LOGIN':
            user = authenticator.getUserName()
            password = authenticator.getPassword()
            if user != '':
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    if password != '':
                        password = password.encode('ascii')
                self._server.login(user, password)
        elif authentication.upper() == 'OAUTH2':
            user = authenticator.getUserName()
            token = getOAuth2Token(self.ctx, self._sessions, server, user)
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
        self.ctx = ctx
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
        #xConnectionContext = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext", self.ctx)
        #xConnectionContext.setPropertyValue("MailServiceType", POP3)
        #xAuthenticator = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.Authenticator", self.ctx)
        self._context = context
        server = context.getValueByName('ServerName')
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType')
        authentication = context.getValueByName('AuthenticationType')
        if connection.upper() == 'SSL':
            self._server = poplib.POP3_SSL(host=server, port=port, timeout=timeout)
        else:
            self._server = poplib.POP3(host=server, port=port, timeout=timeout)
        if connection.upper() == 'TLS':
            self._server.stls()
        if authentication.upper() == 'LOGIN':
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
        self.ctx = ctx

    def create(self, service):
        if service == SMTP:
            return SmtpService(self.ctx)
        elif service == POP3:
            return Pop3Service(self.ctx)
        elif service == IMAP:
            return ImapService(self.ctx)
        else:
            e = self._getNoMailServiceProviderException(101, service)
            logMessage(self.ctx, SEVERE, e.Message, 'MailServiceProvider', 'create()')
            raise e

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_providerImplName, service)
    def getImplementationName(self):
        return g_providerImplName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_providerImplName)

    def _getNoMailServiceProviderException(self, code, format):
        e = NoMailServiceProviderException()
        e.Message = getMessage(self.ctx, g_message, code, format)
        e.Context = self
        return e


class MailMessage(unohelper.Base,
                  XMailMessage):
    def __init__(self, ctx, to='', sender='', subject='', body=None, attachment=None):
        self.ctx = ctx
        self._recipients = [to]
        self._ccrecipients = []
        self._bccrecipients = []
        self._attachments = []
        if attachment is not None:
            self._attachments.append(attachment)
        self._sendername, self._senderaddress = parseaddr(sender)
        self._replytoaddress = sender
        self._subject = subject
        self._body = body

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
                                         ('com.sun.star.mail.MailServiceProvider', ), )

g_ImplementationHelper.addImplementation(MailMessage,
                                         g_messageImplName,
                                         ('com.sun.star.mail.MailMessage', ), )


def getOAuth2Token(ctx, sessions, server, user, encode=False):
    key = '%s/%s' % (server, user)
    if key not in sessions:
        sessions[key] = getOAuth2(ctx, server, user)
    token = sessions[key].getToken('')
    authstring = 'user=%s\1auth=Bearer %s\1\1' % (user, token)
    if encode:
        authstring = base64.b64encode(authstring.encode("ascii"))
    return 
