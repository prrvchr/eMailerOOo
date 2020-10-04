# Caolan McNamara caolanm@redhat.com
# a simple email mailmerge component

# manual installation for hackers, not necessary for users
# cp mailmerge.py /usr/lib/libreoffice/program
# cd /usr/lib/libreoffice/program
# ./unopkg add --shared mailmerge.py
# edit ~/.openoffice.org2/user/registry/data/org/openoffice/Office/Writer.xcu
# and change EMailSupported to as follows...
#  <prop oor:name="EMailSupported" oor:type="xs:boolean">
#   <value>true</value>
#  </prop>

from __future__ import print_function

import unohelper
import uno
import re
import os

#to implement com::sun::star::mail::XMailServiceProvider
#and
#to implement com.sun.star.mail.XMailMessage

from com.sun.star.mail import XMailServiceProvider
from com.sun.star.mail import XMailService
from com.sun.star.mail import XSmtpService
from com.sun.star.mail import XConnectionListener
from com.sun.star.mail import XAuthenticator
from com.sun.star.mail import XMailMessage
from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import POP3
from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.uno import XCurrentContext
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.lang import EventObject
from com.sun.star.lang import XServiceInfo
from com.sun.star.mail import SendMailMessageFailedException


from email.mime.base import MIMEBase
from email.message import Message
from email.charset import Charset
from email.charset import QP
from email.encoders import encode_base64
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from email.utils import parseaddr
#from socket import _GLOBAL_DEFAULT_TIMEOUT

import sys
import smtplib
import imaplib
import poplib

dbg = False

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_providerImplName = "org.openoffice.pyuno.MailServiceProvider2"
g_messageImplName = "org.openoffice.pyuno.MailMessage2"


#no stderr under windows, output to pymailmerge.log
#with no buffering
if dbg and os.name == 'nt':
    dbgout = open('pymailmerge.log', 'w', 0)
else:
    dbgout = sys.stderr

class PyMailSMTPService(unohelper.Base,
                        XSmtpService):
    def __init__(self, ctx):
        self.ctx = ctx
        self.listeners = []
        self.connectiontype = None
        self.authenticationtype = None
        self.supportedconnection = ('Insecure', 'Ssl', 'Tls')
        self.supportedauthentication = ('None', 'Login', 'OAuth2')
        self.server = None
        self.connectioncontext = None
        self.notify = EventObject(self)
        if dbg:
            print("PyMailSMTPService init", file=dbgout)
            print("python version is: " + sys.version, file=dbgout)
    def _getConfiguration(self, nodepath, update=False):
        config = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        service = "com.sun.star.configuration.ConfigurationUpdateAccess" if update else \
                  "com.sun.star.configuration.ConfigurationAccess"
        namedvalue = uno.createUnoStruct("com.sun.star.beans.NamedValue")
        namedvalue.Name = "nodepath"
        namedvalue.Value = nodepath
        return config.createInstanceWithArguments(service, (namedvalue,))
    def addConnectionListener(self, xListener):
        if dbg:
            print("PyMailSMTPService addConnectionListener", file=dbgout)
        self.listeners.append(xListener)
    def removeConnectionListener(self, xListener):
        if dbg:
            print("PyMailSMTPService removeConnectionListener", file=dbgout)
        self.listeners.remove(xListener)
    def getConnectionType(self):
        if dbg:
            print("PyMailSMTPService getConnectionType", file=dbgout)
        return self.connectiontype
    def getAuthenticationType(self):
        if dbg:
            print("PyMailSMTPService getAuthenticationType", file=dbgout)
        return self.authenticationtype
    def getSupportedConnectionTypes(self):
        if dbg:
            print("PyMailSMTPService getSupportedConnectionTypes", file=dbgout)
        return self.supportedauthentication
    #    return self.supportedconnection
    def getSupportedAuthenticationTypes(self):
        if dbg:
            print("PyMailSMTPService getSupportedAuthenticationTypes", file=dbgout)
        return self.supportedauthentication
    def connect(self, xConnectionContext, xAuthenticator):
        if dbg:
            print("PyMailSMTPService connect", file=dbgout)
        #mri = self.ctx.ServiceManager.createInstance("mytools.Mri")
        #mri.inspect(xConnectionContext)
        #mri.inspect(xAuthenticator)
        xConnectionContext = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext", self.ctx)
        xConnectionContext.setPropertyValue("MailServiceType", SMTP)
        xAuthenticator = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.Authenticator", self.ctx)

        self.connectioncontext = xConnectionContext
        server = xConnectionContext.getValueByName("ServerName")
        if dbg:
            print("ServerName: " + server, file=dbgout)
        port = int(xConnectionContext.getValueByName("Port"))
        if dbg:
            print("Port: " + str(port), file=dbgout)
        timeout = int(xConnectionContext.getValueByName("ConnectionTimeout"))
        if dbg:
            print("Timeout: " + str(timeout), file=dbgout)
        self.connectiontype = xConnectionContext.getValueByName("ConnectionType")
        if dbg:
            print("ConnectionType: " + self.connectiontype, file=dbgout)
        self.authenticationtype = xConnectionContext.getValueByName("AuthenticationType")
        if dbg:
            print("AuthenticationType: " + self.authenticationtype, file=dbgout)
        if self.connectiontype.upper() == 'SSL':
            self.server = smtplib.SMTP_SSL(host=server, port=port, timeout=timeout)
        else:
            self.server = smtplib.SMTP(host=server, port=port, timeout=timeout)
        if self.connectiontype.upper() == 'TLS':
            self.server.starttls()

        #stderr not available for us under windows, but
        #set_debuglevel outputs there, and so throw
        #an exception under windows on debugging mode
        #with this enabled
        if dbg and os.name != 'nt':
            self.server.set_debuglevel(1)
        if self.authenticationtype.upper() == 'LOGIN':
            user = xAuthenticator.getUserName()
            password = xAuthenticator.getPassword()
            if user != '':
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    if password != '':
                        password = password.encode('ascii')
                if dbg:
                    print("Logging in, username of: " + user, file=dbgout)
                self.server.login(user, password)
        elif self.authenticationtype.upper() == 'OAUTH2':
            authstring = xAuthenticator.escapeString(server)
            self.server.docmd('AUTH', 'XOAUTH2 ' + authstring)
            if dbg:
                print("OAuth2 authentication, authstring of: " + authstring, file=dbgout)
        for listener in self.listeners:
            listener.connected(self.notify)
    def disconnect(self):
        if dbg:
            print("PyMailSMTPService disconnect", file=dbgout)
        if self.server:
            self.server.quit()
            self.server = None
        for listener in self.listeners:
            listener.disconnected(self.notify)
    def isConnected(self):
        if dbg:
            print("PyMailSMTPService isConnected", file=dbgout)
        return self.server != None
    def getCurrentConnectionContext(self):
        if dbg:
            print("PyMailSMTPService getCurrentConnectionContext", file=dbgout)
        return self.connectioncontext
    def sendMailMessage(self, xMailMessage):
        COMMASPACE = ', '

        if dbg:
            print("PyMailSMTPService sendMailMessage", file=dbgout)
        recipients = xMailMessage.getRecipients()
        sendermail = xMailMessage.SenderAddress
        sendername = xMailMessage.SenderName
        subject = xMailMessage.Subject
        ccrecipients = xMailMessage.getCcRecipients()
        bccrecipients = xMailMessage.getBccRecipients()
        if dbg:
            print("PyMailSMTPService subject: " + subject, file=dbgout)
            print("PyMailSMTPService from:  " + sendername, file=dbgout)
            print("PyMailSMTPService from:  " + sendermail, file=dbgout)
            print("PyMailSMTPService send to: %s" % (recipients,), file=dbgout)

        attachments = xMailMessage.getAttachments()

        textmsg = Message()

        content = xMailMessage.Body
        flavors = content.getTransferDataFlavors()
        if dbg:
            print("PyMailSMTPService flavors len: %d" % (len(flavors),), file=dbgout)

        #Use first flavor that's sane for an email body
        for flavor in flavors:
            if flavor.MimeType.find('text/html') != -1 or flavor.MimeType.find('text/plain') != -1:
                if dbg:
                    print("PyMailSMTPService mimetype is: " + flavor.MimeType, file=dbgout)
                textbody = content.getTransferData(flavor)

                if len(textbody):
                    mimeEncoding = re.sub("charset=.*", "charset=UTF-8", flavor.MimeType)
                    if mimeEncoding.find('charset=UTF-8') == -1:
                        mimeEncoding = mimeEncoding + "; charset=UTF-8"
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

        hdr = Header(sendername, 'utf-8')
        hdr.append('<'+sendermail+'>','us-ascii')
        msg['Subject'] = subject
        msg['From'] = hdr
        msg['To'] = COMMASPACE.join(recipients)
        if len(ccrecipients):
            msg['Cc'] = COMMASPACE.join(ccrecipients)
        if xMailMessage.ReplyToAddress != '':
            msg['Reply-To'] = xMailMessage.ReplyToAddress

        mailerstring = "LibreOffice via Caolan's mailmerge component"
        try:
            configuration = self._getConfiguration("/org.openoffice.Setup/Product")
            mailerstring = "%s %s via Caolan's mailmerge component" % (configuration.getByName("ooName"), configuration.getByName("ooSetupVersion"))
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
            if dbg:
                print(("PyMailSMTPService attachmentheader: ", str(msgattachment)), file=dbgout)

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

        if dbg:
            print(("PyMailSMTPService recipients are: ", truerecipients), file=dbgout)

        self.server.sendmail(sendermail, truerecipients, msg.as_string())


class PyMailIMAPService(unohelper.Base,
                        XMailService):
    def __init__( self, ctx):
        self.ctx = ctx
        self.listeners = []
        self.connectiontype = None
        self.authenticationtype = None
        self.supportedconnection = ('Insecure', 'Ssl', 'Tls')
        self.supportedauthentication = ('None', 'Login', 'OAuth2')
        self.server = None
        self.connectioncontext = None
        self.notify = EventObject(self)
        if dbg:
            print("PyMailIMAPService init", file=dbgout)
    def addConnectionListener(self, xListener):
        if dbg:
            print("PyMailIMAPService addConnectionListener", file=dbgout)
        self.listeners.append(xListener)
    def removeConnectionListener(self, xListener):
        if dbg:
            print("PyMailIMAPService removeConnectionListener", file=dbgout)
        self.listeners.remove(xListener)
    def getConnectionType(self):
        if dbg:
            print("PyMailIMAPService getConnectionType", file=dbgout)
        return self.connectiontype
    def getAuthenticationType(self):
        if dbg:
            print("PyMailIMAPService getAuthenticationType", file=dbgout)
        return self.authenticationtype
    def getSupportedConnectionTypes(self):
        if dbg:
            print("PyMailIMAPService getSupportedConnectionTypes", file=dbgout)
        return self.supportedauthentication
    #    return self.supportedconnection
    def getSupportedAuthenticationTypes(self):
        if dbg:
            print("PyMailIMAPService getSupportedAuthenticationTypes", file=dbgout)
        return self.supportedauthentication
    def connect(self, xConnectionContext, xAuthenticator):
        if dbg:
            print("PyMailIMAPService connect", file=dbgout)

        xConnectionContext = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext", self.ctx)
        xConnectionContext.setPropertyValue("MailServiceType", IMAP)
        xAuthenticator = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.Authenticator", self.ctx)

        self.connectioncontext = xConnectionContext
        server = xConnectionContext.getValueByName("ServerName")
        if dbg:
            print("ServerName: " + server, file=dbgout)
        port = int(xConnectionContext.getValueByName("Port"))
        if dbg:
            print("Port: " + str(port), file=dbgout)
        self.connectiontype = xConnectionContext.getValueByName("ConnectionType")
        if dbg:
            print("ConnectionType: " + self.connectiontype, file=dbgout)
        self.authenticationtype = xConnectionContext.getValueByName("AuthenticationType")
        if dbg:
            print("AuthenticationType: " + self.authenticationtype, file=dbgout)
        if self.connectiontype.upper() == 'SSL':
            self.server = imaplib.IMAP4_SSL(host=server, port=port)
        else:
            self.server = imaplib.IMAP4(host=server, port=port)
        if self.connectiontype.upper() == 'TLS':
            self.server.starttls()
        if self.authenticationtype.upper() == 'LOGIN':
            user = xAuthenticator.getUserName()
            password = xAuthenticator.getPassword()
            if user != '':
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    if password != '':
                        password = password.encode('ascii')
                if dbg:
                    print("Logging in, username of: " + user, file=dbgout)
                self.server.login(user, password)
        elif self.authenticationtype.upper() == 'OAUTH2':
            authstring = xAuthenticator.unescapeString(server)
            self.server.authenticate('XOAUTH2', lambda x: authstring)
            if dbg:
                print("OAuth2 authentication, authstring of: " + authstring, file=dbgout)
        for listener in self.listeners:
            listener.connected(self.notify)
    def disconnect(self):
        if dbg:
            print("PyMailIMAPService disconnect", file=dbgout)
        if self.server:
            self.server.logout()
            self.server = None
        for listener in self.listeners:
            listener.disconnected(self.notify)
    def isConnected(self):
        if dbg:
            print("PyMailIMAPService isConnected", file=dbgout)
        return self.server != None
    def getCurrentConnectionContext(self):
        if dbg:
            print("PyMailIMAPService getCurrentConnectionContext", file=dbgout)
        return self.connectioncontext


class PyMailPOP3Service(unohelper.Base,
                        XMailService):
    def __init__( self, ctx):
        self.ctx = ctx
        self.listeners = []
        self.connectiontype = None
        self.authenticationtype = None
        self.supportedconnection = ('Insecure', 'Ssl', 'Tls')
        self.supportedauthentication = ('None', 'Login')
        self.server = None
        self.connectioncontext = None
        self.notify = EventObject(self)
        if dbg:
            print("PyMailPOP3Service init", file=dbgout)
    def addConnectionListener(self, xListener):
        if dbg:
            print("PyMailPOP3Service addConnectionListener", file=dbgout)
        self.listeners.append(xListener)
    def removeConnectionListener(self, xListener):
        if dbg:
            print("PyMailPOP3Service removeConnectionListener", file=dbgout)
        self.listeners.remove(xListener)
    def getConnectionType(self):
        if dbg:
            print("PyMailPOP3Service getConnectionType", file=dbgout)
        return self.connectiontype
    def getAuthenticationType(self):
        if dbg:
            print("PyMailPOP3Service getAuthenticationType", file=dbgout)
        return self.authenticationtype
    def getSupportedConnectionTypes(self):
        if dbg:
            print("PyMailPOP3Service getSupportedConnectionTypes", file=dbgout)
        return self.supportedauthentication
    #    return self.supportedconnection
    def getSupportedAuthenticationTypes(self):
        if dbg:
            print("PyMailPOP3Service getSupportedAuthenticationTypes", file=dbgout)
        return self.supportedauthentication
    def connect(self, xConnectionContext, xAuthenticator):
        if dbg:
            print("PyMailPOP3Service connect", file=dbgout)

        xConnectionContext = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.ConnectionContext", self.ctx)
        xConnectionContext.setPropertyValue("MailServiceType", POP3)
        xAuthenticator = self.ctx.ServiceManager.createInstanceWithContext("com.gmail.prrvchr.extensions.gMailOOo.Authenticator", self.ctx)

        self.connectioncontext = xConnectionContext
        server = xConnectionContext.getValueByName("ServerName")
        if dbg:
            print("ServerName: " + server, file=dbgout)
        port = int(xConnectionContext.getValueByName("Port"))
        if dbg:
            print("Port: " + str(port), file=dbgout)
        timeout = int(xConnectionContext.getValueByName("ConnectionTimeout"))
        if dbg:
            print("Timeout: " + str(timeout), file=dbgout)
        self.connectiontype = xConnectionContext.getValueByName("ConnectionType")
        if dbg:
            print("ConnectionType: " + self.connectiontype, file=dbgout)
        self.authenticationtype = xConnectionContext.getValueByName("AuthenticationType")
        if dbg:
            print("AuthenticationType: " + self.authenticationtype, file=dbgout)
        if self.connectiontype.upper() == 'SSL':
            self.server = poplib.POP3_SSL(host=server, port=port, timeout=timeout)
        else:
            self.server = poplib.POP3(host=server, port=port, timeout=timeout)
        if self.connectiontype.upper() == 'TLS':
            self.server.stls()
        if self.authenticationtype.upper() == 'LOGIN':
            user = xAuthenticator.getUserName()
            password = xAuthenticator.getPassword()
            if user != '':
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    if password != '':
                        password = password.encode('ascii')
                if dbg:
                    print("Logging in, username of: " + user, file=dbgout)
                self.server.user(user)
                self.server.pass_(password)
        for listener in self.listeners:
            listener.connected(self.notify)
    def disconnect(self):
        if dbg:
            print("PyMailPOP3Service disconnect", file=dbgout)
        if self.server:
            self.server.quit()
            self.server = None
        for listener in self.listeners:
            listener.disconnected(self.notify)
    def isConnected(self):
        if dbg:
            print("PyMailPOP3Service isConnected", file=dbgout)
        return self.server != None
    def getCurrentConnectionContext(self):
        if dbg:
            print("PyMailPOP3Service getCurrentConnectionContext", file=dbgout)
        return self.connectioncontext


class PyMailServiceProvider(unohelper.Base,
                            XMailServiceProvider,
                            XServiceInfo):
    def __init__( self, ctx ):
        if dbg:
            print("PyMailServiceProvider init", file=dbgout)
        self.ctx = ctx
    def create(self, aType):
        if dbg:
            print("PyMailServiceProvider create with", aType, file=dbgout)
        if aType == SMTP:
            return PyMailSMTPService(self.ctx)
        elif aType == POP3:
            return PyMailPOP3Service(self.ctx)
        elif aType == IMAP:
            return PyMailIMAPService(self.ctx)
        else:
            print("PyMailServiceProvider, unknown TYPE " + aType, file=dbgout)

    # XServiceInfo
    def supportsService(self, ServiceName):
        return g_ImplementationHelper.supportsService(g_providerImplName, ServiceName)
    def getImplementationName(self):
        return g_providerImplName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_providerImplName)


class PyMailMessage(unohelper.Base,
                    XMailMessage):
    def __init__( self, ctx, sTo='', sFrom='', Subject='', Body=None, aMailAttachment=None ):
        if dbg:
            print("PyMailMessage init", file=dbgout)
        self.ctx = ctx

        self.recipients = [sTo]
        self.ccrecipients = []
        self.bccrecipients = []
        self.aMailAttachments = []
        if aMailAttachment != None:
            self.aMailAttachments.append(aMailAttachment)

        self.SenderName, self.SenderAddress = parseaddr(sFrom)
        self.ReplyToAddress = sFrom
        self.Subject = Subject
        self.Body = Body
        if dbg:
            print("post PyMailMessage init", file=dbgout)
    def addRecipient( self, recipient ):
        if dbg:
            print("PyMailMessage.addRecipient: " + recipient, file=dbgout)
        self.recipients.append(recipient)
    def addCcRecipient( self, ccrecipient ):
        if dbg:
            print("PyMailMessage.addCcRecipient: " + ccrecipient, file=dbgout)
        self.ccrecipients.append(ccrecipient)
    def addBccRecipient( self, bccrecipient ):
        if dbg:
            print("PyMailMessage.addBccRecipient: " + bccrecipient, file=dbgout)
        self.bccrecipients.append(bccrecipient)
    def getRecipients( self ):
        if dbg:
            print("PyMailMessage.getRecipients: " + str(self.recipients), file=dbgout)
        return tuple(self.recipients)
    def getCcRecipients( self ):
        if dbg:
            print("PyMailMessage.getCcRecipients: " + str(self.ccrecipients), file=dbgout)
        return tuple(self.ccrecipients)
    def getBccRecipients( self ):
        if dbg:
            print("PyMailMessage.getBccRecipients: " + str(self.bccrecipients), file=dbgout)
        return tuple(self.bccrecipients)
    def addAttachment( self, aMailAttachment ):
        if dbg:
            print("PyMailMessage.addAttachment", file=dbgout)
        self.aMailAttachments.append(aMailAttachment)
    def getAttachments( self ):
        if dbg:
            print("PyMailMessage.getAttachments", file=dbgout)
        return tuple(self.aMailAttachments)

    # XServiceInfo
    def supportsService(self, ServiceName):
        return g_ImplementationHelper.supportsService(g_messageImplName, ServiceName)
    def getImplementationName(self):
        return g_messageImplName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_messageImplName)


g_ImplementationHelper.addImplementation(PyMailServiceProvider,
                                         g_providerImplName,
                                         ("com.sun.star.mail.MailServiceProvider", ), )

g_ImplementationHelper.addImplementation(PyMailMessage,
                                         g_messageImplName,
                                         ("com.sun.star.mail.MailMessage", ), )
