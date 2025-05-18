#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
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

from com.sun.star.datatransfer import UnsupportedFlavorException

from com.sun.star.lang import XServiceInfo

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import XMailMessage2

from emailer import getConfiguration
from emailer import getCurrentLocale
from emailer import getExtensionVersion
from emailer import getLogger
from emailer import getMimeTypeFactory

from emailer import g_identifier
from emailer import g_mailservicelog

from email.mime.base import MIMEBase
from email.message import Message
from email.charset import Charset
from email.charset import QP
from email.encoders import encode_base64
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from email.utils import make_msgid
from email.utils import parseaddr
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'org.openoffice.pyuno.MailMessage2'
g_ServiceNames = ('com.sun.star.mail.MailMessage', )


class MailMessage(unohelper.Base,
                  XMailMessage2,
                  XServiceInfo):
    def __init__(self, ctx, to='', sender='', subject='', body=None, attachment=None):
        self._ctx = ctx
        self._datatypes = {'[]byte': lambda x, y: x.getTransferData(y).value,
                           'string': lambda x, y: x.getTransferData(y)}
        self._mtf = getMimeTypeFactory(ctx)
        self._recipients = [to]
        self._ccrecipients = []
        self._bccrecipients = []
        self._charset = 'charset'
        self._encoding = 'UTF-8'
        self._attachments = []
        if attachment is not None:
            self._attachments.append(attachment)
        self._language = None
        self._sname, self._saddress = parseaddr(sender)
        self._headers = {}
        host = self._saddress.replace('@', '.')
        self._headers['Message-Id'] = make_msgid(None, host)
        self._xmailer = self._getXMailer(ctx)
        self._foreignid = ''
        self._custom = ('MessageId', 'ThreadId', 'ForeignId')
        self.ReplyToAddress = sender
        self.Subject = subject
        self.Body = body

    @property
    def SenderName(self):
        return self._sname
    @property
    def SenderAddress(self):
        return self._saddress

    @property
    def MessageId(self):
        return self._headers.get('Message-Id', '')
    @MessageId.setter
    def MessageId(self, value):
        self._headers['Message-Id'] = value

    @property
    def ThreadId(self):
        return self._headers.get('References', '')
    @ThreadId.setter
    def ThreadId(self, value):
        self._headers['References'] = value

    @property
    def ForeignId(self):
        return self._foreignid
    @ForeignId.setter
    def ForeignId(self, value):
        self._foreignid = value

    def hasHeaders(self, name):
        return name in self._headers

    def getHeaderNames(self):
        return tuple(self._headers.keys())

    def getHeader(self, name):
        if name not in self._custom:
            return self._headers.get(name, '')
        if name == 'ForeignId':
            return self._foreignid
        if name == 'MessageId':
            return self.MessageId
        if name == 'ThreadId':
            return self.ThreadId

    def setHeader(self, name, value):
        if name not in self._custom:
            self._headers[name] = value
        elif name == 'ForeignId':
            self._foreignid = value
        elif name == 'MessageId':
            self.MessageId = value
        elif name == 'ThreadId':
            self.ThreadId = value

    def addRecipient(self, recipient):
        self._recipients.append(recipient)

    def getRecipients(self):
        return tuple(self._recipients)

    def hasRecipients(self):
        return len(self._recipients) > 0

    def addCcRecipient(self, ccrecipient):
        self._ccrecipients.append(ccrecipient)

    def getCcRecipients(self):
        return tuple(self._ccrecipients)

    def hasCcRecipients(self):
        return len(self._ccrecipients) > 0

    def addBccRecipient(self, bccrecipient):
        self._bccrecipients.append(bccrecipient)

    def getBccRecipients(self):
        return tuple(self._bccrecipients)

    def hasBccRecipients(self):
        return len(self._bccrecipients) > 0

    def addAttachment(self, attachment):
        self._attachments.append(attachment)

    def getAttachments(self):
        return tuple(self._attachments)

    def hasAttachments(self):
        return len(self._attachments) > 0

    def asBytes(self):
        message = self._getMessage()
        return uno.ByteSequence(message.as_bytes())

    def asString(self):
        message = self._getMessage()
        return message.as_string()

    def _getMessage(self):
        body = None
        COMMASPACE = ', '
        encoding = self._encoding
        #Use first flavor that's sane for an email body
        for flavor in self.Body.getTransferDataFlavors():
            mimetype = flavor.MimeType
            if mimetype.startswith('text/html') or mimetype.startswith('text/plain'):
                mct = self._mtf.createMimeContentType(mimetype)
                if mct.hasParameter(self._charset):
                    encoding = mct.getParameterValue(self._charset)
                else:
                    mimetype += '; %s=%s' % (self._charset, encoding)
                for typename in self._datatypes:
                    flavor.DataType = uno.getTypeByName(typename)
                    if self.Body.isDataFlavorSupported(flavor):
                        data = self._datatypes[typename](self.Body, flavor)
                        body = self._getBody(data, encoding, mimetype)
                        break
                else:
                    continue
                break
        # FIXME: We need to check if the body has been parsed,
        # FIXME: if not we raise a UnsupportedFlavorException
        if body is None:
            msg = self._getLogger().resolveString(2001, self.Subject)
            raise UnsupportedFlavorException(msg, self)
        if self.hasAttachments():
            message = self._getMultipart(body)
        else:
            message = body
        message['Subject'] = self.Subject
        message['From'] = self._getHeader()
        message['To'] = COMMASPACE.join(self.getRecipients())
        message['Reply-To'] = self.ReplyToAddress
        if self.hasCcRecipients():
            message['Cc'] = COMMASPACE.join(self.getCcRecipients())
        for header, value in self._headers.items():
            message[header] = value
        message['X-Mailer'] = self._xmailer
        message['Date'] = formatdate(localtime=True)
        for attachment in self.getAttachments():
            name = attachment.ReadableName
            # FIXME: We need to check if the attachments have a ReadableName,
            # FIXME: if not we raise a MailMessageException
            if not name:
                msg = self._getLogger().resolveString(2002, self.Subject)
                raise MailMessageException(msg, self, ())
            transferable = attachment.Data
            flavors = transferable.getTransferDataFlavors()
            # FIXME: We need to check if the attachments have at least
            # FIXME: one DataFlavor, if not we raise a UnsupportedFlavorException
            if not isinstance(flavors, tuple) or not len(flavors):
                msg = self._getLogger().resolveString(2003, name, self.Subject)
                raise UnsupportedFlavorException(msg, self)
            flavor = flavors[0]
            msgattachment, encoding = self._getAttachment(transferable, flavor, name)
            encode_base64(msgattachment)
            msgattachment.add_header('Content-Disposition',
                                     'attachment',
                                      filename=(encoding, self._getLanguage(), name))
            message.attach(msgattachment)
        return message

    def _getBody(self, data, encoding, mimetype):
        body = Message()
        body['Content-Type'] = mimetype
        body['MIME-Version'] = '1.0'
        body.set_payload(data, self._getCharset(encoding))
        return body

    def _getCharset(self, encoding):
        charset = Charset(encoding)
        charset.body_encoding = QP
        return charset

    def _getMultipart(self, body):
        message = MIMEMultipart()
        message.epilogue = ''
        message.attach(body)
        return message

    def _getHeader(self):
        header = Header(self.SenderName, 'utf-8')
        header.append('<' + self.SenderAddress + '>','us-ascii')
        return header

    def _getAttachment(self, transferable, flavor, name):
        for typename in self._datatypes:
            flavor.DataType = uno.getTypeByName(typename)
            if transferable.isDataFlavorSupported(flavor):
                data = self._datatypes[typename](transferable, flavor)
                mct = self._mtf.createMimeContentType(flavor.MimeType)
                if mct.hasParameter(self._charset):
                    encoding = mct.getParameterValue(self._charset)
                else:
                    # Default to utf-8
                    encoding = self._encoding
                msgattachment = MIMEBase(mct.getMediaType(), mct.getMediaSubtype())
                msgattachment.set_payload(data)
                return msgattachment, encoding
        msg = self._getLogger().resolveString(2031, name, self.Subject)
        raise UnsupportedFlavorException(msg, self)

    def _getXMailer(self, ctx):
        try:
            config = getConfiguration(ctx, '/org.openoffice.Setup/Product')
            name = config.getByName('ooName')
            version1 = config.getByName('ooSetupVersion')
            version2 = getExtensionVersion(ctx, g_identifier)
            xmailer = "%s %s via eMailerOOo v%s extension" % (name, version1, version2)
        except:
            xmailer = "LibreOffice via eMailerOOo extension"
        return xmailer

    def _getLanguage(self):
        if self._language is None:
            self._language = getCurrentLocale(self._ctx).Language
        return self._language

    def _getLogger(self):
        return getLogger(self._ctx, g_mailservicelog)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

g_ImplementationHelper.addImplementation(MailMessage,                     # UNO object class
                                         g_ImplementationName,            # Implementation name
                                         g_ServiceNames)                  # List of implemented services
