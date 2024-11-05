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
from emailer import getStreamSequence
from emailer import hasInterface

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


class MailMessage(unohelper.Base,
                  XMailMessage2,
                  XServiceInfo):
    def __init__(self, ctx, to='', sender='', subject='', body=None, attachment=None):
        self._ctx = ctx
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
        host = self._saddress.replace('@', '.')
        self.MessageId = make_msgid(None, host)
        self.ThreadId = ''
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
                data = self._getData(self.Body, flavor)
                if len(data):
                    body = self._getBody(data, encoding, mimetype)
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
        message['Message-ID'] = self.MessageId
        message['Reply-To'] = self.ReplyToAddress
        if self.ThreadId:
            message['References'] = self.ThreadId
        if self.hasCcRecipients():
            message['Cc'] = COMMASPACE.join(self.getCcRecipients())
        try:
            configuration = getConfiguration(self._ctx, '/org.openoffice.Setup/Product')
            name = configuration.getByName('ooName')
            version1 = configuration.getByName('ooSetupVersion')
            version2 = getExtensionVersion(self._ctx, g_identifier)
            xmailer = "%s %s via eMailerOOo v%s extension" % (name, version1, version2)
        except:
            xmailer = "LibreOffice via eMailerOOo extension"
        message['X-Mailer'] = xmailer
        message['Date'] = formatdate(localtime=True)
        for attachment in self.getAttachments():
            name = attachment.ReadableName
            # FIXME: We need to check if the attachments have a ReadableName,
            # FIXME: if not we raise a MailMessageException
            if not name:
                msg = self._getLogger().resolveString(2002, self.Subject)
                raise MailMessageException(msg, self, ())
            content = attachment.Data
            flavors = content.getTransferDataFlavors()
            # FIXME: We need to check if the attachments have at least
            # FIXME: one DataFlavor, if not we raise a UnsupportedFlavorException
            if not isinstance(flavors, tuple) or not len(flavors):
                msg = self._getLogger().resolveString(2003, name, self.Subject)
                raise UnsupportedFlavorException(msg, self)
            flavor = flavors[0]
            msgattachment, encoding = self._getAttachment(content, flavor, name)
            encode_base64(msgattachment)
            msgattachment.add_header('Content-Disposition',
                                     'attachment',
                                      filename=(encoding, self._getLanguage(), name))
            message.attach(msgattachment)
        return message

    def _getData(self, transferable, flavor):
        interface = 'com.sun.star.io.XInputStream'
        data = transferable.getTransferData(flavor)
        if flavor.DataType == uno.getTypeByName(interface):
            if not hasInterface(data, interface):
                msg = self._getLogger().resolveString(2011, self.Subject, interface, flavor.DataType.typeName)
                raise UnsupportedFlavorException(msg, self)
            data = getStreamSequence(data)
        return data

    def _getBody(self, data, encoding, mimetype):
        # If it's a byte sequence, decode it
        if isinstance(data, uno.ByteSequence):
            data = data.value.decode(encoding)
        # If it's a string, nothing to do
        elif isinstance(data, str):
            pass
        # No data is available, we need to raise an Exception
        else:
            msg = self._getLogger().resolveString(2021, self.Subject, repr(type(data)))
            raise UnsupportedFlavorException(msg, self)
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

    def _getAttachment(self, content, flavor, name):
        data = self._getData(content, flavor)
        if not len(data):
            msg = self._getLogger().resolveString(2031, name, self.Subject)
            raise UnsupportedFlavorException(msg, self)
        mct = self._mtf.createMimeContentType(flavor.MimeType)
        if mct.hasParameter(self._charset):
            encoding = mct.getParameterValue(self._charset)
        else:
            # Default to utf-8
            encoding = self._encoding
        # Normally it's a bytesequence, get raw bytes
        if isinstance(data, uno.ByteSequence):
            data = data.value
        # If it's a string, get it as 'encoding' bytes
        elif isinstance(data, str):
            data = data.encode(encoding)
        # No data is available, we need to raise an Exception
        else:
            msg = self._getLogger().resolveString(2032, name, self.Subject, repr(type(data)))
            raise UnsupportedFlavorException(msg, self)
        msgattachment = MIMEBase(mct.getMediaType(), mct.getMediaSubtype())
        msgattachment.set_payload(data)
        return msgattachment, encoding

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


g_ImplementationHelper.addImplementation(MailMessage,
                                         g_ImplementationName,
                                        ('com.sun.star.mail.MailMessage', ))

