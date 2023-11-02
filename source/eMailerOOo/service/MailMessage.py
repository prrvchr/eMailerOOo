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

from com.sun.star.datatransfer import UnsupportedFlavorException

from com.sun.star.lang import XServiceInfo

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import XMailMessage2

from emailer import createService
from emailer import getConfiguration
from emailer import getLogger
from emailer import getMimeTypeFactory
from emailer import getStreamSequence
from emailer import hasInterface

from emailer import g_mailservicelog
from emailer import g_version

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

import six
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'org.openoffice.pyuno.MailMessage2'


class MailMessage(unohelper.Base,
                  XMailMessage2,
                  XServiceInfo):
    def __init__(self, ctx, to='', sender='', subject='', body=None, attachment=None):
        print("MailMessage.__init__() 1")
        self._ctx = ctx
        self._logger = getLogger(ctx, g_mailservicelog)
        self._mtf = getMimeTypeFactory(ctx)
        self._recipients = [to]
        self._ccrecipients = []
        self._bccrecipients = []
        self._sname, self._saddress = parseaddr(sender)
        host = self._saddress.replace('@', '.')
        self._messageid = make_msgid(None, host)
        self.ReplyToAddress = sender
        self.Subject = subject
        self.Body = body
        self._attachments = []
        self._charset = 'charset'
        self._encode = 'UTF-8'
        if attachment is not None:
            self._attachments.append(attachment)
        self.ThreadId = ''
        print("MailMessage.__init__() 2 %s - %s - %s" % (sender, self._sname, self._saddress))

    @property
    def SenderName(self):
        return self._sname

    @property
    def SenderAddress(self):
        return self._saddress

    @property
    def MessageId(self):
        return self._messageid
    @MessageId.setter
    def MessageId(self, messageid):
        self._messageid = messageid

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
        parsed = False
        COMMASPACE = ', '
        body = Message()
        encode = self._encode
        #Use first flavor that's sane for an email body
        for flavor in self.Body.getTransferDataFlavors():
            mimetype = flavor.MimeType
            if mimetype.startswith('text/html') or mimetype.startswith('text/plain'):
                mct = self._mtf.createMimeContentType(mimetype)
                if mct.hasParameter(self._charset):
                    encode = mct.getParameterValue(self._charset)
                else:
                    mimetype += '; %s=%s' % (self._charset, encode)
                data = self._getTransferData(self.Body, flavor)
                if len(data):
                    print("MailMessage._getMessage() 1 mimetype: %s - charset: %s" % (mimetype, encode))
                    body['Content-Type'] = mimetype
                    body['MIME-Version'] = '1.0'
                    data = self._getBodyData(data, encode).decode(encode)
                    c = Charset(encode)
                    c.body_encoding = QP
                    body.set_payload(data, c)
                    parsed = True
                    break
        # FIXME: We need to check if the body has been parsed,
        # FIXME: if not we raise a UnsupportedFlavorException
        if not parsed:
            msg = self._logger.resolveString(2001, self.Subject)
            raise UnsupportedFlavorException(msg, self)
        if self.hasAttachments():
            message = MIMEMultipart()
            message.epilogue = ''
            message.attach(body)
        else:
            message = body
        message['Subject'] = self.Subject
        header = Header(self.SenderName, 'utf-8')
        header.append('<' + self.SenderAddress + '>','us-ascii')
        message['From'] = header
        message['To'] = COMMASPACE.join(self.getRecipients())
        message['Message-ID'] = self.MessageId
        if self.ThreadId:
            message['References'] = self.ThreadId
        if self.hasCcRecipients():
            message['Cc'] = COMMASPACE.join(self.getCcRecipients())
        if self.ReplyToAddress:
            message['Reply-To'] = self.ReplyToAddress
        try:
            configuration = getConfiguration(self._ctx, '/org.openoffice.Setup/Product')
            name = configuration.getByName('ooName')
            version = configuration.getByName('ooSetupVersion')
            xmailer = "%s %s via eMailerOOo v%s extension" % (name, version, g_version)
        except:
            xmailer = "LibreOffice / OpenOffice via eMailerOOo v%s extension" % g_version
        message['X-Mailer'] = xmailer
        message['Date'] = formatdate(localtime=True)
        for attachment in self.getAttachments():
            name = attachment.ReadableName
            # FIXME: We need to check if the attachments have a ReadableName,
            # FIXME: if not we raise a MailMessageException
            if not name:
                msg = self._logger.resolveString(2002, self.Subject)
                raise MailMessageException(msg, self, ())
            content = attachment.Data
            flavors = content.getTransferDataFlavors()
            # FIXME: We need to check if the attachments have at least
            # FIXME: one DataFlavor, if not we raise a UnsupportedFlavorException
            if not isinstance(flavors, tuple) or not len(flavors):
                msg = self._logger.resolveString(2003, name, self.Subject)
                raise UnsupportedFlavorException(msg, self)
            flavor = flavors[0]
            msgattachment = self._getMessageAttachment(content, flavor, name)
            encode_base64(msgattachment)
            try:
                msgattachment.add_header('Content-Disposition',
                                         'attachment',
                                         filename=name)
            except:
                print("MailMessage._getMessage() ERROR **********************************************")
                msgattachment.add_header('Content-Disposition',
                                         'attachment',
                                         filename=('utf-8', '', name))
            message.attach(msgattachment)
        return message

    def _getTransferData(self, transferable, flavor):
        interface = 'com.sun.star.io.XInputStream'
        data = transferable.getTransferData(flavor)
        if flavor.DataType == uno.getTypeByName(interface):
            if not hasInterface(data, interface):
                msg = self._logger.resolveString(2011, self.Subject, interface, flavor.DataType.typeName)
                raise UnsupportedFlavorException(msg, self)
            data = getStreamSequence(data)
        return data

    def _getBodyData(self, data, encode):
        # Normally it's a bytesequence, get raw bytes
        if isinstance(data, uno.ByteSequence):
            data = data.value
        # If it's a string, get it as 'encode' bytes
        elif isinstance(data, str):
            data = data.encode(encode)
        # No data is available, we need to raise an Exception
        else:
            msg = self._logger.resolveString(2021, self.Subject, repr(type(data)))
            raise UnsupportedFlavorException(msg, self)
        return data

    def _getMessageAttachment(self, content, flavor, name):
        mct = self._mtf.createMimeContentType(flavor.MimeType)
        data = self._getTransferData(content, flavor)
        # Normally it's a bytesequence, get raw bytes
        if isinstance(data, uno.ByteSequence):
            data = data.value
        # If it's a string, we need to get its encoding
        elif isinstance(data, str):
            if mct.hasParameter(self._charset):
                encode = mct.getParameterValue(self._charset)
            else:
                # Default to utf-8
                encode = self._encode
            data = data.encode(encode)
        # No data is available, we need to raise an Exception
        else:
            msg = self._logger.resolveString(2031, name, self.Subject, repr(type(data)))
            raise UnsupportedFlavorException(msg, self)
        msgattachment = MIMEBase(mct.getMediaType(), mct.getMediaSubtype())
        msgattachment.set_payload(data)
        return msgattachment

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

