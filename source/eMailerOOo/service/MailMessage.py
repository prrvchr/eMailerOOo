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

from com.sun.star.lang import XServiceInfo
from com.sun.star.mail import XMailMessage2

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

from emailer import getConfiguration
from emailer import getLogger
from emailer import g_mailservicelog
from emailer import g_version

import re
import sys
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'org.openoffice.pyuno.MailMessage2'


class MailMessage(unohelper.Base,
                  XMailMessage2,
                  XServiceInfo):
    def __init__(self, ctx, to='', frm='', subject='', body=None, attachment=None):
        print("MailMessage.__init__() 1")
        self._ctx = ctx
        self._logger = getLogger(ctx, g_mailservicelog)
        self._recipients = [to]
        self._ccrecipients = []
        self._bccrecipients = []
        self._sname, self._saddress = parseaddr(frm)
        name, part, domain = frm.partition('@')
        host = '%s.%s' % (name, domain)
        self._messageid = make_msgid(None, host)
        self.ReplyToAddress = frm
        self.Subject = subject
        self.Body = body
        self._attachments = []
        if attachment is not None:
            self._attachments.append(attachment)
        self.ThreadId = ''
        print("MailMessage.__init__() 2 %s - %s - %s" % (frm, self._sname, self._saddress))

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
        msg = self._getMessage()
        return uno.ByteSequence(msg.as_bytes())

    def asString(self):
        msg = self._getMessage()
        return msg.as_string()

    def _getMessage(self):
        COMMASPACE = ', '
        self._logger.logprb(INFO, 'SmtpService', 'sendMailMessage()', 251, self.Subject)
        textmsg = Message()
        #Use first flavor that's sane for an email body
        for flavor in self.Body.getTransferDataFlavors():
            if flavor.MimeType.find('text/html') != -1 or flavor.MimeType.find('text/plain') != -1:
                textbody = self.Body.getTransferData(flavor)
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
        if self.hasAttachments():
            msg = MIMEMultipart()
            msg.epilogue = ''
            msg.attach(textmsg)
        else:
            msg = textmsg
        msg['Subject'] = self.Subject
        header = Header(self.SenderName, 'utf-8')
        header.append('<' + self.SenderAddress + '>','us-ascii')
        msg['From'] = header
        msg['To'] = COMMASPACE.join(self.getRecipients())
        msg['Message-ID'] = self.MessageId
        if self.ThreadId:
            msg['References'] = self.ThreadId
        if self.hasCcRecipients():
            msg['Cc'] = COMMASPACE.join(self.getCcRecipients())
        if self.ReplyToAddress:
            msg['Reply-To'] = self.ReplyToAddress
        xmailer = "LibreOffice / OpenOffice via eMailerOOo v%s extension" % g_version
        try:
            configuration = getConfiguration(self._ctx, '/org.openoffice.Setup/Product')
            name = configuration.getByName('ooName')
            version = configuration.getByName('ooSetupVersion')
            xmailer = "%s %s via eMailerOOo v%s extension" % (name, version, g_version)
        except:
            pass
        msg['X-Mailer'] = xmailer
        msg['Date'] = formatdate(localtime=True)
        for attachment in self.getAttachments():
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
                msgattachment.add_header('Content-Disposition',
                                         'attachment',
                                         filename=fname)
            except:
                msgattachment.add_header('Content-Disposition',
                                         'attachment',
                                         filename=('utf-8', '', fname))
            msg.attach(msgattachment)
        return msg


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
