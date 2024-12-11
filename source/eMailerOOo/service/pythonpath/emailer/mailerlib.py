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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.datatransfer import UnsupportedFlavorException
from com.sun.star.datatransfer import XTransferable

from com.sun.star.mail import XAuthenticator

from com.sun.star.uno import XCurrentContext

from .oauth2 import setItemsIdentifier
from .oauth2 import setParametersArguments

from .unotool import getStreamSequence

from string import Template
from collections import UserDict
import json
import traceback


def getTransferable(logger, flavor, data):
    return Transferable(logger, flavor, data)


class Transferable(unohelper.Base,
                   XTransferable):
    def __init__(self, logger, flavor, data):
        self._logger = logger
        self._flavor = flavor
        self._data = data
        self._typemap = self._getTypeMap()

# XTransferable
    def getTransferData(self, flavor):
        if not self.isDataFlavorSupported(flavor):
            msg = self._logger.resolveString(3001, flavor.MimeType, flavor.DataType.typeName)
            raise UnsupportedFlavorException(msg, self)
        map = (self._flavor.DataType.typeName, flavor.DataType.typeName)
        return self._typemap[map](self._data)

    def getTransferDataFlavors(self):
        return (self._flavor,)

    def isDataFlavorSupported(self, flavor):
        if self._flavor.MimeType != flavor.MimeType:
            return False
        map = (self._flavor.DataType.typeName, flavor.DataType.typeName)
        return True if map in self._typemap else False

    def _getTypeMap(self):
        ipstream = 'com.sun.star.io.XInputStream'
        return {(ipstream, '[]byte'): lambda x: getStreamSequence(x),
                (ipstream, ipstream): lambda x: x,
                ('string', 'string'): lambda x: x,
                ('[]byte', '[]byte'): lambda x: x}


class Authenticator(unohelper.Base,
                    XAuthenticator):
    def __init__(self, user):
        self._user = user

# XAuthenticator
    def getUserName(self):
        return self._user.get('Login')

    def getPassword(self):
        return self._user.get('Password')


class CurrentContext(unohelper.Base,
                     XCurrentContext):
    def __init__(self, context):
        self._context = context

# XCurrentContext
    def getValueByName(self, name):
        return self._context.get(name)


class CustomMessage(UserDict):
    def __init__(self, logger, cls, message, request):
        UserDict.__init__(self)
        self._keys = ('MessageId', 'ThreadId', 'ForeignId',
                      'Recipients', 'CcRecipients', 'BccRecipients',
                      'Subject', 'Body', 'MimeType', 'Attachments', 'Message')
        self._datatypes = ('[]byte', 'string')
        self._message = message
        self._parameters = request.getByName('Parameters')
        self._logger = logger
        self._cls = cls

    def keys(self):
        return self._keys + tuple(self.data.keys())

    def __contains__(self, key):
        return key in self._keys or key in self.data

    def __missing__(self, key):
        if key == 'MessageId':
            value = self._message.MessageId
        elif key == 'ThreadId':
            value = self._message.ThreadId
        elif key == 'ForeignId':
            value = self._message.ForeignId
        elif key == 'Recipients':
            getter = lambda x: x.getRecipients()
            parameters, template = self._getParameters(key)
            value = self._getRecipients('Recipient', getter, parameters, template)
        elif key == 'CcRecipients':
            getter = lambda x: x.getCcRecipients()
            parameters, template = self._getParameters(key)
            value = self._getRecipients('CcRecipient', getter, parameters, template)
        elif key == 'BccRecipients':
            getter = lambda x: x.getBccRecipients()
            parameters, template = self._getParameters(key)
            value = self._getRecipients('BccRecipient', getter, parameters, template)
        elif key == 'Subject':
            value = self._message.Subject
        elif key == 'Body':
            value = self._getBody()
        elif key == 'MimeType':
            value = self._message.Body.getTransferDataFlavors()[0].MimeType
        elif key == 'Attachments':
            parameters, template = self._getParameters(key)
            value = self._getAttachments(parameters, template)
        elif key == 'Message':
            value = self._getMessage()
        else:
            # Key not found we use UserDict to raise KeyError
            value = self.data[key]
        return value

    def _getRecipients(self, identifier, getter, parameters, template):
        mtd = '_getRecipients'
        recipients = []
        self._logger.logprb(INFO, self._cls, mtd, 421, identifier)
        for recipient in getter(self._message):
            items = json.loads(template)
            arguments = {identifier: recipient}
            setParametersArguments(parameters, arguments)
            setItemsIdentifier(items, arguments)
            recipients.append(items)
        self._logger.logprb(INFO, self._cls, mtd, 422, identifier, len(recipients))
        return recipients

    def _getBody(self):
        mtd = '_getBody'
        flavor = self._message.Body.getTransferDataFlavors()[0]
        self._logger.logprb(INFO, self._cls, mtd, 411, flavor.DataType.typeName)
        for typename in self._datatypes:
            flavor.DataType = uno.getTypeByName(typename)
            if self._message.Body.isDataFlavorSupported(flavor):
                if flavor.DataType.typeName == 'string':
                    body = self._message.Body.getTransferData(flavor).encode()
                else:
                    body = self._message.Body.getTransferData(flavor).value
                self._logger.logprb(INFO, self._cls, mtd, 412, flavor.DataType.typeName)
                arguments = {'Body': body}
                setParametersArguments(self._parameters, arguments)
                return arguments['Body']
        self._logger.logprb(SEVERE, self._cls, mtd, 413, flavor.DataType.typeName)

    def _getAttachments(self, parameters, template):
        mtd = '_getAttachments'
        attachments = []
        self._logger.logprb(INFO, self._cls, mtd, 431)
        for attachment in self._message.getAttachments():
            flavor = attachment.Data.getTransferDataFlavors()[0]
            self._logger.logprb(INFO, self._cls, mtd, 432, attachment.ReadableName, flavor.DataType.typeName)
            for typename in self._datatypes:
                flavor.DataType = uno.getTypeByName(typename)
                if attachment.Data.isDataFlavorSupported(flavor):
                    items = json.loads(template)
                    if flavor.DataType.typeName == 'string':
                        data = attachement.Data.getTransferData(flavor).encode()
                    else:
                        data = attachement.Data.getTransferData(flavor).value
                    arguments = {'Data':         data,
                                 'MimeType':     flavor.MimeType,
                                 'ReadableName': attachement.ReadableName}
                    setParametersArguments(parameters, arguments)
                    setItemsIdentifier(items, arguments)
                    attachments.append(items)
                    self._logger.logprb(INFO, self._cls, mtd, 433, attachment.ReadableName, flavor.DataType.typeName, flavor.MimeType)
                    break
            else:
                self._logger.logprb(SEVERE, self._cls, mtd, 434, attachment.ReadableName, flavor.DataType.typeName)
        self._logger.logprb(INFO, self._cls, mtd, 435, len(attachments))
        return attachments

    def _getMessage(self):
        arguments = {'Message': self._message.asBytes().value}
        setParametersArguments(self._parameters, arguments)
        return arguments['Message']

    def _getParameters(self, key):
        parameters = template = None
        for name in self._parameters.getElementNames():
            parameter = self._parameters.getByName(name)
            if parameter.getByName('Name') == key:
                parameters = parameter.getByName('Parameters')
                template = parameter.getByName('Template')
                break
        return parameters, template

