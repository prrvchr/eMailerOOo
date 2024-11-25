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

from .mailertool import setParametersArguments

from .unotool import getStreamSequence

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
    def __init__(self, logger, cls, message):
        self._keys = ('MessageId', 'ThreadId', 'ForeignId', 'Subject',
                      'Recipients', 'CcRecipients', 'BccRecipients',
                      'Body', 'Attachments', 'Message')
        self._body = lambda x, y: x.getTransferData(y).value \
                                  if y.DataType.typeName != 'string' else \
                                  x.getTransferData(y).encode()
        self._recipients =  {'Recipients':    lambda x: x.getRecipients(),
                             'CcRecipients':  lambda x: x.getCcRecipients(),
                             'BccRecipients': lambda x: x.getBccRecipients()}
        getdata =    lambda x, y: x.Data.getTransferData(y).value \
                                  if y.DataType.typeName != 'string' else \
                                  x.Data.getTransferData(y).encode()
        self._attachments = {'Data':          getdata,
                             'DataFlavor':    lambda x, y: y.DataFlavor,
                             'ReadableName':  lambda x, y: x.ReadableName}
        self._datatypes = ('[]byte', 'string')
        self._message = message
        self._prefix = '${'
        self._suffix = '}'
        self._templates = {}
        self._logger = logger
        self._cls = cls
        self.data = {}
        UserDict.__init__(self)

    def __getitem__(self, key):
        if key in self.data or key not in self._keys:
            return self.data[key]
        return self._getItem(key)

    def setTemplate(self, name, template, parameters):
        self._templates[name] = (template, parameters)

    def toJson(self, template):
        items = json.loads(template)
        identifiers = {self._getKey(key): key for key in self._keys}
        self._setItems(items, identifiers, self.__getitem__)
        return json.dumps(items)

    def _getItem(self, key):
        if key == 'Body':
            value = self._getBody()
        elif key == 'Subject':
            value = self._message.Subject
        elif key == 'MessageId':
            value = self._message.MessageId
        elif key == 'ThreadId':
            value = self._message.ThreadId
        elif key == 'ForeignId':
            value = self._message.ForeignId
        elif key == 'Message':
            value = self._message.asBytes().value
        elif key == 'Recipients':
            template, dummy = self._templates[key]
            value = self._getRecipients('Recipient', template, self._recipients[key])
        elif key == 'CcRecipients':
            template, dummy = self._templates[key]
            value = self._getRecipients('CcRecipient', template, self._recipients[key])
        elif key == 'BccRecipients':
            template, dummy = self._templates[key]
            value = self._getRecipients('BccRecipient', template, self._recipients[key])
        elif key == 'Attachments':
            value = self._getAttachments(key)
        return value

    def _getBody(self):
        mtd = '_getBody'
        flavor = self._message.Body.getTransferDataFlavors()[0]
        self._logger.logprb(INFO, self._cls, mtd, 411, flavor.DataType.typeName)
        for typename in self._datatypes:
            flavor.DataType = uno.getTypeByName(typename)
            if self._message.Body.isDataFlavorSupported(flavor):
                body = self._body(self._message.Body, flavor)
                self._logger.logprb(INFO, self._cls, mtd, 412, flavor.DataType.typeName)
                return body
        self._logger.logprb(SEVERE, self._cls, mtd, 413, flavor.DataType.typeName)

    def _getRecipients(self, identifier, template, getRecipients):
        mtd = '_getRecipients'
        recipients = []
        self._logger.logprb(INFO, self._cls, mtd, 421)
        identifiers = {self._getKey(identifier): identifier}
        for recipient in getRecipients(self._message):
            items = json.loads(template)
            self._setItems(items, identifiers, lambda x: recipient)
            recipients.append(items)
        self._logger.logprb(INFO, self._cls, mtd, 422, len(recipients))
        return recipients

    def _getAttachments(self, identifier):
        mtd = '_getAttachments'
        attachments = []
        self._logger.logprb(INFO, self._cls, mtd, 431)
        identifiers = {self._getKey(key): key for key in self._attachments}
        template, parameters = self._templates.get(identifier)
        for attachment in self._message.getAttachments():
            flavor = attachment.Data.getTransferDataFlavors()[0]
            for typename in self._datatypes:
                flavor.DataType = uno.getTypeByName('[]byte')
                if self._message.Body.isDataFlavorSupported(flavor):
                    arguments = {key: data(attachment, flavor) for key, data in self._attachments.items()}
                    setParametersArguments(parameters, arguments)
                    items = json.loads(template) 
                    self._setItems(items, identifiers, lambda x: arguments[x])
                    attachments.append(items)
                    break
            else:
                self._logger.logprb(SEVERE, self._cls, mtd, 432, attachment.ReadableName, flavor.DataType.typeName)
        self._logger.logprb(INFO, self._cls, mtd, 433, len(attachments))
        return attachments

    def _setItems(self, items, identifiers, getter):
        for key, value in items.items():
            if isinstance(value, dict):
                self._setItems(value, identifiers, getter)
            elif isinstance(value, list):
                self._setItemsList(value, identifiers, getter)
            elif value in identifiers:
                items[key] = getter(identifiers[value])

    def _setItemsList(self, values, identifiers, getter):
        for i, value in enumerate(values):
            if isinstance(value, dict):
                self._setItems(value, identifiers, getter)
            elif isinstance(value, list):
                self._setItemsList(value, identifiers, getter)
            elif value in identifiers:
                values[i] = getter(identifiers[value])

    def _getKey(self, identifier):
        return  self._prefix + identifier + self._suffix


class CustomParser():
    def __init__(self, keys, items, triggers):
        self._keys = keys
        self._items = items
        self._triggers = triggers
        self._key = None

    def hasItems(self):
        return any((self._items, self._triggers))

    def parse(self, results, prefix, event, value):
        if (prefix, event) in self._items:
            key = self._items[(prefix, event)]
            results[key] = value
            if self._key == (prefix, event):
                del self._items[(prefix, event)]
                self._key = None
        elif (prefix, event, value) in self._triggers:
            item = self._triggers[(prefix, event, value)]
            key = self._keys[item]
            self._items[key] = item
            del self._triggers[(prefix, event, value)]
            self._key = key

