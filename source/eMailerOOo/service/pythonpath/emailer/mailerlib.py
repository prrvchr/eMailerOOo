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
from com.sun.star.datatransfer import XTransferable

from com.sun.star.mail import XAuthenticator

from com.sun.star.uno import XCurrentContext

import traceback


def getTransferable(logger, flavor, data):
    return Transferable(logger, flavor, data)


class Transferable(unohelper.Base,
                   XTransferable):
    def __init__(self, logger, flavor, data):
        self._logger = logger
        self._flavor = flavor
        self._data = data

# XTransferable
    def getTransferData(self, flavor):
        if not self.isDataFlavorSupported(flavor):
            msg = self._logger.resolveString(3001, flavor.MimeType, flavor.DataType.typeName)
            raise UnsupportedFlavorException(msg, self)
        return self._data

    def getTransferDataFlavors(self):
        return (self._flavor,)

    def isDataFlavorSupported(self, flavor):
        return flavor.MimeType == self._flavor.MimeType and flavor.DataType == self._flavor.DataType


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

