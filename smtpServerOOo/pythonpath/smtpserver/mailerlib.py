#!
# -*- coding: utf_8 -*-

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

from com.sun.star.datatransfer import XTransferable
from com.sun.star.mail import XAuthenticator
from com.sun.star.uno import XCurrentContext

from .unotool import createService
from .unotool import getFileSequence

from .mailertool import getAttachmentType

import traceback


class Authenticator(unohelper.Base,
                    XAuthenticator):
    def __init__(self, user):
        self._user = user

# XAuthenticator
    def getUserName(self):
        return self._user['LoginName']

    def getPassword(self):
        return self._user['Password']


class CurrentContext(unohelper.Base,
                     XCurrentContext):
    def __init__(self, context):
        self._context = context

# XCurrentContext
    def getValueByName(self, name):
        return self._context[name]


class MailTransferable(unohelper.Base,
                       XTransferable):
    def __init__(self, ctx, data, html=None):
        print("MailTransferable.__init__() 1")
        self._ctx = ctx
        self._data = data
        self._html = html
        self.texttype = 'text/plain; charset=utf-16'
        if html is None:
            self._mimetype = getAttachmentType(ctx, data)
        elif html:
            self._mimetype = 'text/html; charset=utf-8'
        else:
            self._mimetype = self.texttype
        print("MailTransferable.__init__() 2 %s" % self._mimetype)

# XTransferable
    def getTransferData(self, flavor):
        #mri = createService(self._ctx, 'mytools.Mri')
        #mri.inspect(flavor)
        if flavor.MimeType == self.texttype:
            print("MailTransferable.getTransferData() 1")
            data = self._data
        else:
            print("MailTransferable.getTransferData() 2 %s" % flavor.MimeType)
            lenght, sequence = getFileSequence(self._ctx, self._data)
            data = sequence
        return data

    def getTransferDataFlavors(self):
        flavor = uno.createUnoStruct('com.sun.star.datatransfer.DataFlavor')
        flavor.MimeType = self._mimetype
        if self._html is None:
            flavor.HumanPresentableName = 'E-Documents'
        elif self._html:
            flavor.HumanPresentableName = 'HTML-Documents'
        else:
            flavor.HumanPresentableName = 'Unicode text'
        print("MailTransferable.getTransferDataFlavors() 1 %s" % flavor.MimeType)
        return (flavor,)

    def isDataFlavorSupported(self, flavor):
        support = flavor.MimeType == self._mimetype
        print("MailTransferable.isDataFlavorSupported() 1 %s - %s - %s" % (flavor.MimeType, flavor.DataType, support))
        return support
