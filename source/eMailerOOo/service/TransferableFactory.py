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

from com.sun.star.datatransfer import XTransferableFactory

from com.sun.star.lang import XServiceInfo

from emailer import Transferable

from emailer import getLogger

from emailer import g_identifier
from emailer import g_mailservicelog

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'io.github.prrvchr.eMailerOOo.TransferableFactory'
g_ServiceNames = ('com.sun.star.datatransfer.TransferableFactory', )


class TransferableFactory(unohelper.Base,
                          XServiceInfo,
                          XTransferableFactory):
    def __init__(self, ctx):
        self._transferable = Transferable(ctx, getLogger(ctx, g_mailservicelog))

    @property
    def Encoding(self):
        return self._transferable.Encoding
    @Encoding.setter
    def Encoding(self, encoding):
        self._transferable.Encoding = encoding


# XTransferableFactory
    def getBySequence(self, sequence):
        return self._transferable.getBySequence(sequence)

    def getByUrl(self, url):
        return self._transferable.getByUrl(url)

    def getByStream(self, stream):
        return self._transferable.getByStream(stream)

    def getByString(self, data):
        return self._transferable.getByString(data)


    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

g_ImplementationHelper.addImplementation(TransferableFactory,             # UNO object class
                                         g_ImplementationName,            # Implementation name
                                         g_ServiceNames)                  # List of implemented services
