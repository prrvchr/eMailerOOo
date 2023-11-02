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

from .mailerlib import getTransferable

from .mailertool import getTransferableMimeValues

from .unotool import getMimeTypeFactory
from .unotool import getSequenceInputStream
from .unotool import getSimpleFile
from .unotool import hasInterface

import traceback


class Transferable():
    def __init__(self, ctx, logger):
        self._ctx = ctx
        self._logger = logger
        self._mtf = getMimeTypeFactory(ctx)
        self._charset = 'charset'
        self._default = 'utf-8'
        self._encoding = self._default
        self._encode = False
        self._uiname = 'E Documents'

    @property
    def Encoding(self):
        return self._encoding
    @Encoding.setter
    def Encoding(self, encoding):
        self._encode = True
        self._encoding = encoding

# XTransferableFactory
    def getBySequence(self, sequence):
        stream = getSequenceInputStream(self._ctx, sequence)
        uiname, mimetype = self._getMimeValues({'InputStream': stream})
        return self._getTransferable(uiname, self._getMineType(mimetype), sequence)

    def getByUrl(self, url):
        stream = getSimpleFile(self._ctx).openFileRead(url)
        uiname, mimetype = self._getMimeValues({'URL': url, 'InputStream': stream})
        return self._getTransferable(uiname, self._getMineType(mimetype), stream)

    def getByStream(self, stream):
        uiname, mimetype = self._getMimeValues({'InputStream': stream})
        return self._getTransferable(uiname, self._getMineType(mimetype), stream)

    def getByString(self, data):
        sequence = uno.ByteSequence(data.encode(self._encoding))
        stream = getSequenceInputStream(self._ctx, sequence)
        uiname, mimetype = self._getMimeValues({'InputStream': stream})
        return self._getTransferable(uiname, self._getMineType(mimetype, True), data)

# Private methods
    def _getMimeValues(self, descriptor):
        return getTransferableMimeValues(self._ctx, descriptor, self._uiname)

    def _getMineType(self, mimetype, force=False):
        if force or self._encode:
            mct = self._mtf.createMimeContentType(mimetype)
            if not mct.hasParameter(self._charset):
                mimetype += '; %s=%s' % (self._charset, self._encoding)
        self._encoding = self._default
        self._encode = False
        return mimetype

    def _getTransferable(self, uiname, mimetype, data):
        flavor = self._getFlavor(uiname, mimetype, data)
        return getTransferable(self._logger, flavor, data)

    def _getFlavor(self, uiname, mimetype, data):
        flavor = uno.createUnoStruct('com.sun.star.datatransfer.DataFlavor')
        flavor.HumanPresentableName = uiname
        flavor.MimeType = mimetype
        flavor.DataType = self._getDataType(data)
        return flavor

    def _getDataType(self, data):
        if isinstance(data, str):
            dtype = uno.getTypeByName('string')
        elif  isinstance(data, uno.ByteSequence):
            dtype = uno.getTypeByName('[]byte')
        elif hasInterface(data, 'com.sun.star.io.XInputStream'):
            dtype = uno.getTypeByName('com.sun.star.io.XInputStream')
        return dtype
