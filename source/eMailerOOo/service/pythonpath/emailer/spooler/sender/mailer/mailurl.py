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

from com.sun.star.frame.FrameSearchFlag import CREATE

from ....unotool import createService
from ....unotool import getDesktop
from ....unotool import getDocument
from ....unotool import getLastNamedParts
from ....unotool import getPathSettings
from ....unotool import getPropertyValueSet
from ....unotool import getTempFile
from ....unotool import getTypeDetection

from ....helper import hasDocumentFilter
from ....helper import saveDocumentTo
from ....helper import saveTempDocument

import traceback


class MailUrl():
    def __init__(self, ctx, uri, merge, filter=''):
        self._ctx = ctx
        self._merge = merge
        self._url = uri.UriReference
        self._name = self._title = uri.getPathSegment(uri.PathSegmentCount -1)
        self._folder = self._getTempFolder(ctx)
        self._temp = document = None
        if merge or filter:
            document = getDocument(ctx, self._url)
            if document and filter and not hasDocumentFilter(document, filter):
                filter = ''
        self._filter = filter
        self._document = document
        self._invalidField = ()
        print("MailUrl.__init__() 1")

    @property
    def Document(self):
        return self._document
    @property
    def Filter(self):
        return self._filter
    @property
    def Merge(self):
        return self._merge
    @property
    def Name(self):
        return self._name
    @property
    def Title(self):
        return self._title
    @property
    def Url(self):
        return self._url
    @property
    def Folder(self):
        return self._folder
    @property
    def Main(self):
        if self._temp:
            url = self._temp
        else:
            url = self._url
        return url

    def setTemp(self, temp):
        self._temp = temp

    def isDocumentInvalid(self):
        return self._isTemp and self._document is None

    def saveTempDocument(self):
        if self._document is not None:
            self._name = saveTempDocument(self._document, self._temp, self._title, self._filter)

    def dispose(self, sf):
        if self._document is not None:
            self._document.close(True)
        if sf.exists(self._folder):
            sf.kill(self._folder)
        self._temp = None

    def export(self, folder, export=None):
        if self._filter:
            self._temp = saveDocumentTo(self._document, folder, self._filter, export)

# Private Procedures call
    def _isTemp(self):
        return self._merge or self._filter

    def _getTempFolder(self, ctx):
        folder, name = getLastNamedParts(getTempFile(ctx).Uri, '/')
        return folder

