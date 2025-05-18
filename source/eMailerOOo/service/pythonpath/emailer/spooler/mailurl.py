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

from ..unotool import executeFrameDispatch
from ..unotool import getDocument
from ..unotool import getTempFile

from ..helper import saveTempDocument

import traceback


class MailUrl():
    def __init__(self, ctx, url, merge, filter=None, descriptor=None):
        self._ctx = ctx
        self._url = url.UriReference
        self._merge = merge
        self._filter = filter
        name = url.getPathSegment(url.PathSegmentCount -1)
        self._name = self._title = name
        self._temp = self._document = None
        if self._isTemp():
            self._temp = getTempFile(ctx).Uri
            self._document = getDocument(ctx, self._url)
            if descriptor is not None:
                self.merge(descriptor)
            elif not self._merge:
                self._name = self._saveTempDocument()

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
    def Main(self):
        if self._isTemp():
            url = self._temp
        else:
            url = self._url
        return url

    def merge(self, descriptor):
        self._setDocumentRecord(descriptor)
        self._name = self._saveTempDocument()

    def dispose(self):
        if self._isTemp():
            self._document.close(True)
            self._temp = None

# Private Procedures call
    def _isTemp(self):
        return any((self._merge, self._filter))

    def _setDocumentRecord(self, descriptor):
        url = None
        if self._document.supportsService('com.sun.star.text.TextDocument'):
            url = '.uno:DataSourceBrowser/InsertContent'
        elif self._document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
            url = '.uno:DataSourceBrowser/InsertColumns'
        if url is not None:
            frame = self._document.CurrentController.Frame
            executeFrameDispatch(self._ctx, frame, url, None, *descriptor)

    def _saveTempDocument(self):
        return saveTempDocument(self._document, self._temp, self._title, self._filter)

