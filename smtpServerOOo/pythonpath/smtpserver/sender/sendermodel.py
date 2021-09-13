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

from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import getDesktop
from smtpserver import getFileUrl
from smtpserver import getMessage
from smtpserver import logMessage
from smtpserver import getPropertyValueSet
from smtpserver import getStringResource
from smtpserver import getUrlPresentation
from smtpserver import g_identifier
from smtpserver import g_extension

from collections import OrderedDict
from threading import Thread
import json
import traceback


class SenderModel(unohelper.Base):
    def __init__(self, ctx, datasource, path, close):
        self._ctx = ctx
        self._path = path
        self._datasource = datasource
        self._close = close
        self._disposed = False
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Title': 'SenderDialog.Title',
                           'PickerTitle': 'Sender.FilePicker.Title',
                           'PickerFilters': 'Sender.FilePicker.Filters'}

    @property
    def DataSource(self):
        return self._datasource
    @property
    def Path(self):
        return self._path

# SenderModel getter methods
    def isDisposed(self):
        return self._disposed

    def getDocumentUrl(self):
        title = self._getFilePickerTitle()
        filters = self._getFilePickerFilters()
        url, self._path = getFileUrl(self._ctx, title, self._path, filters)
        return url

    def getDocument(self, *args):
        Thread(target=self._getDocument, args=args).start()

# SenderModel setter methods
    def dispose(self):
        self._disposed = True
        if self._close:
            self.DataSource.dispose()

# SenderModel private getter methods
    def _getDocument(self, url, initMailer):
        # TODO: Document can be <None> if a lock or password exists !!!
        # TODO: It would be necessary to test a Handler on the descriptor...
        location = getUrlPresentation(self._ctx, url)
        properties = {'Hidden': True, 'MacroExecutionMode': ALWAYS_EXECUTE_NO_WARN}
        descriptor = getPropertyValueSet(properties)
        document = getDesktop(self._ctx).loadComponentFromURL(location, '_blank', 0, descriptor)
        title = self._getDocumentTitle(document.URL)
        initMailer(document, title)

    def _getDocumentTitle(self, url):
        resource = self._resources.get('Title')
        title = self._resolver.resolveString(resource)
        return title + url

    def _getFilePickerTitle(self):
        resource = self._resources.get('PickerTitle')
        title = self._resolver.resolveString(resource)
        return title

    def _getFilePickerFilters(self):
        resource = self._resources.get('PickerFilters')
        filter = self._resolver.resolveString(resource)
        filters = json.loads(filter, object_pairs_hook=OrderedDict)
        return filters.items()
