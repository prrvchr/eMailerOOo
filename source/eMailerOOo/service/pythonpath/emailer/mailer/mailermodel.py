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

from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from ..mail import MailModel

from ..unotool import getDesktop
from ..unotool import getDocument
from ..unotool import getFileUrl
from ..unotool import getPropertyValueSet
from ..unotool import getUrlPresentation

from collections import OrderedDict
from threading import Thread
import json
import traceback


class MailerModel(MailModel):
    def __init__(self, ctx, datasource, path, close):
        MailModel.__init__(self, ctx, datasource, close)
        self._path = path
        self._url = None
        self._resources = {'DialogTitle':   'MailerDialog.Title',
                           'PickerTitle':   'Mail.FilePicker.Title',
                           'PickerFilters': 'Mail.FilePicker.Filters',
                           'Property':      'Mail.Document.Property.%s',
                           'Document':      'MailWindow.Label8.Label'}

# MailerModel getter methods
    def isSubjectValid(self, subject):
        return subject != ''

    def getDocumentUrl(self):
        title = self.getFilePickerTitle()
        filters = self._getFilePickerFilters()
        url, self._path = getFileUrl(self._ctx, title, self._path, filters)
        return url

    def getPath(self):
        return self._path

# MailerModel setter methods
    def loadDocument(self, *args):
        Thread(target=self._loadDocument, args=args).start()

# MailerModel private getter methods
    def _getFilePickerFilters(self):
        resource = self._resources.get('PickerFilters')
        filter = self._resolver.resolveString(resource)
        filters = json.loads(filter, object_pairs_hook=OrderedDict)
        return filters.items()

    def _getDocumentTitle(self, url):
        resource = self._resources.get('DialogTitle')
        title = self._resolver.resolveString(resource)
        return title + url

# SenderModel private setter methods
    def _loadDocument(self, url, initView):
        # TODO: Document can be <None> if a lock or password exists !!!
        # TODO: It would be necessary to test a Handler on the descriptor...
        location = getUrlPresentation(self._ctx, url)
        properties = {'Hidden': True, 'MacroExecutionMode': ALWAYS_EXECUTE_NO_WARN}
        descriptor = getPropertyValueSet(properties)
        document = getDesktop(self._ctx).loadComponentFromURL(location, '_blank', 0, descriptor)
        title = self._getDocumentTitle(document.URL)
        initView(document, title)








    def getUrl(self):
        return self._url

    def getDocument(self, url=None):
        if url is None:
            url = self._url
        document = getDocument(self._ctx, url)
        return document

    def setUrl(self, url):
        self._url = url
