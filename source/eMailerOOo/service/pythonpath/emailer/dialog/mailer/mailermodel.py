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

from com.sun.star.frame.DispatchResultState import FAILURE
from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from ..mail import MailModel

from ...helper import saveDocumentTo

from ...unotool import createService
from ...unotool import getDesktop
from ...unotool import getDocument
from ...unotool import getFileUrl
from ...unotool import getPropertyValueSet
from ...unotool import getUrlPresentation

from collections import OrderedDict
from threading import Thread
from time import sleep
import json
import traceback


class MailerModel(MailModel):
    def __init__(self, ctx, datasource, path, close):
        super().__init__(ctx, datasource, close)
        self._path = path
        self._url = None
        self._resources = {'DialogTitle':     'MailerDialog.Title',
                           'PickerTitle':     'Mail.FilePicker.Title',
                           'PickerFilters':   'Mail.FilePicker.Filters',
                           'Property':        'Mail.Document.Property.%s',
                           'Document':        'MailWindow.Label8.Label',
                           'MsgBoxTitle':     'MessageBox.Error.Title',
                           'MsgBoxMsg1':      'MessageBox.Error.Message.1',
                           'MsgBoxMsg2':      'MessageBox.Error.Message.2',
                           'MsgBoxMsg3':      'MessageBox.Error.Message.3',
                           'MsgBoxMsg4':      'MessageBox.Error.Message.4'}

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

    def closeDocument(self, document):
        document.close(True)

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
    def _loadDocument(self, url, caller):
        # TODO: Document can be <None> if a lock or password exists !!!
        # TODO: It would be necessary to test a Handler on the descriptor...
        location = getUrlPresentation(self._ctx, url)
        print("MailerModel._loadDocument() url: %s - presentation: %s" % (url, location))
        properties = {'Hidden': True, 'MacroExecutionMode': ALWAYS_EXECUTE_NO_WARN}
        descriptor = getPropertyValueSet(properties)
        document = getDesktop(self._ctx).loadComponentFromURL(location, '_blank', 0, descriptor)
        title = self._getDocumentTitle(document.URL)
        result = document, title
        data = {'call': 'init', 'status': SUCCESS, 'result': result}
        self._callback.addCallback(caller, getPropertyValueSet(data))

    def _viewDocument(self, selection, url, merge, filter, caller):
        # XXX: Breathe
        sleep(0.2)
        status, result = SUCCESS, None
        print("MailerModel._viewDocument() url: %s" % filter)
        if not self._sf.exists(url):
            status, result = FAILURE, self._getMsgBoxMessage(1) % attachment
        else:
            try:
                document = self.getDocument(url)
                if document is None:
                    status, result = FAILURE, self._getMsgBoxMessage(2) % attachment
            except Exception as e:
                status, result = FAILURE, self._getMsgBoxMessage(3) % (attachment, e)
        if status:
            folder = self._getTempFolder()
            export = self._getPdfExport(filter)
            result, name = saveDocumentTo(document, folder, filter, export)
            self.closeDocument(document)
        data = {'call': filter, 'status': status, 'result': result}
        self._callback.addCallback(caller, getPropertyValueSet(data))

    def getUrl(self):
        return self._url

    def getDocument(self, url):
        if url == self._url:
            document = self._document
        else:
            document = getDocument(self._ctx, url)
        return document

    def setUrl(self, url):
        self._url = url
