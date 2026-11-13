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

from ..mail import MailModel

from ..unotool import getDesktop
from ..unotool import getDocument


class MailerModel(MailModel):
    def __init__(self, ctx, path):
        super().__init__(ctx)
        self._url = None
        self._path = path
        self._resources = {'DialogTitle':     'MailerDialog.Title',
                           'Property':        'Mail.Document.Property.%s',
                           'Document':        'MailWindow.Label8.Label',
                           'MsgBoxTitle':     'MessageBox.Error.Title',
                           'MsgBoxMsg1':      'MessageBox.Error.Message.1',
                           'MsgBoxMsg2':      'MessageBox.Error.Message.2',
                           'MsgBoxMsg3':      'MessageBox.Error.Message.3',
                           'MsgBoxMsg4':      'MessageBox.Error.Message.4'}

# MailerModel getter methods
    def getInitSteps(self):
        yield self._getDocument
        yield self._getDocumentTitle

    def isSubjectValid(self, subject):
        return subject != ''

    def getPath(self):
        return self._path

    def getEmailData(self):
        return self.getUrl(), False

    def getAttachmentData(self, attachment):
        url, _, export = self.parseUriFragment(attachment)
        return url, False, export

    def getDocument(self, url=None, readonly=True):
        if url is None:
            url = self._url
        return getDocument(self._ctx, url, readonly)

    def getUrl(self):
        return self._url

# MailerModel setter methods
    def setUrl(self, url):
        self._url = url

    def viewAttachment(self, attachment):
        url, _, _ = self.parseUriFragment(attachment)
        getDesktop(self._ctx).loadComponentFromURL(url, '_default', 0, ())

    def closeDocument(self, document):
        document.close(True)

# MailerModel private getter methods
    def _getDocument(self, url):
        document = getDocument(self._ctx, url, readonly=True, template=True)
        return document, url

    def _getDocumentTitle(self, document, url):
        resource = self._resources.get('DialogTitle')
        title = self._resolver.resolveString(resource) + url
        return document, title
