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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpmailer import MailModel

from smtpmailer import getDocument
from smtpmailer import getDesktop
from smtpmailer import getMessage
from smtpmailer import getPropertyValueSet
from smtpmailer import getStringResource
from smtpmailer import logMessage
from smtpmailer import g_identifier
from smtpmailer import g_extension

import traceback


class MailerModel(MailModel):
    def __init__(self, ctx, datasource, path):
        self._ctx = ctx
        self._datasource = datasource
        self._path = path
        self._url = None
        self._disposed = False
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'PickerTitle': 'Mail.FilePicker.Title',
                           'Property': 'Mail.Document.Property.%s',
                           'Document': 'MailWindow.Label8.Label.1'}

# MailerModel getter methods
    def getUrl(self):
        return self._url

    def getDocument(self, url=None):
        if url is None:
            url = self._url
        document = getDocument(self._ctx, url)
        return document

# MailerModel setter methods
    def setUrl(self, url):
        self._url = url
