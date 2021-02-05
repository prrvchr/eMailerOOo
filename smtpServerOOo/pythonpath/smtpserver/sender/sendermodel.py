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

from unolib import getPathSettings
from unolib import getStringResource
from unolib import getPropertyValueSet
from unolib import getDesktop

from smtpserver import g_identifier
from smtpserver import g_extension

from smtpserver import logMessage
from smtpserver import getMessage

from threading import Thread
import traceback


class SenderModel(unohelper.Base):
    def __init__(self, ctx, path):
        self._ctx = ctx
        self._path = path
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)

    @property
    def Path(self):
        return self._path

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def getDocument(self, *args):
        thread = Thread(target=self._getDocument, args=args)
        thread.start()

    def getDocumentTitle(self, url, resource):
        title = self.resolveString(resource)
        return title + url

    def _getDocument(self, url, callback):
        properties = {'Hidden': True, 'MacroExecutionMode': ALWAYS_EXECUTE_NO_WARN}
        descriptor = getPropertyValueSet(properties)
        document = getDesktop(self._ctx).loadComponentFromURL(url, '_blank', 0, descriptor)
        callback(document)
