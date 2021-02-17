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

from unolib import getStringResource
from unolib import getPropertyValueSet
from unolib import getDesktop
from unolib import getUrlPresentation

from smtpserver import g_identifier
from smtpserver import g_extension

from smtpserver import logMessage
from smtpserver import getMessage

from threading import Thread
import traceback


class MergerModel(unohelper.Base):
    def __init__(self, ctx):
        self._ctx = ctx
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

# MergerModel StringRessoure methods
    def getDocumentTitle(self, url):
        resource = self._getTitleRessource()
        title = self.resolveString(resource)
        return title + url

    def getFilePickerTitle(self):
        resource = self._getFilePickerTitleResource()
        title = self.resolveString(resource)
        return title

    def getFilePickerFilter(self):
        resource = self._getFilePickerFilterResource()
        filter = self.resolveString(resource)
        return filter

# MergerModel StringRessoure private methods
    def _getTitleRessource(self):
        return 'MergerDialog.Title'


# MergerModel private methods
