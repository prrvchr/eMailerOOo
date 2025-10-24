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

import unohelper

from ..jdbcdriver import isInstrumented

from ..unotool import getConfiguration
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile
from ..unotool import getStringResource

from ..dbconfig  import g_folder

from ..configuration import g_basename
from ..configuration import g_identifier
from ..configuration import g_separator

import traceback


class OptionsModel():
    def __init__(self, ctx):
        self._ctx = ctx
        self._config = getConfiguration(ctx, g_identifier, True)
        folder = g_folder + g_separator + g_basename
        location = getResourceLocation(ctx, g_identifier, folder)
        self._url = location + '.odb'
        self._instrumented = isInstrumented(ctx, 'xdbc:hsqldb')
        resolver = getStringResource(ctx, g_identifier, 'dialogs', 'OptionsDialog')
        self._link = resolver.resolveString('OptionsDialog.Hyperlink1.Url')

    @property
    def _Timeout(self):
        return self._config.getByName('ConnectTimeout')

    def isInstrumented(self):
        return self._instrumented

    def getViewData(self):
        exist = self.getDataBaseStatus()
        return self._link, self._instrumented, exist, self._Timeout

    def saveTimeout(self, timeout):
        if timeout != self._Timeout:
            self._config.replaceByName('ConnectTimeout', timeout)
        if self._config.hasPendingChanges():
            self._config.commitChanges()
            return True
        return False

    def getDataBaseStatus(self):
        return getSimpleFile(self._ctx).exists(self._url)

    def getDataBaseUrl(self):
        return self._url

