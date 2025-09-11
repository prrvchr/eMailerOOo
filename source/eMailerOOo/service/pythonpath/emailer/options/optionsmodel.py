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

from ..helper import getMailSpooler

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
        self._spooler = getMailSpooler(ctx)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'OptionsDialog')
        self._resources = {'SpoolerStatus':  'OptionsDialog.Label5.Label.%s'}
        folder = g_folder + g_separator + g_basename
        location = getResourceLocation(ctx, g_identifier, folder)
        self._url = location + '.odb'

    @property
    def _Timeout(self):
        return self._config.getByName('ConnectTimeout')

    def dispose(self, listener):
        if self._spooler is not None:
            self._spooler.removeListener(listener)
            self._spooler.dispose()

    def addStreamListener(self, listener):
        if self._spooler is not None:
            self._spooler.addListener(listener)

    def getViewData(self):
        exist = self.getDataBaseStatus()
        state, status = self._getSpoolerStatus()
        return exist, self._Timeout, state, status

    def saveTimeout(self, timeout):
        if timeout != self._Timeout:
            self._config.replaceByName('ConnectTimeout', timeout)
        if self._config.hasPendingChanges():
            self._config.commitChanges()
            return True
        return False

    def getSpoolerStatus(self, started):
        resource = self._resources.get('SpoolerStatus') % started
        return started, self._resolver.resolveString(resource)

    def getDataBaseStatus(self):
        return getSimpleFile(self._ctx).exists(self._url)

    def getDataBaseUrl(self):
        return self._url

    # OptionsModel private methods
    def _getSpoolerStatus(self):
        started = 0
        if self._spooler is not None:
            started = int(self._spooler.isStarted())
        return self.getSpoolerStatus(started)

