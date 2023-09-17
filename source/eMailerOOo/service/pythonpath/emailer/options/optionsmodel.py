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

import unohelper

from ..unotool import createService
from ..unotool import getConfiguration
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile
from ..unotool import getStringResource

from ..dbconfig  import g_folder

from ..configuration import g_basename
from ..configuration import g_identifier
from ..configuration import g_extension
from ..configuration import g_separator

import traceback


class OptionsModel():
    def __init__(self, ctx):
        self._ctx = ctx
        self._configuration = getConfiguration(ctx, g_identifier, True)
        self._spooler = createService(ctx, 'com.sun.star.mail.SpoolerService')
        self._resolver = getStringResource(ctx, g_identifier, g_extension, 'OptionsDialog')
        self._resources = {'SpoolerStatus': 'OptionsDialog.Label4.Label.%s'}
        folder = g_folder + g_separator + g_basename
        location = getResourceLocation(ctx, g_identifier, folder)
        self._url = location + '.odb'

    def addSpoolerListener(self, listener):
        self._spooler.addListener(listener)

    def removeSpoolerListener(self, listener):
        self._spooler.removeListener(listener)

    def getViewData(self):
        exist = getSimpleFile(self._ctx).exists(self._url)
        timeout = self._configuration.getByName('ConnectTimeout')
        msg, state = self._getSpoolerStatus()
        return exist, timeout, msg, state

    def saveTimeout(self, timeout):
        self._configuration.replaceByName('ConnectTimeout', timeout)
        if self._configuration.hasPendingChanges():
            self._configuration.commitChanges()

    def toogleSpooler(self, state):
        if state:
            self._spooler.start()
        else:
            self._spooler.stop()

    def getSpoolerStatus(self, started):
        resource = self._resources.get('SpoolerStatus') % started
        return self._resolver.resolveString(resource), started

    def getDataBaseUrl(self):
        return self._url

    # OptionsModel private methods
    def _getSpoolerStatus(self):
        started = int(self._spooler.isStarted())
        resource = self._resources.get('SpoolerStatus') % started
        return self._resolver.resolveString(resource), started

