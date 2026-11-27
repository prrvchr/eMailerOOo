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

from ..unotool import getStringResource
from ..unotool import getConfiguration

from ..dbinit import getDataSourceUrl

from ..configuration import g_identifier


class OptionsModel():
    def __init__(self, ctx):
        self._ctx = ctx
        self._url = getDataSourceUrl(ctx)
        self._config = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'OptionsDialog')
        self._resources = {'Link': 'OptionsDialog.Label8.Url.%s'}

    @property
    def _timeout(self):
        return self._config.getByName('ConnectTimeout')

    @_timeout.setter
    def _timeout(self, timeout):
        self._config.replaceByName('ConnectTimeout', timeout)

    @property
    def _mode(self):
        return self._config.getByName('ErrorMode')

    @_mode.setter
    def _mode(self, mode):
        self._config.replaceByName('ErrorMode', mode)

    @property
    def _delay(self):
        return self._config.getByName('RunDelay')

    @_delay.setter
    def _delay(self, delay):
        self._config.replaceByName('RunDelay', delay)

    def getConfig(self):
        return self._config

    def getInitData(self, state):
        return self._timeout, self._mode, self._delay, *self.getRefreshData(state)

    def getRefreshData(self, state):
        return state, state.getMessage(self._ctx), self._getLink(state.code)

    def saveData(self, timeout, mode, delay):
        if timeout != self._timeout:
            self._timeout = timeout
        if mode != self._mode:
            self._mode = mode
        if delay != self._delay:
            self._delay = delay
        if self._config.hasPendingChanges():
            self._config.commitChanges()
            return True
        return False

    def getDataSourceUrl(self):
        return self._url

    def _getLink(self, code):
        resource = self._resources.get('Link') % code
        return self._resolver.resolveString(resource)
