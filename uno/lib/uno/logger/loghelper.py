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

from ..unotool import createService
from ..unotool import getResourceLocation
from ..unotool import getStringResourceWithLocation

from ..configuration import g_identifier
from ..configuration import g_resource
from ..configuration import g_basename

import traceback


def getLogger(ctx, logger, basename=g_basename):
    return LoggerWrapper(ctx, logger, basename)

def getLoggerName(name):
    return '%s.%s' % (g_identifier, name)


# This logger wrapper allows using variable number of argument in python
# while the UNO API does not allow it
class LoggerWrapper():
    def __init__(self, ctx, name, basename):
        self._ctx = ctx
        self._basename = basename
        self._url = getResourceLocation(ctx, g_identifier, g_resource)
        pool = createService(ctx, 'io.github.prrvchr.jdbcDriverOOo.LoggerPool')
        self._logger = pool.getLocalizedLogger(getLoggerName(name), self._url, basename)

    # XLogger
    @property
    def Name(self):
        return self._logger.Name

    @property
    def Level(self):
        return self._logger.Level
    @Level.setter
    def Level(self, value):
        self._logger.Level = value

    def addLogHandler(self, handler):
        self._logger.addLogHandler(handler)

    def removeLogHandler(selfself, handler):
        self._logger.removeLogHandler(handler)

    def isLoggable(self, level):
        return self._logger.isLoggable(level)

    def log(self, level, message):
        self._logger.log(level, message)

    def logp(self, level, clazz, method, message):
        self._logger.logp(level, clazz, method, message)

    def logrb(self, level, resource, *args):
        if self._logger.hasEntryForId(resource):
            self._logger.logrb(level, resource, args)
        else:
            self._logger.log(level, self._getErrorMessage(resource))

    def logprb(self, level, clazz, method, resource, *args):
        if self._logger.hasEntryForId(resource):
            self._logger.logprb(level, clazz, method, resource, args)
        else:
            self._logger.logp(level, clazz, method, self._getErrorMessage(resource))

    def resolveString(self, resource, *args):
        if self._logger.hasEntryForId(resource):
            return self._logger.resolveString(resource, args)
        else:
            return self._getErrorMessage(resource)

    def addModifyListener(self, listener):
        self._logger.addModifyListener(listener)

    def removeModifyListener(self, listener):
        self._logger.removeModifyListener(listener)

    def _getErrorMessage(self, resource):
        resolver = getStringResourceWithLocation(self._ctx, self._url, 'Logger')
        return resolver.resolveString(141) % (resource, self._url, self._basename)

