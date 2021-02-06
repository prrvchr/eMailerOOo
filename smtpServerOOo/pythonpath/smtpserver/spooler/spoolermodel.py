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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import getPathSettings
from unolib import getStringResource

from ..datasource import DataSource

from ..configuration import g_identifier
from ..configuration import g_extension

from ..logger import logMessage
from ..logger import getMessage

import validators
import traceback


class SpoolerModel(unohelper.Base):
    def __init__(self, ctx):
        self._ctx = ctx
        self._path = getPathSettings(ctx).Work
        self._datasource = DataSource(ctx)
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)
        self._started = False

    @property
    def DataSource(self):
        return self._datasource

    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def getRowSet(self):
        return self._datasource.getRowSet()

    def executeRowSet(self, rowset):
        if self._datasource.isInitialized():
            rowset.execute()
