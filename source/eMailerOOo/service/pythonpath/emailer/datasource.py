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

from com.sun.star.util import XCloseListener

from .datacall import DataCall

from .database import DataBase

from .dbtool import getConnectionUrl

from .helper import checkConfiguration

from .dbconfig import g_folder

from .configuration import g_basename
from .configuration import g_separator

import traceback


class DataSource(unohelper.Base,
                 XCloseListener):
    def __init__(self, ctx, source, logger, warn=False):
        self._ctx = ctx
        checkConfiguration(ctx, source, logger, warn)
        url = getConnectionUrl(ctx, g_folder + g_separator + g_basename)
        self._database = DataBase(ctx, source, logger, url, warn)

    @property
    def DataBase(self):
        return self._database

    def getDataCall(self):
        return DataCall(self._ctx, self.DataBase.getConnection())

    def dispose(self):
        self.waitForDataBase()
        self.DataBase.dispose()

    def waitForDataBase(self):
        self.DataBase.wait()

    # XCloseListener
    def queryClosing(self, source, ownership):
        self.DataBase.shutdownDataBase()

    def notifyClosing(self, source):
        pass

