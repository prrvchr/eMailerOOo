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

from .database import DataBase

from unolib import createService

from .logger import setDebugMode
from .logger import logMessage
from .logger import getMessage

from threading import Thread

import traceback


class DataSource(unohelper.Base):
    def __init__(self, ctx, *args):
        print("DataSource.__init__() 1")
        self.ctx = ctx
        self._dbcontext = createService(self.ctx, 'com.sun.star.sdb.DatabaseContext')
        if not self._isInitialized():
            print("DataSource.__init__() 2")
            DataSource._Init = Thread(target=self._initDataBase)
            DataSource._Init.start()
        print("DataSource.__init__() 3")

    _Init = None
    _Call = None
    _DataBase = None

    @property
    def DataBase(self):
        return DataSource._DataBase

    def getAvailableDataSources(self):
        return self._dbcontext.getElementNames()

    def setDataSource(self, *args):
        init = Thread(target=self._setDataSource, args=args)
        init.start()

    def setRowSet(self, *args):
        DataSource._Call = Thread(target=self._setRowSet, args=args)
        DataSource._Call.start()

    def initPage2(self, *args):
        init = Thread(target=self._initPage2, args=args)
        init.start()

    def executeRecipient(self, *args):
        init = Thread(target=self._executeRecipient, args=args)
        init.start()

    def executeAddress(self, *args):
        init = Thread(target=self._executeAddress, args=args)
        init.start()

    def executeRowSet(self, *args):
        init = Thread(target=self._executeRowSet, args=args)
        init.start()

# Procedures called internally
    def _isInitialized(self):
        return DataSource._Init is not None

    def _initDataBase(self):
        DataSource._DataBase = DataBase(self.ctx)

    def _waitForDataBase(self):
        DataSource._Init.join()

    def _waitForRowSet(self):
        DataSource._Call.join()

    def _setDataSource(self, progress, *args):
        progress(5)
        self._waitForDataBase()
        progress(10)
        self.DataBase.setDataSource(self._dbcontext, progress, *args)

    def _setRowSet(self, callback, *args):
        if self.DataBase.setRowSet(*args):
            callback()

    def _initPage2(self, callback, *args):
        self._waitForRowSet()
        self.DataBase.initPage2(*args)
        callback()

    def _executeRecipient(self, callback, *args):
        self.DataBase.executeRecipient(*args)
        callback()

    def _executeAddress(self, callback, *args):
        self.DataBase.executeAddress(*args)
        callback()

    def _executeRowSet(self, callback, *args):
        self.DataBase.executeRowSet(*args)
        callback()
