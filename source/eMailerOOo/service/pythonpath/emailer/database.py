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

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .unotool import checkVersion
from .unotool import getSimpleFile

from .dbqueries import getSqlQuery

from .dbinit import createDataBase
from .dbinit import getDataBaseConnection

from .helper import checkConnection

from .configuration import g_basename

from .dbconfig import g_version

from threading import Lock
from threading import Thread
import traceback


class DataBase(unohelper.Base):
    def __init__(self, ctx, source, logger, url, warn, user='', pwd=''):
        self._ctx = ctx
        self._lock = Lock()
        self._statement = None
        self._url = url
        odb = url + '.odb'
        new = not getSimpleFile(ctx).exists(odb)
        connection = getDataBaseConnection(ctx, url, user, pwd, new)
        checkConnection(ctx, source, connection, logger, new, warn)
        self._version = connection.getMetaData().getDriverVersion()
        if new and DataBase._init is None:
            DataBase._init = Thread(target=createDataBase, args=(ctx, connection, odb))
            DataBase._init.start()
        else:
            connection.close()
        self._new = new

    _init = None

    @property
    def Version(self):
        return self._version

    @property
    def Url(self):
        return self._url

    @property
    def Connection(self):
        if self._statement is None:
            with self._lock:
                if self._statement is None:
                    self._statement = self.getConnection().createStatement()
        return self._statement.getConnection()

    def dispose(self):
        if self._statement is not None:
            with self._lock:
                if self._statement is not None:
                    connection = self._statement.getConnection()
                    self._statement.close()
                    connection.close()
                    self._statement = None
                    print("smtpMailer.DataBase.dispose() *** database: %s closed!!!" % g_basename)

    def wait(self):
        if DataBase._init and DataBase._init.is_alive():
            DataBase._init.join()
            #DataBase._init = None

    def getConnection(self, user='', password=''):
        return getDataBaseConnection(self._ctx, self._url, user, password)

# Procedures called by the DataSource
    def shutdownDataBase(self):
        if self._new:
            query = getSqlQuery(self._ctx, 'shutdownCompact')
        else:
            query = getSqlQuery(self._ctx, 'shutdown')
        self._statement.execute(query)

