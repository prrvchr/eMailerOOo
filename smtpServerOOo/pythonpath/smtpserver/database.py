#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE
from com.sun.star.sdb.CommandType import QUERY

from unolib import KeyMap
from unolib import parseDateTime

from .dbqueries import getSqlQuery
from .dbconfig import g_role
from .dbconfig import g_dba

from .dbtools import getDataSource
from .dbtools import checkDataBase
from .dbtools import createStaticTable
from .dbtools import executeSqlQueries
from .dbtools import getDataSourceCall
from .dbtools import executeQueries
from .dbtools import getKeyMapFromResult

from .dbinit import getStaticTables
from .dbinit import getQueries
from .dbinit import getTablesAndStatements

from .configuration import g_identifier

from .logger import logMessage
from .logger import getMessage

import traceback


class DataBase(unohelper.Base):
    def __init__(self, ctx, dbname, name='', password='', sync=None):
        print("DataBase.__init__() 1")
        self.ctx = ctx
        self._error = None
        datasource, url, created = getDataSource(self.ctx, dbname, g_identifier, True)
        self._statement = datasource.getConnection(name, password).createStatement()
        print("DataBase.__init__() 2")
        if created:
            print("DataBase.__init__() 3")
            self._error = self._createDataBase()
            if self._error is None:
                self._storeDataBase(url)
        print("DataBase.__init__() 4")
        self.sync = sync

    @property
    def Connection(self):
        return self._statement.getConnection()

# Procedures called by the DataSource
    def getDataSource(self):
        return self.Connection.getParent().DatabaseDocument.DataSource

    def addCloseListener(self, listener):
        self.Connection.Parent.DatabaseDocument.addCloseListener(listener)

    def shutdownDataBase(self, compact=False):
        if compact:
            query = getSqlQuery(self.ctx, 'shutdownCompact')
        else:
            query = getSqlQuery(self.ctx, 'shutdown')
        self._statement.execute(query)

    def getSmtpConfig(self, domain):
        config = None
        call = self._getCall('getSmtpConfig')
        call.setString(1, domain)
        call.setString(2, domain)
        result = call.executeQuery()
        if result.next():
            config = getKeyMapFromResult(result)
        call.close()
        return config

# Procedures called by the Identifier

# Procedures called by the Replicator

# Procedures called internally
    def _createDataBase(self):
        version, error = checkDataBase(self.ctx, self.Connection)
        if error is None:
            createStaticTable(self.ctx, self._statement, getStaticTables(), True)
            tables, statements = getTablesAndStatements(self.ctx, self._statement, version)
            executeSqlQueries(self._statement, tables)
            executeQueries(self.ctx, self._statement, getQueries())
        return error

    def _storeDataBase(self, url):
        self.Connection.getParent().DatabaseDocument.storeAsURL(url, ())

    def _getCall(self, name, format=None):
        return getDataSourceCall(self.ctx, self.Connection, name, format)

    def _getPreparedCall(self, name):
        # TODO: cannot use: call = self.Connection.prepareCommand(name, QUERY)
        # TODO: it trow a: java.lang.IncompatibleClassChangeError
        #query = self.Connection.getQueries().getByName(name).Command
        #self._CallsPool[name] = self.Connection.prepareCall(query)
        return self.Connection.prepareCommand(name, QUERY)
