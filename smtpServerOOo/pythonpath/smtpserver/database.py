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

from .dbtools import checkDataBase
from .dbtools import createStaticTable
from .dbtools import executeSqlQueries
from .dbtools import getDataSourceCall
from .dbtools import executeQueries
from .dbtools import getKeyMapFromResult

from .dbinit import getStaticTables
from .dbinit import getQueries
from .dbinit import getTablesAndStatements

from .logger import logMessage
from .logger import getMessage

import traceback


class DataBase(unohelper.Base):
    def __init__(self, ctx, datasource, name='', password='', sync=None):
        self.ctx = ctx
        self._statement = datasource.getConnection(name, password).createStatement()
        self.sync = sync

    @property
    def Connection(self):
        return self._statement.getConnection()

# Procedures called by the DataSource
    def createDataBase(self):
        version, error = checkDataBase(self.ctx, self.Connection)
        if error is None:
            createStaticTable(self.ctx, self._statement, getStaticTables(), True)
            tables, statements = getTablesAndStatements(self.ctx, self._statement, version)
            executeSqlQueries(self._statement, tables)
            executeQueries(self.ctx, self._statement, getQueries())
        return error

    def getDataSource(self):
        return self.Connection.getParent().DatabaseDocument.DataSource

    def storeDataBase(self, url):
        self.Connection.getParent().DatabaseDocument.storeAsURL(url, ())

    def addCloseListener(self, listener):
        self.Connection.Parent.DatabaseDocument.addCloseListener(listener)

    def shutdownDataBase(self, compact=False):
        if compact:
            query = getSqlQuery(self.ctx, 'shutdownCompact')
        else:
            query = getSqlQuery(self.ctx, 'shutdown')
        self._statement.execute(query)

# Procedures called by the Identifier

# Procedures called by the Replicator

# Procedures called internally
    def _getCall(self, name, format=None):
        return getDataSourceCall(self.ctx, self.Connection, name, format)

    def _getPreparedCall(self, name):
        # TODO: cannot use: call = self.Connection.prepareCommand(name, QUERY)
        # TODO: it trow a: java.lang.IncompatibleClassChangeError
        #query = self.Connection.getQueries().getByName(name).Command
        #self._CallsPool[name] = self.Connection.prepareCall(query)
        return self.Connection.prepareCommand(name, QUERY)
