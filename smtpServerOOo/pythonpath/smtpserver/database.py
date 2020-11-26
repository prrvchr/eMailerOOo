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
        try:
            print("DataBase.__init__() 1")
            self.ctx = ctx
            self._error = None
            datasource, url, created = getDataSource(self.ctx, dbname, g_identifier, True)
            print("DataBase.__init__() 2")
            connection = datasource.getConnection(name, password)
            print("DataBase.__init__() 3")
            self._statement = connection.createStatement()
            print("DataBase.__init__() 4")
            if created:
                print("DataBase.__init__() 5")
                self._error = self._createDataBase()
                if self._error is None:
                    self._storeDataBase(url)
            print("DataBase.__init__() 6")
            self.sync = sync
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

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

    def getSmtpConfig(self, email):
        user = KeyMap()
        servers = []
        call = self._getCall('getUser')
        call.setString(1, email)
        result = call.executeQuery()
        user.setValue('Server', call.getString(2))
        user.setValue('Port', call.getShort(3))
        user.setValue('LoginName', call.getString(4))
        user.setValue('Password', call.getString(5))
        user.setValue('Domain', call.getString(6))
        print("DataBase.getSmtpConfig() %s" % (user, ))
        while result.next():
            servers.append(getKeyMapFromResult(result))
        call.close()
        return user, servers

    def setSmtpConfig(self, config):
        timestamp = parseDateTime()
        provider = config.getValue('Provider')
        call = self._getCall('mergeProvider')
        call.setString(1, provider)
        call.setString(2, config.getValue('DisplayName'))
        call.setString(3, config.getValue('DisplayShortName'))
        call.setTimestamp(4, timestamp)
        result = call.executeUpdate()
        call.close()
        call = self._getCall('mergeDomain')
        call.setString(1, provider)
        for domain in config.getValue('Domains'):
            call.setString(2, domain)
            call.setTimestamp(3, timestamp)
            result = call.executeUpdate()
        call.close()
        call = self._getCall('mergeServer')
        call.setString(1, provider)
        i = 1
        servers = config.getValue('Servers')
        for server in servers:
            print("DataBase.setSmtpConfig() server: %s" % i)
            i += 1
            call.setString(2, server.getValue('Server'))
            call.setShort(3, server.getValue('Port'))
            call.setByte(4, server.getValue('Connection'))
            call.setByte(5, server.getValue('Authentication'))
            call.setByte(6, server.getValue('LoginMode'))
            call.setTimestamp(7, timestamp)
            result = call.executeUpdate()
        call.close()
        return servers

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
