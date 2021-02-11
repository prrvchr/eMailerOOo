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
from com.sun.star.sdb.CommandType import QUERY

from unolib import KeyMap
from unolib import parseDateTime

from .connection import Connection

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
from .dbtools import getSequenceFromResult

from .dbinit import getStaticTables
from .dbinit import getQueries
from .dbinit import getTablesAndStatements

from .configuration import g_identifier

from .logger import logMessage
from .logger import getMessage

import time
import traceback


class DataBase(unohelper.Base):
    def __init__(self, ctx, dbname, user='SA', password='', sync=None):
        try:
            print("DataBase.__init__() 1")
            self._ctx = ctx
            self._error = None
            time.sleep(0.5)
            datasource, url, self._created = getDataSource(ctx, dbname, g_identifier, True)
            print("DataBase.__init__() 2")
            #connection = Connection(ctx, datasource, datasource.URL, user, password, sync, True)
            connection = datasource.getConnection(user, password)
            print("DataBase.__init__() 3")
            self._statement = connection.createStatement()
            print("DataBase.__init__() 4")
            if self._created:
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
        return self.Connection.getParent()

    def addCloseListener(self, listener):
        self.getDataSource().DatabaseDocument.addCloseListener(listener)

    def shutdownDataBase(self):
        if self._created:
            query = getSqlQuery(self._ctx, 'shutdownCompact')
        else:
            query = getSqlQuery(self._ctx, 'shutdown')
        self._statement.execute(query)

# Procedures called by the Server
    def getSmtpConfig(self, email):
        domain = email.partition('@')[2]
        user = KeyMap()
        servers = []
        call = self._getCall('getServers')
        call.setString(1, email)
        call.setString(2, domain)
        result = call.executeQuery()
        user.setValue('Server', call.getString(3))
        user.setValue('Port', call.getShort(4))
        user.setValue('LoginName', call.getString(5))
        user.setValue('Password', call.getString(6))
        user.setValue('Domain', domain)
        while result.next():
            servers.append(getKeyMapFromResult(result))
        call.close()
        return user, servers

    def setSmtpConfig(self, config):
        timestamp = parseDateTime()
        provider = config.getValue('Provider')
        name = config.getValue('DisplayName')
        shortname = config.getValue('DisplayShortName')
        self._mergeProvider(provider, name, shortname, timestamp)
        domains = config.getValue('Domains')
        self._mergeDomain(provider, domains, timestamp)
        call = self._getMergeServerCall(provider)
        i = 1
        servers = config.getValue('Servers')
        for server in servers:
            print("DataBase.setSmtpConfig() server: %s" % i)
            i += 1
            self._callMergeServer(call, server, timestamp)
        call.close()
        return servers

    def mergeProvider(self, provider):
        timestamp = parseDateTime()
        self._mergeProvider(provider, provider, provider, timestamp)

    def mergeServer(self, provider, server):
        timestamp = parseDateTime()
        call = self._getMergeServerCall(provider)
        self._callMergeServer(call, server, timestamp)
        call.close()

    def updateServer(self, host, port, server):
        timestamp = parseDateTime()
        call = self._getCall('updateServer')
        call.setString(1, host)
        call.setShort(2, port)
        call.setString(3, server.getValue('Server'))
        call.setShort(4, server.getValue('Port'))
        call.setByte(5, server.getValue('Connection'))
        call.setByte(6, server.getValue('Authentication'))
        call.setTimestamp(7, timestamp)
        result = call.executeUpdate()
        print("DataBase.updateServer() %s" % result)
        call.close()

    def mergeUser(self, email, user):
        timestamp = parseDateTime()
        call = self._getCall('mergeUser')
        call.setString(1, email)
        call.setString(2, user.getValue('Server'))
        call.setShort(3, user.getValue('Port'))
        call.setString(4, user.getValue('LoginName'))
        call.setString(5, user.getValue('Password'))
        call.setTimestamp(6, timestamp)
        result = call.executeUpdate()
        print("DataBase.mergeUser() %s" % result)
        call.close()

# Procedures called by the Spooler
    def getRowSetCommand(self):
        return getSqlQuery(self._ctx, 'getRowSetCommand')

    def getQueryCommand(self):
        return getSqlQuery(self._ctx, 'getQueryCommand')

    def getRowSetOrder(self):
        return getSqlQuery(self._ctx, 'getRowSetOrder')

    def insertJob(self, sender, subject, document, recipient, attachment, separator):
        call = self._getCall('insertJob')
        call.setString(1, sender)
        call.setString(2, subject)
        call.setString(3, document)
        call.setString(4, recipient)
        call.setString(5, attachment)
        call.setString(6, separator)
        status = call.executeUpdate()
        id = call.getInt(7)
        call.close()
        return id

# Procedures called by the Mailer
    def getSenders(self):
        senders = []
        call = self._getCall('getSenders')
        result = call.executeQuery()
        senders = getSequenceFromResult(result)
        call.close()
        return tuple(senders)

    def deleteUser(self, user):
        call = self._getCall('deleteUser')
        call.setString(1, user)
        status = call.executeUpdate()
        call.close()
        return status

# Procedures called internally by the Server
    def _mergeProvider(self, provider, name, shortname, timestamp):
        call = self._getCall('mergeProvider')
        call.setString(1, provider)
        call.setString(2, name)
        call.setString(3, shortname)
        call.setTimestamp(4, timestamp)
        result = call.executeUpdate()
        call.close()

    def _mergeDomain(self, provider, domains, timestamp):
        call = self._getMergeDomainCall(provider)
        for domain in domains:
            call.setString(2, domain)
            call.setTimestamp(3, timestamp)
            result = call.executeUpdate()
        call.close()

    def _getMergeServerCall(self, provider):
        call = self._getCall('mergeServer')
        call.setString(1, provider)
        return call

    def _getMergeDomainCall(self, provider):
        call = self._getCall('mergeDomain')
        call.setString(1, provider)
        return call

    def _callMergeServer(self, call, server, timestamp):
        call.setString(2, server.getValue('Server'))
        call.setShort(3, server.getValue('Port'))
        call.setByte(4, server.getValue('Connection'))
        call.setByte(5, server.getValue('Authentication'))
        call.setByte(6, server.getValue('LoginMode'))
        call.setTimestamp(7, timestamp)
        result = call.executeUpdate()

# Procedures called internally
    def _createDataBase(self):
        version, error = checkDataBase(self._ctx, self.Connection)
        if error is None:
            createStaticTable(self._ctx, self._statement, getStaticTables(), True)
            tables, statements = getTablesAndStatements(self._ctx, self._statement, version)
            executeSqlQueries(self._statement, tables)
            executeQueries(self._ctx, self._statement, getQueries())
        return error

    def _storeDataBase(self, url):
        self.Connection.getParent().DatabaseDocument.storeAsURL(url, ())

    def _getCall(self, name, format=None):
        return getDataSourceCall(self._ctx, self.Connection, name, format)

    def _getPreparedCall(self, name):
        # TODO: cannot use: call = self.Connection.prepareCommand(name, QUERY)
        # TODO: it trow a: java.lang.IncompatibleClassChangeError
        return self.Connection.prepareCommand(name, QUERY)
