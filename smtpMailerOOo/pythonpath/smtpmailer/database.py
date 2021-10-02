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

from .unolib import KeyMap

from .unotool import createService
from .unotool import getDateTime
from .unotool import getResourceLocation
from .unotool import getSimpleFile
from .unotool import parseDateTime

from .dbqueries import getSqlQuery

from .dbconfig import g_folder
from .dbconfig import g_jar

from .dbtool import checkDataBase
from .dbtool import createDataSource
from .dbtool import createStaticTable
from .dbtool import getConnectionInfo
from .dbtool import getDataBaseConnection
from .dbtool import getDataSource
from .dbtool import getDataSourceCall
from .dbtool import getDataSourceConnection
from .dbtool import getKeyMapFromResult
from .dbtool import getObjectFromResult
from .dbtool import getRowDict
from .dbtool import getSequenceFromResult
from .dbtool import getValueFromResult
from .dbtool import executeQueries
from .dbtool import executeSqlQueries

from .dbinit import getStaticTables
from .dbinit import getQueries
from .dbinit import getTablesAndStatements

from .configuration import g_identifier

from .logger import logMessage
from .logger import getMessage

import time
import traceback


class DataBase(unohelper.Base):
    def __init__(self, ctx, dbname):
        try:
            print("smtpMailer.DataBase.__init__() 1")
            self._ctx = ctx
            self._dbname = dbname
            self._statement = None
            self._embedded = False
            time.sleep(0.2)
            print("smtpMailer.DataBase.__init__() 2")
            url = getResourceLocation(ctx, g_identifier, g_folder)
            self._url = url + '/' + dbname
            if self._embedded:
                self._path = url + '/' + g_jar
            else:
                self._path = None
            odb = self._url + '.odb'
            exist = getSimpleFile(ctx).exists(odb)
            print("smtpMailer.DataBase.__init__() 3")
            if not exist:
                print("smtpMailer.DataBase.__init__() 4")
                connection = getDataSourceConnection(ctx, self._url)
                error = self._createDataBase(connection)
                if error is None:
                    connection.getParent().DatabaseDocument.storeAsURL(odb, ())
                connection.close()
                print("smtpMailer.DataBase.__init__() 5")
            print("smtpMailer.DataBase.__init__() 6")
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

    @property
    def Connection(self):
        if self._statement is None:
            connection = self.getConnection()
            self._statement = connection.createStatement()
        return self._statement.getConnection()

    def dispose(self):
        if self._statement is not None:
            connection = self._statement.getConnection()
            self._statement.dispose()
            self._statement = None
            #connection.getParent().dispose()
            connection.close()
            print("smtpMailer.DataBase.dispose() *** database: %s closed!!!" % self._dbname)

    def getConnection(self, user='', password=''):
        connection = getDataSourceConnection(self._ctx, self._url, user, password, False)
        return connection

# Procedures called by the DataSource
    def getDataSource(self):
        return self.Connection.getParent()

    def shutdownDataBase(self):
        if self._exist:
            query = getSqlQuery(self._ctx, 'shutdown')
        else:
            query = getSqlQuery(self._ctx, 'shutdownCompact')
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
    def getSpoolerViewQuery(self):
        return getSqlQuery(self._ctx, 'getSpoolerViewQuery')

    def getViewQuery(self):
        return getSqlQuery(self._ctx, 'getViewQuery')

    def insertJob(self, sender, subject, document, recipients, attachments):
        call = self._getCall('insertJob')
        call.setString(1, sender)
        call.setString(2, subject)
        call.setString(3, document)
        call.setArray(4, recipients)
        call.setArray(5, attachments)
        status = call.executeUpdate()
        id = call.getInt(6)
        call.close()
        return id

    def insertMergeJob(self, sender, subject, document, datasource, query, table, identifier, bookmark, recipients, indexes, attachments):
        call = self._getCall('insertMergeJob')
        call.setString(1, sender)
        call.setString(2, subject)
        call.setString(3, document)
        call.setString(4, datasource)
        call.setString(5, query)
        call.setString(6, table)
        call.setString(7, identifier)
        call.setString(8, bookmark)
        call.setArray(9, recipients)
        call.setArray(10, indexes)
        call.setArray(11, attachments)
        status = call.executeUpdate()
        id = call.getInt(12)
        call.close()
        return id

    def deleteJob(self, jobs):
        call = self._getCall('deleteJobs')
        call.setArray(1, jobs)
        status = call.executeUpdate()
        print("DataBase.deleteJob() %s" % status)
        call.close()
        return True

    def getJobState(self, job):
        state = 3
        call = self._getCall('getJobState')
        call.setInt(1, job)
        result = call.executeQuery()
        if result.next():
            state = getValueFromResult(result)
        call.close()
        return state

    def getJobIds(self, batch):
        jobs = ()
        call = self._getCall('getJobIds')
        call.setInt(1, batch)
        result = call.executeQuery()
        if result.next():
            values = getValueFromResult(result)
            jobs = () if values is None else values.getArray(None)
        call.close()
        return jobs

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

# Procedures called by the MailSpooler
    def getSpoolerJobs(self, connection, state=0):
        print("DataBase.getSpoolerJobs() 1")
        jobid = []
        print("DataBase.getSpoolerJobs() 2")
        call = self._getDataBaseCall(connection, 'getSpoolerJobs')
        print("DataBase.getSpoolerJobs() 3")
        call.setInt(1, state)
        print("DataBase.getSpoolerJobs() 4")
        result = call.executeQuery()
        print("DataBase.getSpoolerJobs() 5")
        jobids = getSequenceFromResult(result)
        print("DataBase.getSpoolerJobs() 6")
        call.close()
        print("DataBase.getSpoolerJobs() 7")
        return jobids

    def getRecipient(self, connection, job):
        print("DataBase.getRecipient() 1")
        recipient = None
        call = self._getDataBaseCall(connection, 'getRecipient')
        print("DataBase.getRecipient() 2")
        call.setInt(1, job)
        result = call.executeQuery()
        print("DataBase.getRecipient() 3")
        if result.next():
            print("DataBase.getRecipient() 4")
            recipient = getObjectFromResult(result)
        call.close()
        print("DataBase.getRecipient() 5")
        return recipient

    def getSender(self, connection, batch):
        sender = None
        call = self._getDataBaseCall(connection, 'getSender')
        call.setInt(1, batch)
        result = call.executeQuery()
        if result.next():
            sender = getObjectFromResult(result)
        call.close()
        return sender

    def getAttachments(self, connection, batch):
        attachments = ()
        print("DataBase.getAttachments() 1")
        call = self._getDataBaseCall(connection, 'getAttachments')
        call.setInt(1, batch)
        print("DataBase.getAttachments() 2")
        result = call.executeQuery()
        print("DataBase.getAttachments() 3")
        attachments = getSequenceFromResult(result)
        call.close()
        print("DataBase.getAttachments() 4")
        return attachments

    def getServer(self, connection, user, timeout):
        server = None
        call = self._getDataBaseCall(connection, 'getServer')
        call.setString(1, user)
        call.setInt(2, timeout)
        result = call.executeQuery()
        if result.next():
            server = getRowDict(result)
        call.close()
        return server

    def getBookmark(self, connection, format, identifier):
        bookmark = None
        call = self._getDataBaseCall(connection, 'getBookmark', format)
        call.setString(1, identifier)
        result = call.executeQuery()
        if result.next():
            bookmark = getValueFromResult(result)
        call.close()
        return bookmark

    def setJobState(self, connection, state, jobid):
        call = self._getDataBaseCall(connection, 'setJobState')
        call.setInt(1, state)
        call.setTimestamp(2, getDateTime())
        call.setInt(3, jobid)
        result = call.executeUpdate()
        call.close()

    def setBatchState(self, connection, state, batchid):
        call = self._getDataBaseCall(connection, 'setBatchState')
        call.setInt(1, state)
        call.setTimestamp(2, getDateTime())
        call.setInt(3, batchid)
        result = call.executeUpdate()
        call.close()

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
    def _createDataBase(self, connection):
        version, error = checkDataBase(self._ctx, connection)
        if error is None:
            statement = connection.createStatement()
            createStaticTable(self._ctx, statement, getStaticTables(), True)
            tables, statements = getTablesAndStatements(self._ctx, connection, version)
            executeSqlQueries(statement, tables)
            executeQueries(self._ctx, statement, getQueries())
            statement.close()
        return error

    def _getCall(self, name, format=None):
        return self._getDataBaseCall(self.Connection, name, format)

    def _getDataBaseCall(self, connection, name, format=None):
        return getDataSourceCall(self._ctx, connection, name, format)

    def _getPreparedCall(self, name):
        # TODO: cannot use: call = self.Connection.prepareCommand(name, QUERY)
        # TODO: it trow a: java.lang.IncompatibleClassChangeError
        return self.Connection.prepareCommand(name, QUERY)
