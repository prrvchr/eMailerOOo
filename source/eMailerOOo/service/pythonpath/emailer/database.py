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

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import IMAP

from com.sun.star.sdb.CommandType import QUERY

from com.sun.star.sdbc.DataType import VARCHAR
from com.sun.star.sdbc.DataType import SMALLINT

from .unotool import createService
from .unotool import getDateTime
from .unotool import getResourceLocation
from .unotool import getSimpleFile
from .unotool import getUrlPresentation

from .dbqueries import getSqlQuery

from .dbconfig import g_folder
from .dbconfig import g_jar
from .dbconfig import g_csv

from .dbtool import checkDataBase
from .dbtool import createDataSource
from .dbtool import createStaticTable
from .dbtool import getConnectionInfo
from .dbtool import getDataBaseConnection
from .dbtool import getDataSource
from .dbtool import getDataSourceCall
from .dbtool import getDataSourceConnection
from .dbtool import getDataFromResult
from .dbtool import getObjectFromResult
from .dbtool import getResultValue
from .dbtool import getRowDict
from .dbtool import getRowValue
from .dbtool import getSequenceFromResult
from .dbtool import executeQueries
from .dbtool import executeSqlQueries

from .dbinit import getStaticTables
from .dbinit import getQueries
from .dbinit import getTablesAndStatements

from .configuration import g_identifier

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
            location = getResourceLocation(ctx, g_identifier, g_folder)
            url = getUrlPresentation(ctx, location)
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

# Procedures called by the Merger
    def getInnerJoinTable(self, subquery, identifiers, table):
         return subquery + self._getInnerJoinPart(identifiers, table)

    def _getInnerJoinPart(self, identifiers, table):
        query = ' INNER JOIN %s ON ' % table
        conditions = []
        for identifier in identifiers:
            conditions.append('%s = %s.%s' % (identifier, table, identifier))
        return query + ' AND '.join(conditions)

    def getRecipientColumns(self, emails):
        columns = ', '.join(emails)
        if len(emails) > 1:
            columns = 'COALESCE(%s)' % columns
        return columns

# Procedures called by the DataSource
    def getDataSource(self):
        return self.Connection.getParent()

    def shutdownDataBase(self):
        if self._exist:
            query = getSqlQuery(self._ctx, 'shutdown')
        else:
            query = getSqlQuery(self._ctx, 'shutdownCompact')
        self._statement.execute(query)

# Procedures called by the IspDB Wizard
    def getServerConfig(self, servers, email):
        smtp, imap = SMTP.value, IMAP.value
        domain = email.partition('@')[2]
        user = {}
        call = self._getCall('getServers')
        call.setString(1, email)
        call.setString(2, domain)
        call.setString(3, smtp)
        call.setString(4, imap)
        result = call.executeQuery()
        user['ThreadId'] = call.getString(5)
        user[smtp + 'Server'] = call.getString(6)
        user[imap + 'Server'] = call.getString(7)
        user[smtp + 'Port'] = call.getInt(8)
        user[imap + 'Port'] = call.getInt(9)
        user[smtp + 'Login'] = call.getString(10)
        user[imap + 'Login'] = call.getString(11)
        #TODO: Password need default value to empty string not None
        pwd = call.getString(12)
        user[smtp + 'Password'] = '' if call.wasNull() else pwd
        pwd = call.getString(13)
        user[imap + 'Password'] = '' if call.wasNull() else pwd
        user['Domain'] = domain
        while result.next():
            service = getResultValue(result)
            servers.appendServer(service, getDataFromResult(result))
        call.close()
        return user

    def setServerConfig(self, services, servers, config):
        provider = config.get('Provider')
        name = config.get('DisplayName')
        shortname = config.get('DisplayShortName')
        self._mergeProvider(provider, name, shortname)
        domains = config.get('Domains')
        self._mergeDomain(provider, domains)
        self._mergeServices(services, servers, config, provider)

    def _mergeServices(self, services, servers, config, provider):
        call = self._getMergeServerCall(provider)
        for service in services:
            servers.appendServers(service, self._mergeServers(call, config.get(service)))
        call.close()

    def _mergeServers(self, call, servers):
        for server in servers:
            self._callMergeServer(call, server)
        return servers

    def mergeProvider(self, provider):
        self._mergeProvider(provider, provider, provider)

    def mergeServer(self, provider, server):
        call = self._getMergeServerCall(provider)
        self._callMergeServer(call, server)
        call.close()

    def updateServer(self, server, host, port):
        call = self._getCall('updateServer')
        call.setString(1, server.get('Service'))
        call.setString(2, host)
        call.setShort(3, port)
        call.setString(4, server.get('Server'))
        call.setShort(5, server.get('Port'))
        call.setByte(6, server.get('Connection'))
        call.setByte(7, server.get('Authentication'))
        result = call.executeUpdate()
        call.close()

    def mergeUser(self, email, user):
        smtp, imap = SMTP.value, IMAP.value
        call = self._getCall('mergeUser')
        call.setString(1, email)
        if user.hasThread():
            call.setString(2, user.ThreadId)
        else:
            call.setNull(2, VARCHAR)
        call.setString(3, smtp)
        call.setString(4, user.getServer(smtp))
        call.setShort(5, user.getPort(smtp))
        call.setString(6, user.getLogin(smtp))
        call.setString(7, user.getPassword(smtp))
        if user.hasImapConfig():
            call.setString(8, imap)
            call.setString(9, user.getServer(imap))
            call.setShort(10, user.getPort(imap))
            call.setString(11, user.getLogin(imap))
            call.setString(12, user.getPassword(imap))
        else:
            call.setNull(8, VARCHAR)
            call.setNull(9, VARCHAR)
            call.setNull(10, SMALLINT)
            call.setNull(11, VARCHAR)
            call.setNull(12, VARCHAR)
        result = call.executeUpdate()
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

    def insertMergeJob(self, sender, subject, document, datasource, query, table, recipients, filters, attachments):
        call = self._getCall('insertMergeJob')
        call.setString(1, sender)
        call.setString(2, subject)
        call.setString(3, document)
        call.setString(4, datasource)
        call.setString(5, query)
        call.setString(6, table)
        call.setArray(7, recipients)
        call.setArray(8, filters)
        call.setArray(9, attachments)
        status = call.executeUpdate()
        id = call.getInt(10)
        call.close()
        return id

    def deleteJob(self, jobs):
        call = self._getCall('deleteJobs')
        call.setArray(1, jobs)
        status = call.executeUpdate()
        call.close()
        return True

    def getJobState(self, job):
        state = 3
        call = self._getCall('getJobState')
        call.setInt(1, job)
        result = call.executeQuery()
        if result.next():
            state = getResultValue(result)
        call.close()
        return state

    def getJobIds(self, batch):
        jobs = ()
        call = self._getCall('getJobIds')
        call.setInt(1, batch)
        result = call.executeQuery()
        if result.next():
            jobs = getResultValue(result)
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
        jobid = []
        call = self._getDataBaseCall(connection, 'getSpoolerJobs')
        call.setInt(1, state)
        result = call.executeQuery()
        jobids = getSequenceFromResult(result)
        call.close()
        return jobids, len(jobids)

    def getRecipient(self, connection, job):
        recipient = None
        call = self._getDataBaseCall(connection, 'getRecipient')
        call.setInt(1, job)
        result = call.executeQuery()
        if result.next():
            recipient = getObjectFromResult(result)
        call.close()
        return recipient

    def getMailer(self, connection, batch, timeout):
        mailer = None
        call = self._getDataBaseCall(connection, 'getMailer')
        call.setInt(1, batch)
        call.setInt(2, timeout)
        call.setString(3, SMTP.value)
        call.setString(4, IMAP.value)
        result = call.executeQuery()
        if result.next():
            mailer = getDataFromResult(result)
        call.close()
        return mailer

    def updateMailer(self, connection, batch, thread):
        call = self._getDataBaseCall(connection, 'updateMailer')
        call.setInt(1, batch)
        call.setString(2, thread)
        state = call.executeUpdate()
        call.close()

    def getAttachments(self, connection, batch):
        attachments = ()
        call = self._getDataBaseCall(connection, 'getAttachments')
        call.setInt(1, batch)
        result = call.executeQuery()
        attachments = getSequenceFromResult(result)
        call.close()
        return attachments

    def updateRecipient(self, connection, state, messageid, jobid):
        call = self._getDataBaseCall(connection, 'updateRecipient')
        call.setInt(1, state)
        call.setString(2, messageid)
        call.setInt(3, jobid)
        state = call.executeUpdate()
        call.close()

    def setBatchState(self, connection, state, batchid):
        call = self._getDataBaseCall(connection, 'setBatchState')
        call.setInt(1, state)
        call.setTimestamp(2, getDateTime())
        call.setInt(3, batchid)
        result = call.executeUpdate()
        call.close()

# Procedures called internally by the IspDB Wizard
    def _mergeProvider(self, provider, name, shortname):
        call = self._getCall('mergeProvider')
        call.setString(1, provider)
        call.setString(2, name)
        call.setString(3, shortname)
        result = call.executeUpdate()
        call.close()

    def _mergeDomain(self, provider, domains):
        call = self._getMergeDomainCall(provider)
        for domain in domains:
            call.setString(2, domain)
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

    def _callMergeServer(self, call, server):
        call.setString(2, server.get('Service'))
        call.setString(3, server.get('Server'))
        call.setShort(4, server.get('Port'))
        call.setByte(5, server.get('Connection'))
        call.setByte(6, server.get('Authentication'))
        call.setByte(7, server.get('LoginMode'))
        result = call.executeUpdate()

# Procedures called internally
    def _createDataBase(self, connection):
        version, error = checkDataBase(self._ctx, connection)
        if error is None:
            statement = connection.createStatement()
            createStaticTable(self._ctx, statement, getStaticTables(), g_csv, True)
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