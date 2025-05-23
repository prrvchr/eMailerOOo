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

from com.sun.star.sdb.CommandType import QUERY

from .unotool import checkVersion
from .unotool import getSimpleFile

from .dbqueries import getSqlQuery

from .dbtool import getDataSourceCall
from .dbtool import getDataFromResult
from .dbtool import getObjectFromResult
from .dbtool import getResultValue
from .dbtool import getSequenceFromResult

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
        self._init = None
        self._statement = None
        self._url = url
        odb = url + '.odb'
        new = not getSimpleFile(ctx).exists(odb)
        connection = getDataBaseConnection(ctx, url, user, pwd, new)
        checkConnection(ctx, source, connection, logger, new, warn)
        self._version = connection.getMetaData().getDriverVersion()
        self._quote = connection.getMetaData().getIdentifierQuoteString()
        if new:
            self._init = Thread(target=createDataBase, args=(ctx, connection, odb))
            self._init.start()
        else:
            connection.close()
        self._new = new

    @property
    def Version(self):
        return self._version
    @property
    def Error(self):
        return self._error if self._error else ''
    @property
    def Url(self):
        return self._url
    @property
    def IdentifierQuoteString(self):
        return self._quote

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

    def canConnect(self):
        return self._error is None

    def isUptoDate(self):
        return checkVersion(self._version, g_version)

    def wait(self):
        if self._init is not None:
            self._init.join()

    def getConnection(self, user='', password=''):
        return getDataBaseConnection(self._ctx, self._url, user, password)

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
        if self._new:
            query = getSqlQuery(self._ctx, 'shutdownCompact')
        else:
            query = getSqlQuery(self._ctx, 'shutdown')
        self._statement.execute(query)

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

# Procedures called by the Spooler
    def getSpoolerJobs(self, state=0):
        call = self._getCall('getSpoolerJobs')
        call.setInt(1, state)
        result = call.executeQuery()
        jobids = getSequenceFromResult(result)
        call.close()
        return jobids, len(jobids)

    def getRecipient(self, job):
        recipient = None
        call = self._getCall('getRecipient')
        call.setInt(1, job)
        result = call.executeQuery()
        if result.next():
            recipient = getObjectFromResult(result)
        call.close()
        return recipient

    def getMailer(self, batch):
        mailer = None
        call = self._getCall('getMailer')
        call.setInt(1, batch)
        result = call.executeQuery()
        if result.next():
            mailer = getDataFromResult(result)
        call.close()
        return mailer

    def getAttachments(self, batch):
        attachments = ()
        call = self._getCall('getAttachments')
        call.setInt(1, batch)
        result = call.executeQuery()
        attachments = getSequenceFromResult(result)
        call.close()
        return attachments

    def updateSpooler(self, batchid, jobid, threadid, messageid, foreignid, state):
        call = self._getCall('updateSpooler')
        call.setInt(1, batchid)
        call.setInt(2, jobid)
        call.setString(3, threadid)
        call.setString(4, messageid)
        call.setString(5, foreignid)
        call.setInt(6, state)
        call.executeUpdate()
        call.close()

# Procedures called internally
    def _getCall(self, name, format=None):
        return self._getDataBaseCall(self.Connection, name, format)

    def _getDataBaseCall(self, connection, name, format=None):
        return getDataSourceCall(self._ctx, connection, name, format)

    def _getPreparedCall(self, name):
        # TODO: cannot use: call = self.Connection.prepareCommand(name, QUERY)
        # TODO: it trow a: java.lang.IncompatibleClassChangeError
        return self.Connection.prepareCommand(name, QUERY)
