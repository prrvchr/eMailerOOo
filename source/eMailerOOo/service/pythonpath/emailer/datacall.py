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

from .dbqueries import getSqlQuery

from .dbtool import Array

from .dbtool import getDataSourceCall
from .dbtool import getDataFromResult
from .dbtool import getObjectFromResult
from .dbtool import getResultValue
from .dbtool import getSequenceFromResult

import traceback


class DataCall():
    def __init__(self, ctx, connection):
        self._ctx = ctx
        self._connection = connection
        self._calls = {}

    def dispose(self):
        print("DataCall.dispose() 1")
        for name, call in self._calls.items():
            call.close()
            print("DataCall.dispose() 2 call: %s closed!!!" % name)

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

# Procedures called by the Spooler
    def getSpoolerViewQuery(self):
        return getSqlQuery(self._ctx, 'getSpoolerViewQuery')

    def getViewQuery(self):
        return getSqlQuery(self._ctx, 'getViewQuery')

    def addJob(self, sender, subject, document, recipients, attachments):
        call = self._getCall('insertJob')
        call.setString(1, sender)
        call.setString(2, subject)
        call.setString(3, document)
        call.setArray(4, Array('VARCHAR', recipients))
        call.setArray(5, Array('VARCHAR', attachments))
        status = call.executeUpdate()
        id = call.getInt(6)
        return id

    def addMergeJob(self, sender, subject, document, datasource, query, table, recipients,
                          filters, predicates, addresses, identifiers, attachments):
        call = self._getCall('insertMergeJob')
        call.setString(1, sender)
        call.setString(2, subject)
        call.setString(3, document)
        call.setString(4, datasource)
        call.setString(5, query)
        call.setString(6, table)
        call.setArray(7, Array('VARCHAR', recipients))
        call.setArray(8, Array('VARCHAR', filters))
        call.setArray(9, Array('VARCHAR', predicates))
        call.setArray(10, Array('VARCHAR', addresses))
        call.setArray(11, Array('VARCHAR', identifiers))
        call.setArray(12, Array('VARCHAR', attachments))
        status = call.executeUpdate()
        id = call.getInt(13)
        return id

    def deleteJobs(self, jobs):
        call = self._getCall('deleteJobs')
        call.setArray(1, Array('INTEGER', jobs))
        status = call.executeUpdate()
        return True

    def getJobState(self, job):
        state = 3
        call = self._getCall('getJobState')
        call.setInt(1, job)
        result = call.executeQuery()
        if result.next():
            state = getResultValue(result)
        return state

    def getJobIds(self, batch):
        jobs = ()
        call = self._getCall('getJobIds')
        call.setInt(1, batch)
        result = call.executeQuery()
        if result.next():
            jobs = getResultValue(result)
        return jobs

    def resubmitJobs(self, jobs):
        ids = ()
        call = self._getCall('resubmitJobs')
        call.setArray(1, Array('INTEGER', jobs))
        call.execute()
        value = call.getArray(2)
        if not call.wasNull() and value is not None:
            ids = value.getArray(None)
        print("DataCall.resubmitJobs() ids: %s" % (ids, ))
        return ids

# Procedures called by the Spooler
    def getSendJobs(self):
        call = self._getCall('getSendJobs')
        result = call.executeQuery()
        return result

    def getJobs(self, jobs):
        call = self._getCall('getJobs')
        call.setArray(1, Array('INTEGER', jobs))
        result = call.executeQuery()
        return result

    def getSpoolerJobs(self, state):
        call = self._getCall('getSpoolerJobs')
        call.setInt(1, state)
        result = call.executeQuery()
        jobids = getSequenceFromResult(result)
        return jobids

    def getRecipient(self, job):
        recipient = None
        call = self._getCall('getRecipient')
        call.setInt(1, job)
        result = call.executeQuery()
        if result.next():
            recipient = getObjectFromResult(result)
        return recipient

    def getJobRecipient(self, job):
        call = self._getCall('getRecipient')
        call.setInt(1, job)
        rs = call.executeQuery()
        return rs

    def getMailer(self, batch):
        mailer = None
        call = self._getCall('getMailer')
        call.setInt(1, batch)
        result = call.executeQuery()
        if result.next():
            mailer = getDataFromResult(result)
        return mailer

    def getBatchMailer(self, batch):
        call = self._getCall('getMailer')
        call.setInt(1, batch)
        rs = call.executeQuery()
        return rs

    def getAttachments(self, batch):
        attachments = ()
        call = self._getCall('getAttachments')
        call.setInt(1, batch)
        result = call.executeQuery()
        attachments = getSequenceFromResult(result)
        return attachments

    def updateSpooler(self, batch, job, recipient, threadid, messageid, foreignid, state):
        call = self._getCall('updateSpooler')
        call.setInt(1, batch)
        call.setInt(2, job)
        call.setString(3, recipient)
        call.setString(4, threadid)
        call.setString(5, messageid)
        call.setString(6, foreignid)
        call.setInt(7, state)
        call.executeUpdate()

    def updateJobState(self, job, state):
        call = self._getCall('updateJobState')
        call.setInt(1, job)
        call.setInt(2, state)
        call.executeUpdate()

    def updateJobsState(self, jobs, state):
        call = self._getCall('updateJobsState')
        call.setArray(1, Array('INTEGER', jobs))
        call.setInt(2, state)
        call.executeUpdate()


# Procedures called internally
    def _getCall(self, name, format=None):
        return self._getDataBaseCall(name, format)

    def _getDataBaseCall(self, name, format=None):
        if name not in self._calls:
            self._calls[name] = getDataSourceCall(self._ctx, self._connection, name, format)
        return self._calls[name]
