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

import unohelper

from com.sun.star.lang import XComponent

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import XMailSpooler

from com.sun.star.sdb.CommandType import TABLE

from emailer import DataSource

from emailer import createService
from emailer import getConfiguration

from emailer import getLogger

from emailer import g_fetchsize
from emailer import g_spoolerlog
from emailer import g_identifier

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'io.github.prrvchr.eMailerOOo.MailSpooler'
g_ServiceNames = ('com.sun.star.mail.MailSpooler', )


class MailSpooler(unohelper.Base,
                  XMailSpooler,
                  XComponent):
    def __init__(self, ctx):
        self._ctx = ctx
        self._logger = getLogger(ctx, g_spoolerlog)
        self._datasource = DataSource(ctx, self, self._logger)
        self._table = getConfiguration(ctx, g_identifier, False).getByName('SpoolerTable')
        self._datacall = None
        self._close = False
        print("MailSpooler.__init__() 1")

    _rowset = None

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    # XMailSpooler
    def getConnection(self):
        self._datasource.waitForDataBase()
        return self._datasource.DataBase.getConnection()

    def getContent(self):
        self._close = True
        return self._getRowSet()

    def getSendJobs(self):
        return self._getDataCall().getSendJobs()

    def getJobs(self, jobs):
        return self._getDataCall().getJobs(jobs)

    def addContentListener(self, listener):
        if self._hasRowSet():
            self._getRowSet().addRowSetListener(listener)
            self._executeRowSet()

    def removeContentListener(self, listener):
        if self._hasRowSet():
            self._getRowSet().removeRowSetListener(listener)

    def getRecipientContent(self, job):
        return self._getDataCall().getJobRecipient(job)

    def getBatchAttachments(self, batch):
        return self._getDataCall().getAttachments(batch)

    def getBatchContent(self, batch):
        return self._getDataCall().getBatchMailer(batch)

    def addJob(self, sender, subject, document, recipients, attachments):
        added = self._getDataCall().addJob(sender, subject, document, recipients, attachments)
        self._executeRowSet()
        return added

    def addMergeJob(self, sender, subject, document, datasource, query, table, recipients,
                          filters, predicates, addresses, identifiers, attachments):
        added = self._getDataCall().addMergeJob(sender, subject, document, datasource, query, table, recipients,
                                                filters, predicates, addresses, identifiers, attachments)
        self._executeRowSet()
        return added

    def removeJobs(self, jobs):
        removed = self._getDataCall().deleteJobs(jobs)
        self._executeRowSet()
        return removed

    def getJobState(self, job):
        return self._getDataCall().getJobState(job)

    def getJobIds(self, batch):
        return self._getDataCall().getJobIds(batch)

    def getSpoolerJobs(self, state):
        return self._getDataCall().getSpoolerJobs(state)

    def setJobState(self, job, state):
        self._getDataCall().updateJobState(job, state)
        self._executeRowSet()

    def setJobsState(self, jobs, state):
        self._getDataCall().updateJobsState(jobs, state)
        self._executeRowSet()

    def updateJob(self, batch, job, recipient, threadid, messageid, foreignid, state):
        self._getDataCall().updateSpooler(batch, job, recipient, threadid, messageid, foreignid, state)
        self._executeRowSet()

    def resubmitJobs(self, jobs):
        ids = self._getDataCall().resubmitJobs(jobs)
        self._executeRowSet()
        return ids

    # com.sun.star.lang.XComponent
    def dispose(self):
        if self._datacall:
            self._datacall.dispose()
            self._datacall = None
        if self._close and self._hasRowSet():
            self._getRowSet().close()
            MailSpooler._rowset = None

    def addEventListener(self, listener):
        pass
    def removeEventListener(self, listener):
        pass

    # method caller internally
    def _getDataCall(self):
        print("MailSpooler._getDataCall() 1")
        if self._datacall is None:
            self._datasource.waitForDataBase()
            self._datacall = self._datasource.getDataCall()
        print("MailSpooler._getDataCall() 2")
        return self._datacall

    def _getRowSet(self):
        if MailSpooler._rowset is None:
            table = getConfiguration(self._ctx, g_identifier, False).getByName('SpoolerTable')
            rowset = createService(self._ctx, 'com.sun.star.sdb.RowSet')
            rowset.Command = table
            rowset.CommandType = TABLE
            rowset.FetchSize = g_fetchsize
            rowset.ActiveConnection = self.getConnection()
            MailSpooler._rowset = rowset
        return MailSpooler._rowset

    def _hasRowSet(self):
        return MailSpooler._rowset is not None

    def _executeRowSet(self):
        if MailSpooler._rowset:
            # FIXME: If RowSet.Filter is not assigned then unassigned, RowSet.RowCount is always 1
            rowset = MailSpooler._rowset
            rowset.ApplyFilter = True
            rowset.ApplyFilter = False
            rowset.execute()


g_ImplementationHelper.addImplementation(MailSpooler,                     # UNO object class
                                         g_ImplementationName,            # Implementation name
                                         g_ServiceNames)                  # List of implemented services
