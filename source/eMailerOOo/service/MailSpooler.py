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

from com.sun.star.frame import TerminationVetoException

from com.sun.star.lang import EventObject
from com.sun.star.lang import XComponent

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import XMailSpooler

from com.sun.star.sdb.CommandType import TABLE

from emailer import TerminateListener

from emailer import Cancel
from emailer import DataSource
from emailer import Runner
from emailer import Sender

from emailer import getConfiguration
from emailer import getDesktop
from emailer import getRowSet

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
        self._listeners = []
        self._datacall = None
        self._datasource = None
        self._table = 'SpoolerTable'
        self._logger = getLogger(ctx, g_spoolerlog)
        self._datasource = DataSource(ctx, self, self._logger)
        if MailSpooler._count == 0:
            listener = TerminateListener(self)
            getDesktop(ctx).addTerminateListener(listener)
            MailSpooler._listener = listener
        MailSpooler._count += 1
        self._initRowSet()
        print("MailSpooler.__init__() count: %s" % MailSpooler._count)

    _rowset = None
    _count = 0
    _runner = None
    _cancel = Cancel()
    _listeners = []
    _close = False
    _listener = None

    # XMailSpooler
    def start(self, indicator=None):
        if self.isRunning():
            return
        cancel = MailSpooler._cancel
        cancel.reset()
        job = Sender(self._ctx, self, self._onFinished)
        runner = Runner(self._ctx, job, cancel, indicator, self._logger)
        MailSpooler._runner = runner
        runner.start()
        self._notifySpoolerToggleState()

    def stop(self):
        if self.isRunning():
            MailSpooler._cancel.set()

    def isRunning(self):
        return MailSpooler._runner is not None

    def addSpoolerListener(self, listener):
        if listener not in MailSpooler._listeners:
            MailSpooler._listeners.append(listener)

    def removeSpoolerListener(self, listener):
        if listener in MailSpooler._listeners:
            MailSpooler._listeners.remove(listener)

    def addStatusIndicator(self, indicator):
        if MailSpooler._runner is not None:
            MailSpooler._runner.addStatusIndicator(indicator)

    def removeStatusIndicator(self, indicator):
        if MailSpooler._runner is not None:
            MailSpooler._runner.removeStatusIndicator(indicator)

    def getConnection(self):
        return self._datasource.DataBase.getConnection()

    def getContent(self):
        return self._getRowSet()

    def getSendJobs(self):
        return self._getDataCall().getSendJobs()

    def getJobs(self, jobs):
        return self._getDataCall().getJobs(jobs)

    def addContentListener(self, listener):
        if self._hasRowSet():
            self._getRowSet().addRowSetListener(listener)
            self._executeRowSet()
        else:
            self._listeners.append(listener)

    def removeContentListener(self, listener):
        if self._hasRowSet():
            self._getRowSet().removeRowSetListener(listener)

    def getBatchAttachments(self, batch):
        return self._getDataCall().getAttachments(batch)

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

# XTerminateListener
    def queryTermination(self, source):
        if MailSpooler._runner is not None:
            MailSpooler._close = True
            MailSpooler._cancel.set(disposed=True)
            raise TerminationVetoException()

    def notifyTermination(self, source):
        pass

    # com.sun.star.lang.XComponent
    def dispose(self):
        try:
            MailSpooler._count -= 1
            print("MailSpooler.dispose() count: %s" % MailSpooler._count)
            if self._datacall is not None:
                self._datacall.dispose()
                self._datacall = None
            if MailSpooler._count == 0:
                if self._hasRowSet():
                    self._getRowSet().close()
                    MailSpooler._rowset = None
                if self.isRunning():
                    MailSpooler._cancel.set(disposed=True)
                elif MailSpooler._listener:
                    getDesktop(self._ctx).removeTerminateListener(MailSpooler._listener)
                MailSpooler._listener = None
        except Exception as e:
            print("MailSpooler.dispose ERROR: %s" % traceback.format_exc())

    def addEventListener(self, listener):
        pass
    def removeEventListener(self, listener):
        pass

    # method caller internally
    def _initRowSet(self):
        if not self._hasRowSet():
            rowset = getRowSet(self._ctx)
            rowset.Command = getConfiguration(self._ctx, g_identifier).getByName(self._table)
            rowset.CommandType = TABLE
            rowset.FetchSize = g_fetchsize
            rowset.ActiveConnection = self._datasource.DataBase.getConnection()
            MailSpooler._rowset = rowset
            for listener in self._listeners:
                rowset.addRowSetListener(listener)
            rowset.execute()

    def _getDataCall(self):
        if self._datacall is None:
            self._datacall = self._datasource.getDataCall()
        return self._datacall

    def _getRowSet(self):
        return MailSpooler._rowset

    def _hasRowSet(self):
        return MailSpooler._rowset is not None

    def _executeRowSet(self):
        if self._hasRowSet():
            # FIXME: If RowSet.Filter is not assigned then unassigned, RowSet.RowCount is always 1
            rowset = self._getRowSet()
            rowset.ApplyFilter = True
            rowset.ApplyFilter = False
            rowset.execute()

    def _onFinished(self, success):
        MailSpooler._runner = None
        if MailSpooler._close:
            MailSpooler._close = False
            desktop = getDesktop(self._ctx)
            if MailSpooler._listener:
                desktop.removeTerminateListener(MailSpooler._listener)
            try:
                desktop.terminate()
            except Exception:
                pass
        elif not MailSpooler._cancel.isDisposed():
            self._notifySpoolerToggleState(success)

    def _notifySpoolerToggleState(self, success=True):
        event = EventObject(self)
        for listener in MailSpooler._listeners:
            listener.spoolerToggleState(event, success)


g_ImplementationHelper.addImplementation(MailSpooler,                     # UNO object class
                                         g_ImplementationName,            # Implementation name
                                         g_ServiceNames)                  # List of implemented services

