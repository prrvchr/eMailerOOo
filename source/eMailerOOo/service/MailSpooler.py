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

from emailer import DataSource

from emailer import getLogger

from emailer import g_spoolerlog

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'io.github.prrvchr.eMailerOOo.MailSpooler'
g_ServiceNames = ('com.sun.star.mail.MailSpooler', )


class MailSpooler(unohelper.Base,
                  XMailSpooler,
                  XComponent):
    def __init__(self, ctx):
        self._logger = getLogger(ctx, g_spoolerlog)
        self._datasource = DataSource(ctx, self, self._logger)
        self._datacall = None
        print("MailSpooler.__init__() 1")

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

    def addJob(self, sender, subject, document, recipients, attachments):
        return self._getDataCall().insertJob(sender, subject, document, recipients, attachments)

    def addMergeJob(self, sender, subject, document, datasource, query, table, recipients, filters, attachments):
        return self._getDataCall().insertMergeJob(sender, subject, document, datasource, query, table, recipients, filters, attachments)

    def removeJobs(self, jobids):
        return self._getDataCall().deleteJob(jobids)

    def getJobState(self, jobid):
        return self._getDataCall().getJobState(jobid)

    def getJobIds(self, batchid):
        return self._getDataCall().getJobIds(batchid)

    def getSpoolerJobs(self, state):
        return self._getDataCall().getSpoolerJobs(state)

    def setJobState(self, job, state):
        self._getDataCall().updateJobState(job, state)

    def setJobsState(self, jobs, state):
        self._getDataCall().updateJobsState(jobs, state)

    def updateJob(self, batch, job, threadid, messageid, foreignid, state):
        self._getDataCall().updateSpooler(batch, job, threadid, messageid, foreignid, state)


    # com.sun.star.lang.XComponent
    def dispose(self):
        if self._datacall:
            self._datacall.dispose()
            self._datacall = None

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


g_ImplementationHelper.addImplementation(MailSpooler,                     # UNO object class
                                         g_ImplementationName,            # Implementation name
                                         g_ServiceNames)                  # List of implemented services
