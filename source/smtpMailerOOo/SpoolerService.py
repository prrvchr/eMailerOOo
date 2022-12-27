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

from com.sun.star.lang import XServiceInfo

from com.sun.star.mail import XSpoolerService

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.ucb.ConnectionMode import OFFLINE

from smtpmailer import DataSource
from smtpmailer import MailSpooler
from smtpmailer import Pool

from smtpmailer import getConnectionMode
from smtpmailer import g_identifier
from smtpmailer import g_dns

g_service = 'SpoolerService'

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'com.sun.star.mail.%s' % g_service


class SpoolerService(unohelper.Base,
                     XServiceInfo,
                     XSpoolerService):
    def __init__(self, ctx):
        self._ctx = ctx
        self._datasource = DataSource(ctx)
        self._logger = Pool(ctx).getLogger('SpoolerLogger')
        self._logger.setDebugMode(True)
        self._listeners = []

    __spooler = None

    @property
    def _spooler(self):
        return SpoolerService.__spooler

    # XSpoolerService
    def start(self):
        if not self.isStarted():
            self._logger.logResource(INFO, 101)
            if self._isOffLine():
                self._logger.logResource(INFO, 102)
            else:
                SpoolerService.__spooler = self._getSpooler()
                self._logger.logResource(INFO, 103)
                self._spooler.started()
                self._spooler.start()

    def stop(self):
        if self.isStarted():
            self._spooler.setDisposed()
            self._spooler.join()

    def isStarted(self):
        return self._spooler is not None and self._spooler.is_alive()

    def addJob(self, sender, subject, document, recipients, attachments):
        return self._datasource.insertJob(sender, subject, document, recipients, attachments)

    def addMergeJob(self, sender, subject, document, datasource, query, table, recipients, filters, attachments):
        return self._datasource.insertMergeJob(sender, subject, document, datasource, query, table, recipients, filters, attachments)

    def removeJobs(self, jobids):
        return self._datasource.deleteJob(jobids)

    def getJobState(self, jobid):
        return self._datasource.getJobState(jobid)

    def getJobIds(self, batchid):
        return self._datasource.getJobIds(batchid)

    def addListener(self, listener):
        self._listeners.append(listener)
        if self.isStarted():
            self._spooler.addListener(listener)

    def removeListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)
        if self.isStarted():
            self._spooler.removeListener(listener)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

# Private methods
    def _getSpooler(self):
        self._datasource.waitForDataBase()
        spooler = MailSpooler(self._ctx, self._datasource.DataBase, self._logger, self._listeners)
        return spooler

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE


g_ImplementationHelper.addImplementation(SpoolerService,                            # UNO object class
                                         g_ImplementationName,                      # Implementation name
                                        (g_ImplementationName,))                    # List of implemented services
