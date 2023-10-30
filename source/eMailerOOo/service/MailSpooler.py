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

from com.sun.star.mail import XMailSpooler

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from emailer import Spooler

from threading import Lock
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'com.sun.star.mail.MailSpooler'


class MailSpooler(unohelper.Base,
                  XServiceInfo,
                  XMailSpooler):
    def __init__(self, ctx):
        if self._spooler is None:
            with self._lock:
                if self._spooler is None:
                    MailSpooler.__spooler = Spooler(ctx, self)

    __lock = Lock()
    __spooler = None
    @property
    def _lock(self):
        return MailSpooler.__lock
    @property
    def _spooler(self):
        return MailSpooler.__spooler

    # XMailSpooler
    def start(self):
        if not self.isStarted():
            self._spooler.start()

    def terminate(self):
        if self.isStarted():
            self._spooler.terminate()

    def isStarted(self):
        return self._spooler.isStarted()

    def addListener(self, listener):
        self._spooler.addListener(listener)

    def removeListener(self, listener):
        self._spooler.removeListener(listener)

    def addJob(self, sender, subject, document, recipients, attachments):
        return self._spooler.addJob(sender, subject, document, recipients, attachments)

    def addMergeJob(self, sender, subject, document, datasource, query, table, recipients, filters, attachments):
        return self._spooler.addMergeJob(sender, subject, document, datasource, query, table, recipients, filters, attachments)

    def removeJobs(self, jobids):
        return self._spooler.removeJobs(jobids)

    def getJobState(self, jobid):
        return self._spooler.getJobState(jobid)

    def getJobIds(self, batchid):
        return self._spooler.getJobIds(batchid)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(MailSpooler,                               # UNO object class
                                         g_ImplementationName,                      # Implementation name
                                        (g_ImplementationName,))                    # List of implemented services
