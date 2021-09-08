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

from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XInitialization

from com.sun.star.mail import XSpoolerService

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import DataSource
from smtpserver import MailSpooler
from smtpserver import TerminateListener

from smtpserver import executeDispatch
from smtpserver import getDesktop
from smtpserver import g_identifier

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
        if self.Spooler is None:
            print("SpoolerService.__init__() 1")
            self._datasource.waitForDataBase()
            print("SpoolerService.__init__() 2")
            event = uno.createUnoStruct('com.sun.star.lang.EventObject')
            event.Source = self
            spooler = MailSpooler(ctx, self._datasource.DataBase, event)
            SpoolerService._spooler = spooler
            listener = TerminateListener(spooler)
            getDesktop(ctx).addTerminateListener(listener)

    _spooler = None

    @property
    def Spooler(self):
        return SpoolerService._spooler

    # XSpoolerService
    def start(self):
        self.Spooler.start()

    def stop(self):
        self.Spooler.stop()

    def isStarted(self):
        return self.Spooler.isStarted()

    def addJob(self, sender, subject, document, recipients, attachments):
        try:
            print("SpoolerService.addJob() %s - %s - %s - %s - %s" % (sender, subject, document, recipients, attachments))
            batchid = self._datasource.insertJob(sender, subject, document, recipients, attachments)
            print("SpoolerService.addJob() %s" % batchid)
            return batchid
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

    def addMergeJob(self, sender, subject, document, datasource, query, table, identifier, bookmark, recipients, identifiers, attachments):
        try:
            print("SpoolerService.addMergeJob() %s - %s - %s - %s - %s - %s - %s - %s - %s - %s - %s" % (sender, subject, document, datasource, query, table, identifier, bookmark, recipients, identifiers, attachments))
            batchid = self._datasource.insertMergeJob(sender, subject, document, datasource, query, table, identifier, bookmark, recipients, identifiers, attachments)
            print("SpoolerService.addMergeJob() %s" % batchid)
            return batchid
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

    def removeJobs(self, jobids):
        return self._datasource.deleteJob(jobids)

    def getJobState(self, jobid):
        return self._datasource.getJobState(jobid)

    def getJobIds(self, batchid):
        return self._datasource.getJobIds(batchid)

    def addListener(self, listener):
        self.Spooler.addListener(listener)

    def removeListener(self, listener):
        self.Spooler.removeListener(listener)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(SpoolerService,                            # UNO object class
                                         g_ImplementationName,                      # Implementation name
                                        (g_ImplementationName,))                    # List of implemented services
