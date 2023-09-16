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

from com.sun.star.util import XCloseListener

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.sdbc.DataType import CHAR
from com.sun.star.sdbc.DataType import VARCHAR
from com.sun.star.sdbc.DataType import LONGVARCHAR

from .database import DataBase

from .dbtool import Array

from threading import Thread
from threading import Lock
import traceback
import time


class DataSource(unohelper.Base,
                 XCloseListener):
    def __init__(self, ctx, dbname='SmtpMailer'):
        self._ctx = ctx
        self._dbtypes = (CHAR, VARCHAR, LONGVARCHAR)
        if self.DataBase is None:
            with self._lock:
                if self.DataBase is None:
                    DataSource.__database = DataBase(ctx, dbname)

    __lock = Lock()
    __database = None

    @property
    def DataBase(self):
        return DataSource.__database
    @property
    def _lock(self):
        return DataSource.__lock

    def dispose(self):
        self.waitForDataBase()
        self.DataBase.dispose()

    # XCloseListener
    def queryClosing(self, source, ownership):
        self.DataBase.shutdownDataBase()

    def notifyClosing(self, source):
        pass

# Procedures called by IspdbManager
    # called by WizardPage2.activatePage()
    def getServerConfig(self, servers, email):
        return self.DataBase.getServerConfig(servers, email)

    def setServerConfig(self, services, servers, config):
        self.DataBase.setServerConfig(services, servers, config)

    # called by WizardPage3 / WizardPage4
    def saveUser(self, *args):
        self.DataBase.mergeUser(*args)

    def saveServer(self, new, provider, server, host, port):
        if new:
            self.DataBase.mergeProvider(provider)
            self.DataBase.mergeServer(provider, server)
        else:
            self.DataBase.updateServer(host, port, server)

    def waitForDataBase(self):
        self.DataBase.wait()

# Procedures called by the Mailer
    def getSenders(self, *args):
        Thread(target=self._getSenders, args=args).start()

    def removeSender(self, sender):
        return self.DataBase.deleteUser(sender)

# Procedures called by the Merger
    def getFilterValue(self, value, dbtype):
        return "'%s'" % value if dbtype in self._dbtypes else "%s" % value

    def getFilter(self, identifier, value, dbtype):
        return '"%s" = %s' % (identifier, self.getFilterValue(value, dbtype))

# Procedures called by the SpoolerService
    def insertJob(self, sender, subject, document, recipient, attachment):
        recipients = Array('VARCHAR', recipient)
        attachments = Array('VARCHAR', attachment)
        id = self.DataBase.insertJob(sender, subject, document, recipients, attachments)
        return id

    def insertMergeJob(self, sender, subject, document, datasource, query, table, recipient, filter, attachment):
        recipients = Array('VARCHAR', recipient)
        filters = Array('VARCHAR', filter)
        attachments = Array('VARCHAR', attachment)
        id = self.DataBase.insertMergeJob(sender, subject, document, datasource, query, table, recipients, filters, attachments)
        return id

    def deleteJob(self, job):
        jobs = Array('INTEGER', job)
        return self.DataBase.deleteJob(jobs)

    def getJobState(self, job):
        return self.DataBase.getJobState(job)

    def getJobIds(self, batch):
        return self.DataBase.getJobIds(batch)

# Procedures called internally by the Mailer
    def _getSenders(self, callback):
        self.waitForDataBase()
        senders = self.DataBase.getSenders()
        callback(senders)

# Procedures called internally by the Spooler
    def _initSpooler(self, callback):
        self.waitForDataBase()
        callback()

# Private methods
    def _initDataBase(self, dbname):
        database = DataBase(self._ctx, dbname)
        with self._lock:
            DataSource.__database = database
