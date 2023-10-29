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

from com.sun.star.ucb.ConnectionMode import OFFLINE

from com.sun.star.uno import Exception as UnoException

from ..listener import TerminateListener

from ..unotool import getConnectionMode
from ..unotool import getSimpleFile
from ..unotool import getTempFile

from ..mailertool import getDataSource

from ..logger import getLogger
from ..logger import RollerHandler

from ..configuration import g_dns
from ..configuration import g_spoolerlog

from .mailer import Mailer

from threading import Thread
from threading import Lock
from threading import Event
import traceback


class Spooler():
    def __init__(self, ctx, source):
        self._ctx = ctx
        self._source = source
        self._lock = Lock()
        self._stop = Event()
        self._listeners = []
        self._logger = getLogger(ctx, g_spoolerlog)
        self._thread = Thread(target=self._run)
        self.DataSource = None

# Procedures called by TerminateListener and SpoolerService
    def stop(self):
        with self._lock:
            if self._thread.is_alive():
                self._stop.set()
                self._thread.join()

# Procedures called by SpoolerService
    def isStarted(self):
        with self._lock:
            return self._thread.is_alive()

    def start(self):
        with self._lock:
            if not self._thread.is_alive():
                self._thread = Thread(target=self._run)
                self._stop.clear()
                self._thread.start()
                self._started()

    def addListener(self, listener):
        self._listeners.append(listener)

    def removeListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def addJob(self, sender, subject, document, recipients, attachments):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('addJob')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.insertJob(sender, subject, document, recipients, attachments)

    def addMergeJob(self, sender, subject, document, datasource, query, table, recipients, filters, attachments):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('addMergeJob')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.insertMergeJob(sender, subject, document, datasource, query, table, recipients, filters, attachments)

    def removeJobs(self, jobids):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('removeJobs')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.deleteJob(jobids)

    def getJobState(self, jobid):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('getJobState')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.getJobState(jobid)

    def getJobIds(self, batchid):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('getJobIds')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.getJobIds(batchid)

# Private methods
    def _isInitialized(self):
        return self.DataSource is not None

    def _initialize(self, method):
        self.DataSource = getDataSource(self._ctx, method, ' ', self._logError)

    def _logError(self, method, code, *args):
        self._logger.logprb(SEVERE, 'MailSpooler', method, code, *args)

    def _started(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject')
        for listener in self._listeners:
            listener.started(event)

    def _stopped(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject')
        for listener in self._listeners:
            listener.stopped(event)

    def _run(self):
        handler = RollerHandler(self._ctx, self._logger.Name)
        self._logger.addRollerHandler(handler)
        self._logger.logprb(INFO, 'MailSpooler', 'start()', 1001)
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('start')
        if self._isInitialized():
            # FIXME: Configuration has been checked we can continue
            if self._isOffLine():
                self._logger.logprb(INFO, 'MailSpooler', 'start()', 1002)
            else:
                self._logger.logprb(INFO, 'MailSpooler', 'start()', 1003)
                self.DataSource.waitForDataBase()
                jobs, total = self.DataSource.DataBase.getSpoolerJobs()
                if total > 0:
                    self._logger.logprb(INFO, 'MailSpooler', 'start()', 1011, total)
                    count = self._sendMails(jobs)
                    self._logger.logprb(INFO, 'MailSpooler', 'stop()', 1012, count, total)
                else:
                    self._logger.logprb(INFO, 'MailSpooler', 'start()', 1013)
        self._logger.logprb(INFO, 'MailSpooler', 'stop()', 1014)
        self._logger.removeRollerHandler(handler)
        self._stopped()

    def _sendMails(self, jobs):
        mailer = Mailer(self._ctx, self._source, self.DataSource.DataBase, self._logger, True)
        count = 0
        for job in jobs:
            self._logger.logprb(INFO, 'MailSpooler', '_sendMails()', 1021, job)
            if self._stop.is_set():
                self._logger.logprb(INFO, 'MailSpooler', '_sendMails()', 1022, job)
                break
            try:
                mail = mailer.getMail(job)
                mailer.sendMail(mail)
            except UnoException as e:
                self._logger.logprb(SEVERE, 'MailSpooler', '_sendMails()', 1023, job, e.Message)
                continue
            except Exception as e:
                self._logger.logprb(SEVERE, 'MailSpooler', '_sendMails()', 1024, job, str(e), traceback.format_exc())
                continue
            self.DataSource.DataBase.updateRecipient(1, mail.MessageId, job)
            self._logger.logprb(INFO, 'MailSpooler', '_sendMails()', 1025, job)
            count += 1
        mailer.dispose()
        return count

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE

