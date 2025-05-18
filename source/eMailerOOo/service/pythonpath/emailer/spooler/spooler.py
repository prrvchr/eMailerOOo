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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.lang import XServiceInfo

from com.sun.star.mail import XMailSpooler

from com.sun.star.ucb.ConnectionMode import OFFLINE

from com.sun.star.uno import Exception as UnoException

from .mailer import Mailer

from ..datasource import DataSource

from ..logger import RollerHandler

from ..unotool import getConnectionMode

from ..configuration import g_dns

from threading import Thread
from threading import Event
import traceback


class Spooler(unohelper.Base,
              XServiceInfo,
              XMailSpooler):
    def __init__(self, ctx, lock, logger, implementation):
        self._ctx = ctx
        self._stop = Event()
        self._listeners = []
        self._lock = lock
        self._dataSource = DataSource(ctx, self, logger)
        self._logger = logger

    _thread = Thread()

    # XMailSpooler
    def start(self):
        with self._lock:
            if not self._getThread().is_alive():
                Spooler._thread = Thread(target=self._run)
                self._stop.clear()
                self._getThread().start()
                self._started()

    def terminate(self):
        with self._lock:
            if self._getThread().is_alive():
                self._stop.set()

    def isStarted(self):
        with self._lock:
            return self._getThread().is_alive()

    def addListener(self, listener):
        self._listeners.append(listener)

    def removeListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def addJob(self, sender, subject, document, recipients, attachments):
        self._dataSource.waitForDataBase()
        return self._dataSource.insertJob(sender, subject, document, recipients, attachments)

    def addMergeJob(self, sender, subject, document, datasource, query, table, recipients, filters, attachments):
        self._dataSource.waitForDataBase()
        return self._dataSource.insertMergeJob(sender, subject, document, datasource, query, table, recipients, filters, attachments)

    def removeJobs(self, jobids):
        self._dataSource.waitForDataBase()
        return self._dataSource.deleteJob(jobids)

    def getJobState(self, jobid):
        self._dataSource.waitForDataBase()
        return self._dataSource.getJobState(jobid)

    def getJobIds(self, batchid):
        self._dataSource.waitForDataBase()
        return self._dataSource.getJobIds(batchid)

# Private setter methods
    def _started(self):
        for listener in self._listeners:
            listener.started()

    def _closed(self):
        for listener in self._listeners:
            listener.closed()

    def _terminated(self):
        for listener in self._listeners:
            listener.terminated()

    def _error(self, e):
        for listener in self._listeners:
            listener.error(e)

    def _run(self):
        handler = RollerHandler(self._ctx, self._logger.Name)
        self._logger.addRollerHandler(handler)
        self._logger.logprb(INFO, 'MailSpooler', 'start', 1011)
        if self._isOffLine():
            self._logger.logprb(INFO, 'MailSpooler', 'start', 1012)
        else:
            self._logger.logprb(INFO, 'MailSpooler', 'start', 1013)
            self._dataSource.waitForDataBase()
            jobs, total = self._dataSource.DataBase.getSpoolerJobs()
            if total > 0:
                self._logger.logprb(INFO, 'MailSpooler', 'start', 1014, total)
                count = self._sendMails(jobs)
                self._logger.logprb(INFO, 'MailSpooler', self._getMethod(), 1015, count, total)
            else:
                self._logger.logprb(INFO, 'MailSpooler', 'start', 1016)
        self._logger.logprb(INFO, 'MailSpooler', self._getMethod(), self._getResource(1017))
        self._logger.removeRollerHandler(handler)
        if self._stop.is_set():
            self._terminated()
        else:
            self._closed()

# Private getter methods
    def _getThread(self):
        return Spooler._thread

    def _sendMails(self, jobs):
        count = 0
        mailer = Mailer(self._ctx, self, self._dataSource.DataBase, self._logger, True)
        for job in jobs:
            self._logger.logprb(INFO, 'MailSpooler', '_sendMails', 1021, job)
            if self._stop.is_set():
                self._logger.logprb(INFO, 'MailSpooler', '_sendMails', 1022, job)
                break
            try:
                batch, mail = mailer.getMail(job)
                if self._stop.is_set():
                    self._logger.logprb(INFO, 'MailSpooler', '_sendMails', 1022, job)
                    break
                mailer.sendMail(mail)
            except UnoException as e:
                self._logger.logprb(SEVERE, 'MailSpooler', '_sendMails', 1023, job, e.__class__.__name__, e.Message)
                continue
            except Exception as e:
                self._logger.logprb(SEVERE, 'MailSpooler', '_sendMails', 1024, job, e.__class__.__name__, traceback.format_exc())
                continue
            self._dataSource.DataBase.updateSpooler(batch, job, mail.ThreadId, mail.MessageId, mail.ForeignId, 1)
            self._logger.logprb(INFO, 'MailSpooler', '_sendMails', 1025, job)
            count += 1
        mailer.dispose()
        return count

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE

    def _getMethod(self):
        return 'terminate' if self._stop.is_set() else 'close'

    def _getResource(self, code):
        return code + int(self._stop.is_set())

