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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.ucb.ConnectionMode import OFFLINE

from .unotool import getConnectionMode

from .logger import Logger

from .configuration import g_dns

from threading import Thread
from threading import Event
import traceback


class MailSpooler(unohelper.Base):
    def __init__(self, ctx, database):
        self._ctx = ctx
        self.DataBase = database
        self._count = 0
        self._disposed = Event()
        self._thread = None
        self._logger = Logger(ctx, 'MailSpooler', 'MailSpooler')
        self._logger.setDebugMode(True)

# Procedures called by MailServiceSpooler
    def stop(self):
        if self.isStarted():
            self._disposed.set()
            self._thread.join()

    def start(self):
        if not self.isStarted():
            if self._isOffLine():
                self._logMessage(INFO, 101)
            else:
                jobs = self.DataBase.getSpoolerJobs(0)
                if jobs:
                    self._disposed.clear()
                    self._thread = Thread(target=self._run, args=jobs)
                    self._thread.start()
                #else:
                #    self.DataBase.dispose()

    def isStarted(self):
        return self._thread is not None and self._thread.is_alive()

# Private methods
    def _run(self, *jobs):
        try:
            print("MailSpooler._run()1 begin ****************************************")
            for job in jobs:
                if self._disposed.is_set():
                    print("MailSpooler._run() 2 break")
                    break
                self._count = 0
                self._send(job)
            #self.DataBase.dispose()
            print("MailSpooler._run() 3 canceled *******************************************")
        except Exception as e:
            msg = "MailSpooler _run(): Error: %s" % traceback.print_exc()
            print(msg)

    def _send(self, job):
        print("MailSpooler._send() 1")
        self._logMessage(INFO, 102, job)
        mail = self.DataBase.getJobMail(job)
        print("MailSpooler._send() 2 %s - %s - %s - %s - %s - %s - %s - %s" % (mail.Sender, mail.Subject, mail.Document , mail.DataSource , mail.Query , mail.Recipient , mail.Identifier , mail.Attachments))
        server = self.DataBase.getJobServer(job)
        print("MailSpooler._send() 3 %s - %s - %s - %s - %s - %s - %s - %s" % (server.Server, server.Port, server.Connection , server.Authentication , server.LoginMode , server.User , server.LoginName , server.Password))
        self.DataBase.setJobState(job, 1)
        print("MailSpooler._send() 4")
        self._logMessage(INFO, 102, job)


    def _logMessage(self, level, resource, format=None):
        msg = self._logger.getMessage(resource, format)
        self._logger.logMessage(level, msg)

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE
