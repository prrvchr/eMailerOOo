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

from .logger import Logger
#from .logger import logMessage
#from .logger import getMessage
#g_message = 'MailSpooler'

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

    def stop(self):
        print("MailSpooler.stop() 1")
        if self.isStarted():
            self._disposed.set()
            self._thread.join()
        print("MailSpooler.stop() 2")

    def start(self):
        print("MailSpooler.start() 1")
        if self.isStarted():
            print("MailSpooler.start() allready started!!!")
            return
        self._disposed.clear()
        jobids = self.DataBase.getSpoolerJobs(0)
        if jobids:
            self._thread = Thread(target=self._run, args=jobids)
            self._thread.start()
        print("MailSpooler.start() 2")

    def isStarted(self):
        return self._thread is not None and self._thread.is_alive()

    def _run(self, *jobids):
        try:
            print("MailSpooler._run()1 begin ****************************************")
            for jobid in jobids:
                if self._disposed.is_set():
                    print("MailSpooler._run() 2 break")
                    break
                self._count = 0
                self._send(jobid)
            #self.DataBase.dispose()
            print("MailSpooler._run() 3 canceled *******************************************")
        except Exception as e:
            msg = "MailSpooler _run(): Error: %s" % traceback.print_exc()
            print(msg)

    def _send(self, jobid):
        print("MailSpooler._send() 1")
        self._logMessage(INFO, 101, jobid)
        #print("MailSpooler._send() JobId: %s" % jobid)
        self.DataBase.setJobState(jobid, 1)
        #print("MailSpooler._send() JobId: %s envoyer" % jobid)
        print("MailSpooler._send() 2")
        self._logMessage(INFO, 102, jobid)
        #if self.Provider.isOffLine():
        #    msg = getMessage(self._ctx, g_message, 101)
        #    logMessage(self._ctx, INFO, msg, 'Replicator', '_synchronize()')
        #else:
        #    self._syncData()

    def _logMessage(self, level, resource, format=None):
        msg = self._logger.getMessage(resource, format)
        self._logger.logMessage(level, msg)
