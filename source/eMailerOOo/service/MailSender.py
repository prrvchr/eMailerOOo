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

from com.sun.star.frame.DispatchResultState import FAILURE
from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.lang import IllegalArgumentException

from com.sun.star.io import XActiveDataControl

from com.sun.star.lang import XServiceInfo

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.task import XAsyncJob

from com.sun.star.ucb.ConnectionMode import ONLINE

from com.sun.star.uno import Exception as UNOException

from emailer import Mailer

from emailer import DataCall

from emailer import getConnectionMode
from emailer import getNamedValueSet

from emailer import getLogger

from emailer import getMailSpooler

from emailer import g_spoolerlog
from emailer import g_dns

from threading import Thread
from threading import Event
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'io.github.prrvchr.eMailerOOo.MailSender'
g_ServiceNames = ('com.sun.star.mail.MailSender', 'com.sun.star.task.AsyncJob')


class MailSender(unohelper.Base,
                 XAsyncJob,
                 XActiveDataControl,
                 XServiceInfo):

    def __init__(self, ctx):
        self._ctx = ctx
        self._cls = 'MailSender'
        self._logger = getLogger(ctx, g_spoolerlog)
        self._listeners = []
        self._stop = Event()
        self._logger.logprb(INFO, self._cls, '__init__', 1011)
        self._thread = None


    # com.sun.star.task.XAsyncJob:
    def executeAsync(self, args, listener):
        connection, jobs, env = self._getJobArguments(args)
        if connection is None:
            msg = 'Unable to retrieve database connection'
            raise IllegalArgumentException(msg)
        self._start(connection, jobs, env, listener)

    # com.sun.star.io.XActiveDataControl
    def addListener(self, listener):
        if listener not in self._listeners:
            self._listeners.append(listener)
        # XXX: Adding listener permit to known the actual state
        if self._thread is None:
            listener.closed()
        elif self._isCanceled():
            listener.terminated()
        else:
            listener.started()

    def removeListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def start(self):
        print("MailSender.start() 1")
        spooler = getMailSpooler(self._ctx)
        connection, jobs = spooler.getConnection(), spooler.getSpoolerJobs(0)
        spooler.dispose()
        print("MailSender.start() 2")
        self._start(connection, jobs)

    def terminate(self):
        print("MailSender.terminate() 1")
        self._stop.set()
        self._notifyTerminated()
        print("MailSender.terminate() 2")

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    # method caller by the XTerminateListener
    def dispose(self):
        print("MailSender.dispose() 1")
        if self._thread is not None:
            self._stop.set()
            self._thread.join()
        print("MailSender.dispose() 2")

    # Private getter methods
    def _isOffLine(self):
        return getConnectionMode(self._ctx, *g_dns) != ONLINE

    def _getJobArguments(self, args):
        connection = None
        env = 'EXECUTOR'
        jobs = ()
        for arg in args:
            if arg.Name == 'Environment':
                properties = arg.Value
                for property in properties:
                    if property.Name == 'EnvType':
                        env = property.Value
            elif arg.Name == 'DynamicData':
                properties = arg.Value
                for property in properties:
                    if property.Name == 'Connection':
                        connection = property.Value
                    elif property.Name == 'Jobs':
                        jobs = property.Value
        if not connection or not jobs:
            spooler = getMailSpooler(self._ctx)
            connection, jobs = spooler.getConnection, spooler.getSpoolerJobs(0)
            spooler.dispose()
        return connection, jobs, env

    def _isCanceled(self):
        return self._stop.is_set()

    def _getMethod(self):
        return 'terminate' if self._isCanceled() else 'close'

    # Private setter methods
    def _start(self, connection, jobs, env='EXECUTOR', listener=None):
        try:
            if self._thread is None:
                print("MailSender._start() 1")
                self._stop.clear()
                print("MailSender._start() 2")
                self._thread = Thread(target=self._execute, args=(connection, jobs, env, listener))
                self._notifyStarted()
                self._thread.start()
            print("MailSender._start() 3")
        except Exception as e:
            print("MailSender._start() ERROR: %s" % traceback.print_exc())

    def _execute(self, connection, jobs, env, listener):
        count = 0
        state = FAILURE
        mtd = 'start'
        self._logger.logprb(INFO, self._cls, mtd, 1021)

        if self._isOffLine():
            self._logger.logprb(INFO, self._cls, mtd, 1022)
        elif jobs:
            self._logger.logprb(INFO, self._cls, mtd, 1023)
            total = len(jobs)
            print("MailSender._execute() 1")
            self._logger.logprb(INFO, self._cls, mtd, 1024, total)

            database = DataCall(self._ctx, connection)
            count = self._sendMails(database, jobs)
            state = SUCCESS
            print("MailSender._execute() 2")
            database.dispose()
            connection.close()
            self._logger.logprb(INFO, self._cls, self._getMethod(), 1025, count, total)
        else:
            print("MailSender.start() 3")
            self._logger.logprb(INFO, self._cls, mtd, 1026)

        # XXX: We need to notify listener if given
        if listener:
            properties = ()
            if env == 'DISPATCH':
                result = uno.createUnoStruct('com.sun.star.frame.DispatchResultEvent')
                result.Source = self
                result.State = state
                result.Result = count
                properties = getNamedValueSet({'SendDispatchResult', result})
            any = uno.Any('[]com.sun.star.beans.NamedValue', properties)
            uno.invoke(listener, 'jobFinished', (self, any))
            print("MailSender._execute() 4")

        resource = 1027 if self._isCanceled() else 1028
        self._logger.logprb(INFO, self._cls, self._getMethod(), resource)

        self._thread = None
        self._notifyClosed()

    def _sendMails(self, database, jobs):
        count = 0
        mtd = '_sendMails'
        mailer = Mailer(self._ctx, self, database, self._logger, True, self._stop)
        for job in jobs:
            mail = None
            self._logger.logprb(INFO, self._cls, mtd, 1031, job)
            if self._checkCanceled(job, mtd):
                break
            try:
                batch, mail = mailer.getMail(job)
                if self._checkCanceled(job, mtd):
                    break
                mailer.sendMail(mail)
            except UnoException as e:
                print("MailSync._sendMails() 1 ERROR: %s" % traceback.format_exc())
                self._logger.logprb(SEVERE, self._cls, mtd, 1033, job, e.__class__.__name__, e.Message)
                database.setJobState(job, 2)
                if self._checkCanceled(job, mtd):
                    break
                else:
                    continue
            except Exception as e:
                print("MailSync._sendMails() 2 ERROR: %s" % traceback.format_exc())
                self._logger.logprb(SEVERE, self._cls, mtd, 1034, job, e.__class__.__name__, traceback.format_exc())
                database.setJobState(job, 2)
                if self._checkCanceled(job, mtd):
                    break
                else:
                    continue
            database.updateSpooler(batch, job, mail.ThreadId, mail.MessageId, mail.ForeignId, 1)
            self._logger.logprb(INFO, self._cls, mtd, 1035, job)
            count += 1
        mailer.dispose()
        return count

    def _checkCanceled(self, job, mtd):
        canceled = self._isCanceled()
        if canceled:
            self._logger.logprb(INFO, self._cls, mtd, 1032, job)
        return canceled

    def _notifyStarted(self):
        for listener in self._listeners:
            listener.started()

    def _notifyClosed(self):
        for listener in self._listeners:
            listener.closed()

    def _notifyTerminated(self):
        for listener in self._listeners:
            listener.terminated()

    def _error(self, e):
        for listener in self._listeners:
            listener.error(e)


g_ImplementationHelper.addImplementation(MailSender,                      # UNO object class
                                         g_ImplementationName,            # Implementation name
                                         g_ServiceNames)                  # List of implemented services

