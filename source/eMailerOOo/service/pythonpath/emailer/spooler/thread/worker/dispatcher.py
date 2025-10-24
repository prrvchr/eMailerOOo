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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .composer import Composer

from .executor import Executor

from .watchdog import WatchDog

from ..task import Mail

from ..worker import Master

from ...export import PdfExport

from ....unotool import getSimpleFile
from ....unotool import getUriFactory

from time import sleep
from queue import Queue
import traceback


class Dispatcher(Master):
    def __init__(self, ctx, cancel, progress, logger):
        super().__init__(cancel, progress)
        self._ctx = ctx
        self._logger = logger
        self._sf = getSimpleFile(ctx)
        self._uf = getUriFactory(ctx)
        self._export = PdfExport(ctx)
        self._output = Queue()
        self._queue = None
        self._result = None
        self._datasources = {}
        self._wait = 1.0

    def run(self):
        mtd = 'run'
        sleep(self._wait)
        if self._isOffLine():
            self._notifyStarted(mtd)
            sleep(self._wait)
            msg = self._logger.resolveString(self._resource + 2)
            self._setProgress(msg, 50)
            self._logger.logp(INFO, self._cls, mtd, msg)
            sleep(self._wait)
            self._notifyClosed()
        else:
            batchs = self._getBatchs()
            if batchs:
                nbatchs, njobs, ntasks, progress = self._getBatchCount(batchs)
                self._notifyStarted(mtd, progress)
                msg = self._logger.resolveString(self._resource + 4)
                self._setProgress(msg, -10)
                self._logger.logp(INFO, self._cls, mtd, msg)
                composer = self._getComposer()
                msg = self._logger.resolveString(self._resource + 5, njobs)
                self._setProgress(msg, -10)
                self._logger.logp(INFO, self._cls, mtd, msg)
                watchdog = WatchDog(composer)
                for batch in batchs:
                    if self.isCanceled():
                        break
                    mail = self._getMail(batch)
                    self._setProgressValue(-10)
                    if mail.hasError():
                        mail.close()
                        if mail.hasException():
                            msg = self._logger.resolveString(self._resource + 6, mail.getException())
                            self._logger.logp(SEVERE, self._cls, mtd, msg)
                        elif not mail.hasFile():
                            msg = self._logger.resolveString(self._resource + 7, mail.Document)
                            self._logger.logp(SEVERE, self._cls, mtd, msg)
                        else:
                            msg = self._logger.resolveString(self._resource + 8, mail.getMissingAttachments())
                            self._logger.logp(SEVERE, self._cls, mtd, msg)
                    else:
                        executor = self._getExecutor(mail)
                        if watchdog.startExecutor(executor, mail):
                            composer.addTasks()
                        else:
                            # XXX: If task has been aborted then we need to close all resources.
                            mail.close()
                # XXX: We need to start the composer to make sure that any initiated merge is
                # XXX: completed so that we can cleanly close all resources from the DataSource.
                composer.start()
                composer.join()
                self._closeDataSource()
                self._notifyClosed(progress)

            else:
                self._notifyStarted(mtd)
                sleep(self._wait)
                msg = self._logger.resolveString(self._resource + 3)
                self._setProgress(msg, 50)
                self._logger.logp(INFO, self._cls, mtd, msg)
                sleep(self._wait)
                self._notifyClosed()

    # Private getter methods
    def _getBatchs(self):
        raise NotImplementedError('Need to be implemented!')

    def _isOffLine(self):
        raise NotImplementedError('Need to be implemented!')

    def _getBatchCount(self, batchs):
        jobs = tasks = 0
        # XXX: Progress in: Dispatcher
        progress = 20
        for batch in batchs:
            job, task, progress = self._getBatchProgress(batch, progress)
            jobs += job
            tasks += task
        return len(batchs), jobs, tasks, progress

    def _getBatchProgress(self, batch, progress):
        job = len(batch.Jobs)
        task = len(batch.Attachments)
        # XXX: Progress in: Dispatcher + Executor  + Composer
        progress += (10) + (30) +  (10 + (20 * job))
        if self._output:
            # XXX: Progress in: Sender if used
            progress += (20 * job)
        return job, task, progress

    def _getComposer(self):
        return Composer(self._ctx, self._cls, self._logger, self._resource, self._cancel,
                        self._progress, self._sf, self._output, self._queue, self._result)

    def _getExecutor(self, mail):
        return Executor(self._logger, self._resource,
                        self._cancel, self._progress, mail, self._output)

    def _getMail(self, batch):
        mail = Mail(self._ctx, self._cancel, self._getResolver, self._sf, self._uf, self._export, batch)
        # XXX: If we want to avoid a memory dump when exiting LibreOffice,
        # XXX: it is imperative to dispose all used DataSources.
        if mail.hasDataSource():
            datasource = mail.getDataSource()
            self._datasources[datasource.Name] = datasource
        return mail

    def _getMethod(self):
        return 'terminate' if self.isCanceled() else 'close'

    # Private setter methods
    def _notifyStarted(self, mtd, range=100):
        msg = self._logger.resolveString(self._resource + 1)
        self._startProgress(msg, range)
        self._logger.logp(INFO, self._cls, mtd, msg)

    def _notifyClosed(self, value=100):
        resource = self._resource + 41 if self.isCanceled() else self._resource + 42
        msg = self._logger.resolveString(resource)
        self._setProgress(msg, value)
        self._logger.logp(INFO, self._cls, self._getMethod(), msg)
        sleep(self._wait)
        self._endProgress()

    def _getResolver(self, code, *args):
        return self._logger.resolveString(code, *args)

    def _closeDataSource(self):
        for datasource in self._datasources.values():
            datasource.dispose()

