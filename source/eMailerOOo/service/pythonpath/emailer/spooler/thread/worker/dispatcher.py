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

from .merger import Merger

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
        self._url = None
        self._datasources = {}
        self._wait = 1.0

    def run(self):
        mtd = 'run'
        sleep(self._wait)
        if self._isOffLine():
            self._notifyStarted(mtd)
            sleep(self._wait)
            msg = self._logger.resolveString(self._res + 2)
            self._setProgress(msg, 50)
            self._logger.logp(INFO, self._cls, mtd, msg)
            sleep(self._wait)
            self._notifyClosed()
        else:
            batchs = self._getBatchs()
            if batchs:
                nbatchs, njobs, ntasks, progress = self._getBatchCount(batchs)
                self._notifyStarted(mtd, progress)
                msg = self._logger.resolveString(self._res + 3)
                self._setProgress(msg, -10)
                self._logger.logp(INFO, self._cls, mtd, msg)
                composer = self._getComposer()
                msg = self._logger.resolveString(self._res + 4, njobs)
                self._setProgress(msg, -10)
                self._logger.logp(INFO, self._cls, mtd, msg)
                i = 1
                total = 0
                watchdog = WatchDog(composer)
                for batch in batchs:
                    job = len(batch.Jobs)
                    total += job
                    if self.isCanceled():
                        break
                    msg = self._logger.resolveString(self._res + 5, i, nbatchs)
                    self._setProgress(msg, -10)
                    self._logger.logp(INFO, self._cls, mtd, msg)
                    print("Dispatcher.run() 1 batch: %s" % (batch.Batch, ))
                    print("Dispatcher.run() 2 jobs: %s" % (batch.Jobs, ))
                    print("Dispatcher.run() 3 predicates: %s" % (batch.Predicates, ))
                    print("Dispatcher.run() 4 identifiers: %s" % (batch.Identifiers, ))
                    print("Dispatcher.run() 5 adresses: %s" % (batch.Addresses, ))
                    print("Dispatcher.run() 6 attachments: %s" % (batch.Attachments, ))
                    mail = self._getMail(batch)
                    if mail.hasFile() and mail.hasAttachments():
                        merger = self._getMerger(mail)
                        if watchdog.startJob(merger, mail):
                            composer.addTasks()
                            i += 1
                    elif not mail.hasFile():
                        print("Dispatcher.run() Missing file: %s" % mail.Document)
                    elif not mail.hasAttachments():
                        print("Dispatcher.run() Missing file: %s" % mail.getMissingAttachments())
                # XXX: We need to start the composer to make sure that any initiated merge is
                # XXX: completed so that we can cleanly close all resources from the DataSource.
                composer.start()
                composer.join()
                print("Dispatcher.run() normal thread end")
                self._closeDataSource()
                self._notifyClosed(progress)

            else:
                self._notifyStarted(mtd)
                sleep(self._wait)
                msg = self._logger.resolveString(self._res + 6)
                print("Dispatcher.run() 9 msg: %s" % msg)
                self._setProgress(msg, 50)
                self._logger.logp(INFO, self._cls, mtd, msg)
                sleep(self._wait)
                self._notifyClosed()
        print("Dispatcher.run() normal thread end")

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
        # XXX: Progress in: Dispatcher + Executor + Merger + Assembler + Composer
        progress += (10) + (30 + 20 * task) + (30 * task) +  (20 + 20 * task) + (10 + 20 * job)
        if self._output:
            # XXX: Progress in: Sender if used
            progress += (30 * job)
        return job, task, progress

    def _getComposer(self):
        return Composer(self._ctx, self._cancel, self._progress,
                        self._logger, self._sf, self._output, self._queue, self._url)

    def _getMerger(self, mail):
        return Merger(self._ctx, self._cancel, self._progress, mail, self._export, self._output)

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
        msg = self._logger.resolveString(self._res + 1)
        self._startProgress(msg, range)
        self._logger.logp(INFO, self._cls, mtd, msg)

    def _notifyClosed(self, value=100):
        resource = self._res + 7 if self.isCanceled() else self._res + 8
        msg = self._logger.resolveString(resource)
        print("Dispatcher.run() 2 msg: %s" % msg)
        self._setProgress(msg, value)
        self._logger.logp(INFO, self._cls, self._getMethod(), msg)
        sleep(self._wait)
        self._endProgress()

    def _getResolver(self, code, *args):
        return self._logger.resolveString(code, *args)

    def _closeDataSource(self):
        for datasource in self._datasources.values():
            datasource.dispose()

