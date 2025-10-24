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

import uno

from com.sun.star.frame.DispatchResultState import SUCCESS

from .worker import Dispatcher

from ...dbtool import getObjectSequenceFromResult

from ...unotool import getTempFile

from ...helper import getMailSpooler

import traceback


class Mailer(Dispatcher):
    def __init__(self, ctx, cancel, progress, logger, jobs, notifier, source):
        super().__init__(ctx, cancel, progress, logger)
        self._cls = 'Mailer'
        self._jobs = jobs
        self._notifier = notifier
        folder = getTempFile(self._ctx).Uri
        struct = 'com.sun.star.frame.DispatchResultEvent'
        self._result = uno.createUnoStruct(struct, source, SUCCESS, '%s/Email.eml' % folder)
        self._resource = 800
        self.start()

    # Private getter methods
    def _getBatchs(self):
        spooler = getMailSpooler(self._ctx)
        result = spooler.getJobs(self._jobs)
        batchs = getObjectSequenceFromResult(result)
        result.close()
        spooler.dispose()
        return batchs

    def _isOffLine(self):
        return False

    def _notifyClosed(self, value=100):
        super()._notifyClosed(value)
        self._notifier.dispatchFinished(self._result)

