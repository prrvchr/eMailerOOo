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

from .type import Worker

import traceback


class Executor(Worker):
    def __init__(self, logger, resource, cancel, progress, input, output=None):
        super().__init__(cancel, progress, input, output)
        self._logger = logger
        self._resource = resource

    def run(self):
        task = self._input
        if task.Merge:
            msg = self._logger.resolveString(self._resource + 11, task.Name, task.JobCount)
        else:
            msg = self._logger.resolveString(self._resource + 12, task.Name, task.JobCount)
        self._setProgress(msg, -10)
        task.execute(self._cancel)
        if task.Merge:
            msg = self._logger.resolveString(self._resource + 13, task.Name)
        else:
            msg = self._logger.resolveString(self._resource + 14, task.Name)
        self._setProgress(msg, -10)

        if self._output:
            msg = self._logger.resolveString(self._resource + 15)
            self._setProgress(msg, -10)
            self._output.put(task)

