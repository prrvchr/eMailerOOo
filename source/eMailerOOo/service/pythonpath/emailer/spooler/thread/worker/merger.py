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

from com.sun.star.text.MailMergeType import FILE

from com.sun.star.sdb.CommandType import TABLE

from .type import Worker

from ....unotool import getMailMerge
from ....unotool import getMri

import traceback


class Merger(Worker):
    def __init__(self, ctx, cancel, progress, input, export, output=None):
        super().__init__(cancel, progress, input, output)
        self._export = export

    def run(self):
        task = self._input
        print("Merger.run() 1 name: %s - Title: %s" % (self.name, task.Name))

        if task.Merge:
            self._setProgressText("Merge file: %s for %s copies" % (task.Name, task.JobCount))
            self._setProgressValue(-10)
            print("Merger.run() 2 name: %s - Title: %s" % (self.name, task.Name))
        else:
            self._setProgressText("Export file: %s for %s copies" % (task.Name, task.JobCount))
            self._setProgressValue(-10)
        task.execute(self._cancel, self._progress)
        if task.Merge:
            self._setProgressText("End merging file: %s" % task.Name)
            self._setProgressValue(-10)
            #self._job.dispose()
        else:
            self._setProgressText("End exporting file: %s" % task.Name)
            self._setProgressValue(-10)
        #task.close()
        print("Merger.run() 3 name: %s - Title: %s" % (self.name, task.Name))

        if self._output:
            self._setProgressText("Queuing merged/exported file: %s" % task.Name)
            self._setProgressValue(-10)
            self._output.put(task)
        print("Merger.run() normal thread end")

