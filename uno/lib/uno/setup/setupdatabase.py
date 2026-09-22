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

from com.sun.star.awt import XCallback

from ..dbinit import DataBaseSetup

from ..unotool import getCallBack
from ..unotool import getSimpleFile

from ..dbinit import getDataBaseConnection
from ..dbinit import getDataBaseUrl

from threading import Timer


class SetupDataBase(unohelper.Base,
                    XCallback):
    def __init__(self, ctx, progress, callback, user='', pwd=''):
        self._ctx = ctx
        self._database = DataBaseSetup(ctx)
        self._progress = progress
        self._callback = callback
        self._step = 0
        self._asyncCall = getCallBack(ctx)

    def start(self):
        if getSimpleFile(self._ctx).exists(self._database.path):
            Timer(0.1, self._callback, args=(True, False)).start()
        else:
            if self._progress:
                self._progress.start(self._database.label, self._database.total)
            Timer(0.1, self._call).start()

    def notify(self, data):
        try:
            resource, call = next(self._database.steps)
            self._step += 1
            if self._progress:
                self._progress.setText(self._database.resolver.resolveString(resource))
                self._progress.setValue(self._step)
            call()

            Timer(0.1, self._call).start()

        except StopIteration:
            if self._progress:
                self._progress.end()
            self._callback(True, True)
        except Exception as e:
            if self._database.statement:
                try: self._database.statement.close()
                except: pass
            if self._database.connection:
                try: self._database.connection.close()
                except: pass
            if self._progress:
                self._progress.end()
            self._callback(False, True)

    def _call(self):
       self._asyncCall.addCallback(self, None)


class ProgressControl():
    def __init__(self, progress, label):
        self._progress = progress
        self._label = label

    def start(self, text, value):
        if self._progress:
            self._progress.ProgressValue = 0
            self._progress.ProgressValueMax = int(value)
            
        if self._label:
            self._label.Text = text

    def setValue(self, value):
        if self._progress:
            self._progress.ProgressValue = int(value)

    def setText(self, text):
        if self._label:
            self._label.Text = text

    def end(self):
        if self._progress:
            self._progress.ProgressValue = 0
        if self._label:
            self._label.Text = ''

