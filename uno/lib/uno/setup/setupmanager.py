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

from .setupmodel import SetupModel

from .setupview import SetupView

from .setuphandler import WindowHandler

import traceback


class SetupManager():
    def __init__(self, ctx, name):
        self._model = SetupModel(ctx, name)
        self._view = SetupView(ctx, WindowHandler(self), name, self._model.getTitle())

    def cancel(self):
        self._view.close()

    def next(self):
        step = self._view.getStep()
        if step == 1 or step == 7:
            self._checkRequirements()
        elif step == 3:
            self._installPackages()
        elif step == 6:
            if self._view.isDeregistered():
                self._model.deregisterJob()
            self._view.close()
        else:
            self._view.setPage(*self._model.getPage(step + 1))

    def setMaxProgress(self, value):
        self._view.setMaxProgress(value)

    def setProgress(self, text, progress):
        self._view.setProgess(text, progress)

    def _checkRequirements(self):
        self._view.setPage(*self._model.getPage(2))
        self._view.enableNext(False)
        result = self._model.checkRequirements(self.setMaxProgress, self.setProgress)
        self._view.enableNext(True)
        if result:
            self._view.setPage(*self._model.getPage(3))
            self._view.setResult(result)
        else:
            self._view.setPage(*self._model.getPage(6))

    def _installPackages(self):
        self._view.setPage(*self._model.getPage(4))
        self._view.enableNext(False)
        success, result = self._model.installPackages(self.setMaxProgress, self.setProgress)
        self._view.enableNext(True)
        if success:
            self._view.setPage(*self._model.getPage(5))
        else:
            self._view.setPage(*self._model.getPage(7))
        self._view.setResult(result)
