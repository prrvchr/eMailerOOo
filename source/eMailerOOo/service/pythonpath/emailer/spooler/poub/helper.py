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

from ..unotool import getCallBack

import traceback


class AsyncCall(unohelper.Base,
                XCallback):
    def __init__(self, ctx, progress, cancel):
        self._ctx = ctx
        self._asyncCall = getCallBack(ctx)
        self._progress = progress
        self._cancel = cancel

# com.sun.star.awt.XCallback
    def notify(self, data):
        raise NotImplementedError('Need to be implemented!')

    def finalize(self):
        raise NotImplementedError('Need to be implemented!')

    def start(self):
        self._asyncCall.addCallback(self, None)

    def cancel(self):
        self._cancel.set()

    def isCanceled(self):
        return self._cancel.isSet()

    def _startProgress(self, text, value):
        self._progress.start(text, value)

    def _setProgress(self, text, value):
        self._progress.setValue(value)
        self._progress.setText(text)

    def _setProgressText(self, text):
        self._progress.setText(text)

    def _setProgressValue(self, value):
        self._progress.setValue(value)

    def _endProgress(self):
        self._progress.end()

