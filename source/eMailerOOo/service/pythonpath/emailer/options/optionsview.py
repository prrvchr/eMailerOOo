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
import unohelper

import traceback


class OptionsView(unohelper.Base):
    def __init__(self, window):
        self._window = window

# OptionsView getter methods
    def getWindow(self):
        return self._window

    def getTimeout(self):
        return int(self._getTimeout().Value)

# OptionsView setter methods
    def dispose(self):
        self._window.dispose()

    def initView(self, restart, url, instrumented, exist, timeout):
        control = self._getWarning()
        control.URL = url
        self._setWarning(control, restart, instrumented)
        self.updateDataBase(exist)
        self._getTimeout().Value = timeout

    def setWarning(self, restart, instrumented):
        self._setWarning(self._getWarning(), restart, instrumented)

    def updateDataBase(self, exist):
        self._getDataBaseButton().Model.Enabled = exist

    def setRestart(self, enabled):
        self._getRestart().setVisible(enabled)

# OptionsView private methods
    def _setWarning(self, control, restart, instrumented):
        if restart:
            control.setVisible(False)
            self._getRestart().setVisible(True)
        else:
            self._getRestart().setVisible(False)
            control.setVisible(not instrumented)

# OptionsView private control methods
    def _getTimeout(self):
        return self._window.getControl('NumericField1')

    def _getDataBaseButton(self):
        return self._window.getControl('CommandButton2')

    def _getRestart(self):
        return self._window.getControl('Label5')

    def _getWarning(self):
        return self._window.getControl('Hyperlink1')

