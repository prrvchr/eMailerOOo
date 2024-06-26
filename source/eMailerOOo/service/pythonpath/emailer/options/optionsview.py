#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020 https://prrvchr.github.io                                     ║
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
    def __init__(self, dialog, exist, timeout, msg, state):
        self._dialog = dialog
        self.initControl(exist, timeout, msg, state)

# OptionsView getter methods
    def getTimeout(self):
        return int(self._getTimeout().Value)

# OptionsView setter methods
    def initControl(self, exist, timeout, msg, state):
        self._getDataBaseButton().Model.Enabled = exist
        self._getTimeout().Value = timeout
        self.setSpoolerStatus(msg, state)

    def setSpoolerStatus(self, msg, state):
        control = self._getSpoolerButton().Model
        control.State = state
        control.Enabled = True
        self._getSpoolerStatus().Text = msg

    def setSpoolerError(self, error):
        self._getSpoolerError().Text = error

    def clearSpoolerError(self):
        self._getSpoolerError().Text = ''

# OptionsView private control methods
    def _getTimeout(self):
        return self._dialog.getControl('NumericField1')

    def _getSpoolerStatus(self):
        return self._dialog.getControl('Label5')

    def _getSpoolerError(self):
        return self._dialog.getControl('Label7')

    def _getDataBaseButton(self):
        return self._dialog.getControl('CommandButton2')

    def _getSpoolerButton(self):
        return self._dialog.getControl('CommandButton3')

