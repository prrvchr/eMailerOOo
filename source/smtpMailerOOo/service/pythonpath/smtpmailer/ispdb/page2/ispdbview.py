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

import unohelper

from ...unotool import getContainerWindow

from ...configuration import g_extension

import traceback


class IspdbView(unohelper.Base):
    def __init__(self, ctx, manager, parent):
        self._window = getContainerWindow(ctx, parent, None, g_extension, 'IspdbPage2')

# IspdbView getter methods
    def getWindow(self):
        return self._window

# IspdbView setter methods
    def setPageLabel(self, text):
        self._getPageLabel().Text = text

    def updateProgress(self, value, message, style):
        self._getProgressBar().Value = value
        control = self._getProgressMessage()
        control.Text = message
        control.Model.FontWeight = style

# IspdbView private getter control methods
    def _getPageLabel(self):
        return self._window.getControl('Label1')

    def _getProgressBar(self):
        return self._window.getControl('ProgressBar1')

    def _getProgressMessage(self):
        return self._window.getControl('Label2')
