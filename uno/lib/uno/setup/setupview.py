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

from com.sun.star.awt import Size
from com.sun.star.awt.PosSize import POSSIZE

from com.sun.star.util.MeasureUnit import APPFONT

from ..unotool import getContainerWindow
from ..unotool import getToolKit
from ..unotool import getTopWindow

from ..configuration import g_identifier

import traceback


class SetupView():
    def __init__(self, ctx, handler, name, title):
        self._frame = getTopWindow(ctx, name)
        peer = self._frame.getContainerWindow()
        self._window = getContainerWindow(ctx, peer, handler, g_identifier, name)
        # XXX: setComponent is needed if we want a StatusIndicator at the bottom
        self._frame.setComponent(self._window, None)
        self._initWindow(ctx, title)

# SetupView getter methods
    def getWindow(self):
        return self._window

    def getDialogWindow(self):
        return self._frame.getContainerWindow()

    def isDeregistered(self):
        return bool(self._getCheck().State)

    def getStep(self):
        return self._window.Model.Step

# SetupView setter methods
    def close(self):
        self._frame.close(True)

    def setPage(self, step, header):
        self._window.Model.Step = step
        self._getPageHeader().Text = header

    def setMaxProgress(self, value):
        model = self._getProgressBar().Model
        model.ProgressValue = 0
        model.ProgressValueMax = value

    def setProgess(self, text, progress):
        # FIXME: To ensure the label is correctly updated, the progress bar must be updated last.
        self._getProgressText().Text = text
        self._getProgressBar().Model.ProgressValue = progress

    def setResult(self, text):
        self._getResult().Text = text

    def enableNext(self, enabled):
        self._getNextButton().Model.Enabled = enabled

# SetupView private methods
    def _initWindow(self, ctx, title):
        dialog = self._window.Model
        size = self._window.convertSizeToPixel(Size(dialog.Width, dialog.Height), APPFONT)
        x, y = self._getWindowPosition(ctx, size)
        self._frame.getContainerWindow().setPosSize(x, y, size.Width, size.Height, POSSIZE)
        self._frame.setTitle(title)
        # XXX: Visibility should be done after size adjustment
        self._window.setVisible(True)

    def _getWindowPosition(self, ctx, size):
        device = getToolKit(ctx).getWorkArea()
        x = device.X + ((device.Width - size.Width) // 2)
        y = device.Y + ((device.Height - size.Height) // 2)
        return x, y

# SetupView private control methods
    def _getProgressBar(self):
        return self._window.getControl('ProgressBar%s' % self.getStep())

    def _getProgressText(self):
        return self._window.getControl('Label%s' % self.getStep())

    def _getPageHeader(self):
        return self._window.getControl('Label1')

    def _getResult(self):
        return self._window.getControl('Label%s' % self.getStep())

    def _getCheck(self):
        return self._window.getControl('CheckBox1')

    def _getNextButton(self):
        return self._window.getControl('CommandButton2')

