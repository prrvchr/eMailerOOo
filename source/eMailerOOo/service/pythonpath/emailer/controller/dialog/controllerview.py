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

from com.sun.star.awt.PosSize import SIZE

from com.sun.star.util.MeasureUnit import APPFONT

from ...unotool import getContainerWindow
from ...unotool import getTopWindow
from ...unotool import getWindowHeight
from ...unotool import getWindowPosition
from ...unotool import setWindowHeight
from ...unotool import setWindowPosition

from ...configuration import g_identifier
from ...configuration import g_controllerframe

import traceback

class ControllerView():
    def __init__(self, ctx, handler, listener, point, title):
        self._ctx = ctx
        self._frame = getTopWindow(ctx, g_controllerframe)
        window = self._frame.getContainerWindow()
        self._window = getContainerWindow(ctx, window, handler, g_identifier, 'SpoolerController')
        # XXX: setComponent is needed if we want a StatusIndicator at the bottom
        self._frame.setComponent(self._window, None)
        self._frame.addCloseListener(listener)
        setWindowPosition(ctx, window, self._window, point)
        self._frame.setTitle(title)
        self._window.setVisible(True)
        window.setVisible(True)
        self._point = getWindowHeight(self._window, 10)

# ControllerView getter methods
    def getWindowPosition(self):
        return getWindowPosition(self._frame.getContainerWindow())

    def getWindow(self):
        return self._frame.getContainerWindow()

    def getIndicator(self):
        indicator = self._frame.createStatusIndicator()
        setWindowHeight(self._frame, self._point.Y)
        return indicator

    def resetIndicator(self):
        setWindowHeight(self._frame, self._point.X)

# ControllerView setter methods
    def setVisible(self, enabled):
        self._frame.getContainerWindow().setVisible(enabled)
        self._window.setVisible(enabled)

    def dispose(self):
        self._frame.close(True)

    def enableButtons(self, enabled):
        self._getToggleButton().Model.Enabled = enabled
        self._getSpoolerButton().Model.Enabled = enabled
        self._getCloseButton().Model.Enabled = enabled

    def enableToggleButon(self, enabled):
        self._getToggleButton().Model.Enabled = enabled

    def setLabelState(self, text):
        self._getLabelState().Text = text

    def setToggleState(self, text):
        self._getToggleButton().Model.Label = text

    def setToggleButton(self, state):
        self._getToggleButton().Model.State = state

    def getToggleButton(self):
        return self._getToggleButton().Model.State

# ControllerView private control methods
    def _getLabelState(self):
        return self._window.getControl('Label1')

    def _getToggleButton(self):
        return self._window.getControl('CommandButton1')

    def _getSpoolerButton(self):
        return self._window.getControl('CommandButton2')

    def _getCloseButton(self):
        return self._window.getControl('CommandButton3')

