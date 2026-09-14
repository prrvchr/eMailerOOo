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

from ..unotool.layout import Layout

from ..unotool.linklistener import LinkListener

from ..unotool.sublayout import SubLayout


class OptionsView():
    def __init__(self, ctx, window):
        self._window = window
        self._link = LinkListener(ctx, self._getLabel())
        self._layout = SubLayout(Layout(window, size=Size(279, 290), **self._getLayouts()))
        window.addWindowListener(self._layout)

# OptionsView getter methods
    def getLayout(self):
        return self._layout

    def getFrame(self):
        return self._getFrame()

    def getSetup(self):
        return self._getSetup()

    def getWindow(self):
        return self._window

    def getData(self):
        timeout = int(self._getTimeout().Value)
        mode = self._getErrorMode().getSelectedItemPos()
        delay = self._getDelay().Value
        return timeout, mode, delay

# OptionsView setter methods
    def dispose(self):
        if self._layout:
            self._layout.dispose()
            self._layout = None
        if self._link:
            self._link.dispose()
            self._link = None

    def initView(self, timeout, mode, delay, state, text, url):
        self._getTimeout().Value = timeout
        self._getErrorMode().selectItemPos(mode, True)
        self._getDelay().Value = delay
        self.refreshView(state, text, url)

    def refreshView(self, state, text, url):
        self._setLabel(text, url)
        self._enableButtons(state)

# OptionsView private methods
    def _getLayouts(self):
        return {'horizontals': self._getHorizontals(),
                'tops': self._getTops()}

    def _getHorizontals(self):
        yield self._getFrame()
        yield self._getSetup()
        for index in range(1, 3):
            yield self._getFixedLine(index)
        yield self._getErrorMode()
        yield self._getLabel()

    def _getTops(self):
        yield self._getDelayLabel()
        for index in range(1, 3):
            yield self._getNumericField(index)
        for index in range(1, 5):
            yield self._getButton(index)

    def _setLabel(self, label, url):
        self._getLabel().Model.Label = label
        self._link.setUrl(url)

    def _enableButtons(self, state):
        enabled = state.extension and state.java and state.python
        self._getShowIspdb().Model.Enabled = enabled
        self._getShowDataBase().Model.Enabled = state.extension and state.java and state.database
        self._getShowSpooler().Model.Enabled = enabled
        self._getShowController().Model.Enabled = enabled

# OptionsView private control methods
    def _getFrame(self):
        return self._window.getControl('FrameControl1')

    def _getSetup(self):
        return self._window.getControl('FrameControl2')

    def _getFixedLine(self, index):
        return self._window.getControl(f'FixedLine{index}')

    def _getTimeout(self):
        return self._window.getControl('NumericField1')

    def _getDelay(self):
        return self._window.getControl('NumericField2')

    def _getShowIspdb(self):
        return self._window.getControl('CommandButton1')

    def _getShowDataBase(self):
        return self._window.getControl('CommandButton2')

    def _getShowSpooler(self):
        return self._window.getControl('CommandButton3')

    def _getShowController(self):
        return self._window.getControl('CommandButton4')

    def _getErrorMode(self):
        return self._window.getControl('ListBox1')

    def _getNumericField(self, index):
        return self._window.getControl(f'NumericField{index}')

    def _getButton(self, index):
        return self._window.getControl(f'CommandButton{index}')

    def _getDelayLabel(self):
        return self._window.getControl('Label5')

    def _getLabel(self):
        return self._window.getControl('Label8')
