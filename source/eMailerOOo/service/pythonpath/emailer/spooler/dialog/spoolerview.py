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

from ...unotool import getContainerWindow
from ...unotool import getDialogPosSize
from ...unotool import getTopWindow
from ...unotool import getTopWindowPosition

from ...configuration import g_identifier
from ...configuration import g_spoolerframe

import traceback

class SpoolerView(unohelper.Base):
    def __init__(self, ctx, handler, listener, handler1, listener1, point, titles):
        rectangle = getDialogPosSize(ctx, g_identifier, 'SpoolerDialog', point)
        self._frame = getTopWindow(ctx, g_spoolerframe, rectangle)
        self._frame.setTitle(titles[0])
        parent = self._frame.getContainerWindow()
        self._dialog = getContainerWindow(ctx, parent, handler, g_identifier, 'SpoolerDialog')
        # XXX: setComponent is needed if we want a StatusIndicator at the bottom
        self._frame.setComponent(self._dialog, None)
        self._dialog.setVisible(True)
        self._frame.addCloseListener(listener)
        self._tab, tab1, tab2, tab3 = self._getTabPages('Tab1', 1, titles)
        self._tab1 = getContainerWindow(ctx, tab1.getPeer(), handler1, g_identifier, 'SpoolerTab1')
        self._setModelSize(self._tab.Model, self._tab1.Model)
        self._tab1.setVisible(True)
        self._tab2 = getContainerWindow(ctx, tab2.getPeer(), None, g_identifier, 'SpoolerTab2')
        self._tab2.setVisible(True)
        self._tab3 = getContainerWindow(ctx, tab3.getPeer(), None, g_identifier, 'SpoolerTab3')
        self._tab3.setVisible(True)
        self._tab.addTabPageContainerListener(listener1)

# SpoolerView getter methods
    def getDialogPosition(self):
        return getTopWindowPosition(self._frame.getContainerWindow())

    def getParent(self):
        return self._dialog.getPeer()

    def getGridWindow(self):
        return self._getGridWindow()

    def getActiveTab(self):
        return self._tab.ActiveTabPageID

# SpoolerView setter methods
    def close(self):
        self._frame.close(True)

    def initView(self):
        self._enableButtonStartSpooler(True)
        self._enableButtonClose(True)
        self._enableButtonAdd(True)

    def activateTab(self, tab):
        self._setDialogStep(int(tab == 1))

    def enableButtons(self, selected, sent, link):
        self._getButtonEmlView().Model.Enabled = selected
        self._getButtonClientView().Model.Enabled = selected and sent
        self._getButtonWebView().Model.Enabled = selected and sent and link
        self._getButtonResubmit().Model.Enabled = selected
        self._getButtonRemove().Model.Enabled = selected

    def disableButtons(self, add=False):
        self._getButtonEmlView().Model.Enabled = False
        self._getButtonClientView().Model.Enabled = False
        self._getButtonWebView().Model.Enabled = False
        self._getButtonResubmit().Model.Enabled = False
        if add:
            self._enableButtonAdd(False)

    def setSpoolerState(self, text, state, enabled):
        model = self._getButtonStartSpooler().Model
        model.State = state
        model.Enabled = enabled
        self._getButtonClose().Model.Enabled = enabled
        self._getLabelState().Text = text

    def updateLog1(self, text, length):
        control = self._getSpoolerLog()
        self._updateLog(control, text, length)

    def updateLog2(self, text, length):
        control = self._getMailServiceLog()
        self._updateLog(control, text, length)

# SpoolerView private setter methods
    def _updateLog(self, control, text, length):
        selection = uno.createUnoStruct('com.sun.star.awt.Selection')
        selection.Min = length
        selection.Max = length
        control.Text = text
        control.setSelection(selection)

    def _enableButtonStartSpooler(self, enabled):
        self._getButtonStartSpooler().Model.Enabled = enabled

    def _enableButtonClose(self, enabled):
        self._getButtonClose().Model.Enabled = enabled

    def _enableButtonAdd(self, enabled):
        self._getButtonAdd().Model.Enabled = enabled

    def _setDialogStep(self, step):
        self._dialog.Model.Step = step

# SpoolerView private control methods
    def _getButtonStartSpooler(self):
        return self._dialog.getControl('CommandButton1')

    def _getButtonClose(self):
        return self._dialog.getControl('CommandButton3')

    def _getButtonEmlView(self):
        return self._tab1.getControl('CommandButton1')

    def _getButtonClientView(self):
        return self._tab1.getControl('CommandButton2')

    def _getButtonWebView(self):
        return self._tab1.getControl('CommandButton3')

    def _getButtonResubmit(self):
        return self._tab1.getControl('CommandButton4')

    def _getButtonAdd(self):
        return self._tab1.getControl('CommandButton5')

    def _getButtonRemove(self):
        return self._tab1.getControl('CommandButton6')

    def _getGridWindow(self):
        return self._tab1.getControl('FrameControl1')

    def _getLabelState(self):
        return self._dialog.getControl('Label2')

    def _getSpoolerLog(self):
        return self._tab2.getControl('TextField1')

    def _getMailServiceLog(self):
        return self._tab3.getControl('TextField1')

# SpoolerView private methods
    def _getTabPages(self, name, tabid, titles):
        service = 'com.sun.star.awt.tab.UnoControlTabPageContainerModel'
        model = self._dialog.Model.createInstance(service)
        self._dialog.Model.insertByName(name, model)
        tab = self._dialog.getControl(name)
        tab1 = self._getTabPage(tab, model, titles[1], 0)
        tab2 = self._getTabPage(tab, model, titles[2], 1)
        tab3 = self._getTabPage(tab, model, titles[3], 2)
        tab.ActiveTabPageID = tabid
        return tab, tab1, tab2, tab3

    def _getTabPage(self, tab, model, title, tabid):
        page = model.createTabPage(tabid + 1)
        page.Title = title
        index = model.getCount()
        model.insertByIndex(index, page)
        return tab.getControls()[tabid]

    def _setModelSize(self, model, size):
        model.Width = size.Width
        model.Height = size.Height

