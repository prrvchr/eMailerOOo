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
from ...unotool import getWindowPosition
from ...unotool import setWindowPosition

from ...configuration import g_identifier
from ...configuration import g_spoolerframe

import traceback


class SpoolerView():
    def __init__(self, ctx, handler, listener, handler1, listener1, point, title, titles):
        self._ctx = ctx
        self._tabName = 'Tab1'
        self._frame = getTopWindow(ctx, g_spoolerframe)
        window = self._frame.getContainerWindow()
        self._window = getContainerWindow(ctx, window, handler, g_identifier, 'SpoolerDialog')
        tab, tab1, tab2, tab3 = self._getTabPages(titles)
        self._tab1 = getContainerWindow(ctx, tab1, handler1, g_identifier, 'SpoolerTab1')
        self._setModelSize(tab.Model, self._tab1.Model)
        self._tab1.setVisible(True)
        self._tab2 = getContainerWindow(ctx, tab2, None, g_identifier, 'SpoolerTab2')
        self._tab2.setVisible(True)
        self._tab3 = getContainerWindow(ctx, tab3, None, g_identifier, 'SpoolerTab3')
        self._tab3.setVisible(True)
        # XXX: setComponent is needed if we want a StatusIndicator at the bottom
        self._frame.setComponent(self._window, None)
        self._frame.addCloseListener(listener)
        tab.addTabPageContainerListener(listener1)
        setWindowPosition(ctx, window, self._window, point)
        self._frame.setTitle(title)
        self._window.setVisible(True)
        window.setVisible(True)
        self._offset = self._getOffset(10)
        print("SpoolerView.__init__()")

# SpoolerView getter methods
    def getDialogPosition(self):
        return getWindowPosition(self._frame.getContainerWindow())

    def getFrame(self):
        return self._frame

    def getWindow(self):
        return self._window

    def getGridWindow(self):
        return self._getGridWindow()

    def getActiveTab(self):
        return self._getTab().ActiveTabPageID

    def initTab2(self):
        return self._tab2 is None

    def initTab3(self):
        return self._tab3 is None

    def getIndicator(self):
        indicator = self._frame.createStatusIndicator()
        self._setWindowHeight()
        return indicator

    def resetIndicator(self):
        self._setWindowHeight(-1)

# SpoolerView setter methods
    def dispose(self):
        self._frame.close(True)

    def initView(self):
        self._getButtonSpooler().Model.Enabled = True
        self._getButtonClose().Model.Enabled = True
        self._getButtonAdd().Model.Enabled = True
        self._getLabelState().Text = ''

    def setTab2(self):
        peer = self._getTab().getControls()[1].getPeer()
        self._tab2 = getContainerWindow(self._ctx, peer, None, g_identifier, 'SpoolerTab2')
        self._tab2.setVisible(True)

    def setTab3(self):
        peer = self._getTab().getControls()[2].getPeer()
        self._tab3 = getContainerWindow(self._ctx, peer, None, g_identifier, 'SpoolerTab3')
        self._tab3.setVisible(True)

    def activateTab(self, ctx, tab):
        self._setDialogStep(int(tab == 1))

    def enableButtons(self, selected, sent, link):
        self._getButtonEmlView().Model.Enabled = selected
        self._getButtonClientView().Model.Enabled = selected and sent
        self._getButtonWebView().Model.Enabled = selected and sent and link
        self._getButtonResubmit().Model.Enabled = selected
        self._getButtonAdd().Model.Enabled = True
        self._getButtonRemove().Model.Enabled = selected
        self._getButtonSpooler().Model.Enabled = True

    def disableButtons(self):
        self._getButtonEmlView().Model.Enabled = False
        self._getButtonClientView().Model.Enabled = False
        self._getButtonWebView().Model.Enabled = False
        self._getButtonResubmit().Model.Enabled = False
        self._getButtonAdd().Model.Enabled = False
        self._getButtonRemove().Model.Enabled = False
        self._getButtonSpooler().Model.Enabled = False

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

    def _setDialogStep(self, step):
        self._window.Model.Step = step

# SpoolerView private getter methods
    def _getOffset(self, offset):
        point = uno.createUnoStruct('com.sun.star.awt.Point', 0, offset)
        return self._window.convertPointToPixel(point, APPFONT).Y

    def _setWindowHeight(self, factor=1):
        window = self._frame.getContainerWindow()
        size = window.getPosSize()
        offset = self._offset * factor
        window.setPosSize(0, 0, size.Width, size.Height + offset, SIZE)

# SpoolerView private control methods
    def _getButtonSpooler(self):
        return self._window.getControl('CommandButton1')

    def _getButtonClose(self):
        return self._window.getControl('CommandButton3')

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
        return self._window.getControl('Label1')

    def _getSpoolerLog(self):
        return self._tab2.getControl('TextField1')

    def _getMailServiceLog(self):
        return self._tab3.getControl('TextField1')

# SpoolerView private methods
    def _getTabPages(self, titles):
        service = 'com.sun.star.awt.tab.UnoControlTabPageContainerModel'
        model = self._window.Model.createInstance(service)
        self._window.Model.insertByName(self._tabName, model)
        tab = self._getTab()
        tab1 = self._getTabPage(tab, model, titles[0], 0)
        tab2 = self._getTabPage(tab, model, titles[1], 1)
        tab3 = self._getTabPage(tab, model, titles[2], 2)
        tab.ActiveTabPageID = 1
        return tab, tab1, tab2, tab3

    def _getTabPeer(self, tab, index):
        return tab.getControls()[0].getPeer()

    def _getTabPage(self, tab, model, title, tabid):
        page = model.createTabPage(tabid + 1)
        page.Title = title
        index = model.getCount()
        model.insertByIndex(index, page)
        return tab.getControls()[tabid].getPeer()

    def _setModelSize(self, model, size):
        model.Width = size.Width
        model.Height = size.Height

    def _getTab(self):
        return self._window.getControl(self._tabName)

