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

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from ...unotool import getContainerWindow
from ...unotool import getDialog

from ...configuration import g_identifier


class SpoolerView(unohelper.Base):
    def __init__(self, ctx, handler, listener, handler1, parent, title, title1, title2, title3):
        self._dialog = getDialog(ctx, g_identifier, 'SpoolerDialog', handler, parent)
        rectangle = uno.createUnoStruct('com.sun.star.awt.Rectangle', 0, 0, 400, 175)
        self._tab, tab1, tab2, tab3 = self._getTabPages(self._dialog, 'Tab1', title1, title2, title3, rectangle, 1)
        self._tab1 = getContainerWindow(ctx, tab1.getPeer(), handler1, g_identifier, 'SpoolerTab1')
        self._tab1.setVisible(True)
        self._tab2 = getContainerWindow(ctx, tab2.getPeer(), None, g_identifier, 'SpoolerTab2')
        self._tab2.setVisible(True)
        self._tab3 = getContainerWindow(ctx, tab3.getPeer(), None, g_identifier, 'SpoolerTab3')
        self._tab3.setVisible(True)
        self._dialog.setTitle(title)
        self._tab.addTabPageContainerListener(listener)

# SpoolerView getter methods
    def execute(self):
        return self._dialog.execute()

    def getParent(self):
        return self._dialog.getPeer()

    def getGridWindow(self):
        return self._getGridWindow()

    def getActiveTab(self):
        return self._tab.ActiveTabPageID

# SpoolerView setter methods
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
        self._getButtonRemove().Model.Enabled = selected

    def disableButtons(self):
        self._getButtonEmlView().Model.Enabled = False
        self._getButtonClientView().Model.Enabled = False
        self._getButtonWebView().Model.Enabled = False

    def setSpoolerState(self, text, state, enabled):
        model = self._getButtonStartSpooler().Model
        model.State = state
        model.Enabled = enabled
        self._getButtonClose().Model.Enabled = enabled
        self._getLabelState().Text = text

    def endDialog(self):
        self._dialog.endDialog(OK)

    def dispose(self):
        self._dialog.dispose()

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

    def _getButtonAdd(self):
        return self._tab1.getControl('CommandButton4')

    def _getButtonRemove(self):
        return self._tab1.getControl('CommandButton5')

    def _getGridWindow(self):
        return self._tab1.getControl('FrameControl1')

    def _getLabelState(self):
        return self._dialog.getControl('Label2')

    def _getSpoolerLog(self):
        return self._tab2.getControl('TextField1')

    def _getMailServiceLog(self):
        return self._tab3.getControl('TextField1')

# SpoolerView private methods
    def _getTabPages(self, page, name, title1, title2, title3, rectangle, id):
        model = self._getTabModel(page, rectangle)
        page.Model.insertByName(name, model)
        tab = page.getControl(name)
        tab1 = self._getTabPage(tab, model, title1, 0)
        tab2 = self._getTabPage(tab, model, title2, 1)
        tab3 = self._getTabPage(tab, model, title3, 2)
        tab.ActiveTabPageID = id
        return tab, tab1, tab2, tab3

    def _getTabModel(self, page, rectangle):
        service = 'com.sun.star.awt.tab.UnoControlTabPageContainerModel'
        model = page.Model.createInstance(service)
        model.PositionX = rectangle.X
        model.PositionY = rectangle.Y
        model.Width = rectangle.Width
        model.Height = rectangle.Height
        return model

    def _getTabPage(self, tab, model, title, id):
        page = model.createTabPage(id +1)
        page.Title = title
        index = model.getCount()
        model.insertByIndex(index, page)
        return tab.getControls()[id]
