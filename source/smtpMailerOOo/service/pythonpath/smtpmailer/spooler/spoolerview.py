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

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from ..unotool import getContainerWindow
from ..unotool import getDialog

from ..configuration import g_extension


class SpoolerView(unohelper.Base):
    def __init__(self, ctx, handler, handler1, handler2, parent, title, title1, title2):
        self._dialog = getDialog(ctx, g_extension, 'SpoolerDialog', handler, parent)
        rectangle = uno.createUnoStruct('com.sun.star.awt.Rectangle', 0, 0, 400, 175)
        tab1, tab2 = self._getTabPages(self._dialog, 'Tab1', title1, title2, rectangle, 1)
        parent = tab1.getPeer()
        self._tab1 = getContainerWindow(ctx, parent, handler1, g_extension, 'SpoolerTab1')
        self._tab1.setVisible(True)
        parent = tab2.getPeer()
        self._tab2 = getContainerWindow(ctx, parent, handler2, g_extension, 'SpoolerTab2')
        self._tab2.setVisible(True)
        self._dialog.setTitle(title)

# SpoolerView getter methods
    def execute(self):
        return self._dialog.execute()

    def getParent(self):
        return self._dialog.getPeer()

    def getGridWindow(self):
        return self._getGridWindow()

# SpoolerView setter methods
    def initView(self):
        self._enableButtonStartSpooler(True)
        self._enableButtonCancel(True)
        self._enableButtonClose(True)
        self._enableButtonAdd(True)

    def enableButtonRemove(self, enabled):
        self._getButtonRemove().Model.Enabled = enabled

    def setSpoolerState(self, label):
        self._getLabelState().Text = label

    def endDialog(self):
        self._dialog.endDialog(OK)

    def dispose(self):
        self._dialog.dispose()

    def updateLog(self, text, length):
        print("SpoolerView.updateLog()")
        control = self._getActivityLog()
        selection = uno.createUnoStruct('com.sun.star.awt.Selection')
        selection.Min = length
        selection.Max = length
        control.Text = text
        control.setSelection(selection)

# SpoolerView private setter methods
    def _enableButtonStartSpooler(self, enabled):
        self._getButtonStartSpooler().Model.Enabled = enabled

    def _enableButtonCancel(self, enabled):
        self._getButtonCancel().Model.Enabled = enabled

    def _enableButtonClose(self, enabled):
        self._getButtonClose().Model.Enabled = enabled

    def _enableButtonAdd(self, enabled):
        self._getButtonAdd().Model.Enabled = enabled

# SpoolerView private control methods
    def _getButtonStartSpooler(self):
        return self._dialog.getControl('CommandButton1')

    def _getButtonCancel(self):
        return self._dialog.getControl('CommandButton2')

    def _getButtonClose(self):
        return self._dialog.getControl('CommandButton3')

    def _getButtonAdd(self):
        return self._tab1.getControl('CommandButton1')

    def _getButtonRemove(self):
        return self._tab1.getControl('CommandButton2')

    def _getGridWindow(self):
        return self._tab1.getControl('FrameControl1')

    def _getLabelState(self):
        return self._dialog.getControl('Label2')

    def _getActivityLog(self):
        return self._tab2.getControl('TextField1')

# SpoolerView private methods
    def _getTabPages(self, page, name, title1, title2, rectangle, id):
        model = self._getTabModel(page, rectangle)
        page.Model.insertByName(name, model)
        tab = page.getControl(name)
        tab1 = self._getTabPage(tab, model, title1, 0)
        tab2 = self._getTabPage(tab, model, title2, 1)
        tab.ActiveTabPageID = id
        return tab1, tab2

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
