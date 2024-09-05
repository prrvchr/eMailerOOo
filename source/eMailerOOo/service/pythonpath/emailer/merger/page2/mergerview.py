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

from com.sun.star.view.SelectionType import MULTI

from ...unotool import createService
from ...unotool import getContainerWindow

from ...configuration import g_identifier

import traceback


class MergerView(unohelper.Base):
    def __init__(self, ctx, handler1, handler2, parent, tables, title1, title2, label):
        self._ctx = ctx
        self._window = getContainerWindow(ctx, parent, None, g_identifier, 'MergerPage2')
        tab1, tab2 = self._getTabPages(title1, title2, 'Tab1', 1)
        self._tab1 = getContainerWindow(ctx, tab1.getPeer(), handler1, g_identifier, 'MergerTab1')
        self._tab1.setVisible(True)
        self._tab2 = getContainerWindow(ctx, tab2.getPeer(), handler2, g_identifier, 'MergerTab2')
        self._tab2.setVisible(True)
        self._initTables(tables)
        self.setQueryLabel(label)

# MergerView getter methods
    def getWindow(self):
        return self._window

    def getGridWindows(self):
        return self._getGrid1Windows(), self._getGrid2Windows()

    def getTable(self):
        return self._getTable().getSelectedItem()

# MergerView setter methods
    def setTable(self, table):
        self._getTable().selectItem(table, True)

    def initTables(self, tables, table):
        control = self._getTable()
        control.Model.removeAllItems()
        control.Model.StringItemList = tables
        control.selectItem(table, True)

    def enableAdd(self, enabled):
        self._getAdd().Model.Enabled = enabled

    def enableRemove(self, enabled):
        self._getRemove().Model.Enabled = enabled

    def enableAddAll(self, enabled):
        self._getAddAll().Model.Enabled = enabled

    def enableRemoveAll(self, enabled):
        self._getRemoveAll().Model.Enabled = enabled

    def setQueryLabel(self, label):
        self._getQuery().Text = label

# MergerView private getter control methods
    def _getTable(self):
        return self._tab1.getControl('ListBox1')

    def _getAddAll(self):
        return self._tab1.getControl('CommandButton1')

    def _getAdd(self):
        return self._tab1.getControl('CommandButton2')

    def _getGrid1Windows(self):
        return self._tab1.getControl('FrameControl1')

    def _getRemoveAll(self):
        return self._tab2.getControl('CommandButton1')

    def _getRemove(self):
        return self._tab2.getControl('CommandButton2')

    def _getQuery(self):
        return self._tab2.getControl('Label1')

    def _getGrid2Windows(self):
        return self._tab2.getControl('FrameControl1')

# MergerView private methods
    def _getTabPages(self, title1, title2, name, i):
        service = 'com.sun.star.awt.tab.UnoControlTabPageContainerModel'
        model = self._window.Model.createInstance(service)
        model.PositionX = 0
        model.PositionY = 0
        model.Width = self._window.Model.Width
        model.Height = self._window.Model.Height
        self._window.Model.insertByName(name, model)
        tab = self._window.getControl(name)
        tab1 = self._getTabPage(model, tab, title1, 0)
        tab2 = self._getTabPage(model, tab, title2, 1)
        tab.ActiveTabPageID = i
        return tab1, tab2

    def _getTabPage(self, model, tab, title, i):
        page = model.createTabPage(i +1)
        page.Title = title
        index = model.getCount()
        model.insertByIndex(index, page)
        return tab.getControls()[i]

    def _initTables(self, tables):
        self._getTable().Model.StringItemList = tables
