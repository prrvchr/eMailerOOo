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

import unohelper

from com.sun.star.awt import Size

from ...unotool import Layout
from ...unotool import SubLayout

from ...unotool import getContainerWindow

from ...configuration import g_identifier


class MergerView(unohelper.Base):
    def __init__(self, ctx, handler1, handler2, parent, tables, title1, title2, label):
        self._ctx = ctx
        self._tabName = 'Tab1'
        self._window = getContainerWindow(ctx, parent, None, g_identifier, 'MergerPage2')
        tab1, tab2 = self._getTabPages(title1, title2, 1)
        self._tab1 = getContainerWindow(ctx, tab1.getPeer(), handler1, g_identifier, 'MergerTab1')
        self._size1 = Size(self._getGrid1Window().Model.Width, self._getGrid1Window().Model.Height)
        self._tab1.setVisible(True)
        self._tab2 = getContainerWindow(ctx, tab2.getPeer(), handler2, g_identifier, 'MergerTab2')
        self._size2 = Size(self._getGrid2Window().Model.Width, self._getGrid2Window().Model.Height)
        self._tab2.setVisible(True)
        self._initTables(tables)
        self.setQueryLabel(label)
        self._layout = SubLayout(Layout(self._window, controls=(self._getTabFrame(), self._getTab())))
        self._layout.addLayout(Layout(self._tab1, **self._getTab1Layouts()))
        self._layout.addLayout(Layout(self._tab2, **self._getTab2Layouts()))
        print("MergerView.__init__() 2 parent: %s" % parent)

# MergerView getter methods
    def getWindow(self):
        return self._window

    def getLayout(self):
        return self._layout

    def getGridData(self):
        return self._getGrid1Window(), self._size1, self._getGrid2Window(), self._size2

    def getGridWindows(self):
        return self._getGrid1Window(), self._getGrid2Window()

    def getGridSizes(self):
        return self._size1, self._size2

    def getTab1(self):
        return self._tab1

    def getTab2(self):
        return self._tab2

    def getGrid1Window(self):
        return self._getGrid1Window()

    def getGrid2Window(self):
        return self._getGrid2Window()

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

    def enableAddresTable(self, enabled):
        self._getTable().Model.Enabled = enabled

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
    def _getTabFrame(self):
        return self._window.getControl('FrameControl1')

    def _getTable(self):
        return self._tab1.getControl('ListBox1')

    def _getAddAll(self):
        return self._tab1.getControl('CommandButton1')

    def _getAdd(self):
        return self._tab1.getControl('CommandButton2')

    def _getGrid1Window(self):
        return self._tab1.getControl('FrameControl1')

    def _getRemoveAll(self):
        return self._tab2.getControl('CommandButton1')

    def _getRemove(self):
        return self._tab2.getControl('CommandButton2')

    def _getQuery(self):
        return self._tab2.getControl('Label1')

    def _getGrid2Window(self):
        return self._tab2.getControl('FrameControl1')

# MergerView private methods
    def _getTab1Layouts(self):
        return {'horizontals': (self._getTable(), ),
                'tops':        (self._getAddAll(), self._getAdd()),
                'controls':    (self._getGrid1Window(), )}

    def _getTab2Layouts(self):
        return {'tops':     (self._getRemoveAll(), self._getRemove()),
                'controls': (self._getGrid2Window(), )}

    def _getTabPages(self, title1, title2, i):
        service = 'com.sun.star.awt.tab.UnoControlTabPageContainerModel'
        model = self._window.Model.createInstance(service)
        frame = self._getTabFrame().Model
        model.PositionX = frame.PositionX
        model.PositionY = frame.PositionY
        model.Width = frame.Width
        model.Height = frame.Height
        self._window.Model.insertByName(self._tabName, model)
        tab = self._getTab()
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

    def _getTab(self):
        return self._window.getControl(self._tabName)
