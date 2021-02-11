#!
# -*- coding: utf_8 -*-

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

from unolib import getDialog
from unolib import getContainerWindow

from .spoolerhandler import DialogHandler
from .spoolerhandler import Page1Handler
from .spoolerhandler import Page2Handler
from .spoolerhandler import GridHandler

from smtpserver import g_extension


class SpoolerView(unohelper.Base):
    def __init__(self, ctx, manager, parent):
        handler = DialogHandler(manager)
        self._dialog = getDialog(ctx, g_extension, 'SpoolerDialog', handler, parent)
        point = uno.createUnoStruct('com.sun.star.awt.Point', 0, 0)
        size = uno.createUnoStruct('com.sun.star.awt.Size', 400, 180)
        title1 = manager.Model.getTabPageTitle(1)
        title2 = manager.Model.getTabPageTitle(2)
        page1, page2 = self._getTabPages('Tab1', point, size, title1, title2, 1)
        parent = page1.getPeer()
        handler = Page1Handler(manager)
        self._page1 = getContainerWindow(ctx, parent, handler, g_extension, 'SpoolerPage1')
        self._page1.setVisible(True)
        parent = page2.getPeer()
        handler = Page2Handler(manager)
        self._page2 = getContainerWindow(ctx, parent, handler, g_extension, 'SpoolerPage2')
        self._page2.setVisible(True)
        point = uno.createUnoStruct('com.sun.star.awt.Point', 4, 25)
        size = uno.createUnoStruct('com.sun.star.awt.Size', 390, 130)
        model = manager.Model.getGridDataModel(size.Width)
        grid = self._getGrid(model, 'Grid1', point, size)
        handler = GridHandler(manager)
        grid.addSelectionListener(handler)
        self._setTitle(manager.Model)

# SpoolerView setter methods
    def enableButtonRemove(self, enabled):
        self._getButtonRemove().Model.Enabled = enabled

    def setSpoolerState(self, label):
        self._getLabelState().Text = label

    def setColumnsList(self, columns, enumeration):
        control = self.getColumnsList()
        control.Model.StringItemList = columns
        while enumeration.hasMoreElements():
            column = enumeration.nextElement()
            index = columns.index(column.Name)
            control.selectItemPos(index, True)

    def setOrdersList(self, columns, enumeration):
        control = self.getOrdersList()
        control.Model.StringItemList = columns
        while enumeration.hasMoreElements():
            column = enumeration.nextElement()
            index = columns.index(column.Name)
            control.selectItemPos(index, True)

    def dispose(self):
        self._dialog.dispose()
        self._page1.dispose()
        self._page2.dispose()
        self._dialog = None
        self._page1 = None
        self._page2 = None

# SpoolerView getter methods
    def execute(self):
        return self._dialog.execute()

    def getParent(self):
        return self._dialog.getPeer()

# SpoolerView private setter methods
    def _setTitle(self, model):
        title = model.getDialogTitle()
        self._dialog.setTitle(title)

# SpoolerView private getter methods
    def _getTabPageTitle(self, model, id):
        return model.getTabPageTitle(id)

# SpoolerView StringRessoure methods
    def _getGridColumnResource(self, column):
        return 'SpoolerDialog.Grid1.Column.%s' % column



    def getStateResource(self, state):
        return 'SpoolerDialog.Label2.Label.%s' % state

# SpoolerView private control methods
    def _getButtonRemove(self):
        return self._dialog.getControl('CommandButton2')

    def _getLabelState(self):
        return self._dialog.getControl('Label2')

    def getColumnsList(self):
        return self._page1.getControl('ListBox1')

    def getOrdersList(self):
        return self._page1.getControl('ListBox2')

# SpoolerView private methods
    def _getTabPages(self, name, point, size, title1, title2, id):
        service = 'com.sun.star.awt.tab.UnoControlTabPageContainerModel'
        model = self._dialog.Model.createInstance(service)
        model.PositionX = point.X
        model.PositionY = point.Y
        model.Width = size.Width
        model.Height = size.Height
        self._dialog.Model.insertByName(name, model)
        tab = self._dialog.getControl(name)
        page1 = self._getTabPage(tab, model, title1, 0)
        page2 = self._getTabPage(tab, model, title2, 1)
        tab.ActiveTabPageID = id
        return page1, page2

    def _getTabPage(self, tab, model, title, id):
        page = model.createTabPage(id + 1)
        page.Title = title
        index = model.getCount()
        model.insertByIndex(index, page)
        return tab.getControls()[id]

    def _getGrid(self, model, name, point, size):
        grid = self._getGridModel(model, name, point, size)
        self._page1.Model.insertByName(name, grid)
        return self._page1.getControl(name)

    def _getGridModel(self, model, name, point, size):
        service = 'com.sun.star.awt.grid.UnoControlGridModel'
        grid = self._page1.Model.createInstance(service)
        grid.Name = name
        grid.PositionX = point.X
        grid.PositionY = point.Y
        grid.Height = size.Height
        grid.Width = size.Width
        grid.GridDataModel = model
        grid.ColumnModel = model.ColumnModel
        grid.SelectionModel = MULTI
        #grid.ShowRowHeader = True
        grid.BackgroundColor = 16777215
        return grid
