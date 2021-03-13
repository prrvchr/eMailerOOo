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
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK


from unolib import getDialog
from unolib import getContainerWindow

from .spoolerhandler import DialogHandler
from .spoolerhandler import Tab1Handler
from .spoolerhandler import Tab2Handler
from .spoolerhandler import GridHandler

from smtpserver import g_extension


class SpoolerView(unohelper.Base):
    def __init__(self, ctx, manager, parent):
        handler = DialogHandler(manager)
        self._dialog = getDialog(ctx, g_extension, 'SpoolerDialog', handler, parent)
        rectangle = uno.createUnoStruct('com.sun.star.awt.Rectangle', 0, 0, 400, 180)
        title1 = manager.Model.getTabPageTitle(1)
        title2 = manager.Model.getTabPageTitle(2)
        tab1, tab2 = self._getTabPages(self._dialog, 'Tab1', title1, title2, rectangle, 1)
        parent = tab1.getPeer()
        handler = Tab1Handler(manager)
        self._tab1 = getContainerWindow(ctx, parent, handler, g_extension, 'SpoolerTab1')
        self._tab1.setVisible(True)
        parent = tab2.getPeer()
        handler = Tab2Handler(manager)
        self._tab2 = getContainerWindow(ctx, parent, handler, g_extension, 'SpoolerTab2')
        self._tab2.setVisible(True)
        title = manager.Model.getDialogTitle()
        self._dialog.setTitle(title)

# SpoolerView setter methods
    def initGrid(self, manager, titles):
        rectangle = uno.createUnoStruct('com.sun.star.awt.Rectangle', 4, 25, 390, 130)
        data, column = manager.Model.getGridModels(titles, rectangle.Width)
        grid = self._createGrid(self._tab1, data, column, 'Grid1', rectangle)
        handler = GridHandler(manager)
        grid.addSelectionListener(handler)

    def initColumnsList(self, columns):
        control = self.getColumnsList()
        self._initListBox(control, columns)

    def initOrdersList(self, columns, orders):
        control = self.getOrdersList()
        self._initListBox(control, columns)
        items = control.Model.StringItemList
        while orders.hasMoreElements():
            column = orders.nextElement()
            index = items.index(columns[column.Name])
            control.selectItemPos(index, True)

    def initButtons(self):
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

    def isDisposed(self):
        return self._dialog is None

    def dispose(self):
        self._dialog.dispose()
        self._tab1.dispose()
        self._tab2.dispose()
        self._dialog = None
        self._tab1 = None
        self._tab2 = None

    def showGridColumnHeader(self, enabled):
        self._getGrid().Model.ShowColumnHeader = enabled

# SpoolerView getter methods
    def execute(self):
        return self._dialog.execute()

    def getParent(self):
        return self._dialog.getPeer()

    def getGridWidth(self):
        width = self._getGrid().Model.Width
        return width

    def getSortDirection(self):
        ascending = not bool(self._getSortDirection().Model.State)
        return ascending

# SpoolerView private setter methods
    def _initListBox(self, control, columns):
        index = 0
        for column, name in columns.items():
            control.Model.insertItemText(index, name)
            control.Model.setItemData(index, column)
            index += 1

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

    def _getGrid(self):
        return self._tab1.getControl('Grid1')

    def _getSortDirection(self):
        return self._tab1.getControl('CheckBox1')

    def _getButtonAdd(self):
        return self._tab1.getControl('CommandButton1')

    def _getButtonRemove(self):
        return self._tab1.getControl('CommandButton2')

    def _getLabelState(self):
        return self._dialog.getControl('Label2')

    def getColumnsList(self):
        return self._tab1.getControl('ListBox1')

    def getOrdersList(self):
        return self._tab1.getControl('ListBox2')

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

    def _createGrid(self, page, data, column, name, rectangle):
        model = self._getGridModel(page, data, column, name, rectangle)
        page.Model.insertByName(name, model)
        return page.getControl(name)

    def _getGridModel(self, page, data, column, name, rectangle):
        service = 'com.sun.star.awt.grid.UnoControlGridModel'
        model = page.Model.createInstance(service)
        model.Name = name
        model.PositionX = rectangle.X
        model.PositionY = rectangle.Y
        model.Height = rectangle.Height
        model.Width = rectangle.Width
        model.GridDataModel = data
        model.ColumnModel = column
        model.SelectionModel = MULTI
        model.ShowColumnHeader = True
        #model.ShowRowHeader = True
        model.BackgroundColor = 16777215
        return model
