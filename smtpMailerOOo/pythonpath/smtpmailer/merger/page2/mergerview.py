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

from smtpmailer import createService
from smtpmailer import getContainerWindow
from smtpmailer import logMessage
from smtpmailer import g_extension

from .mergerhandler import Tab1Handler
from .mergerhandler import Tab2Handler

import traceback


class MergerView(unohelper.Base):
    def __init__(self, ctx, manager, parent, tables, enabled, message):
        self._ctx = ctx
        self._window = getContainerWindow(ctx, parent, None, g_extension, 'MergerPage2')
        rectangle = uno.createUnoStruct('com.sun.star.awt.Rectangle', 0, 5, 285, 195)
        tab1, tab2 = self._getTabPages(manager, 'Tab1', rectangle, 1)
        parent = tab1.getPeer()
        handler = Tab1Handler(manager)
        self._tab1 = getContainerWindow(ctx, parent, handler, g_extension, 'MergerTab1')
        self._tab1.setVisible(True)
        parent = tab2.getPeer()
        handler = Tab2Handler(manager)
        self._tab2 = getContainerWindow(ctx, parent, handler, g_extension, 'MergerTab2')
        self._tab2.setVisible(True)
        self._rectangle = uno.createUnoStruct('com.sun.star.awt.Rectangle', 2, 25, 277, 152)
        self._initTables(tables, enabled)
        self.setMessage(message)

# MergerView getter methods
    def getGridParents(self):
        return self._tab1.getPeer(), self._tab2.getPeer()

    def getGridPosSize(self):
        return self._rectangle

    def getWindow(self):
        return self._window

    def getTable(self):
        return self._getTable().getSelectedItem()

# MergerView setter methods
    def setTable(self, table):
        self._getTable().selectItem(table, True)

    def initTables(self, tables, table, enabled):
        control = self._getTable()
        control.Model.removeAllItems()
        control.Model.StringItemList = tables
        control.selectItem(table, True)
        control.Model.Enabled = enabled

    def enableAdd(self, enabled):
        self._getAdd().Model.Enabled = enabled

    def enableRemove(self, enabled):
        self._getRemove().Model.Enabled = enabled

    def enableAddAll(self, enabled):
        self._getAddAll().Model.Enabled = enabled

    def enableRemoveAll(self, enabled):
        self._getRemoveAll().Model.Enabled = enabled

    def setMessage(self, message):
        self._getMessage().Text = message

# MergerView private getter control methods
    def _getTable(self):
        return self._tab1.getControl('ListBox1')

    def _getAddAll(self):
        return self._tab1.getControl('CommandButton1')

    def _getAdd(self):
        return self._tab1.getControl('CommandButton2')

    def _getRemoveAll(self):
        return self._tab2.getControl('CommandButton1')

    def _getRemove(self):
        return self._tab2.getControl('CommandButton2')

    def _getMessage(self):
        return self._tab2.getControl('Label1')

# MergerView private methods
    def _getTabPages(self, manager, name, rectangle, i):
        model = self._getTabModel(rectangle)
        self._window.Model.insertByName(name, model)
        tab = self._window.getControl(name)
        tab1 = self._getTabPage(manager, model, tab, 0)
        tab2 = self._getTabPage(manager, model, tab, 1)
        tab.ActiveTabPageID = i
        return tab1, tab2

    def _getTabModel(self, rectangle):
        service = 'com.sun.star.awt.tab.UnoControlTabPageContainerModel'
        model = self._window.Model.createInstance(service)
        model.PositionX = rectangle.X
        model.PositionY = rectangle.Y
        model.Width = rectangle.Width
        model.Height = rectangle.Height
        return model

    def _getTabPage(self, manager, model, tab, i):
        page = model.createTabPage(i +1)
        page.Title = manager.getTabTitle(i +1)
        index = model.getCount()
        model.insertByIndex(index, page)
        return tab.getControls()[i]

    def _initTables(self, tables, enabled):
        control = self._getTable()
        control.Model.StringItemList = tables
        control.Model.Enabled = enabled
