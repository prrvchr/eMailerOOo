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

from .griddatamodel import GridDataModel
from .wizardtools import getStringItemList

from com.sun.star.view.SelectionType import MULTI

from .logger import logMessage

import traceback


class PageView(unohelper.Base):
    def __init__(self, ctx, window):
        self.ctx = ctx
        self.Window = window
        print("PageView.__init__()")

# PageView setter methods
    def initPage1(self, datasources, datasource):
        self._initDataSource(datasources, datasource)

    def initPage2(self, address, recipient, handler):
        area = uno.createUnoStruct('com.sun.star.awt.Rectangle', 10, 60, 115, 115)
        grid1 = self._getGridControl(address, 'GridControl1', area, 'Addresses')
        area.X = 160
        grid2 = self._getGridControl(recipient, 'GridControl2', area, 'Recipients')
        grid1.addSelectionListener(handler)
        grid2.addSelectionListener(handler)

    def initDataSource(self, datasources):
        datasource = self._initDataSource(datasources)
        self.selectDataSource(datasource)

    def selectDataSource(self, datasource):
        self._getDataSourceList().selectItem(datasource, True)

    def initAddressBook(self, model):
        control = self._getAddressBookList()
        tables = model.TableNames
        #control.Model.StringItemList = ()
        control.Model.StringItemList = tables
        table = model._address.Command
        table = tables[0] if table == '' and len(tables) != 0 else table
        control.selectItem(table, True)
        print("PageView.refreshAddress() %s -%s" % (table, tables))
        return table

    def initOrderColumn(self, model):
        control = self._getOrderList()
        control.Model.StringItemList = ()
        control.Model.StringItemList = model.ColumnNames
        columns = model.getOrderIndex()
        control.selectItemsPos(columns, True)

    def refreshGridButton(self):
        self.updateControl(self._getAddressGrid())
        self.updateControl(self._getRecipientGrid())

    def setTables(self, tables, table):
        control = self._getTableList()
        control.Model.StringItemList = tables
        control.selectItem(table, True)

    def setColumns(self, columns):
        control = self._getColumnList()
        control.Model.StringItemList = columns
        control.selectItemPos(0, True)

    def setEmailAddress(self, emails):
        self._getEmailList().Model.StringItemList = emails

    def setPrimaryKey(self, keys):
        self._getIndexList().Model.StringItemList = keys

    def updateControlByTag(self, tag):
        control = self._getControlByTag(tag)
        self.updateControl(control)

    def updateControl(self, control):
        try:
            print("PageView.updateControl() 1")
            tag = control.Model.Tag
            if tag == 'Columns':
                button = self._getAddEmailButton()
                button.Model.Enabled = self._canAddItem(control, self._getEmailList())
                button = self._getAddIndexButton()
                button.Model.Enabled = self._canAddItem(control, self._getIndexList())
            elif tag == 'EmailAddress':
                imax = control.ItemCount -1
                position = control.getSelectedItemPos()
                self._getRemoveEmailButton().Model.Enabled = position != -1
                self._getBeforeButton().Model.Enabled = position > 0
                self._getAfterButton().Model.Enabled = -1 < position < imax
            elif tag == 'PrimaryKey':
                print("PageView.updateControl() PrimaryKey ***************")
                position = control.getSelectedItemPos()
                self._getRemoveIndexButton().Model.Enabled = position != -1
            elif tag == 'Addresses':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                self._getAddAllButton().Model.Enabled = enabled
                self._getAddButton().Model.Enabled = selected
            elif tag == 'Recipients':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                self._getRemoveButton().Model.Enabled = selected
                self._getRemoveAllButton().Model.Enabled = enabled
            else:
                print("PageView.updateControl() Error ***************")
            print("PageView.updateControl() 2")
        except Exception as e:
            print("WizardHandler._updateControl() ERROR: %s - %s" % (e, traceback.print_exc()))

    def updateTables(self, state):
        self._getTableList().Model.Enabled = bool(state)

    def updateAddRecipient(self, enabled):
        self._getAddButton().Model.Enabled = enabled

    def updateRemoveRecipient(self, enabled):
        self._getRemoveButton().Model.Enabled = enabled

    def addEmailAdress(self):
        self._getAddEmailButton().Model.Enabled = False
        self._getRemoveEmailButton().Model.Enabled = False
        self._getBeforeButton().Model.Enabled = False
        self._getAfterButton().Model.Enabled = False
        item = self._getColumnList().getSelectedItem()
        self._addItemColumn(self._getEmailList(), item)

    def addPrimaryKey(self):
        self._getAddIndexButton().Model.Enabled = False
        self._getRemoveIndexButton().Model.Enabled = False
        item = self._getColumnList().getSelectedItem()
        self._addItemColumn(self._getIndexList(), item)

    def removeEmailAdress(self):
        self._getRemoveEmailButton().Model.Enabled = False
        self._getBeforeButton().Model.Enabled = False
        self._getAfterButton().Model.Enabled = False
        control = self._getEmailList()
        item = control.getSelectedItem()
        control.Model.StringItemList = self._removeItem(control, item)
        enabled = self._canAddItem(self._getColumnList(), control)
        self._getAddEmailButton().Model.Enabled = enabled

    def removePrimaryKey(self):
        self._getRemoveIndexButton().Model.Enabled = False
        control = self._getIndexList()
        item = control.getSelectedItem()
        control.Model.StringItemList = self._removeItem(control, item)
        enabled = self._canAddItem(self._getColumnList(), control)
        self._getAddIndexButton().Model.Enabled = enabled

    def moveEmailAdress(self, tag):
        control = self._getEmailList()
        index = control.getSelectedItemPos()
        item = control.getSelectedItem()
        items = self._removeItem(control, item, True)
        if tag == 'Before':
            index -= 1
        else:
            index += 1
        control.Model.StringItemList = self._addItem(items, item, index)
        control.selectItemPos(index, True)
        return True

    def canAdvancePage1(self):
        return all((self._getIndexList().ItemCount != 0,
                    self._getDataSourceList().getSelectedItemPos() != -1))

# PageView getter methods
    def getPageTitle(self, model, pageid):
        return model.resolveString(self._getPageTitle(pageid))

    def getSelectedRecipients(self):
        control = self._getRecipientGrid()
        rows = control.getSelectedRows()
        control.deselectAllRows()
        return rows

    def deselectAllRecipients(self):
        self._getRecipientGrid().deselectAllRows()

    def getSelectedAddress(self):
        return self._getAddressGrid().getSelectedRows()

    def getEmailColumns(self):
        return getStringItemList(self._getEmailList())

    def getIndexColumns(self):
        return getStringItemList(self._getIndexList())

# PageView private methods
    def _getControlByTag(self, tag):
        if tag == 'Columns':
            control = self._getColumnList()
        elif tag == 'EmailAddress':
            control = self._getEmailList()
        elif tag == 'PrimaryKey':
            control = self._getIndexList()
        elif tag == 'Addresses':
            control = self._getAddressBookList()
        elif tag == 'Recipients':
            control = self._getRecipientGrid()
        else:
            print("PageView._getControlByTag() Error: '%s' don't exist!!! **************" % tag)
        return control

    def _canAddItem(self, control, listbox):
        enabled = control.getSelectedItemPos() != -1
        if enabled and listbox.ItemCount != 0:
            column = control.getSelectedItem()
            columns = getStringItemList(listbox)
            return column not in columns
        return enabled

    def _addItemColumn(self, control, item):
        items = list(getStringItemList(control))
        control.Model.StringItemList = self._addItem(items, item)

    def _addItem(self, items, item, index=-1):
        if index != -1:
            items.insert(index, item)
        else:
            items.append(item)
        return tuple(items)

    def _removeItem(self, control, item, mutable=False):
        items = list(getStringItemList(control))
        items.remove(item)
        return items if mutable else tuple(items)

# PageView private message methods
    def _getPageTitle(self, pageid):
        return 'PageWizard%s.Title' % pageid

# PageView private getter control methods for Page1
    def _getDataSourceList(self):
        return self.Window.getControl('ListBox1')

    def _getColumnList(self):
        return self.Window.getControl('ListBox2')

    def _getTableList(self):
        return self.Window.getControl('ListBox3')

    def _getEmailList(self):
        return self.Window.getControl('ListBox4')

    def _getIndexList(self):
        return self.Window.getControl('ListBox5')

    def _getAddEmailButton(self):
        return self.Window.getControl('CommandButton2')

    def _getRemoveEmailButton(self):
        return self.Window.getControl('CommandButton3')

    def _getBeforeButton(self):
        return self.Window.getControl('CommandButton4')

    def _getAfterButton(self):
        return self.Window.getControl('CommandButton5')

    def _getAddIndexButton(self):
        return self.Window.getControl('CommandButton6')

    def _getRemoveIndexButton(self):
        return self.Window.getControl('CommandButton7')


# PageView private getter control methods for Page2
    def _getAddressBookList(self):
        return self.Window.getControl('ListBox1')

    def _getOrderList(self):
        return self.Window.getControl('ListBox2')

    def _getAddressGrid(self):
        return self.Window.getControl('GridControl1')

    def _getRecipientGrid(self):
        return self.Window.getControl('GridControl2')

    def _getAddAllButton(self):
        return self.Window.getControl('CommandButton1')

    def _getAddButton(self):
        return self.Window.getControl('CommandButton2')

    def _getRemoveButton(self):
        return self.Window.getControl('CommandButton3')

    def _getRemoveAllButton(self):
        return self.Window.getControl('CommandButton4')


# PageView private control methods
    def _initDataSource(self, datasources, datasource=None):
        control = self._getDataSourceList()
        if datasource is None:
            datasource = control.getSelectedItem()
        if getStringItemList(control) != datasources:
            control.Model.StringItemList = datasources
        return datasource

    def _getGridControl(self, rowset, name, area, tag):
        dialog = self.Window.getModel()
        model = self._getGridModel(rowset, dialog, name, area, tag)
        dialog.insertByName(name, model)
        return self.Window.getControl(name)

    def _getGridModel(self, rowset, dialog, name, area, tag):
        data = GridDataModel(self.ctx, rowset)
        model = dialog.createInstance('com.sun.star.awt.grid.UnoControlGridModel')
        model.Name = name
        model.PositionX = area.X
        model.PositionY = area.Y
        model.Height = area.Height
        model.Width = area.Width
        model.Tag = tag
        model.GridDataModel = data
        model.ColumnModel = data.ColumnModel
        model.SelectionModel = MULTI
        #model.ShowRowHeader = True
        model.BackgroundColor = 16777215
        return model
