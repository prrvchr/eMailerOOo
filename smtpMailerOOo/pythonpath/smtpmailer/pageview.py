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
    def initPage1(self, model):
        return self._initDataSources(model)

    def initPage2(self, model, handler):
        area = uno.createUnoStruct('com.sun.star.awt.Rectangle', 10, 60, 115, 115)
        grid1 = self._getGridControl(model._address, 'GridControl1', area, 'Addresses')
        grid1.addSelectionListener(handler)
        area.X = 160
        grid2 = self._getGridControl(model._recipient, 'GridControl2', area, 'Recipients')
        grid2.addSelectionListener(handler)

    def initAddressBook(self, model):
        control = self._getAddressBook()
        tables = model.TableNames
        #control.Model.StringItemList = ()
        control.Model.StringItemList = tables
        table = model._address.Command
        table = tables[0] if table == '' and len(tables) != 0 else table
        control.selectItem(table, True)
        print("PageView.refreshAddress() %s -%s" % (table, tables))
        return table

    def initOrderColumn(self, model):
        control = self._getOrderColumn()
        control.Model.StringItemList = ()
        control.Model.StringItemList = model.ColumnNames
        columns = model.getOrderIndex()
        control.selectItemsPos(columns, True)

    def refreshGridButton(self):
        self.updateControl(self._getAddress())
        self.updateControl(self._getRecipient())

    def setTables(self, tables, table):
        control = self._getTables()
        control.Model.StringItemList = tables
        control.selectItem(table, True)

    def setColumns(self, columns, column):
        control = self._getColumns()
        control.Model.StringItemList = columns
        control.selectItem(column, True)

    def setEmailAddress(self, emails):
        self._getEmailAddress().Model.StringItemList = emails

    def setPrimaryKey(self, keys):
        self._getPrimaryKey().Model.StringItemList = keys

    def updateControlByTag(self, tag):
        control = self._getControlByTag(tag)
        self.updateControl(control)

    def updateControl(self, control):
        try:
            print("PageView.updateControl() 1")
            tag = control.Model.Tag
            #if tag == 'DataSource':
            #    control.Model.StringItemList = model.DataSources
            if tag == 'Columns':
                button = self._getAddEmailAddress()
                button.Model.Enabled = self._canAddItem(control, self._getEmailAddress())
                button = self._getAddPrimaryKey()
                button.Model.Enabled = self._canAddItem(control, self._getPrimaryKey())
            elif tag == 'EmailAddress':
                imax = control.ItemCount -1
                position = control.getSelectedItemPos()
                self._getRemoveEmailAddress().Model.Enabled = position != -1
                self._getMoveBefore().Model.Enabled = position > 0
                self._getMoveAfter().Model.Enabled = -1 < position < imax
            elif tag == 'PrimaryKey':
                print("PageView.updateControl() PrimaryKey ***************")
                position = control.getSelectedItemPos()
                self._getRemovePrimaryKey().Model.Enabled = position != -1
            elif tag == 'Addresses':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                self._getAddAllRecipient().Model.Enabled = enabled
                self._getAddRecipient().Model.Enabled = selected
            elif tag == 'Recipients':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                self._getRemoveRecipient().Model.Enabled = selected
                self._getRemoveAllRecipient().Model.Enabled = enabled
            else:
                print("PageView.updateControl() Error ***************")
            print("PageView.updateControl() 2")
        except Exception as e:
            print("WizardHandler._updateControl() ERROR: %s - %s" % (e, traceback.print_exc()))

    def updateTables(self, state):
        self._getTables().Model.Enabled = bool(state)

    def updateAddRecipient(self, enabled):
        self._getAddRecipient().Model.Enabled = enabled

    def updateRemoveRecipient(self, enabled):
        self._getRemoveRecipient().Model.Enabled = enabled

    def addEmailAdress(self):
        self._getAddEmailAddress().Model.Enabled = False
        self._getRemoveEmailAddress().Model.Enabled = False
        self._getMoveBefore().Model.Enabled = False
        self._getMoveAfter().Model.Enabled = False
        item = self._getColumns().getSelectedItem()
        self._addItemColumn(self._getEmailAddress(), item)

    def addPrimaryKey(self):
        self._getAddPrimaryKey().Model.Enabled = False
        self._getRemovePrimaryKey().Model.Enabled = False
        item = self._getColumns().getSelectedItem()
        self._addItemColumn(self._getPrimaryKey(), item)

    def removeEmailAdress(self):
        self._getRemoveEmailAddress().Model.Enabled = False
        self._getMoveBefore().Model.Enabled = False
        self._getMoveAfter().Model.Enabled = False
        control = self._getEmailAddress()
        item = control.getSelectedItem()
        control.Model.StringItemList = self._removeItem(control, item)
        enabled = self._canAddItem(self._getColumns(), control)
        self._getAddEmailAddress().Model.Enabled = enabled

    def removePrimaryKey(self):
        self._getRemovePrimaryKey().Model.Enabled = False
        control = self._getPrimaryKey()
        item = control.getSelectedItem()
        control.Model.StringItemList = self._removeItem(control, item)
        enabled = self._canAddItem(self._getColumns(), control)
        self._getAddPrimaryKey().Model.Enabled = enabled

    def moveEmailAdress(self, tag):
        control = self._getEmailAddress()
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

    def isPrimaryKeySet(self):
        return self._getPrimaryKey().ItemCount != 0

# PageView getter methods
    def getPageTitle(self, model, pageid):
        return model.resolveString(self._getPageTitle(pageid))

# PageView private methods
    def _getControlByTag(self, tag):
        if tag == 'Columns':
            control = self._getColumns()
        elif tag == 'EmailAddress':
            control = self._getEmailAddress()
        elif tag == 'PrimaryKey':
            control = self._getPrimaryKey()
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
    def _getDataSource(self):
        return self.Window.getControl('ListBox1')

    def _getColumns(self):
        return self.Window.getControl('ListBox2')

    def _getTables(self):
        return self.Window.getControl('ListBox3')

    def _getEmailAddress(self):
        return self.Window.getControl('ListBox4')

    def _getPrimaryKey(self):
        return self.Window.getControl('ListBox5')

    def _getAddEmailAddress(self):
        return self.Window.getControl('CommandButton2')

    def _getRemoveEmailAddress(self):
        return self.Window.getControl('CommandButton3')

    def _getMoveBefore(self):
        return self.Window.getControl('CommandButton4')

    def _getMoveAfter(self):
        return self.Window.getControl('CommandButton5')

    def _getAddPrimaryKey(self):
        return self.Window.getControl('CommandButton6')

    def _getRemovePrimaryKey(self):
        return self.Window.getControl('CommandButton7')


# PageView private getter control methods for Page2
    def _getAddressBook(self):
        return self.Window.getControl('ListBox1')

    def _getOrderColumn(self):
        return self.Window.getControl('ListBox2')

    def _getAddress(self):
        return self.Window.getControl('GridControl1')

    def _getRecipient(self):
        return self.Window.getControl('GridControl2')

    def _getAddAllRecipient(self):
        return self.Window.getControl('CommandButton1')

    def _getAddRecipient(self):
        return self.Window.getControl('CommandButton2')

    def _getRemoveRecipient(self):
        return self.Window.getControl('CommandButton3')

    def _getRemoveAllRecipient(self):
        return self.Window.getControl('CommandButton4')


# PageView private setter control methods
    def _initDataSources(self, model):
        initialized = False
        datasources = model.DataSources
        control = self._getDataSource()
        control.Model.StringItemList = datasources
        datasource = model.getDocumentDataSource()
        if datasource in datasources:
            if model.setDataSource(self, datasource):
                control.selectItem(datasource, True)
                initialized = True
        return initialized

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
