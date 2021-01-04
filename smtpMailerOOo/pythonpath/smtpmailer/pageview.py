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
    def initPage1(self, datasources):
        self._initDataSource(datasources)
        self.setPageStep(3)

    def setPageStep(self, step):
        self.Window.Model.Step = step

    def updatePage(self, value):
        control = self._getProgressBar()
        if control is not None:
            control.Value = value
        else:
            print("PageView.updatePage() *********************")

    def setMessageText(self, text):
        self._getColumnList().Model.StringItemList = ()
        self._getTableList().Model.StringItemList = ()
        self._getEmailList().Model.StringItemList = ()
        self._getIndexList().Model.StringItemList = ()
        self.enableDatasource(True)
        self.enableButton(False)
        self._getMessageLabel().Text = text

    def initGrid(self, address, recipient, handler):
        area = uno.createUnoStruct('com.sun.star.awt.Rectangle', 10, 60, 124, 122)
        grid1 = self._getGridControl(address, 'GridControl1', area, 'Addresses')
        area.X = 156
        grid2 = self._getGridControl(recipient, 'GridControl2', area, 'Recipients')
        grid1.addSelectionListener(handler)
        grid2.addSelectionListener(handler)

    def initDataSource(self, datasources):
        # TODO: Update the list of data sources and keep the selection if possible
        datasource = self._getSelectedDataSource()
        self._initDataSource(datasources)
        if datasource in datasources:
            self.selectDataSource(datasource)

    def selectDataSource(self, datasource):
        self._getDataSourceList().selectItem(datasource, True)

    def enablePage(self, enabled):
        self.enableDatasource(enabled)
        self._getColumnList().Model.Enabled = enabled
        self._getTableList().Model.Enabled = enabled
        self._getEmailList().Model.Enabled = enabled
        self._getIndexList().Model.Enabled = enabled
        self._getGeneralOption().Model.Enabled = enabled
        self._getDataSourceOption().Model.Enabled = enabled
        self._getTemplateOption().Model.Enabled = enabled
        self._getTitleOption().Model.Enabled = enabled
        self._getTitleField().Model.Enabled = enabled

    def enableDatasource(self, enabled):
        self._getDataSourceList().Model.Enabled = enabled
        self._getNewDataSourceButton().Model.Enabled = enabled

    def enableButton(self, enabled):
        self._getAddEmailButton().Model.Enabled = enabled
        self._getRemoveEmailButton().Model.Enabled = enabled
        self._getBeforeButton().Model.Enabled = enabled
        self._getAfterButton().Model.Enabled = enabled
        self._getAddIndexButton().Model.Enabled = enabled
        self._getRemoveIndexButton().Model.Enabled = enabled

    def initPage2(self, tables, columns):
        self._getAddressBookList().Model.StringItemList = tables
        self._getOrderList().Model.StringItemList = columns
        #control.selectItem(table, True)

    def setPage2(self, table, index):
        self._getAddressBookList().selectItem(table, True)
        self.setOrder(index)

    def setOrder(self, index):
        self._getOrderList().selectItemsPos(index, True)

    def refreshGridButton(self):
        self.updateControl(self._getAddressGrid())
        self.updateControl(self._getRecipientGrid())

    def initTables(self, tables):
        self._getTableList().Model.StringItemList = tables

    def setTables(self, table):
        self._getTableList().selectItem(table, True)

    def initColumns(self, columns):
        self._getColumnList().Model.StringItemList = columns

    def setColumns(self, index):
        self._getColumnList().selectItemPos(index, True)

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
                print("PageView.updateControl() 1")
                selected = control.hasSelectedRows()
                print("PageView.updateControl() 2")
                enabled = control.Model.GridDataModel.RowCount != 0
                print("PageView.updateControl() 3")
                self._getAddAllButton().Model.Enabled = enabled
                print("PageView.updateControl() 4")
                self._getAddButton().Model.Enabled = selected
                print("PageView.updateControl() 5")
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

    def getOrderColumns(self):
        return getStringItemList(self._getOrderList())

    def getSelectedTable(self):
        return self._getTableList().getSelectedItem()

    def getSelectedAbbressBook(self):
        return self._getAddressBookList().getSelectedItem()

    def getTables(self):
        return getStringItemList(self._getTableList())

    def getColumns(self):
        return getStringItemList(self._getColumnList())

# PageView private methods
    def _getControlByTag(self, tag):
        if tag == 'Columns':
            control = self._getColumnList()
        elif tag == 'EmailAddress':
            control = self._getEmailList()
        elif tag == 'PrimaryKey':
            control = self._getIndexList()
        elif tag == 'Addresses':
            control = self._getAddressGrid()
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

    def _getNewDataSourceButton(self):
        return self.Window.getControl('CommandButton1')

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

    def _getProgressBar(self):
        return self.Window.getControl('ProgressBar1')

    def _getGeneralOption(self):
        return self.Window.getControl('OptionButton1')

    def _getDataSourceOption(self):
        return self.Window.getControl('OptionButton2')

    def _getTemplateOption(self):
        return self.Window.getControl('OptionButton3')

    def _getTitleOption(self):
        return self.Window.getControl('OptionButton4')

    def _getMessageLabel(self):
        return self.Window.getControl('Label10')

    def _getTitleField(self):
        return self.Window.getControl('TextField1')

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
    def _initDataSource(self, datasources):
        self._getDataSourceList().Model.StringItemList = datasources

    def _getSelectedDataSource(self):
        return self._getDataSourceList().SelectedItem

    def _getGridControl(self, model, name, area, tag):
        dialog = self.Window.getModel()
        grid = self._getGridModel(model, dialog, name, area, tag)
        dialog.insertByName(name, grid)
        return self.Window.getControl(name)

    def _getGridModel(self, model, dialog, name, area, tag):
        grid = dialog.createInstance('com.sun.star.awt.grid.UnoControlGridModel')
        grid.Name = name
        grid.PositionX = area.X
        grid.PositionY = area.Y
        grid.Height = area.Height
        grid.Width = area.Width
        grid.Tag = tag
        grid.GridDataModel = model
        grid.ColumnModel = model.ColumnModel
        grid.SelectionModel = MULTI
        #grid.ShowRowHeader = True
        grid.BackgroundColor = 16777215
        return grid
