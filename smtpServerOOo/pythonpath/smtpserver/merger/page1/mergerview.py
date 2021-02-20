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

from unolib import getContainerWindow

from smtpserver.wizard import getStringItemList

from .mergerhandler import WindowHandler

from smtpserver import g_extension

from smtpserver import logMessage

import traceback


class MergerView(unohelper.Base):
    def __init__(self, ctx, manager, parent, datasources):
        self._ctx = ctx
        handler = WindowHandler(manager)
        self._window = getContainerWindow(ctx, parent, handler, g_extension, 'MergerPage1')
        self._initDataSource(datasources)
        self.setPageStep(3)

# PageView getter methods
    def getWindow(self):
        return self._window

    def hasIndex(self):
        return self._getIndexList().ItemCount != 0
        
    def isDataSourceSelected(self):
        return self._getDataSourceList().getSelectedItemPos() != -1

# PageView setter methods
    def enablePage(self, enabled):
        self.enableDatasource(enabled)
        self._getColumnList().Model.Enabled = enabled
        self._getTableList().Model.Enabled = enabled
        self._getEmailList().Model.Enabled = enabled
        self._getIndexList().Model.Enabled = enabled
        self._getQueryList().Model.Enabled = enabled

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
        self._getAddQueryButton().Model.Enabled = enabled
        self._getRemoveQueryButton().Model.Enabled = enabled

    def setPageStep(self, step):
        self._window.Model.Step = step

    def updatePage(self, value):
        control = self._getProgressBar()
        if control is not None:
            control.Value = value
        else:
            print("MergerView.updatePage() ERROR *********************")

    def setMessageText(self, text):
        self._getColumnList().Model.StringItemList = ()
        self._getTableList().Model.StringItemList = ()
        self._getEmailList().Model.StringItemList = ()
        self._getIndexList().Model.StringItemList = ()
        self._getQueryList().Model.StringItemList = ()
        self.enableDatasource(True)
        self.enableButton(False)
        self._getMessageLabel().Text = text

    def initDataSource(self, datasources):
        # TODO: Update the list of data sources and keep the selection if possible
        datasource = self._getSelectedDataSource()
        self._initDataSource(datasources)
        if datasource in datasources:
            self.selectDataSource(datasource)

    def selectDataSource(self, datasource):
        self._getDataSourceList().selectItem(datasource, True)

    def initTables(self, tables):
        control = self._getTableList()
        print("MergerView.initTables() 1")
        control.Model.StringItemList = tables
        print("MergerView.initTables() 2")
        table = control.getItem(0)
        control.selectItem(table, True)
        print("MergerView.initTables() 3")

    def initColumns(self, columns):
        control= self._getColumnList()
        print("MergerView.initColumns() 1")
        control.Model.StringItemList = columns
        print("MergerView.initColumns() 2")
        column = control.getItem(0)
        control.selectItem(column, True)
        print("MergerView.initColumns() 3")
        self._updateColumns(control)
        print("MergerView.initColumns() 4")

    def initQuery(self, queries):
        control = self._getQueryList()
        control.Model.StringItemList = queries

    def setTables(self, table):
        self._getTableList().selectItem(table, True)

    def setEmail(self, emails):
        self._getEmailList().Model.StringItemList = emails

    def setIndex(self, keys):
        self._getIndexList().Model.StringItemList = keys

    def setColumns(self, index):
        self._getColumnList().selectItemPos(index, True)

    def updateColumns(self):
        control = self._getColumnList()
        self._updateColumns(control)

    def updateButtonAdd(self, enabled, column):
        button = self._getAddEmailButton()
        email = self._getEmailList()
        button.Model.Enabled = self._canAddItem(enabled, column, email)
        button = self._getAddIndexButton()
        index = self._getIndexList()
        button.Model.Enabled = self._canAddItem(enabled, column, index)

    def enabledRemoveEmailButton(self, enabled):
        self._getRemoveEmailButton().Model.Enabled = enabled

    def enabledBeforeButton(self, enabled):
        self._getBeforeButton().Model.Enabled = enabled

    def enabledAfterButton(self, enabled):
        self._getAfterButton().Model.Enabled = enabled

    def addEmail(self):
        self._getAddEmailButton().Model.Enabled = False
        self._getRemoveEmailButton().Model.Enabled = False
        self._getBeforeButton().Model.Enabled = False
        self._getAfterButton().Model.Enabled = False
        control = self._getEmailList()
        column = self._getColumnList().getSelectedItem()
        self._addItem(control, column)

    def removeEmail(self):
        self._getRemoveEmailButton().Model.Enabled = False
        self._getBeforeButton().Model.Enabled = False
        self._getAfterButton().Model.Enabled = False
        control = self._getEmailList()
        column = control.getSelectedItem()
        position = control.getSelectedItemPos()
        control.Model.removeItem(position)
        selected = self._getColumnList().getSelectedItemPos() != -1
        enabled = self._canAddItem(selected, column, control)
        self._getAddEmailButton().Model.Enabled = enabled

    def moveBefore(self):
        control = self._getEmailList()
        column = control.getSelectedItem()
        position = control.getSelectedItemPos()
        control.Model.removeItem(position)
        position -= 1
        self._addItem(control, column, position)
        control.selectItemPos(position, True)

    def moveAfter(self):
        control = self._getEmailList()
        column = control.getSelectedItem()
        position = control.getSelectedItemPos()
        control.Model.removeItem(position)
        position += 1
        self._addItem(control, column, position)
        control.selectItemPos(position, True)

    def addIndex(self):
        self._getAddIndexButton().Model.Enabled = False
        self._getRemoveIndexButton().Model.Enabled = False
        control = self._getIndexList()
        column = self._getColumnList().getSelectedItem()
        self._addItem(control, column)

    def removeIndex(self):
        self._getRemoveIndexButton().Model.Enabled = False
        control = self._getIndexList()
        column = control.getSelectedItem()
        position = control.getSelectedItemPos()
        control.Model.removeItem(position)
        selected = self._getColumnList().getSelectedItemPos() != -1
        enabled = self._canAddItem(selected, column, control)
        self._getAddIndexButton().Model.Enabled = enabled

    def enabledRemoveIndexButton(self, enabled):
        self._getRemoveIndexButton().Model.Enabled = enabled

    def enabledRemoveQueryButton(self, enabled):
        self._getRemoveQueryButton().Model.Enabled = enabled

# MergerView private methods
    def _updateColumns(self, control):
        enabled = control.getSelectedItemPos() != -1
        column = control.getSelectedItem()
        self.updateButtonAdd(enabled, column)

    def _canAddItem(self, enabled, column, listbox):
        if enabled and listbox.ItemCount != 0:
            columns = listbox.Model.StringItemList
            return column not in columns
        return enabled

    def _addItem(self, control, column, position=-1):
        if position != -1:
            control.Model.insertItemText(position, column)
        else:
            position = control.getItemCount()
            control.Model.insertItemText(position, column)









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



    def removeEmailAdress(self):
        self._getRemoveEmailButton().Model.Enabled = False
        self._getBeforeButton().Model.Enabled = False
        self._getAfterButton().Model.Enabled = False
        control = self._getEmailList()
        item = control.getSelectedItem()
        control.Model.StringItemList = self._removeItem(control, item)
        enabled = self._canAddItem(self._getColumnList(), control)
        self._getAddEmailButton().Model.Enabled = enabled

    def moveEmailAdress(self, tag):
        control = self._getEmailList()
        index = control.getSelectedItemPos()
        column = control.getSelectedItem()
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




    def _removeItem(self, control, item, mutable=False):
        items = list(getStringItemList(control))
        items.remove(item)
        return items if mutable else tuple(items)

# PageView private message methods
    def _getPageTitle(self, pageid):
        return 'MergerPage%s.Title' % pageid

# PageView private getter control methods
    def _getDataSourceList(self):
        return self._window.getControl('ListBox1')

    def _getColumnList(self):
        return self._window.getControl('ListBox2')

    def _getTableList(self):
        return self._window.getControl('ListBox3')

    def _getEmailList(self):
        return self._window.getControl('ListBox4')

    def _getIndexList(self):
        return self._window.getControl('ListBox5')

    def _getQueryList(self):
        return self._window.getControl('ComboBox1')

    def _getNewDataSourceButton(self):
        return self._window.getControl('CommandButton1')

    def _getAddEmailButton(self):
        return self._window.getControl('CommandButton2')

    def _getRemoveEmailButton(self):
        return self._window.getControl('CommandButton3')

    def _getBeforeButton(self):
        return self._window.getControl('CommandButton4')

    def _getAfterButton(self):
        return self._window.getControl('CommandButton5')

    def _getAddIndexButton(self):
        return self._window.getControl('CommandButton6')

    def _getRemoveIndexButton(self):
        return self._window.getControl('CommandButton7')

    def _getAddQueryButton(self):
        return self._window.getControl('CommandButton8')

    def _getRemoveQueryButton(self):
        return self._window.getControl('CommandButton9')

    def _getMessageLabel(self):
        return self._window.getControl('Label10')

    def _getProgressBar(self):
        return self._window.getControl('ProgressBar1')

# PageView private control methods
    def _initDataSource(self, datasources):
        self._getDataSourceList().Model.StringItemList = datasources

    def _getSelectedDataSource(self):
        return self._getDataSourceList().SelectedItem
