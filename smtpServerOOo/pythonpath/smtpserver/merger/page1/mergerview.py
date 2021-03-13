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

import unohelper

from unolib import getContainerWindow

from .mergerhandler import WindowHandler

from smtpserver import g_extension

from smtpserver import logMessage

import traceback


class MergerView(unohelper.Base):
    def __init__(self, ctx, manager, parent, datasources):
        handler = WindowHandler(manager)
        self._window = getContainerWindow(ctx, parent, handler, g_extension, 'MergerPage1')
        self.initDataSource(datasources)
        self.setPageStep(3)

# MergerView getter methods
    def isDisposed(self):
        return self._window is None

    def getWindow(self):
        return self._window

    def hasEmail(self):
        return self._getEmail().getItemCount() > 0

    def hasIndex(self):
        return self._getIndex().getItemCount() > 0

    def isQuerySelected(self):
        control = self._getQuery()
        if control.getItemCount() > 0:
            query = control.getText()
            queries = control.getItems()
            return query in queries
        return False

    def getDataSource(self):
        datasource = self._getDataSource().getSelectedItem()
        return datasource

    def getQueries(self):
        control = self._getQuery()
        queries = control.getItems()
        query = control.getText()
        return queries, query

    def getQuery(self):
        control = self._getQuery()
        query = control.getText().strip()
        return query

    def getTable(self):
        table = self._getTable().getSelectedItem()
        return table

    def getEmails(self):
        columns = self._getEmail().Model.StringItemList
        return columns

    def getIndexes(self):
        columns = self._getIndex().Model.StringItemList
        return columns

# MergerView setter methods
    def dispose(self):
        self._window.dispose()
        self._window = None

    def setPageStep(self, step):
        self._window.Model.Step = step

    def enablePage(self, enabled):
        self.enableDatasource(enabled)
        control = self._getTable()
        self._enableListBox(control, enabled)
        control = self._getColumn()
        self._enableListBox(control, enabled)
        self._getQuery().Model.Enabled = enabled
        self._getEmail().Model.Enabled = enabled
        self._getIndex().Model.Enabled = enabled

    def enableDatasource(self, enabled):
        self._getDataSource().Model.Enabled = enabled
        self._getNewDataSource().Model.Enabled = enabled

    def enableButton(self, enabled):
        self._getAddEmail().Model.Enabled = enabled
        self._getRemoveEmail().Model.Enabled = enabled
        self._getBefore().Model.Enabled = enabled
        self._getAfter().Model.Enabled = enabled
        self._getAddIndex().Model.Enabled = enabled
        self._getRemoveIndex().Model.Enabled = enabled
        self._getAddQuery().Model.Enabled = enabled
        self._getRemoveQuery().Model.Enabled = enabled

    def updateProgress(self, value):
        if not self.isDisposed():
            self._getProgress().Value = value
        else:
            print("MergerView.updateProgress() ERROR *********************")

    def setMessageText(self, text):
        self._clearContent()
        self.enableDatasource(True)
        self.enableButton(False)
        self._getMessage().Text = text

    def initDataSource(self, datasources):
        self._getDataSource().Model.StringItemList = datasources

    def selectDataSource(self, datasource):
        self._getDataSource().selectItem(datasource, True)

    def initTables(self, tables):
        control = self._getTable()
        control.Model.StringItemList = tables
        if control.getItemCount() > 0:
            control.selectItemPos(0, True)

    def setEmailLabel(self, text):
        self._getEmailLabel().Text = text

    def setIndexLabel(self, text):
        self._getIndexLabel().Text = text

    def initColumns(self, columns):
        control= self._getColumn()
        control.Model.StringItemList = columns
        if control.getItemCount() > 0:
            control.selectItemPos(0, True)
        self._updateColumns(control)

    def initQuery(self, queries):
        control = self._getQuery()
        control.Model.StringItemList = queries
        if control.getItemCount() > 0:
            query = control.getItem(0)
            control.setText(query)
        else:
            control.setText('')

    def setTables(self, table):
        self._getTable().selectItem(table, True)

    def setEmail(self, emails):
        self._getEmail().Model.StringItemList = emails

    def setIndex(self, indexes):
        self._getIndex().Model.StringItemList = indexes

    def setColumns(self, index):
        self._getColumn().selectItemPos(index, True)

    def updateColumns(self):
        control = self._getColumn()
        self._updateColumns(control)

    def updateAddEmail(self):
        email = self._getEmail()
        column = self._getColumn().getSelectedItem()
        enabled = self._canAddItem(True, True, column, email)
        self._getAddEmail().Model.Enabled = enabled

    def updateAddIndex(self):
        email = self._getIndex()
        column = self._getColumn().getSelectedItem()
        enabled = self._canAddItem(True, True, column, email)
        self._getAddIndex().Model.Enabled = enabled

    def updateButtonAdd(self, selected, column):
        button = self._getAddEmail()
        email = self._getEmail()
        enabled = self._canAddItem(True, selected, column, email)
        button.Model.Enabled = enabled
        button = self._getAddIndex()
        index = self._getIndex()
        enable = self.isQuerySelected()
        enabled = self._canAddItem(enable, selected, column, index)
        button.Model.Enabled = enabled

    def enableRemoveQuery(self, enabled):
        self._getRemoveQuery().Model.Enabled = enabled

    def enableAddQuery(self, enabled):
        self._getAddQuery().Model.Enabled = enabled

    def addQuery(self, query):
        control = self._getQuery()
        control.setText('')
        count = control.getItemCount()
        control.addItem(query, count)
        self._getRemoveQuery().Model.Enabled = False

    def removeQuery(self, query):
        self._getRemoveQuery().Model.Enabled = False
        control = self._getQuery()
        query = control.getText()
        queries = control.getItems()
        if query in queries:
            control.setText('')
            position = queries.index(query)
            control.removeItems(position, 1)

    def enableRemoveEmail(self, enabled):
        self._getRemoveEmail().Model.Enabled = enabled

    def enableBefore(self, enabled):
        self._getBefore().Model.Enabled = enabled

    def enableAfter(self, enabled):
        self._getAfter().Model.Enabled = enabled

    def addEmail(self):
        self._getAddEmail().Model.Enabled = False
        self._getRemoveEmail().Model.Enabled = False
        self._getBefore().Model.Enabled = False
        self._getAfter().Model.Enabled = False
        control = self._getEmail()
        column = self._getColumn().getSelectedItem()
        self._addItem(control, column)

    def removeEmail(self):
        self._getRemoveEmail().Model.Enabled = False
        self._getBefore().Model.Enabled = False
        self._getAfter().Model.Enabled = False
        control = self._getEmail()
        column = control.getSelectedItem()
        position = control.getSelectedItemPos()
        control.Model.removeItem(position)
        selected = self._getColumn().getSelectedItemPos() != -1
        enabled = self._canAddItem(True, selected, column, control)
        self._getAddEmail().Model.Enabled = enabled

    def moveBefore(self):
        control = self._getEmail()
        column = control.getSelectedItem()
        position = control.getSelectedItemPos()
        control.Model.removeItem(position)
        position -= 1
        self._addItem(control, column, position)
        control.selectItemPos(position, True)

    def moveAfter(self):
        control = self._getEmail()
        column = control.getSelectedItem()
        position = control.getSelectedItemPos()
        control.Model.removeItem(position)
        position += 1
        self._addItem(control, column, position)
        control.selectItemPos(position, True)

    def enableAddIndex(self, enabled):
        self._getAddIndex().Model.Enabled = enabled

    def addIndex(self):
        self._getAddIndex().Model.Enabled = False
        self._getRemoveIndex().Model.Enabled = False
        control = self._getIndex()
        column = self._getColumn().getSelectedItem()
        self._addItem(control, column)

    def removeIndex(self):
        self._getRemoveIndex().Model.Enabled = False
        control = self._getIndex()
        column = control.getSelectedItem()
        position = control.getSelectedItemPos()
        control.Model.removeItem(position)
        selected = self._getColumn().getSelectedItemPos() != -1
        enabled = self._canAddItem(True, selected, column, control)
        self._getAddIndex().Model.Enabled = enabled

    def enableRemoveIndex(self, enabled):
        self._getRemoveIndex().Model.Enabled = enabled

# MergerView private setter methods
    def _enableListBox(self, control, enabled):
        control.Model.Enabled = enabled
        if not enabled:
            control.Model.StringItemList = ()

    def _updateColumns(self, control):
        enabled = control.getSelectedItemPos() != -1
        column = control.getSelectedItem()
        self.updateButtonAdd(enabled, column)

    def _canAddItem(self, enabled, selected, column, control):
        if enabled and selected:
            columns = control.Model.StringItemList
            enabled = column not in columns
        return enabled

    def _addItem(self, control, column, position=-1):
        if position != -1:
            control.Model.insertItemText(position, column)
        else:
            position = control.getItemCount()
            control.Model.insertItemText(position, column)

    def _addQuery(self, control, query):
        control.setText('')
        count = control.getItemCount()
        control.addItem(query, count)
        self._getRemoveQuery().Model.Enabled = False
        #self._getAddQuery().Model.Enabled = False

    def _clearContent(self):
        self._getColumn().Model.StringItemList = ()
        self._getTable().Model.StringItemList = ()
        self._getEmail().Model.StringItemList = ()
        self._getIndex().Model.StringItemList = ()
        control = self._getQuery()
        count = control.getItemCount()
        control.removeItems(0, count)
        control.setText('')

# MergerView private getter control methods
    def _getDataSource(self):
        return self._window.getControl('ListBox1')

    def _getTable(self):
        return self._window.getControl('ListBox2')

    def _getColumn(self):
        return self._window.getControl('ListBox3')

    def _getEmail(self):
        return self._window.getControl('ListBox4')

    def _getIndex(self):
        return self._window.getControl('ListBox5')

    def _getQuery(self):
        return self._window.getControl('ComboBox1')

    def _getProgress(self):
        return self._window.getControl('ProgressBar1')

    def _getNewDataSource(self):
        return self._window.getControl('CommandButton1')

    def _getAddQuery(self):
        return self._window.getControl('CommandButton2')

    def _getRemoveQuery(self):
        return self._window.getControl('CommandButton3')

    def _getAddEmail(self):
        return self._window.getControl('CommandButton4')

    def _getRemoveEmail(self):
        return self._window.getControl('CommandButton5')

    def _getBefore(self):
        return self._window.getControl('CommandButton6')

    def _getAfter(self):
        return self._window.getControl('CommandButton7')

    def _getAddIndex(self):
        return self._window.getControl('CommandButton8')

    def _getRemoveIndex(self):
        return self._window.getControl('CommandButton9')

    def _getMessage(self):
        return self._window.getControl('Label7')

    def _getEmailLabel(self):
        return self._window.getControl('Label11')

    def _getIndexLabel(self):
        return self._window.getControl('Label12')
