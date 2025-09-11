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

from ....unotool import getContainerWindow

from ....configuration import g_identifier

import traceback


class MergerView(unohelper.Base):
    def __init__(self, ctx, handler, parent, addressbooks):
        self._window = getContainerWindow(ctx, parent, handler, g_identifier, 'MergerPage1')
        self.initAddressBook(addressbooks)
        self.setPageStep(3)

# MergerView getter methods
    def isDisposed(self):
        return self._window is None

    def getWindow(self):
        return self._window

    # AddressBook getter methods
    def getAddressBook(self):
        return self._getAddressBook().getSelectedItem()

    # Table getter methods
    def getTable(self):
        return self._getTable().getSelectedItem()

    # Column getter methods
    def getColumn(self):
        return self._getColumn().getSelectedItem()

    # Query getter methods
    def getQuery(self):
        return self._getQuery().getText().strip()

    def getSubQuery(self):
        subquery = None
        control = self._getQuery()
        item = control.getText().strip()
        items = control.getItems()
        if item in items:
            subquery = control.Model.getItemData(items.index(item))
        return subquery

    def isQuerySelected(self):
        control = self._getQuery()
        item = control.getText()
        items = control.getItems()
        return item in items

    def getQueries(self):
        subquery = None
        control = self._getQuery()
        query = control.getText().strip()
        queries = control.getItems()
        if query in queries:
            subquery = control.Model.getItemData(queries.index(query))
        return query, subquery

    def isTableSelected(self):
        return self._getTable().getSelectedItemPos() != -1

    # Email getter method
    def getEmail(self):
        return self._getEmail().getSelectedItem()

    def getEmails(self):
        return self._getEmail().Model.StringItemList

    def getEmailPosition(self):
        return self._getEmail().getSelectedItemPos()

    def hasEmail(self):
        return self._getEmail().getItemCount() > 0

    # Identifier getter method
    def getIdentifier(self):
        return self._getIdentifier().getSelectedItem()

    def getIdentifiers(self):
        return self._getIdentifier().Model.StringItemList

    def getIdentifierPosition(self):
        return self._getIdentifier().getSelectedItemPos()

    def hasIdentifier(self):
        return self._getIdentifier().getItemCount() > 0

    def getIdentifierCount(self):
        return self._getIdentifier().getItemCount()

# MergerView setter methods
    def setPageStep(self, step):
        self._window.Model.Step = step

    def enablePage(self, enabled):
        self.enableAddressBook(enabled)
        self._enableBox(enabled)

    def enableAddressBook(self, enabled):
        self._getAddressBook().Model.Enabled = enabled
        self._getNewAddressBook().Model.Enabled = enabled

    def enableButton(self, enabled):
        self._getAddQuery().Model.Enabled = enabled
        self._getRemoveQuery().Model.Enabled = enabled
        self._getAddEmail().Model.Enabled = enabled
        self._getRemoveEmail().Model.Enabled = enabled
        self._getUpEmail().Model.Enabled = enabled
        self._getDownEmail().Model.Enabled = enabled
        self._getAddIdentifier().Model.Enabled = enabled
        self._getRemoveIdentifier().Model.Enabled = enabled

    def updateProgress(self, value, message):
        if not self.isDisposed():
            self._getProgressBar().Value = value
            self._getProgressMessage().Text = message
        else:
            print("MergerView.updateProgress() ERROR *********************")

    def setMessageText(self, text):
        self.enableAddressBook(True)
        self._enableBox(False)
        self.enableButton(False)
        self._getMessage().Text = text

    def initAddressBook(self, addressbooks):
        self._getAddressBook().Model.StringItemList = addressbooks

    def selectAddressBook(self, addressbook):
        self._getAddressBook().selectItem(addressbook, True)

    def initTables(self, tables):
        self._getTable().Model.StringItemList = tables

    def setDefaultTable(self):
        self._getTable().selectItemPos(0, True)

    def setTable(self, table):
        self._getTable().selectItem(table, True)

    def setColumnLabel(self, text):
        self._getColumnLabel().Text = text

    def initColumns(self, columns):
        control= self._getColumn()
        # TODO: We need to reset the control in order the handler been called
        control.Model.removeAllItems()
        control.Model.StringItemList = columns
        if control.getItemCount() > 0:
            control.selectItemPos(0, True)

    def initQuery(self, queries):
        control = self._getQuery()
        control.Model.removeAllItems()
        for query in queries:
            index = control.Model.ItemCount
            control.Model.insertItemText(index, query.Name)
            control.Model.setItemData(index, query.Value)
        if queries:
            control.Text = control.Model.getItemText(0)

    # Query methods
    def enableAddQuery(self, enabled):
        self._getAddQuery().Model.Enabled = enabled

    def enableRemoveQuery(self, enabled):
        self._getRemoveQuery().Model.Enabled = enabled

    def setQuery(self, query):
        self._getQuery().setText(query)

    def addQuery(self, query, subquery):
        control = self._getQuery()
        control.setText('')
        index = control.Model.ItemCount
        control.Model.insertItemText(index, query)
        control.Model.setItemData(index, subquery)

    def removeQuery(self, query):
        self._getRemoveQuery().Model.Enabled = False
        control = self._getQuery()
        queries = control.getItems()
        if query in queries:
            control.Model.removeItem(queries.index(query))
        control.setText('')

    # Email column setter methods
    def disableEmailButton(self):
        self._getUpEmail().Model.Enabled = False
        self._getDownEmail().Model.Enabled = False
        self._getRemoveEmail().Model.Enabled = False

    def enableAddEmail(self, enabled):
        self._getAddEmail().Model.Enabled = enabled

    def updateAddEmail(self, enabled):
        items = self.getEmails()
        button = self._getAddEmail()
        self._updateAddButton(items, button, enabled)

    def enableRemoveEmail(self, enabled):
        self._getRemoveEmail().Model.Enabled = enabled

    def enableUpEmail(self, enabled):
        self._getUpEmail().Model.Enabled = enabled

    def enableDownEmail(self, enabled):
        self._getDownEmail().Model.Enabled = enabled

    def setEmail(self, emails, index=-1):
        control = self._getEmail()
        control.Model.StringItemList = emails
        if index != -1:
            control.selectItemPos(index, True)

    # Identifier column methods
    def disableIdentifierButton(self):
        self._getUpIdentifier().Model.Enabled = False
        self._getDownIdentifier().Model.Enabled = False
        self._getRemoveIdentifier().Model.Enabled = False

    def enableAddIdentifier(self, enabled):
        self._getAddIdentifier().Model.Enabled = enabled

    def updateAddIdentifier(self, enabled):
        items = self.getIdentifiers()
        button = self._getAddIdentifier()
        self._updateAddButton(items, button, enabled)

    def enableRemoveIdentifier(self, enabled):
        self._getRemoveIdentifier().Model.Enabled = enabled

    def enableUpIdentifier(self, enabled):
        self._getUpIdentifier().Model.Enabled = enabled

    def enableDownIdentifier(self, enabled):
        self._getDownIdentifier().Model.Enabled = enabled

    def setIdentifier(self, identifiers, index=-1):
        control = self._getIdentifier()
        control.Model.StringItemList = identifiers
        if index != -1:
            control.selectItemPos(index, True)

# MergerView private setter methods
    def _updateAddButton(self, items, button, enabled):
        if enabled:
            column = self.getColumn()
            enabled = column not in items
        button.Model.Enabled = enabled

    def _enableBox(self, enabled):
        control = self._getQuery()
        self._enableComboBox(control, enabled)
        control = self._getColumn()
        self._enableListBox(control, enabled)
        control = self._getTable()
        self._enableListBox(control, enabled)
        control = self._getEmail()
        self._enableListBox(control, enabled)
        control = self._getIdentifier()
        self._enableListBox(control, enabled)

    def _enableComboBox(self, control, enabled):
        self._enableListBox(control, enabled)
        if not enabled:
            control.setText('')

    def _enableListBox(self, control, enabled):
        control.Model.Enabled = enabled
        if not enabled:
            control.Model.removeAllItems()

# MergerView private getter control methods
    def _getAddressBook(self):
        return self._window.getControl('ListBox1')

    def _getTable(self):
        return self._window.getControl('ListBox2')

    def _getColumn(self):
        return self._window.getControl('ListBox3')

    def _getEmail(self):
        return self._window.getControl('ListBox4')

    def _getIdentifier(self):
        return self._window.getControl('ListBox5')

    def _getQuery(self):
        return self._window.getControl('ComboBox1')

    def _getProgressBar(self):
        return self._window.getControl('ProgressBar1')

    def _getNewAddressBook(self):
        return self._window.getControl('CommandButton1')

    def _getAddQuery(self):
        return self._window.getControl('CommandButton2')

    def _getRemoveQuery(self):
        return self._window.getControl('CommandButton3')

    def _getAddEmail(self):
        return self._window.getControl('CommandButton4')

    def _getUpEmail(self):
        return self._window.getControl('CommandButton5')

    def _getDownEmail(self):
        return self._window.getControl('CommandButton6')

    def _getRemoveEmail(self):
        return self._window.getControl('CommandButton7')

    def _getAddIdentifier(self):
        return self._window.getControl('CommandButton8')

    def _getUpIdentifier(self):
        return self._window.getControl('CommandButton9')

    def _getDownIdentifier(self):
        return self._window.getControl('CommandButton10')

    def _getRemoveIdentifier(self):
        return self._window.getControl('CommandButton11')

    def _getProgressMessage(self):
        return self._window.getControl('Label6')

    def _getMessage(self):
        return self._window.getControl('Label8')

    def _getColumnLabel(self):
        return self._window.getControl('Label14')
