#!
# -*- coding: utf-8 -*-

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

from smtpmailer import getContainerWindow
from smtpmailer import logMessage
from smtpmailer import g_extension

from .mergerhandler import WindowHandler

import traceback


class MergerView(unohelper.Base):
    def __init__(self, ctx, manager, parent, addressbooks):
        handler = WindowHandler(manager)
        self._window = getContainerWindow(ctx, parent, handler, g_extension, 'MergerPage1')
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

    def getQueryTable(self):
        table = None
        control = self._getQuery()
        item = control.getText()
        items = control.getItems()
        if item in items:
            index = items.index(item)
            table = control.Model.getItemData(index)
        return table

    def isQuerySelected(self):
        control = self._getQuery()
        item = control.getText()
        items = control.getItems()
        return item in items

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
        return self._getIdentifier().getItem(0)

    def hasIdentifier(self):
        return self._getIdentifier().getItemCount() > 0

    # Bookmark getter method
    def getBookmark(self):
        return self._getBookmark().getItem(0)

    def hasBookmark(self):
        return self._getBookmark().getItemCount() > 0

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
        self._getBefore().Model.Enabled = enabled
        self._getAfter().Model.Enabled = enabled
        self._getAddIdentifier().Model.Enabled = enabled
        self._getRemoveIdentifier().Model.Enabled = enabled
        self._getAddBookmark().Model.Enabled = enabled
        self._getRemoveBookmark().Model.Enabled = enabled

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

    def initTablesSelection(self):
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
        for query, table in queries.items():
            index = control.Model.ItemCount
            control.Model.insertItemText(index, query)
            control.Model.setItemData(index, table)
        query = ''
        initialized = control.Model.ItemCount > 0
        if initialized:
            query = control.Model.getItemText(0)
        control.Text = query
        return not initialized

    # Query methods
    def enableAddQuery(self, enabled):
        self._getAddQuery().Model.Enabled = enabled

    def enableRemoveQuery(self, enabled):
        self._getRemoveQuery().Model.Enabled = enabled

    def addQuery(self, query, table):
        control = self._getQuery()
        control.setText('')
        index = control.Model.ItemCount
        control.Model.insertItemText(index, query)
        control.Model.setItemData(index, table)

    def removeQuery(self, query):
        self._getRemoveQuery().Model.Enabled = False
        control = self._getQuery()
        query = control.getText()
        queries = control.getItems()
        if query in queries:
            control.setText('')
            index = queries.index(query)
            control.removeItems(index, 1)

    # Email column setter methods
    def enableAddEmail(self, enabled):
        self._getAddEmail().Model.Enabled = enabled

    def updateAddEmail(self, emails, enabled):
        if enabled:
            column = self.getColumn()
            enabled = column not in emails
        self.enableAddEmail(enabled)

    def enableRemoveEmail(self, enabled):
        self._getRemoveEmail().Model.Enabled = enabled

    def enableBefore(self, enabled):
        self._getBefore().Model.Enabled = enabled

    def enableAfter(self, enabled):
        self._getAfter().Model.Enabled = enabled

    def setEmail(self, emails, index=-1):
        control = self._getEmail()
        control.Model.StringItemList = emails
        if index != -1:
            control.selectItemPos(index, True)

    # Identifier column methods
    def setIdentifier(self, identifier, exist):
        identifiers = () if identifier is None else (identifier, )
        add = identifier is None if exist else False
        remove = identifier is not None if exist else False
        self._getIdentifier().Model.StringItemList = identifiers
        self.enableAddIdentifier(add)
        self.enableRemoveIdentifier(remove)

    def addIdentifier(self, identifier):
        self._getIdentifier().Model.insertItemText(0, identifier)
        self.enableRemoveIdentifier(True)

    def removeIdentifier(self, enabled):
        self._getIdentifier().Model.removeItem(0)
        self.enableAddIdentifier(enabled)

    def enableAddIdentifier(self, enabled):
        enabled = not self.hasIdentifier() if enabled else False
        self._getAddIdentifier().Model.Enabled = enabled

    def enableRemoveIdentifier(self, enabled):
        self._getRemoveIdentifier().Model.Enabled = enabled

    # Bookmark column methods
    def setBookmark(self, bookmark, exist):
        bookmarks = () if bookmark is None else (bookmark, )
        add = bookmark is None if exist else False
        remove = bookmark is not None if exist else False
        self._getBookmark().Model.StringItemList = bookmarks
        self.enableAddBookmark(add)
        self.enableRemoveBookmark(remove)

    def addBookmark(self, bookmark):
        self._getBookmark().Model.insertItemText(0, bookmark)
        self.enableRemoveBookmark(True)

    def removeBookmark(self, enabled):
        self._getBookmark().Model.removeItem(0)
        self.enableAddBookmark(enabled)

    def enableAddBookmark(self, enabled):
        enabled = not self.hasBookmark() if enabled else False
        self._getAddBookmark().Model.Enabled = enabled

    def enableRemoveBookmark(self, enabled):
        self._getRemoveBookmark().Model.Enabled = enabled

# MergerView private setter methods
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
        control = self._getBookmark()
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

    def _getBookmark(self):
        return self._window.getControl('ListBox6')

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

    def _getRemoveEmail(self):
        return self._window.getControl('CommandButton5')

    def _getBefore(self):
        return self._window.getControl('CommandButton6')

    def _getAfter(self):
        return self._window.getControl('CommandButton7')

    def _getAddIdentifier(self):
        return self._window.getControl('CommandButton8')

    def _getRemoveIdentifier(self):
        return self._window.getControl('CommandButton9')

    def _getAddBookmark(self):
        return self._window.getControl('CommandButton10')

    def _getRemoveBookmark(self):
        return self._window.getControl('CommandButton11')

    def _getProgressMessage(self):
        return self._window.getControl('Label6')

    def _getMessage(self):
        return self._window.getControl('Label8')

    def _getColumnLabel(self):
        return self._window.getControl('Label14')
