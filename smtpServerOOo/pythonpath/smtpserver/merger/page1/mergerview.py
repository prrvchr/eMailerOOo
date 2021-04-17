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

from smtpserver import getContainerWindow
from smtpserver import logMessage
from smtpserver import g_extension

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

    def isQuerySelected(self):
        selected = False
        control = self._getQuery()
        if control.getItemCount() > 0:
            query = control.getText()
            queries = control.getItems()
            selected = query in queries
        return selected

    # Email getter method
    def getEmail(self):
        return self._getEmail().getSelectedItem()

    def getEmails(self):
        return self._getEmail().Model.StringItemList

    def getEmailPosition(self):
        return self._getEmail().getSelectedItemPos()

    def hasEmail(self):
        return self._getEmail().getItemCount() > 0

    # Index getter method
    def getIndex(self):
        return self._getIndex().getItem(0)

    #def getIndexes(self):
    #    return self._getIndex().Model.StringItemList

    def hasIndex(self):
        return self._getIndex().getItemCount() > 0

# MergerView setter methods
    def dispose(self):
        self._window.dispose()
        self._window = None

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
        self._getAddIndex().Model.Enabled = enabled
        self._getRemoveIndex().Model.Enabled = enabled

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
        control = self._getTable()
        control.Model.StringItemList = tables
        if control.getItemCount() > 0:
            control.selectItemPos(0, True)

    def setIndexLabel(self, text):
        self._getIndexLabel().Text = text

    def initColumns(self, columns):
        control= self._getColumn()
        control.Model.StringItemList = columns
        if control.getItemCount() > 0:
            control.selectItemPos(0, True)

    def initQuery(self, queries):
        control = self._getQuery()
        control.Model.StringItemList = queries
        if control.getItemCount() > 0:
            query = control.getItem(0)
            control.setText(query)
        else:
            control.setText('')

    # Query methods
    def enableAddQuery(self, enabled):
        self._getAddQuery().Model.Enabled = enabled

    def enableRemoveQuery(self, enabled):
        self._getRemoveQuery().Model.Enabled = enabled

    def addQuery(self, query):
        control = self._getQuery()
        control.setText('')
        count = control.getItemCount()
        control.addItem(query, count)
        #self._getRemoveQuery().Model.Enabled = False

    def removeQuery(self, query):
        self._getRemoveQuery().Model.Enabled = False
        control = self._getQuery()
        query = control.getText()
        queries = control.getItems()
        if query in queries:
            control.setText('')
            position = queries.index(query)
            control.removeItems(position, 1)

    # Email column setter methods
    def enableAddEmail(self, enabled):
        self._getAddEmail().Model.Enabled = enabled

    def enableRemoveEmail(self, enabled):
        self._getRemoveEmail().Model.Enabled = enabled

    def enableBefore(self, enabled):
        self._getBefore().Model.Enabled = enabled

    def enableAfter(self, enabled):
        self._getAfter().Model.Enabled = enabled

    def setEmail(self, emails, index=None):
        control = self._getEmail()
        control.Model.StringItemList = emails
        if index is not None:
            control.selectItemPos(index, True)

    # Index column methods
    def setIndex(self, index, add):
        control = self._getIndex()
        if control.getItemCount() > 0:
            if index is None:
                self._removeIndex(control, add)
            else:
                self.enableAddIndex(False)
                control.Model.setItemText(0, index)
                self.enableRemoveIndex(True)
        elif index is not None:
            self._addIndex(control, index)
        else:
            self.enableAddIndex(add)
            self.enableRemoveIndex(False)

    def addIndex(self, index):
        control = self._getIndex()
        self._addIndex(control, index)

    def removeIndex(self):
        control = self._getIndex()
        self._removeIndex(control, True)

    def enableAddIndex(self, enabled):
        self._getAddIndex().Model.Enabled = enabled

    def enableRemoveIndex(self, enabled):
        self._getRemoveIndex().Model.Enabled = enabled

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
        control = self._getIndex()
        self._enableListBox(control, enabled)

    def _enableComboBox(self, control, enabled):
        self._enableListBox(control, enabled)
        if not enabled:
            control.setText('')

    def _enableListBox(self, control, enabled):
        control.Model.Enabled = enabled
        if not enabled:
            control.Model.StringItemList = ()

    def _addIndex(self, control, index):
        self.enableAddIndex(False)
        control.Model.insertItemText(0, index)
        self.enableRemoveIndex(True)

    def _removeIndex(self, control, add):
        self.enableRemoveIndex(False)
        control.Model.removeItem(0)
        self.enableAddIndex(add)

# MergerView private getter control methods
    def _getAddressBook(self):
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

    def _getAddIndex(self):
        return self._window.getControl('CommandButton8')

    def _getRemoveIndex(self):
        return self._window.getControl('CommandButton9')

    def _getProgressMessage(self):
        return self._window.getControl('Label6')

    def _getMessage(self):
        return self._window.getControl('Label8')

    def _getIndexLabel(self):
        return self._window.getControl('Label13')
