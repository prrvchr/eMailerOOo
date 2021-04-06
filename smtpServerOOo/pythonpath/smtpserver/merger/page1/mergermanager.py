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

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import createService
from smtpserver import executeDispatch
from smtpserver import logMessage
from smtpserver import getMessage

from .mergerview import MergerView

import traceback


class MergerManager(unohelper.Base,
                    XWizardPage):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._disabled = False
        addressbooks = self._model.getAvailableAddressBooks()
        self._view = MergerView(ctx, self, parent, addressbooks)
        addressbook = self._model.getDocumentAddressBook()
        if addressbook in addressbooks:
            self._view.setPageStep(1)
            # TODO: We must disable the handler "ChangeAddressBook" otherwise it activates twice
            self._disableHandler()
            self._view.selectAddressBook(addressbook)
        else:
            self._view.enableAddressBook(True)

    @property
    def Model(self):
        return self._model

    # TODO: One shot disabler handler
    def isHandlerEnabled(self):
        enabled = True
        if self._disabled:
            self._disabled = enabled = False
        return enabled
    def _disableHandler(self):
        self._disabled = True

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        pass

    def commitPage(self, reason):
        try:
            query = self._view.getQuery()
            self._model.setQuery(query)
            updated = self._model.isUpdated()
            print("MergerManager.commitPage() %s" % updated)
            if updated:
                self._model.updateRecipient()
            return True
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

    def canAdvance(self):
        return all((self._view.isQuerySelected(),
                    self._view.hasEmail(),
                    self._view.hasIndex()))

# MergerManager setter methods
    # AddressBook setter methods
    def changeAddressBook(self, addressbook):
        self._view.enablePage(False)
        self._view.enableButton(False)
        self._view.setPageStep(1)
        self._model.setAddressBook(addressbook, self.progress, self.setAddressBook)

    def progress(self, value):
        self._view.updateProgress(value)

    def setAddressBook(self, step, queries, tables, label1, label2, msg):
        if step == 2:
            self._view.setMessageText(msg)
        elif step == 3:
            # TODO: We must disable the handler "EditQuery" otherwise it activates twice
            self._disableHandler()
            self._view.initQuery(queries)
            # TODO: We must disable the handler "ChangeAddressBookTable" otherwise it activates twice
            self._disableHandler()
            self._view.initTables(tables)
            self._view.setEmailLabel(label1)
            self._view.setIndexLabel(label2)
            self._view.enablePage(True)
        self._view.setPageStep(step)
        self._wizard.updateTravelUI()

    def newAddressBook(self):
        url = '.uno:AutoPilotAddressDataSource'
        executeDispatch(self._ctx, url)
        #url = '.uno:AddressBookSource'
        # TODO: Update the list of AddressBooks and keep the selection if possible
        addressbook = self._view.getAddressBook()
        addressbooks = self._model.getAvailableAddressBooks()
        self._view.initAddressBook(addressbooks)
        if addressbook in addressbooks:
            # TODO: We must disable the handler "ChangeAddressBook" otherwise it activates twice
            self._disableHandler()
            self._view.selectAddressBook(addressbook)

    # AddressBook Table setter methods
    def changeAddressBookTable(self, table):
        columns, emails = self._model.setAddressBookTable(table)
        # TODO: We must disable the handler "ChangeAddressBookColumn" otherwise it activates twice
        self._disableHandler()
        self._view.initColumns(columns)
        self._view.setEmail(emails)

    # AddressBook Column setter methods
    def changeAddressBookColumn(self, column):
        emails = self._view.getEmails()
        enabled = column not in emails
        self._view.enableAddEmail(enabled)
        indexes = self._view.getIndexes()
        enabled = all((self._view.isQuerySelected(),
                       column not in indexes))
        self._view.enableAddIndex(enabled)

    # Query setter methods
    def editQuery(self, query, exist):
        indexes = self._model.getQueryIndex(query)
        self._setQuery(indexes, exist)
        enabled = self._model.validateQuery(query, exist)
        self._view.enableAddQuery(enabled)
        self._view.enableAddIndex(exist)
        self._wizard.updateTravelUI()

    def changeQuery(self, query):
        indexes = self._model.getQueryIndex(query)
        self._setQuery(indexes, True)
        column = self._view.getColumn()
        enabled = column not in indexes
        self._view.enableAddIndex(enabled)
        self._wizard.updateTravelUI()

    def enterQuery(self, query, exist):
        if self._model.validateQuery(query, exist):
            self._addQuery(query)
            self._wizard.updateTravelUI()

    def addQuery(self):
        query = self._view.getQuery()
        self._addQuery(query)
        self._wizard.updateTravelUI()

    def removeQuery(self):
        query = self._view.getQuery()
        self._model.removeQuery(query)
        self._view.removeQuery(query)
        self._wizard.updateTravelUI()

    # Email column setter methods
    def changeEmail(self, imax, position):
        self._view.enableRemoveEmail(position != -1)
        self._view.enableBefore(position > 0)
        self._view.enableAfter(-1 < position < imax)

    def addEmail(self):
        self._view.enableAddEmail(False)
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        table = self._view.getTable()
        email = self._view.getColumn()
        emails = self._model.addEmail(table, email)
        self._view.setEmail(emails)
        self._wizard.updateTravelUI()

    def removeEmail(self):
        self._view.enableAddEmail(False)
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        table = self._view.getTable()
        email = self._view.getEmail()
        emails = self._model.removeEmail(table, email)
        self._view.setEmail(emails)
        column = self._view.getColumn()
        if column not in emails:
            self._view.enableAddEmail(True)
        self._wizard.updateTravelUI()

    def moveBefore(self):
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        table = self._view.getTable()
        email = self._view.getEmail()
        index = self._view.getEmailPosition() -1
        emails = self._model.moveEmail(table, email, index)
        self._view.setEmail(emails, index)

    def moveAfter(self):
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        table = self._view.getTable()
        email = self._view.getEmail()
        index = self._view.getEmailPosition() +1
        emails = self._model.moveEmail(table, email, index)
        self._view.setEmail(emails, index)

    # Index column setter methods
    def changeIndex(self, enabled):
        self._view.enableRemoveIndex(enabled)

    def addIndex(self):
        self._view.enableAddIndex(False)
        self._view.enableRemoveIndex(False)
        query = self._view.getQuery()
        index = self._view.getColumn()
        indexes = self._model.addIndex(query, index)
        self._view.setIndexes(indexes)
        self._wizard.updateTravelUI()

    def removeIndex(self):
        self._view.enableRemoveIndex(False)
        self._view.enableAddIndex(False)
        query = self._view.getQuery()
        index = self._view.getIndex()
        indexes = self._model.removeIndex(query, index)
        self._view.setIndexes(indexes)
        column = self._view.getColumn()
        if column not in indexes:
            self._view.enableAddIndex(True)
        self._wizard.updateTravelUI()

# MergerManager private setter methods
    # Query private setter methods
    def _setQuery(self, indexes, exist):
        self._view.setIndexes(indexes)
        self._view.enableRemoveQuery(exist)

    def _addQuery(self, query):
        table = self._view.getTable()
        self._model.addQuery(table, query)
        self._view.addQuery(query)
        self._view.enableAddIndex(False)
        self._view.enableRemoveQuery(False)
