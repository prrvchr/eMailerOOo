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

from smtpmailer import createService
from smtpmailer import executeDispatch
from smtpmailer import logMessage
from smtpmailer import getMessage

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
        addressbook = self._model.getDefaultAddressBook()
        if addressbook in addressbooks:
            self._view.setPageStep(1)
            # FIXME: We must disable the "ChangeAddressBook"
            # FIXME: handler otherwise it activates twice
            self._disableHandler()
            self._view.selectAddressBook(addressbook)
        else:
            self._view.enableAddressBook(True)

    # FIXME: One shot disabler handler
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
            table = self._view.getQueryTable()
            self._model.commitPage1(query, table)
            return True
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

    def canAdvance(self):
        return all((self._view.hasEmail(),
                    self._view.hasIdentifier(),
                    self._view.hasBookmark()))

# MergerManager setter methods
    # AddressBook setter methods
    def changeAddressBook(self, addressbook):
        if self._model.isAddressBookNotLoaded(addressbook):
            self._view.enablePage(False)
            self._view.enableButton(False)
            self._view.setPageStep(1)
            message = self._model.getProgressMessage(0)
            self._view.updateProgress(0, message)
            self._model.setAddressBook(addressbook, self.progress, self.setAddressBook)

    def progress(self, value):
        message = self._model.getProgressMessage(value)
        self._view.updateProgress(value, message)

    def setAddressBook(self, step, queries, tables, label, msg):
        if step == 2:
            self._view.setMessageText(msg)
        elif step == 3:
            self._view.enablePage(True)
            self._view.setColumnLabel(label)
            self._view.initTables(tables)
            if self._view.initQuery(queries) and tables:
                # FIXME: We must disable the "ChangeAddressBookTable"
                # FIXME: handler otherwise it activates twice
                self._disableHandler()
                self._view.initTablesSelection()
        self._view.setPageStep(step)
        self._wizard.updateTravelUI()

    def newAddressBook(self):
        url = '.uno:AutoPilotAddressDataSource'
        executeDispatch(self._ctx, url)
        # url = '.uno:AddressBookSource'
        # Update the list of AddressBooks and keep the selection if possible
        addressbook = self._view.getAddressBook()
        addressbooks = self._model.getAvailableAddressBooks()
        self._view.initAddressBook(addressbooks)
        if addressbook in addressbooks:
            # FIXME: We must disable the "ChangeAddressBook"
            # FIXME: handler otherwise it activates twice
            self._disableHandler()
            self._view.selectAddressBook(addressbook)

    # Table setter methods
    def changeAddressBookTable(self, table):
        columns = self._model.setAddressBookTable(table)
        # FIXME: We must disable the "ChangeAddressBookColumn"
        # FIXME: handler otherwise it activates twice
        self._disableHandler()
        #self._view.updateAddQuery()
        self._view.initColumns(columns)

    # Column setter methods
    def changeAddressBookColumn(self, column):
        enabled = False
        query = self._view.getQueryTable()
        if query is not None:
            table = self._view.getTable()
            enabled = self._model.isSimilar() or query == table
        emails = self._view.getEmails()
        self._view.updateAddEmail(emails, enabled)
        self._view.enableAddIdentifier(enabled)
        self._view.enableAddBookmark(enabled)

    # Query setter methods
    def editQuery(self, query, table):
        exist = table is not None
        if exist:
            self._model.setQuery(query)
            if self._view.getTable() != table:
                # FIXME: We must disable the "ChangeAddressBookTable"
                # FIXME: handler otherwise it activates twice
                self._disableHandler()
                self._view.setTable(table)
            enabled = False
            identifier = self._model.getIdentifier()
            bookmark = self._model.getBookmark()
            emails = self._model.getEmails()
        else:
            enabled = self._model.isQueryValid(query)
            identifier = None
            bookmark = None
            emails = ()
        self._view.enableAddQuery(enabled)
        self._view.enableRemoveQuery(exist)
        self._view.setBookmark(bookmark, exist)
        self._view.setIdentifier(identifier, exist)
        self._view.setEmail(emails)
        self._view.updateAddEmail(emails, exist)
        self._view.enableRemoveEmail(False)
        self._wizard.updateTravelUI()

    def enterQuery(self, query):
        if self._model.isQueryValid(query):
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

    # Query private setter methods
    def _addQuery(self, query):
        table = self._view.getTable()
        self._model.addQuery(query, table)
        self._view.addQuery(query, table)
        self._view.enableRemoveQuery(False)

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
        query = self._view.getQuery()
        email = self._view.getColumn()
        emails = self._model.addEmail(query, email)
        self._view.setEmail(emails)
        self._wizard.updateTravelUI()

    def removeEmail(self):
        self._view.enableAddEmail(False)
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        query = self._view.getQuery()
        email = self._view.getEmail()
        emails = self._model.removeEmail(query, email)
        enabled = self._canAddColumn()
        self._view.setEmail(emails)
        self._view.updateAddEmail(emails, enabled)
        self._wizard.updateTravelUI()

    def moveBefore(self):
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        query = self._view.getQuery()
        email = self._view.getEmail()
        position = self._view.getEmailPosition() -1
        emails = self._model.moveEmail(query, email, position)
        self._view.setEmail(emails, position)

    def moveAfter(self):
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        query = self._view.getQuery()
        email = self._view.getEmail()
        position = self._view.getEmailPosition() +1
        emails = self._model.moveEmail(query, email, position)
        self._view.setEmail(emails, position)

    # Identifier column setter methods
    def addIdentifier(self):
        self._view.enableAddIdentifier(False)
        self._view.enableRemoveIdentifier(False)
        query = self._view.getQuery()
        identifier = self._view.getColumn()
        self._model.addIdentifier(query, identifier)
        self._view.addIdentifier(identifier)
        self._wizard.updateTravelUI()

    def removeIdentifier(self):
        self._view.enableAddIdentifier(False)
        self._view.enableRemoveIdentifier(False)
        query = self._view.getQuery()
        identifier = self._view.getIdentifier()
        self._model.removeIdentifier(query, identifier)
        enabled = self._canAddColumn()
        self._view.removeIdentifier(enabled)
        self._wizard.updateTravelUI()

    # Bookmark column setter methods
    def addBookmark(self):
        self._view.enableAddBookmark(False)
        self._view.enableRemoveBookmark(False)
        query = self._view.getQuery()
        bookmark = self._view.getColumn()
        self._model.addBookmark(query, bookmark)
        self._view.addBookmark(bookmark)
        self._wizard.updateTravelUI()

    def removeBookmark(self):
        self._view.enableAddBookmark(False)
        self._view.enableRemoveBookmark(False)
        query = self._view.getQuery()
        bookmark = self._view.getBookmark()
        self._model.removeBookmark(query, bookmark)
        enabled = self._canAddColumn()
        self._view.removeBookmark(enabled)
        self._wizard.updateTravelUI()

    def _canAddColumn(self):
        return self._model.isSimilar() or self._isSameTable()

    def _isSameTable(self):
        table = self._view.getTable()
        query = self._view.getQueryTable()
        return table == query
