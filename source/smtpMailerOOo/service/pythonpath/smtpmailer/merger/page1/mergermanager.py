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

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .mergerview import MergerView
from .mergerhandler import WindowHandler

from ...unotool import createMessageBox
from ...unotool import createService
from ...unotool import executeDispatch

from ...logger import getMessage
from ...logger import logMessage

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
        self._view = MergerView(ctx, WindowHandler(self), parent, addressbooks)
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
            self._model.commitPage1()
            return True
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

    def canAdvance(self):
        return self._view.hasEmail() and self._view.hasIdentifier()

# MergerManager setter methods
    # Methods called by MergerHandler
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
            self._view.initQuery(queries)
            if len(queries) > 0:
                self._view.setDefaultQuery()
            elif len(tables):
                # FIXME: We must disable the "ChangeAddressBookTable"
                # FIXME: handler otherwise it activates twice
                self._disableHandler()
                self._view.setDefaultTable()
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
        subquery = self._view.getSubQuery()
        if subquery is not None:
            enabled = self._model.isSimilar() or subquery.Second == self._view.getTable()
        self._view.updateAddEmail(enabled)
        self._view.updateAddIdentifier(enabled)

    # Query setter methods
    def editQuery(self, query, subquery, exist):
        table = self._view.getTable()
        self._model.setQuery(query, subquery, exist, table, self.setQuery)

    def setQuery(self, identifiers, emails, exist, table, enabled):
        if self._view.getTable() != table:
            # FIXME: We must disable the "ChangeAddressBookTable"
            # FIXME: handler otherwise it activates twice
            self._disableHandler()
            self._view.setTable(table)
        self._view.enableAddQuery(enabled)
        self._view.enableRemoveQuery(exist)
        self._view.setEmail(emails)
        self._view.setIdentifier(identifiers)
        self._view.updateAddEmail(exist)
        self._view.updateAddIdentifier(exist)
        self._view.enableRemoveEmail(False)
        self._wizard.updateTravelUI()

    def enterQuery(self, query):
        if self._view.isTableSelected() and self._model.isQueryValid(query):
            self._addQuery(query)
            #self._wizard.updateTravelUI()

    def addQuery(self):
        query = self._view.getQuery()
        self._addQuery(query)
        #self._wizard.updateTravelUI()

    def removeQuery(self):
        query, subquery = self._view.getQueries()
        self._model.removeQuery(query, subquery)
        self._view.removeQuery(query)
        self._wizard.updateTravelUI()

    # Query private setter methods
    def _addQuery(self, query):
        table = self._view.getTable()
        subquery = self._model.addQuery(query, table)
        self._view.addQuery(query, subquery)
        self._view.enableRemoveQuery(False)
        self._view.setQuery(query)

    # Email column setter methods
    def changeEmail(self, imax, position):
        self._view.enableRemoveEmail(position != -1)
        self._view.enableUpEmail(position > 0)
        self._view.enableDownEmail(-1 < position < imax)

    def addEmail(self):
        self._view.enableAddEmail(False)
        self._view.disableEmailButton()
        email = self._view.getColumn()
        emails = self._model.addEmail(email)
        self._view.setEmail(emails)
        self._wizard.updateTravelUI()

    def removeEmail(self):
        # FIXME: If we remove an Email column to make sure there is
        # FIXME: no orphan filter, we need to remove all filters
        query = self._view.getQuery()
        count = self._model.getFilterCount()
        if count > 1 and self._cancelAction(query):
            return
        self._view.enableAddEmail(False)
        self._view.disableEmailButton()
        email = self._view.getEmail()
        emails = self._model.removeEmail(email, count > 1)
        enabled = self._canAddColumn()
        self._view.setEmail(emails)
        self._view.updateAddEmail(emails, enabled)
        self._wizard.updateTravelUI()

    def upEmail(self):
        self._view.disableEmailButton()
        email = self._view.getEmail()
        position = self._view.getEmailPosition() -1
        emails = self._model.moveEmail(email, position)
        self._view.setEmail(emails, position)

    def downEmail(self):
        self._view.disableEmailButton()
        email = self._view.getEmail()
        position = self._view.getEmailPosition() +1
        emails = self._model.moveEmail(email, position)
        self._view.setEmail(emails, position)

    # Identifier column setter methods
    def changeIdentifier(self, imax, position):
        self._view.enableRemoveIdentifier(position != -1)
        self._view.enableUpIdentifier(position > 0)
        self._view.enableDownIdentifier(-1 < position < imax)

    def addIdentifier(self):
        count = self._model.getFilterCount()
        if count > 1 and self._cancelAction():
            return
        self._view.enableAddIdentifier(False)
        self._view.disableIdentifierButton()
        identifier = self._view.getColumn()
        first = self._view.getIdentifierCount() == 0 
        identifiers = self._model.addIdentifier(identifier, first or count > 0)
        self._view.setIdentifier(identifiers)
        self._wizard.updateTravelUI()

    def removeIdentifier(self):
        count = self._model.getFilterCount()
        last = self._view.getIdentifierCount() == 1
        if (count > 1 or last) and self._cancelAction():
            return
        self._view.enableAddIdentifier(False)
        self._view.disableIdentifierButton()
        identifier = self._view.getIdentifier()
        identifiers = self._model.removeIdentifier(identifier, count > 0)
        enabled = self._canAddColumn()
        enabled = self._canAddColumn()
        self._view.setIdentifier(identifiers)
        self._view.updateAddIdentifier(enabled)
        self._wizard.updateTravelUI()

    def upIdentifier(self):
        count = self._model.getFilterCount()
        if count > 1 and self._cancelAction():
            return
        self._view.disableIdentifierButton()
        identifier = self._view.getIdentifier()
        position = self._view.getIdentifierPosition() -1
        identifiers = self._model.moveIdentifier(identifier, count > 0, position)
        self._view.setIdentifier(identifiers, position)
        self._wizard.updateTravelUI()

    def downIdentifier(self):
        count = self._model.getFilterCount()
        if count > 1 and self._cancelAction():
            return
        self._view.disableIdentifierButton()
        identifier = self._view.getIdentifier()
        position = self._view.getIdentifierPosition() +1
        identifiers = self._model.moveIdentifier(identifier, count > 0, position)
        self._view.setIdentifier(identifiers, position)
        self._wizard.updateTravelUI()

    def _cancelAction(self):
        query = self._view.getQuery()
        parent = self._view.getWindow().Peer
        message, title = self._model.getMessageBoxData(query)
        dialog = createMessageBox(parent, message, title, 'warning')
        status = dialog.execute()
        dialog.dispose()
        return status != OK

    def _canAddColumn(self):
        return self._model.isSimilar() or self._isSameTable()

    def _isSameTable(self):
        subquery = self._view.getSubQuery()
        if subquery is not None:
            return self._view.getTable() == subquery.Second
        return False
