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

from smtpmailer import getMessage
from smtpmailer import logMessage

from .mergerview import MergerView
from .mergerhandler import AddressHandler
from .mergerhandler import RecipientHandler

import traceback


class MergerManager(unohelper.Base,
                    XWizardPage):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._disabled = False
        tables, table, enabled, message = self._model.getPageInfos(True)
        self._view = MergerView(ctx, self, parent, tables, enabled, message)
        print("mergerManager.__init__() 1")
        address = AddressHandler(self)
        recipient = RecipientHandler(self)
        print("mergerManager.__init__() 2")
        self._model.initGrid(table, address, recipient, self.initGrid1, self.initGrid2)
        print("mergerManager.__init__() 3")
        # TODO: We must disable the handler "ChangeAddressBook" otherwise it activates twice
        self._disableHandler()
        print("mergerManager.__init__() 4")
        self._view.setTable(table)
        print("mergerManager.__init__() 5")

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
        if self._model.isChanged():
            tables, table, enabled, message = self._model.getPageInfos()
            self._view.setMessage(message)
            # TODO: We must disable the handler "ChangeAddressTable"
            # TODO: otherwise it activates twice
            self._disableHandler()
            self._view.initTables(tables, table, enabled)

    def commitPage(self, reason):
        return True

    def canAdvance(self):
        print("MergerManager2.canAdvance() 1")
        advance = self._model.getRecipientCount() > 0
        print("MergerManager2.canAdvance() 2 %s" % advance)
        return advance

# MergerManager getter methods
    def getGridModels(self, tab):
        return self._model.getGridModels(tab)

    def getTabTitle(self, tab):
        return self._model.getTabTitle(tab)

# MergerManager setter methods
    def initGrid1(self, titles, orders):
        self._view.initGrid1(self)
        self._view.initColumn1(titles)
        self._disableHandler()
        self._view.initOrder1(titles, orders)

    def initGrid2(self, titles, orders):
        self._view.initGrid2(self)
        self._view.initColumn2(titles)
        self._disableHandler()
        self._view.initOrder2(titles, orders)

    def setAddressTable(self, table):
        self._model.setAddressTable(table)

    def changeRecipientRowSet(self, enabled):
        self._model.changeRecipientRowSet()
        self._view.enableRemoveAll(enabled)
        self._wizard.updateTravelUI()

    def changeAddressRowSet(self, enabled):
        self._model.changeAddressRowSet()
        self._view.enableAddAll(enabled)

    def setAddressColumn(self, titles, reset):
        self._model.setAddressColumn(titles, reset)

    def setAddressOrder(self, titles):
        ascending = self._view.getAddressSort()
        self._model.setAddressOrder(titles, ascending)

    def setRecipientColumn(self, titles, reset):
        self._model.setRecipientColumn(titles, reset)

    def setRecipientOrder(self, titles):
        ascending = self._view.getRecipientSort()
        self._model.setRecipientOrder(titles, ascending)

    def changeGrid1Selection(self, selected, index):
        self._view.enableAdd(selected)
        if index != -1:
            self._model.setAddressRecord(index)

    def changeGrid2Selection(self, selected, index):
        self._view.enableRemove(selected)
        if index != -1:
            self._model.setRecipientRecord(index)

    def addItem(self):
        rows = self._view.getSelectedAddress()
        self._model.addItem(rows)

    def addAllItem(self):
        rows = range(self._model.getAddressCount())
        self._model.addItem(rows)

    def removeItem(self):
        rows = self._view.getSelectedRecipient()
        self._model.removeItem(rows)

    def removeAllItem(self):
        rows = range(self._model.getRecipientCount())
        self._model.removeItem(rows)
