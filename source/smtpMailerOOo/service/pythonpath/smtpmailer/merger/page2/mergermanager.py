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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .mergerview import MergerView

from .mergerhandler import Tab1Handler
from .mergerhandler import Tab2Handler

from .mergerhandler import AddressHandler
from .mergerhandler import RecipientHandler

from ...grid import GridListener

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
        tables, title1, title2, message = self._model.getPageInfos()
        self._view = MergerView(ctx, Tab1Handler(self), Tab2Handler(self), parent, tables, title1, title2, message)
        window1, window2 = self._view.getGridWindows()
        self._model.initPage2(window1, window2, self.initPage)

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
        if self._model.hasQueryChanged():
            label = self._model.getQueryLabel()
            print("MergerManager.activatePage() Label changed: %s" % label)
            self._view.setQueryLabel(label)
        if self._model.hasTablesChanged():
            tables, table = self._model.getQueryTables()
            print("MergerManager.activatePage() Tables changed: %s" % table)
            # FIXME: We must disable the handler "ChangeAddressTable"
            # FIXME: otherwise it activates twice
            self._disableHandler()
            self._view.initTables(tables, table)

    def commitPage(self, reason):
        self._model.resetPendingChanges()
        return True

    def canAdvance(self):
        print("MergerManager2.canAdvance() 1")
        advance = self._model.getRecipientCount() > 0
        print("MergerManager2.canAdvance() 2 %s" % advance)
        return advance

# MergerManager setter methods
    def initPage(self, grid1, grid2, rowset1, rowset2, table):
        grid1.addSelectionListener(GridListener(self, 1))
        grid2.addSelectionListener(GridListener(self, 2))
        rowset1.addRowSetListener(AddressHandler(self))
        rowset2.addRowSetListener(RecipientHandler(self))
        # FIXME: We must disable the handler "ChangeAddressBook"
        # FIXME: otherwise it activates twice
        self._disableHandler()
        self._view.setTable(table)

    def changeAddressRowSet(self, rowset):
        self._model.setGrid1Data(rowset)

    def changeRecipientRowSet(self, rowset):
        count = self._model.setGrid2Data(rowset)
        self._view.enableAddAll(count == 1)
        self._view.enableRemoveAll(count == 0)
        self._wizard.updateTravelUI()

    def setAddressTable(self, table):
        self._model.setAddressTable(table)

    def changeGridSelection(self, index, grid):
        selected = index != -1
        enabled = self._model.hasFilters()
        if grid == 1:
            self._view.enableAdd(selected and enabled)
            if selected:
                self._model.setAddressRecord(index)
        elif grid == 2:
            print("MergerModel.changeGrid2Selection() Selected: %s" % selected)
            self._view.enableRemove(selected and enabled)
            if selected:
                self._model.setRecipientRecord(index)

    def addItem(self):
        filters = self._model.getGrid1SelectedStructuredFilters()
        self._model.addItem(filters, self.enableRemoveAll)

    def addAllItem(self):
        table = self._view.getTable()
        self._model.addAllItem(table)

    def removeItem(self):
        filters = self._model.getGrid2SelectedStructuredFilters()
        self._model.removeItem(filters, self.enableAddAll)

    def removeAllItem(self):
        self._model.removeAllItem()

    def enableAddAll(self, hasnofilter):
        self._view.enableAddAll(hasnofilter)

    def enableRemoveAll(self, hasnofilter):
        self._view.enableRemoveAll(hasnofilter)

