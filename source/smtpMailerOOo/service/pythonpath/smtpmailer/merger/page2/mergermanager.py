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
        tables, table, message = self._model.getPageInfos()
        self._view = MergerView(ctx, self, parent, tables, message)
        window1, window2 = self._view.getGridWindows()
        self._model.initPage2(table, window1, window2, self.initPage)
        # FIXME: We must disable the handler "ChangeAddressBook"
        # FIXME: otherwise it activates twice
        self._disableHandler()
        self._view.setTable(table)

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
        print("MergerManager.activatePage() %s" % self._model.isChanged())
        if self._model.isChanged():
            tables, table, enabled, message = self._model.getPageInfos()
            self._view.setMessage(message)
            # FIXME: We must disable the handler "ChangeAddressTable"
            # FIXME: otherwise it activates twice
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
    def getTabTitle(self, tab):
        return self._model.getTabTitle(tab)

# MergerManager setter methods
    def initPage(self, grid1, grid2, rowset1, rowset2):
        grid1.addSelectionListener(GridListener(self, 1))
        grid2.addSelectionListener(GridListener(self, 2))
        rowset1.addRowSetListener(AddressHandler(self))
        rowset2.addRowSetListener(RecipientHandler(self))

    def changeAddressRowSet(self, rowset):
        self._model.setGrid1Data(rowset)
        enabled = rowset.RowCount > 0
        self._view.enableAddAll(enabled)

    def changeRecipientRowSet(self, rowset):
        self._model.setGrid2Data(rowset)
        enabled = rowset.RowCount > 0
        self._view.enableRemoveAll(enabled)
        self._wizard.updateTravelUI()

    def setAddressTable(self, table):
        self._model.setAddressTable(table)

    def changeGridSelection(self, index, grid):
        selected = index != -1
        if grid == 1:
            self._view.enableAdd(selected)
            if selected:
                self._model.setAddressRecord(index)
        elif grid == 2:
            self._view.enableRemove(selected)
            if selected:
                self._model.setRecipientRecord(index)

    def addItem(self):
        filters = self._model.getGrid1SelectedStructuredFilters()
        self._model.addItem(filters)

    def addAllItem(self):
        table = self._view.getTable()
        self._model.addAllItem(table)

    def removeItem(self):
        filters = self._model.getGrid2SelectedStructuredFilters()
        self._model.removeItem(filters)

    def removeAllItem(self):
        pass
        #rows = range(self._model.getRecipientCount())
        #self._model.removeItem(rows)
